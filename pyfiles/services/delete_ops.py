"""Pure delete/safe-delete/restore logic."""

from __future__ import annotations

import json
import os
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Set
from uuid import uuid4


@dataclass
class DeleteResult:
    operation: str
    total: int
    success: int = 0
    skipped: int = 0
    logs: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    error_paths: Set[str] = field(default_factory=set)
    removed_ids: List[str] = field(default_factory=list)
    restored_ids: List[str] = field(default_factory=list)

    @property
    def failed(self) -> int:
        return len(self.errors)


@dataclass
class RecycleRestoreResult:
    operation: str = "restore_from_recycle_bin"
    restored: bool = False
    logs: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    source: str | None = None
    destination: str | None = None


def load_meta(meta_path: str) -> tuple[dict, str | None]:
    """Read meta file safely. Returns (meta, error)."""

    if not os.path.exists(meta_path):
        return {}, None

    try:
        with open(meta_path, "r", encoding="utf-8") as f:
            return json.load(f), None
    except Exception as e:
        return {}, str(e)


def save_meta(meta_path: str, meta: dict) -> str | None:
    """Write meta file and return error string on failure."""

    try:
        os.makedirs(os.path.dirname(meta_path), exist_ok=True)
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=4)
        return None
    except Exception as e:
        return str(e)


def is_safe_mode(enabled: bool) -> bool:
    return bool(enabled)


def _delete_real_path(path: str) -> None:
    if os.path.isfile(path):
        os.remove(path)
    else:
        shutil.rmtree(path)


def restore_from_recycle_bin(
    file_name: str,
    folder: str,
    *,
    recycle_root: str | None = None,
) -> RecycleRestoreResult:
    """Restore a single entry by file name from Windows $Recycle.Bin."""

    result = RecycleRestoreResult()

    if not file_name:
        result.errors.append("Missing file name")
        return result

    if not folder:
        result.errors.append("Missing destination folder")
        return result

    recycle_root_resolved = Path(
        recycle_root
        if recycle_root is not None
        else Path(os.environ.get("USERPROFILE", os.getcwd())) / "$Recycle.Bin"
    )

    if not recycle_root_resolved.exists():
        result.errors.append("Recycle bin folder not found")
        return result

    for root_dir, _dirs, files in os.walk(str(recycle_root_resolved)):
        if file_name not in files:
            continue

        src = os.path.join(root_dir, file_name)
        dst = os.path.join(folder, file_name)

        result.source = src
        result.destination = dst

        try:
            os.makedirs(folder, exist_ok=True)
            if os.path.exists(dst):
                if os.path.isdir(dst):
                    shutil.rmtree(dst)
                else:
                    os.remove(dst)

            shutil.move(src, dst)
            result.restored = True
            result.logs.append(f"Restored {file_name}")
        except Exception as e:
            result.errors.append(f"{file_name} >> ERROR: {str(e)}")
        return result

    result.errors.append("File not found in recycle bin")
    return result


def delete_items(
    paths: Iterable[str],
    *,
    safe_mode: bool,
    bin_dir: str,
    meta_path: str,
    overwrite_bin: bool = True,
) -> DeleteResult:
    """Delete paths directly or move them into a local `.pyfiles_bin`.

    Returns explicit counts and error messages, no raw raises.
    """

    path_list = [str(p) for p in paths]
    result = DeleteResult(operation="delete", total=len(path_list))

    meta, meta_error = load_meta(meta_path)
    if meta_error:
        result.errors.append(f"Meta error: {meta_error}")

    os.makedirs(bin_dir, exist_ok=True)

    for path in path_list:
        path_norm = str(Path(path).resolve())
        filename = os.path.basename(path_norm)

        if not os.path.exists(path_norm):
            err = f"{filename} >> ERROR: File not found"
            result.errors.append(err)
            result.error_paths.add(path_norm)
            continue

        try:
            if safe_mode:
                file_id = str(uuid4())
                dst = os.path.join(bin_dir, file_id)

                # keep old behavior (strict move to unique id inside bin)
                if os.path.exists(dst):
                    if overwrite_bin:
                        if os.path.isdir(dst):
                            shutil.rmtree(dst)
                        else:
                            os.remove(dst)
                    else:
                        err = f"{filename} >> ERROR: ID collision in bin"
                        result.errors.append(err)
                        result.error_paths.add(path_norm)
                        continue

                shutil.move(path_norm, dst)
                meta[file_id] = {
                    "original_path": path_norm,
                    "name": filename,
                }
                result.success += 1
                result.removed_ids.append(file_id)
                result.logs.append(f"[BIN] {filename} moved to bin")
            else:
                _delete_real_path(path_norm)
                result.success += 1
                result.logs.append(f"[DEL] {filename} deleted")
        except Exception as e:
            err = f"{filename} >> ERROR: {str(e)}"
            result.errors.append(err)
            result.error_paths.add(path_norm)

    if result.failed == 0:
        if not result.errors and not result.success:
            result.skipped = result.total

    save_err = save_meta(meta_path, meta)
    if save_err:
        result.errors.append(f"Meta error: {save_err}")

    return result


def delete_from_bin(
    file_ids: Iterable[str],
    *,
    bin_dir: str,
    meta_path: str,
) -> DeleteResult:
    ids = list(file_ids)
    result = DeleteResult(operation="delete_from_bin", total=len(ids))

    meta, meta_error = load_meta(meta_path)
    if meta_error:
        result.errors.append(f"Meta error: {meta_error}")

    for file_id in ids:
        name = file_id
        path = os.path.join(bin_dir, str(file_id))
        try:
            data = meta.get(str(file_id)) if isinstance(meta, dict) else None
            if data and isinstance(data, dict):
                name = data.get("name", name)

            if not os.path.exists(path):
                raise FileNotFoundError("item not found in bin")

            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)

            meta.pop(str(file_id), None)
            result.success += 1
            result.logs.append(f"[BIN DEL] {name} permanently deleted")
            result.removed_ids.append(str(file_id))
        except Exception as e:
            err = f"{name} >> ERROR: {str(e)}"
            result.errors.append(err)
            result.error_paths.add(str(file_id))

    save_err = save_meta(meta_path, meta)
    if save_err:
        result.errors.append(f"Meta error: {save_err}")

    return result


def restore_items(
    file_ids: Iterable[str],
    *,
    bin_dir: str,
    meta_path: str,
    overwrite: bool = False,
) -> DeleteResult:
    ids = list(file_ids)
    result = DeleteResult(operation="restore", total=len(ids))

    meta, meta_error = load_meta(meta_path)
    if meta_error:
        result.errors.append(f"Meta error: {meta_error}")

    for file_id in ids:
        file_id = str(file_id)
        data = meta.get(file_id) if isinstance(meta, dict) else None

        if not isinstance(data, dict):
            result.errors.append(f"{file_id} >> ERROR: Unknown bin item")
            result.error_paths.add(file_id)
            continue

        src = os.path.join(bin_dir, file_id)
        dst = data.get("original_path")
        name = data.get("name", file_id)

        try:
            if not os.path.exists(src):
                raise FileNotFoundError("file not found in bin")

            if dst is None:
                raise ValueError("missing original_path")

            if os.path.exists(dst):
                if not overwrite:
                    raise FileExistsError(f"{name} already exists at original path")
                if os.path.isdir(dst):
                    shutil.rmtree(dst)
                else:
                    os.remove(dst)

            parent = os.path.dirname(dst)
            if parent:
                os.makedirs(parent, exist_ok=True)

            shutil.move(src, dst)
            meta.pop(file_id, None)
            result.success += 1
            result.restored_ids.append(file_id)
            result.logs.append(f"Restored {name}")
        except Exception as e:
            err = f"{name} >> ERROR: {str(e)}"
            result.errors.append(err)
            result.error_paths.add(file_id)

    save_err = save_meta(meta_path, meta)
    if save_err:
        result.errors.append(f"Meta error: {save_err}")

    return result
