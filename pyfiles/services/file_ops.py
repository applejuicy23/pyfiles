"""Pure file operation logic (move / copy / scan)."""

from __future__ import annotations

import os
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, List, Literal, Sequence, Set, Tuple


action_t = Literal["replace", "rename", "skip"]


@dataclass
class OperationItem:
    source: str
    status: str
    target: str | None = None
    error: str | None = None


@dataclass
class FileOpResult:
    operation: str
    total: int
    success: int = 0
    skipped: int = 0
    logs: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    error_paths: Set[str] = field(default_factory=set)
    items: List[OperationItem] = field(default_factory=list)

    @property
    def done(self) -> int:
        return self.success + self.skipped

    @property
    def failed(self) -> int:
        return len(self.errors)


def split_paths_for_operations(paths: Iterable[str]) -> Tuple[List[str], List[str]]:
    """Split paths list into files and directories."""

    files: List[str] = []
    folders: List[str] = []

    for p in paths:
        pth = Path(p)
        if pth.is_dir():
            folders.append(str(pth))
        elif pth.is_file():
            files.append(str(pth))
        else:
            # non-existent path handled by caller
            files.append(str(pth))

    return files, folders


def build_target_path(src: str, destination: str) -> tuple[str, str]:
    """Return display name and full destination path for source name in folder."""

    src_path = Path(src)
    return src_path.name, str(Path(destination) / src_path.name)


def resolve_conflict_path(destination: str, *, on_exists: Callable[[str, str], action_t] | None = None) -> tuple[str, str]:
    """Resolve name collisions for destination path.

    `on_exists(path, candidate)` called with file name and existing destination.
    Returns `(action, final_target)` where action is one of:
    - replace: remove existing target and use it
    - rename: choose free name `<name> (n)<ext>`
    - skip: keep same target as is but no write operation should happen
    """

    if not os.path.exists(destination):
        return "replace", destination

    if on_exists is None:
        # Safe default for headless callers.
        return "replace", destination

    action = on_exists("", destination)
    if action == "rename":
        base, ext = os.path.splitext(os.path.basename(destination))
        parent = Path(destination).parent
        counter = 1

        while True:
            new_name = f"{base} ({counter}){ext}"
            candidate = str(parent / new_name)
            if not os.path.exists(candidate):
                return "replace", candidate
            counter += 1

    return action, destination


def _safe_remove_if_exists(path: str) -> None:
    if os.path.isdir(path):
        shutil.rmtree(path)
    elif os.path.exists(path):
        os.remove(path)


def move_to_cache(path: str, base_src: str, cache_dir: str) -> None:
    rel_path = os.path.relpath(path, base_src)
    target = os.path.join(cache_dir, rel_path)

    os.makedirs(os.path.dirname(target), exist_ok=True)
    shutil.move(path, target)


def move_items(
    paths: Iterable[str],
    destination: str,
    *,
    conflict_resolver: Callable[[str, str], action_t] | None = None,
    excluded_paths: Iterable[str] | None = None,
    cache_dir: str | None = None,
) -> FileOpResult:
    """Move files/directories to destination preserving substructure for dirs.

    Returns explicit statuses/logs/errors instead of raising.
    """

    path_list = [str(p) for p in paths]
    excluded = {os.path.normcase(os.path.normpath(p)) for p in (excluded_paths or set())}
    result = FileOpResult(operation="move", total=len(path_list))

    def _is_excluded(p: str) -> bool:
        return os.path.normcase(os.path.normpath(p)) in excluded

    for src in path_list:
        src_abs = str(Path(src))
        if not os.path.exists(src_abs):
            err = f"{os.path.basename(src_abs)} >> ERROR: File not found"
            result.errors.append(err)
            result.error_paths.add(src_abs)
            result.items.append(OperationItem(source=src_abs, status="error", error=err))
            continue

        name = os.path.basename(src_abs)

        try:
            if os.path.isfile(src_abs):
                _, target = build_target_path(src_abs, destination)
                action, target = resolve_conflict_path(
                    target,
                    on_exists=lambda _n, dst: conflict_resolver(name, dst) if conflict_resolver else "replace",
                )

                if action == "skip":
                    result.skipped += 1
                    result.items.append(OperationItem(source=src_abs, status="skip", target=target))
                    continue

                if action == "replace":
                    _safe_remove_if_exists(target)

                if _is_excluded(src_abs):
                    if cache_dir is None:
                        raise RuntimeError("cache_dir is required for excluded files")
                    move_to_cache(src_abs, os.path.dirname(src_abs), cache_dir)
                    result.success += 1
                    result.items.append(OperationItem(source=src_abs, status="moved_to_cache", target=target))
                else:
                    shutil.move(src_abs, target)
                    result.success += 1
                    result.items.append(OperationItem(source=src_abs, status="moved", target=target))

                result.logs.append(f"Moved: {name}")
                continue

            if os.path.isdir(src_abs):
                _, root_target = build_target_path(src_abs, destination)
                action, root_target = resolve_conflict_path(
                    root_target,
                    on_exists=lambda _n, dst: conflict_resolver(name, dst) if conflict_resolver else "replace",
                )

                if action == "skip":
                    result.skipped += 1
                    result.items.append(OperationItem(source=src_abs, status="skip", target=root_target))
                    continue

                if action == "replace":
                    _safe_remove_if_exists(root_target)

                os.makedirs(root_target, exist_ok=True)

                had_error = False

                for root_dir, _dirs, files in os.walk(src_abs):
                    for fname in files:
                        full = os.path.join(root_dir, fname)

                        if _is_excluded(full):
                            if cache_dir is None:
                                err = f"{os.path.basename(full)} >> ERROR: cache_dir is required for excluded files"
                                result.errors.append(err)
                                result.error_paths.add(full)
                                result.items.append(OperationItem(source=full, status="error", error=err))
                                had_error = True
                                continue

                            try:
                                move_to_cache(full, os.path.dirname(full), cache_dir)
                                result.logs.append(f"Moved to cache: {os.path.basename(full)}")
                                result.items.append(OperationItem(source=full, status="moved_to_cache", target=os.path.join(root_target, os.path.relpath(full, src_abs))))
                            except Exception as e:
                                err = f"{os.path.basename(full)} >> ERROR: {str(e)}"
                                result.errors.append(err)
                                result.error_paths.add(full)
                                result.items.append(OperationItem(source=full, status="error", error=err))
                                had_error = True
                            continue

                        try:
                            if not os.path.exists(full):
                                err = f"{fname} >> ERROR: File not found"
                                result.errors.append(err)
                                result.error_paths.add(full)
                                result.items.append(OperationItem(source=full, status="error", error=err))
                                had_error = True
                                continue

                            rel = os.path.relpath(full, src_abs)
                            dest_file = os.path.join(root_target, rel)

                            os.makedirs(os.path.dirname(dest_file), exist_ok=True)

                            # Keep legacy behavior: move file directly.
                            shutil.move(full, dest_file)
                            result.items.append(OperationItem(source=full, status="moved", target=dest_file))
                        except Exception as e:
                            err = f"{os.path.basename(full)} >> ERROR: {str(e)}"
                            result.errors.append(err)
                            result.error_paths.add(full)
                            result.items.append(OperationItem(source=full, status="error", error=err))
                            had_error = True

                if not had_error:
                    shutil.rmtree(src_abs, ignore_errors=False)
                    result.success += 1
                    result.items.append(OperationItem(source=src_abs, status="moved", target=root_target))
                    result.logs.append(f"Moved folder: {name}")
                else:
                    # keep source folder to inspect partially moved content
                    result.errors.append(f"Folder {name} moved with errors")
                    result.error_paths.add(src_abs)

                continue

            # neither file nor dir
            err = f"{name} >> ERROR: Unsupported path type"
            result.errors.append(err)
            result.error_paths.add(src_abs)
            result.items.append(OperationItem(source=src_abs, status="error", error=err))

        except Exception as e:
            err = f"{name} >> ERROR: {str(e)}"
            result.errors.append(err)
            result.error_paths.add(src_abs)
            result.items.append(OperationItem(source=src_abs, status="error", error=err))

    return result



def copy_items(
    paths: Iterable[str],
    destination: str,
    *,
    conflict_resolver: Callable[[str, str], action_t] | None = None,
) -> FileOpResult:
    """Copy files/directories to destination and return explicit operation report."""

    path_list = [str(p) for p in paths]
    result = FileOpResult(operation="copy", total=len(path_list))

    for src in path_list:
        src_abs = str(Path(src))

        if not os.path.exists(src_abs):
            err = f"{os.path.basename(src_abs)} >> ERROR: File not found"
            result.errors.append(err)
            result.error_paths.add(src_abs)
            result.items.append(OperationItem(source=src_abs, status="error", error=err))
            continue

        try:
            name = os.path.basename(src_abs)

            if os.path.isfile(src_abs):
                _, target = build_target_path(src_abs, destination)
                action, target = resolve_conflict_path(
                    target,
                    on_exists=lambda _n, dst: conflict_resolver(name, dst) if conflict_resolver else "replace",
                )

                if action == "skip":
                    result.skipped += 1
                    result.items.append(OperationItem(source=src_abs, status="skip", target=target))
                    continue

                if action == "replace":
                    _safe_remove_if_exists(target)

                shutil.copy2(src_abs, target)
                result.success += 1
                result.items.append(OperationItem(source=src_abs, status="copied", target=target))
                result.logs.append(f"Copied: {name}")

            elif os.path.isdir(src_abs):
                _, target = build_target_path(src_abs, destination)
                action, target = resolve_conflict_path(
                    target,
                    on_exists=lambda _n, dst: conflict_resolver(name, dst) if conflict_resolver else "replace",
                )

                if action == "skip":
                    result.skipped += 1
                    result.items.append(OperationItem(source=src_abs, status="skip", target=target))
                    continue

                if action == "replace":
                    _safe_remove_if_exists(target)

                # Copy whole folder in one step to preserve structure.
                shutil.copytree(src_abs, target)
                result.success += 1
                result.items.append(OperationItem(source=src_abs, status="copied", target=target))
                result.logs.append(f"Copied: {name}")

            else:
                err = f"{name} >> ERROR: Unsupported path type"
                result.errors.append(err)
                result.error_paths.add(src_abs)
                result.items.append(OperationItem(source=src_abs, status="error", error=err))
        except Exception as e:
            err = f"{os.path.basename(src_abs)} >> ERROR: {str(e)}"
            result.errors.append(err)
            result.error_paths.add(src_abs)
            result.items.append(OperationItem(source=src_abs, status="error", error=err))

    return result


def scan_files(base_path: str, max_depth: int) -> tuple[list[str], list[str]]:
    """Scan folder tree with depth limit and return found files and collected errors."""

    result: list[str] = []
    errors: list[str] = []

    if not os.path.isdir(base_path):
        errors.append(f"{base_path} >> ERROR: Source folder not found")
        return result, errors

    try:
        base_path_norm = os.path.normpath(base_path)
        base_depth = base_path_norm.count(os.sep)

        for root_dir, dirs, files in os.walk(base_path_norm):
            current_depth = root_dir.count(os.sep) - base_depth
            if current_depth >= max_depth:
                dirs[:] = []

            for fname in files:
                result.append(os.path.join(root_dir, fname))
    except Exception as e:
        errors.append(str(e))

    return result, errors
