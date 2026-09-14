"""Shared pure utility functions extracted from legacy UI application.

These helpers keep core behavior stable while gradually removing global utility
functions from the monolithic legacy file.
"""

from __future__ import annotations

import os
import re
from typing import Any, Iterable, Sequence, Tuple


def resource_path(relative_path: str, base_path: str | None = None) -> str:
    """Build path to bundled or local resource.

    Supports both standalone and development layouts where assets are in `icons/`.
    """

    try:
        import sys

        root = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        root = None

    from pathlib import Path

    file_dir = Path(__file__).resolve().parent
    cwd = Path(os.getcwd())
    provided = Path(base_path) if base_path else None

    candidates = [
        p
        for p in [
            Path(root) if root else None,
            provided,
            file_dir,
            file_dir / "icons",
            file_dir.parent,
            file_dir.parent / "icons",
            file_dir.parent.parent,
            file_dir.parent.parent / "icons",
            cwd,
            cwd / "icons",
        ]
        if p is not None
    ]

    for base in candidates:
        candidate = base / relative_path
        if candidate.exists():
            return str(candidate)

    fallback = file_dir.parent.parent / relative_path
    if fallback.exists():
        return str(fallback)

    return str(file_dir.parent.parent / "icons" / relative_path)


def norm_path(p: str) -> str:
    """Normalize path for path-comparison helpers."""

    return os.path.abspath(os.path.normcase(os.path.normpath(p)))


def format_items(files: int, folders: int, names: Sequence[str] | None = None, action: str = "Added") -> str:
    """Human-readable operation summary with grammar for singular/plural."""

    if names and len(names) == 1:
        name = names[0]
        if files:
            return f"{action} file {name}"
        return f"{action} folder {name}"

    parts = []
    if folders:
        parts.append(f"{folders} folder{'s' if folders != 1 else ''}")
    if files:
        parts.append(f"{files} file{'s' if files != 1 else ''}")

    return f"{action} {' and '.join(parts)}"


def count_inside_folder(folder: str) -> Tuple[int, int]:
    """Return total subfolders and files count for a folder tree."""

    files = 0
    folders = 0

    for _root, dirs, filenames in os.walk(folder):
        folders += len(dirs)
        files += len(filenames)

    return folders, files


def natural_sort_key(value: str) -> list[Any]:
    """Create key for natural sorting (alnum chunks)."""

    return [int(text) if text.isdigit() else text.lower() for text in re.split(r"(\d+)", value)]
