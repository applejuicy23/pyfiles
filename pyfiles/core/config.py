"""Configuration and theme definitions for PyFiles.

This module is a new home for app-wide constants after splitting monolithic code.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Final


APP_TITLE: Final = "PyFiles"
APP_VERSION: Final = "0.11.1"


THEMES = {
    "default": {
        "bg": "SystemButtonFace",
        "fg": "black",
        "panel": "SystemButtonFace",
        "accent": "#e8e8e8",
        "scroll": "SystemButtonFace",
    },
    "light": {
        "bg": "#ffffff",
        "fg": "#000000",
        "panel": "#ffffff",
        "accent": "#eaeaea",
    },
    "dark": {
        "bg": "#1e1e1e",
        "fg": "#ffffff",
        "panel": "#2a2a2a",
        "accent": "#3a3a3a",
        "scroll": "#1e1e1e",
    },
}

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".avif", ".bmp", ".gif", ".tiff", ".ico"}


@dataclass(frozen=True)
class FileOperationConfig:
    """Костыльные значения для операций. Вынесены сюда для безопасного редизайна."""

    cache_dirname: str = "PyFiles_Cache"
    bin_dirname: str = ".pyfiles_bin"
    log_dirname: str = "LOGS"
