"""UI window composition (placeholder for staged migration)."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class UiComponents:
    """Контейнер для ссылок на виджеты Tkinter."""

    root = None
    files_tree = None
    dest_tree = None
    delete_tree = None
    delete_bin_tree = None
    status_label = None
    progress_bar = None


__all__ = ["UiComponents"]
