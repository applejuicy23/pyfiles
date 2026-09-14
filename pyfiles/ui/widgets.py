"""Reusable widget helpers (placeholder)."""

from __future__ import annotations


def attach_mousewheel_scroll(widget, scroll_fn):
    if widget is None:
        return
    widget.bind("<MouseWheel>", lambda event: scroll_fn(int(-1 * (event.delta / 120)), "units"))


def bind_tree_copy_like(item_widget, select_all_handler):
    if item_widget is None:
        return
    item_widget.bind("<Control-a>", select_all_handler)


__all__ = ["attach_mousewheel_scroll", "bind_tree_copy_like"]
