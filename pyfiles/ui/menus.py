"""Menu builders and context actions for Tkinter (placeholder)."""

from __future__ import annotations


def build_file_menu(*_args, **_kwargs):
    """Возвращает заготовку File menu. После миграции сюда вынесем binding-логика."""
    return None


def build_delete_menu(*_args, **_kwargs):
    return None


def build_theme_handlers(app):
    """Hook for theme switch callbacks; currently placeholder."""
    return {
        "light": lambda: None,
        "default": lambda: None,
        "dark": lambda: None,
    }


__all__ = ["build_file_menu", "build_delete_menu", "build_theme_handlers"]
