"""Mutable application state (plain object passed explicitly between modules).

We keep this small and explicit to kill global-variable chaos from the legacy file.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set


@dataclass
class AppState:
    # ---- navigation/state ----
    current_mode: str = "MOVE"
    current_theme: str = "default"
    selected_files: List[str] = field(default_factory=list)
    added_paths: Set[str] = field(default_factory=set)
    excluded_paths: Set[str] = field(default_factory=set)
    destination_folder: str = ""
    last_status: Optional[tuple[str, str]] = None
    status_job: Optional[str] = None

    # ---- delete bin ----
    delete_files_list: List[str] = field(default_factory=list)
    delete_bin_meta: Dict[str, Any] = field(default_factory=dict)

    # ---- runtime/job ----
    progress_job: Optional[str] = None
    remove_hold_job: Optional[str] = None

    # ---- caches/handlers ----
    icons_cache: List[Any] = field(default_factory=list)
    icon_cache_map: Dict[str, Any] = field(default_factory=dict)
    preview_icon_cache: Dict[str, Any] = field(default_factory=dict)
    bin_icons_cache: List[Any] = field(default_factory=list)

    # ---- misc ----
    selected_file: Optional[str] = None


# Global singleton context (for gradual migration from implicit globals)
state = AppState()


def reset_state(st: AppState) -> None:
    """Полный сброс runtime-состояния, кроме тех полей, которые считаем конфигом."""

    st.current_mode = "MOVE"
    st.current_theme = "default"
    st.selected_files.clear()
    st.added_paths.clear()
    st.excluded_paths.clear()
    st.destination_folder = ""
    st.last_status = None
    st.status_job = None
    st.delete_files_list.clear()
    st.delete_bin_meta.clear()
    st.progress_job = None
    st.remove_hold_job = None
    st.icons_cache.clear()
    st.icon_cache_map.clear()
    st.preview_icon_cache.clear()
    st.bin_icons_cache.clear()
    st.selected_file = None


def set_destination(st: AppState, folder: Optional[Path | str]) -> None:
    st.destination_folder = str(folder) if folder else ""
