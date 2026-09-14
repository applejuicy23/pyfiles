"""Create-file utility helpers.

Only pure transformations are kept here to keep legacy behavior stable.
"""

from __future__ import annotations

from typing import List, Sequence


FORBIDDEN_FILENAME_CHARS = r'<>:"/\\|?*'


def sanitize_filename(name: str, forbidden: str = FORBIDDEN_FILENAME_CHARS) -> str:
    """Replace characters forbidden in Windows filenames with underscores."""

    clean = name
    for ch in forbidden:
        clean = clean.replace(ch, "_")
    return clean


def parse_count_input(raw: str) -> List[int]:
    """Parse user input like `1`, `1,2,3`, `1:5` into list of ints.

    Behaviour kept compatible with legacy parser.
    """

    text = raw.strip()
    if not text:
        return [1]

    result: List[int] = []
    for part in text.split(","):
        part = part.strip()

        if not part:
            continue

        if ":" in part:
            try:
                start, end = map(int, part.split(":"))
                result.extend(range(start, end + 1))
                continue
            except Exception:
                continue

        try:
            result.append(int(part))
        except Exception:
            continue

    if len(result) == 1 and ":" not in text and "," not in text:
        n = result[0]
        return list(range(1, n + 1))

    return result


def apply_prefix(prefix: str, name: str, number: int, ext: str,
                 date_value: str | None = None, time_value: str | None = None) -> str:
    """Expand placeholders used by create-mode.

    Parameters `date_value` / `time_value` may be passed by caller;
    this keeps the function pure while remaining compatible with legacy logic.
    """

    result = prefix

    result = result.replace("{name}", name)
    result = result.replace("{file}", name)

    if "{num}" in result:
        result = result.replace("{num}", str(number))

    if "{date}" in result and date_value is not None:
        result = result.replace("{date}", date_value)

    if "{time}" in result and time_value is not None:
        result = result.replace("{time}", time_value)

    return result


def build_filename(prefix: str, base_name: str, number: int, ext: str) -> str:
    """Backward-compatible filename generator using prefix macroes."""

    safe_name = sanitize_filename(f"{prefix.replace('{name}', base_name).replace('{file}', base_name)}{ext}".replace("{num}", str(number)))
    return safe_name


__all__: Sequence[str] = (
    "sanitize_filename",
    "parse_count_input",
    "apply_prefix",
    "build_filename",
)
