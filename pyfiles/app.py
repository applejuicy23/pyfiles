"""Application entrypoint."""

from __future__ import annotations

import runpy
from pathlib import Path


def main() -> None:
    app_file = Path(__file__).with_name("pyfiles0.11.1.py")
    runpy.run_path(str(app_file), run_name="__main__")


if __name__ == "__main__":
    main()
