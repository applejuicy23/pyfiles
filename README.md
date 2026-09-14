# pyfiles

Simple GUI tool for file operations.

## Current architecture

- `pyfiles/pyfiles0.11.1.py` — current application source (Tkinter UI + refactored service boundaries).
- `pyfiles/app.py` — lightweight entrypoint that runs `pyfiles0.11.1.py`.
- `pyfiles/services/*` — pure service layer (move/copy/delete/create helpers, resource/path utils, recycle/bin operations).

Legacy single-version script dumps from earlier releases are intentionally removed from source control to keep the repository focused on the current refactor branch.
