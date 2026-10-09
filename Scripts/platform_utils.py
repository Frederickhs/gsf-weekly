from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


APP_NAME = "Success Factors Weekly Training Report"


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def get_project_root(current_file: str | Path) -> Path:
    """Return the source project root or the Windows distribution folder."""
    if is_frozen() and sys.platform == "win32":
        return Path(sys.executable).resolve().parent

    return Path(current_file).resolve().parent.parent


def get_app_data_root(current_file: str | Path) -> Path:
    """
    Return the writable folder used for Output, Archive, and Logs.

    Windows frozen build:
        Folder containing the EXE.
    macOS frozen build:
        ~/Documents/Success Factors Weekly Training Report
    Development:
        Project root.
    """
    if is_frozen() and sys.platform == "darwin":
        return Path.home() / "Documents" / APP_NAME

    return get_project_root(current_file)


def ensure_app_folders(current_file: str | Path) -> Path:
    root = get_app_data_root(current_file)
    for folder_name in ("Output", "Archive", "Logs"):
        (root / folder_name).mkdir(parents=True, exist_ok=True)
    return root


def open_path(path: str | Path) -> None:
    target = str(Path(path).resolve())

    if sys.platform == "win32":
        os.startfile(target)  # type: ignore[attr-defined]
        return

    if sys.platform == "darwin":
        subprocess.run(["open", target], check=False)
        return

    subprocess.run(["xdg-open", target], check=False)
