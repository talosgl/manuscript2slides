"""Utilities for use across the entire program."""

import io

import logging
import platform
import sys
from pathlib import Path
import subprocess

log = logging.getLogger("manuscript2slides")


# region setup_console_encoding
def setup_console_encoding() -> None:
    """Configure UTF-8 encoding for Windows console to prevent UnicodeEncodeError when printing non-ASCII characters (like emojis)."""
    if platform.system() == "Windows":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")


# endregion


# region open_folder_in_os_explorer
def open_folder_in_os_explorer(folder_path: Path | str) -> None:
    """
    Open the folder in the system file explorer, platform-specific.

    Args:
        folder_path: Path to the folder to open
    """
    try:
        folder_path = Path(folder_path)  # Convert to Path if string
        system = platform.system()

        if system == "Windows":
            subprocess.run(["explorer", str(folder_path)])
        elif system == "Darwin":  # macOS
            subprocess.run(["open", str(folder_path)])
        else:  # Linux and others
            subprocess.run(["xdg-open", str(folder_path)])

        log.info(f"Opened folder: {folder_path}")

    except Exception as e:
        log.error(f"Failed to open folder: {folder_path}: {e}")


# endregion
