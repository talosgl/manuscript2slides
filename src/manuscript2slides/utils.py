"""Utilities for use across the entire program."""

import io
import logging
import os
import platform
import subprocess
import sys
from pathlib import Path

from manuscript2slides.internals import constants

log = logging.getLogger("manuscript2slides")


# region setup_console_encoding
def setup_console_encoding() -> None:
    """Configure UTF-8 encoding for Windows console to prevent UnicodeEncodeError when printing non-ASCII characters (like emojis)."""
    if platform.system() == "Windows":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")


# endregion


# region get_debug_mode
def get_debug_mode() -> bool:
    """Determine debug mode by checking whether there's an env variable set; otherwise fallback to bool constant."""

    # 1. Check env variable
    env_debug_str = os.environ.get("MANUSCRIPT2SLIDES_DEBUG")
    if env_debug_str is not None:
        try:
            # If a valid value is found, return it immediately
            return str_to_bool(env_debug_str)
        except ValueError:
            # If the env var is set but invalid ("bob"), log a warning and fall through to default
            log.warning(
                f"Warning: Invalid value for MANUSCRIPT2SLIDES_DEBUG env var: '{env_debug_str}'. Using default."
            )

    # 2. Lowest Priority / Fallback: The system default constant
    return constants.DEBUG_MODE_DEFAULT


# endregion


# region str_to_bool
def str_to_bool(value: str) -> bool:
    """Convert strings "True"/"False" to  booleans"""
    if value.lower().strip() in {"false", "f", "0", "no", "n"}:
        return False
    elif value.lower().strip() in {"true", "t", "1", "yes", "y"}:
        return True
    else:
        log.warning(f"{value} is not a valid boolean value.")
        raise ValueError(f"{value} is not a valid boolean value.")


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
