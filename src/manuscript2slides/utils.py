"""Utilities for use across the entire program."""

import io

import logging
import platform
import sys

log = logging.getLogger("manuscript2slides")


# region Basic Utils
def setup_console_encoding() -> None:
    """Configure UTF-8 encoding for Windows console to prevent UnicodeEncodeError when printing non-ASCII characters (like emojis)."""
    if platform.system() == "Windows":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")


# endregion
