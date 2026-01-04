"""Startup logic needed by both CLI and GUI interfaces before anything else happens.

Handles common setup tasks required by both CLI and GUI interfaces:
- Logging configuration
- User directory scaffolding (templates, output folders)
- Some CLI-specific setup that is harmless to GUI (console encoding setup)
"""

import logging
import sys
from pathlib import Path

from manuscript2slides.internals.logger import setup_logger
from manuscript2slides.internals.paths import user_base_dir
from manuscript2slides.internals.scaffold import ensure_user_scaffold, user_log_dir_path
from manuscript2slides.utils import get_debug_mode, setup_console_encoding


# region initialize_application
def initialize_application() -> logging.Logger:
    """Common startup tasks for both CLI and GUI."""

    # Windows console encoding must be set before any console output
    # This is CLI-specific setup, but it is harmless to call for GUI,
    # and we need to call this prior to setting up the logger.
    setup_console_encoding()

    # Start up logging.

    # Compute expected log path for error messages (before trying to create it)
    expected_log_dir = user_base_dir() / "logs"

    try:
        log = setup_logger(enable_trace=_should_enable_trace_on_startup())
    except PermissionError as e:
        print(
            f"ERROR: Cannot create log files. Check permissions for:", file=sys.stderr
        )
        print(f"  {expected_log_dir}", file=sys.stderr)
        raise SystemExit(1) from e
    except OSError as e:
        print(
            f"ERROR: Cannot write log files (disk full or I/O error).", file=sys.stderr
        )
        raise SystemExit(1) from e

    log.info("Starting manuscript2slides Log.")

    # Ensure user folders exist and templates are copied
    log.debug(
        "Checking for existing manuscripts2slides user folders and scaffolding if needed."
    )

    try:
        ensure_user_scaffold()
    except PermissionError as e:
        log.error("Cannot create user directories. Check permissions.")
        raise SystemExit(1) from e
    except OSError as e:
        log.error(f"Failed to create user scaffold: {e}")
        raise SystemExit(1) from e

    return log


# endregion


# region _should_enable_trace_on_startup
def _should_enable_trace_on_startup() -> bool:
    """
    Determine if trace logging should start immediately based on Debug Mode switch.

    Checks:
    - Environment variable (MANUSCRIPT2SLIDES_DEBUG)
    - System default (DEBUG_MODE_DEFAULT in constants.py)

    Later we may add other ways to enable this like a --verbose CLI arg.
    """
    return get_debug_mode()


# endregion
