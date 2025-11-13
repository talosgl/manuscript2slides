"""Startup logic needed by both CLI and GUI interfaces before anything else happens.

Handles common setup tasks required by both CLI and GUI interfaces:
- Logging configuration
- User directory scaffolding (templates, output folders)
- Some CLI-specific setup that is harmless to GUI (console encoding setup)
"""

import logging

from manuscript2slides.internals.logger import setup_logger
from manuscript2slides.internals.scaffold import ensure_user_scaffold
from manuscript2slides.internals.config.define_config import get_debug_mode

from manuscript2slides.utils import setup_console_encoding


# region initialize_application
def initialize_application() -> logging.Logger:
    """Common startup tasks for both CLI and GUI."""

    # Windows console encoding must be set before any console output
    # This is CLI-specific setup, but it is harmless to call for GUI,
    # and we need to call this prior to setting up the logger.
    setup_console_encoding()

    # Start up logging.
    log = setup_logger(enable_trace=_should_enable_trace_on_startup())
    log.info("Starting manuscript2slides Log.")

    # Ensure user folders exist and templates are copied
    log.debug(
        "Checking for existing manuscripts2slides user folders and scaffolding if needed."
    )
    ensure_user_scaffold()

    return log


# endregion


# region _should_enable_trace_on_startup
def _should_enable_trace_on_startup() -> bool:
    """
    Determine if trace logging should start immediately.

    At startup, only checks:
    - Environment variable (MANUSCRIPT2SLIDES_DEBUG)
    - System default (DEBUG_MODE_DEFAULT in constants.py)
    (CLI args, config files, and GUI preferences are not available yet at this point.)

    Note: Once trace logging is enabled, it remains active for the entire session.
    Later changes to debug_mode via CLI/config/GUI won't disable already-active
    trace logging. This is intentonal. We do not want to turn off trace logging midway
    through a session.
    """
    return get_debug_mode()


# endregion
