"""Startup logic needed by both CLI and GUI interfaces before anything else happens.

Handles common setup tasks required by both CLI and GUI interfaces:
- Logging configuration
- User directory scaffolding (templates, output folders)
- Some CLI-specific setup that is harmless to GUI (console encoding setup)
"""

import logging

from manuscript2slides.internals.constants import DEBUG_MODE
from manuscript2slides.internals.logger import setup_logger
from manuscript2slides.internals.scaffold import ensure_user_scaffold

from manuscript2slides.utils import setup_console_encoding


def initialize_application() -> logging.Logger:
    """Common startup tasks for both CLI and GUI."""

    # Windows console encoding must be set before any console output
    # This is CLI-specific setup, but it is harmless to call for GUI,
    # and we need to call this prior to setting up the logger.
    setup_console_encoding()

    # Start up logging
    log = setup_logger(enable_trace=DEBUG_MODE)
    log.info("Starting manuscript2slides Log.")

    # Ensure user folders exist and templates are copied
    log.debug(
        "Checking for existing manuscripts2slides user folders and scaffolding if needed."
    )
    ensure_user_scaffold()

    return log
