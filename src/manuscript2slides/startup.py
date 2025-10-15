"""Startup logic needed by both CLI and GUI interfaces before anything else happens.

Handles common setup tasks required by both CLI and GUI interfaces:
- Logging configuration
- User directory scaffolding (templates, output folders)
"""

import logging

from manuscript2slides.internals.constants import DEBUG_MODE
from manuscript2slides.internals.logger import setup_logger
from manuscript2slides.internals.scaffold import ensure_user_scaffold


def initialize_application() -> logging.Logger:
    """Common startup tasks for both CL and GUI."""
    # Start up logging
    log = setup_logger(enable_trace=DEBUG_MODE)
    log.info("Starting manuscript2slides Log.")

    # A hello-world to probably remove someday, but I am sentimental. :)
    log.info("Hello, manuscript parser!")

    # Ensure user folders exist and templates are copied
    ensure_user_scaffold()

    return log
