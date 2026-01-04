"""
Basic logging setup; creates console and file handlers with session_id in every log line.
"""

import logging

from manuscript2slides.internals.paths import user_log_dir_path
from manuscript2slides.internals.run_context import get_session_id


# region setup_logger
def setup_logger(
    name: str = "manuscript2slides",
    level: int = logging.DEBUG,
    enable_trace: bool = False,
) -> logging.Logger:
    """
    Setup logging with console and file output.

    The session_id is included in every log line for traceability.
    Safe to call multiple times (won't create duplicate handlers).

    Args:
        name: Logger name (default: "manuscript2slides")
        level: Minimum log level (default: DEBUG)

    Returns:
        Configured logger instance

    Example:
        >>> log = setup_logger()
        >>> log.info("Starting conversion")
        2025-01-09 14:23:45 [INFO] Starting conversion [run:a1b2c3d4]
    """

    logger = logging.getLogger(name)

    if not logger.handlers:
        # First-time setup: console + basic file handlers
        logger.setLevel(level)

        # Don't pass logs up to parent loggers, like Python's root logger.
        # Why: If you have other libraries that log, you don't want their logs mixed with yours. This keeps "manuscript2slides" logs separate.
        logger.propagate = False

        # Get the session_id once for this logger setup
        session_id = get_session_id()

        # Create formatters (same format for both console and file)
        log_format = f"%(asctime)s [%(levelname)s] %(message)s [session:{session_id}]"
        formatter = logging.Formatter(log_format, datefmt="%Y-%m-%d %H:%M:%S")

        # Console handler (prints to terminal)
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        console_handler.setLevel(logging.INFO)  # Filter to less verbose for console
        logger.addHandler(console_handler)

        # File handler (writes to ~/Documents/manuscript2slides/logs/manuscript2slides.log)
        log_dir = user_log_dir_path()
        log_dir.mkdir(parents=True, exist_ok=True)  # Ensure log directory exists
        log_file = log_dir / "manuscript2slides.log"
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.DEBUG)  # Everything goes to file
        logger.addHandler(file_handler)

        logger.info(f"Logger initialized. Writing to {log_file}")

        # Add trace handler if requested and not already present
        if enable_trace and not _has_trace_handler(logger):
            _add_trace_handler(logger)

    # Add trace handler if requested and not already present
    if enable_trace and not _has_trace_handler(logger):
        _add_trace_handler(logger)

    return logger


# endregion


# region _has_trace_handler
def _has_trace_handler(logger: logging.Logger) -> bool:
    """Check if the logger already has a trace handler."""
    trace_log_file_name = user_log_dir_path() / "trace_manuscript2slides.log"
    for handler in logger.handlers:
        if isinstance(handler, logging.FileHandler):
            if handler.baseFilename == str(trace_log_file_name):
                return True
    return False


# endregion


# region _add_trace_handler
def _add_trace_handler(logger: logging.Logger) -> None:
    """Add trace file handler to existing logger."""
    session_id = get_session_id()

    trace_log_format = f"%(filename)s: %(funcName)s(), Line: %(lineno)d: - [%(levelname)s] %(asctime)s - %(message)s -- [session_id={session_id}]"
    trace_log_formatter = logging.Formatter(
        trace_log_format, datefmt="%Y-%m-%d %H:%M:%S"
    )
    log_dir = user_log_dir_path()
    log_dir.mkdir(parents=True, exist_ok=True)  # Ensure log directory exists
    trace_log_file = log_dir / "trace_manuscript2slides.log"
    trace_file_handler = logging.FileHandler(trace_log_file, encoding="utf-8")
    trace_file_handler.setFormatter(trace_log_formatter)
    trace_file_handler.setLevel(logging.DEBUG)
    logger.addHandler(trace_file_handler)
    logger.info(f"Trace logging enabled. Writing to {trace_log_file}")


# endregion
