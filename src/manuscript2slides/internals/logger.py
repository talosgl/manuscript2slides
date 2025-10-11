"""
Basic logging setup; creates console and file handlers with run_id in every log line.
"""

import logging

from manuscript2slides.internals.paths import user_log_dir_path
from manuscript2slides.internals.run_context import get_run_id


def setup_logger(
    name: str = "manuscript2slides",
    level: int = logging.DEBUG,
    enable_trace: bool = False,
) -> logging.Logger:
    """
    Setup logging with console and file output.

    The run_id is included in every log line for traceability.
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

    # If it's already configured, return the existing logger
    if logger.handlers:
        return logger

    logger.setLevel(level)

    # Don't pass logs up to parent loggers, like Python's root logger.
    # Why: If you have other libraries that log, you don't want their logs mixed with yours. This keeps "manuscript2slides" logs separate.
    logger.propagate = False

    # Get the run_id once for this logger setup
    run_id = get_run_id()

    # Create formatters (same format for both console and file)
    log_format = f"%(asctime)s [%(levelname)s] %(message)s [run:{run_id}]"
    formatter = logging.Formatter(log_format, datefmt="%Y-%m-%d %H:%M:%S")

    # Console handler (prints to terminal)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    console_handler.setLevel(logging.INFO)  # Filter to less verbose for console
    logger.addHandler(console_handler)

    # File handler (writes to ~/Documents/manuscript2slides/logs/manuscript2slides.log)
    log_file = user_log_dir_path() / "manuscript2slides.log"
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.DEBUG)  # Everything goes to file
    logger.addHandler(file_handler)

    # Trace log file handler
    if enable_trace:
        # Putting this here for later, maybe...
        trace_log_format = f"%(filename)s: %(funcName)s(), Line: %(lineno)d: - [%(levelname)s] %(asctime)s - %(message)s -- [run_id={run_id}]"
        trace_log_formatter = logging.Formatter(
            trace_log_format, datefmt="%Y-%m-%d %H:%M:%S"
        )
        trace_log_file = user_log_dir_path() / "trace_manuscript2slides.log"
        trace_file_handler = logging.FileHandler(trace_log_file, encoding="utf-8")
        trace_file_handler.setFormatter(trace_log_formatter)
        trace_file_handler.setLevel(logging.DEBUG)
        logger.addHandler(trace_file_handler)

    logger.info(f"Logger initialized. Writing to {log_file}")

    return logger
