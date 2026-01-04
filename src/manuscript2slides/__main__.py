"""Entry point for manuscript2slides desktop application."""

from __future__ import annotations

import logging

from manuscript2slides import startup
from manuscript2slides.gui import run as run_gui


def main() -> None:
    """GUI entry point for source code development.

    In dev, you can call the app any of these ways:
        # Launch GUI
        python -m manuscript2slides
        python -m manuscript2slides.gui

        # Launch CLI
        python -m manuscript2slides.cli

    After pip install, use the following (and see pyproject.toml):
        manuscript2slides        # Launch GUI
        manuscript2slides-cli    # Launch CLI
    """

    # Set up logging and user folder scaffold.
    log: logging.Logger = startup.initialize_application()

    # A hello-world to probably remove someday, but I am sentimental. :)
    log.info("Hello, manuscript parser!")

    try:
        run_gui()
    except Exception:
        log.exception("Unhandled exception - program crashed.")  # Logs full traceback
        raise  # Still crash, but now it's logged


if __name__ == "__main__":
    main()
