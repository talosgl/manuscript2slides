"""Entry point for manuscript2slides desktop application."""

from __future__ import annotations

import logging
import sys

from manuscript2slides import startup
from manuscript2slides.cli import run as run_cli


def main() -> None:
    """Optional source code development entry point; defaults to GUI.

    In dev, you can call the app any of these ways:
        # Launch GUI
        python -m manuscript2slides
        python -m manuscript2slides.gui (skips __main__.py)

        # Launch CLI
        python -m manuscript2slides --cli
        python -m manuscript2slides.cli (skips __main__.py)

    After pip install, use the following (and see pyproject.toml):
        manuscript2slides        # Launch GUI (skips __main__.py)
        manuscript2slides-cli    # Launch CLI (skips __main__.py)
    """

    # Set up logging and user folder scaffold.
    log: logging.Logger = startup.initialize_application()

    # A hello-world to probably remove someday, but I am sentimental. :)
    log.info("Hello, manuscript parser!")

    try:
        # Route to the appropriate interface
        if "--cli" in sys.argv:
            run_cli()
        else:
            # Only import the GUI stuff if we're going to use the GUI
            from manuscript2slides.gui import run as run_gui

            # GUI mode: no arguments, launch GUI
            run_gui()
    except Exception:
        log.exception("Unhandled exception - program crashed.")  # Logs full traceback
        raise  # Still crash, but now it's logged


if __name__ == "__main__":
    main()
