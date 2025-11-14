"""Entry point for manuscript2slides desktop application."""

from __future__ import annotations
import sys
from manuscript2slides import startup
from manuscript2slides.cli import run as run_cli
import logging


def main() -> None:
    """Application entry point - handles initialization and interface routing.

    Call like:
    ```
    python -m manuscript2slides # GUI will launch by default
    python -m manuscript2slides --cli # CLI will launch
    ```

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
