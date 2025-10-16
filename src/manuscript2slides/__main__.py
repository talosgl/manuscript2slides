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
    python -m manuscript2slides # GUI will launch if no args passed in
    python -m manuscript2slides --help # CLI because an arg is passed in
    ```

    """

    # Set up logging and user folder scaffold.
    log: logging.Logger = startup.initialize_application()

    # A hello-world to probably remove someday, but I am sentimental. :)
    log.info("Hello, manuscript parser!")

    # Route to the appropriate interface (CLI vs GUI based on command-line args)
    # NOTE: For testing CLI without any args, just pass in `--cli`
    if "--cli" in sys.argv or any(arg.startswith("--") for arg in sys.argv[1:]):
        run_cli()  # Any CLI flags = CLI mode
    else:
        # Only import the GUI stuff if we're going to use the GUI
        from manuscript2slides.gui import run as run_gui

        # GUI mode: no arguments, launch GUI
        run_gui()


if __name__ == "__main__":

    main()
