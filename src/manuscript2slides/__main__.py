"""Entry point for manuscript2slides desktop application.

We use C-style single entry point architecture because this is a desktop
application, not a library. All initialization happens once in main()
(console encoding, logging, user folder setup) before routing to CLI or GUI.

This pattern ensures initialization always happens correctly and in the right
order, unlike library-style patterns where each module (cli.main(), gui.main())
would handle its own initialization independently.
"""

# TODO: I can't find a great resource to explain the two possible patterns above,
# so I will plan to write one up for https://github.com/talosgl/jojos-tech-guides/tree/main
# alongside the single-file/multi-file python anatomy article(s).

from __future__ import annotations
import sys
from manuscript2slides import startup
from manuscript2slides.cli import run as run_cli
import logging


def main() -> None:
    """Application entry point - handles initialization and interface routing."""

    # Set up logging and user folder scaffold.
    log: logging.Logger = startup.initialize_application()

    # A hello-world to probably remove someday, but I am sentimental. :)
    log.info("Hello, manuscript parser!")

    # Route to the appropriate interface (CLI vs GUI based on command-line args)
    # NOTE: For testing CLI without any args, just pass in a dummy arg, like so:
    # `python -m manuscript2slides --dummy-arg-so-that-CLI-will-run-in-dev`
    if len(sys.argv) > 1:
        # CLI mode
        run_cli()
    else:
        # Only import the GUI stuff if we're going to use the GUI
        from manuscript2slides.gui import run as run_gui

        # GUI mode: no arguments, launch GUI
        run_gui()


if __name__ == "__main__":

    main()
