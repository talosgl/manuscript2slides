"""Tkinter desktop UI for manuscript2slides."""

import tkinter as tk
from manuscript2slides.startup import initialize_application


def main():
    """Tkinter UI entry point."""

    # Do standard setup
    log = initialize_application()
    log.info("Running in Tkinter UI mode")

    # Create the main window object
    # Every Tkinter app has exactly one Tk() object - it's the "root" of our entire UI.
    root = tk.Tk()
    root.title("manuscript2slides")
    root.geometry("600x400")

    # === EXPERIMENT ZONE === #
    # Set a minimum size height and width
    root.minsize(400, 300)

    # Start maximized (!!)
    root.state("zoomed")
    # ======================= #

    # Start the event loop (program waits here)
    root.mainloop()

    # This line only run after the window is closed
    log.info("UI closed (I hope by the user!)")

    # TODO: Separate _run_id into: and _pipeline_run_id and _session_id
    # Q: I'm noticing, now that I'm looking at the logs for the UI, that the run_id is going to be the same for any UI session. It's not going to be
    # per-pipeline-run, it's going to be per-UI-run. When we were doing CLI, those were the same things, but not with UI. You could, presumably, leave
    # the UI open for days, and run it dozens of times, with the same run_id. Dangit!


if __name__ == "__main__":
    main()
