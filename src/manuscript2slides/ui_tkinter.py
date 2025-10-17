"""Tkinter desktop UI for manuscript2slides."""

import tkinter as tk
from manuscript2slides.startup import initialize_application


def main() -> None:
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
    # root.state("zoomed")

    # ======================= #

    # === WIDGET ZONE === #

    # Create a label widget
    label = tk.Label(root, text="Button not yet clicked", fg="blue")

    # Make the widget visible with pack(), Tkinter's "geometry manager"
    # Put this widget in its parent" + "Stack it with other widgets"
    # Without `.pack()`, the widget exists but is invisible!
    label.pack(
        pady=20, side="bottom"
    )  # Add to the window with 20px padding top & bottom (on the y-axis)

    # Create a button widget
    counter = 0

    # Define what will happen when the button is clicked by providing a named, callable object (function) containing the logic
    def on_button_click() -> (
        None
    ):  # Q: I assume typically event handlers are going to have None as return typehint?
        """This function runs when button is clicked."""
        nonlocal counter
        counter += 1  # Q: pylance says this unbound and, yeah, it crashes. Why do we have this global thing right above?
        # I thought since it was defined outside the scope of this func def, it would be able to access it, like in JS. But, it can't.
        # I also only ever seem to see nested functions like this in JS/UI contexts, and it's always confusing AF. Why do people like it?
        # Maybe since it's defined right before it's going to be called, it's easier to find or something? Just feels very spaghetti-code
        # architecturally...

        label.config(text=f"Button clicked {counter} times")
        log.info(f"Click count: {counter}")

        log.info("Button was clicked!")
        print(
            "Button clicked!"
        )  # Extra print to the console for fun # NOTE: in vs code debugpy, this won't get printed till we exit the loop!

    button = tk.Button(
        root,
        text="Click me!",
        command=on_button_click,  # last arg is an example of "event binding" - we pass the function, not call it, and we're telling the function when this button is clicked, then call this function -- NOT When you create the button.
    )  # we're passing in the function, named, as an arg. Q: are we passing in a pointer to the func, or a "copy" of the logic? Doesn't really matter, just curious
    button.pack(pady=10)
    # NOTE: This isn't from the top of the window, it's from the last-added object.
    # There's semantic meaning encoded in the order we add the widgets; that ordering
    # doesn't look to be captured in data elsewhere.... Q: Am I wrong? Can you somehow query the
    # ordering of widgets? Maybe they're technically like in an ordered array/list.

    # ======================= #

    # === EXPERIMENT ZONE 2 === #

    label1 = tk.Label(root, text="First label", bg="red", fg="white")
    label1.pack(side="left")

    label2 = tk.Label(root, text="Second label", fg="pink")
    label2.pack(
        side="top"
    )  # Q this moves it "above" the previous label, but after the earlier ones, so I am curious in relation to "what" these are meaning.

    button1 = tk.Button(
        root, text="what if you click me INSTEAD??", command=on_button_click
    )  # NOTE: if I put this code above the func def, it gets mad, like in C
    button1.pack(side="bottom")
    # ======================== #

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
