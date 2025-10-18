"""Tkinter desktop UI for manuscript2slides."""

import tkinter as tk
import logging
from pathlib import Path
import threading

from tkinter import filedialog

from manuscript2slides.startup import initialize_application
from manuscript2slides.internals.config.define_config import (
    UserConfig,
    PipelineDirection,
    ChunkType,
)
from manuscript2slides.orchestrator import run_pipeline

log = logging.getLogger("manuscript2slides")


class Manuscript2SlidesUI:
    """Main UI application class."""

    def __init__(self, root: tk.Tk) -> None:
        """Initialize the UI."""
        self.root = root

        # Set up UI State

        # Get defaults from UserConfig
        cfg_defaults = UserConfig()

        try:
            self.selected_file: Path | None = cfg_defaults.get_input_docx_file()
            log.info(f"Loaded default file: {self.selected_file.name}")
        except Exception as e:
            self.selected_file = None
            log.warning(f"No default file available: {e}")
            log.info(
                "We won't be able to run a sample dry run, but you should still be able to select your own file for input."
            )

        # TODO: Add all the other fields from UserConfig, later, in a collapsible "Advanced" section. Later.
        """
        # Add a button to show/hide advanced options
        advanced_btn = tk.Button(
            self.root,
            text="▶ Advanced Options",
            command=self.toggle_advanced
        )

        # Then in toggle_advanced(), show/hide a frame with checkboxes:
        self.advanced_frame = tk.Frame(self.root)
        # ... add checkboxes for all the bools ...
        """
        # Create StringVars for widgets
        self.direction_var = tk.StringVar(value=cfg_defaults.direction.value)
        self.chunk_var = tk.StringVar(value=cfg_defaults.chunk_type.value)

        # Set up window
        self.root.title("manuscript2slides")
        self.root.geometry("600x400")
        self.root.minsize(400, 300)

        # Build the UI
        self.create_widgets()

    def create_widgets(self) -> None:
        """Create all UI widgets."""

        # === FILE SELECTION === #

        # File Select verb label
        file_label = tk.Label(
            self.root,
            text="Select an input file, or use the default to do a dry run:",
        )
        file_label.grid(
            row=0, column=0, sticky="w", padx=10, pady=10
        )  # Sticky is "alignment" and uses NSEW

        # Selected file label
        if self.selected_file:
            file_text = self.selected_file.name  # Show sample filename
        else:
            file_text = "No file selected"

        self.file_display = tk.Label(self.root, text=file_text, fg="gray")
        self.file_display.grid(row=0, column=1, sticky="w", padx=10, pady=10)

        # browse
        browse_btn = tk.Button(self.root, text="Browse...", command=self.browse_file)
        browse_btn.grid(row=0, column=2, padx=10, pady=10)

        # === DIRECTION SELECTION === #
        direction_label = tk.Label(self.root, text="Pipeline direction:")
        direction_label.grid(row=1, column=0, sticky="w", padx=10, pady=10)

        docx2pptx_radio = tk.Radiobutton(
            self.root,
            text="Word to PowerPoint",
            variable=self.direction_var,
            value=PipelineDirection.DOCX_TO_PPTX.value,
            command=self.update_chunk_dropdown,  # add a callback so the rest of the UI updates when this radio changes
        )
        docx2pptx_radio.grid(row=1, column=1, sticky="w", padx=10, pady=5)

        pptx2docx_radio = tk.Radiobutton(
            self.root,
            text="PowerPoint to Word",
            variable=self.direction_var,
            value=PipelineDirection.PPTX_TO_DOCX.value,
            command=self.update_chunk_dropdown,
        )
        pptx2docx_radio.grid(row=2, column=1, sticky="w", padx=10, pady=0)

        # === CHUNK TYPE (only for docx2pptx) === #
        chunk_label = tk.Label(self.root, text="Split by:")
        chunk_label.grid(row=3, column=0, sticky="w", padx=10, pady=10)

        # Chunk dropdown
        self.chunk_dropdown = tk.OptionMenu(
            self.root,
            self.chunk_var,
            *[chunk.value for chunk in ChunkType],  # Generate list from enum
        )
        self.chunk_dropdown.grid(row=3, column=1, sticky="w", padx=10, pady=10)

        # === CONVERT BUTTON === #
        self.convert_btn = tk.Button(
            self.root,
            text="Convert",
            command=self.on_convert_click,
            bg="green",
            fg="white",
            font=("Arial", 12, "bold"),
        )
        self.convert_btn.grid(row=4, column=1, pady=20)

        # Status label
        self.status_label = tk.Label(self.root, text="", fg="blue")
        self.status_label.grid(row=4, column=0, pady=(0, 10))

        # Configure grid weights (makes things resize properly)
        self.root.columnconfigure(0, weight=1)

    def get_direction(self) -> PipelineDirection:
        """Get current direction from UI."""
        return PipelineDirection(self.direction_var.get())

    def get_chunk_type(self) -> ChunkType:
        """Get current chunk type from UI."""
        return ChunkType(self.chunk_var.get())

    def browse_file(self) -> None:
        """Open file browser dialog."""

        filetypes = [
            ("Word Documents", "*.docx"),
            ("PowerPoint Files", "*.pptx"),
            ("All Files", "*.*"),
        ]

        filename = filedialog.askopenfilename(
            title="Select input file", filetypes=filetypes
        )

        if filename:  # User selected a file (didn't cancel)
            self.selected_file = Path(filename)
            self.file_display.config(text=self.selected_file.name)
            log.info(f"File selected: {self.selected_file}")

    def update_chunk_dropdown(self) -> None:
        """Enable/disable chunk dropdown based on direction."""
        if self.get_direction() == PipelineDirection.DOCX_TO_PPTX:
            self.chunk_dropdown.config(state="normal")  # enable / show
        else:
            self.chunk_dropdown.config(state="disabled")  # hide

    # === Threading Methods === #
    def on_convert_click(self) -> None:
        """Handle convert button click; start conversion in background."""
        log.info("Convert button clicked!")

        if not self.selected_file:
            # TODO: show error dialog
            log.error(
                "No file was selected."
            )  # Q: Well shouldn't the button be disabled then? Lol
            return

        # Build config from UI
        cfg = UserConfig()
        cfg.direction = self.get_direction()

        # set input file based on direction
        if cfg.direction and cfg.direction == PipelineDirection.DOCX_TO_PPTX:
            cfg.input_docx = str(self.selected_file)
            cfg.chunk_type = self.get_chunk_type()
        else:
            cfg.input_pptx = str(self.selected_file)

        # Update the UI to show we're working
        self.status_label.config(text="Converting...", fg="blue")
        # TODO: Also change the button
        self.convert_btn.config(
            state="disabled", bg="yellow", text="Converting in Process..."
        )
        self.root.update()  # Force UI update

        # Start conversion thread in a background thread
        thread = threading.Thread(
            target=self.run_conversion_thread,
            args=(cfg,),
            daemon=True,  # Tells the thread to "die when the program exits"
        )
        thread.start()  # Starts the thread, then IMMEDIATELY returns out of this function so the UI isn't hung up

    def run_conversion_thread(self, cfg: UserConfig) -> None:
        """Run the convversion in a background thread."""  # this way, the UI thread is free to handle clicks, redraws, etc.
        if self.selected_file:  # For pylance's sake.
            log.info(f"Starting conversion: {self.selected_file.name}")
        try:
            cfg.validate()
            run_pipeline(cfg)

            # Notify UI of success
            self.root.after(ms=0, func=self.on_conversion_success)
            """
            Why root.after()?
            CRITICAL RULE: You cannot modify Tkinter widgets from a background thread. Tkinter is NOT thread-safe!
            root.after(0, function) says:
                "Schedule this function to run on the UI thread"
                "As soon as possible (0 milliseconds)"
            This is thread-safe - it's the ONLY safe way to communicate from background → UI in Tkinter.
            """

        except Exception as e:
            # Notify UI of error
            self.root.after(0, self.on_conversion_error, e)

    def on_conversion_success(self) -> None:
        """Called on UI thread when conversion succeeds."""
        self.status_label.config(text="✓ Conversion complete!", fg="green")
        self.convert_btn.config(state="normal", bg="green", text="Convert")
        log.info("Conversion complete!")

    def on_conversion_error(self, error: Exception) -> None:
        """Called on UI thread when conversion fails."""
        self.status_label.config(text=f"ERROR: {str(error)}", fg="red")
        self.convert_btn.config(state="normal")
        log.error(f"Conversion failed: {error}", exc_info=True)


def main() -> None:
    """Tkinter UI entry point."""
    initialize_application()  # configure the log & other startup tasks

    log.info("Initializing Tkinter UI")
    root = tk.Tk()
    app = Manuscript2SlidesUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
