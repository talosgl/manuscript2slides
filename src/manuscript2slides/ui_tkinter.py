"""Tkinter desktop UI for manuscript2slides."""

import tkinter as tk
from tkinter import messagebox
import logging
from pathlib import Path
import threading
import platform
import subprocess

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

        # Setup log viewer and handler
        self.setup_log_handler()

    # region create_widgets()
    def create_widgets(self) -> None:
        """Create all UI widgets."""

        # === Input Frame === #
        # LabelFrame = Frame + built-in title + border. Perfect for sections
        input_frame = tk.LabelFrame(
            self.root,
            text="Input File",
            padx=15,
            pady=10,
            relief="groove",  # Try: flat, raised, sunken, groove, ridge
            borderwidth=2,
            font=("Arial", 10, "bold"),
        )
        input_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)

        #### children of input_frame
        # File dispay/selection
        if self.selected_file:
            file_text = self.selected_file.name
        else:
            file_text = "No file selected"

        self.file_display = tk.Label(input_frame, text=file_text, fg="gray", bg="white")
        self.file_display.grid(row=0, column=0, sticky="w", padx=5)

        # browse button
        browse_btn = tk.Button(input_frame, text="Browse...", command=self.browse_file)
        browse_btn.grid(row=0, column=1, padx=5)

        # === Config Frame === #
        config_frame = tk.LabelFrame(self.root, text="Configuration", padx=15, pady=10)
        config_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)

        ### Config children
        # Direction
        direction_label = tk.Label(config_frame, text="Direction:")
        direction_label.grid(row=0, column=0, sticky="w", pady=5)

        docx2pptx_radio = tk.Radiobutton(
            config_frame,
            text="Word → PowerPoint",
            variable=self.direction_var,
            value=PipelineDirection.DOCX_TO_PPTX.value,  # Copying value
            command=self.update_chunk_dropdown,
        )
        docx2pptx_radio.grid(row=0, column=1, sticky="w", padx=(20, 0))

        pptx2docx_radio = tk.Radiobutton(
            config_frame,
            text="PowerPoint → Word",
            variable=self.direction_var,
            value=PipelineDirection.PPTX_TO_DOCX.value,
            command=self.update_chunk_dropdown,
        )
        pptx2docx_radio.grid(row=0, column=2, sticky="w", padx=(10, 0))

        # Chunk type
        chunk_label = tk.Label(config_frame, text="Split by:")
        chunk_label.grid(row=1, column=0, sticky="w", pady=5)

        self.chunk_dropdown = tk.OptionMenu(
            config_frame,
            self.chunk_var,
            *[chunk.value for chunk in ChunkType],
        )
        self.chunk_dropdown.grid(row=1, column=1, sticky="w", padx=(20, 0))

        # === Action Frame === #
        action_frame = tk.LabelFrame(self.root, text="Actions", padx=15, pady=10)
        action_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        ### action children
        # TODO, polish: make it so this button is disabled until selected_file is populated.
        # Convert button
        self.convert_btn = tk.Button(
            action_frame,
            text="Convert",
            command=self.on_convert_click,
            bg="green",
            fg="white",
            font=("Arial", 12, "bold"),
            width=15,
        )
        self.convert_btn.grid(row=0, column=0, padx=5)

        # Status label
        self.status_label = tk.Label(action_frame, text="Ready", fg="blue")
        self.status_label.grid(row=0, column=1, padx=20)

        # === Log Frame === #
        log_frame = tk.LabelFrame(self.root, text="Log Output", padx=10, pady=10)
        log_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=10)

        # Text widget with scrollbar (inside log_frame)
        text_frame = tk.Frame(log_frame)  # Sub-frame for text+scrollbar
        text_frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")

        self.log_text = tk.Text(
            text_frame,
            height=10,
            state="disabled",
            yscrollcommand=scrollbar.set,
            wrap="word",
            bg="#f0f0f0",
            font=("Courier", 9),
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar.config(command=self.log_text.yview)

        # === CONFIGURE GRID WEIGHTS === #
        # Make the window resizable properly
        self.root.columnconfigure(0, weight=1)  # Main column expands
        self.root.rowconfigure(3, weight=1)  # Log row expands

    # endregion

    def get_direction(self) -> PipelineDirection:
        """Get current direction from UI."""
        return PipelineDirection(self.direction_var.get())

    def get_chunk_type(self) -> ChunkType:
        """Get current chunk type from UI."""
        return ChunkType(self.chunk_var.get())

    # region browse_file
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

    # endregion

    def update_chunk_dropdown(self) -> None:
        """Enable/disable chunk dropdown based on direction."""
        if self.get_direction() == PipelineDirection.DOCX_TO_PPTX:
            self.chunk_dropdown.config(state="normal")  # enable / show
        else:
            self.chunk_dropdown.config(state="disabled")  # hide

    # region on_convert_click
    # === Threading Methods === #
    def on_convert_click(self) -> None:
        """Handle convert button click; start conversion in background."""
        log.info("Convert button clicked!")

        # Validate file selection
        if not self.selected_file:
            # TODO: show error dialog
            err_msg = "No file was selected"
            messagebox.showerror(
                err_msg, "Please select an input file before converting."
            )
            log.error(err_msg + ".")
            return

        # Validate file exists
        if not self.selected_file.exists():
            messagebox.showerror(
                "File Not Found",
                f"The selected file does not exist:\n\n{self.selected_file}",
            )
            log.error(f"File not found: {self.selected_file}")
            self.status_label.config(text="ERROR: File not found", fg="red")
            return

        # Build config from UI
        cfg = UserConfig()
        cfg.direction = self.get_direction()

        # store the cfg as an instance var so we can reference it later
        self.last_config = cfg

        # set input file based on direction
        if cfg.direction and cfg.direction == PipelineDirection.DOCX_TO_PPTX:
            cfg.input_docx = str(self.selected_file)
            cfg.chunk_type = self.get_chunk_type()
        else:
            cfg.input_pptx = str(self.selected_file)

        # Update the UI to show we're working
        self.status_label.config(text="Converting...", fg="blue")
        # also change the button
        self.convert_btn.config(state="disabled", bg="orange", text="Converting...")
        self.root.update()  # Force UI update

        # Start conversion thread in a background thread
        thread = threading.Thread(
            target=self.run_conversion_thread,
            args=(cfg,),
            daemon=True,  # Tells the thread to "die when the program exits"
        )
        thread.start()  # Starts the thread, then IMMEDIATELY returns out of this function so the UI isn't hung up

    # endregion

    # region run_conversion_thread
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

        # Get the output location
        if self.last_config:
            output_folder = self.last_config.get_output_folder()
        else:
            log.warning(
                "Couldn't access the last config; cannot determine output folder."
            )
            return

        # Show success dialog with output location
        input_name = (
            self.selected_file.name if self.selected_file else ""
        )  # For pylance
        message = (
            f"Conversion completed successfully!\n\n"
            f"Input: {input_name}\n"
            f"Output location:\n{output_folder}"
        )

        result = messagebox.askokcancel("Success", message + "\n\nOpen output folder?")

        # If user clicks OK, open the folder for them.
        if result:
            self.open_output_folder(output_folder)

    def on_conversion_error(self, error: Exception) -> None:
        """Called on UI thread when conversion fails."""
        self.status_label.config(text=f"ERROR: {str(error)}", fg="red")
        self.convert_btn.config(state="normal", bg="green", text="Convert")
        log.error(f"Conversion failed: {error}", exc_info=True)

        # prepare error for messagebox
        error_msg = str(error)
        # truncate very long errors
        if len(error_msg) > 300:
            error_msg = error_msg[:300] + "...\n\n(See log for full details)"

        messagebox.showerror(
            "Conversion Failed",
            f"An error occurred during conversion:\n\n{error_msg}\n\n"
            f"Check the log output below for more details.",
        )

    # endregion

    # region open_output_folder
    def open_output_folder(self, folder_path: Path) -> None:
        """
        Open the output folder in the system file explorer, platform-specific.

        Args:
            folder_path: Path to the folder to open
        """
        try:
            system = platform.system()

            if system == "Windows":
                # Windows: use 'explorer'
                subprocess.run(["explorer", str(folder_path)])
            elif system == "Darwin":  # macOS
                # macOS: use 'open'
                subprocess.run(["open", str(folder_path)])
            else:  # Linux and others
                # Linux: use 'xdg-open'
                subprocess.run(["xdg-open", str(folder_path)])

            log.info(f"Opened output folder: {folder_path}")

        except Exception as e:
            log.error(f"Failed to open folder: {e}")
            messagebox.showwarning(
                "Cannot Open Folder",
                f"Could not open the output folder automatically.\n\n"
                f"Location: {folder_path}",
            )

    # endregion

    # region setup_log_handler
    def setup_log_handler(self) -> None:
        """Connect the log viewer text widget to the logging system via our custom handler"""

        # We already got the logger at the top of the file, with log = logging.getLogger("manuscript2slides") after our imports,
        # but this is more readable for someone coming to the code later
        logger = logging.getLogger("manuscript2slides")

        text_handler = TextWidgetHandler(self.log_text, self.root)

        # format log messages
        formatter = logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s", datefmt="%H:%M:%S"
        )
        text_handler.setFormatter(formatter)

        # add our new handler to the logger
        logger.addHandler(text_handler)

        log.info("Log viewer initialized in tkinter UI")

    # endregion


# region Custom Log Handler
class TextWidgetHandler(logging.Handler):
    """Custom logging handler that writes to a Tkinter Text widget"""

    def __init__(self, text_widget: tk.Text, root: tk.Tk) -> None:
        """
        Initialize handler.

        Args:
            text_widget: The Text widget to write to
            root: The root Tk instance (for thread-safe updates)
        """
        super().__init__()  # I think this means yes, we're inheriting from logging.Handler. This is calling that guy's constructor to start init
        self.text_widget = text_widget
        self.root = root

    # This is called automatically by Python's logging system
    # whenever a log message is generated, because we made a subclass.
    # In fact it's required/expected that subclasses implement this method.
    # "This version is intended to be implemented by subclasses and so raises a NotImplementedError."

    def emit(self, record: logging.LogRecord) -> None:
        """Called by logging system when a log message is generated."""
        # Format the message
        msg = self.format(record=record)

        # Schedule the UI to update on the main thread, using after to be thread-safe
        # Critical: emit() might be called from a background thread; we use root.after() to safely update the UI.
        self.root.after(0, self._append_log, msg)

    def _append_log(self, message: str) -> None:
        """Append message to text widget (must run on UI thread)."""
        # Enable editing temporarily
        self.text_widget.config(state="normal")

        # Add the new message
        self.text_widget.insert("end", message + "\n")

        # Auto-scroll to bottom
        self.text_widget.see("end")

        # disable editing again to make it read-only
        self.text_widget.config(state="disabled")


# endregion


# region main
def main() -> None:
    """Tkinter UI entry point."""
    initialize_application()  # configure the log & other startup tasks

    log.info("Initializing Tkinter UI")
    root = tk.Tk()
    app = Manuscript2SlidesUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
# endregion
