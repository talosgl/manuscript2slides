"""Tkinter and ttk GUI interface entry point."""

from __future__ import annotations


import tkinter as tk
from tkinter import ttk
import sv_ttk
import darkdetect

from tkinter import scrolledtext
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
import sys


log = logging.getLogger("manuscript2slides")


# region MainWindow
class MainWindow(tk.Tk):
    """Main UI Window application class for manuscript2slides."""

    def __init__(self) -> None:
        """Constructor for the Main Window UI."""
        super().__init__()  # Initialize tk.Tk

        # Set up window
        self.title("manuscript2slides")
        self.geometry("600x500")
        self.minsize(400, 300)

        # Apply theme BEFORE creating widgets
        self._apply_theme()

        # Build the UI
        self._create_widgets()

    # region MainWindow_create_widgets()
    def _create_widgets(self) -> None:
        """Create all UI widgets in turn, calling components and their constructors as needed."""

        # Create and pack the notebook (tab container)
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=5, pady=5)

        # Create the log_viewer first, so we can pass to the other widgets, but wait to pack it into the UI geo.
        log_viewer = LogViewer(self)

        docx2pptx_tab = Docx2PptxTab(notebook, log_viewer)
        notebook.add(docx2pptx_tab, text="DOCX → PPTX")

        pptx2docx_tab = Pptx2DocxTab(notebook, log_viewer)
        notebook.add(pptx2docx_tab, text="PPTX → DOCX")

        demo_tab = DemoTab(notebook, log_viewer)
        notebook.add(demo_tab, text="DEMO")

        # Add the log_viewer to the end of the UI Geo.
        log_viewer.pack(fill="both", expand=True, padx=10, pady=10)

        # Configure grid weights (for resizing)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

    # endregion

    def _apply_theme(self) -> None:
        """Apply a modern(ish) theme to the UI."""
        sv_ttk.set_theme(darkdetect.theme())  # type:ignore
        style = ttk.Style()

        available_themes = style.theme_names()
        log.debug(f"Available themes: {available_themes}")

        # Try to use the best theme for the platform
        if "aqua" in available_themes:  # macOS
            style.theme_use("aqua")
            log.info("Using 'aqua' theme.")
        # elif "winnative" in available_themes:  # Windows '95
        #     style.theme_use("winnative")
        #     log.info("Usting 'winnative' theme.")
        elif "vista" in available_themes:  # Windows
            style.theme_use("vista")
            log.info("Usting 'vista' theme.")
        elif "clam" in available_themes:  # Linux/cross-plat
            style.theme_use("clam")
            log.info("Using 'clam' theme.")
        else:
            style.theme_use("default")
            log.info("Using 'default' theme.")

        # (Optional TODO) Customize specific elements


# endregion


# ============= Tab Classes =============
# region Docx2PptxTab class
class Docx2PptxTab(ttk.Frame):
    """UI Tab for the docx2pptx pipeline."""

    def __init__(self, parent: tk.Widget, log_viewer: LogViewer) -> None:
        """Constructor for docx2pptx Tab"""
        super().__init__(parent)
        self.log_viewer = log_viewer  # Store reference so we can write to it

        self._create_widgets()

    # Styled as _private; this is definitely internal setup.
    def _create_widgets(self) -> None:

        # Get defaults from backend
        default_cfg = UserConfig()

        # === IO Section ===
        io_section = ttk.LabelFrame(self, text="Input/Output Selection")
        io_section.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        # Input file
        self.input_selector = PathSelector(
            io_section, "Input .docx File:", filetypes=[("Word Documents", "*.docx")]
        )
        self.input_selector.grid(row=0, column=0, sticky="ew", pady=2)

        # Advanced (collapsible)
        advanced = CollapsibleFrame(io_section, title="Advanced")
        advanced.grid(row=1, column=0, sticky="ew", pady=5)

        self.output_selector = PathSelector(
            advanced.content_frame,
            "Output Folder:",
            is_dir=True,
            default=str(default_cfg.get_output_folder()),
        )
        self.output_selector.pack(fill="x", pady=2)

        self.template_selector = PathSelector(
            advanced.content_frame,
            "Custom Template:",
            filetypes=[("PowerPoint", "*.pptx")],
            default=str(default_cfg.get_template_pptx_path()),
        )
        self.template_selector.pack(fill="x", pady=2)

        # Configure column weights so widgets stretch
        self.columnconfigure(0, weight=1)
        io_section.columnconfigure(0, weight=1)

        # TODO: Create Options Frame
        # TODO: Create Basic Options frame
        # TODO: Chunk Type Dropdown/combobox
        # TODO: Experimental Formatting bool
        # TODO: Advanced Options Frame (collapsible)
        # TODO: Create Annotations master toggle & 3 child bools
        # TODO: create SaveLoadConfig with on_save=set_config, on_load=get_config
        # TODO: Create ActionFrame for convert button
        pass

    # We could make this _private since it is only called inside this class,
    # but conventionally callbacks are usually public in python.
    def on_convert_click(self) -> None:
        """Handle convert button click."""
        # TODO: Prepare data for us to call the pipeline with by performing basic validation of selected options,
        # building a UserConfig object from valid UI selections, and starting a background thread for the pipeline
        # to be run on.

        pass

    def _run_conversion_thread(self, cfg) -> None:
        """Run the conversion in a background thread."""

        pass

    def ui_to_config(self):
        # TODO: gather UI values into UserConfig
        pass

    def config_to_ui(self):
        # TODO: Populate UI values from UserConfig
        pass


# endregion


# region Pptx2DocxTab
class Pptx2DocxTab(ttk.Frame):
    """Tab frame for the Pptx2Docx Pipeline."""

    def __init__(self, parent, log_viewer) -> None:
        super().__init__(parent)
        self.log_viewer = log_viewer

        # Get defaults from backend
        default_cfg = UserConfig()
        self.default_template = str(default_cfg.get_template_docx_path())
        self._create_widgets()

    def _create_widgets(self) -> None:
        # TODO: Create IO Frame
        # TODO: Create Options Frame
        # TODO: Create SaveLoadConfig with on_save=set_config, on_load=get_config
        # TODO: Create ActionFrame for Convert button
        pass


# endregion


# region DemoTab
class DemoTab(ttk.Frame):
    """Tab for running demo dry-runs"""

    def __init__(self, parent, log_viewer) -> None:
        super().__init__(parent)
        self.log_viewer = log_viewer
        self._create_widgets()

    def _create_widgets(self) -> None:
        # TODO: Demo Selection Frame
        # Button: "DOCX → PPTX Demo" (command=self.run_docx2pptx_demo)
        # Button: "PPTX → DOCX Demo" (command=self.run_pptx2docx_demo)
        # Button: "Round-trip Demo" (command=self.run_roundtrip_demo)
        # Separator
        # Button: "Load Config & Run" (command=self.run_custom_demo)
        pass

    def run_docx2pptx_demo(self) -> None:
        pass

    def run_pptx2docx_demo(self) -> None:
        pass

    def run_roundtrip_demo(self) -> None:
        pass

    def run_custom_demo(self) -> None:
        # TODO: Browse for .toml, load it, run pipeline
        pass


# endregion

# =============== Component Classes ============ #


class CollapsibleFrame(ttk.Frame):
    """A frame that can be collapsed/expanded with a toggle button."""

    def __init__(
        self, parent: tk.Widget, title: str, start_collapsed: bool = True
    ) -> None:
        super().__init__(parent)

        self.title = title
        self.is_collapsed = start_collapsed

        # Toggle button with arrow
        arrow = "▶" if start_collapsed else "▼"
        self.toggle_btn = ttk.Button(
            self, text=f"{arrow} {title}", command=self.toggle, style="Collapse.TButton"
        )
        self.toggle_btn.pack(fill="x", padx=5, pady=2)

        # Content frame (for child widgets)
        self.content_frame = ttk.Frame(self)
        if not start_collapsed:
            self.content_frame.pack(fill="both", expand=True, padx=10, pady=5)

    def toggle(self) -> None:
        """Toggle between collapsed and expanded states."""
        if self.is_collapsed:
            # Expand
            self.content_frame.pack(fill="both", expand=True, padx=10, pady=5)
            self.toggle_btn.config(text=self.toggle_btn.cget("text").replace("▶", "▼"))
            self.is_collapsed = False
        else:
            # Collapse
            self.content_frame.pack_forget()
            self.toggle_btn.config(text=self.toggle_btn.cget("text").replace("▼", "▶"))
            self.is_collapsed = True


class PathSelector(ttk.Frame):
    """Shared file/directory path selector component."""

    def __init__(
        self,
        parent: tk.Widget,
        label: str,
        is_dir: bool = False,
        filetypes: list[tuple[str, str]] | None = None,
        default: str | None = None,
    ) -> None:
        super().__init__(parent)

        self.label_text = label
        self.is_dir = is_dir
        self.filetypes = filetypes  # Ignored if is_dir=True
        self.default = default

        # TODO: Study - this is data binding-- use this example to understand that UI concept better.
        self.selected_path = tk.StringVar(value=default or "No selection")

        self._create_widgets()

    def _create_widgets(self):
        # Label showing what this selector is for
        label = ttk.Label(self, text=self.label_text)
        label.grid(row=0, column=0, sticky="w", padx=5)

        # Entry showing the selected path
        self.path_entry = ttk.Entry(
            self, textvariable=self.selected_path, width=40, state="readonly"
        )
        self.path_entry.grid(row=0, column=1, sticky="ew", padx=5)

        # Browse button
        browse_btn = ttk.Button(self, text="Browse...", command=self.browse)
        browse_btn.grid(row=0, column=2, padx=(5, 0))

        # Make entry stretch with window
        self.columnconfigure(1, weight=1)

    def browse(self):
        """Open file or directory dialog based on is_dir flag."""

        # Get initial directory from default (extract dir from file path if needed)
        initial_dir = None
        if self.default:
            default_path = Path(self.default)
            if default_path.exists():
                # If default is a file, use its parent directory
                initial_dir = (
                    str(default_path.parent)
                    if default_path.is_file()
                    else str(default_path)
                )

        if self.is_dir:
            path = filedialog.askdirectory(
                title=f"Select {self.label_text}", initialdir=initial_dir
            )
        else:
            path = filedialog.askopenfilename(
                title=f"Select {self.label_text}",
                filetypes=self.filetypes if self.filetypes else [("All files", "*.*")],
                initialdir=initial_dir,
            )

        if path:
            self.selected_path.set(path)


class SaveLoadConfig(ttk.Frame):
    pass


class ActionFrame(ttk.Frame):
    pass


# =============== Log Viewer ============ #


# region LogViewer
class LogViewer(ttk.LabelFrame):
    """The LogViewer Frame which we want visible on all tabs."""

    def __init__(self, parent: tk.Tk) -> None:
        super().__init__(parent, text="Log Viewer")
        self.root = parent

        self._create_widgets()
        self._setup_log_handler()

    def _create_widgets(self) -> None:
        """Create the text widget and clear button."""
        self.text_widget = scrolledtext.ScrolledText(
            self,
            height=10,
            state="disabled",
            wrap="word",
            bg="#f0f0f0",
            font=("Courier", 9),
        )
        self.text_widget.pack(fill="both", expand=True, padx=5, pady=5)

        self.clear_btn = ttk.Button(self, text="Clear Log", command=self.clear_log)
        self.clear_btn.pack(side="left", pady=(0, 5))

    def clear_log(self) -> None:
        """Clear all text from the log viewer."""
        self.text_widget.config(state="normal")  # enable editing
        # Tkinter's Text widget uses string indices. "1.0" = line 1, character 0 (the very start)
        self.text_widget.delete("1.0", "end")
        self.text_widget.config(state="disabled")

    # region setup_log_handler
    def _setup_log_handler(self) -> None:
        """Connect the log viewer text widget to the logging system via our custom handler"""

        # We already got the logger at the top of the file, with log = logging.getLogger("manuscript2slides") after our imports,
        # but this is more readable for someone coming to the code later
        logger = logging.getLogger("manuscript2slides")

        text_handler = TextWidgetHandler(self.text_widget, self.root)

        # format log messages
        formatter = logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s", datefmt="%H:%M:%S"
        )
        text_handler.setFormatter(formatter)

        # add our new handler to the logger
        logger.addHandler(text_handler)

        log.info("Log viewer initialized in tkinter UI")

    # endregion


# endregion


# region TextWidgetHandler extending logging.Handler
class TextWidgetHandler(logging.Handler):
    """Custom logging handler that writes to a Tkinter Text widget"""

    def __init__(self, text_widget: tk.Text, root: tk.Tk) -> None:
        """
        Initialize handler.

        Args:
            text_widget: The Text widget to write to
            root: The root Tk instance (for thread-safe updates)
        """
        super().__init__()  # We're inheriting from logging.Handler. This is calling that guy's constructor to start init
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


# region Helper functions (class agnostic)


def show_completion_dialog(output_folder):
    """Helper to pop a message box for any tab that's completed a pipeline run & offer a button for the user to click if they want us to open the output folder for them."""
    pass


def open_folder_in_os_explorer(folder_path):
    """Helper to open the output folder for a pipeline run for the user."""
    pass


# endregion


# region main
def main() -> None:
    """Tkinter UI entry point."""
    initialize_application()  # configure the log & other startup tasks

    log.info("Initializing Tkinter UI")
    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
# endregion
