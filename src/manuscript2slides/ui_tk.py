"""Tkinter and ttk GUI interface entry point."""

from __future__ import annotations


import tkinter as tk
from tkinter import ttk

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
from manuscript2slides.internals.constants import DEBUG_MODE
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
        # TODO: gets smooshed/covered if Advanced Options are expanded in docx2pptx tab; fix to make responsive.
        log_viewer.pack(fill="both", expand=True, padx=10, pady=10)

        # Configure grid weights (for resizing)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

    # endregion

    def _apply_theme(self) -> None:
        """Apply a modern(ish) theme to the UI."""
        self.style = ttk.Style()

        available_themes = self.style.theme_names()
        log.debug(f"Available themes: {available_themes}")

        # Try to use the best theme for the platform
        if "aqua" in available_themes:  # macOS
            # TODO: Test aesthetics on macOS and see if we need to do any workarounds
            self.style.theme_use("aqua")
            log.info("Using 'aqua' theme.")
        # elif "winnative" in available_themes:  # Windows '95
        #     style.theme_use("winnative")
        #     log.info("Usting 'winnative' theme.")
        # elif "vista" in available_themes:  # Windows
        #     style.theme_use("vista")
        #     log.info("Usting 'vista' theme.")
        elif "clam" in available_themes:  # Linux/cross-plat
            self.style.theme_use("clam")
            log.info("Using 'clam' theme.")

            self._fix_clam()

        else:
            self.style.theme_use("default")
            log.info("Using 'default' theme.")

        # (Optional TODO) Customize specific elements or give up and switch to PySide
        self.style.configure(
            "Convert.TButton",
            background="#4CAF50",  # Green
            foreground="white",
            font=("Arial", 11, "bold"),
            padding=10,
        )

        self.style.map(
            "Convert.TButton",
            background=[
                ("active", "#45a049"),  # Darker green on hover
                ("disabled", "#cccccc"),  # Gray when disabled
            ],
        )

    def _fix_clam(self) -> None:
        """Apply some styling fixes for clam theme if used."""

        log.debug("Applying combobox fix for clam for Linux or Windows.")
        # Configure Combobox to look better in clam
        self.style.configure(
            "TCombobox",
            fieldbackground="white",  # Background of the text field
            background="white",  # Background of the dropdown
            foreground="black",  # Text color
            selectbackground="#0078d7",  # Selected item background
            selectforeground="white",  # Selected item text
        )

        # Map states for better visual feedback
        self.style.map(
            "TCombobox",
            fieldbackground=[
                ("disabled", "#e0e0e0"),  # Gray when disabled
                ("readonly", "white"),  # White when enabled but readonly
            ],
            foreground=[
                ("disabled", "#808080"),  # Gray text when disabled
                ("readonly", "black"),  # Black text when enabled
            ],
        )

        # clam's main window background is not respected on Windows; this is a workaround
        # TODO: Test clam and the other possible windows themes on a few more PCs
        if platform.system() == "Windows":
            log.debug(
                "Windows only clam fix: set the background color to what it should be, explicitly. Use TButton for color lookup."
            )
            self.configure(bg=self.style.lookup("TButton", "background"))

        # Configure the collapse button to look like part of the TFrame, clam-only
        tframe_background = self.style.lookup("TFrame", "background")
        self.style.configure(
            "Collapse.TButton",
            # font=("Arial", 11, "bold"),
            background=tframe_background,
            borderwidth=0,
            relief="flat",
            anchor="w",
        )
        self.style.map(
            "Collapse.TButton",
            background=[
                ("active", tframe_background),
                ("pressed", tframe_background),
            ],
        )


# endregion

# region Tabs
# ============= Tab Classes =============
# endregion


# region Docx2PptxTab class
class Docx2PptxTab(ttk.Frame):
    """UI Tab for the docx2pptx pipeline."""

    def __init__(self, parent: tk.Widget, log_viewer: LogViewer) -> None:
        """Constructor for docx2pptx Tab"""
        super().__init__(parent)
        self.log_viewer = log_viewer  # Store reference so we can write to it
        self.loaded_config = None  # Store loaded config here
        self.last_run_config = (
            None  # Config actually used for last conversion (for finding output)
        )

        # Get defaults from UserConfig
        self.cfg_defaults = UserConfig()
        self.chunk_var = tk.StringVar(value=self.cfg_defaults.chunk_type.value)

        # BooleanVars for checkboxes
        self.exp_fmt_var = tk.BooleanVar(
            value=self.cfg_defaults.experimental_formatting_on
        )
        self.keep_metadata = tk.BooleanVar(
            value=self.cfg_defaults.preserve_docx_metadata_in_speaker_notes
        )
        self.keep_all_annotations = tk.BooleanVar(value=False)
        self.keep_comments = tk.BooleanVar(value=self.cfg_defaults.display_comments)
        self.keep_footnotes = tk.BooleanVar(value=self.cfg_defaults.display_footnotes)
        self.keep_endnotes = tk.BooleanVar(value=self.cfg_defaults.display_endnotes)

        # These feel excessive to have in the UI.
        # self.c_srt_by_date = tk.BooleanVar(value=self.cfg_defaults.comments_sort_by_date)
        # self.c_keep_authordate = tk.BooleanVar(value=self.cfg_defaults.comments_keep_author_and_date)

        self._create_widgets()

    # Styled as _private; this is definitely internal setup.
    def _create_widgets(self) -> None:

        self._create_io_section()

        self._create_basic_options()

        self._create_advanced_options()

        self._create_action_section()

    def _create_io_section(self) -> None:
        """Create docx2pptx tab's io section."""
        io_section = ttk.LabelFrame(self, text="Input/Output Selection")
        io_section.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        # Input file
        self.input_selector = PathSelector(
            io_section, "Input .docx File:", filetypes=[("Word Documents", "*.docx")]
        )
        self.input_selector.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=5,
            padx=5,
        )

        # Advanced (collapsible)
        advanced = CollapsibleFrame(io_section, title="Advanced")
        advanced.grid(row=1, column=0, sticky="ew", pady=5)

        self.output_selector = PathSelector(
            advanced.content_frame,
            "Output Folder:",
            is_dir=True,
            default=str(self.cfg_defaults.get_output_folder()),
        )
        self.output_selector.pack(fill="x", pady=2)

        self.template_selector = PathSelector(
            advanced.content_frame,
            "Custom Template:",
            filetypes=[("PowerPoint", "*.pptx")],
            default=str(self.cfg_defaults.get_template_pptx_path()),
        )
        self.template_selector.pack(fill="x", pady=2)

        # TODO: create SaveLoadConfig with on_save=set_config, on_load=get_config
        # try parenting to io_section first and seeing how that looks.

        io_section.columnconfigure(0, weight=1)

    def _create_basic_options(self) -> None:
        """Create pipeline options widgets."""
        # Create Basic Options frame
        options_frame = ttk.Labelframe(self, text="Basic Options")
        options_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)

        chunk_label = ttk.Label(options_frame, text="Chunk manuscript into slides by:")
        chunk_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        self.chunk_dropdown = ttk.Combobox(
            options_frame,
            textvariable=self.chunk_var,
            values=[chunk.value for chunk in ChunkType],
            state="readonly",  # Can't type custom values
            width=20,
        )
        self.chunk_dropdown.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        # Collapsible explanation
        explain_chunks = CollapsibleFrame(options_frame, title="What do these mean?")
        explain_chunks.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5, padx=5)
        explanation_text = (
            "PARAGRAPH (default): One slide per paragraph break.\n"
            "PAGE: One slide for every page break.\n"
            "Heading (Flat): New slides for every heading, regardless of parent-child hierarchy.\n"
            "Heading (Nested): New slides only on finding a 'parent/grandparent' heading to the previously found. \n"
            "All options create a new slide if there is a page break in the middle of a section."
        )
        ttk.Label(
            explain_chunks.content_frame, text=explanation_text, wraplength=500
        ).pack(padx=5, pady=5)

        exp_fmt_chk = ttk.Checkbutton(
            options_frame,
            text="Preserve advanced formatting (experimental)",
            variable=self.exp_fmt_var,  # ← Bind to BooleanVar,
        )
        exp_fmt_chk.grid(
            row=2, column=0, columnspan=2, sticky="w", padx=5, pady=(10, 0)
        )

        # Tip below checkbox (wraps automatically)
        tip_label = ttk.Label(
            options_frame,
            text="Tip: Disable this if conversion crashes or freezes",
            wraplength=400,  # Wraps at 400px
            foreground="gray",
        )
        tip_label.grid(row=3, column=0, columnspan=2, sticky="w", padx=25, pady=(0, 5))

        options_frame.columnconfigure(0, weight=0)
        options_frame.columnconfigure(1, weight=1)

    def _create_advanced_options(self) -> None:
        """Create advanced options (collapsible)."""
        # TODO: Advanced Options inner-frame (collapsible)
        advanced = CollapsibleFrame(
            self, title="Advanced Options", start_collapsed=True
        )
        advanced.grid(row=3, column=0, sticky="ew", padx=5, pady=5)

        # Metadata preservation
        keep_metadata_chk = ttk.Checkbutton(
            advanced.content_frame,
            text="Preserve metadata in speaker notes",
            variable=self.keep_metadata,
        )
        keep_metadata_chk.pack(anchor="w", padx=5, pady=(5, 2))

        ttk.Label(
            advanced.content_frame,
            text="Tip: Enable for round-trip conversion (maintains comments, heading formatting, etc.)",
            wraplength=400,
            foreground="gray",
        ).pack(anchor="w", padx=25, pady=(0, 10))

        # Annotations explanation
        ttk.Label(
            advanced.content_frame,
            text="Annotations cannot be replicated in slides, but can be copied into the slides' speaker notes.",
            wraplength=500,
        ).pack(anchor="w")

        # Checkboxes
        self.keep_all_annotations_chk = ttk.Checkbutton(
            advanced.content_frame,
            text="Keep all annotations",
            variable=self.keep_all_annotations,
            command=self._on_parent_annotation_toggle,  # When parent is clicked
        )
        self.keep_all_annotations_chk.pack(anchor="w")
        ttk.Checkbutton(
            advanced.content_frame,
            text="Keep comments",
            variable=self.keep_comments,
            command=self._on_child_annotation_toggle,
        ).pack(
            anchor="w",
            padx=25,
        )
        ttk.Checkbutton(
            advanced.content_frame,
            text="Keep footnotes",
            variable=self.keep_footnotes,
            command=self._on_child_annotation_toggle,
        ).pack(anchor="w", padx=25)
        ttk.Checkbutton(
            advanced.content_frame,
            text="Keep endnotes",
            variable=self.keep_endnotes,
            command=self._on_child_annotation_toggle,
        ).pack(anchor="w", padx=25)

    def _on_child_annotation_toggle(self) -> None:
        """When any child is toggled, update parent state."""
        children_checked = [
            self.keep_comments.get(),
            self.keep_footnotes.get(),
            self.keep_endnotes.get(),
        ]

        if all(children_checked):
            # All checked → parent checked
            self.keep_all_annotations.set(True)
            self.keep_all_annotations_chk.state(["!alternate"])  # Clear indeterminate
        elif any(children_checked):
            # Some checked → parent indeterminate (dash/gray)
            self.keep_all_annotations_chk.state(["alternate"])
        else:
            # None checked → parent unchecked
            self.keep_all_annotations.set(False)
            self.keep_all_annotations_chk.state(["!alternate"])  # Clear indeterminate

    def _on_parent_annotation_toggle(self) -> None:
        """When parent is toggled, set all children to match."""
        parent_value = self.keep_all_annotations.get()
        self.keep_comments.set(parent_value)
        self.keep_footnotes.set(parent_value)
        self.keep_endnotes.set(parent_value)

    def _create_action_section(self) -> None:
        """Create convert button."""
        # ActionFrame for convert button
        # v1 inline
        action_frame = ttk.Frame(self)
        action_frame.grid(row=4, column=0, sticky="ew", padx=5, pady=5)

        self.convert_btn = ttk.Button(
            action_frame,
            text="Convert",
            command=self.on_convert_click,
            state="disabled",  # Start disabled
            padding=10,
            style="Convert.TButton",
        )
        self.convert_btn.grid(row=0, column=0, padx=5, sticky="ew")

        # Watch for file selection to enable button
        self.input_selector.selected_path.trace_add("write", self._on_file_selected)

        # Configure column weights so widgets stretch
        self.columnconfigure(0, weight=1)

        action_frame.columnconfigure(
            0, weight=1
        )  # Convert button - stretches east-west

    def _on_file_selected(self, *args) -> None:  # noqa: ANN002
        """Enable convert button when a file is selected."""
        path = self.input_selector.selected_path.get()
        if path and path != "No selection":
            self.convert_btn.config(
                state="normal",
            )
        else:
            self.convert_btn.config(state="disabled")

    # We could make this _private since it is only called inside this class,
    # but callbacks are conventionally public in python.
    def on_convert_click(self) -> None:
        """Handle convert button click."""
        # Disable button BEFORE starting thread (on UI thread)
        self.convert_btn.config(state="disabled", text="Converting...")
        # self.update_idletasks()  # Force UI to refresh NOW

        # Start with loaded config (if any) or defaults
        cfg = self.loaded_config if self.loaded_config else UserConfig()

        # Update with UI values (preserves fields not in UI)
        cfg = self.ui_to_config(cfg)

        # TODO: Prepare data for us to call the pipeline with by performing basic validation of selected options,
        # building a UserConfig object from valid UI selections, and starting a background thread for the pipeline
        # to be run on.

        # Start background thread
        thread = threading.Thread(
            target=self._run_conversion_thread, args=(cfg,), daemon=True
        )
        thread.start()

    def _run_conversion_thread(self, cfg: UserConfig) -> None:
        """Run the conversion in a background thread."""
        # == DEBUGGING == #
        # TODO: remove
        if DEBUG_MODE:
            import time

            time.sleep(5)  # Fake work for 5 seconds
        # == ========= == #

        try:
            run_pipeline(cfg)
            # Success! Schedule UI update on main thread
            self.winfo_toplevel().after(0, self._on_conversion_success)
        except Exception as e:
            # Error! Schedule UI update on main thread
            self.winfo_toplevel().after(0, self._on_conversion_error, e)
        pass

    def ui_to_config(self, cfg: UserConfig) -> UserConfig:
        """Gather UI-selected values and update the UserConfig object"""
        # TODO: Gather UI values into UserConfig
        # Only update fields that have UI controls
        cfg.input_docx = self.input_selector.selected_path.get()
        cfg.chunk_type = ChunkType(self.chunk_var.get())
        cfg.experimental_formatting_on = self.exp_fmt_var.get()  # get from BooleanVar
        cfg.preserve_docx_metadata_in_speaker_notes = self.keep_metadata.get()
        # TODO: the rest of the bools
        return cfg

    def config_to_ui(self, cfg: UserConfig) -> None:
        """Populate UI values from a loaded UserConfig"""
        # TODO: Populate UI values from a loaded UserConfig
        # Only populate fields that have UI controls
        pass

    def _on_conversion_success(self) -> None:
        """Inform the user of pipeline success"""
        log.info("Re-enabling convert button (success)")
        self.convert_btn.config(state="normal", text="Convert")
        # TODO: Popup message box and offer to open the output folder (call the helper)

    def _on_conversion_error(self, error: Exception) -> None:
        """Inform the user of pipeline failure and error."""
        log.error(f"Re-enabling convert button (error): {error}")
        self.convert_btn.config(state="normal", text="Convert")
        # TODO: Popup error message box

    def load_and_validate_config(self, path: Path) -> None:
        """Load config from file, validating it matches this tab's direction."""

        # Load a config from disk into memory
        try:
            cfg = UserConfig.from_toml(path)  # Load from disk
        except Exception as e:
            error_msg = f"Failed to load config:\n\n{str(e)}"
            log.info(error_msg)
            messagebox.showerror("Load Failed", error_msg)
            return

        # Validate direction matches this tab
        if cfg.direction != PipelineDirection.DOCX_TO_PPTX:
            log.info("Wrong config type loaded; rejecting and informing user.")
            messagebox.showerror(
                "Invalid Config",
                "This config is for PPTX→DOCX.\n"
                "Please use the PPTX→DOCX tab to load this config.",
            )
            # TODO: Offer to swap tabs and load the config there for them or cancel.
            # Note they'll still need to make sure an input file for conversion is selected on the new tab of the right type.
            return

        self.config_to_ui(cfg)  # Populate UI
        self.loaded_config = cfg  # Store it as THE config

        success_msg = f"Loaded config from {Path(path).name}"
        log.info(success_msg)
        messagebox.showinfo("Config Loaded", success_msg)


# endregion


# region Pptx2DocxTab
class Pptx2DocxTab(ttk.Frame):
    """Tab frame for the Pptx2Docx Pipeline."""

    def __init__(self, parent: tk.Widget, log_viewer: tk.Widget) -> None:
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

    def __init__(self, parent: tk.Widget, log_viewer: tk.Widget) -> None:
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

# region components
# =============== Component Classes ============ #
# endregion


# region CollapsibleFrame
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
            self,
            text=f"{arrow} {title}",
            command=self.toggle,
            style="Collapse.TButton",
        )
        self.toggle_btn.pack(
            fill="x",
            padx=5,
            pady=2,
        )

        # Content frame (for child widgets)
        self.content_frame = ttk.Frame(self)
        if not start_collapsed:
            self.content_frame.pack(fill="both", expand=True, padx=5, pady=5)

    def toggle(self) -> None:
        """Toggle between collapsed and expanded states."""
        if self.is_collapsed:
            # Expand
            self.content_frame.pack(fill="both", expand=True, padx=5, pady=5)
            self.toggle_btn.config(text=self.toggle_btn.cget("text").replace("▶", "▼"))
            self.is_collapsed = False
        else:
            # Collapse
            self.content_frame.pack_forget()
            self.toggle_btn.config(text=self.toggle_btn.cget("text").replace("▼", "▶"))
            self.is_collapsed = True


# endregion


# region PathSelector
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

    def _create_widgets(self) -> None:
        """Create the path widgets"""
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

    def browse(self) -> None:
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


# endregion


# region SaveLoadConfig
class SaveLoadConfig(ttk.Frame):
    pass


# endregion


# region ActionFrame
class ActionFrame(ttk.Frame):
    pass


# endregion


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
            font=("Courier", 10),
            padx=5,
            pady=5,
        )
        self.text_widget.pack(fill="both", expand=True, padx=5, pady=5)

        self.clear_btn = ttk.Button(self, text="Clear Log", command=self.clear_log)
        self.clear_btn.pack(side="left", pady=(0, 5), padx=5)

    def clear_log(self) -> None:
        """Clear all text from the log viewer."""
        self.text_widget.config(state="normal")  # enable editing
        # Tkinter's Text widget uses string indices. "1.0" = line 1, character 0 (the very start)
        self.text_widget.delete("1.0", "end")
        self.text_widget.config(state="disabled")

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


def show_completion_dialog(output_folder: Path) -> None:
    """Helper to pop a message box for any tab that's completed a pipeline run & offer a button for the user to click if they want us to open the output folder for them."""
    pass


def open_folder_in_os_explorer(folder_path: Path) -> None:
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
