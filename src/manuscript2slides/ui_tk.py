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
from typing import Callable

from manuscript2slides.startup import initialize_application
from manuscript2slides.internals.config.define_config import (
    UserConfig,
    PipelineDirection,
    ChunkType,
)
from manuscript2slides.internals.constants import DEBUG_MODE
from manuscript2slides.orchestrator import run_pipeline, run_roundtrip_test
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

    # region _create_widgets()
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

        # (Optional) Customize specific elements (or give up and switch to PySide)
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


# region BaseConversionTab
class BaseConversionTab(ttk.Frame):
    """Base class for conversion tabs with shared threading & button logic."""

    def __init__(self, parent: tk.Widget, log_viewer: LogViewer) -> None:
        super().__init__(parent)
        self.log_viewer = log_viewer
        self.last_run_config = None
        self.buttons = []
        # children should call self._create_widgets()

    def _create_convert_button(
        self, button_text: str, cmd: Callable[[], None]
    ) -> ttk.Button:
        """Create a button widget styled for conversion without grid/pack. Caller must grid/pack."""
        return ttk.Button(self, text=button_text, style="Convert.TButton", command=cmd)

    def _load_config(self, path: Path) -> UserConfig | None:
        try:
            cfg = UserConfig.from_toml(path)  # Load from disk
            log.info(f"Loaded config from {path.name}")
            return cfg
        except Exception as e:
            log.error(
                f"Try again; something went wrong when we tried to load that config from disk: {e}"
            )
            return None

    def start_conversion(
        self, cfg: UserConfig, pipeline_func: Callable | None = None
    ) -> None:
        """
        Disable buttons for the tab and start the conversion background thread.

        NOTE: Child must handle cfg prep and any other unique prep.
        """
        # "None sentinal pattern" to set the default pipeline
        if pipeline_func is None:
            pipeline_func = run_pipeline  # Resolved at runtime

        self.disable_buttons()
        self.last_run_config = cfg
        log.info("Starting conversion in background thread.")
        thread = threading.Thread(
            target=self._run_in_thread, args=(cfg, pipeline_func), daemon=True
        )
        thread.start()

    def disable_buttons(self) -> None:
        """Disable all buttons inside self.buttons[] on this tab. Use to prevent button clicks during conversion pipeline runs."""
        log.debug("Disabling button(s) during conversion.")
        for button in self.buttons:
            button._original_text = button.cget("text")
            button.config(state="disabled", text="Converting...")

    def enable_buttons(self) -> None:
        """Re-enable all buttons inside self.buttons[] on this tab. Use after conversion pipeline is complete."""
        log.debug("Renabling convert button(s).")
        for button in self.buttons:
            button.config(state="normal", text=button._original_text)

    def _run_in_thread(self, cfg: UserConfig, pipeline_func: Callable) -> None:
        """Run a pipeline_func call inside a background thread."""

        # == DEBUGGING == #
        # Pause the UI for a few seconds so we can verify button disable/enable
        if DEBUG_MODE:
            import time

            time.sleep(3)
        # =============== #

        try:
            cfg.validate()
            pipeline_func(cfg)
            # Success! Schedule UI update on main thread
            self.winfo_toplevel().after(0, self._on_conversion_success)
        except Exception as e:
            # Error! Schedule UI update on main thread
            self.winfo_toplevel().after(0, self._on_conversion_error, e)

    def _on_conversion_success(self) -> None:
        """Inform the user of pipeline success"""
        self.enable_buttons()

        # Get the output folder location
        cfg = self.last_run_config if self.last_run_config else UserConfig()
        output_folder = cfg.get_output_folder()

        # Show success message
        message = (
            f"Successfully ran conversion!\n\n"
            f"Output location:\n{output_folder}\n\n"
            f"Open output folder?"
        )

        # Ask if user wants to open folder
        result = messagebox.askokcancel("Conversion Complete!", message)

        # User clicked OK
        if result:
            open_folder_in_os_explorer(output_folder)

    def _on_conversion_error(self, error: Exception) -> None:
        """Inform the user of pipeline failure and error."""
        log.error(f"Re-enabling buttons (error): {error}")
        self.enable_buttons()

        error_msg = str(error)
        if len(error_msg) > 300:
            error_msg = error_msg[:300] + "...\n\n(See log for full details)"

        message = (
            f"An error occurred during conversion:\n\n"
            f"{error_msg}\n\n"
            f"Check the log viewer below for details.\n\n"
            f"Open log folder?"
        )
        result = messagebox.askokcancel("Demo Conversion Failed", message, icon="error")

        if result:
            # Get log folder from config
            log_folder = UserConfig().get_log_folder()
            open_folder_in_os_explorer(log_folder)


# endregion


# region ConfigurableConversionTab
class ConfigurableConversionTab(BaseConversionTab):
    """Extends BaseConversionTab with logic and UI very similar between Docx2PptxTab and Pptx2DocxTab."""

    def __init__(self, parent: tk.Widget, log_viewer: LogViewer) -> None:
        super().__init__(parent, log_viewer)
        self.loaded_config = None
        # Get defaults from UserConfig
        self.cfg_defaults = UserConfig()
        self.convert_btn = None
        # children to call _create_widgets()

    # Abstract - children implement
    def ui_to_config(self, cfg: UserConfig) -> UserConfig:
        """Gather values from UI widgets into config."""
        raise NotImplementedError

    def config_to_ui(self, cfg: UserConfig) -> None:
        """Populate UI widgets from config."""
        raise NotImplementedError

    def get_pipeline_direction(self) -> PipelineDirection:
        """Return this tab's direction for validation."""
        raise NotImplementedError

    def _get_input_path(self) -> str:
        """Get current input path. Child implements."""
        raise NotImplementedError

    def _validate_input(self, cfg: UserConfig) -> bool:
        """Validate input before conversion. Child must implement."""
        raise NotImplementedError

    # Concrete - shared by both conversion tabs
    # region _create_action_section + helpers
    def _create_action_section(self) -> None:
        """Create convert button section."""
        # ActionFrame for convert button
        action_frame = ttk.Frame(self)
        action_frame.grid(
            row=99, column=0, sticky="ew", padx=5, pady=5
        )  # High row number so it goes to bottom

        self.convert_btn = ttk.Button(
            action_frame,
            text="Convert",
            command=self.on_convert_click,
            state="disabled",  # Start disabled
            padding=10,
            style="Convert.TButton",
        )
        self.convert_btn.grid(row=0, column=0, padx=5, sticky="ew")
        self.buttons.append(self.convert_btn)  # For base class disable/enable

        action_frame.columnconfigure(
            0, weight=1
        )  # Convert button - stretches east-west

    def _on_file_selected(self, *args) -> None:  # noqa: ANN002
        """Enable convert button when a file is selected."""
        # Child needs to wire this up with trace_add
        if self.convert_btn:
            path = self._get_input_path()
            if path and path != "No selection":
                self.convert_btn.config(state="normal")
            else:
                self.convert_btn.config(state="disabled")

    def on_convert_click(self) -> None:
        """Handle convert button click with validation."""
        cfg = self.loaded_config if self.loaded_config else UserConfig()
        cfg = self.ui_to_config(cfg)

        # Call child's validation (they implement specifics)
        if not self._validate_input(cfg):
            return  # Validation failed, error already shown

        self.start_conversion(cfg)

    def on_save_config_click(self) -> None:
        """Handle Save Config button click"""
        path = filedialog.asksaveasfilename(
            title="Save Config As",
            defaultextension=".toml",
            filetypes=[("TOML Config", "*.toml")],
            initialfile="my_config.toml",
        )
        if path:
            cfg = self.ui_to_config(UserConfig())
            cfg.save_toml(Path(path))
            messagebox.showinfo("Config Saved", f"Saved config to {Path(path).name}")

    def on_load_config_click(self) -> None:
        """Handle load config button click."""
        path = browse_for_file(
            title="Load Config file", filetypes=[("TOML Config", "*.toml")]
        )
        if path:
            cfg = self._load_config(Path(path))
            if cfg:
                self._validate_loaded_config(cfg)

    def _validate_loaded_config(self, cfg: UserConfig) -> None:
        """Load config with direction validation."""

        # Validate direction
        if cfg.direction != self.get_pipeline_direction():
            messagebox.showerror(
                "Invalid Config",
                f"This config is for {cfg.direction.value}.\n"
                f"Please use the correct tab.",
            )
            # TODO, v2: Offer to swap tabs and load the config there for them or cancel.
            # Note they'll still need to make sure an input file for conversion is selected
            # on the new tab of the right type.
            return

        self.config_to_ui(cfg)
        self.loaded_config = cfg
        messagebox.showinfo("Config Loaded", f"Loaded config successfully")


# endregion


# region Docx2PptxTab class
class Docx2PptxTab(ConfigurableConversionTab):
    """UI Tab for the docx2pptx pipeline."""

    # region d2p init + _create_widgets()
    def __init__(self, parent: tk.Widget, log_viewer: LogViewer) -> None:
        """Constructor for docx2pptx Tab"""
        super().__init__(parent, log_viewer)
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

    def _create_widgets(self) -> None:

        self._create_io_section()

        self._create_basic_options()

        self._create_advanced_options()

        self._create_action_section()

        self.columnconfigure(0, weight=1)

    # endregion

    # region d2p _create_io_section
    def _create_io_section(self) -> None:
        """Create docx2pptx tab's io section."""
        io_section = ttk.LabelFrame(self, text="Input/Output Selection")
        io_section.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        io_section.columnconfigure(0, weight=1)

        # Input file
        self.input_selector = PathSelector(
            io_section, "Input .docx File:", filetypes=[("Word Document", "*.docx")]
        )
        self.input_selector.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=5,
            padx=5,
        )
        # Watch for file selection to enable convert button
        self.input_selector.selected_path.trace_add("write", self._on_file_selected)

        # Advanced (collapsible)
        advanced = CollapsibleFrame(io_section, title="Advanced")
        advanced.grid(row=1, column=0, sticky="ew", pady=5)
        advanced.columnconfigure(0, weight=1)

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

        save_btn = tk.Button(
            advanced.content_frame,
            text="Save Options to Config",
            command=self.on_save_config_click,
        )
        save_btn.pack(side="left", padx=5)

        load_btn = tk.Button(
            advanced.content_frame,
            text="Load Config",
            command=self.on_load_config_click,
        )
        load_btn.pack(side="left", padx=5)

    # endregion

    # region d2p _create_basic_options
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

    # endregion

    # region d2p _create_advanced_options + helpers
    def _create_advanced_options(self) -> None:
        """Create advanced options (collapsible)."""
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
        )
        self.keep_all_annotations_chk.pack(anchor="w")
        ttk.Checkbutton(
            advanced.content_frame,
            text="Keep comments",
            variable=self.keep_comments,
        ).pack(
            anchor="w",
            padx=25,
        )
        ttk.Checkbutton(
            advanced.content_frame,
            text="Keep footnotes",
            variable=self.keep_footnotes,
        ).pack(anchor="w", padx=25)
        ttk.Checkbutton(
            advanced.content_frame,
            text="Keep endnotes",
            variable=self.keep_endnotes,
        ).pack(anchor="w", padx=25)

        self._setup_annotation_observers()

    def _setup_annotation_observers(self) -> None:
        """Wire up parent/child checkbox relationships using observers."""
        # Children notify parent when they change
        self.keep_comments.trace_add("write", self._on_child_annotation_changed)
        self.keep_footnotes.trace_add("write", self._on_child_annotation_changed)
        self.keep_endnotes.trace_add("write", self._on_child_annotation_changed)

        # Parent notifies children when it changes
        self.keep_all_annotations.trace_add("write", self._on_parent_annotation_changed)

    def _on_child_annotation_changed(self, *args) -> None:  # noqa: ANN002
        """Observer: When any child changes, update parent state."""
        children_checked = [
            self.keep_comments.get(),
            self.keep_footnotes.get(),
            self.keep_endnotes.get(),
        ]

        if all(children_checked):
            self.keep_all_annotations.set(True)
            self.keep_all_annotations_chk.state(["!alternate"])  # Clear indeterminate
        elif any(children_checked):
            self.keep_all_annotations_chk.state(["alternate"])
        else:
            self.keep_all_annotations.set(False)
            self.keep_all_annotations_chk.state(["!alternate"])

    def _on_parent_annotation_changed(self, *args) -> None:  # noqa: ANN002
        """Observer: When parent changes, update all children."""
        parent_value = self.keep_all_annotations.get()

        # Setting all these children will actually trigger the child's observer. 
        # We handle this in cycle within the children's observer (`if all(children_checked): / self.keep_all_annotations.set(True)`)
        # so it is actually fine. But if we didn't, we'd need to temporarily 
        # disable child observers here before setting them, in order to avoid infinite loop
        self.keep_comments.set(parent_value)
        self.keep_footnotes.set(parent_value)
        self.keep_endnotes.set(parent_value)

    # endregion

    # region d2p child implementation of parent methods

    def ui_to_config(self, cfg: UserConfig) -> UserConfig:
        """Gather UI-selected values and update the UserConfig object"""

        cfg.direction = self.get_pipeline_direction()

        # Only update fields that have UI controls
        cfg.input_docx = self.input_selector.selected_path.get()
        cfg.chunk_type = ChunkType(self.chunk_var.get())
        cfg.experimental_formatting_on = self.exp_fmt_var.get()
        cfg.preserve_docx_metadata_in_speaker_notes = self.keep_metadata.get()
        cfg.display_comments = self.keep_comments.get()
        cfg.display_footnotes = self.keep_footnotes.get()
        cfg.display_endnotes = self.keep_endnotes.get()

        # Handle optional paths (might be "No selection")
        output = self.output_selector.selected_path.get()
        cfg.output_folder = output if output != "No selection" else None

        template = self.template_selector.selected_path.get()
        cfg.template_pptx = template if template != "No selection" else None
        return cfg

    def config_to_ui(self, cfg: UserConfig) -> None:
        """Populate UI values from a loaded UserConfig"""
        # Only populate fields that have UI controls

        # Set Path selectors
        self.input_selector.selected_path.set(cfg.input_docx or "No selection")
        self.output_selector.selected_path.set(cfg.output_folder or "No selection")
        self.template_selector.selected_path.set(cfg.template_pptx or "No selection")

        # Set dropdown
        self.chunk_var.set(cfg.chunk_type.value)

        # Set checkboxes
        self.exp_fmt_var.set(cfg.experimental_formatting_on)
        self.keep_metadata.set(cfg.preserve_docx_metadata_in_speaker_notes)
        self.keep_comments.set(cfg.display_comments)
        self.keep_footnotes.set(cfg.display_footnotes)
        self.keep_endnotes.set(cfg.display_endnotes)

    def _get_input_path(self) -> str:
        return self.input_selector.selected_path.get()

    def get_pipeline_direction(self) -> PipelineDirection:
        """Return this tab's direction for validation."""
        return PipelineDirection.DOCX_TO_PPTX

    def _validate_input(self, cfg: UserConfig) -> bool:
        """Validate docx-specific input."""
        # Validate required fields
        if not cfg.input_docx or cfg.input_docx == "No selection":
            messagebox.showerror("Missing Input", "Please select an input .docx file.")
            return False

        if not Path(cfg.input_docx).exists():
            messagebox.showerror(
                "File Not Found", f"Input file does not exist:\n{cfg.input_docx}"
            )
            return False

        # Validate it's actually a .docx
        if not cfg.input_docx.endswith(".docx"):
            messagebox.showerror("Invalid File", "Input file must be a .docx file.")
            return False

        return True

    # endregion


# endregion


# region Pptx2DocxTab
class Pptx2DocxTab(ConfigurableConversionTab):
    """Tab frame for the Pptx2Docx Pipeline."""

    def __init__(self, parent: tk.Widget, log_viewer: LogViewer) -> None:
        super().__init__(parent, log_viewer)
        # Get defaults from backend
        self._create_widgets()

    def _create_widgets(self) -> None:
        self._create_io_section()

        # We don't yet offer options for the reverse pipeline;
        # behavior is all inferred from the available data in the .docx.
        # Eventually we might want to; this is where it would be long in the UI.
        # self._create_options()

        self._create_action_section()  # defined by parent

        self.columnconfigure(0, weight=1)

    # NOTE: This could probably be moved to the ConfigurableConversionTab class
    # if we added logic to resolve the slightly different strings based on the
    # tab's pipeline direction. For the time being we've decided the duplication
    # is more readable.
    def _create_io_section(self) -> None:
        """Create Pptx2Docx tab's IO section."""
        io_section = ttk.LabelFrame(self, text="Input/Output Selection")
        io_section.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        io_section.columnconfigure(0, weight=1)

        # Input file
        self.input_selector = PathSelector(
            io_section, "Input .pptx File:", filetypes=[("PowerPoint", "*.pptx")]
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
        advanced.columnconfigure(0, weight=1)

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
            filetypes=[("Word Document", "*.docx")],
            default=str(self.cfg_defaults.get_template_docx_path()),
        )
        self.template_selector.pack(fill="x", pady=2)

        ttk.Separator(advanced.content_frame, orient="horizontal").pack(
            fill="x", pady=5
        )

        save_btn = tk.Button(
            advanced.content_frame,
            text="Save Options to Config",
            command=self.on_save_config_click,
        )
        save_btn.pack(side="left", padx=5)

        load_btn = tk.Button(
            advanced.content_frame,
            text="Load Config",
            command=self.on_load_config_click,
        )
        load_btn.pack(side="left", padx=5)

        # Watch for file selection to enable convert button; callback is defined on parent class
        self.input_selector.selected_path.trace_add("write", self._on_file_selected)

    def _create_options(self) -> None:
        """UNUSED. Create pipeline options widget(s)."""
        options_frame = ttk.Labelframe(self, text="Basic Options")
        options_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)

        # NOTE: Here is where we would add user-configurable options to the pptx2docx UI.

    # region p2d child implementation of parent abstract methods
    def ui_to_config(self, cfg: UserConfig) -> UserConfig:
        """Gather UI-selected values and update the UserConfig object"""

        # Set the direction based on what tab we're in.
        cfg.direction = self.get_pipeline_direction()

        # Only update fields that have UI controls
        cfg.input_pptx = self.input_selector.selected_path.get()

        # Handle optional paths (might be "No selection")
        output = self.output_selector.selected_path.get()
        cfg.output_folder = output if output != "No selection" else None

        template = self.template_selector.selected_path.get()
        cfg.template_docx = template if template != "No selection" else None
        return cfg

    def config_to_ui(self, cfg: UserConfig) -> None:
        """Populate UI values from a loaded UserConfig"""
        # Set Path selectors
        self.input_selector.selected_path.set(cfg.input_pptx or "No selection")
        self.output_selector.selected_path.set(cfg.output_folder or "No selection")
        self.template_selector.selected_path.set(cfg.template_docx or "No selection")

    def _get_input_path(self) -> str:
        return self.input_selector.selected_path.get()

    def get_pipeline_direction(self) -> PipelineDirection:
        """Return this tab's direction for validation."""
        return PipelineDirection.PPTX_TO_DOCX

    def _validate_input(self, cfg: UserConfig) -> bool:
        """Validate pptx-specific input."""
        if not cfg.input_pptx or cfg.input_pptx == "No selection":
            messagebox.showerror("Missing Input", "Please select an input .pptx file.")
            return False

        if not Path(cfg.input_pptx).exists():
            messagebox.showerror(
                "File Not Found", f"Input file does not exist:\n{cfg.input_pptx}"
            )
            return False

        # Validate it's actually a .pptx
        if not cfg.input_pptx.endswith(".pptx"):
            messagebox.showerror("Invalid File", "Input file must be a .pptx file.")
            return False

        return True


# endregion


# region DemoTab
class DemoTab(BaseConversionTab):
    """Tab for running demo dry-runs"""

    # region init & _create_widgets

    def __init__(self, parent: tk.Widget, log_viewer: LogViewer) -> None:
        # Call parent constructor
        super().__init__(parent, log_viewer)
        self._create_widgets()

    def _create_widgets(self) -> None:
        info = ttk.Label(
            self,
            text="Run demos with built-in sample files to try out the pipeline.",
            font=("Arial", 10),
        )
        info.pack(pady=10)

        self.d2p_btn = ttk.Button(
            self,
            text="DOCX → PPTX Demo",
            style="Convert.TButton",
            command=self.on_docx2pptx_demo_click,
        )
        self.d2p_btn.pack(pady=5)
        self.buttons.append(self.d2p_btn)

        self.p2d_btn = ttk.Button(
            self,
            text="PPTX → DOCX Demo",
            style="Convert.TButton",
            command=self.on_pptx2docx_demo_click,
        )
        self.p2d_btn.pack(pady=5)
        self.buttons.append(self.p2d_btn)

        self.round_trip_btn = ttk.Button(
            self,
            text="Round-trip Demo (DOCX → PPTX → DOCX)",
            style="Convert.TButton",
            command=self.on_roundtrip_demo_click,
        )
        self.round_trip_btn.pack(pady=5)
        self.buttons.append(self.round_trip_btn)

        self.load_demo_btn = ttk.Button(
            self,
            text="Load & Run Config",
            style="Convert.TButton",
            command=self.on_load_demo_click,
        )
        self.load_demo_btn.pack(
            pady=5,
        )
        self.buttons.append(self.load_demo_btn)

    # endregion

    # region on_clicks
    def on_docx2pptx_demo_click(self) -> None:
        """Handle DOCX → PPTX Demo button click."""
        direction = PipelineDirection.DOCX_TO_PPTX
        cfg = UserConfig().for_demo(direction=direction)
        self.start_conversion(cfg, run_pipeline)

    def on_pptx2docx_demo_click(self) -> None:
        """Handle PPTX → DOCX Demo button click."""
        direction = PipelineDirection.PPTX_TO_DOCX
        cfg = UserConfig().for_demo(direction=direction)
        self.start_conversion(cfg, run_pipeline)

    def on_load_demo_click(self) -> None:
        """Handle Load & Run Config button click."""
        path = browse_for_file(
            title="Load Config file", filetypes=[("TOML Config", "*.toml")]
        )
        if path:
            cfg = self._load_config(Path(path))
            if cfg:
                # No specific validation in this tab's version
                self.start_conversion(cfg, run_pipeline)

    def on_roundtrip_demo_click(self) -> None:
        """Handle Roundtrip demo button click."""
        cfg = UserConfig().with_defaults()
        self.start_conversion(cfg, run_roundtrip_test)

    # endregion


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
        if self.is_dir:
            path = browse_for_dir(title=f"Select {self.label_text}")
        else:
            path = browse_for_file(
                title=f"Select {self.label_text}", filetypes=self.filetypes
            )

        if path:
            self.selected_path.set(path)


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
def open_folder_in_os_explorer(folder_path: Path | str) -> None:
    """
    Open the folder in the system file explorer, platform-specific.

    Args:
        folder_path: Path to the folder to open
    """
    try:
        folder_path = Path(folder_path)  # Convert to Path if string
        system = platform.system()

        if system == "Windows":
            subprocess.run(["explorer", str(folder_path)])
        elif system == "Darwin":  # macOS
            subprocess.run(["open", str(folder_path)])
        else:  # Linux and others
            subprocess.run(["xdg-open", str(folder_path)])

        log.info(f"Opened folder: {folder_path}")

    except Exception as e:
        log.error(f"Failed to open folder: {e}")
        messagebox.showwarning(
            "Cannot Open Folder",
            f"Could not open the folder automatically.\n\n" f"Location: {folder_path}",
        )


def browse_for_file(
    title: str,
    filetypes: list[tuple[str, str]] | None = None,
    initial_dir: str | None = None,
) -> str | None:
    """Open the fial dialog for the user to pick a file and return the selected path (or None, if cancelled.)"""
    path = filedialog.askopenfilename(
        title=title,
        filetypes=filetypes if filetypes else [("All files", "*.*")],
        initialdir=initial_dir,
    )
    return path if path else None


def browse_for_dir(title: str, initial_dir: str | None = None) -> str | None:
    """Open directory dialog and return selected path (or None if cancelled)."""
    path = filedialog.askdirectory(title=title, initialdir=initial_dir)
    return path if path else None


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
