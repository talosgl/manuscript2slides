"""Toy UI to learn Qt/PySide"""

# ruff: noqa
# region imports
from __future__ import annotations
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QPushButton,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QFileDialog,
    QTabWidget,
    QSplitter,
    QMessageBox,
    QPlainTextEdit,
    QGroupBox,
    QFrame,
    QComboBox,
    QCheckBox,
)
from PySide6.QtCore import QObject, Signal, QSettings, QThread, Qt
import sys
import time
from pathlib import Path
from typing import Callable

from manuscript2slides.internals.config.define_config import (
    ChunkType,
    PipelineDirection,
    UserConfig,
)
from manuscript2slides.utils import open_folder_in_os_explorer
from manuscript2slides.internals.constants import DEBUG_MODE
from manuscript2slides.orchestrator import run_pipeline, run_roundtrip_test
from manuscript2slides.startup import initialize_application

import logging

log = logging.getLogger("manuscript2slides")

# We have to use QSettings to store preference persistence and things like last-selected-dialog because
# QT doesn't integrate as cleanly with OS memory as Tkinter did.
APP_SETTINGS = QSettings("manuscript2slides", "manuscript2slides")
# endregion


# region MainWindow
class MainWindow(QMainWindow):
    """Main Qt Application Window."""

    # TODO: How do we change the favicon? Also what do you call a favicon when it's for a desktop app and not a website tab...

    # region init
    def __init__(self) -> None:
        """Constructor for the Main Window UI."""
        super().__init__()  # Initialize using QMainWindow's constructor

        self.setWindowTitle("manuscript2slides")
        # self.resize(800, 600)  # Initial size, but resizable

        # In Tk, we applied theme BEFORE creating widgets (do we need at all in Qt?)
        # self._apply_theme()

        # Build the UI
        self._create_widgets()
        self._create_layout()

    # endregion

    # region _create_widgets()
    def _create_widgets(self):
        """Create main components."""
        self.tabs = QTabWidget()
        self.log_viewer = LogViewer()

        self.d2p_tab_view = Docx2PptxTabView()
        self.d2p_tab_presenter = Docx2PptxTabPresenter(self.d2p_tab_view)
        self.tabs.addTab(self.d2p_tab_view, "DOCX -> PPTX")

        self.p2d_tab_view = Pptx2DocxTabView()
        self.p2d_tab_presenter = Pptx2DocxTabPresenter(self.p2d_tab_view)
        self.tabs.addTab(self.p2d_tab_view, "PPTX -> DOCX")

        self.demo_tab_view = DemoTabView()
        self.demo_presenter = DemoTabPresenter(self.demo_tab_view)
        self.tabs.addTab(self.demo_tab_view, "DEMO")

    # endregion

    # region _create_layout
    def _create_layout(self):
        """Arrange components."""
        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.addWidget(self.tabs)
        splitter.addWidget(self.log_viewer)
        splitter.setSizes([420, 180])

        self.setCentralWidget(splitter)

    # endregion

    # region _apply_theme?

    # endregion


# endregion


# region Abstract Parent Tab Classes
# =============
# endregion


# region BaseConversionTabView
class BaseConversionTabView(QWidget):
    """Base view for all conversion tabs."""

    # region signals
    # endregion

    # region init
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.buttons: list[QPushButton] = []  # Track buttons for enable/disable
        self.button_original_texts: dict[QPushButton, str] = {}
        # Subclasses must define and call their own _create_widgets() and _create_layout()

    # endregion

    # region disable/enable buttons
    # Public interface for Presenter to control the view's widgets
    def disable_buttons(self) -> None:
        """Disable all tracked buttons during conversion. Use to prevent button clicks during conversion pipeline runs."""
        log.debug("Disabling button(s) during conversion.")
        for button in self.buttons:
            # Store original text so we can restore it later
            self.button_original_texts[button] = button.text()
            button.setText("Converting...")

            # Disable button
            button.setEnabled(False)

    def enable_buttons(self) -> None:
        """Re-enable all tracked buttons after conversion."""
        log.debug("Renabling button(s).")
        for button in self.buttons:
            # Restore original text
            original_text = self.button_original_texts[button]
            if original_text:
                button.setText(original_text)

            # Re-enable
            button.setEnabled(True)

    # endregion


# endregion


# region ConversionWorker
class ConversionWorker(QObject):
    """Worker object for running conversions in a background thread."""

    # region Signals
    finished = Signal()  # Emitted when conversion succeeds; passes no args.
    error = Signal(Exception)  # Emitted when conversion fails; passes Exception object.
    # endregion

    # region init
    def __init__(self, cfg: UserConfig, pipeline_func: Callable):
        super().__init__()
        self.cfg = cfg
        self.pipeline_func = pipeline_func

    # endregion

    # region run
    def run(self):
        """Run the conversion (called in a background thread)."""
        # == DEBUGGING == #
        # Pause the UI for a few seconds so we can verify button disable/enable
        if DEBUG_MODE:
            import time

            time.sleep(2)
        # =============== #

        try:
            self.cfg.validate()
            self.pipeline_func(self.cfg)
            self.finished.emit()  # Success
        except Exception as e:
            self.error.emit(e)  # Failure!

    # endregion


# endregion


# region BaseConversionTabPresenter
class BaseConversionTabPresenter(QObject):
    """
    Base Presenter for conversion tabs.

    Inherit from QObject, rather than being a basic Python class, for threading to work correctly.
    Qt needs the Presenter to be part of its object system to route signals properly.
    This is REQUIRED for showing dialogs from signal handlers (like _on_conversion_success/error)
    """

    # region init
    def __init__(self, view: BaseConversionTabView) -> None:
        super().__init__()  # Initialize QObject
        self.view = view
        self.last_run_config = None

        # Instance variables to prevent garbage collection; storing self.worker_thread and self.worker keeps Python references alive during execution.
        self.worker_thread = None
        self.worker = None

    # endregion

    # region _load_config
    def _load_config(self, path: Path) -> UserConfig | None:
        """Load config from disk."""
        try:
            cfg = UserConfig.from_toml(path)  # Load from disk
            log.info(f"Loaded config from {path.name}")
            return cfg
        except Exception as e:

            log.error(
                f"Try again; something went wrong when we tried to load that config from disk: {e}"
            )

            QMessageBox.critical(
                self.view,
                "Title",
                "Try again; something went wrong when we tried to load that config from disk. See log for details.",
            )

            return None

    # endregion

    # region start_conversion
    def start_conversion(
        self, cfg: UserConfig, pipeline_func: Callable | None = None
    ) -> None:
        """
        Disable buttons for the tab and start the conversion background thread.

        NOTE: Subclasses must handle cfg prep and any other unique prep.
        """
        # "None sentinal pattern" to set the default pipeline
        if pipeline_func is None:
            log.debug(
                "No pipeline_func was passed into start_conversion(), so we'll use run_pipeline."
            )
            pipeline_func = run_pipeline  # Resolved at runtime

        self.view.disable_buttons()
        self.last_run_config = cfg
        log.info("Starting conversion in background thread.")

        log.debug("Create thread and worker.")

        # === Qt Threading ===
        # Create thread and worker (local variables)
        self.worker_thread = QThread(
            self
        )  # parents the thread to the Presenter. This prevents garbage collection while keeping Qt's ownership clear.
        self.worker = ConversionWorker(cfg, pipeline_func)

        # Move worker to thread
        self.worker.moveToThread(self.worker_thread)

        # Connect signals
        self.worker_thread.started.connect(
            self.worker.run
        )  # Start work when thread starts
        self.worker.finished.connect(self._on_conversion_success)
        self.worker.error.connect(self._on_conversion_error)

        # Cleanup when done
        self.worker.finished.connect(
            self.worker_thread.quit
        )  # Note multi-slots per signal :)
        self.worker.error.connect(self.worker_thread.quit)

        # The deleteLater calls tell Qt to safely clean up the objects after the thread finishes.
        # This prevents segfaults from accessing deleted objects.
        # I don't know if we actually need this.
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.error.connect(self.worker.deleteLater)
        self.worker_thread.finished.connect(self.worker.deleteLater)

        # Start the thread
        log.debug("Actually start the thread.")
        self.worker_thread.start()

    # endregion

    # region _show_question_dialog
    def _show_question_dialog(
        self, title: str, text: str, info_text: str, icon: QMessageBox.Icon
    ) -> bool:
        """Helper to show a dialog with OK/Cancel."""
        msg = QMessageBox(parent=self.view)
        msg.setIcon(icon)
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.setInformativeText(info_text)
        msg.setStandardButtons(
            QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel
        )  # Remember: Qt uses the pipe | for flag composition, so we're telling it to add both of these items with this syntax

        return (
            msg.exec() == QMessageBox.StandardButton.Ok
        )  # Returns True if OK, False if Cancel

    # endregion

    # region on_conversion_success/error
    def _on_conversion_success(self) -> None:
        """Handle successful conversion."""
        self.view.enable_buttons()
        log.info("Successful conversion complete!")

        # Get output folder
        cfg = self.last_run_config if self.last_run_config else UserConfig()
        output_folder = cfg.get_output_folder()

        # Pop message box with option to open folder
        result = self._show_question_dialog(
            title="Conversion Complete",
            text="Successfully ran conversion!",
            info_text=f"Output location:\n{output_folder}\n\nOpen output folder?",
            icon=QMessageBox.Icon.Information,
        )

        # If the user clicked OK, open the output folder. Otherwise, they hit cancel, so do nothing.
        if result:
            open_folder_in_os_explorer(output_folder)

    def _on_conversion_error(self, error: Exception) -> None:
        """Handle conversion failure."""
        self.view.enable_buttons()
        log.error(f"Conversion failed: {error}")

        error_msg = str(error)
        if len(error_msg) > 300:
            error_msg = error_msg[:300] + "...\n\n(See log for full details)"

        # Get log folder from config
        cfg = self.last_run_config if self.last_run_config else UserConfig()
        log_folder = cfg.get_log_folder()

        # Pop message box with error information and option to open logs folder
        result = self._show_question_dialog(
            title="Conversion Failed",
            text="An error occurred during conversion:",
            info_text=f"{error_msg}\n\nCheck the log viewer for details.\n\nOpen log folder?",
            icon=QMessageBox.Icon.Critical,
        )

        if result:
            open_folder_in_os_explorer(log_folder)

    # endregion


# endregion


# done~ (apart from styling)
# region ConfigurableConversionTabView
class ConfigurableConversionTabView(BaseConversionTabView):
    """View class for the ConfigurableConversionTab."""

    # region init
    def __init__(self, parent=None) -> None:
        super().__init__(parent)

        # Get defaults from UserConfig
        self.cfg_defaults = UserConfig()

        # Subclasses must define these attributes in their own _create_io_widgets()
        self.input_label = "Input File:"
        self.input_filetypes = self.template_filetypes = ["*.*"]
        self.input_typenames = self.template_typenames = "Files"
        self.template_default = "No Selection"  # This is a bit gross but is there another option? None, with checks?

        # children to call _create_widgets(), _create_layouts(), _connect_internal_signals()

    # endregion

    # region _create_[widgets] pieces (concrete/shared)
    def _create_io_section(self) -> None:
        """Create and layout IO section"""
        self._create_io_widgets()
        self._create_io_layout()

    # region _create_io_widgets (partial)
    def _create_io_widgets(self) -> None:
        """Create I/O section.

        Subclasses must define special attributes in their own _create_io_widgets() BEFORE
        calling super()._create_io_widgets() to override. If they don't, the defaults from
        this parent class's init will be used:

        self.input_label = "Input File:"
        self.input_filetypes = ["*.*"]
        self.input_typenames = "Files"
        self.template_filetypes = ["*.*"]
        self.template_typenames = "Files"
        self.template_default = "No selection"
        """

        # Create I/O Section Group
        self.io_section = QGroupBox("Input/Output Selection")

        # Input file
        self.input_selector = PathSelector(
            parent=self.io_section,
            label_text=self.input_label,
            is_dir=False,
            filetypes=self.input_filetypes,
            typenames=self.input_typenames,
            default_path="No Selection",
            read_only=True,
        )

        # Create Advanced I/O Collapsible Frame/Group
        self.advanced_io = CollapsibleFrame(title="Advanced", start_collapsed=True)

        # Create items that'll go under the collapse
        self.output_selector = PathSelector(
            parent=self.advanced_io.content_frame,
            label_text="Output Folder:",
            is_dir=True,
            default_path=str(self.cfg_defaults.get_output_folder()),
        )

        self.template_selector = PathSelector(
            parent=self.advanced_io.content_frame,
            label_text="Custom Template:",
            filetypes=self.template_filetypes,
            typenames=self.template_typenames,
            default_path=self.template_default,
        )

        self.save_btn = QPushButton("Save Config")
        self.load_btn = QPushButton("Load Config")

    # endregion

    # region _create_io_layout (concrete/shared)
    def _create_io_layout(self) -> None:
        """Arrange the I/O section's widgets & subsections."""

        # Arrange items in the the "Advanced" CollapsibleFrame subsection
        # (NOTE: the advanced_io subsection creates it own layout)
        self.advanced_io.content_layout.addWidget(self.output_selector)

        self.advanced_io.content_layout.addWidget(self.template_selector)

        # Create horizontal line separator to go between paths and buttons
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        self.advanced_io.content_layout.addWidget(separator)

        # Put save/load buttons in their own self-contained sub-layout
        button_layout = QHBoxLayout()

        button_layout.addWidget(self.save_btn)  # Pylance, are you happy?
        button_layout.addWidget(self.load_btn)

        # button_layout.addStretch()  # Push buttons to left
        self.advanced_io.content_layout.addLayout(button_layout)

        # Create the outermost main I/O Layout
        io_layout = QVBoxLayout()

        # Add the input_selector to it
        io_layout.addWidget(self.input_selector)
        io_layout.addWidget(self.advanced_io)

        self.io_section.setLayout(io_layout)

    # endregion

    # region _connect_internal_signals
    def _connect_internal_signals(self):
        """Wire up view's internal logic."""
        self.input_selector.path_changed.connect(self._update_convert_button)

    # endregion

    # region _get_input_path (concrete/shared)
    def _get_input_path(self) -> str:
        return self.input_selector.get_path()

    # endregion

    # region _create_convert_section (concrete/shared)
    def _create_convert_section(self) -> None:
        """Create convert button section."""
        self.convert_section = QGroupBox("Let's Go!")

        self.convert_btn = QPushButton("Convert!")  # TODO: start disabled?
        # TODO: Style the button to be big!
        # TODO: And to be green when ready, grayed-out when not ready/disabled

        self.buttons.append(self.convert_btn)  # For base class disable/enable

        self._update_convert_button(self.input_selector.get_path())

        convert_layout = QVBoxLayout()
        convert_layout.addWidget(self.convert_btn)
        self.convert_section.setLayout(convert_layout)

    # endregion

    # region _update_convert_button (concrete/shared) signal handler/slot
    def _update_convert_button(self, path: str) -> None:
        """Enable/disable convert button based on path validity."""
        if self.convert_btn:
            if path and path != "No selection" and Path(path).exists():
                self.convert_btn.setEnabled(True)
            else:
                self.convert_btn.setEnabled(False)

    # endregion

    # region config_to_ui (abstract)
    def config_to_ui(self, cfg: UserConfig) -> None:
        """Populate UI widgets from config."""
        raise NotImplementedError

    # endregion

    # region get_pipeline_direction (abstract)
    def get_pipeline_direction(self) -> PipelineDirection:
        """Return this tab's direction for validation."""
        raise NotImplementedError

    # endregion


# endregion

# endregion


# done???
# region ConfigurableConversionTabPresenter
class ConfigurableConversionTabPresenter(BaseConversionTabPresenter):
    """Presenter class for the ConfigurableConversionTab."""

    # region init
    def __init__(self, view: ConfigurableConversionTabView) -> None:
        super().__init__(view)
        self.view = view
        self.loaded_config = None

        # subclasses must call self._connect_signals()

    # endregion

    # region shared concrete methods

    # region _connect_signals base method
    # (docx2pptx will probably want to extend)
    def _connect_signals(self) -> None:
        """Wire up view signals to presenter handlers."""

        # Button click handlers
        if self.view.convert_btn:  # Please Pylance
            self.view.convert_btn.clicked.connect(self.on_convert_click)

        if self.view.save_btn:
            self.view.save_btn.clicked.connect(self.on_save_config_click)

        if self.view.load_btn:
            self.view.load_btn.clicked.connect(self.on_load_config_click)

    # endregion

    # region on_convert_click
    def on_convert_click(self) -> None:
        """Handle convert button click with validation."""
        cfg = self.loaded_config if self.loaded_config else UserConfig()
        cfg = self.ui_to_config(cfg)

        # Call child's validation (they implement specifics)
        if not self._validate_input(cfg):
            return  # Validation failed, error already shown

        self.start_conversion(cfg)

    # endregion

    # region on_save_config_click
    def on_save_config_click(self) -> None:
        """Handle Save Config button click"""

        # load the last-used directory from QSettings, if it's there
        last_dir = str(APP_SETTINGS.value("last_browse_directory", ""))

        # Combine directory + filename
        initial_path = (
            str(Path(last_dir) / "my_config.toml") if last_dir else "my_config.toml"
        )

        path, _ = QFileDialog.getSaveFileName(
            parent=self.view,
            caption="Save Config As",
            dir=initial_path,  # Sets BOTH starting directory to "look" in, and the initial filename
            filter="TOML Config (*.toml);;All Files (*)",
        )

        if path:
            # Qt doesn't auto-add extension, so ensure it
            if not path.endswith(".toml"):
                path += ".toml"

            cfg = self.ui_to_config(UserConfig())
            cfg.save_toml(Path(path))
            QMessageBox.information(
                self.view, "Config Saved", f"Saved config to {Path(path).name}"
            )

            # Save the selected path to QSettings so we can load it next session.
            selected_dir = str(Path(path).parent)
            APP_SETTINGS.setValue("last_browse_directory", selected_dir)

    # endregion

    # region on_load_config_click
    def on_load_config_click(self) -> None:
        """Handle load config button click."""
        # load the last-used directory from QSettings, if it's there
        last_dir = str(APP_SETTINGS.value("last_browse_directory", ""))
        path, _ = QFileDialog.getOpenFileName(
            self.view, "Load Config", last_dir, "TOML Config (*.toml)"
        )
        if path:
            # Save the selected path to QSettings so we can load it next session.
            selected_dir = str(Path(path).parent)
            APP_SETTINGS.setValue("last_browse_directory", selected_dir)

            cfg = self._load_config(Path(path))
            if cfg:
                self._validate_loaded_config(cfg)

    # endregion

    # region p2d _validate_loaded_config
    def _validate_loaded_config(self, cfg: UserConfig) -> None:
        """Load config with direction validation."""
        if cfg.direction != self.view.get_pipeline_direction():
            QMessageBox.critical(
                self.view,
                "Invalid Config",
                f"This config is for {cfg.direction.value}.\n"
                f"Please use the correct tab.",
            )
            # TODO, v2: Offer to swap tabs and load the config there for them or cancel.
            # Note they'll still need to make sure an input file for conversion is selected
            # on the new tab of the right type.
            return

        self.view.config_to_ui(cfg)
        self.loaded_config = cfg
        QMessageBox.information(
            self.view, "Config Loaded", f"Loaded config successfully"
        )

    # endregion

    # endregion

    # region abstract methods
    # subclasses must implement

    def ui_to_config(self, cfg: UserConfig) -> UserConfig:
        """Gather values from UI widgets into config."""
        raise NotImplementedError

    def _validate_input(self, cfg: UserConfig) -> bool:
        """Validate input before conversion. Child must implement."""
        raise NotImplementedError

    # endregion


# endregion


# region Real Tab Classes
# =============
# endregion


# region DemoTabView
class DemoTabView(BaseConversionTabView):
    """Demo Tab with sample conversion buttons."""

    # region init
    def __init__(
        self, parent=None
    ) -> (
        None
    ):  # Note: Unlike in Tkinter, where parents are absolutely required at all times, parent=None is a conventional pattern in Qt to allow flexibility
        super().__init__(parent)

        # Create widgets
        self._create_widgets()

        # Arrange them in a layout
        self._create_layout()

    # endregion

    # region _create_widgets
    def _create_widgets(self) -> None:
        # Info Label
        self.info_label = QLabel(
            "Run demos with built-in sample files to try out the pipeline."
        )
        self.info_label.setAlignment(
            Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignTop
        )

        # Create the 4 Demo Buttons (no connections; that's Presenter's job)
        self.docx2pptx_btn = QPushButton("DOCX → PPTX Demo")
        self.pptx2docx_btn = QPushButton("PPTX → DOCX Demo")
        self.round_trip_btn = QPushButton("Round-trip Demo (DOCX → PPTX → DOCX)")
        self.load_demo_btn = QPushButton("Load & Run Config")

        self.buttons.extend(
            [
                self.docx2pptx_btn,
                self.pptx2docx_btn,
                self.round_trip_btn,
                self.load_demo_btn,
            ]
        )

    # endregion

    # region _create_layout
    def _create_layout(self) -> None:
        # Create vertical layout
        layout = QVBoxLayout()

        # Add the widgets in display order

        layout.addWidget(self.info_label)
        layout.addSpacing(10)  # Like pady in Tk

        layout.addWidget(self.docx2pptx_btn)
        layout.addWidget(self.pptx2docx_btn)
        layout.addWidget(self.round_trip_btn)
        layout.addWidget(self.load_demo_btn)

        # Add stretch at bottom to push everything up
        layout.addStretch()

        # Actually apply the layout to this widget
        self.setLayout(layout)

    # endregion


# endregion


# region DemoTabPresenter
class DemoTabPresenter(BaseConversionTabPresenter):
    """Presenter for Demo tab. Handles coordination between the view and the backend (a.k.a model, business logic)"""

    # region init
    def __init__(self, view: DemoTabView) -> None:
        super().__init__(view)  # Call base class init
        self.view = (
            view  # instantiated and passed in from MainWindow's _create_widgets()
        )

        # Wire up the View's buttons (signals) to Presenter's handlers (slots)
        self._connect_signals()

    # endregion

    # region _connect_signals
    def _connect_signals(self) -> None:
        """Connect view's button signals to presenter's handler methods."""

        # Connect Signals to Slots
        self.view.docx2pptx_btn.clicked.connect(self.on_docx2pptx_demo)
        self.view.pptx2docx_btn.clicked.connect(self.on_pptx2docx_demo)
        self.view.round_trip_btn.clicked.connect(self.on_roundtrip_demo)
        self.view.load_demo_btn.clicked.connect(self.on_load_demo)

        # `button.clicked` is a Signal (Qt emits it when button is clicked)
        # `.connect(method)` connects that signal to a Slot (your handler method)
        # This replaces Tkinter's button.config(command=...)

    # endregion

    # region on_{btn}_demo click Handler Methods/"Slots"
    def on_docx2pptx_demo(self) -> None:
        """Handle DOCX → PPTX Demo button click."""
        log.debug("DOCX → PPTX Demo clicked!")
        cfg = UserConfig().for_demo(direction=PipelineDirection.DOCX_TO_PPTX)
        self.start_conversion(cfg, run_pipeline)

    def on_pptx2docx_demo(self) -> None:
        """Handle PPTX → DOCX Demo button click."""
        log.debug("PPTX → DOCX Demo clicked!")
        cfg = UserConfig().for_demo(direction=PipelineDirection.PPTX_TO_DOCX)
        self.start_conversion(cfg, run_pipeline)

    def on_roundtrip_demo(self) -> None:
        """Handle Load & Run Config button click."""
        log.debug("Round-trip Demo clicked!")
        cfg = UserConfig().with_defaults()
        self.start_conversion(cfg, run_roundtrip_test)

    def on_load_demo(self) -> None:
        """Handle Roundtrip demo button click."""
        log.debug("Load & Run Config clicked!")

    # endregion


# endregion


# done~
# region Pptx2DocxView
class Pptx2DocxTabView(ConfigurableConversionTabView):
    """View Tab for the Pptx2Docx Pipeline."""

    # region init
    def __init__(self, parent=None) -> None:
        super().__init__(parent)

        # Create widgets
        self._create_widgets()

        # Arrange them in a layout
        self._create_layout()

        # Wire up internal signals
        self._connect_internal_signals()

    # endregion

    # region _create_widgets
    def _create_widgets(self) -> None:

        self._create_io_section()

        # We don't yet offer options for the reverse pipeline;
        # behavior is all inferred from the available data in the .docx.
        # Eventually we might want to; this is where it would be long in the UI.
        # self._create_basic_options()

        self._create_convert_section()

    # endregion

    # region p2d create_io_widgets() (extended)
    def _create_io_widgets(self) -> None:
        # Define this subclass's unique attributes
        self.input_label = "Input .pptx File:"
        self.input_filetypes = ["*.pptx"]
        self.input_typenames = "PowerPoint"
        self.template_filetypes = ["*.docx"]
        self.template_typenames = "Word Document"
        self.template_default = str(self.cfg_defaults.get_template_docx_path())

        # Call parent's method
        super()._create_io_widgets()

    # endregion

    # region (UNUSED) create basic options
    def _create_basic_options(self) -> None:
        """UNUSED. Create pipeline options widget(s)."""
        self.options_frame = QGroupBox("Basic Options")

        # NOTE: Here is where we could add user-configurable options widgets to the pptx2docx UI.
        # If things get complex, split this into _create_options_widgets() and _create_options_layouts()

    # endregion

    # region _create_layout
    def _create_layout(self) -> None:
        """Arrange sections into main layout."""
        layout = QVBoxLayout()
        layout.addWidget(self.io_section)
        # layout.addWidget(self.options_frame)
        layout.addWidget(self.convert_section)
        layout.addStretch()
        self.setLayout(layout)

    # endregion

    # region p2d _get_pipeline_direction
    def get_pipeline_direction(self) -> PipelineDirection:
        """Return this tab's direction for validation."""
        return PipelineDirection.PPTX_TO_DOCX

    # endregion

    # region p2d config_to_ui
    def config_to_ui(self, cfg: UserConfig) -> None:
        """Populate UI values from a loaded UserConfig"""
        # Set Path selectors
        self.input_selector.set_path(cfg.input_pptx or "No selection")
        self.output_selector.set_path(cfg.output_folder or "No selection")
        self.template_selector.set_path(cfg.template_docx or "No selection")

    # endregion


# endregion


# done?~?~??
# region Pptx2DocxPresenter
class Pptx2DocxTabPresenter(ConfigurableConversionTabPresenter):
    """Presenter class for the PPTX -> Docx Tab."""

    def __init__(self, view: Pptx2DocxTabView) -> None:
        super().__init__(view)
        self.view = view  # self.view already exists from base class, but we have this for typehints
        self.loaded_config = None

        self._connect_signals()

    # region p2d _validate_input
    def _validate_input(self, cfg: UserConfig) -> bool:
        """Validate pptx-specific input."""
        if not cfg.input_pptx or cfg.input_pptx == "No selection":
            QMessageBox.critical(
                self.view, "Missing Input", "Please select a valid .pptx input file."
            )
            return False

        if not Path(cfg.input_pptx).exists():
            QMessageBox.critical(
                self.view,
                "File Not Found",
                f"Input file does not exist:\n{cfg.input_pptx}",
            )
            return False

        # Validate it's actually a .pptx
        if not cfg.input_pptx.endswith(".pptx"):
            QMessageBox.critical(
                self.view, "Invalid File", "Input file must be a .pptx file."
            )
            return False

        return True

    # endregion

    # region p2d ui_to_config
    def ui_to_config(self, cfg: UserConfig) -> UserConfig:
        """Gather UI-selected values and update the UserConfig object"""

        # Set the direction based on what tab we're in.
        cfg.direction = self.view.get_pipeline_direction()

        # Only update fields that have UI controls
        cfg.input_pptx = self.view.input_selector.get_path()

        # Handle optional paths (might be "No selection")
        output = self.view.output_selector.get_path()
        cfg.output_folder = output if output != "No selection" else None

        template = self.view.template_selector.get_path()
        cfg.template_docx = template if template != "No selection" else None
        return cfg

    # endregion


# endregion


# done?!
# region Docx2PptxView
class Docx2PptxTabView(ConfigurableConversionTabView):
    """View Tab for the DOCX -> PPTX Pipeline."""

    # region init _create_widgets()
    def __init__(self, parent=None) -> None:
        """Constructor for docx2pptx Tab"""
        super().__init__(parent)

        # Unlike in tk, no need to declare a bunch of BooleanVar or StringVar in Qt

        # Create widgets
        self._create_widgets()

        # Arrange them in a layout
        self._create_layout()

        # Wire up internal signals
        self._connect_internal_signals()

    def _create_widgets(self) -> None:
        self._create_io_section()

        self._create_basic_options()

        self._create_advanced_options()

        self._create_convert_section()

    # endregion

    # region d2p create_io_widgets() (extended)
    def _create_io_widgets(self) -> None:

        # Define this subclass's unique attributes
        self.input_label = "Input .docx File:"
        self.input_filetypes = ["*.docx"]
        self.input_typenames = "Word Document"
        self.template_filetypes = ["*.pptx"]
        self.template_typenames = "PowerPoint"
        self.template_default = str(self.cfg_defaults.get_template_pptx_path())

        # Call parent's method
        super()._create_io_widgets()

    # endregion

    # region d2p _create_basic_options
    def _create_basic_options(self) -> None:
        """Create pipeline options widgets."""
        self.basic_options = QGroupBox("Basic Options")

        chunk_label = QLabel("Chunk manuscript into slides by:")

        # Use self.* because we know we'll need to read from it later.
        self.chunk_dropdown = QComboBox()
        self.chunk_dropdown.addItems([chunk.value for chunk in ChunkType])
        self.chunk_dropdown.setCurrentText(self.cfg_defaults.chunk_type.value)
        # read with selected_chunk = self.chunk_dropdown.currentText()

        explain_chunks = CollapsibleFrame(
            self.basic_options, title="What do these mean?"
        )
        explain_chunks_text = QPlainTextEdit(parent=explain_chunks.content_frame)
        explain_chunks_text.setPlainText(
            "PARAGRAPH (default): One slide per paragraph break.\n"
            "PAGE: One slide for every page break.\n"
            "Heading (Flat): New slides for every heading, regardless of parent-child hierarchy.\n"
            "Heading (Nested): New slides only on finding a 'parent/grandparent' heading to the previously found. \n"
            "All options create a new slide if there is a page break in the middle of a section.",
        )
        explain_chunks_text.setReadOnly(True)

        # Set up min/max for height so it doesn't get squished
        explain_chunks_text.setMinimumHeight(100)  # Give it some breathing room
        explain_chunks_text.setMaximumHeight(150)  # But not infinite

        # Optional: Remove scrollbars if text fits
        explain_chunks_text.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        explain_chunks_text.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )

        explain_chunks.content_layout.addWidget(explain_chunks_text)

        self.experimental_fmt_chk = QCheckBox(
            "Preserve advanced formatting (experimental)"
        )
        # Unlike in Tk, in Qt it is good practice to split widget creation and state configuration into separate lines
        self.experimental_fmt_chk.setChecked(
            self.cfg_defaults.experimental_formatting_on
        )

        tip_label = QLabel(r"Tip: Disable this if conversion crashes or freezes")
        tip_label.setWordWrap(True)  # wraps at the layout width
        tip_label.setContentsMargins(25, 0, 0, 0)
        tip_label.setStyleSheet("color: gray;")  # same as foreground="gray" in tk

        # Add to layout
        layout = QVBoxLayout()
        layout.addWidget(chunk_label)
        layout.addWidget(self.chunk_dropdown)

        # Explanation section
        layout.addWidget(explain_chunks)

        # Experimental formatting section
        layout.addWidget(self.experimental_fmt_chk)
        layout.addWidget(tip_label)

        self.basic_options.setLayout(layout)

    # endregion

    # region _create_advanced_options
    def _create_advanced_options(self) -> None:
        """Create advanced options (collapsible)."""

        self.advanced_options = CollapsibleFrame(
            title="Advanced Options", start_collapsed=True
        )

        self.keep_metadata_chk = QCheckBox("Preserve metadata in speaker notes")
        self.keep_metadata_chk.setChecked(
            self.cfg_defaults.preserve_docx_metadata_in_speaker_notes
        )
        tip_label = QLabel(
            "Tip: Enable for round-trip conversion (maintains comments, heading formatting, etc.)"
        )
        tip_label.setWordWrap(True)
        tip_label.setMaximumWidth(400)
        tip_label.setContentsMargins(25, 0, 0, 0)
        tip_label.setStyleSheet("color: gray;")

        annotations_label = QLabel(
            "Annotations cannot be replicated in slides, but can be copied into the slides' speaker notes.",
        )
        annotations_label.setWordWrap(True)

        # Parent checkbox
        self.keep_all_annotations_chk = QCheckBox("Keep all annotations")
        self.keep_all_annotations_chk.setTristate(True)

        # Children checkboxes
        self.keep_comments_chk = QCheckBox("Keep comments")
        self.keep_footnotes_chk = QCheckBox("Keep footnotes")
        self.keep_endnotes_chk = QCheckBox("Keep endnotes")

        # Set baseline states before connecting signals
        self.keep_comments_chk.setChecked(self.cfg_defaults.display_comments)
        self.keep_footnotes_chk.setChecked(self.cfg_defaults.display_footnotes)
        self.keep_endnotes_chk.setChecked(self.cfg_defaults.display_endnotes)

        # Set parent state by explicitly calling handler
        self._on_child_annotation_changed()

        # Create sub-layout & indent children
        # TODO: Fix that these checkboxes are squished top-bottom too
        child_layout = QVBoxLayout()
        child_layout.setContentsMargins(25, 0, 0, 0)  # Indent 25px from left
        child_layout.addWidget(self.keep_comments_chk)
        child_layout.addWidget(self.keep_footnotes_chk)
        child_layout.addWidget(self.keep_endnotes_chk)

        # Add everything to layout
        self.advanced_options.content_layout.addWidget(self.keep_metadata_chk)

        # TODO: This is still a bit squished
        self.advanced_options.content_layout.addWidget(tip_label, stretch=1)
        self.advanced_options.content_layout.addWidget(annotations_label, stretch=1)
        self.advanced_options.content_layout.addWidget(self.keep_all_annotations_chk)
        self.advanced_options.content_layout.addLayout(child_layout)

        # Connect signals
        self._setup_annotation_observers()

    # endregion

    # region _setup_annotation_observers()
    def _setup_annotation_observers(self):
        """Wire up parent/child checkbox relationships."""
        # Children notify parent when they change
        self.keep_comments_chk.stateChanged.connect(self._on_child_annotation_changed)
        self.keep_footnotes_chk.stateChanged.connect(self._on_child_annotation_changed)
        self.keep_endnotes_chk.stateChanged.connect(self._on_child_annotation_changed)

        # Parent notifies children when it changes
        self.keep_all_annotations_chk.stateChanged.connect(
            self._on_parent_annotation_changed
        )

    def _on_child_annotation_changed(self, *args) -> None:  # noqa: ANN002
        """Observer: When any child changes, update parent state."""
        children_checked = [
            self.keep_comments_chk.isChecked(),
            self.keep_footnotes_chk.isChecked(),
            self.keep_endnotes_chk.isChecked(),
        ]

        if all(children_checked):
            self.keep_all_annotations_chk.setCheckState(Qt.CheckState.Checked)
        elif any(children_checked):
            self.keep_all_annotations_chk.setCheckState(Qt.CheckState.PartiallyChecked)
        else:
            self.keep_all_annotations_chk.setCheckState(Qt.CheckState.Unchecked)

    def _on_parent_annotation_changed(self, *args) -> None:  # noqa: ANN002
        """Observer: When parent changes, update all children."""
        parent_value = self.keep_all_annotations_chk.checkState()
        if parent_value == Qt.CheckState.Checked:
            parent_bool = True
        elif parent_value == Qt.CheckState.Unchecked:
            parent_bool = False
        else:
            return

        # Setting children will trigger the child's observer. In order to avoid infinite loop,
        # we handle this in cycle within the children's observer (`if all(children_checked): / self.keep_all_annotations.set(True)`),
        # but we also disable child observers here before setting them, out of paranoia.

        self.keep_comments_chk.blockSignals(True)
        self.keep_comments_chk.setChecked(parent_bool)
        self.keep_comments_chk.blockSignals(False)

        self.keep_footnotes_chk.blockSignals(True)
        self.keep_footnotes_chk.setChecked(parent_bool)
        self.keep_footnotes_chk.blockSignals(False)

        self.keep_endnotes_chk.blockSignals(True)
        self.keep_endnotes_chk.setChecked(parent_bool)
        self.keep_endnotes_chk.blockSignals(False)

    # endregion

    # region create_layout
    def _create_layout(self) -> None:
        """Arrange sections into main layout."""
        layout = QVBoxLayout()
        layout.addWidget(self.io_section)
        layout.addWidget(self.basic_options)
        layout.addWidget(self.advanced_options)
        layout.addWidget(self.convert_section)
        layout.addStretch()
        self.setLayout(layout)

    # endregion

    # region d2p _get_pipeline_direction
    def get_pipeline_direction(self) -> PipelineDirection:
        """Return this tab's direction for validation."""
        return PipelineDirection.DOCX_TO_PPTX

    # endregion

    # region d2p config_to_ui
    def config_to_ui(self, cfg: UserConfig) -> None:
        """Populate UI values from a loaded UserConfig"""
        # Only populate fields that have UI controls

        # Set Path selectors
        self.input_selector.set_path(cfg.input_docx or "No selection")
        self.output_selector.set_path(cfg.output_folder or "No selection")
        self.template_selector.set_path(cfg.template_pptx or "No selection")

        # Set dropdown
        self.chunk_dropdown.setCurrentText(cfg.chunk_type.value)
        # Qt will search the items in the combo box and select the one matching that text.
        # If the text isn't in the combo, nothing changes.

        # Set checkboxes
        self.experimental_fmt_chk.setChecked(cfg.experimental_formatting_on)
        self.keep_metadata_chk.setChecked(cfg.preserve_docx_metadata_in_speaker_notes)
        self.keep_comments_chk.setChecked(cfg.display_comments)
        self.keep_footnotes_chk.setChecked(cfg.display_footnotes)
        self.keep_endnotes_chk.setChecked(cfg.display_endnotes)

    # endregion


# endregion


# Done?
# region Docx2PptxPresenter
class Docx2PptxTabPresenter(ConfigurableConversionTabPresenter):

    def __init__(self, view: Docx2PptxTabView) -> None:
        super().__init__(view)
        self.view = view  # self.view already exists from base class, but we have this for typehints
        self.loaded_config = None

        self._connect_signals()

    # TODO?: Extend for this subclass's needs
    # RE: Are there any? All the observer stuff is in the UI Only.
    # TODO: Delete?
    def _connect_signals(self):
        super()._connect_signals()  # Get shared signals
        # ... add subclass-specific ones

    # region d2p ui_to_config
    def ui_to_config(self, cfg: UserConfig) -> UserConfig:
        """Gather UI-selected values and update the UserConfig object"""

        cfg.direction = self.view.get_pipeline_direction()

        # Only update fields that have UI controls
        cfg.input_docx = self.view.input_selector.get_path()
        cfg.chunk_type = ChunkType(self.view.chunk_dropdown.currentText())
        cfg.experimental_formatting_on = self.view.experimental_fmt_chk.isChecked()
        cfg.preserve_docx_metadata_in_speaker_notes = (
            self.view.keep_metadata_chk.isChecked()
        )
        cfg.display_comments = self.view.keep_comments_chk.isChecked()
        cfg.display_footnotes = self.view.keep_footnotes_chk.isChecked()
        cfg.display_endnotes = self.view.keep_endnotes_chk.isChecked()

        # Handle optional paths (might be "No selection")
        output = self.view.output_selector.get_path()
        cfg.output_folder = output if output != "No selection" else None

        template = self.view.template_selector.get_path()
        cfg.template_pptx = template if template != "No selection" else None
        return cfg

    # endregion

    # region d2p _validate_input
    def _validate_input(self, cfg: UserConfig) -> bool:
        """Validate input before conversion. Child must implement."""
        if not cfg.input_docx or cfg.input_docx == "No selection":
            QMessageBox.critical(
                self.view, "Missing Input", "Please select a valid .docx input file."
            )
            return False

        if not Path(cfg.input_docx).exists():
            QMessageBox.critical(
                self.view,
                "File Not Found",
                f"Input file does not exist:\n{cfg.input_docx}",
            )
            return False

        # Validate it's actually a .docx
        if not cfg.input_docx.endswith(".docx"):
            QMessageBox.critical(
                self.view, "Invalid File", "Input file must be a .docx file."
            )
            return False

        return True

    # endregion


# endregion

# region Components
# ===========
# endregion


# region CollapsibleFrame
class CollapsibleFrame(QWidget):
    """A widget that can be collapsed/expanded with a toggle button."""

    def __init__(
        self, parent=None, title: str = "Advanced", start_collapsed: bool = True
    ) -> None:
        super().__init__(parent)

        self.title = title
        self.is_collapsed = start_collapsed

        self._create_widgets()
        self._create_layout()

    def _create_widgets(self) -> None:
        """Create toggle button and content frame."""
        # Arrow symbol based on collapsed state
        arrow = "▶" if self.is_collapsed else "▼"

        # Toggle button
        self.toggle_btn = QPushButton(f"{arrow} {self.title}")
        self.toggle_btn.setFlat(True)  # No button border/background
        self.toggle_btn.clicked.connect(self.toggle)

        # Content frame (where child widgets go)
        self.content_frame = QWidget()
        self.content_frame.setVisible(
            not self.is_collapsed
        )  # Hidden if starting collapsed

        self.content_layout = QVBoxLayout()
        self.content_frame.setLayout(self.content_layout)

    def _create_layout(self) -> None:
        """Arrange widgets in layout."""
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)  # No extra padding around the edges

        # Toggle button / title at the top
        layout.addWidget(self.toggle_btn)

        # Content frame below
        layout.addWidget(self.content_frame)

        self.setLayout(layout)

    def toggle(self):
        """Toggle between collapsed and expanded states."""
        if self.is_collapsed:
            # Expand
            self.content_frame.setVisible(True)
            self.toggle_btn.setText(f"▼ {self.title}")
            self.is_collapsed = False
        else:
            # Collapse
            self.content_frame.setVisible(False)
            self.toggle_btn.setText(f"▶ {self.title}")
            self.is_collapsed = True

        # Force layout recalculation
        self.adjustSize()
        self.updateGeometry()


# endregion


# region PathSelector
class PathSelector(QWidget):
    """
    Reusable file/folder path selector component.

    Features:
    - Label, text field, and browse button
    - Support for files or directories
    - Optional file type filtering
    """

    # Custom signal - emit when path changes
    path_changed = Signal(str)  # pass the new path

    def __init__(
        self,
        parent,
        label_text: str = "Path: ",
        is_dir: bool = False,
        filetypes: list = [],
        typenames: str = "Files",
        default_path: str | None = None,
        read_only: bool = True,
    ) -> None:
        """
        Args:
            parent: Parent widget
            label_text: Label to display
            is_dir: True for folder selection, False for file selection
            file_types: List of allowed extensions (e.g., ["*.docx", "*.txt"])
        """
        super().__init__(parent)

        self.is_dir = is_dir
        self.filetypes = filetypes
        self.typenames = typenames

        self._create_widgets(
            label_text, read_only
        )  # No other method need label_text or read_only, so we pass them rather than adding to self.__.
        self._create_layout()  # arrange them
        self._connect_signals()

        # We call this after _create_widgets() so that self.line_edit exists.
        if default_path:
            self.set_path(default_path)

    def _create_widgets(self, label_text, read_only):
        """Create child widgets"""
        self.label = QLabel(label_text)
        self.line_edit = QLineEdit()
        self.line_edit.setReadOnly(read_only)
        self.browse_btn = QPushButton("Browse...")

    def _create_layout(self):
        """Arrange widgets horizontally: label-input-button"""
        # add widgets to layout and arrange them
        layout = QHBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(
            self.line_edit, stretch=1
        )  # give line edit an extra space, width makes it expandable
        layout.addWidget(self.browse_btn)
        layout.setContentsMargins(0, 0, 0, 0)  # No extra padding
        self.setLayout(layout)

    def _connect_signals(self):
        """Wire up internal signals."""
        self.browse_btn.clicked.connect(self.browse)
        # Emit signal when line edit text changes
        self.line_edit.textChanged.connect(self.path_changed.emit)
        self.line_edit.textChanged.connect(self._validate_path)

    def _build_qtfilter_str(self) -> str:
        """Build Qt file filter string from file types."""
        if not self.filetypes:
            # No specific types - just "All Files"
            return "All Files (*)"

        extensions = " ".join(f"{ext}" for ext in self.filetypes)
        return f"{self.typenames} ({extensions});;All Files (*)"

    def browse(self):
        """Open file/folder dialog."""
        # load the last-used directory from QSettings, if it's there
        last_dir = str(APP_SETTINGS.value("last_browse_directory", ""))

        if self.is_dir:
            path = QFileDialog.getExistingDirectory(
                parent=self, caption="Select Folder", dir=last_dir
            )
        else:
            filter_str = self._build_qtfilter_str()
            path, _ = QFileDialog.getOpenFileName(
                parent=self, caption="Select File", filter=filter_str, dir=last_dir
            )

        if path:  # if the user picked something and did not cancel...
            self.line_edit.setText(path)

            save_dir = str(Path(path).parent) if not self.is_dir else path

            # Save the selected path to QSettings so we can load it next session.
            APP_SETTINGS.setValue("last_browse_directory", save_dir)

    def _validate_path(self, path: str) -> None:
        """Validate path and change line_edit color accordingly"""
        if not path or path == "No Selection":
            # Empty or "No Selection" is OK
            self._set_line_edit_color(None)
            return

        try:
            path_obj = Path(path)
        except:
            return

        if self.is_dir:
            is_valid = path_obj.is_dir()
        else:
            is_valid = path_obj.is_file()

        # Set color based on validity
        if is_valid:
            self._set_line_edit_color("green")
        else:
            self._set_line_edit_color("lightcoral")

    def _set_line_edit_color(self, color: str | None):
        """Set background color of line edit."""
        if color:
            self.line_edit.setStyleSheet(f"background-color: {color};")
        else:
            self.line_edit.setStyleSheet("")  # Reset to default

    # Public API / getter/setters
    def get_path(self) -> str:
        """Get currently selected path."""
        return self.line_edit.text()

    def set_path(self, new_path: str) -> None:
        """Set the path programmatically."""
        self.line_edit.setText(new_path)

    def clear(self) -> None:
        """Clear the selected path."""
        self.line_edit.clear()


# endregion


# region LogViewer
class LogViewer(QGroupBox):
    """Scrolling log viewer."""

    # region init
    def __init__(self, parent=None) -> None:
        super().__init__(title="Log Viewer", parent=parent)

        self._create_widgets()
        self._create_layout()
        self._setup_log_handler()

    # endregion

    # region _create_widgets
    def _create_widgets(self) -> None:
        """Create the text widget and clear button."""
        self.text_widget = QPlainTextEdit()
        self.text_widget.setReadOnly(True)  # User can't edit

        # Set a monospace font for logs
        font = self.text_widget.font()
        font.setFamily("Courier")
        self.text_widget.setFont(font)

        # Clear button
        self.clear_btn = QPushButton("Clear Log")
        self.clear_btn.clicked.connect(self.clear_log)

    # endregion

    # region _create_layout
    def _create_layout(self) -> None:
        layout = QVBoxLayout()

        # Text widget to take up most of the space
        layout.addWidget(self.text_widget)

        # Clear button at the bottom
        layout.addWidget(self.clear_btn)

        # Apply layout to this widget
        self.setLayout(layout)

    # endregion

    # region _setup_log_handler
    def _setup_log_handler(self) -> None:
        """Connect the log viewer text widget to the logging system via our custom handler"""

        # We already got the logger at the top of the file, with log = logging.getLogger("manuscript2slides") after our imports,
        # but this is more readable for someone coming to the code later
        logger = logging.getLogger("manuscript2slides")

        # Create our custom handler
        text_handler = QTextEditHandler(self.text_widget)

        # format log messages
        formatter = logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s", datefmt="%H:%M:%S"
        )
        text_handler.setFormatter(formatter)

        # Add handler to logger
        logger.addHandler(text_handler)

        log.info("Log viewer initialized in Qt UI")

    # endregion

    # region clear_log
    def clear_log(self) -> None:
        """Clear all text from the log viewer."""
        log.info("Clearing log view!")
        self.text_widget.clear()

    # endregion


# endregion


# region LogSignaller
class LogSignaller(QObject):
    """Helper class to emit logging signals in a thread-safe way.

    See: https://plumberjack.blogspot.com/2019/11/a-qt-gui-for-logging.html
    """

    log_message = Signal(str)


# endregion


# region QTextEditHandler extending logging.Handler
class QTextEditHandler(logging.Handler):
    """Custom logging handler that writes to a QPlainTextEdit widget."""

    def __init__(self, text_widget: QPlainTextEdit):
        super().__init__()
        self.text_widget = text_widget
        self.signaller = LogSignaller()

        # Connect signal to widget's appendPlainText slot (thread-safe!)
        self.signaller.log_message.connect(self.text_widget.appendPlainText)

    def emit(self, record: logging.LogRecord) -> None:
        """Called by logging system when a log message is generated."""
        # Format the message
        msg = self.format(record=record)

        # Emit signal (Qt routes to main thread automatically)
        self.signaller.log_message.emit(msg)


# region init

# endregion

# region emit
# endregion

# region append_log
# endregion

# endregion


# region main
def main() -> None:
    """Qt UI entry point."""
    initialize_application()  # configure the log & other startup tasks
    log.info("Initializing Qt UI")

    app = QApplication(sys.argv)  # Create App

    # Create the window
    window = MainWindow()
    # Make it visible
    window.show()

    sys.exit(
        app.exec()
    )  # Start event loop with app.exec(); includes sys.exit() for proper cleanup


if __name__ == "__main__":
    main()
# endregion
