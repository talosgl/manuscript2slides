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
        self.setGeometry(100, 100, 800, 600)

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

        # TODO: Add the rest of the tabs

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
        msg = QMessageBox(self.view)
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

# TODO
# region ConfigurableConversionTabView
# endregion
# TODO
# region ConfigurableConversionTabPresenter
# endregion


# region Real Tab Classes
# =============
# endregion


# TODO: clean up collapsible thing
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

        # ======= TODO remove
        # Test collapsible frame
        self.test_collapsible = CollapsibleFrame(
            title="\t\tTest Collapsible", start_collapsed=True
        )
        # Add some dummy content
        test_label = QLabel("This content can be hidden!")
        test_layout = QVBoxLayout()
        test_layout.addWidget(test_label)
        self.test_collapsible.content_frame.setLayout(test_layout)

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

        # ==== TODO REMOVE
        # Test collapsible
        layout.addWidget(self.test_collapsible)

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

# TODO
# region Pptx2DocxView
# endregion
# TODO
# region Pptx2DocxPresenter
# endregion
# TODO
# region Docx2PptxView
# endregion
# TODO
# region Docx2PptxPresenter
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
            file_types: List of allowed extensions (e.g., [".docx", ".txt"])
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

        extensions = " ".join(f"*{ext}" for ext in self.filetypes)
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
        if not path:
            # Empty is OK
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
            self._set_line_edit_color("lightgreen")
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
