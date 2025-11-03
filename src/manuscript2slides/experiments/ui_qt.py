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
    QGridLayout,
    QLabel,
    QLineEdit,
    QCheckBox,
    QFileDialog,
)
from PySide6.QtCore import QObject, Signal, QSettings
import sys
import time
from pathlib import Path

# log = logging.getLogger("manuscript2slides")
# endregion

# We have to use QSettings to store preference persistence and things like last-selected-dialog because
# QT doesn't integrate as cleanly with OS memory as Tkinter did.
APP_SETTINGS = QSettings("manuscript2slides", "manuscript2slides")


# region classes
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

        # Create persistent dialog instance
        self.dialog = QFileDialog(self)
        if is_dir:
            self.dialog.setFileMode(QFileDialog.FileMode.Directory)
        else:
            self.dialog.setFileMode(QFileDialog.FileMode.ExistingFile)

        self._create_widgets(
            label_text, read_only
        )  # No other method needs read_only, so we don't clutter self with it.
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
        connection1 = self.browse_btn.clicked.connect(self.browse)
        # Emit signal when line edit text changes
        connection2 = self.line_edit.textChanged.connect(self.path_changed.emit)
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


# region main
def main() -> None:
    """Qt UI entry point."""
    # initialize_application()  # configure the log & other startup tasks
    # log.info("Initializing Qt UI")

    app = QApplication(sys.argv)  # Create App
    window = QWidget()
    window.setWindowTitle("Path Selector")
    layout = QVBoxLayout()

    # Create three path selctors
    input_selector = PathSelector(
        window, "Input File: ", filetypes=[".docx"], typenames="Word Documents"
    )
    output_selector = PathSelector(window, "Output Folder:", is_dir=True)
    template_selector = PathSelector(
        window, "Template:", filetypes=[".pptx"], typenames="PowerPoints"
    )

    layout.addWidget(input_selector)
    layout.addWidget(output_selector)
    layout.addWidget(template_selector)

    window.setLayout(layout)
    window.show()
    sys.exit(
        app.exec()
    )  # Start event loop with app.exec(); includes sys.exit() for proper cleanup


if __name__ == "__main__":
    main()
# endregion
