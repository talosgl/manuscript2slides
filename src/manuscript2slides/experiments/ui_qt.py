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
)
from PySide6.QtCore import Qt
import sys

# log = logging.getLogger("manuscript2slides")
# endregion


# region event handlers
def on_button_click():
    """Handle button click."""
    print("Yoooo, button clicked!")


# endregion

# region stylers

# endregion


# region main
def main() -> None:
    """Qt UI entry point."""
    # initialize_application()  # configure the log & other startup tasks
    # log.info("Initializing Qt UI")

    app = QApplication(sys.argv)  # Create App
    window = QWidget()
    window.setWindowTitle("Signals & Slots")
    layout = QVBoxLayout()

    checkbox = QCheckBox("Enable button")
    button = QPushButton("Click me")

    # Connect checkbox signal to button's setEnabled slot directly
    checkbox.toggled.connect(button.setEnabled)
    # Q: So this makes it so that button.setEnabled is "listening" for checkbox.toggled to "occur", right? The syntax feels a little backwards, so just making sure.
    # Start with button disabled
    button.setEnabled(False)

    layout.addWidget(checkbox)
    layout.addWidget(button)

    window.setLayout(layout)

    window.show()
    sys.exit(
        app.exec()
    )  # Start event loop with app.exec(); includes sys.exit() for proper cleanup


if __name__ == "__main__":
    main()
# endregion
