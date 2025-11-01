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
from PySide6.QtCore import QObject, Signal
import sys
import time

# log = logging.getLogger("manuscript2slides")
# endregion


# region helpers


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

    firstnum = QLineEdit()
    secondnum = QLineEdit()
    sumnum = QLineEdit()
    sumnum.setDisabled(True)
    sumnum.setPlaceholderText("Sum will print here")

    def calc_if_valid():
        try:
            val1 = int(firstnum.text())
            val2 = int(secondnum.text())
            sumnum.setText(str(val1 + val2))
        except ValueError:
            sumnum.setText("Must be numbers only!")

    # connect the fields to the validation func
    # I think this'd happen in the presenter init
    firstnum.textChanged.connect(calc_if_valid)
    secondnum.textChanged.connect(calc_if_valid)

    layout.addWidget(QLabel("Jojo's calculator"))
    layout.addWidget(firstnum)
    layout.addWidget(secondnum)
    layout.addWidget(sumnum)
    window.setLayout(layout)
    window.show()
    sys.exit(
        app.exec()
    )  # Start event loop with app.exec(); includes sys.exit() for proper cleanup


if __name__ == "__main__":
    main()
# endregion
