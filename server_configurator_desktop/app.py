from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication

from .ui.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Server Configurator Desktop")
    app.setOrganizationName("Open Source Lab")

    base_dir = Path(__file__).resolve().parent.parent
    window = MainWindow(base_dir=base_dir)
    window.show()
    return app.exec()
