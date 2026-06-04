from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..models import TaskRow
from ..services.backend import ConfiguratorBackend, RunOptions
from ..services.task_store import TaskStore
from ..workers.collection_worker import CollectionWorker, start_worker


class MainWindow(QMainWindow):
    TABLE_COLUMNS = ["Enabled", "IP", "User", "Password", "Status"]

    def __init__(self, base_dir: Path):
        super().__init__()
        self.base_dir = Path(base_dir)
        self.csv_path = self.base_dir / "task.csv"
        self.template_pdf_path = self.base_dir / "Server Configurator_fillable.pdf"
        self.task_store = TaskStore(self.csv_path)
        self.backend = ConfiguratorBackend(self.base_dir, self.template_pdf_path)
        self.rows: list[TaskRow] = []
        self.worker = None
        self.thread = None

        self.setWindowTitle("Server Configurator Desktop")
        self.resize(1500, 900)
        self._build_ui()
        self.load_tasks()

    def _build_ui(self):
        central = QWidget(self)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(10)

        top = self._build_top_bar()
        root.addWidget(top)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_center_panel())
        splitter.addWidget(self._build_right_panel())
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 3)
        splitter.setStretchFactor(2, 2)
        root.addWidget(splitter, 1)

        self.setCentralWidget(central)
        self.statusBar().showMessage("Ready")

    def _build_top_bar(self):
        box = QGroupBox("Run Controls")
        layout = QHBoxLayout(box)

        self.vendor_combo = QComboBox()
        self.vendor_combo.addItems(["dell", "hpe", "supermicro"])
        layout.addWidget(QLabel("Vendor:"))
        layout.addWidget(self.vendor_combo)

        self.folder_name_edit = QLineEdit()
        self.folder_name_edit.setPlaceholderText("RUN name or job reference")
        layout.addWidget(QLabel("Output Folder Name:"))
        layout.addWidget(self.folder_name_edit, 1)

        self.output_root_edit = QLineEdit()
        self.output_root_edit.setText(str(self.base_dir))
        layout.addWidget(QLabel("Output Root:"))
        layout.addWidget(self.output_root_edit, 1)

        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.choose_output_root)
        layout.addWidget(browse_btn)

        self.start_btn = QPushButton("Start Run")
        self.start_btn.clicked.connect(self.start_run)
        layout.addWidget(self.start_btn)

        self.reload_btn = QPushButton("Reload CSV")
        self.reload_btn.clicked.connect(self.load_tasks)
        layout.addWidget(self.reload_btn)

        return box

    def _build_left_panel(self):
        box = QGroupBox("Targets")
        layout = QVBoxLayout(box)

        self.table = QTableWidget(0, len(self.TABLE_COLUMNS))
        self.table.setHorizontalHeaderLabels(self.TABLE_COLUMNS)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        layout.addWidget(self.table, 1)

        btn_row = QHBoxLayout()
        for text, slot in [
            ("Add", self.add_row),
            ("Edit", self.edit_selected_row),
            ("Delete", self.delete_selected_row),
            ("Save CSV", self.save_tasks),
            ("Toggle Enabled", self.toggle_selected_row),
        ]:
            btn = QPushButton(text)
            btn.clicked.connect(slot)
            btn_row.addWidget(btn)
        layout.addLayout(btn_row)
        return box

    def _build_center_panel(self):
        box = QGroupBox("Progress and Logs")
        layout = QVBoxLayout(box)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        layout.addWidget(self.progress)

        self.current_label = QLabel("Current host: -")
        layout.addWidget(self.current_label)

        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output, 1)
        return box

    def _build_right_panel(self):
        box = QGroupBox("Options")
        layout = QFormLayout(box)

        self.timeout_edit = QLineEdit("30")
        layout.addRow("Timeout (sec)", self.timeout_edit)

        self.skip_unreachable = QCheckBox()
        self.skip_unreachable.setChecked(True)
        layout.addRow("Skip unreachable", self.skip_unreachable)

        self.individual_pdf = QCheckBox()
        self.individual_pdf.setChecked(True)
        layout.addRow("Individual PDFs", self.individual_pdf)

        self.group_pdf = QCheckBox()
        self.group_pdf.setChecked(True)
        layout.addRow("Grouped PDF", self.group_pdf)

        self.serial_excel = QCheckBox()
        self.serial_excel.setChecked(True)
        layout.addRow("Serial Excel", self.serial_excel)

        open_folder_btn = QPushButton("Open Output Root")
        open_folder_btn.clicked.connect(self.open_output_root)
        layout.addRow(open_folder_btn)
        return box

    def append_log(self, message: str):
        self.log_output.appendPlainText(message)

    def choose_output_root(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose Output Root", self.output_root_edit.text())
        if folder:
            self.output_root_edit.setText(folder)

    def open_output_root(self):
        folder = Path(self.output_root_edit.text().strip() or str(self.base_dir))
        folder.mkdir(parents=True, exist_ok=True)
        if hasattr(QApplication, 'instance'):
            import os, sys, subprocess
            if sys.platform.startswith('win'):
                os.startfile(str(folder))
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(folder)])
            else:
                subprocess.Popen(['xdg-open', str(folder)])

    def load_tasks(self):
        self.rows = self.task_store.load()
        self.refresh_table()
        self.statusBar().showMessage(f"Loaded {len(self.rows)} rows from task.csv")

    def save_tasks(self):
        self.task_store.save(self.rows)
        self.statusBar().showMessage("task.csv saved")
        QMessageBox.information(self, "Saved", "task.csv updated successfully.")

    def refresh_table(self):
        self.table.setRowCount(len(self.rows))
        for row_index, row in enumerate(self.rows):
            enabled_item = QTableWidgetItem("Yes" if row.enabled else "No")
            ip_item = QTableWidgetItem(row.ip)
            user_item = QTableWidgetItem(row.user)
            pw_item = QTableWidgetItem(row.password)
            status_item = QTableWidgetItem(row.status)
            for item in (enabled_item, ip_item, user_item, pw_item, status_item):
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)
            self.table.setItem(row_index, 0, enabled_item)
            self.table.setItem(row_index, 1, ip_item)
            self.table.setItem(row_index, 2, user_item)
            self.table.setItem(row_index, 3, pw_item)
            self.table.setItem(row_index, 4, status_item)

    def _prompt_row(self, existing: TaskRow | None = None) -> TaskRow | None:
        ip, ok = QInputDialog.getText(self, "IP", "IP Address", text=existing.ip if existing else "")
        if not ok or not ip.strip():
            return None
        user, ok = QInputDialog.getText(self, "User", "Username", text=existing.user if existing else "Administrator")
        if not ok:
            return None
        password, ok = QInputDialog.getText(self, "Password", "Password", text=existing.password if existing else "")
        if not ok:
            return None
        return TaskRow(ip=ip.strip(), user=user.strip(), password=password, enabled=existing.enabled if existing else True,
                       status=existing.status if existing else "Pending")

    def add_row(self):
        row = self._prompt_row()
        if row:
            self.rows.append(row)
            self.refresh_table()

    def edit_selected_row(self):
        idx = self.table.currentRow()
        if idx < 0:
            return
        row = self._prompt_row(self.rows[idx])
        if row:
            self.rows[idx] = row
            self.refresh_table()

    def delete_selected_row(self):
        idx = self.table.currentRow()
        if idx < 0:
            return
        self.rows.pop(idx)
        self.refresh_table()

    def toggle_selected_row(self):
        idx = self.table.currentRow()
        if idx < 0:
            return
        self.rows[idx].enabled = not self.rows[idx].enabled
        self.refresh_table()

    def update_row_status(self, ip: str, status: str):
        for row in self.rows:
            if row.ip == ip:
                row.status = status
                break
        self.refresh_table()
        self.current_label.setText(f"Current host: {ip}")

    def start_run(self):
        if not self.template_pdf_path.exists():
            QMessageBox.warning(self, "Missing Template", f"Template PDF not found:\n{self.template_pdf_path}")
            return
        try:
            timeout = max(5, int(self.timeout_edit.text().strip() or "30"))
        except ValueError:
            QMessageBox.warning(self, "Invalid Timeout", "Timeout must be a whole number.")
            return

        options = RunOptions(
            timeout_seconds=timeout,
            skip_unreachable=self.skip_unreachable.isChecked(),
            create_individual_pdfs=self.individual_pdf.isChecked(),
            create_group_pdf=self.group_pdf.isChecked(),
            create_serial_excel=self.serial_excel.isChecked(),
        )
        output_root = Path(self.output_root_edit.text().strip() or str(self.base_dir))
        folder_name = self.folder_name_edit.text().strip()

        self.log_output.clear()
        for row in self.rows:
            row.status = "Pending"
        self.refresh_table()

        self.worker = CollectionWorker(
            backend=self.backend,
            tasks=self.rows,
            vendor_key=self.vendor_combo.currentText().strip(),
            folder_name=folder_name,
            output_root=output_root,
            options=options,
        )
        self.thread = start_worker(self.worker)
        self.worker.log_message.connect(self.append_log)
        self.worker.progress_changed.connect(self.on_progress)
        self.worker.row_status_changed.connect(self.update_row_status)
        self.worker.finished_ok.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)

        self.start_btn.setEnabled(False)
        self.statusBar().showMessage("Run started")
        self.thread.start()

    def on_progress(self, current: int, total: int):
        if total <= 0:
            self.progress.setValue(0)
            return
        self.progress.setValue(int((current / total) * 100))

    def on_finished(self, output_dir: str):
        self.progress.setValue(100)
        self.start_btn.setEnabled(True)
        self.statusBar().showMessage(f"Completed: {output_dir}")
        QMessageBox.information(self, "Run Complete", f"Output ready:\n{output_dir}")

    def on_failed(self, message: str):
        self.start_btn.setEnabled(True)
        self.statusBar().showMessage("Run failed")
        QMessageBox.critical(self, "Run Failed", message)
