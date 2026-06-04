from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, QThread, Signal

from ..models import TaskRow
from ..services.backend import ConfiguratorBackend, RunOptions


class CollectionWorker(QObject):
    log_message = Signal(str)
    progress_changed = Signal(int, int)
    row_status_changed = Signal(str, str)
    finished_ok = Signal(str)
    failed = Signal(str)

    def __init__(self, backend: ConfiguratorBackend, tasks: list[TaskRow], vendor_key: str,
                 folder_name: str, output_root: Path | None, options: RunOptions):
        super().__init__()
        self.backend = backend
        self.tasks = tasks
        self.vendor_key = vendor_key
        self.folder_name = folder_name
        self.output_root = output_root
        self.options = options

    def run(self):
        try:
            out_dir = self.backend.run(
                tasks=self.tasks,
                vendor_key=self.vendor_key,
                folder_name=self.folder_name,
                output_root=self.output_root,
                options=self.options,
                log=self.log_message.emit,
                progress=self.progress_changed.emit,
                status=self.row_status_changed.emit,
            )
            self.finished_ok.emit(str(out_dir))
        except Exception as exc:
            self.failed.emit(str(exc))


def start_worker(worker: CollectionWorker):
    thread = QThread()
    worker.moveToThread(thread)
    thread.started.connect(worker.run)
    worker.finished_ok.connect(thread.quit)
    worker.failed.connect(thread.quit)
    thread.finished.connect(thread.deleteLater)
    return thread
