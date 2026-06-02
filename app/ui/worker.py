"""Background worker for conversion jobs."""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import QObject, pyqtSignal

from app.config import ConversionConfig
from app.services.converter import run_conversion


class ConversionWorker(QObject):
    """Runs conversion in a background thread."""

    log = pyqtSignal(str)
    progress = pyqtSignal(int, int)
    finished = pyqtSignal(int, int, int)
    failed = pyqtSignal(str)

    def __init__(
        self,
        input_path: Path,
        output_dir: Path,
        config: ConversionConfig,
    ) -> None:
        super().__init__()
        self._input_path = input_path
        self._output_dir = output_dir
        self._config = config

    def run(self) -> None:
        try:
            processed, skipped, errors = run_conversion(
                input_path=self._input_path,
                output_dir=self._output_dir,
                config=self._config,
                log=self.log.emit,
                on_progress=self.progress.emit,
            )
            self.finished.emit(processed, skipped, errors)
        except Exception as exc:
            self.failed.emit(str(exc))

