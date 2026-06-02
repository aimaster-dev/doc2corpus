"""Application entrypoint for the PyQt desktop UI."""

from __future__ import annotations

import logging
import sys
from pathlib import Path

def _setup_runtime_logging() -> Path:
    log_dir = Path.home() / "CorpusConverter" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "app.log"
    logging.basicConfig(
        filename=str(log_path),
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s - %(message)s",
    )
    return log_path


def main() -> int:
    if "--doc-com-worker" in sys.argv:
        from app.services.doc_com_worker import run_doc_com_worker_cli

        return run_doc_com_worker_cli(sys.argv[1:])
    if "--probe-com-worker" in sys.argv:
        from app.services.doc_com_worker import run_probe_com_worker_cli

        return run_probe_com_worker_cli(sys.argv[1:])

    from PyQt6.QtGui import QIcon
    from PyQt6.QtWidgets import QApplication

    from app.ui.main_window import MainWindow

    log_path = _setup_runtime_logging()
    logging.info("Starting CorpusConverter app")

    app = QApplication(sys.argv)
    icon_path = Path(__file__).parent / "app" / "assets" / "app_icon.svg"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    try:
        window = MainWindow()
        window.show()
        return app.exec()
    except Exception:
        logging.exception("Unhandled startup/runtime exception. See log: %s", log_path)
        raise


if __name__ == "__main__":
    raise SystemExit(main())

