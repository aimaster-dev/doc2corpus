"""Main PyQt window for professional document conversion workflow."""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import QSettings, QThread, QUrl
from PyQt6.QtGui import QAction, QDesktopServices, QFont, QIcon
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.config import DEFAULT_SUPPORTED_EXTENSIONS, ConversionConfig
from app.services.extractors import get_tesseract_languages
from app.ui.worker import ConversionWorker


class MainWindow(QMainWindow):
    """Professional desktop UI for document-to-corpus conversion."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Professional Document Corpus Converter")
        self.resize(980, 720)
        self._settings = QSettings("Doc2Pdf", "CorpusConverter")
        self._all_languages: list[str] = []
        self._thread: QThread | None = None
        self._worker: ConversionWorker | None = None
        self._is_loading_saved_state = False
        self._icon_path = Path(__file__).resolve().parents[1] / "assets" / "app_icon.svg"
        if self._icon_path.exists():
            self.setWindowIcon(QIcon(str(self._icon_path)))
        self._setup_menu()
        self._setup_ui()
        self._apply_styles()
        self._load_settings()

    def _setup_menu(self) -> None:
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("File")
        tools_menu = menu_bar.addMenu("Tools")
        help_menu = menu_bar.addMenu("Help")

        open_output_action = QAction("Open Output Folder", self)
        open_output_action.triggered.connect(self._open_output_folder)
        file_menu.addAction(open_output_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        load_lang_action = QAction("Load Tesseract Languages", self)
        load_lang_action.triggered.connect(self._load_languages)
        tools_menu.addAction(load_lang_action)

        about_action = QAction("About", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    def _setup_ui(self) -> None:
        root = QWidget(self)
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setSpacing(14)

        title = QLabel("Document to Raw Corpus TXT")
        title.setObjectName("titleLabel")
        subtitle = QLabel(
            "Load Tesseract, select OCR languages, then convert a single file or an entire folder."
        )
        subtitle.setObjectName("subtitleLabel")
        layout.addWidget(title)
        layout.addWidget(subtitle)

        layout.addWidget(self._build_tesseract_group())
        layout.addWidget(self._build_input_group())
        layout.addWidget(self._build_options_group())
        layout.addWidget(self._build_log_group(), stretch=1)

        run_row = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.start_button = QPushButton("Start Conversion")
        self.start_button.clicked.connect(self._start_conversion)
        run_row.addWidget(self.progress_bar, stretch=1)
        run_row.addWidget(self.start_button)
        layout.addLayout(run_row)

    def _build_tesseract_group(self) -> QGroupBox:
        group = QGroupBox("OCR Engine")
        form = QFormLayout(group)

        path_row = QHBoxLayout()
        self.tesseract_input = QLineEdit()
        self.tesseract_input.setPlaceholderText("Path to tesseract.exe")
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self._browse_tesseract)
        load_button = QPushButton("Load & Detect Languages")
        load_button.clicked.connect(self._load_languages)
        path_row.addWidget(self.tesseract_input)
        path_row.addWidget(browse_button)
        path_row.addWidget(load_button)
        form.addRow("Tesseract", path_row)

        self.language_search = QLineEdit()
        self.language_search.setPlaceholderText("Search languages (e.g., eng, kor, deu)")
        self.language_search.textChanged.connect(self._filter_languages)
        form.addRow("Search", self.language_search)

        self.lang_list = QListWidget()
        self.lang_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.lang_list.setMinimumHeight(220)
        form.addRow("Languages", self.lang_list)
        return group

    def _build_input_group(self) -> QGroupBox:
        group = QGroupBox("Input / Output")
        form = QFormLayout(group)

        input_row = QHBoxLayout()
        self.input_path = QLineEdit()
        self.input_path.setPlaceholderText("Input document path or directory")
        input_file_btn = QPushButton("Single File")
        input_file_btn.clicked.connect(self._browse_single_input)
        input_dir_btn = QPushButton("Folder")
        input_dir_btn.clicked.connect(self._browse_input_folder)
        input_row.addWidget(self.input_path)
        input_row.addWidget(input_file_btn)
        input_row.addWidget(input_dir_btn)
        form.addRow("Source", input_row)

        output_row = QHBoxLayout()
        self.output_path = QLineEdit(str(Path.cwd() / "txt_output"))
        output_btn = QPushButton("Browse")
        output_btn.clicked.connect(self._browse_output_folder)
        output_row.addWidget(self.output_path)
        output_row.addWidget(output_btn)
        form.addRow("Output", output_row)
        return group

    def _build_options_group(self) -> QGroupBox:
        group = QGroupBox("Conversion Options")
        vbox = QVBoxLayout(group)

        ext_row = QHBoxLayout()
        self.extension_boxes: list[QCheckBox] = []
        for ext in DEFAULT_SUPPORTED_EXTENSIONS:
            box = QCheckBox(ext)
            box.setChecked(True)
            self.extension_boxes.append(box)
            ext_row.addWidget(box)
        vbox.addLayout(ext_row)

        options_row = QHBoxLayout()
        self.min_chars_spin = QSpinBox()
        self.min_chars_spin.setRange(1, 2000)
        self.min_chars_spin.setValue(30)
        self.min_plain_spin = QSpinBox()
        self.min_plain_spin.setRange(1, 10000)
        self.min_plain_spin.setValue(20)
        self.ocr_fallback_box = QCheckBox("Enable OCR fallback for low-text PDF pages")
        self.ocr_fallback_box.setChecked(True)
        options_row.addWidget(QLabel("Min page chars"))
        options_row.addWidget(self.min_chars_spin)
        options_row.addWidget(QLabel("Min output chars"))
        options_row.addWidget(self.min_plain_spin)
        options_row.addWidget(self.ocr_fallback_box)
        options_row.addStretch(1)
        vbox.addLayout(options_row)
        return group

    def _build_log_group(self) -> QGroupBox:
        group = QGroupBox("Logs")
        layout = QVBoxLayout(group)
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box)
        return group

    def _apply_styles(self) -> None:
        QApplication.setFont(QFont("Segoe UI", 10))
        self.setStyleSheet(
            """
            QMainWindow { background-color: #f8f9fb; }
            QGroupBox {
                font-weight: 600;
                border: 1px solid #d0d6e2;
                border-radius: 8px;
                margin-top: 12px;
                background: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px 0 6px;
                color: #2d3a4f;
            }
            QLineEdit, QSpinBox, QListWidget, QTextEdit {
                border: 1px solid #c9d3e0;
                border-radius: 6px;
                padding: 6px;
                background: #ffffff;
            }
            QListWidget::item { padding: 6px 8px; border-radius: 4px; }
            QListWidget::item:hover { background: #eef4ff; }
            QListWidget::item:selected {
                background: #1f6feb;
                color: white;
            }
            QListWidget::item:selected:hover { background: #1657b8; }
            QPushButton {
                background: #1f6feb;
                border: none;
                color: white;
                border-radius: 6px;
                padding: 7px 12px;
            }
            QPushButton:hover { background: #1657b8; }
            QLabel#titleLabel { font-size: 22px; font-weight: 700; color: #162033; }
            QLabel#subtitleLabel { color: #4d5b73; }
            """
        )

    def _browse_tesseract(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Select tesseract.exe", str(Path.home()), "Executable (*.exe)"
        )
        if path:
            self.tesseract_input.setText(path)

    def _browse_single_input(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "Select Input Document")
        if selected:
            self.input_path.setText(selected)

    def _browse_input_folder(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "Select Input Folder")
        if selected:
            self.input_path.setText(selected)

    def _browse_output_folder(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if selected:
            self.output_path.setText(selected)

    def _append_log(self, message: str) -> None:
        self.log_box.append(message)

    def _load_languages(self) -> None:
        path = self.tesseract_input.text().strip()
        if not path:
            QMessageBox.warning(self, "Missing Path", "Please provide tesseract.exe path.")
            return
        langs = get_tesseract_languages(Path(path))
        self._all_languages = langs
        self.lang_list.clear()
        if not langs:
            QMessageBox.critical(
                self,
                "Language Detection Failed",
                "Unable to load languages from the selected Tesseract path.",
            )
            return
        for lang in langs:
            item = QListWidgetItem(lang)
            self.lang_list.addItem(item)
            if lang in {"eng", "kor"}:
                item.setSelected(True)
        self._restore_selected_languages()
        self._filter_languages(self.language_search.text())
        self._append_log(f"[INFO] Loaded Tesseract with {len(langs)} language packs.")

    def _selected_languages(self) -> list[str]:
        selected: list[str] = []
        for index in range(self.lang_list.count()):
            item = self.lang_list.item(index)
            if item.isSelected():
                selected.append(item.text())
        return selected

    def _filter_languages(self, text: str) -> None:
        query = text.strip().lower()
        for index in range(self.lang_list.count()):
            item = self.lang_list.item(index)
            if not query:
                item.setHidden(False)
                continue
            item.setHidden(query not in item.text().lower())

    def _selected_extensions(self) -> list[str]:
        return [box.text() for box in self.extension_boxes if box.isChecked()]

    def _validate_before_run(self) -> tuple[Path, Path, ConversionConfig] | None:
        input_text = self.input_path.text().strip()
        output_text = self.output_path.text().strip()
        tesseract_text = self.tesseract_input.text().strip()
        if not input_text:
            QMessageBox.warning(self, "Missing Input", "Select a file or folder to convert.")
            return None
        if not output_text:
            QMessageBox.warning(self, "Missing Output", "Select output folder.")
            return None
        if not tesseract_text:
            QMessageBox.warning(
                self, "Missing OCR Engine", "Select tesseract.exe and load languages first."
            )
            return None

        selected_langs = self._selected_languages()
        if not selected_langs:
            QMessageBox.warning(self, "No Languages", "Select at least one OCR language.")
            return None

        selected_exts = self._selected_extensions()
        if not selected_exts:
            QMessageBox.warning(self, "No Extensions", "Select at least one extension.")
            return None

        config = ConversionConfig(
            tesseract_path=Path(tesseract_text),
            ocr_languages=selected_langs,
            supported_extensions=selected_exts,
            min_text_chars_for_page=self.min_chars_spin.value(),
            enable_ocr_fallback=self.ocr_fallback_box.isChecked(),
            min_plain_text_length=self.min_plain_spin.value(),
        )
        return Path(input_text), Path(output_text), config

    def _start_conversion(self) -> None:
        validated = self._validate_before_run()
        if not validated:
            return
        input_path, output_path, config = validated
        self.log_box.clear()
        self.progress_bar.setValue(0)
        self.start_button.setEnabled(False)

        self._thread = QThread(self)
        self._worker = ConversionWorker(input_path, output_path, config)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._append_log)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    def _on_progress(self, current: int, total: int) -> None:
        if total <= 0:
            return
        self.progress_bar.setValue(int((current / total) * 100))

    def _on_finished(self, processed: int, skipped: int, errors: int) -> None:
        self.start_button.setEnabled(True)
        self._append_log(
            f"[DONE] Processed={processed}, Skipped={skipped}, Errors={errors}"
        )
        QMessageBox.information(
            self,
            "Conversion Complete",
            f"Processed: {processed}\nSkipped: {skipped}\nErrors: {errors}",
        )

    def _on_failed(self, message: str) -> None:
        self.start_button.setEnabled(True)
        self._append_log(f"[FATAL] {message}")
        QMessageBox.critical(self, "Conversion Failed", message)

    def _open_output_folder(self) -> None:
        output_path = Path(self.output_path.text().strip()) if self.output_path.text().strip() else None
        if output_path is None:
            QMessageBox.warning(self, "Missing Output", "No output folder is configured.")
            return
        output_path.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_path.resolve())))

    def _show_about(self) -> None:
        QMessageBox.information(
            self,
            "About",
            "Professional Document Corpus Converter\n"
            "Converts PDF/DOC/DOCX/images into raw corpus text files.",
        )

    def _save_settings(self) -> None:
        self._settings.setValue("window/size", self.size())
        self._settings.setValue("window/pos", self.pos())
        self._settings.setValue("paths/tesseract", self.tesseract_input.text().strip())
        self._settings.setValue("paths/input", self.input_path.text().strip())
        self._settings.setValue("paths/output", self.output_path.text().strip())
        self._settings.setValue("options/min_page_chars", self.min_chars_spin.value())
        self._settings.setValue("options/min_output_chars", self.min_plain_spin.value())
        self._settings.setValue("options/ocr_fallback", self.ocr_fallback_box.isChecked())
        self._settings.setValue("options/extensions", self._selected_extensions())
        self._settings.setValue("options/languages", self._selected_languages())
        self._settings.sync()

    def _load_settings(self) -> None:
        self._is_loading_saved_state = True
        size = self._settings.value("window/size")
        pos = self._settings.value("window/pos")
        if size:
            self.resize(size)
        if pos:
            self.move(pos)

        self.tesseract_input.setText(self._settings.value("paths/tesseract", "", type=str))
        self.input_path.setText(self._settings.value("paths/input", "", type=str))
        stored_output = self._settings.value("paths/output", "", type=str)
        if stored_output:
            self.output_path.setText(stored_output)

        self.min_chars_spin.setValue(self._settings.value("options/min_page_chars", 30, type=int))
        self.min_plain_spin.setValue(
            self._settings.value("options/min_output_chars", 20, type=int)
        )
        self.ocr_fallback_box.setChecked(
            self._settings.value("options/ocr_fallback", True, type=bool)
        )

        stored_extensions = self._settings.value("options/extensions", [])
        if isinstance(stored_extensions, str):
            stored_extensions = [stored_extensions]
        if stored_extensions:
            for box in self.extension_boxes:
                box.setChecked(box.text() in stored_extensions)

        if self.tesseract_input.text().strip():
            self._load_languages()

        self._is_loading_saved_state = False

    def _restore_selected_languages(self) -> None:
        stored_languages = self._settings.value("options/languages", [])
        if isinstance(stored_languages, str):
            stored_languages = [stored_languages]
        if not stored_languages:
            return
        for index in range(self.lang_list.count()):
            item = self.lang_list.item(index)
            item.setSelected(item.text() in stored_languages)

    def closeEvent(self, event) -> None:  # type: ignore[override]
        self._save_settings()
        super().closeEvent(event)

