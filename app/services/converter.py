"""High-level conversion service for single and batch jobs."""

from __future__ import annotations

from collections.abc import Callable, Iterable
from pathlib import Path

from app.config import ConversionConfig
from app.services.extractors import (
    configure_tesseract,
    extract_doc_pages_text,
    extract_docx_pages_text,
    extract_image_text,
    extract_pdf_pages_text,
)
from app.services.text_processing import corpus_plain_text, pages_to_corpus

LogFn = Callable[[str], None]
ProgressFn = Callable[[int, int], None]


def collect_input_files(path: Path, extensions: Iterable[str]) -> list[Path]:
    normalized = {ext.lower() for ext in extensions}
    if path.is_file():
        return [path] if path.suffix.lower() in normalized else []
    files: list[Path] = []
    for ext in normalized:
        files.extend(path.glob(f"*{ext}"))
        files.extend(path.glob(f"*{ext.upper()}"))
    return sorted(set(files), key=lambda item: item.name.lower())


def convert_one_file_to_corpus(file_path: Path, config: ConversionConfig) -> str:
    suffix = file_path.suffix.lower()
    if suffix == ".pdf":
        pages_text = extract_pdf_pages_text(file_path, config)
    elif suffix == ".docx":
        pages_text = extract_docx_pages_text(file_path)
    elif suffix == ".doc":
        pages_text = extract_doc_pages_text(file_path)
    elif suffix in {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}:
        pages_text = extract_image_text(file_path, config.ocr_languages)
    else:
        raise ValueError(f"Unsupported file type: {suffix}")
    return pages_to_corpus(pages_text)


def run_conversion(
    input_path: Path,
    output_dir: Path,
    config: ConversionConfig,
    log: LogFn,
    on_progress: ProgressFn | None = None,
) -> tuple[int, int, int]:
    if not input_path.exists():
        raise FileNotFoundError(f"Input path does not exist: {input_path}")
    output_dir.mkdir(parents=True, exist_ok=True)
    configure_tesseract(config.tesseract_path)

    files = collect_input_files(input_path, config.supported_extensions)
    if not files:
        raise FileNotFoundError(
            "No supported files were found. "
            f"Supported extensions: {', '.join(sorted(config.supported_extensions))}"
        )

    processed_count = 0
    skipped_count = 0
    error_count = 0
    total = len(files)

    for index, file_path in enumerate(files, start=1):
        log(f"[PROCESSING] {file_path.name}")
        try:
            corpus = convert_one_file_to_corpus(file_path, config)
            if len(corpus_plain_text(corpus)) < config.min_plain_text_length:
                log(f"[SKIPPED] Too little extracted text: {file_path.name}")
                skipped_count += 1
            else:
                out_path = output_dir / f"{file_path.stem}.txt"
                out_path.write_text(corpus, encoding="utf-8")
                log(f"[WRITTEN] {out_path}")
                processed_count += 1
        except Exception as exc:
            error_count += 1
            log(f"[ERROR] {file_path.name}: {exc}")
        if on_progress:
            on_progress(index, total)

    return processed_count, skipped_count, error_count

