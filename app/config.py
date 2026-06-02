"""Application configuration models and defaults."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


DEFAULT_SUPPORTED_EXTENSIONS = [
    ".pdf",
    ".docx",
    ".doc",
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".tif",
    ".tiff",
]

DEFAULT_COM_PROG_IDS = [
    "Word.Application",
    "KWPS.Application",
    "wps.Application",
]


@dataclass(slots=True)
class ConversionConfig:
    """Runtime options for corpus conversion."""

    tesseract_path: Path | None = None
    ocr_languages: list[str] = field(default_factory=lambda: ["eng"])
    supported_extensions: list[str] = field(
        default_factory=lambda: DEFAULT_SUPPORTED_EXTENSIONS.copy()
    )
    min_text_chars_for_page: int = 30
    enable_ocr_fallback: bool = True
    ocr_dpi: int = 200
    min_plain_text_length: int = 20
    office_com_prog_ids: list[str] = field(
        default_factory=lambda: DEFAULT_COM_PROG_IDS.copy()
    )

