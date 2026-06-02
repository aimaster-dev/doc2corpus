"""Runtime environment validation helpers for deployment diagnostics."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app.config import ConversionConfig
from app.services.extractors import (
    find_libreoffice_executable,
    find_wps_executable,
    get_tesseract_languages,
    probe_com_prog_id,
)


@dataclass(slots=True)
class CheckResult:
    name: str
    ok: bool
    detail: str


def run_system_checks(config: ConversionConfig) -> list[CheckResult]:
    results: list[CheckResult] = []

    tesseract_path = config.tesseract_path
    if tesseract_path is None:
        results.append(CheckResult("Tesseract Path", False, "Not configured"))
    else:
        exists = Path(tesseract_path).is_file()
        results.append(
            CheckResult(
                "Tesseract Path",
                exists,
                f"{tesseract_path}" if exists else f"File not found: {tesseract_path}",
            )
        )

    detected_langs = set(get_tesseract_languages(tesseract_path)) if tesseract_path else set()
    if detected_langs:
        missing = [lang for lang in config.ocr_languages if lang not in detected_langs]
        if missing:
            results.append(
                CheckResult(
                    "OCR Languages",
                    False,
                    f"Missing language packs: {', '.join(missing)}",
                )
            )
        else:
            results.append(
                CheckResult(
                    "OCR Languages",
                    True,
                    f"Ready ({len(config.ocr_languages)} selected)",
                )
            )
    else:
        results.append(
            CheckResult("OCR Languages", False, "Cannot detect languages from Tesseract")
        )

    for prog_id in config.office_com_prog_ids:
        ok, detail = probe_com_prog_id(prog_id)
        results.append(
            CheckResult(
                f"COM {prog_id}",
                ok,
                detail if ok else f"Unavailable: {detail}",
            )
        )

    wps_path = find_wps_executable()
    results.append(
        CheckResult(
            "WPS Executable",
            wps_path is not None,
            wps_path or "Not found",
        )
    )

    libreoffice_path = find_libreoffice_executable()
    results.append(
        CheckResult(
            "LibreOffice Executable",
            libreoffice_path is not None,
            libreoffice_path or "Not found (optional fallback)",
        )
    )

    return results

