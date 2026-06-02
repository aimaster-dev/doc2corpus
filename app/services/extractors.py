"""File type extractors and OCR integration."""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
import json
from pathlib import Path

import fitz
import pdfplumber
import pytesseract
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image

from app.config import ConversionConfig
from app.services.text_processing import clean_text


def _co_initialize() -> object | None:
    """Initialize COM for the current thread when pywin32 is available."""
    try:
        import pythoncom  # type: ignore

        pythoncom.CoInitialize()
        return pythoncom
    except Exception:
        return None


def _co_uninitialize(pythoncom_module: object | None) -> None:
    if pythoncom_module is None:
        return
    try:
        pythoncom_module.CoUninitialize()
    except Exception:
        pass


def probe_com_prog_id(prog_id: str) -> tuple[bool, str]:
    """Probe whether a COM ProgID is dispatchable in this environment."""
    if not prog_id.strip():
        return False, "Empty COM ProgID"

    # Use isolated subprocess probe to avoid UI hangs when COM dispatch stalls.
    if getattr(sys, "frozen", False):
        command = [
            sys.executable,
            "--probe-com-worker",
            "--prog-id",
            prog_id,
        ]
    else:
        command = [
            sys.executable,
            "-m",
            "main",
            "--probe-com-worker",
            "--prog-id",
            prog_id,
        ]

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=False,
            timeout=20,
        )
        output = (result.stdout or "").strip()
        if not output:
            detail = (result.stderr or "").strip() or f"Probe exit code {result.returncode}"
            return False, detail
        payload = json.loads(output)
        if payload.get("ok"):
            return True, str(payload.get("detail", "available"))
        return False, str(payload.get("error", "Unknown probe error"))
    except Exception as exc:
        return False, str(exc)


def configure_tesseract(path: Path | None) -> None:
    if path:
        pytesseract.pytesseract.tesseract_cmd = str(path)


def get_tesseract_languages(path: Path | None) -> list[str]:
    configure_tesseract(path)
    try:
        langs = sorted(set(pytesseract.get_languages(config="")))
        return [lang for lang in langs if lang]
    except Exception:
        return []


def table_to_text_rows(table: list[list[str | None]] | None) -> list[str]:
    if not table or len(table) < 2:
        return []
    cleaned_table: list[list[str]] = []
    for row in table:
        cleaned_row: list[str] = []
        for cell in row:
            cleaned_row.append(clean_text("" if cell is None else str(cell)))
        cleaned_table.append(cleaned_row)

    header = cleaned_table[0]
    data_rows = cleaned_table[1:]
    result_rows: list[str] = []
    for row in data_rows:
        parts: list[str] = []
        for i, value in enumerate(row):
            if not value:
                continue
            if i < len(header) and header[i]:
                parts.append(f"{header[i]}: {value}")
            else:
                parts.append(value)
        row_text = " | ".join(parts).strip()
        if row_text:
            result_rows.append(row_text)
    return result_rows


def extract_text_pages_with_pdfplumber(pdf_path: Path) -> list[str]:
    pages_text: list[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            tables = page.extract_tables()
            for table in tables:
                for row_text in table_to_text_rows(table):
                    text += f"\n{row_text}"
            pages_text.append(text)
    return pages_text


def ocr_single_pdf_page(
    pdf_path: Path, page_index: int, ocr_languages: list[str], ocr_dpi: int
) -> str:
    doc = fitz.open(str(pdf_path))
    try:
        page = doc.load_page(page_index)
        zoom = ocr_dpi / 72
        matrix = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        lang = "+".join(ocr_languages) if ocr_languages else "eng"
        return pytesseract.image_to_string(image, lang=lang) or ""
    finally:
        doc.close()


def extract_pdf_pages_text(pdf_path: Path, config: ConversionConfig) -> list[str]:
    pages_text = extract_text_pages_with_pdfplumber(pdf_path)
    final_pages: list[str] = []
    for i, page_text in enumerate(pages_text):
        cleaned = clean_text(page_text)
        if len(cleaned) >= config.min_text_chars_for_page:
            final_pages.append(cleaned)
            continue
        if config.enable_ocr_fallback:
            ocr_text = ocr_single_pdf_page(
                pdf_path, i, config.ocr_languages, config.ocr_dpi
            )
            final_pages.append(clean_text(ocr_text))
        else:
            final_pages.append("")
    return final_pages


def iter_docx_blocks(parent: Document):
    for child in parent.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def docx_table_to_rows(table: Table) -> list[str]:
    rows = [[cell.text for cell in row.cells] for row in table.rows]
    return table_to_text_rows(rows)


def extract_docx_pages_text(docx_path: Path) -> list[str]:
    lines: list[str] = []
    doc = Document(str(docx_path))
    for block in iter_docx_blocks(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text:
                lines.append(text)
        else:
            lines.extend(docx_table_to_rows(block))

    full_text = "\n".join(lines)
    if "\f" in full_text:
        return [clean_text(part) for part in full_text.split("\f") if part.strip()]
    return [clean_text(full_text)] if full_text.strip() else []


def find_libreoffice_executable() -> str | None:
    candidates = [
        "soffice",
        "libreoffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for name in candidates:
        path = shutil.which(name) if not Path(name).is_file() else name
        if path and Path(path).is_file():
            return path
    return None


def find_wps_executable() -> str | None:
    candidates = [
        "wps",
        "wps.exe",
        r"C:\Program Files\WPS Office\12.1.0.20305\office6\wps.exe",
        r"C:\Program Files\WPS Office\11.2.0.11626\office6\wps.exe",
        r"C:\Program Files\WPS Office\office6\wps.exe",
        r"C:\Program Files (x86)\WPS Office\office6\wps.exe",
        r"C:\Program Files\Kingsoft\WPS Office\office6\wps.exe",
        r"C:\Program Files (x86)\Kingsoft\WPS Office\office6\wps.exe",
    ]
    for name in candidates:
        path = shutil.which(name) if not Path(name).is_file() else name
        if path and Path(path).is_file():
            return path
    return None


def try_register_wps_com() -> str | None:
    """Attempt to register WPS COM server and return status text."""
    wps_path = find_wps_executable()
    if not wps_path:
        return None
    try:
        result = subprocess.run(
            [wps_path, "/regserver"],
            capture_output=True,
            text=True,
            check=False,
            timeout=30,
        )
        output = (result.stdout or "").strip() or (result.stderr or "").strip()
        if result.returncode == 0:
            return output or "WPS COM registration completed."
        return output or f"WPS /regserver returned {result.returncode}."
    except Exception as exc:
        return f"WPS /regserver failed: {exc}"


def extract_doc_with_libreoffice(doc_path: Path) -> str:
    soffice = find_libreoffice_executable()
    if not soffice:
        raise RuntimeError("LibreOffice not found (install LibreOffice for .doc support)")
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        result = subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "txt:Text",
                "--outdir",
                str(tmp_dir_path),
                str(doc_path.resolve()),
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr or result.stdout or "LibreOffice conversion failed")
        txt_path = tmp_dir_path / f"{doc_path.stem}.txt"
        if not txt_path.exists():
            raise RuntimeError(f"LibreOffice did not produce output: {txt_path}")
        return txt_path.read_text(encoding="utf-8", errors="ignore")


def extract_doc_with_word_com(doc_path: Path) -> str:
    import win32com.client  # type: ignore

    pythoncom_module = _co_initialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(str(doc_path.resolve()))
        return doc.Content.Text
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()
        _co_uninitialize(pythoncom_module)


def extract_doc_with_com(doc_path: Path, prog_ids: list[str]) -> str:
    """Try COM automation using a list of ProgIDs."""
    import win32com.client  # type: ignore

    if not prog_ids:
        raise RuntimeError("No COM ProgIDs were provided.")

    pythoncom_module = _co_initialize()
    last_error: Exception | None = None
    app = None
    try:
        for _ in range(2):
            for prog_id in prog_ids:
                try:
                    app = win32com.client.Dispatch(prog_id)
                    break
                except Exception as exc:
                    last_error = exc
                    app = None
            if app is not None:
                break
            # If first COM attempt failed, try registering WPS and retry once.
            try_register_wps_com()

        if app is None:
            tried = ", ".join(prog_ids)
            raise RuntimeError(
                f"COM automation is unavailable. Tried [{tried}]. Last error: {last_error}"
            )

        doc = None
        app.Visible = False
        doc = app.Documents.Open(str(doc_path.resolve()))
        return doc.Content.Text
    finally:
        try:
            if "doc" in locals() and doc is not None:
                doc.Close(False)
        finally:
            if app is not None:
                app.Quit()
            _co_uninitialize(pythoncom_module)


def extract_doc_pages_text(
    doc_path: Path,
    office_com_prog_ids: list[str] | None = None,
) -> list[str]:
    text: str | None = None
    errors: list[str] = []
    configured_prog_ids = office_com_prog_ids or [
        "Word.Application",
        "KWPS.Application",
        "wps.Application",
    ]

    try:
        text = extract_doc_with_com_subprocess(doc_path, configured_prog_ids)
    except Exception as exc:
        errors.append(f"Configured COM: {exc}")

    if text is None:
        try:
            text = extract_doc_with_libreoffice(doc_path)
        except Exception as exc:
            errors.append(f"LibreOffice: {exc}")

    if text is None:
        try:
            text = extract_doc_with_binary_fallback(doc_path)
        except Exception as exc:
            errors.append(f"Binary fallback: {exc}")

    if text is None:
        raise RuntimeError(
            "Could not read .doc file. Install Microsoft Word, WPS Office, or LibreOffice "
            "(with pywin32 for COM automation).\n"
            + "\n".join(errors)
        )

    text = text.replace("\r", "\n")
    if "\f" in text:
        return [clean_text(part) for part in text.split("\f") if part.strip()]
    return [clean_text(text)] if text.strip() else []


def extract_image_text(image_path: Path, ocr_languages: list[str]) -> list[str]:
    image = Image.open(str(image_path))
    lang = "+".join(ocr_languages) if ocr_languages else "eng"
    text = pytesseract.image_to_string(image, lang=lang) or ""
    return [clean_text(text)] if text.strip() else []


def extract_doc_with_binary_fallback(doc_path: Path) -> str:
    """Best-effort text extraction from legacy .doc bytes.

    This fallback is intentionally simple and used only when COM/libreoffice
    are unavailable or blocked. It may include noise, but preserves usable
    corpus content on locked-down systems.
    """
    raw = doc_path.read_bytes()
    if not raw:
        raise RuntimeError("Empty file.")

    candidates = []
    for encoding in ("utf-16le", "gb18030", "cp949", "utf-8", "latin-1"):
        try:
            decoded = raw.decode(encoding, errors="ignore")
        except Exception:
            continue
        decoded = decoded.replace("\x00", "")
        decoded = clean_text(decoded)
        if len(decoded) > 20:
            candidates.append(decoded)

    if not candidates:
        raise RuntimeError("Could not decode text from binary content.")

    def score(text: str) -> int:
        # Score by count of useful language chars.
        return sum(ch.isalnum() or ("\u4e00" <= ch <= "\u9fff") for ch in text)

    best = max(candidates, key=score)
    # Keep only lines likely to contain human-readable content.
    lines = []
    for line in best.splitlines():
        line = clean_text(line)
        if len(line) < 2:
            continue
        useful = sum(
            ch.isalnum() or ("\u4e00" <= ch <= "\u9fff") or ch in " .,;:!?-_/()[]{}"
            for ch in line
        )
        if useful / max(len(line), 1) >= 0.55:
            lines.append(line)

    text = "\n".join(lines).strip()
    if len(text) < 20:
        raise RuntimeError("Binary fallback produced too little usable text.")
    return text


def extract_doc_with_com_subprocess(doc_path: Path, prog_ids: list[str]) -> str:
    """Run COM extraction in isolated subprocess for EXE stability."""
    errors: list[str] = []
    for prog_id in prog_ids:
        if getattr(sys, "frozen", False):
            command = [
                sys.executable,
                "--doc-com-worker",
                "--doc-path",
                str(doc_path),
                "--prog-ids",
                prog_id,
            ]
        else:
            command = [
                sys.executable,
                "-m",
                "main",
                "--doc-com-worker",
                "--doc-path",
                str(doc_path),
                "--prog-ids",
                prog_id,
            ]

        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                timeout=45,
            )
        except subprocess.TimeoutExpired:
            errors.append(f"{prog_id}: timed out after 45 seconds")
            continue

        output = (result.stdout or "").strip()
        if not output:
            detail = result.stderr.strip() or f"worker exit code {result.returncode}"
            errors.append(f"{prog_id}: {detail}")
            continue

        try:
            payload = json.loads(output)
        except Exception:
            errors.append(f"{prog_id}: invalid worker output")
            continue

        if payload.get("ok"):
            return str(payload.get("text", ""))
        errors.append(f"{prog_id}: {payload.get('error', 'Unknown COM worker error')}")

    raise RuntimeError("; ".join(errors) if errors else "No COM ProgID attempts were executed")

