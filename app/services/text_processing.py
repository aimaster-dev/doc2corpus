"""Text normalization and corpus formatting helpers."""

from __future__ import annotations

import re
import unicodedata

SEP_TOKEN = "[SEP]"
MATH_START_TOKEN = "[MATH]"
MATH_END_TOKEN = "[/MATH]"


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = unicodedata.normalize("NFC", text)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s+([.,!?;:%])", r"\1", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def wrap_math_regions(text: str) -> str:
    if not text:
        return text
    text = re.sub(
        r"(\\\(.+?\\\))",
        rf"{MATH_START_TOKEN} \1 {MATH_END_TOKEN}",
        text,
        flags=re.DOTALL,
    )
    text = re.sub(
        r"(\\\[.+?\\\])",
        rf"{MATH_START_TOKEN} \1 {MATH_END_TOKEN}",
        text,
        flags=re.DOTALL,
    )
    text = re.sub(
        r"(?<!\w)\$(?!\s)(.{2,120}?)(?<!\s)\$(?!\w)",
        rf"{MATH_START_TOKEN} $\1$ {MATH_END_TOKEN}",
        text,
        flags=re.DOTALL,
    )
    latex_keywords = r"(\\frac|\\sqrt|\\sum|\\int|\\lim|\\alpha|\\beta|\\gamma|\\theta|\\pi|\\Delta)"
    new_lines: list[str] = []
    for line in text.split("\n"):
        if re.search(latex_keywords, line) and MATH_START_TOKEN not in line:
            line = f"{MATH_START_TOKEN} {line} {MATH_END_TOKEN}"
        new_lines.append(line.strip())
    return "\n".join(new_lines)


def remove_page_noise_lines(lines: list[str]) -> list[str]:
    cleaned: list[str] = []
    for line in lines:
        line = line.strip()
        if not line:
            cleaned.append("")
            continue
        if re.fullmatch(r"\d+", line):
            continue
        if re.fullmatch(r"[-–—]\s*\d+\s*[-–—]", line):
            continue
        if re.fullmatch(r"Page\s+\d+", line, flags=re.IGNORECASE):
            continue
        if re.match(r"^(Figure|Fig\.|Table)\s+\d+", line, flags=re.IGNORECASE):
            continue
        cleaned.append(line)
    return cleaned


def split_into_paragraphs(text: str) -> list[str]:
    text = clean_text(text)
    text = wrap_math_regions(text)
    if not text:
        return []

    lines = remove_page_noise_lines(text.split("\n"))
    paragraphs: list[str] = []
    buffer: list[str] = []
    for line in lines:
        line = line.strip()
        if not line:
            if buffer:
                paragraphs.append(" ".join(buffer).strip())
                buffer = []
            continue
        is_heading = bool(
            re.match(
                r"^(\d+\.|\d+\)|제\s*\d+\s*장|Chapter\s+\d+|CHAPTER\s+\d+)",
                line,
                flags=re.IGNORECASE,
            )
        )
        if is_heading:
            if buffer:
                paragraphs.append(" ".join(buffer).strip())
                buffer = []
            paragraphs.append(line)
            continue
        buffer.append(line)

    if buffer:
        paragraphs.append(" ".join(buffer).strip())
    return [p for p in paragraphs if len(p) >= 2]


def pages_to_corpus(pages_text: list[str]) -> str:
    lines: list[str] = []
    for page_text in pages_text:
        if not page_text.strip():
            continue
        for paragraph in split_into_paragraphs(page_text):
            lines.append(f"{paragraph} {SEP_TOKEN}")
    return "\n".join(lines)


def corpus_plain_text(corpus_text: str) -> str:
    return (
        corpus_text.replace(SEP_TOKEN, "")
        .replace(MATH_START_TOKEN, "")
        .replace(MATH_END_TOKEN, "")
        .strip()
    )

