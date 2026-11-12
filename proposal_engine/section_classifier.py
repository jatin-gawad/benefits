from __future__ import annotations

import re
from typing import Any


STANDARD_SECTIONS = [
    "premium_summary",
    "rate_summary",
    "plan_design",
    "demographics",
    "underwriting",
    "pooling_retention",
]


def classify_sections(document: dict[str, Any], carrier_config: dict[str, Any]) -> list[dict[str, Any]]:
    """Classify PDF pages and Excel sheets into standard proposal sections."""

    sections: list[dict[str, Any]] = []
    keywords = carrier_config.get("section_keywords", {})

    for page in document.get("pages", []):
        text = page.get("text", "")
        tables = page.get("tables", [])
        combined = "\n".join([text, _tables_to_text(tables)])
        section_name, score = _best_section(combined, keywords)
        sections.append(
            {
                "section_name": section_name,
                "source_file": document.get("source_file"),
                "page_number": page.get("page_number"),
                "sheet_name": None,
                "text": text,
                "tables": tables,
                "confidence_score": score,
            }
        )

    for sheet in document.get("sheets", []):
        rows = sheet.get("rows", [])
        text = _rows_to_text(rows)
        sheet_name = sheet.get("sheet_name") or ""
        section_name, score = _best_section(f"{sheet_name}\n{text}", keywords)
        sections.append(
            {
                "section_name": section_name,
                "source_file": document.get("source_file"),
                "page_number": None,
                "sheet_name": sheet.get("sheet_name"),
                "text": text,
                "tables": [rows],
                "confidence_score": score,
            }
        )

    if not sections:
        sections.append(
            {
                "section_name": "unknown",
                "source_file": document.get("source_file"),
                "page_number": None,
                "sheet_name": None,
                "text": document.get("all_text", ""),
                "tables": [],
                "confidence_score": 0,
            }
        )

    return sections


def _best_section(text: str, section_keywords: dict[str, list[str]]) -> tuple[str, int]:
    normalized = _normalize(text)
    best_name = "unknown"
    best_score = 0

    for section_name in STANDARD_SECTIONS:
        score = 0
        for keyword in section_keywords.get(section_name, []):
            if _normalize(keyword) in normalized:
                score += max(1, len(keyword.split()))
        if score > best_score:
            best_name = section_name
            best_score = score

    return best_name, best_score


def _tables_to_text(tables: list[list[list[Any]]]) -> str:
    lines: list[str] = []
    for table in tables:
        lines.append(_rows_to_text(table))
    return "\n".join(lines)


def _rows_to_text(rows: list[list[Any]]) -> str:
    lines: list[str] = []
    for row in rows:
        line = " ".join(str(cell).strip() for cell in row if str(cell).strip())
        if line:
            lines.append(line)
    return "\n".join(lines)


def _normalize(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").lower()
