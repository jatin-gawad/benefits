from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
import pdfplumber
from openpyxl import load_workbook


def read_document(file_path: str | Path) -> dict[str, Any]:
    """Read a PDF, Excel workbook, or CSV into one standard document object."""

    path = Path(file_path)
    extension = path.suffix.lower()
    if not path.exists():
        raise FileNotFoundError(f"Proposal file was not found: {path}")

    if extension == ".pdf":
        return _read_pdf(path)
    if extension in {".xlsx", ".xlsm", ".xls"}:
        return _read_excel(path)
    if extension == ".csv":
        return _read_csv(path)

    raise ValueError(f"Unsupported proposal file type: {extension}")


def _read_pdf(path: Path) -> dict[str, Any]:
    pages: list[dict[str, Any]] = []
    errors: list[str] = []

    with pdfplumber.open(str(path)) as pdf:
        for index, page in enumerate(pdf.pages, start=1):
            text = ""
            tables: list[list[list[str]]] = []
            try:
                text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            except Exception as exc:  # pragma: no cover - defensive for unusual PDFs
                errors.append(f"Page {index} text extraction failed: {exc}")

            try:
                raw_tables = page.extract_tables() or []
                tables = [_clean_table(table) for table in raw_tables]
            except Exception as exc:
                errors.append(f"Page {index} table extraction failed: {exc}")

            pages.append(
                {
                    "page_number": index,
                    "text": text,
                    "tables": tables,
                    "extraction_method": "pdfplumber"
                }
            )

    all_text = "\n".join(page["text"] for page in pages if page.get("text"))
    return {
        "source_file": str(path),
        "source_file_name": path.name,
        "file_type": "pdf",
        "pages": pages,
        "sheets": [],
        "all_text": all_text,
        "errors": errors,
        "extraction_method": "pdfplumber"
    }


def _read_excel(path: Path) -> dict[str, Any]:
    errors: list[str] = []
    sheets: list[dict[str, Any]] = []

    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        workbook = load_workbook(path, data_only=True, read_only=True)
        for worksheet in workbook.worksheets:
            rows = [[_cell_value(value) for value in row] for row in worksheet.iter_rows(values_only=True)]
            sheets.append(
                {
                    "sheet_name": worksheet.title,
                    "rows": rows,
                    "extraction_method": "openpyxl"
                }
            )
    else:
        try:
            excel_file = pd.ExcelFile(path)
            for sheet_name in excel_file.sheet_names:
                frame = excel_file.parse(sheet_name=sheet_name, header=None, dtype=object)
                rows = frame.fillna("").astype(str).values.tolist()
                sheets.append(
                    {
                        "sheet_name": sheet_name,
                        "rows": rows,
                        "extraction_method": "pandas"
                    }
                )
        except Exception as exc:
            errors.append(f"Excel extraction failed: {exc}")

    all_text = _rows_to_text(sheet["rows"] for sheet in sheets)
    return {
        "source_file": str(path),
        "source_file_name": path.name,
        "file_type": "excel",
        "pages": [],
        "sheets": sheets,
        "all_text": all_text,
        "errors": errors,
        "extraction_method": "openpyxl/pandas"
    }


def _read_csv(path: Path) -> dict[str, Any]:
    frame = pd.read_csv(path, header=None, dtype=object).fillna("")
    rows = frame.astype(str).values.tolist()
    return {
        "source_file": str(path),
        "source_file_name": path.name,
        "file_type": "excel",
        "pages": [],
        "sheets": [
            {
                "sheet_name": path.stem,
                "rows": rows,
                "extraction_method": "pandas_csv"
            }
        ],
        "all_text": _rows_to_text([rows]),
        "errors": [],
        "extraction_method": "pandas_csv"
    }


def _clean_table(table: list[list[Any]]) -> list[list[str]]:
    return [[_cell_value(cell) for cell in row] for row in table]


def _cell_value(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _rows_to_text(sheet_rows: Any) -> str:
    lines: list[str] = []
    for rows in sheet_rows:
        for row in rows:
            line = " ".join(_cell_value(cell) for cell in row if _cell_value(cell))
            if line:
                lines.append(line)
    return "\n".join(lines)
