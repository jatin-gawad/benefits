from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from proposal_engine import DEFAULT_DATABASE_PATH, DEFAULT_OUTPUT_DIR, SUPPORTED_EXTENSIONS
from proposal_engine.booklet_validator import validate_against_booklet
from proposal_engine.carrier_config_loader import load_carrier_config
from proposal_engine.carrier_detector import detect_carrier
from proposal_engine.document_reader import read_document
from proposal_engine.excel_writer import write_to_excel
from proposal_engine.field_extractor import extract_fields
from proposal_engine.normalizer import combine_proposals, normalize_proposal_data
from proposal_engine.section_classifier import classify_sections
from proposal_engine.sqlite_loader import load_to_sqlite, sqlite_preview
from proposal_engine.validator import validate_proposal_data


def process_proposal(
    file_path: str | Path,
    client_id: str | None = None,
    booklet_path: str | Path | None = None,
    output_folder: str | Path | None = None,
    database_path: str | Path | None = None,
    write_outputs: bool = True,
    test_mode: bool = False,
) -> dict[str, Any]:
    """Read, detect, extract, normalize, validate, export, and store one proposal."""

    path = Path(file_path)
    document = read_document(path)
    detection = detect_carrier(document, path)
    carrier_config = load_carrier_config(detection.get("carrier_key"))
    sections = classify_sections(document, carrier_config)
    raw_data = extract_fields(sections, carrier_config)

    raw_data["proposal_summary"]["source_file"] = str(path)
    raw_data["proposal_summary"]["detected_carrier"] = detection.get("carrier_key")
    raw_data["proposal_summary"]["detection_confidence"] = detection.get("confidence")
    raw_data["proposal_summary"]["carrier_name"] = (
        raw_data["proposal_summary"].get("carrier_name")
        or detection.get("carrier_name")
        or carrier_config.get("carrier_name")
    )

    normalized_data = normalize_proposal_data(raw_data, carrier_config)
    normalized_data = validate_against_booklet(normalized_data, booklet_path)
    validated_data = validate_proposal_data(normalized_data)
    validated_data["client_id"] = client_id
    validated_data["_diagnostics"] = {
        "detected_carrier": detection.get("carrier_key"),
        "config_used": carrier_config.get("_config_path"),
        "detection_confidence": detection.get("confidence"),
        "sections_found": _section_counts(sections),
        "tables_found": _count_tables(sections),
        "document_errors": document.get("errors", []),
    }

    if write_outputs:
        out_dir = Path(output_folder) if output_folder else DEFAULT_OUTPUT_DIR
        out_dir.mkdir(parents=True, exist_ok=True)
        excel_name = "Carrier_Extraction_Test_Output.xlsx" if test_mode else "Carrier_Proposal_Comparison.xlsx"
        audit = load_to_sqlite(validated_data, database_path or DEFAULT_DATABASE_PATH, client_id=client_id)
        validated_data["sqlite_audit"] = audit
        workbook_path = write_to_excel(validated_data, out_dir / excel_name, test_mode=test_mode)
        validated_data["_diagnostics"]["workbook_path"] = workbook_path
        validated_data["_diagnostics"]["database_path"] = str(database_path or DEFAULT_DATABASE_PATH)

    return validated_data


def process_proposal_folder(
    folder_path: str | Path,
    client_id: str | None = None,
    booklet_path: str | Path | None = None,
    output_folder: str | Path | None = None,
    database_path: str | Path | None = None,
) -> dict[str, Any]:
    files = find_supported_files(folder_path)
    return _process_proposal_files(files, client_id, booklet_path, output_folder, database_path)


def find_supported_files(folder_path: str | Path) -> list[str]:
    folder = Path(folder_path)
    if not folder.exists():
        raise FileNotFoundError(f"Proposal folder was not found: {folder}")

    files: list[str] = []
    for path in sorted(folder.rglob("*")):
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
            files.append(str(path))
    return files


def run_proposal_comparison(
    folder_path: str | Path,
    client_id: str | None = None,
    booklet_path: str | Path | None = None,
    output_folder: str | Path | None = None,
    database_path: str | Path | None = None,
) -> dict[str, Any]:
    return process_proposal_folder(folder_path, client_id, booklet_path, output_folder, database_path)


def run_proposal_comparison_for_files(
    file_paths: list[str | Path],
    client_id: str | None = None,
    booklet_path: str | Path | None = None,
    output_folder: str | Path | None = None,
    database_path: str | Path | None = None,
) -> dict[str, Any]:
    files = [str(Path(path)) for path in file_paths if Path(path).suffix.lower() in SUPPORTED_EXTENSIONS]
    return _process_proposal_files(files, client_id, booklet_path, output_folder, database_path)


def _process_proposal_files(
    file_paths: list[str],
    client_id: str | None,
    booklet_path: str | Path | None,
    output_folder: str | Path | None,
    database_path: str | Path | None,
) -> dict[str, Any]:
    proposals: list[dict[str, Any]] = []
    exceptions: list[dict[str, Any]] = []

    for file_path in file_paths:
        try:
            proposals.append(
                process_proposal(
                    file_path,
                    client_id=client_id,
                    booklet_path=booklet_path,
                    output_folder=output_folder,
                    database_path=database_path,
                    write_outputs=False,
                )
            )
        except Exception as exc:
            exceptions.append(
                {
                    "carrier_name": "Unknown Carrier",
                    "exception_type": "NEEDS REVIEW",
                    "field_name": "source_file",
                    "message": f"{Path(file_path).name} could not be processed: {exc}",
                    "source_file": str(file_path),
                    "source_page": None,
                    "source_sheet": None,
                    "confidence_score": "LOW",
                    "status": "NEEDS REVIEW",
                    "notes": "Processing continued for remaining files.",
                }
            )

    combined = combine_proposals(proposals, client_id=client_id)
    combined["exceptions"].extend(exceptions)
    if not proposals:
        combined = validate_proposal_data(combined)

    out_dir = Path(output_folder) if output_folder else DEFAULT_OUTPUT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    audit = load_to_sqlite(combined, database_path or DEFAULT_DATABASE_PATH, client_id=client_id)
    combined["sqlite_audit"] = audit
    workbook_path = write_to_excel(combined, out_dir / "Carrier_Proposal_Comparison.xlsx")

    combined["_diagnostics"] = {
        "proposal_count": len(proposals),
        "failed_file_count": len(exceptions),
        "workbook_path": workbook_path,
        "database_path": str(database_path or DEFAULT_DATABASE_PATH),
        "sqlite_preview": sqlite_preview(database_path or DEFAULT_DATABASE_PATH),
    }
    return combined


def _section_counts(sections: list[dict[str, Any]]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for section in sections:
        name = section.get("section_name", "unknown")
        counts[name] = counts.get(name, 0) + 1
    return counts


def _count_tables(sections: list[dict[str, Any]]) -> int:
    total = 0
    for section in sections:
        total += len(section.get("tables", []))
    return total


def _cli_summary(data: dict[str, Any]) -> dict[str, Any]:
    diagnostics = data.get("_diagnostics", {})
    summaries = data.get("proposal_summaries") or [data.get("proposal_summary", {})]
    return {
        "proposal_count": len(summaries),
        "carriers": [summary.get("carrier_name") for summary in summaries],
        "workbook_path": diagnostics.get("workbook_path"),
        "database_path": diagnostics.get("database_path"),
        "exceptions": len(data.get("exceptions", [])),
        "market_survey_stats": data.get("_market_survey_stats", {}),
        "sqlite_preview": diagnostics.get("sqlite_preview") or sqlite_preview(diagnostics.get("database_path") or DEFAULT_DATABASE_PATH),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Run dynamic carrier proposal comparison.")
    parser.add_argument("paths", nargs="+", help="Proposal folder or proposal file path(s).")
    parser.add_argument("--client-id", dest="client_id")
    parser.add_argument("--booklet", dest="booklet_path")
    parser.add_argument("--output-folder", dest="output_folder")
    parser.add_argument("--database-path", dest="database_path")
    parser.add_argument("--json", action="store_true", help="Print compact JSON output.")
    args = parser.parse_args()

    input_paths = [Path(path) for path in args.paths]
    if len(input_paths) == 1 and input_paths[0].is_dir():
        result = run_proposal_comparison(
            input_paths[0],
            client_id=args.client_id,
            booklet_path=args.booklet_path,
            output_folder=args.output_folder,
            database_path=args.database_path,
        )
    else:
        result = run_proposal_comparison_for_files(
            input_paths,
            client_id=args.client_id,
            booklet_path=args.booklet_path,
            output_folder=args.output_folder,
            database_path=args.database_path,
        )

    summary = _cli_summary(result)
    if args.json:
        print(json.dumps(summary, indent=2, default=str))
    else:
        print(f"Proposal count: {summary['proposal_count']}")
        print(f"Carriers: {', '.join(str(carrier or 'Unknown') for carrier in summary['carriers'])}")
        print(f"Exceptions: {summary['exceptions']}")
        print(f"Workbook: {summary['workbook_path']}")
        print(f"SQLite: {summary['database_path']}")
        print(f"SQLite preview: {summary['sqlite_preview']}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
