from __future__ import annotations

import argparse
from pathlib import Path

from proposal_engine import DEFAULT_DATABASE_PATH, DEFAULT_OUTPUT_DIR
from proposal_engine.main import process_proposal
from proposal_engine.sqlite_loader import sqlite_preview


def run_test(file_path: str | Path, output_folder: str | Path | None = None, database_path: str | Path | None = None) -> dict:
    out_dir = Path(output_folder) if output_folder else DEFAULT_OUTPUT_DIR
    return process_proposal(
        file_path,
        output_folder=out_dir,
        database_path=database_path or DEFAULT_DATABASE_PATH,
        write_outputs=True,
        test_mode=True,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Test one carrier proposal extraction.")
    parser.add_argument("file_path", help="Single proposal PDF, Excel workbook, or CSV.")
    parser.add_argument("--output-folder", dest="output_folder")
    parser.add_argument("--database-path", dest="database_path")
    args = parser.parse_args()

    result = run_test(args.file_path, args.output_folder, args.database_path)
    diagnostics = result.get("_diagnostics", {})
    market_stats = result.get("_market_survey_stats", {})
    missing = [item for item in result.get("exceptions", []) if item.get("status") == "MISSING"]

    print(f"Detected carrier: {diagnostics.get('detected_carrier')}")
    print(f"Config used: {diagnostics.get('config_used')}")
    print(f"Detection confidence: {diagnostics.get('detection_confidence')}")
    print(f"Sections found: {diagnostics.get('sections_found')}")
    print(f"Tables found: {diagnostics.get('tables_found')}")
    print(
        "Fields extracted: "
        f"benefit_lines={len(result.get('benefit_lines', []))}, "
        f"rates={len(result.get('rates', []))}, "
        f"premiums={len(result.get('premiums', []))}, "
        f"plan_design={len(result.get('plan_design', []))}"
    )
    print(f"Fields missing: {len(missing)}")
    print(f"Rows populated in Total Cost Comparison: {market_stats.get('total_cost_rows', 0)}")
    print(f"Rows populated in Plan Design: {market_stats.get('plan_design_rows', 0)}")
    print(f"Rows populated in Rate Guarantee Sum: {market_stats.get('rate_guarantee_rows', 0)}")
    print(f"Confidence score: {diagnostics.get('detection_confidence')}")
    print(f"Manual review items: {market_stats.get('manual_review_items', len(result.get('exceptions', [])))}")
    print(f"Exceptions: {len(result.get('exceptions', []))}")
    print(f"Workbook: {diagnostics.get('workbook_path')}")
    print(f"SQLite: {diagnostics.get('database_path')}")
    print(f"SQLite preview: {sqlite_preview(diagnostics.get('database_path') or DEFAULT_DATABASE_PATH)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
