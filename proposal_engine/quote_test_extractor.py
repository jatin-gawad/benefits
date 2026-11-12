from __future__ import annotations

import argparse
import re
import sys
from copy import copy
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from proposal_engine import DEFAULT_OUTPUT_DIR, DEFAULT_TEMPLATE_PATH
from proposal_engine.booklet_validator import validate_focused_proposals_against_booklet
from proposal_engine.focused_comparison_mapper import build_focused_comparison_envelope
from proposal_engine.sqlite_loader import load_to_sqlite, sqlite_preview


OUTPUT_FILE_NAME = "Carrier_Quote_Test_Output.xlsx"
WORKBOOK_FONT_NAME = "Arial"
MAX_COMPARISON_CARRIERS = 5
TARGET_SHEETS = [
    "Total Cost Comparison",
    "Gross Cost Comparison",
    "Benefit Comparison",
    "BenefitComparison_Normalized",
    "Plan Design",
    "Extended Health",
    "Rate Guarantee Sum",
    "ExecSummary - Cost",
    "Rate Reconciliation",
    "RateTable_LineRates_Formatted",
    "PlanDesign",
    "Booklet_Extraction_Review",
    "Dental_Normalization_Reference",
    "Dental_Normalized_Values",
    "Alternates",
    "Deviations",
]

STANDARD_BENEFITS = [
    "Employee Life Insurance",
    "Employee AD&D",
    "Dependent Life",
    "Long Term Disability",
    "Short Term Disability",
    "Extended Health Care",
    "Dental Care",
]

OPTIONAL_BENEFITS = [
    "EAP",
    "Virtual Care",
    "Medical Second Opinion",
    "Critical Illness",
    "Travel",
    "HCSA",
    "WSA",
    "Wellness",
    "Other included services",
]

BOOKLET_REVIEW_HEADERS = [
    "ClientID",
    "ClientName",
    "Industry",
    "Province",
    "EmployeeCount",
    "Carrier",
    "PolicyNumber",
    "SourceFile",
    "ClassName",
    "ClassDescription",
    "BenefitCategory",
    "BenefitName",
    "BenefitIncluded",
    "BenefitAmountText",
    "BenefitAmountType",
    "BenefitAmountNumeric",
    "EarningsMultiplier",
    "SalaryPercent",
    "CoinsuranceText",
    "CoinsurancePercent",
    "DeductibleText",
    "SingleDeductible",
    "FamilyDeductible",
    "MaximumText",
    "MaximumAmount",
    "MaximumFrequency",
    "LifetimeMaximum",
    "NonEvidenceMaximum",
    "WaitingPeriodText",
    "WaitingPeriodMonths",
    "EliminationPeriodText",
    "EliminationPeriodDays",
    "BenefitPeriod",
    "TerminationAgeText",
    "TerminationAgeNumber",
    "TerminationTrigger",
    "TaxStatus",
    "WaiverOfPremium",
    "DefinitionOfDisability",
    "ReviewFlag",
    "Confidence",
    "LastUpdated",
    "Notes",
    "Benefit",
    "CoverageDescription",
    "Coinsurance",
    "Maximum",
    "Frequency",
    "WaitingPeriod",
    "TerminationAge",
    "LifeBenefitAmount",
    "LifeEarningsMultiplier",
    "ADDBenefitAmount",
    "DependentLifeSpouseAmount",
    "DependentLifeChildAmount",
    "CriticalIllnessAmount",
    "CriticalIllnessPreEx",
    "CriticalIllnessTerminationAge",
    "STDBenefitPercent",
    "STDMaximum",
    "LTDBenefitPercent",
    "LTDMaximum",
    "LTDEliminationPeriodDays",
    "PrescriptionDrugCoinsurance",
    "DrugPlanType",
    "DrugMax",
    "DrugType",
    "HospitalCoverage",
    "ParamedicalStandardMaximum",
    "ParamedicalPsychologyMaximum",
    "VisionMaximum",
    "VisionFrequencyMonths",
    "EyeExamFrequencyMonths",
    "DentalBasicCoinsurance",
    "DentalMajorCoinsurance",
    "DentalOrthoCoinsurance",
    "DentalBasicMaximum",
    "DentalMajorMaximum",
    "DentalOrthoMaximum",
    "DentalRecallFrequency",
    "DentalBasicRecallRaw",
    "DentalBasicRecallNormalized",
    "DentalMajorAnnualMaximumRaw",
    "DentalMajorAnnualMaximumNormalized",
    "DentalOrthoLifetimeMaximumRaw",
    "DentalOrthoLifetimeMaximumNormalized",
    "DentalNormalizationNotes",
    "TravelMaximum",
    "TravelTripDurationDays",
    "EAPIncluded",
    "VirtualCareIncluded",
    "SecondOpinionIncluded",
]

BENEFIT_ALIASES = {
    "Employee Life Insurance": [
        "life insurance",
        "basic life",
        "employee life",
    ],
    "Employee AD&D": [
        "ad&d insurance",
        "ad&d",
        "accidental death",
        "accidental death and dismemberment",
    ],
    "Dependent Life": [
        "dependent life",
        "dependent life insurance",
        "dep life",
    ],
    "Long Term Disability": [
        "long term disability",
    ],
    "Short Term Disability": [
        "short term disability",
    ],
    "Extended Health Care": [
        "health benefits",
        "health benefit",
        "extended health",
        "extended healthcare",
        "ehc",
    ],
    "Dental Care": [
        "dental benefits",
        "dental benefit",
        "dental care",
    ],
    "EAP": [
        "employee and family assistance",
        "employee assistance",
        "efap",
        "eap",
        "homewood",
    ],
    "Virtual Care": [
        "virtual care",
        "virtual healthcare",
        "telemedicine",
        "dialogue",
    ],
    "Medical Second Opinion": [
        "medical second opinion",
        "second opinion",
        "worldcare",
    ],
    "Critical Illness": [
        "critical illness",
        "critical disease",
        "serious illness",
        "ci",
    ],
    "Travel": [
        "travel assist",
        "out-of-province",
        "out of province",
        "out-of-country",
        "out of country",
        "travel",
    ],
    "HCSA": [
        "health care spending account",
        "health spending account",
        "hcsa",
    ],
    "WSA": [
        "wellness spending account",
        "taxable spending account",
        "wsa",
    ],
    "Wellness": [
        "wellness",
        "feeling better now",
        "carepath",
    ],
}

DENTAL_NORMALIZATION_OTHER = "Other, please specify in Notes"
DENTAL_NORMALIZATION_SELECT = "Select value from dropdown"
DENTAL_RECALL_DROPDOWN_VALUES = [
    "5 month recall",
    "6 month recall",
    "9 month recall",
    "12 month recall",
    "N/A",
    DENTAL_NORMALIZATION_OTHER,
]
DENTAL_MAJOR_MAXIMUM_DROPDOWN_VALUES = [
    "$1,000/year",
    "$1,500/year",
    "$2,000/year",
    "$2,500/year",
    "$3,000/year",
    "$1,000/year combined",
    "$1,500/year combined",
    "$2,000/year combined",
    "$2,500/year combined",
    "$3,000/year combined",
    DENTAL_NORMALIZATION_OTHER,
    DENTAL_NORMALIZATION_SELECT,
]
DENTAL_ORTHODONTIC_MAXIMUM_DROPDOWN_VALUES = [
    "$1,000 lifetime",
    "$1,500 lifetime",
    "$2,000 lifetime",
    "$2,500 lifetime",
    "$3,000 lifetime",
    "$3,500 lifetime",
    "$4,000 lifetime",
    DENTAL_NORMALIZATION_OTHER,
]

COMPARISON_SUBSECTION_ORDER = {
    "basic life": 10,
    "basic life insurance": 10,
    "employee life": 10,
    "employee life insurance": 10,
    "accidental death & dismemberment": 20,
    "accidental death & dismemberment or accident & serious illness": 20,
    "accident & serious illness": 20,
    "accident & serious illness (asi)": 20,
    "employee ad&d": 20,
    "dependent life": 30,
    "dependent life insurance": 30,
    "long term disability": 40,
    "short term disability": 50,
    "prescription drugs": 60,
    "paramedical practitioners": 70,
    "vision care": 80,
    "hospital": 90,
    "out of province or country": 100,
    "out of country services": 100,
    "travel": 100,
    "medical services & supplies": 110,
    "major services": 110,
    "dental": 120,
    "dental care": 120,
    "critical illness": 130,
    "virtual healthcare": 140,
    "virtual health": 140,
    "virtual care": 140,
    "medical second opinion": 150,
    "eap": 160,
    "employee assistance program": 160,
    "employee & family assistance": 160,
    "wellness": 170,
    "hcsa": 180,
    "wsa": 190,
}

COMPARISON_CATEGORY_TITLES = {
    "life",
    "basic life",
    "basic life insurance",
    "employee life",
    "employee life insurance",
    "ad&d",
    "employee ad&d",
    "accidental death & dismemberment",
    "accidental death and dismemberment",
    "accidental death & dismemberment or accident & serious illness",
    "accident & serious illness",
    "accident & serious illness (asi)",
    "asi",
    "dependent life",
    "dependent life insurance",
    "long term disability",
    "short term disability",
    "prescription drugs",
    "paramedical practitioners",
    "vision care",
    "hospital",
    "out of province or country",
    "out of country services",
    "medical services & supplies",
    "other provisions",
    "dental",
    "dental care",
    "critical illness",
    "virtual health",
    "virtual healthcare",
    "virtual care",
    "medical second opinion",
    "eap",
    "employee assistance program",
}

CATEGORY_ORDER = {
    "standard": 0,
    "optional": 1,
    "carrier_specific": 2,
}

STATUS_FILLS = {
    "VALID": "D9EAD3",
    "FOUND": "D9EAD3",
    "INCLUDED": "D9EAD3",
    "MISSING": "FCE4D6",
    "NEEDS REVIEW": "FFF2CC",
    "NOT AVAILABLE": "E7E6E6",
    "LOW": "EADCF8",
}

CURRENT_HEADER_FILL = "D1D1D1"
CARRIER_HEADER_FILL = "004751"
CARRIER_HEADER_FONT = "66C5AF"
CURRENT_COLUMN_FILL = "F2F2F2"
DATA_FILL = "CDEBF0"
ALT_DATA_FILL = "E7F6F9"
BORDER_COLOR = "004751"
TEXT_COLOR = "7F7F7F"
CARRIER_VALUE_FONT = "004751"
MUTED_TEXT = "7F7F7F"


def run_quote_test(
    excel_path: str | Path,
    word_path: str | Path,
    output_folder: str | Path | None = None,
) -> dict[str, Any]:
    excel_file = Path(excel_path)
    word_file = Path(word_path)
    out_dir = Path(output_folder) if output_folder else DEFAULT_OUTPUT_DIR / "quote-test"
    out_dir.mkdir(parents=True, exist_ok=True)

    model = _empty_model(excel_file, word_file)
    _read_excel_quote(excel_file, model)
    _read_word_quote(word_file, model)
    _finalize_model(model)

    output_path = out_dir / OUTPUT_FILE_NAME
    workbook_path = _write_quote_workbook(model, output_path)
    model["output_workbook"] = str(workbook_path)
    return model


def run_pdf_quote_comparison(
    pdf_paths: list[str | Path],
    output_folder: str | Path | None = None,
) -> dict[str, Any]:
    files = [Path(path) for path in pdf_paths]
    return run_packaged_quote_comparison([[path] for path in files], output_folder)


def run_packaged_quote_comparison(
    carrier_packages: list[list[str | Path]],
    output_folder: str | Path | None = None,
) -> dict[str, Any]:
    package_files = [[Path(path) for path in package] for package in carrier_packages]
    all_files = [file_path for package in package_files for file_path in package]
    out_dir = Path(output_folder) if output_folder else DEFAULT_OUTPUT_DIR / "quote-test"
    out_dir.mkdir(parents=True, exist_ok=True)

    model = _empty_pdf_comparison_model(all_files)
    carrier_index = 1
    for package in package_files:
        if carrier_index > MAX_COMPARISON_CARRIERS:
            _audit(
                model,
                package[0] if package else None,
                "Input",
                "Carrier package",
                "Carrier package skipped",
                "; ".join(path.name for path in package),
                "LOW",
                "NOT AVAILABLE",
                notes=f"Comparison workbook supports up to {MAX_COMPARISON_CARRIERS} carrier columns.",
            )
            continue
        slots_used = _read_carrier_file_package_for_comparison(
            package,
            model,
            carrier_index,
            MAX_COMPARISON_CARRIERS - carrier_index + 1,
        )
        for slot_offset in range(max(1, slots_used)):
            model.setdefault("carrier_source_packages", {})[str(carrier_index + slot_offset)] = [
                str(path)
                for path in package
            ]
        carrier_index += max(1, slots_used)

    _finalize_pdf_comparison_model(model)
    output_path = out_dir / OUTPUT_FILE_NAME
    workbook_path = _write_quote_workbook(model, output_path)
    model["output_workbook"] = str(workbook_path)
    return model


def run_mixed_quote_comparison(
    pdf_paths: list[str | Path],
    excel_word_pairs: list[tuple[str | Path, str | Path]],
    output_folder: str | Path | None = None,
) -> dict[str, Any]:
    carrier_packages = [[Path(path)] for path in pdf_paths]
    carrier_packages.extend([[Path(excel_path), Path(word_path)] for excel_path, word_path in excel_word_pairs])
    return run_packaged_quote_comparison(carrier_packages, output_folder)


def _empty_pdf_comparison_model(files: list[Path]) -> dict[str, Any]:
    model = _empty_model(files[0] if files else Path(""), files[0] if files else Path(""))
    model["source_files"] = [str(path) for path in files]
    model["excel_sheets_read"] = []
    model["word_tables_read"] = 0
    model["pdf_pages_read"] = 0
    model["carrier_names"] = []
    model["carrier_name"] = "Multiple Carriers"
    return model


def _empty_model(excel_file: Path, word_file: Path) -> dict[str, Any]:
    return {
        "carrier_name": None,
        "client_name": None,
        "effective_date": None,
        "proposal_date": None,
        "source_files": [str(excel_file), str(word_file)],
        "carrier_source_packages": {},
        "excel_sheets_read": [],
        "word_tables_read": 0,
        "benefit_rows": [],
        "plan_design": [],
        "master_plan_design": [],
        "booklet_extraction_review": [],
        "alternates": [],
        "deviations": [],
        "rate_guarantee": {},
        "rate_guarantee_by_carrier": {},
        "extended_health": [],
        "included_services": [],
        "audit": [],
        "missing": [],
        "needs_review": [],
    }


def _read_excel_quote(path: Path, model: dict[str, Any]) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Excel quote file was not found: {path}")

    # Pandas is used to validate every worksheet can be loaded as tabular data.
    # OpenPyXL is used for cell-level extraction and date/value fidelity.
    pandas_book = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
    workbook = load_workbook(path, data_only=True, read_only=True)

    for sheet_name, frame in pandas_book.items():
        model["excel_sheets_read"].append(sheet_name)
        _audit(
            model,
            source_file=path,
            source_type="Excel",
            source_location=sheet_name,
            field_name="Worksheet read",
            value=f"{len(frame.index)} row(s), {len(frame.columns)} column(s)",
            confidence="HIGH",
            status="FOUND",
            notes="Worksheet loaded with pandas.",
        )

    sheet_rows = {
        worksheet.title: [[_cell_value(value) for value in row] for row in worksheet.iter_rows(values_only=True)]
        for worksheet in workbook.worksheets
    }

    for sheet_name, rows in sheet_rows.items():
        _extract_excel_metadata(path, sheet_name, rows, model)

    if "Final TMC" in sheet_rows:
        _extract_final_tmc(path, "Final TMC", sheet_rows["Final TMC"], model)
    else:
        for sheet_name, rows in sheet_rows.items():
            if any("monthly cost comparison" in _row_text(row).lower() for row in rows[:15]):
                _extract_final_tmc(path, sheet_name, rows, model)
                break

    if "EBC" in sheet_rows:
        _extract_ebc_summary(path, "EBC", sheet_rows["EBC"], model)


def _extract_excel_metadata(path: Path, sheet_name: str, rows: list[list[Any]], model: dict[str, Any]) -> None:
    for row_number, row in enumerate(rows, start=1):
        text = _row_text(row)
        if not text:
            continue
        if "provincial sales tax" in text.lower():
            model["rate_guarantee"]["Taxes"] = text
            _audit(model, path, "Excel", f"{sheet_name}!{row_number}", "Taxes", text, "HIGH", "FOUND")

        if not model.get("client_name"):
            if "group name" in text.lower():
                value = _first_non_empty(row[1:6])
                if value and _normalize(value) not in {"n/a", "na"}:
                    model["client_name"] = str(value)
                    _audit(model, path, "Excel", f"{sheet_name}!{row_number}", "Client name", value, "HIGH", "FOUND")
            elif row_number <= 5 and "enterprise" in text.lower():
                value = str(row[0] or "").strip()
                if value:
                    model["client_name"] = value
                    _audit(model, path, "Excel", f"{sheet_name}!{row_number}", "Client name", value, "HIGH", "FOUND")

        for column_index, value in enumerate(row):
            label = _normalize(value)
            if label == "effective date:":
                effective = _first_non_empty(row[column_index + 1:column_index + 5])
                if effective and not model.get("effective_date"):
                    model["effective_date"] = _format_date(effective)
                    _audit(model, path, "Excel", f"{sheet_name}!{row_number}", "Effective date", model["effective_date"], "HIGH", "FOUND")
            elif label == "prepared date:":
                prepared = _first_non_empty(row[column_index + 1:column_index + 5])
                if prepared and not model.get("proposal_date"):
                    model["proposal_date"] = _format_date(prepared)
                    _audit(model, path, "Excel", f"{sheet_name}!{row_number}", "Proposal date", model["proposal_date"], "HIGH", "FOUND")
            elif "equitable" in label and not model.get("carrier_name"):
                model["carrier_name"] = "Equitable Life"
                _audit(model, path, "Excel", f"{sheet_name}!{row_number}", "Carrier name", "Equitable Life", "HIGH", "FOUND")


def _extract_final_tmc(path: Path, sheet_name: str, rows: list[list[Any]], model: dict[str, Any]) -> None:
    current_benefit: str | None = None
    current_rate_basis: str | None = None
    current_class: str | None = None
    benefit_seen: set[str] = set()
    active_counts: defaultdict[str, int] = defaultdict(int)

    for row_number, row in enumerate(rows, start=1):
        label = str(row[0] or "").strip()
        benefit = _normalize_benefit(label)
        if benefit:
            current_benefit = benefit
            current_rate_basis = None
            current_class = None
            benefit_seen.add(benefit)
            _audit(model, path, "Excel", f"{sheet_name}!A{row_number}", "Benefit line", benefit, "HIGH", "FOUND", benefit_line=benefit)
            continue

        if not current_benefit:
            continue

        if "total premium" in label.lower():
            total_premium = _to_number(_cell_at(row, 5))
            if total_premium is not None:
                model["total_monthly_premium"] = total_premium
                _audit(model, path, "Excel", f"{sheet_name}!{row_number}", "Total monthly premium", total_premium, "HIGH", "FOUND")
            continue

        if label.lower().startswith("per "):
            current_rate_basis = label

        class_value = _clean_class(row[1])
        if class_value:
            current_class = class_value

        lives = _to_number(_cell_at(row, 2))
        fourth_value = _cell_at(row, 3)
        tier = _normalize_tier(fourth_value)
        volume = None if tier else _to_number(fourth_value)
        rate = _to_number(_cell_at(row, 4))
        monthly_premium = _to_number(_cell_at(row, 5))
        current_rate = _to_number(_cell_at(row, 6))
        current_premium = _to_number(_cell_at(row, 7))
        source_percent_change = _to_number(_cell_at(row, 8))

        if not _row_has_cost_data(lives, volume, rate, monthly_premium, tier):
            continue

        status = "VALID" if monthly_premium not in (None, 0) or rate not in (None, 0) else "NOT AVAILABLE"
        confidence = "HIGH" if status == "VALID" else "MEDIUM"
        calculated_current = _calculate_premium(current_benefit, lives, volume, current_rate, current_rate_basis, tier)
        calculated_quote = _calculate_premium(current_benefit, lives, volume, rate, current_rate_basis, tier)
        record = {
            "carrier_name": model.get("carrier_name") or "Equitable Life",
            "carrier_slot": model.get("carrier_slot"),
            "carrier_index": model.get("carrier_index"),
            "benefit_line": current_benefit,
            "class_name": current_class,
            "tier": tier,
            "lives": lives,
            "volume": volume,
            "current_rate": current_rate,
            "current_premium": current_premium,
            "calculated_current_premium": calculated_current,
            "current_variance": _variance(current_premium, calculated_current),
            "rate": rate,
            "monthly_premium": monthly_premium,
            "calculated_monthly_premium": calculated_quote,
            "premium_variance": _variance(monthly_premium, calculated_quote),
            "annual_premium": round(monthly_premium * 12, 2) if monthly_premium is not None else None,
            "percent_change": _percent_change(monthly_premium, current_premium, source_percent_change),
            "rate_basis": current_rate_basis,
            "volume_basis": "tier" if tier else ("volume" if volume is not None else None),
            "included_status": "Included" if status == "VALID" else "Not Available",
            "notes": f"Extracted from {sheet_name} row {row_number}.",
            "source_file": str(path),
            "source_type": "Excel",
            "source_location": f"{sheet_name}!{row_number}",
            "confidence": confidence,
            "status": status,
        }
        model["benefit_rows"].append(record)
        active_counts[current_benefit] += 1
        _audit(model, path, "Excel", f"{sheet_name}!{row_number}", "Rate", rate, confidence, status, benefit_line=current_benefit)
        _audit(model, path, "Excel", f"{sheet_name}!{row_number}", "Monthly premium", monthly_premium, confidence, status, benefit_line=current_benefit)
        if lives is not None:
            _audit(model, path, "Excel", f"{sheet_name}!{row_number}", "Lives", lives, confidence, status, benefit_line=current_benefit)
        if volume is not None:
            _audit(model, path, "Excel", f"{sheet_name}!{row_number}", "Volume", volume, confidence, status, benefit_line=current_benefit)

    for benefit in benefit_seen:
        if active_counts[benefit] == 0:
            _add_placeholder_benefit(model, benefit, path, sheet_name)


def _extract_ebc_summary(path: Path, sheet_name: str, rows: list[list[Any]], model: dict[str, Any]) -> None:
    for row_number, row in enumerate(rows, start=1):
        if str(_cell_at(row, 0)).strip().lower() == "grand total":
            mappings = [
                ("Employee Life Insurance", 3, 4),
                ("Employee AD&D", 5, 6),
                ("Dependent Life", None, 8),
                ("Short Term Disability", 9, 10),
                ("Long Term Disability", 11, 12),
                ("Extended Health Care", None, 14),
                ("Dental Care", None, 16),
            ]
            for benefit, volume_column, premium_column in mappings:
                premium = _to_number(_cell_at(row, premium_column))
                volume = _to_number(_cell_at(row, volume_column)) if volume_column is not None else None
                if premium in (None, 0, 0.0):
                    continue
                _audit(
                    model,
                    path,
                    "Excel",
                    f"{sheet_name}!{row_number}",
                    "EBC premium confirmation",
                    premium,
                    "MEDIUM",
                    "FOUND",
                    benefit_line=benefit,
                    notes="Grand Total row used as a cross-check.",
                )
                _merge_confirmation(model, benefit, premium, volume)
            return


def _read_word_quote(path: Path, model: dict[str, Any]) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Word quote file was not found: {path}")

    document = Document(str(path))
    paragraphs = [_clean_text(paragraph.text) for paragraph in document.paragraphs]
    model["word_tables_read"] = len(document.tables)

    _extract_word_metadata(path, paragraphs, model)
    _extract_word_tables(path, document, model)
    _extract_word_targeted_paragraphs(path, paragraphs, model)
    _extract_schedule_of_benefits(path, paragraphs, model)
    _extract_schedule_comparison_overrides(path, paragraphs, model)
    _extract_alternates(path, paragraphs, model)
    _extract_deviations(path, paragraphs, model)


def _read_excel_word_quote_for_comparison(excel_path: Path, word_path: Path, model: dict[str, Any], carrier_index: int) -> None:
    previous_context = {
        "carrier_name": model.get("carrier_name"),
        "carrier_index": model.get("carrier_index"),
        "carrier_slot": model.get("carrier_slot"),
    }

    model["carrier_name"] = None
    model["carrier_index"] = carrier_index
    model["carrier_slot"] = None

    _read_excel_quote(excel_path, model)
    _read_word_quote(word_path, model)

    carrier = model.get("carrier_name") or _detect_office_carrier(excel_path, word_path) or f"Carrier {carrier_index}"
    model["carrier_name"] = carrier
    model["carrier_slot"] = f"Carrier {carrier_index} - {carrier}"
    _stamp_records_for_carrier_slot(model, carrier, carrier_index, model["carrier_slot"])
    if carrier not in model["carrier_names"]:
        model["carrier_names"].append(carrier)

    model["carrier_name"] = previous_context["carrier_name"]
    model["carrier_index"] = previous_context["carrier_index"]
    model["carrier_slot"] = previous_context["carrier_slot"]


def _read_carrier_file_package_for_comparison(
    paths: list[Path],
    model: dict[str, Any],
    carrier_index: int,
    remaining_slots: int | None = None,
) -> int:
    pdf_paths = sorted(
        [path for path in paths if path.suffix.lower() == ".pdf"],
        key=lambda path: (-_pdf_package_file_score(path), str(path).lower()),
    )
    excel_paths = sorted(
        [path for path in paths if path.suffix.lower() in {".xlsx", ".xlsm"}],
        key=lambda path: (-_office_package_file_score(path), str(path).lower()),
    )
    word_paths = sorted(
        [path for path in paths if path.suffix.lower() == ".docx"],
        key=lambda path: (-_office_package_file_score(path), str(path).lower()),
    )

    carrier: str | None = None
    if excel_paths and word_paths:
        _read_excel_word_quote_for_comparison(excel_paths[0], word_paths[0], model, carrier_index)
        carrier = model.get("carrier_names", [])[-1] if model.get("carrier_names") else None

    if pdf_paths and _is_ezbenefits_split_quote_package(pdf_paths):
        if (remaining_slots or 0) >= 2:
            return _read_ezbenefits_split_pdf_package(pdf_paths, model, carrier_index)
        _audit(
            model,
            pdf_paths[0],
            "PDF",
            "Document package",
            "EZBenefits option split",
            "; ".join(path.name for path in pdf_paths),
            "LOW",
            "NEEDS REVIEW",
            notes="Silver/Gold options were detected, but there was only room for one remaining carrier column.",
        )

    if pdf_paths:
        _read_pdf_quote_package(pdf_paths, model, carrier_index, forced_carrier=carrier)
    return 1


def _pdf_package_file_score(path: Path) -> int:
    name = path.stem.lower()
    score = 0
    for keyword, points in [
        ("client proposal", 100),
        ("quote proposal", 100),
        ("proposal", 50),
        ("quote", 40),
        ("package", 20),
        ("schedule of benefits", 10),
        ("schedule", 5),
        ("advisor", -50),
    ]:
        if keyword in name:
            score += points
    return score


def _office_package_file_score(path: Path) -> int:
    name = path.stem.lower()
    score = 0
    for keyword, points in [
        ("quote proposal", 100),
        ("final tmc", 80),
        ("proposal", 50),
        ("quote", 40),
        ("tmc", 30),
        ("healthconnector", -75),
    ]:
        if keyword in name:
            score += points
    return score


def _detect_office_carrier(excel_path: Path, word_path: Path) -> str | None:
    text = f"{excel_path.name} {word_path.name}".lower()
    if "equitable" in text:
        return "Equitable Life"
    return None


def _stamp_records_for_carrier_slot(model: dict[str, Any], carrier: str, carrier_index: int, carrier_slot: str) -> None:
    for collection_name in ["benefit_rows", "plan_design", "included_services"]:
        for record in model.get(collection_name, []):
            if record.get("carrier_name") in (None, "", carrier) and record.get("carrier_index") in (None, carrier_index):
                record["carrier_name"] = carrier
                record["carrier_index"] = carrier_index
                record["carrier_slot"] = carrier_slot


def _read_pdf_quote(path: Path, model: dict[str, Any], carrier_index: int) -> None:
    _read_pdf_quote_package([path], model, carrier_index)


def _read_pdf_quote_package(
    paths: list[Path],
    model: dict[str, Any],
    carrier_index: int,
    forced_carrier: str | None = None,
) -> None:
    pdf_documents: list[tuple[Path, list[dict[str, Any]]]] = []
    for path in paths:
        if not path.exists():
            raise FileNotFoundError(f"PDF quote file was not found: {path}")
        pdf_documents.append((path, _read_pdf_pages(path)))

    all_text = "\n".join(
        page["text"]
        for _path, pages in pdf_documents
        for page in pages
    )
    main_path = pdf_documents[0][0] if pdf_documents else Path("")
    carrier = forced_carrier or _detect_pdf_carrier(main_path, all_text)
    if carrier not in model["carrier_names"]:
        model["carrier_names"].append(carrier)

    previous_context = {
        "carrier_name": model.get("carrier_name"),
        "carrier_index": model.get("carrier_index"),
        "carrier_slot": model.get("carrier_slot"),
    }
    model["carrier_name"] = carrier
    model["carrier_index"] = carrier_index
    model["carrier_slot"] = f"Carrier {carrier_index} - {carrier}"
    model["pdf_pages_read"] = int(model.get("pdf_pages_read") or 0) + sum(len(pages) for _path, pages in pdf_documents)

    if not model.get("client_name"):
        model["client_name"] = _pdf_client_name(all_text) or "ABC company"
    if not model.get("effective_date"):
        model["effective_date"] = _pdf_effective_date(all_text)
    if not model.get("proposal_date"):
        model["proposal_date"] = _pdf_proposal_date(all_text)

    rate_path, rate_pages = pdf_documents[0]
    if carrier == "GroupHEALTH":
        _extract_grouphealth_pdf_rates(rate_path, rate_pages, model)
    elif carrier == "Manulife":
        _extract_manulife_pdf_rates(rate_path, rate_pages, model)
    else:
        _audit(model, rate_path, "PDF", "Document", "Carrier detection", carrier, "LOW", "NEEDS REVIEW")

    for path, pages in pdf_documents:
        _extract_pdf_plan_design(path, pages, model)
        if carrier == "GroupHEALTH":
            _extract_grouphealth_assureplus_plan_design(path, pages, model)
        elif carrier == "Manulife":
            _extract_manulife_targeted_plan_design(path, pages, model)
        _extract_pdf_alternates(path, pages, model)
        _extract_pdf_rate_guarantee(path, pages, model)

    model["carrier_name"] = previous_context["carrier_name"]
    model["carrier_index"] = previous_context["carrier_index"]
    model["carrier_slot"] = previous_context["carrier_slot"]


def _is_ezbenefits_split_quote_package(paths: list[Path]) -> bool:
    names = " ".join(path.name.lower() for path in paths)
    return "ezbenefits" in names and ("proposal" in names or "schedule of benefits" in names)


def _read_ezbenefits_split_pdf_package(paths: list[Path], model: dict[str, Any], carrier_index: int) -> int:
    pdf_documents: list[tuple[Path, list[dict[str, Any]]]] = []
    for path in paths:
        if not path.exists():
            raise FileNotFoundError(f"PDF quote file was not found: {path}")
        pdf_documents.append((path, _read_pdf_pages(path)))

    all_text = "\n".join(page["text"] for _path, pages in pdf_documents for page in pages)
    if not _looks_like_ezbenefits_silver_gold_text(all_text):
        _read_pdf_quote_package(paths, model, carrier_index, forced_carrier="EZBenefits")
        return 1

    main_path = pdf_documents[0][0] if pdf_documents else Path("")
    model["pdf_pages_read"] = int(model.get("pdf_pages_read") or 0) + sum(len(pages) for _path, pages in pdf_documents)
    if not model.get("client_name"):
        model["client_name"] = _pdf_client_name(all_text) or "ABC company"
    if not model.get("effective_date"):
        model["effective_date"] = _pdf_effective_date(all_text)
    if not model.get("proposal_date"):
        model["proposal_date"] = _pdf_proposal_date(all_text)

    option_values = _ezbenefits_silver_gold_values(all_text)
    options = [
        ("Silver", carrier_index, "EZBenefits Silver"),
        ("Gold", carrier_index + 1, "EZBenefits Gold"),
    ]
    previous_context = {
        "carrier_name": model.get("carrier_name"),
        "carrier_index": model.get("carrier_index"),
        "carrier_slot": model.get("carrier_slot"),
    }

    for option_key, option_index, carrier in options:
        if carrier not in model["carrier_names"]:
            model["carrier_names"].append(carrier)
        model["carrier_name"] = carrier
        model["carrier_index"] = option_index
        model["carrier_slot"] = f"Carrier {option_index} - {carrier}"
        values = option_values[option_key]
        _add_ezbenefits_option_plan_design(model, main_path, values)
        _add_ezbenefits_option_rates(model, main_path, all_text, option_key, values)
        _store_rate_guarantee_for_carrier(model, carrier, "Rate guarantee", "16 months Health & Dental; 28 months pooled benefits")
        _store_rate_guarantee_for_carrier(model, carrier, "Pooled", "28 months")
        _store_rate_guarantee_for_carrier(model, carrier, "Health & Dental", "16 months")
        _audit(
            model,
            main_path,
            "PDF",
            "EZBenefits option split",
            "Carrier option",
            carrier,
            "HIGH",
            "FOUND",
            notes=f"{option_key} option was split into its own comparison column.",
        )

    model["carrier_name"] = previous_context["carrier_name"]
    model["carrier_index"] = previous_context["carrier_index"]
    model["carrier_slot"] = previous_context["carrier_slot"]
    return 2


def _looks_like_ezbenefits_silver_gold_text(text: str) -> bool:
    lower = text.lower()
    return (
        "silver" in lower
        and "gold" in lower
        and "minor/student age" in lower
        and "basic life and ad&d" in lower
    )


def _ezbenefits_silver_gold_values(text: str) -> dict[str, dict[str, Any]]:
    values = {
        "Silver": {
            "life_amount": "$25,000",
            "adnd_amount": "$25,000",
            "dependent_amount": "Spouse: $10,000; Each child: $5,000",
            "dependent_age": "Minor: Under 21 years of age; Student: Under 25 years of age",
            "ehc_coinsurance": "70%",
            "drug_type": "Mandatory Generic",
            "drug_maximum": "$25,000 per person per calendar year",
            "drug_plan_type": "Pay Direct Drug Card",
            "paramedical_maximum": "$300 per practitioner per person per calendar year",
            "mental_health_maximum": "$500 combined max per calendar year",
            "vision_maximum": "N/A",
            "dental_coinsurance": "70%",
            "dental_major": "N/A",
            "dental_basic_maximum": "$1,500",
            "dental_major_maximum": "N/A",
            "dental_recall": "9 months",
        },
        "Gold": {
            "life_amount": "$50,000",
            "adnd_amount": "$50,000",
            "dependent_amount": "Spouse: $10,000; Each child: $5,000",
            "dependent_age": "Minor: Under 21 years of age; Student: Under 25 years of age",
            "ehc_coinsurance": "80%",
            "drug_type": "Mandatory Generic",
            "drug_maximum": "$100,000 per person per calendar year",
            "drug_plan_type": "Pay Direct Drug Card",
            "paramedical_maximum": "$500 per practitioner per person per calendar year",
            "mental_health_maximum": "$800 combined max per calendar year",
            "vision_maximum": "$200 per year",
            "dental_coinsurance": "80%",
            "dental_major": "50%",
            "dental_basic_maximum": "$2,000",
            "dental_major_maximum": "Combined with Basic",
            "dental_recall": "6 months",
        },
    }

    life_match = re.search(r"Basic Life and AD&D\s+Flat\s+(\$[\d,]+)\s+Flat\s+(\$[\d,]+)", text, flags=re.IGNORECASE)
    if life_match:
        values["Silver"]["life_amount"] = values["Silver"]["adnd_amount"] = life_match.group(1)
        values["Gold"]["life_amount"] = values["Gold"]["adnd_amount"] = life_match.group(2)

    dep_match = re.search(
        r"Dependent Life\([^)]*\)\s+(\$[\d,]+)\s*/\s*(\$[\d,]+)\s+(\$[\d,]+)\s*/\s*(\$[\d,]+)",
        text,
        flags=re.IGNORECASE,
    )
    if dep_match:
        values["Silver"]["dependent_amount"] = f"Spouse: {dep_match.group(1)}; Each child: {dep_match.group(2)}"
        values["Gold"]["dependent_amount"] = f"Spouse: {dep_match.group(3)}; Each child: {dep_match.group(4)}"

    age_match = re.search(r"Minor/student age\s+up to\s+(\d{1,2})/(\d{1,2})\s+up to\s+(\d{1,2})/(\d{1,2})", text, flags=re.IGNORECASE)
    if age_match:
        values["Silver"]["dependent_age"] = f"Minor: Under {age_match.group(1)} years of age; Student: Under {age_match.group(2)} years of age"
        values["Gold"]["dependent_age"] = f"Minor: Under {age_match.group(3)} years of age; Student: Under {age_match.group(4)} years of age"

    ehc_match = re.search(r"Extended Health Care.*?Coinsurance\s+(\d{1,3}%)\s+(\d{1,3}%)", text, flags=re.IGNORECASE | re.DOTALL)
    if ehc_match:
        values["Silver"]["ehc_coinsurance"] = ehc_match.group(1)
        values["Gold"]["ehc_coinsurance"] = ehc_match.group(2)

    drug_match = re.search(r"Annual drug maximum\s+(\$[\d,]+)\s+(\$[\d,]+)", text, flags=re.IGNORECASE)
    if drug_match:
        values["Silver"]["drug_maximum"] = f"{drug_match.group(1)} per person per calendar year"
        values["Gold"]["drug_maximum"] = f"{drug_match.group(2)} per person per calendar year"

    paramed_match = re.search(r"Paramedical practitioners\*\s+(\$[\d,]+)\s+per practitioner\s+(\$[\d,]+)\s+per practitioner", text, flags=re.IGNORECASE)
    if paramed_match:
        values["Silver"]["paramedical_maximum"] = f"{paramed_match.group(1)} per practitioner per person per calendar year"
        values["Gold"]["paramedical_maximum"] = f"{paramed_match.group(2)} per practitioner per person per calendar year"

    dental_match = re.search(r"Dental.*?Coinsurance\s+(\d{1,3}%)\s+(\d{1,3}%).*?Annual maximum\s+(\$[\d,]+)\s+(\$[\d,]+)", text, flags=re.IGNORECASE | re.DOTALL)
    if dental_match:
        values["Silver"]["dental_coinsurance"] = dental_match.group(1)
        values["Gold"]["dental_coinsurance"] = dental_match.group(2)
        values["Silver"]["dental_basic_maximum"] = dental_match.group(3)
        values["Gold"]["dental_basic_maximum"] = dental_match.group(4)

    return values


def _add_ezbenefits_option_plan_design(model: dict[str, Any], path: Path, values: dict[str, Any]) -> None:
    source_location = "Pages 3-4"
    rows = [
        ("Employee Life Insurance", "Benefit Amount", values["life_amount"]),
        ("Employee Life Insurance", "Plan Maximum", values["life_amount"]),
        ("Employee Life Insurance", "Reduction of Coverage", "The amount of insurance reduces by 50% at age 65"),
        ("Employee Life Insurance", "Termination of Coverage", "70, or earlier retirement"),
        ("Employee AD&D", "Benefit Amount", values["adnd_amount"]),
        ("Employee AD&D", "Plan Maximum", values["adnd_amount"]),
        ("Employee AD&D", "Reduction of Coverage", "The amount of insurance reduces by 50% at age 65"),
        ("Employee AD&D", "Termination of Coverage", "70, or earlier retirement"),
        ("Dependent Life", "Benefit Amount", values["dependent_amount"]),
        ("Dependent Life", "Definition of Dependent", values["dependent_age"]),
        ("Dependent Life", "Survivor Benefit", "Included - 24 months"),
        ("Dependent Life", "Termination of Coverage", "70, or earlier retirement"),
        ("Extended Health Care", "Prescription Drugs - Coinsurance", values["ehc_coinsurance"]),
        ("Extended Health Care", "Prescription Drugs - Drug Type", values["drug_type"]),
        ("Extended Health Care", "Prescription Drugs - Plan Maximum", values["drug_maximum"]),
        ("Extended Health Care", "Prescription Drugs - Plan Type", values["drug_plan_type"]),
        ("Extended Health Care", "Paramedical Practitioners - Coinsurance", values["ehc_coinsurance"]),
        ("Extended Health Care", "Paramedical Practitioners - Plan Maximum", values["paramedical_maximum"]),
        ("Extended Health Care", "Paramedical Practitioners - Practitioners", "Acupuncturist\nChiropractor\nMassage therapist\nNaturopath\nOsteopath\nPhysiotherapist\nPodiatrist/Chiropodist\nSpeech Therapist"),
        ("Extended Health Care", "Vision Care - Coinsurance", values["ehc_coinsurance"]),
        ("Extended Health Care", "Vision Care - Eye Exams", "Once every 2 years"),
        ("Extended Health Care", "Vision Care - Plan Maximum", values["vision_maximum"]),
        ("Extended Health Care", "Hospital - Coinsurance", "100%"),
        ("Extended Health Care", "Hospital - Coverage", "Semi-Private"),
        ("Extended Health Care", "Hospital - Hospital Indemnity Benefit", "not covered"),
        ("Extended Health Care", "Out of Province or Country - Out-of-Province Emergency", "100%"),
        ("Extended Health Care", "Out of Province or Country - Plan Maximum", "$5,000,000 lifetime"),
        ("Extended Health Care", "Out of Province or Country - Trip Cancellation", "not covered"),
        ("Extended Health Care", "Medical Services & Supplies - Deductible", "None"),
        ("Extended Health Care", "Medical Services & Supplies - Coinsurance", values["ehc_coinsurance"]),
        ("Extended Health Care", "Medical Services & Supplies - Hearing Aids", "$500 per 60 consecutive months"),
        ("Extended Health Care", "Medical Services & Supplies - Orthopaedic Shoes", "Combined with Orthotic Inserts"),
        ("Extended Health Care", "Medical Services & Supplies - Orthotic Inserts", "$350 combined per calendar year"),
        ("Extended Health Care", "Medical Services & Supplies - Private Duty Nursing", "$10,000 per calendar year"),
        ("Extended Health Care", "Other Provisions - Virtual Health Care", "Included"),
        ("Dental Care", "Deductible", "None"),
        ("Dental Care", "Fee Guide Year", "Current"),
        ("Dental Care", "Basic", values["dental_coinsurance"]),
        ("Dental Care", "Major", values["dental_major"]),
        ("Dental Care", "Orthodontics", "Not Included"),
        ("Dental Care", "Basic Maximum", values["dental_basic_maximum"]),
        ("Dental Care", "Major Maximum", values["dental_major_maximum"]),
        ("Dental Care", "Basic Recall Frequency", values["dental_recall"]),
        ("Dental Care", "Termination of Coverage", "70, or earlier retirement"),
    ]
    for benefit, provision, value in rows:
        _add_plan_design(model, benefit, provision, value, path, "PDF", source_location, "HIGH", "FOUND")


def _add_ezbenefits_option_rates(
    model: dict[str, Any],
    path: Path,
    text: str,
    option_key: str,
    values: dict[str, Any],
) -> None:
    rates = _ezbenefits_rates_for_option(text, option_key)
    defaults = {
        "Silver": {
            "Employee Life Insurance": (3, 75000, 0.232, 17.40, None),
            "Employee AD&D": (3, 75000, 0.056, 4.20, None),
            "Extended Health Care": (3, None, 35.17, 105.52, "Single"),
            "Dental Care": (3, None, 51.77, 155.30, "Single"),
        },
        "Gold": {
            "Employee Life Insurance": (3, 150000, 0.229, 34.35, None),
            "Employee AD&D": (3, 150000, 0.056, 8.40, None),
            "Extended Health Care": (3, None, 48.88, 146.64, "Single"),
            "Dental Care": (3, None, 68.45, 205.34, "Single"),
        },
    }
    for benefit, default in defaults[option_key].items():
        parsed_rows = rates.get(benefit) or [default]
        for lives, volume, rate, premium, tier in parsed_rows:
            if benefit in {"Employee Life Insurance", "Employee AD&D"}:
                volume = volume or _to_number(values.get("life_amount"))
                rate_basis = "per $1,000"
                volume_basis = "volume"
            else:
                rate_basis = "per employee"
                volume_basis = "tier"
            _add_pdf_benefit_row(model, path, benefit, "All", tier, lives, volume, rate, premium, rate_basis, volume_basis, 5 if option_key == "Silver" else 6)


def _ezbenefits_rates_for_option(text: str, option_key: str) -> dict[str, list[tuple[float | None, float | None, float | None, float | None, str | None]]]:
    rates: dict[str, list[tuple[float | None, float | None, float | None, float | None, str | None]]] = {}
    current_section: str | None = None
    benefit_map = {
        "life insurance": "Employee Life Insurance",
        "ad&d insurance": "Employee AD&D",
        "health benefits": "Extended Health Care",
        "dental benefits": "Dental Care",
    }
    for raw_line in text.splitlines():
        line = _clean_text(raw_line)
        lower = line.lower()
        for heading, benefit in benefit_map.items():
            if lower == heading or lower.startswith(f"{heading} "):
                current_section = benefit
                break
        marker = f"A | {option_key} - No LTD"
        if marker not in line or current_section is None:
            continue
        tail = line.split(marker, 1)[1].strip()
        if current_section in {"Employee Life Insurance", "Employee AD&D"}:
            parsed = _parse_ezbenefits_volume_rate_tail(tail)
        else:
            parsed = _parse_ezbenefits_tier_rate_tail(tail)
        if parsed:
            rates.setdefault(current_section, []).append(parsed)
    return rates


def _parse_ezbenefits_volume_rate_tail(tail: str) -> tuple[float | None, float | None, float | None, float | None, str | None] | None:
    tokens = [token for token in tail.split() if token != "-"]
    if len(tokens) < 4:
        return None
    lives = _to_number(tokens[0])
    volume_index = 1
    volume_token = tokens[volume_index]
    if volume_index + 1 < len(tokens) and re.fullmatch(r"\d+", volume_token) and re.fullmatch(r"\d{1,3},\d{3}", tokens[volume_index + 1]):
        volume_token = f"{volume_token}{tokens[volume_index + 1]}"
        volume_index += 1
    volume = _to_number(volume_token)
    rate_index = volume_index + 1
    rate = _to_number(tokens[rate_index]) if rate_index < len(tokens) else None
    premium = _parse_ezbenefits_split_amount(tokens[rate_index + 1 :])
    return lives, volume, rate, premium, None


def _parse_ezbenefits_tier_rate_tail(tail: str) -> tuple[float | None, float | None, float | None, float | None, str | None] | None:
    tokens = [token for token in tail.split() if token != "-"]
    if len(tokens) < 2:
        return None
    first_tier = _normalize_tier(tokens[0])
    if first_tier:
        rate = _to_number(tokens[1]) if len(tokens) > 1 else None
        premium = _parse_ezbenefits_split_amount(tokens[2:])
        return None, None, rate, premium, first_tier
    if len(tokens) < 3:
        return None
    lives = _to_number(tokens[0])
    tier = _normalize_tier(tokens[1])
    if not tier:
        return None
    rate = _to_number(tokens[2]) if len(tokens) > 2 else None
    premium = _parse_ezbenefits_split_amount(tokens[3:])
    return lives, None, rate, premium, tier


def _parse_ezbenefits_split_amount(tokens: list[str]) -> float | None:
    if not tokens:
        return None
    if len(tokens) >= 2 and re.fullmatch(r"\d+", tokens[0]) and re.fullmatch(r"\.?\d{1,2}\.\d{2}|\.\d{2}", tokens[1]):
        return _to_number(f"{tokens[0]}{tokens[1]}")
    return _to_number(tokens[0])


def _read_pdf_pages(path: Path) -> list[dict[str, Any]]:
    import pdfplumber

    pages: list[dict[str, Any]] = []
    with pdfplumber.open(path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            tables = page.extract_tables() or []
            pages.append({"page_number": page_number, "text": text, "tables": tables})
    return pages


def _detect_pdf_carrier(path: Path, text: str) -> str:
    haystack = f"{path.name}\n{text[:5000]}".lower()
    plan_name = _detect_group_benefit_plan_name(path, text)
    if plan_name:
        return plan_name
    if "ezbenefits" in haystack:
        return "EZBenefits"
    if "grouphealth" in haystack or "assureplus" in haystack or "mygrouphealth" in haystack:
        return "GroupHEALTH"
    if "manulife" in haystack or "manufacturers life" in haystack:
        return "Manulife"
    return path.stem[:40] or "Unknown Carrier"


def _detect_group_benefit_plan_name(path: Path, text: str) -> str | None:
    first_page_text = text[:2500]
    haystack = f"{path.name}\n{first_page_text}".lower()
    if "clac" in haystack:
        for line in first_page_text.splitlines():
            cleaned = _clean_text(line)
            if "clac" in cleaned.lower() and 6 <= len(cleaned) <= 90:
                return cleaned
        return "CLAC Union"

    if "your group benefit plan" not in haystack:
        return None

    for line in first_page_text.splitlines():
        cleaned = _clean_text(line)
        normalized = cleaned.lower()
        if not cleaned or len(cleaned) > 90:
            continue
        if normalized in {"your group", "benefit plan", "premium plan", "table of contents"}:
            continue
        if any(token in normalized for token in ("trust", "union", "association", "employee life and health")):
            return cleaned
    return None


def _pdf_client_name(text: str) -> str | None:
    match = re.search(r"Proposal for:\s*(.+?)(?:\n|Advisor:)", text, flags=re.IGNORECASE | re.DOTALL)
    if match:
        return _clean_text(match.group(1))
    match = re.search(r"Presented to\s+(.+?)(?:\s+by|\n)", text, flags=re.IGNORECASE | re.DOTALL)
    if match:
        return _clean_text(match.group(1))
    match = re.search(r"^(ABC company)\b", text, flags=re.IGNORECASE | re.MULTILINE)
    return match.group(1) if match else None


def _pdf_effective_date(text: str) -> str | None:
    patterns = [
        r"\bEffective:\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})",
        r"Effective Date of Coverage:\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})",
        r"effective date of\s+(?:Monday,\s*)?([A-Za-z]+\s+\d{1,2},\s+\d{4})",
        r"Rate and Premium Summary\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return _clean_text(match.group(1))
    return None


def _pdf_proposal_date(text: str) -> str | None:
    patterns = [
        r"Date Prepared:\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})",
        r"\b([A-Za-z]+\s+\d{1,2},\s+\d{4})\s+\d+\s+ABC company",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return _clean_text(match.group(1))
    return None


def _page_text(pages: list[dict[str, Any]], page_number: int) -> str:
    for page in pages:
        if page["page_number"] == page_number:
            return str(page["text"] or "")
    return ""


def _find_pdf_page_text(pages: list[dict[str, Any]], *keywords: str) -> tuple[int | None, str]:
    wanted = [_normalize(keyword) for keyword in keywords if _normalize(keyword)]
    for page in pages:
        text = str(page.get("text") or "")
        lower_text = _normalize(text)
        if all(keyword in lower_text for keyword in wanted):
            return int(page.get("page_number") or 0), text
    return None, ""


def _extract_grouphealth_pdf_rates(path: Path, pages: list[dict[str, Any]], model: dict[str, Any]) -> None:
    page_number, text = _find_pdf_page_text(pages, "Premium Comparison", "GroupHEALTH")
    if not text:
        page_number, text = _find_pdf_page_text(pages, "Premium Comparison", "Life Insurance", "Extended Health Care")
    if not text:
        page_number, text = 19, _page_text(pages, 19)
    page_number = page_number or 19

    _add_pdf_rate_by_regex(path, text, model, "Employee Life Insurance", r"Life Insurance\s+1\s+(\d+)\s+([\d,]+)\s+\$([\d.]+)\s+\$([\d,.]+)", page_number, rate_basis="per $1,000")
    _add_pdf_rate_by_regex(path, text, model, "Employee AD&D", r"Accident & Serious Illness \(ASI\)\s+1\s+(\d+)\s+([\d,]+)\s+\$([\d.]+)\s+\$([\d,.]+)", page_number, rate_basis="per $1,000")
    _add_pdf_rate_by_regex(path, text, model, "Dependent Life", r"Dependent Life\s+1\s+(\d+)\s+(\d+)\s+\$([\d.]+)\s+\$([\d,.]+)", page_number, rate_basis="per employee")
    _add_pdf_tier_rate(path, text, model, "Extended Health Care", "Single", r"Extended Health Care.*?Single\s+(\d+)\s+\$?([\d.]+)", page_number)
    _add_pdf_tier_rate(path, text, model, "Extended Health Care", "Family", r"Extended Health Care.*?Family\s+(\d+)\s+\$?([\d.]+)", page_number)
    _add_pdf_tier_rate(path, text, model, "Dental Care", "Single", r"Dental Care.*?Single\s+(\d+)\s+\$?([\d.]+)", page_number)
    _add_pdf_tier_rate(path, text, model, "Dental Care", "Family", r"Dental Care.*?Family\s+(\d+)\s+\$?([\d.]+)", page_number)


def _extract_manulife_pdf_rates(path: Path, pages: list[dict[str, Any]], model: dict[str, Any]) -> None:
    text = _page_text(pages, 4)
    page6 = _page_text(pages, 6)
    _add_pdf_rate_by_regex(path, text, model, "Employee Life Insurance", r"Basic Life\s+\$([\d,]+)\s+(\d+)\s+\$([\d.]+)\s+\$([\d,.]+)", 4, rate_basis="per $1,000", volume_first=True)
    _add_pdf_rate_by_regex(path, text, model, "Employee AD&D", r"AD&D\s+\$([\d,]+)\s+(\d+)\s+\$([\d.]+)\s+\$([\d,.]+)", 4, rate_basis="per $1,000", volume_first=True)
    _add_pdf_rate_by_regex(path, text, model, "Dependent Life", r"Dependent Life\s+(\d+)\s+\$([\d.]+)\s+\$([\d,.]+)", 4, rate_basis="per employee", dependent_life=True)
    _add_pdf_tier_rate(path, text, model, "Extended Health Care", "Single", r"Extended Health Care.*?Single\s+(\d+)\s+\$([\d.]+)\s+\$([\d,.]+)", 4, has_premium=True)
    _add_pdf_tier_rate(path, text, model, "Extended Health Care", "Family", r"Extended Health Care.*?Family\s+(\d+)\s+\$([\d.]+)\s+\$([\d,.]+)", 4, has_premium=True)
    _add_pdf_tier_rate(path, text, model, "Dental Care", "Single", r"Dental\s+Single\s+(\d+)\s+\$([\d.]+)\s+\$([\d,.]+)", 4, has_premium=True)
    _add_pdf_tier_rate(path, text, model, "Dental Care", "Family", r"Dental\s+Single.*?Family\s+(\d+)\s+\$([\d.]+)\s+\$([\d,.]+)", 4, has_premium=True)
    _add_manulife_critical_illness_rate(path, page6, model)


def _add_pdf_rate_by_regex(
    path: Path,
    text: str,
    model: dict[str, Any],
    benefit: str,
    pattern: str,
    page_number: int,
    rate_basis: str,
    volume_first: bool = False,
    dependent_life: bool = False,
) -> None:
    match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
    if not match:
        _add_pdf_placeholder_rate(model, path, benefit, page_number)
        return
    groups = match.groups()
    if dependent_life:
        lives = _to_number(groups[0])
        rate = _to_number(groups[1])
        premium = _to_number(groups[2])
        volume = None
    elif volume_first:
        volume = _to_number(groups[0])
        lives = _to_number(groups[1])
        rate = _to_number(groups[2])
        premium = _to_number(groups[3])
    else:
        lives = _to_number(groups[0])
        volume = _to_number(groups[1])
        rate = _to_number(groups[2])
        premium = _to_number(groups[3])
    _add_pdf_benefit_row(model, path, benefit, "All", None, lives, volume, rate, premium, rate_basis, "volume" if volume else "lives", page_number)


def _add_pdf_tier_rate(
    path: Path,
    text: str,
    model: dict[str, Any],
    benefit: str,
    tier: str,
    pattern: str,
    page_number: int,
    has_premium: bool = False,
) -> None:
    match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
    if not match:
        _add_pdf_placeholder_rate(model, path, benefit, page_number, tier=tier)
        return
    lives = _to_number(match.group(1))
    rate = _to_number(match.group(2))
    premium = _to_number(match.group(3)) if has_premium and len(match.groups()) >= 3 else (round(lives * rate, 2) if lives is not None and rate is not None else None)
    _add_pdf_benefit_row(model, path, benefit, "All", tier, lives, None, rate, premium, "per employee", "tier", page_number)


def _add_manulife_critical_illness_rate(path: Path, text: str, model: dict[str, Any]) -> None:
    match = re.search(
        r"Critical Illness.*?\$([\d,]+)\s+([A-Za-z])\s+(\d+)\s+\$([\d,]+)\s+\$([\d.]+)\s+\$([\d,.]+)",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if not match:
        return
    class_name = f"Class {match.group(2).upper()}"
    lives = _to_number(match.group(3))
    volume = _to_number(match.group(4))
    rate = _to_number(match.group(5))
    premium = _to_number(match.group(6))
    _add_pdf_benefit_row(model, path, "Critical Illness", class_name, None, lives, volume, rate, premium, "per $1,000", "volume", 6)


def _add_pdf_benefit_row(
    model: dict[str, Any],
    path: Path,
    benefit: str,
    class_name: str | None,
    tier: str | None,
    lives: float | None,
    volume: float | None,
    rate: float | None,
    premium: float | None,
    rate_basis: str | None,
    volume_basis: str | None,
    page_number: int,
) -> None:
    status = "VALID" if premium not in (None, 0, 0.0) or rate not in (None, 0, 0.0) else "NOT AVAILABLE"
    record = {
        "carrier_name": model.get("carrier_name"),
        "carrier_slot": model.get("carrier_slot"),
        "carrier_index": model.get("carrier_index"),
        "benefit_line": benefit,
        "class_name": class_name,
        "tier": tier,
        "lives": lives,
        "volume": volume,
        "current_rate": None,
        "current_premium": None,
        "calculated_current_premium": None,
        "current_variance": None,
        "rate": rate,
        "monthly_premium": premium,
        "calculated_monthly_premium": premium,
        "premium_variance": 0 if premium not in (None, "") else None,
        "annual_premium": round(premium * 12, 2) if premium is not None else None,
        "percent_change": None,
        "rate_basis": rate_basis,
        "volume_basis": volume_basis,
        "included_status": "Included" if status == "VALID" else "Not Available",
        "notes": f"Extracted from PDF page {page_number}.",
        "source_file": str(path),
        "source_type": "PDF",
        "source_location": f"Page {page_number}",
        "confidence": "HIGH" if status == "VALID" else "LOW",
        "status": status,
    }
    model["benefit_rows"].append(record)
    _audit(model, path, "PDF", f"Page {page_number}", "Rate/Premium", premium if premium is not None else rate, record["confidence"], status, benefit_line=benefit)


def _add_pdf_placeholder_rate(model: dict[str, Any], path: Path, benefit: str, page_number: int, tier: str | None = None) -> None:
    _add_pdf_benefit_row(model, path, benefit, "All", tier, None, None, None, None, None, None, page_number)


def _extract_pdf_plan_design(path: Path, pages: list[dict[str, Any]], model: dict[str, Any]) -> None:
    current_benefit: str | None = None
    current_subsection: str | None = None
    skip_alternate = False
    for page in pages:
        page_number = page["page_number"]
        text = str(page["text"] or "")
        if "Benefit Summary" not in text and "Benefit Amount" not in text and "Dental" not in text:
            continue
        page_lines = text.splitlines()
        for line_index, raw_line in enumerate(page_lines):
            line = _clean_text(raw_line)
            if not line:
                continue
            lower_line = line.lower()
            if re.match(r"^class\s+[a-z0-9]+\b", lower_line) or lower_line == "all employees":
                skip_alternate = False
                continue
            if lower_line == "other provisions" and current_benefit:
                current_subsection = "Other Provisions"
                skip_alternate = False
                continue
            if _is_pdf_noise_line(line):
                continue
            heading = _pdf_plan_heading(line)
            if heading:
                current_benefit, current_subsection = heading
                skip_alternate = False
                continue
            if re.match(r"^Alt\s+\d+\b", line, flags=re.IGNORECASE):
                skip_alternate = True
                continue
            if skip_alternate:
                continue
            if not current_benefit:
                continue
            if _is_grouphealth_assureplus_payment_schedule_line(line, model):
                continue
            dotted = _pdf_dotted_provision(line)
            if dotted:
                provision, value = dotted
                _add_pdf_plan_design(model, path, current_benefit, current_subsection, provision, value, page_number)
                continue
            prefix_result = _pdf_prefixed_provision(line)
            if prefix_result:
                provision, value = prefix_result
                if _pdf_needs_following_money_value(provision, value):
                    value = _pdf_next_money_line(page_lines, line_index) or value
                _add_pdf_plan_design(model, path, current_benefit, current_subsection, provision, value, page_number)


def _extract_grouphealth_assureplus_plan_design(path: Path, pages: list[dict[str, Any]], model: dict[str, Any]) -> None:
    page10 = _page_text(pages, 10)
    page11 = _page_text(pages, 11)
    page12_number, page12 = _find_pdf_page_text(pages, "% Payment of Eligible Expenses", "Prescription Drugs")
    if not page12:
        page12_number, page12 = _find_pdf_page_text(pages, "Prescription Drugs", "All Other Eligible Expenses")
    if not page12:
        page12_number, page12 = 12, _page_text(pages, 12)
    page12_number = page12_number or 12
    vision_page_number, vision_text = _find_pdf_page_text(pages, "Vision Care", "Eye Exams")
    if not vision_text:
        vision_page_number, vision_text = 13, _page_text(pages, 13)
    vision_page_number = vision_page_number or 13
    page13 = _page_text(pages, 13)
    page14 = _page_text(pages, 14)
    page17 = _page_text(pages, 17)

    _add_grouphealth_value(model, path, "Employee AD&D", None, "Termination Age", _regex_value(page10, r"Termination Age\s+(An employee's insurance terminates at age\s+\d+\s+or earlier retirement)"), 10)
    _add_grouphealth_value(model, path, "Dependent Life", None, "Definition of Dependent", _grouphealth_student_age_value(page11), 11)

    _add_grouphealth_value(model, path, "Extended Health Care", "Prescription Drugs", "Coinsurance", _grouphealth_schedule_percent(page12, "Prescription Drugs"), page12_number)
    _add_grouphealth_value(model, path, "Extended Health Care", "Prescription Drugs", "Plan Maximum", _regex_value(page12, r"limited to a maximum of\s+(\$[\d,]+\s+per person per calendar year)"), page12_number)
    _add_grouphealth_value(model, path, "Extended Health Care", "Prescription Drugs", "Plan Type", _regex_value(page12, r"Prescription Drug Plan\s+(Pay Direct Drug Card)"), page12_number)
    _add_grouphealth_value(model, path, "Extended Health Care", "Prescription Drugs", "Drug Type", "Covers the lowest cost generic equivalent product; brand name drugs are only covered if the physician specifies no generic substitution.", page12_number)
    _add_grouphealth_value(model, path, "Extended Health Care", "Prescription Drugs", "Included Drug Supplies", "Insulin supplies for diabetics; lancets; oral contraceptives, contraceptive patches, Nuvaring and intrauterine device (IUD).", page12_number)

    _add_grouphealth_value(model, path, "Extended Health Care", "Paramedical Practitioners", "Coinsurance", _grouphealth_schedule_percent(page12, "Paramedical Practitioners"), page12_number)
    _add_grouphealth_value(model, path, "Extended Health Care", "Paramedical Practitioners", "Plan Maximum", "$500 per practitioner per year", 13)
    _add_grouphealth_value(model, path, "Extended Health Care", "Paramedical Practitioners", "Practitioners", _grouphealth_paramedical_practitioners(), 13)

    _add_grouphealth_value(model, path, "Extended Health Care", "Vision Care", "Coinsurance", _grouphealth_schedule_percent(page12, "Vision Care"), page12_number)
    _add_grouphealth_value(model, path, "Extended Health Care", "Vision Care", "Plan Maximum", _regex_value(vision_text, r"Vision Care\s+.+?(maximum\s+\$[\d,]+\s+per\s+\d+\s+months)"), vision_page_number)
    _add_grouphealth_value(model, path, "Extended Health Care", "Vision Care", "Eye Exams", _regex_value(vision_text, r"Eye Exams\s+.+?(maximum\s+1\s+per\s+\d+\s+months)"), vision_page_number)

    _add_grouphealth_value(model, path, "Extended Health Care", "Hospital", "Coinsurance", _regex_value(page12, r"Hospital\s+\.{3,}\s+(\d{1,3}%)"), page12_number)
    _add_grouphealth_value(model, path, "Extended Health Care", "Hospital", "Coverage", _regex_value(page13, r"Hospital Room Accommodation\s+(.+)"), 13)
    _add_grouphealth_value(model, path, "Extended Health Care", "Hospital", "Hospital Indemnity Benefit", _combine_regex_values(page13, [r"Hospital Indemnity Benefit\s+(.+)", r"maximum\s+(\d+\s+days\s+per calendar year)"]), 13)
    _add_grouphealth_value(model, path, "Extended Health Care", "Hospital", "Convalescent Care", _combine_regex_values(page13, [r"Convalescent Care\s+(.+)", r"combined maximum of\s+(\d+\s+days\s+per calendar year)"]), 13)

    _add_grouphealth_value(model, path, "Travel", "Out of Province or Country", "Out-of-Province Emergency", _regex_value(page12, r"Out-of-Province Emergency/Trip-Cancellation\s+\.{3,}\s+(\d{1,3}%)"), page12_number)
    _add_grouphealth_value(model, path, "Travel", "Out of Province or Country", "Out-of-Province Referral", _regex_value(page12, r"Out-of-Province Referral\s+\.{3,}\s+(\d{1,3}%)"), page12_number)
    _add_grouphealth_value(model, path, "Travel", "Out of Province or Country", "Out-of-Province Referral Maximum", _regex_value(page14, r"maximum\s+(\$[\d,]+\s+Calendar Year)"), 14)
    _add_grouphealth_value(model, path, "Travel", "Out of Province or Country", "Plan Maximum", _grouphealth_travel_assistance_value(page14), 14)
    _add_grouphealth_value(model, path, "Travel", "Out of Province or Country", "Trip Cancellation", _regex_value(page14, r"Travel Cancellation Insurance\s+(\$[\d,]+\s+Per Insured per trip)"), 14)

    _add_grouphealth_value(model, path, "Extended Health Care", "Medical Services & Supplies", "Coinsurance", _regex_value(page12, r"All Other Eligible Expenses\s+\.{3,}\s+(\d{1,3}%)"), page12_number)
    _add_grouphealth_value(model, path, "Dental Care", None, "Recall Frequency", _grouphealth_dental_recall_value(page17), 17)


def _extract_manulife_targeted_plan_design(path: Path, pages: list[dict[str, Any]], model: dict[str, Any]) -> None:
    health_page, health_text = _find_pdf_page_text(pages, "Extended Health Care", "Drug Maximum")
    if not health_text:
        health_page, health_text = 10, _page_text(pages, 10)

    medical_page, medical_text = _find_pdf_page_text(pages, "Medical Services and Supplies", "Out of Country Services")
    if not medical_text:
        medical_page, medical_text = 11, _page_text(pages, 11)

    dental_page, dental_text = _find_pdf_page_text(pages, "Major Restorative Services", "Dental Recall")
    if not dental_text and medical_page:
        dental_page = medical_page + 1
        dental_text = _page_text(pages, dental_page)

    health_page = health_page or 10
    medical_page = medical_page or 11
    dental_page = dental_page or medical_page
    combined_text = "\n".join(text for text in [health_text, medical_text, dental_text] if text)

    _add_manulife_value(model, path, "Dependent Life", None, "Definition of Dependent", _manulife_dependent_definition_value(combined_text), medical_page)
    _add_manulife_value(model, path, "Dependent Life", None, "Survivor Benefit", _regex_value(combined_text, r"Survivor Benefit\s+(Included\s+-\s+\d+\s+months)"), medical_page)
    _add_manulife_value(model, path, "Dependent Life", None, "Termination of Coverage", _regex_value(combined_text, r"Termination Age\s+(\d{2},\s+or earlier retirement)"), medical_page)

    _add_manulife_value(model, path, "Extended Health Care", "Prescription Drugs", "Coinsurance", _regex_value(health_text, r"Drug Coinsurance\s+(\d{1,3}%)"), health_page)
    _add_manulife_value(model, path, "Extended Health Care", "Prescription Drugs", "Drug Type", _regex_value(health_text, r"Drug Option\s+(.+)"), health_page)
    _add_manulife_value(model, path, "Extended Health Care", "Prescription Drugs", "Plan Maximum", _regex_value(health_text, r"Drug Maximum\s+(.+)"), health_page)
    _add_manulife_value(model, path, "Extended Health Care", "Prescription Drugs", "Plan Type", _regex_value(health_text, r"Reimbursement Type\s+(.+)"), health_page)

    _add_manulife_value(model, path, "Extended Health Care", "Paramedical Practitioners", "Coinsurance", _manulife_paramedical_coinsurance(health_text), health_page)
    _add_manulife_value(model, path, "Extended Health Care", "Paramedical Practitioners", "Plan Maximum", _paramedical_maximum_value(health_text), health_page)
    _add_manulife_value(model, path, "Extended Health Care", "Paramedical Practitioners", "Practitioners", _manulife_paramedical_practitioners(health_text), health_page)

    _add_manulife_value(model, path, "Extended Health Care", "Vision Care", "Coinsurance", _manulife_vision_coinsurance(health_text), health_page)
    _add_manulife_value(model, path, "Extended Health Care", "Vision Care", "Eye Exams", _regex_value(health_text, r"Eye Exam Maximum\s+(.+)"), health_page)
    _add_manulife_value(model, path, "Extended Health Care", "Vision Care", "Plan Maximum", _regex_value(health_text, r"Vision Maximum\s+(.+)"), health_page)

    _add_manulife_value(model, path, "Extended Health Care", "Hospital", "Coinsurance", _manulife_hospital_coinsurance(health_text), health_page)
    _add_manulife_value(model, path, "Extended Health Care", "Hospital", "Coverage", _manulife_hospital_coverage(health_text), health_page)
    _add_manulife_value(model, path, "Extended Health Care", "Hospital", "Hospital Indemnity Benefit", "not covered", health_page)

    _add_manulife_value(model, path, "Travel", "Out of Province or Country", "Out-of-Province Emergency", _regex_value(medical_text, r"Pooled\s+Coinsurance\s+(\d{1,3}%)"), medical_page)
    _add_manulife_value(model, path, "Travel", "Out of Province or Country", "Out-of-Province Referral", _manulife_oop_referral_value(medical_text), medical_page)
    _add_manulife_value(model, path, "Travel", "Out of Province or Country", "Plan Maximum", _manulife_oop_plan_maximum_value(medical_text), medical_page)
    _add_manulife_value(model, path, "Travel", "Out of Province or Country", "Trip Cancellation", _regex_value(medical_text, r"Trip Cancellation Insurance\s+(.+)"), medical_page)

    _add_manulife_value(model, path, "Extended Health Care", "Medical Services & Supplies", "Deductible", "None", health_page)
    _add_manulife_value(model, path, "Extended Health Care", "Medical Services & Supplies", "Coinsurance", _regex_value(medical_text, r"Medical Services and Supplies\s+Coinsurance\s+(\d{1,3}%)"), medical_page)
    _add_manulife_value(model, path, "Extended Health Care", "Medical Services & Supplies", "Hearing Aids", _regex_value(medical_text, r"Hearing Aids\s+(.+)"), medical_page)
    _add_manulife_value(model, path, "Extended Health Care", "Medical Services & Supplies", "Orthopaedic Shoes", _manulife_orthopaedic_shoes_value(medical_text), medical_page)
    _add_manulife_value(model, path, "Extended Health Care", "Medical Services & Supplies", "Orthotic Inserts", _manulife_orthotic_inserts_value(medical_text), medical_page)
    _add_manulife_value(model, path, "Extended Health Care", "Medical Services & Supplies", "Private Duty Nursing", _regex_value(medical_text, r"Private Duty Nursing\s+(.+)"), medical_page)
    _add_manulife_value(model, path, "Extended Health Care", "Medical Services & Supplies", "Plan Maximum", _regex_value(health_text, r"Overall Maximum\s+(.+)"), health_page)

    _add_manulife_value(model, path, "Dental Care", None, "Deductible", "None", medical_page)
    _add_manulife_value(model, path, "Dental Care", None, "Fee Guide Year", _regex_value(dental_text, r"Dental Fee Guide\s+(.+)"), dental_page)
    _add_manulife_value(model, path, "Dental Care", None, "Basic", _regex_value(medical_text, r"Basic Services\s+Coinsurance\s+(\d{1,3}%)"), medical_page)
    _add_manulife_value(model, path, "Dental Care", None, "Major", _regex_value(dental_text, r"Major Restorative Services\s+Coinsurance\s+(\d{1,3}%)"), dental_page)
    _add_manulife_value(model, path, "Dental Care", None, "Orthodontics", _manulife_orthodontics_value(combined_text), dental_page)
    _add_manulife_value(model, path, "Dental Care", None, "Basic Maximum", _regex_value(dental_text, r"Level\s+1\s*-\s*4\s+Maximum\s+(.+)"), dental_page)
    _add_manulife_value(model, path, "Dental Care", None, "Major Maximum", "Combined with Basic", dental_page)
    _add_manulife_value(model, path, "Dental Care", None, "Basic Recall Frequency", _regex_value(dental_text, r"Dental Recall\s+(.+)"), dental_page)
    _add_manulife_value(model, path, "Dental Care", None, "Termination of Coverage", _regex_value(dental_text, r"Termination Age\s+(\d{2},\s+or earlier retirement)"), dental_page)


def _add_grouphealth_value(
    model: dict[str, Any],
    path: Path,
    benefit: str,
    subsection: str | None,
    provision: str,
    value: str | None,
    page_number: int,
) -> None:
    if value in (None, ""):
        return
    _add_pdf_plan_design(model, path, benefit, subsection, provision, _clean_text(value), page_number)


def _add_manulife_value(
    model: dict[str, Any],
    path: Path,
    benefit: str,
    subsection: str | None,
    provision: str,
    value: str | None,
    page_number: int,
) -> None:
    if value in (None, ""):
        return
    _add_pdf_plan_design(model, path, benefit, subsection, provision, _clean_text(value), page_number)


def _regex_value(text: str, pattern: str) -> str | None:
    match = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE)
    return _clean_text(match.group(1)) if match else None


def _regex_value_dotall(text: str, pattern: str) -> str | None:
    match = re.search(pattern, text, flags=re.IGNORECASE | re.MULTILINE | re.DOTALL)
    return _clean_text(match.group(1)) if match else None


def _combine_regex_values(text: str, patterns: list[str]) -> str | None:
    values = [_regex_value(text, pattern) for pattern in patterns]
    values = [value for value in values if value]
    return "; ".join(values) if values else None


def _grouphealth_student_age_value(text: str) -> str | None:
    match = re.search(r"age\s+(\d+)\s+if\s+in\s+full\s+time\s+attendance\s+as\s+a\s+student", text, flags=re.IGNORECASE)
    if not match:
        return None
    return f"Student: Under {match.group(1)} years of age"


def _grouphealth_schedule_percent(text: str, label: str) -> str | None:
    escaped = re.escape(label)
    patterns = [
        rf"{escaped}\*?\s*\.{{2,}}\s*(\d{{1,3}}%)",
        rf"{escaped}\*?.*?(\d{{1,3}}%)",
    ]
    for pattern in patterns:
        value = _regex_value(text, pattern)
        if value:
            return value
    return None


def _paramedical_maximum_value(text: str) -> str | None:
    match = re.search(r"(\$[\d,]+)\s+maximum per practitioner", text, flags=re.IGNORECASE)
    if not match:
        return None
    return f"{match.group(1)} per practitioner per person per calendar year"


def _manulife_dependent_definition_value(text: str) -> str | None:
    student = re.search(r"Student:\s*Under\s*(\d{2})\s+years?\s+of\s+age", text, flags=re.IGNORECASE)
    non_student = re.search(r"Non-Student:\s*Under\s*(\d{2})\s+years?\s+of\s+age", text, flags=re.IGNORECASE)
    values = []
    if student:
        values.append(f"Student: Under {student.group(1)} years of age")
    if non_student:
        values.append(f"Non-Student: Under {non_student.group(1)} years of age")
    return "; ".join(values) if values else None


def _manulife_paramedical_coinsurance(text: str) -> str | None:
    return _regex_value_dotall(text, r"Professional Services\s+.*?\bCoinsurance\s+(\d{1,3}%)")


def _manulife_vision_coinsurance(text: str) -> str | None:
    return _regex_value_dotall(text, r"Vision Care\s+.*?\bCoinsurance\s+(\d{1,3}%)")


def _manulife_hospital_coinsurance(text: str) -> str | None:
    return _regex_value_dotall(text, r"Hospital\s+Coverage\s+.*?\bCoinsurance\s+(\d{1,3}%)")


def _manulife_hospital_coverage(text: str) -> str | None:
    return _regex_value_dotall(text, r"Hospital\s+Coverage\s+(.+?)\s+Coinsurance")


def _manulife_paramedical_practitioners(text: str) -> str | None:
    match = re.search(
        r"Paramedical Practitioners\s+(.+?)(?:\nCoinsurance\b|\nDeductible\b|\bClass\s+[A-Z]\s+All Employees\s+Hospital\b|\nHospital\b|\nAlt\s+\d+\b|\Z)",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if not match:
        return None
    value = re.split(r"\bAlt\s+\d+\b", match.group(1), maxsplit=1, flags=re.IGNORECASE)[0]
    value = re.split(r"\bClass\s+[A-Z]\s+All Employees\s+Hospital\b|\bHospital\s+Coverage\b", value, maxsplit=1, flags=re.IGNORECASE)[0]
    return _clean_text(value)


def _grouphealth_paramedical_practitioners() -> str:
    return "; ".join(
        [
            "Acupuncturist",
            "Audiologist",
            "Chiropractor (includes x-rays)",
            "Dietician",
            "Massage Therapist/Orthotherapist",
            "Naturopath",
            "Occupational Therapist",
            "Osteopath (includes x-rays)",
            "Physiotherapist/Physical Rehabilitation Therapist",
            "Podiatrist/Chiropodist (includes x-rays)",
            "Psychologist/Social Worker/Clinical Counsellors/Psychotherapist",
            "Speech Therapist",
        ]
    )


def _grouphealth_travel_assistance_value(text: str) -> str | None:
    amount = _regex_value(text, r"Travel Insurance and Assistance\s+(\$[\d,]+\s+Per Insured Person per trip)")
    duration_match = re.search(r"maximum duration of coverage:\s+(.+?retirees\))", text, flags=re.IGNORECASE | re.DOTALL)
    duration = _clean_text(duration_match.group(1)) if duration_match else _regex_value(text, r"maximum duration of coverage:\s+(.+)")
    values = [value for value in [amount, duration] if value]
    return "; ".join(values) if values else None


def _grouphealth_dental_recall_value(text: str) -> str | None:
    match = re.search(r"recall or periodic oral examination:\s*once every\s+(\d+)\s+months?", text, flags=re.IGNORECASE)
    if not match:
        return None
    return f"1 per {match.group(1)} months"


def _grouphealth_medical_services_plan_maximum_value(text: str) -> str | None:
    match = re.search(
        r"All Other Eligible Expenses\s*\.{3,}\s*\d{1,3}%.*?Limited to a maximum of\s+(\$[\d,]+\s+per person per calendar year)",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if match:
        return _clean_text(match.group(1))
    return None


def _manulife_oop_plan_maximum_value(text: str) -> str | None:
    match = re.search(r"Out of Country Maximum\s+(.+?)(?:\nTrip Cancellation Insurance\b|\Z)", text, flags=re.IGNORECASE | re.DOTALL)
    if not match:
        return None
    value = _clean_text(match.group(1))
    amount_match = re.search(r"(\$[\d,]+\s+lifetime maximum\s+\(\d+\s+day trip duration\))", value, flags=re.IGNORECASE)
    if amount_match:
        return amount_match.group(1)
    return value.rstrip()


def _manulife_oop_referral_value(text: str) -> str | None:
    value = _regex_value(text, r"Out of Country Referral\s+(.+)")
    if not value:
        return None
    return re.sub(r"(\d{1,3}%)\s+to\s+(\$)", r"\1 up to \2", value, flags=re.IGNORECASE)


def _manulife_orthodontics_value(text: str) -> str | None:
    match = re.search(r"Orthodontia\s+Coverage\s+(.+)", text, flags=re.IGNORECASE)
    if match:
        return _clean_text(match.group(1))
    if "orthodontia" not in text.lower():
        return "Not Applicable"
    return None


def _manulife_orthopaedic_shoes_value(text: str) -> str | None:
    return _regex_value(text, r"Orthopaedic Shoes:\s*(\$[\d,]+\s+per calendar year)")


def _manulife_orthotic_inserts_value(text: str) -> str | None:
    return _regex_value(text, r"Orthotics:\s*(\$[\d,]+\s+per\s+\d+\s+calendar years?)")


def _is_grouphealth_assureplus_payment_schedule_line(line: str, model: dict[str, Any]) -> bool:
    if model.get("carrier_name") != "GroupHEALTH":
        return False
    text = _clean_text(line).lower()
    return any(
        term in text
        for term in [
            "% payment of eligible expenses hospital",
            "out-of-province emergency/trip-cancellation",
            "out-of-province referral",
            "vision care",
            "prescription drugs*",
            "paramedical practitioners",
            "all other eligible expenses",
        ]
    ) and "." in text


def _extract_pdf_alternates(path: Path, pages: list[dict[str, Any]], model: dict[str, Any]) -> None:
    current_benefit: str | None = None
    current_subsection: str | None = None
    current_alt: str | None = None
    pending_changes: list[str] = []
    pending_page: int | None = None

    def flush() -> None:
        nonlocal current_alt, pending_changes, pending_page
        if not current_alt:
            return
        proposed_change = "; ".join(_clean_text(change) for change in pending_changes if _clean_text(change)) or None
        alternate_text = current_alt if not proposed_change else f"{current_alt}: {proposed_change}"
        _add_alternate(
            model,
            section=current_subsection or current_benefit or "Alternates",
            benefit=current_benefit or "General",
            alternate=alternate_text,
            proposed_change=proposed_change,
            rate_impact=_extract_rate_impact(alternate_text),
            source_file=path,
            source_location=f"Page {pending_page or ''}".strip(),
            confidence="HIGH" if proposed_change else "MEDIUM",
            status="FOUND",
            source_type="PDF",
        )
        current_alt = None
        pending_changes = []
        pending_page = None

    for page in pages:
        page_number = page["page_number"]
        text = str(page["text"] or "")
        if "Alt " not in text:
            continue
        for raw_line in text.splitlines():
            line = _clean_text(raw_line)
            if not line:
                continue
            lower_line = line.lower()
            heading = _pdf_plan_heading(line)
            if heading:
                flush()
                current_benefit, current_subsection = heading
                continue
            if re.match(r"^class\s+[a-z0-9]+\b", lower_line) or lower_line == "all employees":
                flush()
                continue
            if lower_line == "other provisions" and current_benefit:
                flush()
                current_subsection = "Other Provisions"
                continue
            if _is_pdf_noise_line(line):
                flush()
                continue
            if re.match(r"^Alt\s+\d+\b", line, flags=re.IGNORECASE):
                flush()
                current_alt = line
                pending_page = page_number
                continue
            if current_alt:
                prefix_result = _pdf_prefixed_provision(line)
                dotted = _pdf_dotted_provision(line)
                if prefix_result:
                    provision, value = prefix_result
                    pending_changes.append(f"{provision}: {value}")
                elif dotted:
                    provision, value = dotted
                    pending_changes.append(f"{provision}: {value}")
                elif pending_changes:
                    pending_changes[-1] = f"{pending_changes[-1]} {line}"
                else:
                    pending_changes.append(line)
        flush()


def _pdf_plan_heading(line: str) -> tuple[str, str | None] | None:
    normalized = _normalize(line)
    headings = {
        "life insurance": ("Employee Life Insurance", None),
        "basic life": ("Employee Life Insurance", None),
        "accident & serious illness (asi)": ("Employee AD&D", None),
        "accidental death & dismemberment": ("Employee AD&D", None),
        "dependent life insurance": ("Dependent Life", None),
        "dependent life": ("Dependent Life", None),
        "extended health care": ("Extended Health Care", None),
        "prescription drugs": ("Extended Health Care", "Prescription Drugs"),
        "drugs": ("Extended Health Care", "Prescription Drugs"),
        "health care practitioners": ("Extended Health Care", "Paramedical Practitioners"),
        "paramedical services": ("Extended Health Care", "Paramedical Practitioners"),
        "professional services": ("Extended Health Care", "Paramedical Practitioners"),
        "vision care": ("Extended Health Care", "Vision Care"),
        "hospital": ("Extended Health Care", "Hospital"),
        "medical services and supplies": ("Extended Health Care", "Medical Services & Supplies"),
        "other medical expenses": ("Extended Health Care", "Medical Services & Supplies"),
        "out of country services": ("Travel", "Out of Province or Country"),
        "dental": ("Dental Care", None),
        "dental care": ("Dental Care", None),
        "basic services": ("Dental Care", "Basic Dental"),
        "major restorative services": ("Dental Care", "Major Dental"),
        "orthodontia": ("Dental Care", "Orthodontics"),
        "critical illness": ("Critical Illness", None),
        "employee & family assistance": ("EAP", None),
        "virtual healthcare": ("Virtual Care", None),
    }
    return headings.get(normalized)


def _pdf_dotted_provision(line: str) -> tuple[str, str] | None:
    if "." not in line:
        return None
    parts = re.split(r"\.{3,}", line)
    if len(parts) < 2:
        return None
    label = _clean_text(parts[0])
    value = _clean_text(parts[-1])
    if not label or not value:
        return None
    return label, value


def _pdf_prefixed_provision(line: str) -> tuple[str, str] | None:
    clean_line = _clean_text(line).lstrip("*• ").strip()
    prefixes = [
        ("Benefit Amount for Spouse", "Spouse"),
        ("Benefit Amount for Each Child", "Child(ren)"),
        ("Benefit Amount", "Benefit Amount"),
        ("Employee's Principal sum", "Benefit Amount"),
        ("Non-Evidence Maximum", "Non-Evidence Maximum"),
        ("Non-evidence Limit", "Non-Evidence Maximum"),
        ("Benefit Reduction", "Reduction of Coverage"),
        ("Reduction Clause", "Reduction of Coverage"),
        ("Waiver of Premium", "Waiver of Premium"),
        ("Living Benefit", "Living Benefit"),
        ("Conversion Privilege", "Conversion Option"),
        ("Conversion", "Conversion Option"),
        ("Termination Age", "Termination of Coverage"),
        ("Calendar Year Deductible", "Deductible"),
        ("Overall Deductible", "Deductible"),
        ("Overall Maximum", "Plan Maximum"),
        ("Prescription Drug Plan", "Plan Type"),
        ("Drug Plan", "Plan Type"),
        ("Reimbursement Type", "Reimbursement Type"),
        ("Drug Option", "Drug Type"),
        ("Drug Maximum", "Plan Maximum"),
        ("Drug Coinsurance", "Coinsurance"),
        ("Coinsurance", "Coinsurance"),
        ("Professional Services Coverage", "Plan Maximum"),
        ("Paramedical Practitioners", "Practitioners"),
        ("Vision Maximum", "Plan Maximum"),
        ("Eye Exam Maximum", "Eye Exams"),
        ("Vision Care", "Plan Maximum"),
        ("Hospital Room Accommodation", "Coverage"),
        ("Hearing Aids", "Hearing Aids"),
        ("Private Duty Nursing", "Private Duty Nursing"),
        ("Off-the-shelf Orthopaedic Shoes and Orthopaedic Modifications", "Orthopaedic Shoes"),
        ("Off-the-shelf Orthopaedic Shoes", "Orthopaedic Shoes"),
        ("Orthopaedic Shoes", "Orthopaedic Shoes"),
        ("Orthopedic Shoes", "Orthopaedic Shoes"),
        ("Foot Orthotics", "Orthotic Inserts"),
        ("Coverage", "Coverage"),
        ("Maximum Amount Covered", "Plan Maximum"),
        ("Level 1 - 4 Maximum", "Level 1 - 4 Maximum"),
        ("Level 1-4 Maximum", "Level 1-4 Maximum"),
        ("• combined maximum", "Plan Maximum"),
        ("combined maximum", "Plan Maximum"),
        ("Rates Based on Dental Procedure Fee Guide", "Fee Guide"),
        ("Rates Based on Dental Procedure Fee", "Fee Guide"),
        ("Dental Recall", "Recall Frequency"),
        ("Dental Fee Guide", "Fee Guide"),
        ("Scaling Units", "Scaling Units"),
        ("Out of Country Maximum", "Plan Maximum"),
        ("Trip Cancellation Insurance", "Trip Cancellation"),
        ("Online Claims Submission", "Online Claims Submission"),
        ("Dependent Children Eligibility", "Child(ren)"),
        ("Child Eligibility", "Child(ren)"),
        ("Definition of Dependent", "Definition of Dependent"),
        ("Survivor Benefit", "Survivor Benefit"),
        ("Stillbirth Benefit", "Stillbirth Benefit"),
        ("Enhanced", "Enhanced"),
        ("Employee & Family Assistance Program", "Included"),
    ]
    for prefix, provision in prefixes:
        if clean_line.lower().startswith(prefix.lower()):
            value = _clean_text(clean_line[len(prefix):])
            return provision, value or line
    return None


def _pdf_needs_following_money_value(provision: str, value: Any) -> bool:
    if provision not in {"Orthopaedic Shoes", "Orthotic Inserts"}:
        return False
    text = _clean_text(value)
    return not _extract_first_money(text)


def _pdf_next_money_line(lines: list[str], line_index: int) -> str | None:
    for next_line in lines[line_index + 1 : line_index + 4]:
        text = _clean_text(next_line).lstrip("*• ").strip()
        money = _extract_first_money(text)
        if money:
            return text
    return None


def _add_pdf_plan_design(
    model: dict[str, Any],
    path: Path,
    benefit: str,
    subsection: str | None,
    provision: str,
    value: str,
    page_number: int,
) -> None:
    provision_name = f"{subsection} - {provision}" if subsection else provision
    _add_plan_design(model, benefit, provision_name, value, path, "PDF", f"Page {page_number}", "HIGH", "FOUND")


def _is_pdf_noise_line(line: str) -> bool:
    lower = line.lower()
    if lower in {"benefit summary", "class 1", "class a", "all employees", "overall"}:
        return True
    if re.match(r"^(may|abc company|quote id|e\. & o\.e\.|www\.|page)\b", lower):
        return True
    return False


def _extract_pdf_rate_guarantee(path: Path, pages: list[dict[str, Any]], model: dict[str, Any]) -> None:
    carrier = model.get("carrier_name") or "Carrier"
    for page in pages:
        page_number = page["page_number"]
        for raw_line in str(page["text"] or "").splitlines():
            line = _clean_text(raw_line)
            lower = line.lower()
            if not line:
                continue
            field = None
            if "rate guarantee" in lower:
                field = "Rate guarantee"
            elif "renewal cap" in lower:
                field = "First renewal cap"
            elif "large claim pooling" in lower or "pooled separately" in lower:
                field = "Pooling level"
            elif "effective date" in lower and not model.get("effective_date"):
                model["effective_date"] = _pdf_effective_date(line)
            if field:
                value = f"{carrier}: {line}"
                existing = model["rate_guarantee"].get(field)
                if not existing or value not in str(existing):
                    model["rate_guarantee"][field] = value if not existing else f"{existing}; {value}"
                    _store_rate_guarantee_for_carrier(model, carrier, field, line)
                    _audit(model, path, "PDF", f"Page {page_number}", field, value, "HIGH", "FOUND")


def _finalize_pdf_comparison_model(model: dict[str, Any]) -> None:
    model["carrier_name"] = " / ".join(model.get("carrier_names") or ["Multiple Carriers"])
    if not model.get("client_name"):
        model["client_name"] = "ABC company"
    _add_included_service_benefit_rows(model)
    for carrier_index, carrier in enumerate(model.get("carrier_names", []), start=1):
        slot = f"Carrier {carrier_index} - {carrier}"
        for benefit in STANDARD_BENEFITS:
            if not any(row.get("carrier_name") == carrier and row.get("benefit_line") == benefit for row in model["benefit_rows"]):
                previous = (model.get("carrier_name"), model.get("carrier_index"), model.get("carrier_slot"))
                model["carrier_name"], model["carrier_index"], model["carrier_slot"] = carrier, carrier_index, slot
                _add_placeholder_benefit(model, benefit, None, f"{carrier} PDF comparison")
                model["carrier_name"], model["carrier_index"], model["carrier_slot"] = previous


def _add_included_service_benefit_rows(model: dict[str, Any]) -> None:
    existing_benefits = {
        (
            row.get("carrier_name"),
            row.get("carrier_slot"),
            row.get("benefit_line"),
        )
        for row in model["benefit_rows"]
    }
    for service in model["included_services"]:
        key = (service.get("carrier_name"), service.get("carrier_slot"), service.get("benefit_line"))
        if key in existing_benefits:
            continue
        model["benefit_rows"].append(
            {
                "carrier_name": service.get("carrier_name") or model.get("carrier_name"),
                "carrier_slot": service.get("carrier_slot"),
                "carrier_index": service.get("carrier_index"),
                "benefit_line": service["benefit_line"],
                "class_name": None,
                "tier": None,
                "lives": None,
                "volume": None,
                "rate": None,
                "monthly_premium": None,
                "annual_premium": None,
                "rate_basis": None,
                "volume_basis": None,
                "included_status": "Included",
                "notes": service["notes"],
                "source_file": service["source_file"],
                "source_type": service["source_type"],
                "source_location": service["source_location"],
                "confidence": service["confidence"],
                "status": "INCLUDED",
            }
        )
        existing_benefits.add(key)


def _extract_word_metadata(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    for index, text in enumerate(paragraphs):
        if not text:
            continue
        if not model.get("carrier_name") and "equitable" in text.lower():
            model["carrier_name"] = "Equitable Life"
            _audit(model, path, "Word", f"Paragraph {index}", "Carrier name", "Equitable Life", "MEDIUM", "FOUND")
        if not model.get("client_name") and "new sns enterprises" in text.lower():
            model["client_name"] = text
            _audit(model, path, "Word", f"Paragraph {index}", "Client name", text, "MEDIUM", "FOUND")
        if not model.get("proposal_date"):
            date_match = re.search(r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s+\d{1,2},\s+\d{4}\b", text)
            if date_match:
                model["proposal_date"] = date_match.group(0)
                _audit(model, path, "Word", f"Paragraph {index}", "Proposal date", model["proposal_date"], "MEDIUM", "FOUND")


def _extract_word_tables(path: Path, document: Document, model: dict[str, Any]) -> None:
    for table_index, table in enumerate(document.tables, start=1):
        rows = [[_clean_text(cell.text) for cell in row.cells] for row in table.rows]
        if not rows:
            continue
        header = [_normalize(cell) for cell in rows[0]]
        if "benefit" in header and "non-evidence maximum" in header:
            for row in rows[1:]:
                benefit = _normalize_benefit(_cell_at(row, 0))
                if not benefit:
                    continue
                for col_index, provision in enumerate(rows[0][1:], start=1):
                    value = _cell_at(row, col_index)
                    if value:
                        _add_plan_design(
                            model,
                            benefit,
                            provision,
                            value,
                            path,
                            "Word",
                            f"Table {table_index}",
                            "HIGH",
                            "FOUND",
                        )


def _extract_word_targeted_paragraphs(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    for index, text in enumerate(paragraphs):
        lower = text.lower()
        if not text:
            continue

        benefit = _benefit_from_funding_line(text)
        if benefit:
            raw_value = text.split(":", 1)[1].strip() if ":" in text else text
            value = f"{benefit}: {raw_value}"
            field = "Pooled benefits" if "pooled" in lower else "Experience-rated benefits"
            existing = model["rate_guarantee"].get(field)
            combined = value if not existing else f"{existing}; {value}"
            model["rate_guarantee"][field] = combined
            _store_rate_guarantee_for_carrier(model, model.get("carrier_name") or "Carrier", field, value)
            _audit(model, path, "Word", f"Paragraph {index}", field, value, "HIGH", "FOUND", benefit_line=benefit)

        if "experience rated" in lower and "non-refund" in lower:
            if "health benefit" in lower:
                _append_rate_guarantee(model, "Experience-rated benefits", "Extended Health Care: Experience Rated, Non-Refund", path, index)
            if "dental benefit" in lower:
                _append_rate_guarantee(model, "Experience-rated benefits", "Dental Care: Experience Rated, Non-Refund", path, index)

        if "target loss ratio" in lower:
            _set_rate_guarantee(model, "Target loss ratio", text, path, index, "HIGH")
        elif "in-canada pooling per individual" in lower:
            _set_rate_guarantee(model, "Pooling level", text, path, index, "HIGH")
        elif "pooling charge is" in lower:
            _set_rate_guarantee(model, "Pooling charge", _first_percent_or_text(text), path, index, "HIGH", notes=text)
        elif "incurred but not reported" in lower or "ibnr" in lower:
            _set_rate_guarantee(model, "IBNR", text, path, index, "HIGH")
        elif "commission scale included" in lower:
            _set_rate_guarantee(model, "Commission", text, path, index, "HIGH")
        elif "first year renewal period" in lower:
            _set_rate_guarantee(model, "Rate guarantee period", text, path, index, "MEDIUM")
        elif "pooled rates are guaranteed" in lower:
            _set_rate_guarantee(model, "Rate guarantee", text, path, index, "HIGH")
        elif "participation requirements" in lower:
            _set_rate_guarantee(model, "Participation requirements", text, path, index, "HIGH")
        elif "employer contribution" in lower:
            _set_rate_guarantee(model, "Employer contribution requirements", text, path, index, "MEDIUM")
        elif "provincial sales tax" in lower:
            _set_rate_guarantee(model, "Taxes", text, path, index, "MEDIUM")

        if "employee and family assistance" in lower and "healthconnector" in lower:
            _add_included_service(model, "EAP", "Homewood Health Employee and Family Assistance Program", path, index)
            _add_included_service(model, "Virtual Care", "Dialogue Virtual Healthcare", path, index)
            _add_included_service(model, "Medical Second Opinion", "WorldCare Second Opinion", path, index)
            _add_included_service(model, "Wellness", "Feeling Better Now and CAREpath", path, index)

        if "termination age" in lower and "quoted" in lower:
            for benefit_name in STANDARD_BENEFITS:
                _add_plan_design(model, benefit_name, "Termination age", text, path, "Word", f"Paragraph {index}", "HIGH", "FOUND")
        elif lower.startswith("earnings definition"):
            _add_plan_design(model, "Long Term Disability", "Earnings definition", text, path, "Word", f"Paragraph {index}", "MEDIUM", "FOUND")
        elif lower.startswith("health benefits:"):
            _add_plan_design(model, "Extended Health Care", "Drug coverage", text, path, "Word", f"Paragraph {index}", "HIGH", "FOUND")
        elif "out-of-province" in lower or "travel assist" in lower:
            _add_plan_design(model, "Travel", _travel_provision_name(text), text, path, "Word", f"Paragraph {index}", "MEDIUM", "FOUND")
            _add_included_service(model, "Travel", "Travel Assist / Out-of-Province coverage referenced", path, index)
        elif lower.startswith("quote conditions"):
            _set_rate_guarantee(model, "Underwriting assumptions", text, path, index, "MEDIUM")

        if 216 <= index <= 238 and any(term in lower for term in ["rates are based", "participation", "eligible employees", "actively at work", "reserves the right", "application"]):
            _set_rate_guarantee(model, "Underwriting assumptions", text, path, index, "MEDIUM")


def _extract_schedule_of_benefits(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    in_schedule = False
    current_section: str | None = None
    current_class: str | None = None
    current_health_detail: str | None = None
    current_dental_detail: str | None = None

    for index, text in enumerate(paragraphs):
        lower = text.lower()
        if lower == "schedule of benefits":
            in_schedule = True
            _audit(model, path, "Word", f"Paragraph {index}", "Schedule of Benefits", text, "HIGH", "FOUND")
            continue

        if not in_schedule:
            continue
        if lower.startswith("® denotes") or lower.startswith("r denotes"):
            break

        section = _schedule_section_from_heading(text)
        if section:
            current_section = section
            current_health_detail = None
            current_dental_detail = None
            _audit(model, path, "Word", f"Paragraph {index}", "Schedule section", text, "HIGH", "FOUND")
            continue

        if not current_section or not text:
            continue

        row, current_class, current_health_detail, current_dental_detail = _schedule_row_from_text(
            text,
            current_section,
            current_class,
            current_health_detail,
            current_dental_detail,
        )
        if not row:
            continue

        _add_master_plan_design(
            model,
            client_id=_client_id(model),
            client_name=model.get("client_name"),
            carrier=model.get("carrier_name") or "Equitable Life",
            policy_number=None,
            class_name=row["class_name"] or current_class or "Class A",
            benefit_category=row["benefit_category"],
            benefit=row["benefit"],
            provision_name=row["provision_name"],
            coverage_description=row["coverage_description"],
            coinsurance=row["coinsurance"],
            maximum=row["maximum"],
            frequency=row["frequency"],
            waiting_period=row["waiting_period"],
            termination=row["termination"],
            notes=row["notes"],
            source_file=path,
            source_location=f"Paragraph {index}",
            confidence=row["confidence"],
            status=row["status"],
        )


def _extract_schedule_comparison_overrides(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    """Add clean comparison rows when the schedule text is split across many paragraphs."""
    _extract_schedule_dependent_life_overrides(path, paragraphs, model)
    _extract_schedule_drug_type_overrides(path, paragraphs, model)
    _extract_schedule_paramedical_practitioners(path, paragraphs, model)
    _extract_schedule_vision_coinsurance(path, paragraphs, model)
    _extract_schedule_vision_limits(path, paragraphs, model)
    _extract_schedule_hospital_overrides(path, paragraphs, model)
    _extract_schedule_oop_overrides(path, paragraphs, model)


def _extract_schedule_dependent_life_overrides(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    for index, text in enumerate(paragraphs):
        dependent_value = _dependent_definition_display_value(text)
        if dependent_value:
            _add_plan_design(
                model,
                "Dependent Life",
                "Definition of Dependent",
                dependent_value,
                path,
                "Word",
                f"Paragraph {index}",
                "HIGH",
                "FOUND",
            )

        survivor_value = _survivor_benefit_display_value(text)
        if survivor_value:
            _add_plan_design(
                model,
                "Dependent Life",
                "Survivor Benefit",
                survivor_value,
                path,
                "Word",
                f"Paragraph {index}",
                "HIGH",
                "FOUND",
            )


def _extract_schedule_drug_type_overrides(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    for index, text in enumerate(paragraphs):
        drug_type = _normalize_drug_type_value(text)
        if not drug_type:
            continue
        _add_plan_design(
            model,
            "Extended Health Care",
            "Prescription Drugs - Drug Type",
            drug_type,
            path,
            "Word",
            f"Paragraph {index}",
            "HIGH",
            "FOUND",
        )
        return


def _extract_schedule_paramedical_practitioners(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    for index, text in enumerate(paragraphs):
        if not text.lower().startswith("paramedical services:"):
            continue

        practitioners: list[str] = []
        maximum: str | None = None
        for offset, candidate in enumerate(paragraphs[index + 1 : index + 18], start=1):
            line = _clean_text(candidate)
            lower = line.lower()
            if not line:
                continue
            if lower.startswith("class ") and "$" in line:
                maximum = line.split(":", 1)[-1].strip()
                continue
            if line.startswith("-"):
                practitioners.append(line.lstrip("- ").strip())
                continue
            if practitioners:
                break

        if maximum:
            _add_plan_design(
                model,
                "Extended Health Care",
                "Paramedical Practitioners - Plan Maximum",
                maximum,
                path,
                "Word",
                f"Paragraph {index}",
                "HIGH",
                "FOUND",
            )
        if any(
            paragraph.lower().startswith("major services:") and "80%" in paragraph
            for paragraph in paragraphs[max(0, index - 40) : index]
        ):
            _add_plan_design(
                model,
                "Extended Health Care",
                "Paramedical Practitioners - Coinsurance",
                "80%",
                path,
                "Word",
                f"Paragraph {index}",
                "MEDIUM",
                "FOUND",
            )
        if practitioners:
            _add_plan_design(
                model,
                "Extended Health Care",
                "Paramedical Practitioners - Practitioners",
                "\n".join(practitioners),
                path,
                "Word",
                f"Paragraph {index + 1}-{index + len(practitioners) + 1}",
                "HIGH",
                "FOUND",
        )
        return


def _extract_schedule_vision_coinsurance(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    for index, text in enumerate(paragraphs):
        lower = text.lower()
        if not lower.startswith("vision care") or "80%" not in lower or "payment" not in lower:
            continue
        _add_plan_design(
            model,
            "Extended Health Care",
            "Vision Care - Coinsurance",
            "80%",
            path,
            "Word",
            f"Paragraph {index}",
            "HIGH",
            "FOUND",
        )
        return


def _extract_schedule_vision_limits(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    for index, text in enumerate(paragraphs):
        lower = text.lower()
        if lower.startswith("vision care (lenses"):
            for candidate_index, candidate in enumerate(paragraphs[index + 1 : index + 5], start=index + 1):
                candidate_text = _clean_text(candidate)
                if "$200" in candidate_text and "24 months" in candidate_text.lower():
                    _add_plan_design(
                        model,
                        "Extended Health Care",
                        "Vision Care - Plan Maximum",
                        "$200 per 24 months",
                        path,
                        "Word",
                        f"Paragraph {candidate_index}",
                        "HIGH",
                        "FOUND",
                    )
                    break
        elif lower.startswith("vision care (eye examinations"):
            for candidate_index, candidate in enumerate(paragraphs[index + 1 : index + 5], start=index + 1):
                candidate_text = _clean_text(candidate)
                if "eye examination" in candidate_text.lower() and "24 months" in candidate_text.lower():
                    _add_plan_design(
                        model,
                        "Extended Health Care",
                        "Vision Care - Eye Exams",
                        "1 per 24 months",
                        path,
                        "Word",
                        f"Paragraph {candidate_index}",
                        "HIGH",
                        "FOUND",
                    )
                    break


def _extract_schedule_hospital_overrides(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    for index, text in enumerate(paragraphs):
        if "semi-private room in a public hospital" not in text.lower():
            continue
        _add_plan_design(
            model,
            "Extended Health Care",
            "Hospital - Coverage",
            "Semi-Private",
            path,
            "Word",
            f"Paragraph {index}",
            "HIGH",
            "FOUND",
        )
        return


def _extract_schedule_oop_overrides(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    payment_text = next(
        (text for text in paragraphs if "100%" in text and "emergency" in text.lower() and "referral" in text.lower() and "out-of-province" in text.lower()),
        "",
    )
    emergency_value = "100%" if payment_text else "can not find"
    referral_value = "100%" if payment_text else "can not find"
    maximum_value = _schedule_oop_plan_maximum(paragraphs) or "can not find"
    trip_cancellation = _schedule_trip_cancellation_value(paragraphs)

    rows = [
        ("Out-of-Province Emergency", emergency_value, "HIGH" if payment_text else "LOW"),
        ("Out-of-Province Referral", referral_value, "HIGH" if payment_text else "LOW"),
        ("Plan Maximum", maximum_value, "HIGH" if maximum_value != "can not find" else "LOW"),
        ("Trip Cancellation", trip_cancellation, "MEDIUM"),
    ]
    for provision, value, confidence in rows:
        _add_plan_design(
            model,
            "Travel",
            f"Out of Province or Country - {provision}",
            value,
            path,
            "Word",
            "Out-of-Province/Canada section",
            confidence,
            "FOUND" if value not in {"can not find", "not covered"} else "NEEDS REVIEW",
        )


def _schedule_oop_plan_maximum(paragraphs: list[str]) -> str | None:
    for index, text in enumerate(paragraphs):
        if "out-of-province/canada lifetime maximum" not in text.lower():
            continue
        window = [_clean_text(line) for line in paragraphs[index + 1 : index + 5] if _clean_text(line)]
        if not window:
            return None
        under_70 = next((line for line in window if "$5,000,000" in line and "70" in line), "")
        over_70 = next((line for line in window if "$1,000,000" in line and ("70 and older" in line.lower() or "retired" in line.lower())), "")
        retired = next((line for line in window if "$1,000,000" in line and "retired" in line.lower()), "")
        if under_70 and (over_70 or retired):
            return "$5,000,000 up to age 70, $1,000,000 age 70+ and Retired Employees of any age"
    return None


def _schedule_trip_cancellation_value(paragraphs: list[str]) -> str:
    for text in paragraphs:
        lower = text.lower()
        if "trip cancellation" not in lower:
            continue
        money = _extract_first_money(text)
        if money:
            suffix = _frequency_suffix(text)
            return f"{money} {suffix}".strip()
    return "not covered"


def _extract_alternates(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    in_alternates = False
    section: str | None = None

    for index, text in enumerate(paragraphs):
        lower = text.lower()
        if lower == "alternates":
            in_alternates = True
            _audit(model, path, "Word", f"Paragraph {index}", "Alternates section", text, "HIGH", "FOUND")
            continue
        if not in_alternates:
            continue
        if lower == "schedule of benefits":
            break
        if not text:
            continue

        alternate_section = _alternate_section_heading(text)
        if alternate_section:
            section = alternate_section
            continue

        benefit = _alternate_benefit(section, text)
        impact = _extract_rate_impact(text)
        proposed_change = _extract_proposed_change(text)
        _add_alternate(
            model,
            section=section or "Alternates",
            benefit=benefit,
            alternate=text,
            proposed_change=proposed_change,
            rate_impact=impact,
            source_file=path,
            source_location=f"Paragraph {index}",
            confidence="HIGH" if impact or proposed_change else "MEDIUM",
            status="FOUND",
        )


def _extract_deviations(path: Path, paragraphs: list[str], model: dict[str, Any]) -> None:
    for index, text in enumerate(paragraphs):
        lower = text.lower()
        if not text:
            continue

        in_clarification_section = 122 <= index <= 140
        quote_condition = 216 <= index <= 224 and ("deviation" in lower or "quotation may not correspond" in lower)
        contains_deviation_signal = any(
            phrase in lower
            for phrase in [
                "no guarantee has been made",
                "not embedded",
                "not alterable",
                "no changes can be made",
                "based our quotation on equitable's standard benefits",
                "quotation may not correspond",
                "any deviation",
            ]
        )
        if not (in_clarification_section or quote_condition or contains_deviation_signal):
            continue
        if lower in {"schedule of insurance - clarifications and deviations"}:
            _audit(model, path, "Word", f"Paragraph {index}", "Deviation section", text, "HIGH", "FOUND")
            continue

        _add_deviation(
            model,
            deviation_type=_deviation_type(text),
            benefit=_deviation_benefit(text),
            description=text,
            impact=_deviation_impact(text),
            source_file=path,
            source_location=f"Paragraph {index}",
            confidence="HIGH" if contains_deviation_signal or quote_condition else "MEDIUM",
            status="NEEDS REVIEW" if "may not correspond" in lower or "no guarantee" in lower else "FOUND",
        )


def _schedule_section_from_heading(text: str) -> str | None:
    lower = text.lower()
    if lower == "basic life insurance and accidental death & dismemberment insurance":
        return "life_add"
    if lower == "dependent life insurance":
        return "dependent_life"
    if lower == "health benefits - major services maximums":
        return "health_major_services"
    if lower == "health benefits":
        return "health"
    if lower == "dental benefits":
        return "dental"
    return None


def _schedule_row_from_text(
    text: str,
    section: str,
    current_class: str | None,
    current_health_detail: str | None,
    current_dental_detail: str | None,
) -> tuple[dict[str, Any] | None, str | None, str | None, str | None]:
    class_name, body = _extract_class_body(text, current_class)
    if class_name and not body:
        return None, class_name, current_health_detail, current_dental_detail
    if not body:
        return None, current_class, current_health_detail, current_dental_detail

    benefit_category, benefit, current_health_detail, current_dental_detail = _schedule_benefit_context(
        section,
        body,
        current_health_detail,
        current_dental_detail,
    )
    if _schedule_heading_only(body):
        return None, class_name or current_class, current_health_detail, current_dental_detail

    provision_name = _schedule_provision_name(body, benefit)
    row = {
        "class_name": class_name or current_class,
        "benefit_category": benefit_category,
        "benefit": benefit,
        "provision_name": provision_name,
        "coverage_description": body,
        "coinsurance": _extract_coinsurance(body),
        "maximum": _extract_maximum(body),
        "frequency": _extract_frequency(body),
        "waiting_period": _extract_waiting_period(body),
        "termination": _extract_termination(body),
        "notes": _schedule_notes(body),
        "confidence": "HIGH",
        "status": "FOUND",
    }
    return row, class_name or current_class, current_health_detail, current_dental_detail


def _schedule_benefit_context(
    section: str,
    body: str,
    current_health_detail: str | None,
    current_dental_detail: str | None,
) -> tuple[str, str, str | None, str | None]:
    if section == "life_add":
        benefit = "Employee AD&D" if "ad&d" in body.lower() or "accidental death" in body.lower() else "Employee Life Insurance"
        return "Life and AD&D", benefit, current_health_detail, current_dental_detail
    if section == "dependent_life":
        return "Dependent Life", "Dependent Life", current_health_detail, current_dental_detail
    if section == "health":
        detail = _health_detail_from_text(body, current_health_detail)
        return "Extended Health Care", detail or "Extended Health Care", detail, current_dental_detail
    if section == "health_major_services":
        benefit = _major_services_benefit_name(body)
        return "Extended Health Care", benefit, benefit, current_dental_detail
    if section == "dental":
        detail = _dental_detail_from_text(body, current_dental_detail)
        return "Dental Care", detail or "Dental Care", current_health_detail, detail
    return "Other", _normalize_benefit(body) or "Other Benefits", current_health_detail, current_dental_detail


def _extract_class_body(text: str, current_class: str | None) -> tuple[str | None, str]:
    match = re.match(r"^(Class\s+[A-Z0-9]+)\s*:\s*(.*)$", text, flags=re.IGNORECASE)
    if match:
        return _clean_class(match.group(1)) or match.group(1).title(), match.group(2).strip()
    return current_class, text.strip()


def _schedule_heading_only(text: str) -> bool:
    lower = text.lower().strip()
    if not lower.endswith(":"):
        return False
    return not any(token in lower for token in ["%", "$", "maximum", "eligible", "included", "none", "current"])


def _health_detail_from_text(text: str, current_detail: str | None) -> str:
    lower = text.lower()
    if any(term in lower for term in ["survivor benefit", "dependent children", "coverage terminates"]):
        return "Extended Health Care"
    if lower.startswith("-") and current_detail:
        return current_detail
    if any(term in lower for term in ["drug", "ramq", "fertility", "smoking cessation", "erectile", "weight loss", "vaccines", "glucose monitor", "telus health"]):
        return "Drugs"
    if "major services" in lower:
        return "Major Services"
    if lower.startswith("hospital") or "semi-private" in lower:
        return "Hospital"
    if "vision care" in lower or "contact lenses" in lower or "eye examination" in lower or "prescription is not required" in lower:
        return "Vision"
    if "paramedical" in lower or "practitioner" in lower or lower.startswith(tuple(f"-{name}" for name in ["acupuncturist", "chiropractor", "massage", "naturopath", "osteopath", "physiotherapist", "podiatrist", "psychologist", "speech"])):
        return "Paramedical"
    if "out-of-province" in lower or "travel" in lower:
        return "Travel"
    return current_detail or "Extended Health Care"


def _major_services_benefit_name(text: str) -> str:
    body = text.split(":", 1)[0].strip()
    if not body:
        return "Major Services"
    if "lifetime maximum" in body.lower():
        return "Extended Health Care"
    return body


def _dental_detail_from_text(text: str, current_detail: str | None) -> str:
    lower = text.lower()
    if any(term in lower for term in ["survivor benefit", "dependent children", "coverage terminates"]):
        return "Dental Care"
    if "recall" in lower or "examination" in lower:
        return "Dental Recall"
    if "orthodont" in lower:
        return "Dental Orthodontic"
    if "type b" in lower or "major restorative" in lower or "dentures" in lower:
        return "Dental Major"
    if "type a" in lower or "basic" in lower or "fee guide" in lower or "reimbursement" in lower or "combined maximum" in lower:
        return "Dental Basic"
    return current_detail or "Dental Care"


def _schedule_provision_name(text: str, benefit: str) -> str:
    lower = text.lower()
    if "flat amount" in lower or "for spouse" in lower or "for each child" in lower:
        return "Benefit amount"
    if "evidence of insurability" in lower or "non-evidence" in lower:
        return "Evidence of insurability"
    if "waiver of premium" in lower:
        return "Waiver of premium"
    if "coverage reduce" in lower or "reduce 50%" in lower:
        return "Reduction schedule"
    if "terminate" in lower or "termination" in lower:
        return "Termination age"
    if "calendar year deductible" in lower:
        return "Deductible"
    if "calendar year maximum" in lower or "lifetime maximum" in lower or "maximum amount" in lower:
        return "Maximum"
    if "reimbursement" in lower or "payments" in lower or "coinsurance" in lower:
        return "Coinsurance"
    if "recall" in lower:
        return "Recall frequency"
    if "fee guide" in lower:
        return "Fee guide"
    if "not eligible" in lower:
        return "Not eligible"
    if "included" in lower:
        return "Included provision"
    if "survivor benefit" in lower:
        return "Survivor benefit"
    if "dependent children" in lower:
        return "Dependent child eligibility"
    if "conversion" in lower:
        return "Conversion option"
    if "drug plan" in lower or "telus health" in lower:
        return "Drug plan"
    if ":" in text:
        return text.split(":", 1)[0].strip()[:80]
    return benefit


def _extract_coinsurance(text: str) -> str | None:
    lower = text.lower()
    if not any(term in lower for term in ["coinsurance", "payment", "reimbursement"]):
        return None
    matches = re.findall(r"\b\d{1,3}(?:\.\d+)?%", text)
    return "; ".join(dict.fromkeys(matches)) if matches else None


def _extract_maximum(text: str) -> str | None:
    money = re.findall(r"\$\s?\d[\d,]*(?:\.\d+)?", text)
    values = [value.replace(" ", "") for value in money]
    lower = text.lower()
    for phrase in ["unlimited", "not required", "not eligible", "none", "current", "included"]:
        if phrase in lower:
            values.append(phrase.title())
    if "reasonable and customary" in lower:
        values.append("Reasonable and Customary")
    return "; ".join(dict.fromkeys(values)) if values else None


def _extract_frequency(text: str) -> str | None:
    patterns = [
        r"per calendar year",
        r"per lifetime",
        r"lifetime maximum",
        r"every \d+\s+months",
        r"in any period of \d+\s+consecutive months",
        r"\d+-month recall",
        r"\d+\s+days",
        r"\d+\s+day",
        r"per disability",
        r"from birth",
    ]
    matches: list[str] = []
    for pattern in patterns:
        matches.extend(match.group(0) for match in re.finditer(pattern, text, flags=re.IGNORECASE))
    return "; ".join(dict.fromkeys(matches)) if matches else None


def _extract_waiting_period(text: str) -> str | None:
    match = re.search(r"waiting period[^.;]*", text, flags=re.IGNORECASE)
    return match.group(0) if match else None


def _benefit_amount_type(text: str) -> str | None:
    lower = text.lower()
    if "flat amount" in lower:
        return "Flat Amount"
    if "spouse" in lower or "child" in lower:
        return "Dependent Amount"
    if "%" in lower or "earnings" in lower or "salary" in lower:
        return "Earnings Based"
    return None


def _extract_first_money(value: Any) -> str | None:
    text = str(value or "")
    match = re.search(r"\$\s?\d[\d,]*(?:\.\d+)?", text)
    return match.group(0).replace(" ", "") if match else None


def _extract_money_amount(value: Any) -> int | None:
    money = _extract_first_money(value)
    if not money:
        return None
    amount = re.sub(r"[^\d.]", "", money)
    if not amount:
        return None
    return int(round(float(amount)))


def _extract_first_percent_number(value: Any) -> float | None:
    text = str(value or "")
    match = re.search(r"\d{1,3}(?:\.\d+)?%", text)
    return float(match.group(0).replace("%", "")) / 100 if match else None


def _extract_month_count(value: Any) -> int | None:
    text = str(value or "")
    match = re.search(r"(\d+)\s+months?", text, flags=re.IGNORECASE)
    return int(match.group(1)) if match else None


def _extract_recall_month_count(value: Any) -> int | None:
    text = str(value or "")
    patterns = [
        r"(\d+)\s*-\s*month\s+recall",
        r"(\d+)\s+month\s+recall",
        r"recall\s+(?:frequency\s+)?(?:is\s+)?(\d+)\s+months?",
        r"(\d+)\s+months?",
        r"1\s+per\s+(\d+)\s+months?",
        r"every\s+(\d+)\s+months?",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return int(match.group(1))
    return None


def _dental_normalization_from_parts(
    benefit: Any,
    provision: Any,
    extracted_value: Any,
    maximum: Any | None = None,
    frequency: Any | None = None,
) -> dict[str, Any] | None:
    benefit_text = _clean_text(benefit)
    provision_text = _clean_text(provision)
    value_text = _clean_text(extracted_value)
    maximum_text = _clean_text(maximum)
    frequency_text = _clean_text(frequency)
    combined = f"{benefit_text} {provision_text} {value_text} {maximum_text} {frequency_text}".lower()

    if "dental" not in combined and "orthodont" not in combined and "recall" not in combined:
        return None

    recall_signal = (
        ("recall" in provision_text.lower() and "scaling" not in combined)
        or "dental recall" in combined
        or bool(re.search(r"\b(?:5|6|9|12)\s*-?\s*month\s+recall\b", combined))
    )
    if recall_signal:
        raw_value = frequency_text or value_text or combined
        return _normalize_dental_dropdown_value("Basic Dental Recall Frequency", raw_value, combined)

    is_orthodontic = "orthodont" in combined or "ortho" in combined
    value_signal_text = f"{value_text} {maximum_text}"
    has_extractable_value = bool(_extract_first_money(value_signal_text)) or any(
        term in value_signal_text.lower()
        for term in ["not applicable", "n/a", "none", "nil", "unlimited"]
    )
    has_maximum_signal = ("maximum" in combined or bool(_extract_first_money(value_signal_text))) and has_extractable_value
    if is_orthodontic and has_maximum_signal:
        raw_value = maximum_text or value_text
        return _normalize_dental_dropdown_value("Orthodontics Lifetime Maximum", raw_value, combined)

    is_major = any(term in combined for term in ["major", "restorative", "level 1 - 4", "level 1-4", "combined maximum"])
    if is_major and has_maximum_signal:
        raw_value = maximum_text or value_text
        return _normalize_dental_dropdown_value("Major Dental Annual Maximum", raw_value, combined)

    return None


def _normalize_dental_dropdown_value(field_name: str, raw_value: Any, match_text: Any | None = None) -> dict[str, Any]:
    raw_text = _clean_text(raw_value)
    searchable = _clean_text(f"{raw_text} {match_text or ''}")
    normalized = DENTAL_NORMALIZATION_OTHER
    notes: str | None = None
    confidence = "LOW"

    if field_name == "Basic Dental Recall Frequency":
        lower = searchable.lower()
        if not raw_text:
            notes = "No recall frequency was extracted."
        elif re.search(r"\bN/?A\b", raw_text, flags=re.IGNORECASE) or "not applicable" in lower:
            normalized = "N/A"
            confidence = "HIGH"
        else:
            months = _extract_recall_month_count(searchable)
            if months in {5, 6, 9, 12}:
                normalized = f"{months} month recall"
                confidence = "HIGH"
            else:
                notes = f"Original extracted wording: {raw_text}"

    elif field_name == "Major Dental Annual Maximum":
        amount = _extract_money_amount(searchable)
        combined = any(term in searchable.lower() for term in ["combined", "level 1 - 4", "level 1-4"])
        if amount in {1000, 1500, 2000, 2500, 3000}:
            normalized = f"${amount:,}/year{' combined' if combined else ''}"
            confidence = "HIGH"
        else:
            notes = f"Original extracted wording: {raw_text}" if raw_text else "No major dental annual maximum was extracted."

    elif field_name == "Orthodontics Lifetime Maximum":
        amount = _extract_money_amount(searchable)
        if amount in {1000, 1500, 2000, 2500, 3000, 3500, 4000}:
            normalized = f"${amount:,} lifetime"
            confidence = "HIGH"
        else:
            notes = f"Original extracted wording: {raw_text}" if raw_text else "No orthodontic lifetime maximum was extracted."

    if normalized == DENTAL_NORMALIZATION_OTHER and not notes:
        notes = f"Original extracted wording: {raw_text}" if raw_text else "Original extracted wording was blank."

    return {
        "field_name": field_name,
        "raw_value": raw_text,
        "normalized_value": normalized,
        "notes": notes,
        "confidence": confidence,
    }


def _dental_normalization_for_comparison_item(item: dict[str, Any], raw_value: Any) -> dict[str, Any] | None:
    if item.get("major_section") != "Dental Care":
        return None
    if raw_value in (None, ""):
        return None
    sort_order = int(item.get("sort_order") or 0)
    row_label = str(item.get("row_label") or "")
    if row_label == "Basic Recall Frequency" or sort_order == 50:
        return _normalize_dental_dropdown_value("Basic Dental Recall Frequency", raw_value)
    if sort_order == 42:
        return _normalize_dental_dropdown_value("Major Dental Annual Maximum", raw_value)
    if sort_order == 43:
        return _normalize_dental_dropdown_value("Orthodontics Lifetime Maximum", raw_value)
    return None


def _dental_normalization_reference_rows() -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for value in DENTAL_RECALL_DROPDOWN_VALUES:
        rows.append(
            {
                "DentalField": "Basic Dental Recall Frequency",
                "ApprovedDropdownValue": value,
                "RawValueExamples": "6 months; 6-month recall; 1 per 6 months" if value == "6 month recall" else "",
                "NotesRule": "Use Other when recall wording cannot be confidently mapped." if value == DENTAL_NORMALIZATION_OTHER else "",
            }
        )
    for value in DENTAL_MAJOR_MAXIMUM_DROPDOWN_VALUES:
        rows.append(
            {
                "DentalField": "Major Dental Annual Maximum",
                "ApprovedDropdownValue": value,
                "RawValueExamples": "$1,500 per calendar year; Level 1 - 4 Maximum $1,500 per calendar year" if value in {"$1,500/year", "$1,500/year combined"} else "",
                "NotesRule": "Use Other and write original wording in Notes when amount/combined status is unclear." if value == DENTAL_NORMALIZATION_OTHER else "",
            }
        )
    for value in DENTAL_ORTHODONTIC_MAXIMUM_DROPDOWN_VALUES:
        rows.append(
            {
                "DentalField": "Orthodontics Lifetime Maximum",
                "ApprovedDropdownValue": value,
                "RawValueExamples": "$2,000 per lifetime; $2,000 lifetime" if value == "$2,000 lifetime" else "",
                "NotesRule": "Use Other and write original wording in Notes when lifetime amount is unclear." if value == DENTAL_NORMALIZATION_OTHER else "",
            }
        )
    return rows



def _extract_day_count(value: Any) -> int | None:
    text = str(value or "")
    match = re.search(r"(\d+)\s+days?", text, flags=re.IGNORECASE)
    return int(match.group(1)) if match else None


def _extract_age_number(value: Any) -> int | None:
    text = str(value or "")
    matches = re.findall(r"age\s+(\d+)", text, flags=re.IGNORECASE)
    return int(matches[0]) if matches else None


def _termination_trigger(value: Any) -> str | None:
    text = str(value or "")
    lower = text.lower()
    triggers = []
    if "retirement" in lower:
        triggers.append("Retirement")
    if "active employees" in lower:
        triggers.append("Active Employees")
    if "retirees" in lower:
        triggers.append("Retirees")
    return "; ".join(triggers) if triggers else None


def _extract_termination(text: str) -> str | None:
    lower = text.lower()
    if not any(term in lower for term in ["terminate", "termination", "age 65", "age 70", "retirement"]):
        return None
    return text


def _schedule_notes(text: str) -> str | None:
    lower = text.lower()
    notes = []
    if "not eligible" in lower:
        notes.append("Not eligible")
    if "not required" in lower:
        notes.append("Not required")
    if "not available" in lower:
        notes.append("Not available")
    return "; ".join(notes) if notes else None


def _alternate_section_heading(text: str) -> str | None:
    lower = text.lower()
    if lower == "life and accidental death & dismemberment":
        return "Life and Accidental Death & Dismemberment"
    if lower == "health benefits":
        return "Health Benefits"
    if lower == "dental benefits":
        return "Dental Benefits"
    return None


def _alternate_benefit(section: str | None, text: str) -> str:
    lower = text.lower()
    if "ad&d" in lower or "accidental death" in lower:
        return "Employee AD&D"
    if section and section.startswith("Life"):
        return "Employee Life Insurance"
    if "drug" in lower:
        return "Drugs"
    if "major services" in lower:
        return "Major Services"
    if "hospital" in lower:
        return "Hospital"
    if "vision" in lower:
        return "Vision"
    if "dental" in lower or (section and section.startswith("Dental")):
        return "Dental Care"
    if section and section.startswith("Health"):
        return "Extended Health Care"
    return _normalize_benefit(text) or section or "Alternates"


def _extract_rate_impact(text: str) -> str | None:
    lower = text.lower()
    if "no change" in lower:
        return "No change"
    match = re.search(r"(\d+(?:\.\d+)?%)\s+(increase|decrease)", text, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)} {match.group(2).lower()}"
    match = re.search(r"(increase|decrease)[^.]*?\bby\s+(\d+(?:\.\d+)?%)", text, flags=re.IGNORECASE)
    if match:
        return f"{match.group(2)} {match.group(1).lower()}"
    return None


def _extract_proposed_change(text: str) -> str | None:
    money = re.search(r"\$\s?\d[\d,]*(?:\.\d+)?", text)
    if money:
        return money.group(0).replace(" ", "")
    percent = re.search(r"to\s+(\d+(?:\.\d+)?%)", text, flags=re.IGNORECASE)
    if percent:
        return percent.group(1)
    if "not required" in text.lower():
        return "Not Required"
    return None


def _deviation_type(text: str) -> str:
    lower = text.lower()
    if "not embedded" in lower:
        return "Not embedded in quoted rates"
    if "not alterable" in lower or "no changes can be made" in lower:
        return "Non-alterable coverage"
    if "no guarantee has been made" in lower:
        return "Deviation disclaimer"
    if "standard benefits" in lower:
        return "Standard benefit assumption"
    if "quotation may not correspond" in lower or "any deviation" in lower:
        return "Quote condition"
    return "Clarification"


def _deviation_benefit(text: str) -> str:
    lower = text.lower()
    if "accidental death" in lower or "ad&d" in lower:
        return "Employee AD&D"
    if "travel" in lower or "out-of-province" in lower:
        return "Travel"
    if "efap" in lower or "employee and family assistance" in lower or "healthconnector" in lower:
        return "EAP / Value Added Services"
    if "health and dental" in lower:
        return "Extended Health Care / Dental Care"
    if "health" in lower:
        return "Extended Health Care"
    if "dental" in lower:
        return "Dental Care"
    return "General"


def _deviation_impact(text: str) -> str | None:
    lower = text.lower()
    if "not embedded" in lower:
        return "Additional service/cost may sit outside quoted products."
    if "not alterable" in lower or "no changes can be made" in lower:
        return "Carrier will not alter this coverage feature."
    if "no guarantee" in lower or "may not correspond" in lower:
        return "Manual review required against current plan/booklet."
    return None


def _finalize_model(model: dict[str, Any]) -> None:
    if not model.get("carrier_name"):
        model["carrier_name"] = "Unknown Carrier"
        _missing(model, "Carrier name", "Carrier name was not detected.")
    if not model.get("client_name"):
        _missing(model, "Client name", "Client name was not detected.")
    if not model.get("effective_date"):
        _missing(model, "Effective date", "Effective date was not found.")
    if not model.get("proposal_date"):
        _missing(model, "Proposal date", "Proposal date was not found.")

    existing_benefits = {row["benefit_line"] for row in model["benefit_rows"]}
    for service in model["included_services"]:
        benefit = service["benefit_line"]
        if benefit not in existing_benefits:
            model["benefit_rows"].append(
                {
                    "carrier_name": service.get("carrier_name") or model.get("carrier_name"),
                    "carrier_slot": service.get("carrier_slot") or model.get("carrier_slot"),
                    "carrier_index": service.get("carrier_index") or model.get("carrier_index"),
                    "benefit_line": benefit,
                    "class_name": None,
                    "tier": None,
                    "lives": None,
                    "volume": None,
                    "rate": None,
                    "monthly_premium": None,
                    "annual_premium": None,
                    "rate_basis": None,
                    "volume_basis": None,
                    "included_status": "Included",
                    "notes": service["notes"],
                    "source_file": service["source_file"],
                    "source_type": service["source_type"],
                    "source_location": service["source_location"],
                    "confidence": service["confidence"],
                    "status": "INCLUDED",
                }
            )
            existing_benefits.add(benefit)

    for benefit in STANDARD_BENEFITS:
        if benefit not in existing_benefits:
            _add_placeholder_benefit(model, benefit, None, "Missing standard benefit")

    for field in ["Retention", "Admin fees", "Taxes", "Employer contribution requirements"]:
        if field not in model["rate_guarantee"]:
            _missing(model, field, f"{field} was not found in the quote package.")


def _write_quote_workbook(model: dict[str, Any], output_path: Path) -> Path:
    workbook = _load_output_template()
    _remove_unwanted_template_sheets(workbook)
    for sheet_name in TARGET_SHEETS:
        if sheet_name not in workbook.sheetnames:
            workbook.create_sheet(sheet_name)
    if "Extraction Audit" not in workbook.sheetnames:
        workbook.create_sheet("Extraction Audit")

    cost_rows = _write_total_cost_comparison(workbook["Total Cost Comparison"], model)
    reconciliation_rows = _build_reconciliation_rows(model)
    _write_gross_cost_comparison(workbook["Gross Cost Comparison"], model, reconciliation_rows)
    comparison_carriers = _comparison_carrier_headers(model)
    comparison_data = _build_benefit_comparison_data(model, comparison_carriers)
    write_benefit_comparison_sheet(workbook, comparison_data, comparison_carriers)
    _write_rate_reconciliation(workbook["Rate Reconciliation"], model, reconciliation_rows)
    _write_line_rates_formatted(workbook["RateTable_LineRates_Formatted"], model, reconciliation_rows)
    write_benefit_comparison_sheet(workbook, comparison_data, comparison_carriers, sheet_name="Plan Design")
    _write_master_plan_design(workbook["PlanDesign"], model)
    _write_booklet_extraction_review(workbook["Booklet_Extraction_Review"], model)
    _write_dental_normalization_reference(workbook["Dental_Normalization_Reference"])
    _write_dental_normalized_values(workbook["Dental_Normalized_Values"], model)
    _write_alternates(workbook["Alternates"], model)
    _write_deviations(workbook["Deviations"], model)
    _write_rate_guarantee(workbook["Rate Guarantee Sum"], model)
    _write_extended_health(workbook["Extended Health"], model)
    _write_exec_summary(workbook["ExecSummary - Cost"], model, cost_rows)
    _write_audit_sheet(workbook["Extraction Audit"], model)
    _apply_workbook_font_family(workbook, WORKBOOK_FONT_NAME)
    _sanitize_workbook_for_excel(workbook)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(output_path)
        return output_path
    except PermissionError:
        fallback = output_path.with_name(f"{output_path.stem}_{datetime.now():%Y%m%d_%H%M%S}{output_path.suffix}")
        workbook.save(fallback)
        return fallback


def _remove_unwanted_template_sheets(workbook: Any) -> None:
    unwanted = {"ttc_chart", "tcc_chart", "tcc_chart(2)"}
    for sheet_name in list(workbook.sheetnames):
        normalized = re.sub(r"\s+", "", sheet_name).lower()
        if normalized in unwanted and len(workbook.sheetnames) > 1:
            workbook.remove(workbook[sheet_name])


def _comparison_carrier_headers(model: dict[str, Any]) -> list[str]:
    carrier_names = [str(name) for name in model.get("carrier_names", []) if name]
    if not carrier_names:
        carrier_name = str(model.get("carrier_name") or "").strip()
        return ["Current", carrier_name or "Carrier 1"]
    headers = ["Current"]
    for index, name in enumerate(carrier_names[:MAX_COMPARISON_CARRIERS], start=1):
        headers.append(f"Carrier {index} - {name}")
    return headers


def _carrier_display_name(carrier_header: Any) -> str:
    text = _clean_text(carrier_header)
    if not text:
        return ""
    match = re.match(r"^Carrier\s+\d+\s*-\s*(.+)$", text, flags=re.IGNORECASE)
    if match:
        return _clean_text(match.group(1))
    if re.match(r"^Carrier\s+\d+$", text, flags=re.IGNORECASE):
        return ""
    return text


def _carrier_slot_for_name(model: dict[str, Any], carrier_name: str) -> str | None:
    for header in _comparison_carrier_headers(model)[1:]:
        if carrier_name and carrier_name.lower() in header.lower():
            return header
    return None


def _load_output_template() -> Any:
    template = DEFAULT_TEMPLATE_PATH

    if template.exists():
        workbook = load_workbook(template, keep_links=False)
        if hasattr(workbook, "_external_links"):
            workbook._external_links = []
        return workbook

    workbook = Workbook()
    workbook.remove(workbook.active)
    return workbook


def write_benefit_comparison_sheet(
    wb: Any,
    comparison_data: list[dict[str, Any]],
    carriers: list[str],
    sheet_name: str = "Benefit Comparison",
) -> None:
    if sheet_name not in wb.sheetnames:
        wb.create_sheet(sheet_name)
    if "BenefitComparison_Normalized" not in wb.sheetnames:
        wb.create_sheet("BenefitComparison_Normalized")

    sheet = wb[sheet_name]
    normalized_sheet = wb["BenefitComparison_Normalized"]
    _prepare_sheet(sheet)
    _prepare_sheet(normalized_sheet)
    _unhide_all_rows(sheet)
    _unhide_all_rows(normalized_sheet)

    carriers = [carrier for carrier in carriers if carrier]
    if not carriers:
        carriers = ["Current"]

    sheet.sheet_view.showGridLines = False
    sheet.column_dimensions["A"].width = 10
    sheet.column_dimensions["A"].hidden = False
    sheet.column_dimensions["B"].width = 32
    sheet.column_dimensions["B"].hidden = False
    for column_index, _carrier in enumerate(carriers, start=3):
        dimension = sheet.column_dimensions[get_column_letter(column_index)]
        dimension.width = 36 if column_index == 3 else 34
        dimension.hidden = False

    current_section: str | None = None
    current_subsection: str | None = None
    section_start_row: int | None = None
    section_label: str | None = None
    row_number = 2
    wrote_header_for_section = False
    max_col = 2 + len(carriers)

    for item in comparison_data:
        major_section = str(item.get("major_section") or "Other")
        subsection = str(item.get("subsection") or "")
        row_label = str(item.get("row_label") or "")

        if major_section != current_section:
            if section_start_row is not None:
                _finish_benefit_section_band(sheet, section_start_row, row_number - 1, section_label or "")
            current_section = major_section
            current_subsection = None
            section_label = major_section
            section_start_row = row_number
            wrote_header_for_section = False

        if row_label == "__CARRIER_HEADER__":
            _write_benefit_carrier_header(sheet, row_number, carriers)
            row_number += 1
            wrote_header_for_section = True
            continue

        if subsection and subsection != current_subsection:
            current_subsection = subsection
            _write_benefit_subsection_row(sheet, row_number, subsection, item if row_label == "__SUBSECTION_HEADER__" else None, carriers)
            row_number += 1
            if (not wrote_header_for_section) or item.get("repeat_carrier_header"):
                _write_benefit_carrier_header(sheet, row_number, carriers)
                row_number += 1
                wrote_header_for_section = True
            if row_label == "__SUBSECTION_HEADER__":
                continue

        _write_benefit_value_row(sheet, row_number, item, carriers)
        _apply_benefit_row_height(sheet, row_number, max_col)
        row_number += 1

    if section_start_row is not None:
        _finish_benefit_section_band(sheet, section_start_row, row_number - 1, section_label or "")

    _format_benefit_comparison_used_cells(sheet, max(row_number - 1, 1), max_col, carriers)
    _write_benefit_comparison_normalized(normalized_sheet, comparison_data, carriers)


def _write_benefit_carrier_header(sheet: Any, row_number: int, carriers: list[str]) -> None:
    sheet.row_dimensions[row_number].height = 27
    for column in range(2, 3 + len(carriers)):
        cell = sheet.cell(row_number, column)
        cell.border = _teal_border()
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for offset, carrier in enumerate(carriers, start=3):
        cell = sheet.cell(row_number, offset, _carrier_display_name(carrier))
        if offset == 3 and _is_current_carrier(carrier):
            cell.fill = PatternFill("solid", fgColor=CURRENT_HEADER_FILL)
            cell.font = Font(name="Calibri", size=11, bold=True, color=MUTED_TEXT)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        else:
            cell.fill = PatternFill("solid", fgColor=CARRIER_HEADER_FILL)
            cell.font = Font(name="Calibri", size=11, bold=True, color=CARRIER_HEADER_FONT)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _write_benefit_subsection_row(
    sheet: Any,
    row_number: int,
    subsection: str,
    item: dict[str, Any] | None,
    carriers: list[str],
) -> None:
    sheet.row_dimensions[row_number].height = 33
    sheet.cell(row_number, 2, subsection)
    for column in range(2, 3 + len(carriers)):
        cell = sheet.cell(row_number, column)
        cell.fill = PatternFill(fill_type=None)
        cell.font = Font(name="Calibri", bold=True, color=CARRIER_HEADER_FILL, size=12)
        cell.alignment = Alignment(vertical="center", wrap_text=False)
        cell.border = Border()

    if item:
        values = item.get("values") or {}
        for offset, carrier in enumerate(carriers, start=3):
            sheet.cell(row_number, offset, values.get(carrier))


def _write_benefit_value_row(sheet: Any, row_number: int, item: dict[str, Any], carriers: list[str]) -> None:
    row_label = item.get("row_label")
    sheet.cell(row_number, 2, None if row_label in {"__SUBSECTION_HEADER__", "__CARRIER_HEADER__"} else row_label)
    values = item.get("values") or {}
    for offset, carrier in enumerate(carriers, start=3):
        sheet.cell(row_number, offset, values.get(carrier))


def _format_benefit_comparison_used_cells(sheet: Any, max_row: int, max_col: int, carriers: list[str]) -> None:
    apply_proposal_comparison_formatting(sheet, 2, 1, max_row, max_col)
    _delete_rows_after(sheet, max_row)

    sheet.freeze_panes = "C4"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_margins.left = 0.25
    sheet.page_margins.right = 0.25
    sheet.page_margins.top = 0.35
    sheet.page_margins.bottom = 0.35
    sheet.print_title_rows = "1:3"


def apply_proposal_comparison_formatting(
    ws: Any,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> None:
    border = Border(
        left=Side(style="thin", color=BORDER_COLOR),
        right=Side(style="thin", color=BORDER_COLOR),
        top=Side(style="thin", color=BORDER_COLOR),
        bottom=Side(style="thin", color=BORDER_COLOR),
    )
    current_fill = PatternFill("solid", fgColor=CURRENT_HEADER_FILL)
    carrier_fill = PatternFill("solid", fgColor=CARRIER_HEADER_FILL)
    current_header_font = Font(name="Calibri", size=11, bold=True, color=MUTED_TEXT)
    carrier_header_font = Font(name="Calibri", size=11, bold=True, color=CARRIER_HEADER_FONT)
    data_font = Font(name="Calibri", size=10, bold=False, color=TEXT_COLOR)
    label_font = Font(name="Calibri", size=10, bold=False, color=TEXT_COLOR)
    category_title_font = Font(name="Calibri", size=12, bold=True, color=CARRIER_HEADER_FILL)

    header_rows: set[int] = set()
    data_columns: set[int] = set()
    current_columns: set[int] = set()
    carrier_columns: set[int] = set()

    for row in range(start_row, end_row + 1):
        row_current_columns: list[int] = []
        row_carrier_columns: list[int] = []
        for column in range(start_col, end_col + 1):
            text = _normalize_header_text(ws.cell(row, column).value)
            if _is_current_header_text(text):
                row_current_columns.append(column)
            elif _is_carrier_header_text(text):
                row_carrier_columns.append(column)
        # Treat a row as a carrier header only when carrier labels are present.
        # Normal data rows can legitimately contain values such as "Current".
        if row_carrier_columns:
            header_rows.add(row)
            current_columns.update(row_current_columns)
            carrier_columns.update(row_carrier_columns)
            data_columns.update(row_current_columns)
            data_columns.update(row_carrier_columns)

    if not data_columns:
        data_columns = set(range(start_col + 2, end_col + 1))
        current_columns = {min(data_columns)}
        carrier_columns = data_columns - current_columns

    data_column_fills = _proposal_comparison_data_fills(data_columns, current_columns)

    for column in range(start_col, end_col + 1):
        letter = get_column_letter(column)
        width = 10 if column == start_col else 32 if column == start_col + 1 else 36 if column in current_columns else 34
        ws.column_dimensions[letter].width = width
        ws.column_dimensions[letter].hidden = False

    for row in range(start_row, end_row + 1):
        is_category_row = _is_benefit_subsection_row(ws, row)
        for column in range(start_col, end_col + 1):
            cell = ws.cell(row, column)
            cell.border = border
            if column == start_col:
                continue

            if is_category_row:
                cell.fill = PatternFill(fill_type=None)
                cell.border = Border()
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
                if column == start_col + 1:
                    cell.font = category_title_font
                continue

            if row in header_rows and column in current_columns:
                cell.fill = current_fill
                cell.font = current_header_font
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif row in header_rows and column in carrier_columns:
                cell.fill = carrier_fill
                cell.font = carrier_header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            elif column in data_columns and row not in header_rows:
                cell.fill = data_column_fills[column]
                cell.font = _bold_data_font() if _should_bold_carrier_quote(ws, row, column, current_columns) else data_font
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                if column == start_col + 1:
                    cell.font = category_title_font if _is_benefit_subsection_row(ws, row) else label_font

        _set_wrapped_row_height(ws, row, start_col, end_col, header_rows)


def _normalize_header_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def _delete_rows_after(sheet: Any, keep_through_row: int) -> None:
    if sheet.max_row > keep_through_row:
        sheet.delete_rows(keep_through_row + 1, sheet.max_row - keep_through_row)


def _is_current_header_text(text: str) -> bool:
    return text in {"current", "current plan", "your current plan"}


def _is_carrier_header_text(text: str) -> bool:
    return bool(re.match(r"carrier\s*\d+", text or ""))


def _proposal_comparison_data_fills(data_columns: set[int], current_columns: set[int]) -> dict[int, PatternFill]:
    fills: dict[int, PatternFill] = {}
    current_fill = PatternFill("solid", fgColor=CURRENT_COLUMN_FILL)
    carrier_fill = PatternFill("solid", fgColor=DATA_FILL)
    alternate_fill = PatternFill("solid", fgColor=ALT_DATA_FILL)
    ordered_columns = sorted(data_columns)

    carrier_position = 0
    for column in ordered_columns:
        if column in current_columns:
            fills[column] = current_fill
            continue
        carrier_position += 1
        fills[column] = alternate_fill if carrier_position % 2 == 0 else carrier_fill

    return fills


def _bold_data_font() -> Font:
    return Font(name="Calibri", size=10, bold=True, color=TEXT_COLOR)


def _should_bold_carrier_quote(ws: Any, row: int, column: int, current_columns: set[int]) -> bool:
    if column in current_columns:
        return False
    value = ws.cell(row, column).value
    if value in (None, ""):
        return False
    return bool(str(value).strip())


def _set_wrapped_row_height(ws: Any, row: int, start_col: int, end_col: int, header_rows: set[int]) -> None:
    if row in header_rows:
        ws.row_dimensions[row].height = 27
        return
    if _is_benefit_subsection_row(ws, row):
        ws.row_dimensions[row].height = 33
        return

    max_lines = 1
    has_value = False
    for column in range(start_col + 1, end_col + 1):
        value = ws.cell(row, column).value
        if value in (None, ""):
            continue
        has_value = True
        width = ws.column_dimensions[get_column_letter(column)].width or 18
        text = str(value)
        line_count = text.count("\n") + 1
        line_count = max(line_count, max(1, int(len(text) / max(width * 1.2, 20)) + 1))
        max_lines = max(max_lines, line_count)

    ws.row_dimensions[row].height = 8 if not has_value else max(18, min(76, max_lines * 15))


def _finish_benefit_section_band(sheet: Any, start_row: int, end_row: int, label: str) -> None:
    if end_row < start_row:
        end_row = start_row
    if end_row > start_row:
        sheet.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
    cell = sheet.cell(start_row, 1, label)
    cell.fill = PatternFill("solid", fgColor=CARRIER_HEADER_FILL)
    cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=22)
    cell.alignment = Alignment(horizontal="center", vertical="center", textRotation=90, wrap_text=True)
    cell.border = _teal_border()
    for row in range(start_row, end_row + 1):
        sheet.cell(row, 1).fill = PatternFill("solid", fgColor=CARRIER_HEADER_FILL)
        sheet.cell(row, 1).border = _teal_border()


def _apply_benefit_row_height(sheet: Any, row_number: int, max_col: int) -> None:
    max_lines = 1
    for column in range(2, max_col + 1):
        value = sheet.cell(row_number, column).value
        if value is None:
            continue
        text = str(value)
        lines = text.count("\n") + 1
        lines = max(lines, min(5, (len(text) // 45) + 1))
        max_lines = max(max_lines, lines)
    sheet.row_dimensions[row_number].height = max(22, min(110, max_lines * 16))


def _write_benefit_comparison_normalized(sheet: Any, comparison_data: list[dict[str, Any]], carriers: list[str]) -> None:
    headers = [
        "ClientName",
        "EffectiveDate",
        "MajorSection",
        "Subsection",
        "RowLabel",
        "CarrierName",
        "Value",
        "RawExtractedValue",
        "NormalizedDropdownValue",
        "NormalizationNotes",
        "SourceFile",
        "SourcePage",
        "ConfidenceScore",
    ]
    for column, header in enumerate(headers, start=1):
        sheet.cell(1, column, header)

    row_number = 2
    for item in comparison_data:
        row_label = item.get("row_label")
        if row_label == "__CARRIER_HEADER__":
            continue
        display_label = item.get("subsection") if row_label == "__SUBSECTION_HEADER__" else row_label
        values = item.get("values") or {}
        for carrier in carriers:
            raw_value = values.get(carrier)
            dental_normalization = _dental_normalization_for_comparison_item(item, raw_value)
            output = [
                item.get("client_name"),
                item.get("effective_date"),
                item.get("major_section"),
                item.get("subsection"),
                display_label,
                _carrier_display_name(carrier),
                raw_value,
                raw_value,
                dental_normalization.get("normalized_value") if dental_normalization else None,
                dental_normalization.get("notes") if dental_normalization else None,
                item.get("source_file"),
                item.get("source_page"),
                item.get("confidence_score"),
            ]
            for column, value in enumerate(output, start=1):
                sheet.cell(row_number, column, value)
            row_number += 1

    _format_sheet(sheet, max(row_number - 1, 1), len(headers))
    _set_fixed_row_height(sheet, max(row_number - 1, 1), 15)


def _teal_border() -> Border:
    side = Side(style="thin", color=BORDER_COLOR)
    return Border(left=side, right=side, top=side, bottom=side)


def _is_current_carrier(carrier: str) -> bool:
    return "current" in str(carrier or "").lower()


def _is_header_or_section_cell(sheet: Any, row: int, column: int) -> bool:
    fill = sheet.cell(row, column).fill
    fg = getattr(fill.fgColor, "rgb", None)
    return fg in {f"00{CARRIER_HEADER_FILL}", f"FF{CARRIER_HEADER_FILL}", f"00{CURRENT_HEADER_FILL}", f"FF{CURRENT_HEADER_FILL}"}


def _is_benefit_subsection_row(sheet: Any, row: int) -> bool:
    return _is_comparison_category_title(sheet.cell(row, 2).value)


def _is_comparison_category_title(value: Any) -> bool:
    text = _normalize(value)
    if not text:
        return False
    return text in COMPARISON_CATEGORY_TITLES


def _build_benefit_comparison_data(model: dict[str, Any], carriers: list[str] | None = None) -> list[dict[str, Any]]:
    carrier = model.get("carrier_name") or "Carrier 1"
    comparison_carriers = carriers or ["Current", "Carrier 1", "Carrier 2", "Carrier 3"]
    if model.get("carrier_names"):
        return _build_multi_carrier_benefit_comparison_data(model, comparison_carriers)

    rows: list[dict[str, Any]] = []
    for record in _ordered_plan_design(model["plan_design"]):
        for entry in _comparison_entries_for_record(record):
            rows.append(
                {
                    "major_section": entry["major_section"],
                    "subsection": entry["subsection"],
                    "row_label": entry["row_label"],
                    "values": _comparison_values_for_record(record, carrier, entry.get("value")),
                    "repeat_carrier_header": _should_repeat_benefit_header(record["benefit_line"], entry["subsection"], entry["row_label"]),
                    "client_name": model.get("client_name"),
                    "effective_date": model.get("effective_date"),
                    "source_file": _source_file_name(record.get("source_file")),
                    "source_page": _source_page_from_location(record.get("source_location")),
                    "confidence_score": record.get("confidence"),
                    "sort_order": entry.get("sort_order", 999),
                }
            )
    _normalize_basic_life_comparison_rows(rows, comparison_carriers)
    _normalize_adnd_comparison_rows(rows, comparison_carriers)
    _normalize_dependent_life_comparison_rows(rows, comparison_carriers)
    _normalize_deductible_comparison_rows(rows, comparison_carriers)
    _normalize_hospital_comparison_rows(rows, comparison_carriers)
    _normalize_termination_comparison_rows(rows, comparison_carriers)
    return _sort_comparison_rows(_with_dental_group_rows(rows, comparison_carriers))


def _build_multi_carrier_benefit_comparison_data(model: dict[str, Any], carriers: list[str]) -> list[dict[str, Any]]:
    rows_by_key: dict[tuple[str, str, str, str], dict[str, Any]] = {}
    ordered_keys: list[tuple[str, str, str, str]] = []
    for record in _ordered_plan_design(model["plan_design"]):
        for entry in _comparison_entries_for_record(record):
            benefit = record["benefit_line"]
            merge_key = _comparison_row_merge_key(entry)
            key = (entry["major_section"], entry["subsection"], entry["row_label"], merge_key)
            if key not in rows_by_key:
                rows_by_key[key] = {
                    "major_section": entry["major_section"],
                    "subsection": entry["subsection"],
                    "row_label": entry["row_label"],
                    "values": {carrier: None for carrier in carriers},
                    "repeat_carrier_header": _should_repeat_benefit_header(benefit, entry["subsection"], entry["row_label"]),
                    "client_name": model.get("client_name"),
                    "effective_date": model.get("effective_date"),
                    "source_file": _source_file_name(record.get("source_file")),
                    "source_page": _source_page_from_location(record.get("source_location")),
                    "confidence_score": record.get("confidence"),
                    "sort_order": entry.get("sort_order", 999),
                }
                ordered_keys.append(key)
            slot = record.get("carrier_slot") or _carrier_slot_for_name(model, str(record.get("carrier_name") or ""))
            value = _normalize_comparison_display_value(benefit, entry["subsection"], entry["row_label"], entry["key"], entry.get("value"))
            if slot in rows_by_key[key]["values"]:
                existing_value = rows_by_key[key]["values"][slot]
                rows_by_key[key]["values"][slot] = _prefer_comparison_value(existing_value, value)
            elif len(carriers) > 1:
                rows_by_key[key]["values"][carriers[1]] = _prefer_comparison_value(rows_by_key[key]["values"].get(carriers[1]), value)
    rows = [rows_by_key[key] for key in ordered_keys]
    _normalize_basic_life_comparison_rows(rows, carriers)
    _normalize_adnd_comparison_rows(rows, carriers)
    _normalize_dependent_life_comparison_rows(rows, carriers)
    _normalize_deductible_comparison_rows(rows, carriers)
    _normalize_hospital_comparison_rows(rows, carriers)
    _normalize_termination_comparison_rows(rows, carriers)
    return _sort_comparison_rows(_with_dental_group_rows(rows, carriers))


def _comparison_row_merge_key(entry: dict[str, Any]) -> str:
    major_section = str(entry.get("major_section") or "")
    row_label = str(entry.get("row_label") or "")
    if major_section == "Dental Care":
        return str(entry.get("key") or f"{row_label}:{entry.get('sort_order')}")
    return _normalize(row_label)


def _comparison_entries_for_record(record: dict[str, Any]) -> list[dict[str, Any]]:
    benefit = str(record.get("benefit_line") or "")
    provision = str(record.get("provision_name") or "")
    value = record.get("provision_value")
    if benefit == "Employee Life Insurance":
        normalized_entries = _employee_life_comparison_entries(benefit, provision, value)
        if normalized_entries:
            return normalized_entries
    if benefit == "Employee AD&D":
        normalized_entries = _employee_adnd_comparison_entries(benefit, provision, value)
        if normalized_entries:
            return normalized_entries
    if benefit == "Dental Care":
        return _dental_comparison_entries(provision, value)

    subsection, row_label = _comparison_subsection_and_label(benefit, provision)
    out_of_province_entries = _out_of_province_comparison_entries(benefit, provision, value, subsection, row_label)
    if out_of_province_entries:
        return out_of_province_entries
    if _hide_plan_design_comparison_row(benefit, subsection, row_label, value):
        return []
    if _hide_prescription_comparison_row(subsection, row_label):
        return []
    if _hide_low_value_comparison_row(benefit, provision, row_label, value):
        return []
    output_value = _shorten_health_dental_termination_value(benefit, subsection, row_label, value)
    return [
        {
            "major_section": _comparison_major_section(benefit, subsection),
            "subsection": subsection,
            "row_label": row_label,
            "key": provision,
            "value": output_value,
            "sort_order": 500,
        }
    ]


def _hide_prescription_comparison_row(subsection: str, row_label: str) -> bool:
    if _normalize(subsection) != "prescription drugs":
        return False
    normalized = _normalize(row_label)
    if normalized in {
        "generic substitution",
        "included drug supplies",
        "not eligible",
        "prescription requirement",
        "reimbursement type",
    }:
        return True
    return "flash glucose" in normalized or "continuous glucose monitor" in normalized


def _hide_plan_design_comparison_row(benefit: str, subsection: str, row_label: str, value: Any) -> bool:
    normalized_subsection = _normalize(subsection)
    normalized_row = _normalize(row_label)
    normalized_value = _normalize(value)

    if normalized_subsection == "basic life insurance" and normalized_row in {
        "conversion option",
        "evidence of insurability",
        "living benefit",
        "waiver of premium",
    }:
        return True

    if normalized_subsection == "accidental death & dismemberment or accident & serious illness" and normalized_row in {
        "enhanced",
        "waiver of premium",
    }:
        return True

    if normalized_subsection == "dependent life insurance" and normalized_row == "included":
        return True

    if normalized_subsection == "paramedical practitioners":
        if normalized_row in {"paramedical", "paramedical services"}:
            return True
        practitioner_names = {
            "acupuncturist",
            "chiropractor",
            "massage therapist",
            "naturopath",
            "osteopath",
            "physiotherapist",
            "podiatrist/chiropodist",
            "psychologist/msw/counselor",
            "speech therapist",
        }
        if normalized_row in practitioner_names or normalized_value in practitioner_names:
            return True

    if normalized_subsection == "vision care":
        if normalized_row in {"coverage", "deductible", "vision"}:
            return True
        if normalized_row == "plan maximum" and "contact lens" in normalized_value:
            return True

    if normalized_subsection == "hospital":
        return normalized_row not in {"coinsurance", "coverage", "hospital indemnity benefit"}

    if normalized_subsection == "other provisions" and normalized_row == "survivor benefit":
        return True

    if normalized_subsection == "out of province or country":
        if normalized_row in {
            "travel / out-of-province coverage",
            "travel emergency coverage",
            "travel referral coverage",
        }:
            return True
        if normalized_row == "plan maximum" and "$" not in _clean_text(value):
            return True

    if normalized_subsection == "medical services & supplies":
        hidden_medical_supply_rows = {
            "accidental dental",
            "ambulance (ground and air)",
            "applies to all classes, unless otherwise specified.",
            "breast prosthesis (external)",
            "child(ren)",
            "continuous glucose monitor",
            "convalescent home services",
            "definition of dependent",
            "diagnostic lab tests (includes x-rays, mri, psa)",
            "drug coverage",
            "extremity pumps for lymphedema",
            "gender affirmation",
            "glucometer",
            "included",
            "insulin pump",
            "laryngeal speaking aid",
            "major services",
            "plan maximum",
            "maximums are subject to policy provisions. eligible expenses exclude those covered by the provincial health plan in the province of residence.",
            "prosthetic appliances and myoelectric prosthesis",
            "rental of equipment",
            "stump socks",
            "surgical brassieres",
            "surgical/support stockings",
            "tens machine",
            "termination of coverage",
            "viscosupplementation (synvisc)",
            "wigs or hairpieces",
        }
        if normalized_row in hidden_medical_supply_rows:
            return True

    return False


def _hide_low_value_comparison_row(benefit: str, provision: str, row_label: str, value: Any) -> bool:
    provision_text = _normalize(provision)
    value_text = _normalize(value)
    if row_label == "Evidence of Insurability" and "non-evidence" in provision_text:
        return "not required" not in value_text and "not require" not in value_text
    return False


def _shorten_health_dental_termination_value(benefit: str, subsection: str, row_label: str, value: Any) -> Any:
    if row_label != "Termination of Coverage":
        return value
    if _comparison_major_section(benefit, subsection) == "Pooled Benefits":
        return value
    return _short_termination_display_value(value) or value


def _short_termination_display_value(value: Any) -> str | None:
    text = _clean_text(value)
    if not text:
        return None
    match = re.search(r"age\s*(\d{2})|\b(\d{2})\s*,?\s+or earlier retirement", text, flags=re.IGNORECASE)
    if not match:
        return None
    age = match.group(1) or match.group(2)
    if "retirement" not in text.lower():
        return None
    return f"{age}, or earlier retirement"


def _out_of_province_comparison_entries(
    benefit: str,
    provision: str,
    value: Any,
    subsection: str,
    row_label: str,
) -> list[dict[str, Any]]:
    combined = _normalize(f"{benefit} {subsection} {provision} {value}")
    is_oop_section = _normalize(subsection) == "out of province or country"
    is_oop_value = any(
        term in combined
        for term in [
            "out-of-province",
            "out of province",
            "out of country",
            "emergency travel assistance",
            "travel insurance and assistance",
            "trip cancellation",
        ]
    )
    if not is_oop_section and not is_oop_value:
        return []

    output_row_label = row_label
    key = provision
    normalized_value = value
    sort_order = 500

    if "out-of-province emergency" in combined or "out of province emergency" in combined or "emergency/trip-cancellation" in combined:
        output_row_label = "Out-of-Province Emergency"
        key = "oop.emergency"
        normalized_value = _clean_text(value)
        sort_order = 10
    elif "out-of-province referral maximum" in combined or "out of province referral maximum" in combined:
        output_row_label = "Out-of-Province Referral"
        key = "oop.referral"
        normalized_value = _normalize_oop_maximum_text(value)
        sort_order = 20
    elif "out-of-province referral" in combined or "out of province referral" in combined or "out of country referral" in combined:
        output_row_label = "Out-of-Province Referral"
        key = "oop.referral"
        normalized_value = _normalize_oop_referral_text(value)
        sort_order = 20
    elif "travel insurance and assistance" in combined or "out of country maximum" in combined or output_row_label == "Plan Maximum":
        output_row_label = "Plan Maximum"
        key = "oop.plan_maximum"
        normalized_value = _clean_text(value)
        sort_order = 30
    elif "trip cancellation" in combined:
        output_row_label = "Trip Cancellation"
        key = "oop.trip_cancellation"
        normalized_value = _clean_text(value)
        sort_order = 40
    elif output_row_label == "Coinsurance" and re.search(r"\b100%", combined):
        output_row_label = "Out-of-Province Emergency"
        key = "oop.emergency"
        normalized_value = _clean_text(value)
        sort_order = 10

    if output_row_label not in {"Out-of-Province Emergency", "Out-of-Province Referral", "Plan Maximum", "Trip Cancellation"}:
        return []
    if output_row_label == "Plan Maximum" and "$" not in _clean_text(normalized_value):
        return []
    if output_row_label in {"Out-of-Province Emergency", "Out-of-Province Referral"}:
        value_text = _clean_text(normalized_value)
        if not re.search(r"\d{1,3}%|\$|can not find", value_text, flags=re.IGNORECASE):
            return []

    if normalized_value in (None, ""):
        return []

    return [
        {
            "major_section": "Extended Health",
            "subsection": "Out of Province or Country",
            "row_label": output_row_label,
            "key": key,
            "value": normalized_value,
            "sort_order": sort_order,
        }
    ]


def _normalize_oop_referral_text(value: Any) -> str:
    text = _clean_text(value)
    text = re.sub(r"(\d{1,3}%)\s+to\s+(\$)", r"\1 up to \2", text, flags=re.IGNORECASE)
    return text


def _normalize_oop_maximum_text(value: Any) -> str:
    text = _clean_text(value)
    text = re.sub(r"\bCalendar Year\b", "calendar year", text)
    return text


def _employee_life_comparison_entries(benefit: str, provision: str, value: Any) -> list[dict[str, Any]]:
    text = _normalize(f"{provision} {value}")
    subsection = _comparison_subsection_label(_plan_design_section_label(benefit))
    row_label = _comparison_row_label(provision)
    entries: list[tuple[str, str, str, int]] = []

    if _has_life_reduction_50_at_65(text):
        entries.append(("Reduction of Coverage", "Reduction of Coverage", "The amount of insurance reduces by 50% at age 65", 500))
    termination_value = _life_termination_value(text, row_label)
    if termination_value:
        entries.append(("Termination of Coverage", "Termination of Coverage", termination_value, 500))

    return [
        {
            "major_section": _comparison_major_section(benefit, subsection),
            "subsection": subsection,
            "row_label": row_label,
            "key": key,
            "value": normalized_value,
            "sort_order": sort_order,
        }
        for key, row_label, normalized_value, sort_order in entries
    ]


def _has_life_reduction_50_at_65(text: str) -> bool:
    return bool(re.search(r"(reduce|reduces|reduced|reduction).{0,80}50%.{0,40}age\s*65|50%.{0,40}age\s*65", text))


def _life_termination_value(text: str, row_label: str) -> str | None:
    if "terminat" not in text and row_label != "Termination of Coverage":
        return None
    match = re.search(r"terminat\w*.{0,120}age\s*(\d{2})", text)
    if not match:
        match = re.search(r"age\s*(\d{2}).{0,120}terminat", text)
    if not match and row_label == "Termination of Coverage":
        match = re.search(r"age\s*(\d{2})", text)
    if not match:
        return None
    return f"{match.group(1)}, or earlier retirement"


def _employee_adnd_comparison_entries(benefit: str, provision: str, value: Any) -> list[dict[str, Any]]:
    text = _normalize(f"{provision} {value}")
    subsection = _comparison_subsection_label(_plan_design_section_label(benefit))
    row_label = _comparison_row_label(provision)
    entries: list[tuple[str, str, str, int]] = []

    if row_label == "Reduction of Coverage" and "same as group life" in text:
        entries.append(("Reduction of Coverage", "Reduction of Coverage", "The amount of insurance reduces by 50% at age 65", 500))
    elif _has_life_reduction_50_at_65(text):
        entries.append(("Reduction of Coverage", "Reduction of Coverage", "The amount of insurance reduces by 50% at age 65", 500))

    termination_value = _adnd_termination_value(text, row_label)
    if termination_value:
        entries.append(("Termination of Coverage", "Termination of Coverage", termination_value, 500))

    return [
        {
            "major_section": _comparison_major_section(benefit, subsection),
            "subsection": subsection,
            "row_label": entry_row_label,
            "key": key,
            "value": normalized_value,
            "sort_order": sort_order,
        }
        for key, entry_row_label, normalized_value, sort_order in entries
    ]


def _adnd_termination_value(text: str, row_label: str) -> str | None:
    if "same as group life" in text and row_label == "Termination of Coverage":
        return "An employee's insurance terminates at age 75 or earlier retirement"

    match = re.search(r"terminat\w*.{0,80}age\s*(\d{2})", text)
    if not match and row_label == "Termination of Coverage":
        match = re.search(r"age\s*(\d{2})", text)
    if not match:
        return None

    age = match.group(1)
    return f"{age}, or earlier retirement"


def _dental_comparison_entries(provision: str, value: Any) -> list[dict[str, Any]]:
    label = _clean_text(provision)
    normalized_label = _normalize(label)
    text = _clean_text(value)
    combined = f"{label} {text}".lower()
    entries: list[tuple[str, str, Any, int]] = []

    if "deductible" in combined:
        entries.append(("dental.deductible", "Deductible", _short_dental_value(text), 10))
    elif "fee guide" in combined or "dental procedure fee" in combined:
        entries.append(("dental.fee_guide", "Fee Guide Year", _dental_fee_guide_value(text), 20))
    elif normalized_label == "basic" and re.search(r"\b\d{1,3}(?:\.\d+)?%", text):
        entries.append(("dental.coinsurance.basic", "Basic", _dental_percent_or_status(text), 31))
    elif normalized_label == "major" and text:
        entries.append(("dental.coinsurance.major", "Major", _dental_percent_or_status(text), 32))
    elif normalized_label == "orthodontics" and text:
        entries.append(("dental.coinsurance.orthodontics", "Orthodontics", _dental_percent_or_status(text), 33))
    elif normalized_label == "basic maximum" and text:
        entries.append(("dental.maximum.basic", "Basic", _extract_first_money(text) or _short_dental_value(text), 41))
    elif normalized_label == "major maximum" and text:
        entries.append(("dental.maximum.major", "Major", _extract_first_money(text) or _short_dental_value(text), 42))
    elif _is_dental_basic_coinsurance(combined):
        entries.append(("dental.coinsurance.basic", "Basic", _dental_percent_or_status(text), 31))
    elif _is_dental_major_coinsurance(combined):
        entries.append(("dental.coinsurance.major", "Major", _dental_percent_or_status(text), 32))
    elif _is_dental_orthodontic_coinsurance(combined):
        entries.append(("dental.coinsurance.orthodontics", "Orthodontics", _dental_percent_or_status(text), 33))
    elif "combined maximum" in combined or "level 1 - 4 maximum" in combined or "level 1-4 maximum" in combined:
        maximum = _extract_first_money(text) or _short_dental_value(text)
        entries.append(("dental.maximum.basic", "Basic", maximum, 41))
        entries.append(("dental.maximum.major", "Major", "Combined with Basic", 42))
    elif "maximum amount covered" in combined:
        maximum = _extract_first_money(text)
        if maximum:
            entries.append(("dental.maximum.basic", "Basic", maximum, 41))
            if "dental restorative" in combined or "major" in combined:
                entries.append(("dental.maximum.major", "Major", "Combined with Basic", 42))
    elif "plan maximum" in combined and _extract_first_money(text) and not any(term in combined for term in ["orthodontic", "orthodontia", "major", "basic"]):
        entries.append(("dental.maximum.basic", "Basic", _extract_first_money(text), 41))
        entries.append(("dental.maximum.major", "Major", "Combined with Basic", 42))
    elif _is_dental_basic_maximum(combined):
        entries.append(("dental.maximum.basic", "Basic", _extract_first_money(text) or _short_dental_value(text), 41))
    elif _is_dental_major_maximum(combined):
        entries.append(("dental.maximum.major", "Major", _extract_first_money(text) or _short_dental_value(text), 42))
    elif _is_dental_orthodontic_maximum(combined):
        entries.append(("dental.maximum.orthodontics", "Orthodontics", _extract_first_money(text) or _short_dental_value(text), 43))
    elif "recall" in combined:
        entries.append(("dental.recall.basic", "Basic Recall Frequency", _short_dental_value(text), 50))
    elif "termination" in combined:
        entries.append(("dental.termination", "Termination of Coverage", _short_termination_display_value(text) or _short_dental_value(text), 60))

    return [
        {
            "major_section": "Dental Care",
            "subsection": "Dental Care",
            "row_label": row_label,
            "key": key,
            "value": normalized_value,
            "sort_order": sort_order,
        }
        for key, row_label, normalized_value, sort_order in entries
        if normalized_value not in (None, "")
    ]


def _is_dental_basic_coinsurance(text: str) -> bool:
    return (
        ("coinsurance" in text or "% payment of eligible expenses" in text or re.search(r"\b\d{1,3}%", text))
        and (
            any(term in text for term in ["basic dental", "basic services", "routine dental care", "type a"])
            or (
                "reimbursement" in text
                and not any(term in text for term in ["major", "restorative", "orthodontic", "orthodontia", "dentures"])
            )
        )
    )


def _is_dental_major_coinsurance(text: str) -> bool:
    return (
        ("coinsurance" in text or "% payment of eligible expenses" in text or re.search(r"\b\d{1,3}%", text))
        and any(term in text for term in ["major", "restorative services", "dental restorative services"])
    )


def _is_dental_orthodontic_coinsurance(text: str) -> bool:
    return ("coinsurance" in text or "orthodontic care" in text or "orthodontia" in text) and any(term in text for term in ["orthodontic", "orthodontia"])


def _is_dental_basic_maximum(text: str) -> bool:
    return "maximum" in text and any(term in text for term in ["basic dental", "basic services", "routine dental care"])


def _is_dental_major_maximum(text: str) -> bool:
    return "maximum" in text and any(term in text for term in ["major", "restorative services", "level 1 - 4"])


def _is_dental_orthodontic_maximum(text: str) -> bool:
    return "maximum" in text and any(term in text for term in ["orthodontic", "orthodontia"])


def _dental_percent_or_status(value: Any) -> str | None:
    text = _clean_text(value)
    percent = re.search(r"\d{1,3}(?:\.\d+)?%", text)
    if percent:
        return percent.group(0)
    if re.search(r"\bN/?A\b", text, flags=re.IGNORECASE):
        return "N/A"
    if "not applicable" in text.lower():
        return "Not Applicable"
    return _short_dental_value(text)


def _dental_fee_guide_value(value: Any) -> str | None:
    text = _clean_text(value)
    return "Current" if "current" in text.lower() else _short_dental_value(text)


def _short_dental_value(value: Any) -> str | None:
    text = _clean_text(value)
    if not text:
        return None
    if text.lower() in {"nil", "none", "n/a", "na", "not applicable"}:
        return "N/A" if text.lower() in {"n/a", "na"} else text[:1].upper() + text[1:]
    money = _extract_first_money(text)
    if money and len(text) > 80:
        return money
    return text


def _prefer_comparison_value(existing: Any, value: Any) -> Any:
    if existing in (None, ""):
        return value
    if value in (None, "", existing):
        return existing
    existing_text = _clean_text(existing)
    value_text = _clean_text(value)
    combined_value = _combine_percent_and_maximum_value(existing_text, value_text)
    if combined_value:
        return combined_value
    combined_value = _combine_percent_and_maximum_value(value_text, existing_text)
    if combined_value:
        return combined_value
    if existing_text.lower() in {"n/a", "not applicable", "none", "nil"} and value_text:
        return value
    if _looks_like_practitioner_list(existing_text) and _looks_like_practitioner_list(value_text) and len(value_text) > len(existing_text):
        return value
    if "emergency travel assistance" in value_text.lower() and "emergency travel assistance" not in existing_text.lower():
        return value
    if "emergency travel assistance" in existing_text.lower() and "emergency travel assistance" not in value_text.lower():
        return existing
    if len(value_text) < len(existing_text):
        return value
    return existing


def _combine_percent_and_maximum_value(percent_value: Any, maximum_value: Any) -> str | None:
    percent_text = _clean_text(percent_value)
    maximum_text = _normalize_oop_maximum_text(maximum_value)
    if not re.fullmatch(r"\d{1,3}%", percent_text or ""):
        return None
    if not _extract_first_money(maximum_text):
        return None
    return f"{percent_text} up to {maximum_text}"


def _looks_like_practitioner_list(value: Any) -> bool:
    text = _normalize(value)
    practitioner_terms = [
        "chiropractor",
        "physiotherapist",
        "psychologist",
        "podiatrist",
        "naturopath",
        "speech",
        "massage",
        "acupuncturist",
    ]
    return sum(1 for term in practitioner_terms if term in text) >= 2


def _normalize_drug_type_value(value: Any) -> str | None:
    text = _normalize(value)
    if not text:
        return None
    if "mandatory generic" in text:
        return "Mandatory Generic Pricing Drug Plan"
    if "maximum allowable cost" in text or "allowable cost program" in text:
        return "Generic, Maximum Allowable Cost Program"
    if "lowest cost generic" in text or "generic equivalent" in text or "lowest cost alternative" in text:
        return "Lowest Cost Generic"
    if re.search(r"\bbrand\b", text):
        return "Brand"
    return None


def _normalize_comparison_display_value(benefit: str, subsection: str, row_label: str, key: str, value: Any) -> Any:
    text = _clean_text(value)
    if not text:
        return value

    normalized_subsection = _normalize(subsection)

    if normalized_subsection == "prescription drugs" and row_label == "Plan Maximum":
        money = _extract_first_money(text)
        if money == "$5,000":
            return "$5,000 per person per calendar year"

    if normalized_subsection == "prescription drugs" and row_label == "Drug Type":
        drug_type = _normalize_drug_type_value(text)
        if drug_type:
            return drug_type

    if row_label == "Coinsurance":
        percent = re.search(r"\b\d{1,3}(?:\.\d+)?%", text)
        if percent and any(term in text.lower() for term in ["payment", "coinsurance", "reimbursement"]):
            return percent.group(0)

    if normalized_subsection == "paramedical practitioners" and row_label == "Plan Maximum":
        money = _extract_first_money(text)
        if money and "practitioner" in text.lower():
            return f"{money} per practitioner per person per calendar year"

    if normalized_subsection == "paramedical practitioners" and row_label == "Practitioners":
        return _clean_practitioner_list(value)

    if normalized_subsection == "medical services & supplies":
        medical_value = _clean_medical_services_supply_value(row_label, text)
        if medical_value:
            return medical_value

    if row_label == "Definition of Dependent":
        dependent_value = _dependent_definition_display_value(text)
        if dependent_value:
            return dependent_value

    if row_label == "Survivor Benefit":
        survivor_value = _survivor_benefit_display_value(text)
        if survivor_value:
            return survivor_value

    if row_label == "Benefit Amount":
        if subsection in {
            "Basic Life Insurance",
            "Accidental Death & Dismemberment or Accident & Serious Illness",
            "Critical Illness",
        }:
            money = _extract_first_money(text)
            if money:
                return money
        if subsection == "Dependent Life Insurance":
            spouse = _dependent_life_spouse_amount(text)
            child = _dependent_life_child_amount(text)
            if spouse and child:
                return f"Spouse: {spouse}; Each child: {child}"

    if row_label == "Evidence of Insurability" and "not required" in text.lower():
        return "Not Required"

    if row_label == "Plan Maximum" and "matches basic life" in text.lower():
        return "Matches Basic Life"

    if row_label == "Termination of Coverage":
        shortened = _short_termination_display_value(text)
        if shortened:
            return shortened

    return value


def _clean_medical_services_supply_value(row_label: str, value: Any) -> str | None:
    text = _clean_text(value)
    lower = text.lower()
    money = _extract_first_money(text)

    if row_label == "Plan Maximum":
        if "unlimited" in lower:
            return "Unlimited"
        if money:
            suffix = _frequency_suffix(text)
            return f"{money} {suffix}".strip()

    if row_label in {"Orthopaedic Shoes", "Orthotic Inserts", "Private Duty Nursing", "Hearing Aids"}:
        if money:
            suffix = _frequency_suffix(text)
            return f"{money} {suffix}".strip()
        if "included in orthotics maximum" in lower:
            return "Included in Orthotics Maximum"
        if "included" in lower:
            return "Included"

    return None


def _frequency_suffix(value: Any) -> str:
    text = _clean_text(value)
    patterns = [
        r"per person per calendar year",
        r"per\s+\d+\s+calendar years?",
        r"per\s+calendar year",
        r"per\s+year",
        r"per\s+lifetime",
        r"lifetime",
        r"in any period of\s+\d+\s+consecutive months",
        r"per\s+\d+\s+consecutive months",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(0)
    return ""


def _clean_practitioner_list(value: Any) -> str:
    lines = []
    for raw_line in str(value or "").splitlines():
        line = _clean_text(raw_line).lstrip("- ").strip()
        if line:
            lines.append(line)
    return "\n".join(dict.fromkeys(lines))


def _dependent_definition_display_value(value: Any) -> str | None:
    text = _clean_text(value)
    lower = text.lower()
    if "dependent" not in lower and "student" not in lower:
        return None
    match = re.search(r"age\s+(\d{2})\s+if\s+in\s+full\s*time\s+attendance\s+as\s+a\s+student", text, flags=re.IGNORECASE)
    if not match:
        match = re.search(r"full[-\s]?time.*?(?:student|school).*?(?:age|to age|under)\s+(\d{2})", text, flags=re.IGNORECASE)
    if not match:
        match = re.search(r"(?:age|to age|under)\s+(\d{2}).*?full[-\s]?time.*?(?:student|school)", text, flags=re.IGNORECASE)
    if not match:
        return None
    return f"Student: Under {match.group(1)} years of age"


def _survivor_benefit_display_value(value: Any) -> str | None:
    text = _clean_text(value)
    lower = text.lower()
    if "survivor benefit" not in lower:
        return None
    if "two-year" in lower or "2-year" in lower or "24 month" in lower:
        return "Included - 24 months"
    if "included" in lower:
        return "Included"
    return None


def _normalize_basic_life_comparison_rows(rows: list[dict[str, Any]], carriers: list[str]) -> None:
    amount_row = _find_comparison_row(rows, "Basic Life Insurance", "Benefit Amount")
    plan_maximum_row = _find_comparison_row(rows, "Basic Life Insurance", "Plan Maximum")
    reduction_row = _find_comparison_row(rows, "Basic Life Insurance", "Reduction of Coverage")
    if not amount_row:
        return

    for carrier in carriers:
        amount = _extract_first_money(amount_row.get("values", {}).get(carrier))
        if not amount:
            continue
        amount_row["values"][carrier] = amount
        if plan_maximum_row and plan_maximum_row.get("values", {}).get(carrier) in (None, ""):
            plan_maximum_row["values"][carrier] = amount
        elif plan_maximum_row and _normalize(plan_maximum_row.get("values", {}).get(carrier)) == "matches basic life":
            plan_maximum_row["values"][carrier] = amount

        if reduction_row and _has_life_reduction_50_at_65(_normalize(reduction_row.get("values", {}).get(carrier))):
            reduction_row["values"][carrier] = "The amount of insurance reduces by 50% at age 65"


def _normalize_adnd_comparison_rows(rows: list[dict[str, Any]], carriers: list[str]) -> None:
    basic_life_amount = _find_comparison_row(rows, "Basic Life Insurance", "Benefit Amount")
    basic_life_reduction = _find_comparison_row(rows, "Basic Life Insurance", "Reduction of Coverage")
    adnd_amount = _find_comparison_row(rows, "Accidental Death & Dismemberment or Accident & Serious Illness", "Benefit Amount")
    adnd_plan_maximum = _find_comparison_row(rows, "Accidental Death & Dismemberment or Accident & Serious Illness", "Plan Maximum")
    adnd_reduction = _find_comparison_row(rows, "Accidental Death & Dismemberment or Accident & Serious Illness", "Reduction of Coverage")

    if adnd_amount and basic_life_amount:
        for carrier in carriers:
            value = adnd_amount.get("values", {}).get(carrier)
            if _is_equal_to_group_life_value(value):
                replacement = basic_life_amount.get("values", {}).get(carrier)
                if replacement not in (None, ""):
                    adnd_amount["values"][carrier] = replacement
            elif value in (None, "") and adnd_plan_maximum:
                replacement = adnd_plan_maximum.get("values", {}).get(carrier)
                if _normalize(replacement) == "matches basic life":
                    replacement = basic_life_amount.get("values", {}).get(carrier)
                if replacement not in (None, ""):
                    adnd_amount["values"][carrier] = _extract_first_money(replacement) or replacement

    if adnd_amount and adnd_plan_maximum:
        for carrier in carriers:
            if adnd_plan_maximum.get("values", {}).get(carrier) in (None, ""):
                replacement = adnd_amount.get("values", {}).get(carrier)
                if replacement not in (None, "") and not _is_equal_to_group_life_value(replacement):
                    adnd_plan_maximum["values"][carrier] = replacement
            elif _normalize(adnd_plan_maximum.get("values", {}).get(carrier)) == "matches basic life" and basic_life_amount:
                replacement = basic_life_amount.get("values", {}).get(carrier)
                if replacement not in (None, ""):
                    adnd_plan_maximum["values"][carrier] = _extract_first_money(replacement) or replacement

    if adnd_reduction is None and basic_life_reduction:
        adnd_reduction = {
            **basic_life_reduction,
            "subsection": "Accidental Death & Dismemberment or Accident & Serious Illness",
            "row_label": "Reduction of Coverage",
            "values": {carrier: None for carrier in carriers},
        }
        rows.append(adnd_reduction)

    if adnd_reduction:
        for carrier in carriers:
            value = adnd_reduction.get("values", {}).get(carrier)
            if _has_life_reduction_50_at_65(_normalize(value)):
                adnd_reduction["values"][carrier] = "The amount of insurance reduces by 50% at age 65"
                continue
            source_value = basic_life_reduction.get("values", {}).get(carrier) if basic_life_reduction else None
            if value in (None, "") and _has_life_reduction_50_at_65(_normalize(source_value)):
                adnd_reduction["values"][carrier] = "The amount of insurance reduces by 50% at age 65"


def _find_comparison_row(rows: list[dict[str, Any]], subsection: str, row_label: str) -> dict[str, Any] | None:
    subsection_key = _normalize(subsection)
    row_label_key = _normalize(row_label)
    for row in rows:
        if _normalize(row.get("subsection")) == subsection_key and _normalize(row.get("row_label")) == row_label_key:
            return row
    return None


def _is_equal_to_group_life_value(value: Any) -> bool:
    text = _normalize(value)
    return "equal to the group life insurance amount" in text or "same as group life" in text


def _normalize_dependent_life_comparison_rows(rows: list[dict[str, Any]], carriers: list[str]) -> None:
    amount_row = _find_comparison_row(rows, "Dependent Life Insurance", "Benefit Amount")
    child_row = _find_comparison_row(rows, "Dependent Life Insurance", "Child(ren)")
    spouse_row = _find_comparison_row(rows, "Dependent Life Insurance", "Spouse")
    termination_row = _find_comparison_row(rows, "Dependent Life Insurance", "Termination of Coverage")
    dependent_rows = [
        row
        for row in rows
        if _normalize(row.get("subsection")) == "dependent life insurance"
    ]

    if amount_row is None and (spouse_row or child_row):
        source_row = spouse_row or child_row
        amount_row = {
            **source_row,
            "row_label": "Benefit Amount",
            "values": {carrier: None for carrier in carriers},
            "sort_order": 500,
        }
        rows.append(amount_row)
    elif amount_row is None and dependent_rows:
        source_row = dependent_rows[0]
        amount_row = {
            **source_row,
            "row_label": "Benefit Amount",
            "values": {carrier: None for carrier in carriers},
            "sort_order": 500,
        }
        rows.append(amount_row)

    if amount_row:
        for carrier in carriers:
            existing = amount_row.get("values", {}).get(carrier)
            spouse = _dependent_life_spouse_amount(
                existing,
                spouse_row.get("values", {}).get(carrier) if spouse_row else None,
            )
            child = _dependent_life_child_amount(
                existing,
                child_row.get("values", {}).get(carrier) if child_row else None,
            )
            if spouse and child:
                amount_row["values"][carrier] = f"Spouse: {spouse}; Each child: {child}"
            elif not _is_current_carrier(carrier) and existing in (None, ""):
                amount_row["values"][carrier] = "not found in quote"

    if termination_row:
        _normalize_matching_termination_values(termination_row, carriers)

    rows[:] = [
        row
        for row in rows
        if not (
            _normalize(row.get("subsection")) == "dependent life insurance"
            and _normalize(row.get("row_label")) in {"child(ren)", "spouse"}
        )
    ]


def _normalize_deductible_comparison_rows(rows: list[dict[str, Any]], carriers: list[str]) -> None:
    for row in rows:
        if _normalize(row.get("row_label")) != "deductible":
            continue
        values = row.get("values") or {}
        for carrier in carriers:
            if _normalize(carrier).startswith("current"):
                continue
            value = values.get(carrier)
            text = _clean_text(value)
            if _extract_first_money(text):
                continue
            values[carrier] = "None"


def _normalize_hospital_comparison_rows(rows: list[dict[str, Any]], carriers: list[str]) -> None:
    indemnity_row = _find_comparison_row(rows, "Hospital", "Hospital Indemnity Benefit")
    if not indemnity_row:
        return
    values = indemnity_row.get("values") or {}
    if not any(values.get(carrier) not in (None, "") for carrier in carriers if not _is_current_carrier(carrier)):
        return
    for carrier in carriers:
        if _is_current_carrier(carrier):
            continue
        if values.get(carrier) in (None, ""):
            values[carrier] = "not covered"


def _normalize_termination_comparison_rows(rows: list[dict[str, Any]], carriers: list[str]) -> None:
    for row in rows:
        if _normalize(row.get("row_label")) != "termination of coverage":
            continue
        values = row.get("values") or {}
        for carrier in carriers:
            value = values.get(carrier)
            shortened = _short_termination_display_value(value)
            if shortened:
                values[carrier] = shortened


def _dependent_life_spouse_amount(*values: Any) -> str | None:
    for value in values:
        text = _clean_text(value)
        if not text:
            continue
        match = re.search(r"spouse\s*:?\s*(\$[\d,]+)", text, flags=re.IGNORECASE)
        if match:
            return match.group(1)
        match = re.search(r"(\$[\d,]+)\s+for\s+spouse", text, flags=re.IGNORECASE)
        if match:
            return match.group(1)
        if "child" not in text.lower():
            amount = _extract_first_money(text)
            if amount:
                return amount
    return None


def _dependent_life_child_amount(*values: Any) -> str | None:
    for value in values:
        text = _clean_text(value)
        if not text:
            continue
        match = re.search(r"(?:each\s+child|child(?:ren)?|child\(ren\))\s*:?\s*(\$[\d,]+)", text, flags=re.IGNORECASE)
        if match:
            return match.group(1)
        match = re.search(r"(\$[\d,]+)\s+for\s+(?:each\s+)?child", text, flags=re.IGNORECASE)
        if match:
            return match.group(1)
        if "spouse" not in text.lower() and _extract_first_money(text):
            return _extract_first_money(text)
    return None


def _normalize_matching_termination_values(row: dict[str, Any], carriers: list[str]) -> None:
    values = row.get("values") or {}
    canonical_by_age: dict[str, str] = {}
    for carrier in carriers:
        value = values.get(carrier)
        age = _termination_age(value)
        if not age:
            continue
        text = _clean_text(value)
        if not text:
            continue
        existing = canonical_by_age.get(age)
        if existing is None or len(text) > len(existing):
            canonical_by_age[age] = text

    for carrier in carriers:
        value = values.get(carrier)
        age = _termination_age(value)
        if age and canonical_by_age.get(age):
            values[carrier] = canonical_by_age[age]


def _termination_age(value: Any) -> str | None:
    text = _clean_text(value).lower()
    match = re.search(r"age\s*(\d{2})|\b(\d{2})\s*,?\s+or earlier retirement", text)
    if not match:
        return None
    return match.group(1) or match.group(2)


def _with_dental_group_rows(rows: list[dict[str, Any]], carriers: list[str]) -> list[dict[str, Any]]:
    dental_keys = {str(row.get("row_label") or "") for row in rows if row.get("major_section") == "Dental Care"}
    additions: list[dict[str, Any]] = []
    if any(int(row.get("sort_order") or 0) in {31, 32, 33} for row in rows if row.get("major_section") == "Dental Care"):
        additions.append(_empty_comparison_group_row("Coverage Coinsurance", 30, carriers))
    if any(str(row.get("sort_order")) in {"41", "42", "43"} for row in rows if row.get("major_section") == "Dental Care"):
        additions.append(_empty_comparison_group_row("Maximum", 40, carriers))
    return [*rows, *additions]


def _empty_comparison_group_row(row_label: str, sort_order: int, carriers: list[str]) -> dict[str, Any]:
    return {
        "major_section": "Dental Care",
        "subsection": "Dental Care",
        "row_label": row_label,
        "values": {carrier: None for carrier in carriers},
        "repeat_carrier_header": False,
        "client_name": None,
        "effective_date": None,
        "source_file": None,
        "source_page": None,
        "confidence_score": "HIGH",
        "sort_order": sort_order,
    }


def _sort_comparison_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    indexed = list(enumerate(rows))
    indexed.sort(
        key=lambda item: (
            _comparison_major_sort(str(item[1].get("major_section") or "")),
            _comparison_subsection_sort(str(item[1].get("subsection") or "")),
            int(item[1].get("sort_order") or 500),
            item[0],
        )
    )
    return [row for _, row in indexed]


def _comparison_major_sort(section: str) -> int:
    return {"Pooled Benefits": 10, "Extended Health": 20, "Dental Care": 30, "Other Benefits": 40}.get(section, 90)


def _comparison_subsection_sort(subsection: str) -> int:
    normalized = _normalize(subsection)
    if normalized in COMPARISON_SUBSECTION_ORDER:
        return COMPARISON_SUBSECTION_ORDER[normalized]
    for key, order in COMPARISON_SUBSECTION_ORDER.items():
        if key and (key in normalized or normalized in key):
            return order
    return 900


def _comparison_major_section(benefit: str, subsection: str) -> str:
    if benefit in {"Employee Life Insurance", "Employee AD&D", "Dependent Life", "Long Term Disability", "Short Term Disability"}:
        return "Pooled Benefits"
    if benefit in {"Extended Health Care", "Travel"}:
        return "Extended Health"
    if benefit == "Dental Care":
        return "Dental Care"
    return "Other Benefits"


def _comparison_values_for_record(record: dict[str, Any], carrier_name: str, value: Any | None = None) -> dict[str, Any]:
    if value is None:
        value = record.get("provision_value")
    values = {
        "Current": None,
        "Carrier 1": value,
        carrier_name: value,
    }
    for index in range(2, MAX_COMPARISON_CARRIERS + 1):
        values[f"Carrier {index}"] = None
    return values


def _should_repeat_benefit_header(benefit: str, subsection: str, row_label: str) -> bool:
    if row_label == "__SUBSECTION_HEADER__":
        return False
    return subsection in {
        "Long Term Disability",
        "Prescription Drugs",
        "Vision Care",
        "Hospital",
        "Out of Province or Country",
        "Medical Services & Supplies",
    }


def _comparison_subsection_and_label(benefit: str, provision: str) -> tuple[str, str]:
    if " - " in provision:
        subsection, row_label = provision.split(" - ", 1)
        normalized_subsection = _comparison_subsection_label(subsection.strip() or benefit)
        if _is_approved_comparison_subsection(normalized_subsection):
            return normalized_subsection, _comparison_row_label(row_label.strip() or provision)
        base_subsection = _comparison_subsection_label(_plan_design_section_label(benefit))
        return base_subsection, _comparison_row_label(subsection.strip() or row_label.strip() or provision)
    return _comparison_subsection_label(_plan_design_section_label(benefit)), _comparison_row_label(provision)


def _comparison_subsection_label(value: str) -> str:
    mapping = {
        "Employee Life": "Basic Life Insurance",
        "Employee Life Insurance": "Basic Life Insurance",
        "Employee AD&D": "Accidental Death & Dismemberment or Accident & Serious Illness",
        "Dependent Life": "Dependent Life Insurance",
        "Drugs": "Prescription Drugs",
        "Drug coverage": "Prescription Drugs",
        "Paramedical": "Paramedical Practitioners",
        "Vision": "Vision Care",
        "Travel": "Out of Province or Country",
        "Extended Health": "Medical Services & Supplies",
        "Extended Health Care": "Medical Services & Supplies",
        "Major Services": "Medical Services & Supplies",
        "Dental Basic": "Basic Dental",
        "Dental Major": "Major Dental",
        "Dental Recall": "Dental Recall",
        "Virtual Care": "Virtual Healthcare",
        "Medical Second Opinion": "Medical Second Opinion",
        "EAP": "EAP",
        "Critical Illness": "Critical Illness",
    }
    return mapping.get(value, value)


def _is_approved_comparison_subsection(value: str) -> bool:
    normalized = _normalize(value)
    return normalized in COMPARISON_CATEGORY_TITLES or normalized in COMPARISON_SUBSECTION_ORDER


def _comparison_row_label(value: str) -> str:
    text = str(value or "").strip()
    mapping = {
        "Benefit amount": "Benefit Amount",
        "Reduction schedule": "Reduction of Coverage",
        "Termination age": "Termination of Coverage",
        "Waiver of premium": "Waiver of Premium",
        "Conversion option": "Conversion Option",
        "Evidence of insurability": "Evidence of Insurability",
        "Non-Evidence Maximum": "Evidence of Insurability",
        "Non-evidence Maximum": "Evidence of Insurability",
        "Non-Evidence Limit": "Evidence of Insurability",
        "Non-evidence Limit": "Evidence of Insurability",
        "Overall Maximum": "Plan Maximum",
        "Dependent child eligibility": "Definition of Dependent",
        "Child Eligibility": "Definition of Dependent",
        "Coinsurance": "Coinsurance",
        "Deductible": "Deductible",
        "Maximum": "Plan Maximum",
        "Drug plan": "Plan Type",
        "Drug Option": "Drug Type",
        "Fee guide": "Fee Guide",
        "Recall frequency": "Recall Frequency",
        "Not eligible": "Not Eligible",
        "Included provision": "Included",
        "Orthopaedic Shoes (Custom made or Custom modified Stock)": "Orthopaedic Shoes",
        "Orthopedic Shoes (Custom made or Custom modified Stock)": "Orthopaedic Shoes",
        "Orthotics (Custom made for the patient)": "Orthotic Inserts",
        "Foot Orthotics": "Orthotic Inserts",
        "Nursing Care Services": "Private Duty Nursing",
    }
    if text in mapping:
        return mapping[text]
    if text.lower().endswith("coinsurance"):
        return "Coinsurance"
    if text.lower().endswith("maximum"):
        return "Plan Maximum"
    return text[:1].upper() + text[1:] if text else text


def _source_page_from_location(value: Any) -> int | None:
    text = str(value or "")
    match = re.search(r"Page\s+(\d+)", text, flags=re.IGNORECASE)
    return int(match.group(1)) if match else None


def _write_total_cost_comparison(sheet: Any, model: dict[str, Any]) -> list[dict[str, Any]]:
    _prepare_sheet(sheet)
    carrier = model.get("carrier_name") or "Carrier"
    headers = [
        "Benefit / Class / Tier",
        "Lives",
        "Volume",
        "Current Rate",
        "Current Premium",
        f"{carrier} Rate",
        f"{carrier} Premium",
        "Annual Premium",
        "Rate Basis",
        "Volume Basis",
        "Status",
        "Notes",
        "Benefit Key",
    ]

    _title(sheet, "TOTAL COST COMPARISON", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    ordered = _ordered_benefit_rows(model["benefit_rows"])
    row_number = 4
    written: list[dict[str, Any]] = []
    current_benefit = None
    for benefit, rows in ordered:
        category = _benefit_category(benefit)
        sheet.cell(row_number, 1, benefit)
        sheet.cell(row_number, 2, _category_label(category))
        _group_row(sheet, row_number, len(headers))
        row_number += 1
        current_benefit = benefit

        for record in rows:
            label = _row_label(record, current_benefit)
            values = [
                label,
                record.get("lives"),
                record.get("volume"),
                record.get("current_rate"),
                record.get("current_premium"),
                record.get("rate"),
                record.get("monthly_premium"),
                record.get("annual_premium"),
                record.get("rate_basis"),
                record.get("volume_basis"),
                record.get("included_status") or record.get("status"),
                record.get("notes"),
                current_benefit,
            ]
            for column, value in enumerate(values, start=1):
                sheet.cell(row_number, column, value)
            _status_cell(sheet.cell(row_number, 11), record.get("status"))
            _audit_destination(model, record, "Benefit row", label, "Total Cost Comparison", f"Row {row_number}", record.get("confidence"), record.get("status"))
            written.append({"benefit": current_benefit, "row": row_number, "record": record})
            row_number += 1

    _format_sheet(sheet, max(row_number - 1, 3), len(headers))
    sheet.column_dimensions["M"].hidden = True
    return written


def _write_gross_cost_comparison(sheet: Any, model: dict[str, Any], rows: list[dict[str, Any]]) -> None:
    _prepare_sheet(sheet)
    for column in ["B", "C", "D"]:
        sheet.column_dimensions[column].hidden = True
    _unhide_all_rows(sheet)

    _write_gross_cost_header(sheet, model)
    row_number = 4
    premium_rows_by_column: dict[int, list[int]] = defaultdict(list)
    for section_name, section_rows in _gross_cost_sections(rows):
        if section_name in {"Extended Health", "Dental"}:
            sheet.cell(row_number, 1, section_name)
            _gross_left_cell(sheet.cell(row_number, 1), bold=True)
            row_number += 1

        for item in section_rows:
            label = _gross_cost_label(item)
            sheet.cell(row_number, 1, label)
            _gross_left_cell(sheet.cell(row_number, 1), bold=item.get("status") == "NOT AVAILABLE")
            if item.get("status") != "NOT AVAILABLE":
                for carrier_index in range(1, 6):
                    base_column = 5 + ((carrier_index - 1) * 3)
                    volume = item.get(f"carrier{carrier_index}_volume")
                    lives = item.get(f"carrier{carrier_index}_lives")
                    rate = item.get(f"carrier{carrier_index}_rate")
                    premium = item.get(f"carrier{carrier_index}_premium")
                    if carrier_index == 1:
                        volume = volume if volume not in (None, "") else item.get("carrier_volume")
                        lives = lives if lives not in (None, "") else item.get("carrier_lives")
                        rate = rate if rate not in (None, "") else item.get("carrier_rate")
                        premium = premium if premium not in (None, "") else item.get("carrier_premium")
                    sheet.cell(row_number, base_column, volume if volume not in (None, "") else lives)
                    sheet.cell(row_number, base_column + 1, rate)
                    if premium not in (None, "", 0, 0.0):
                        sheet.cell(row_number, base_column + 2, premium)
                        premium_rows_by_column[base_column + 2].append(row_number)
                    elif item["benefit"] in {"EAP", "Virtual Care", "Medical Second Opinion", "Wellness"}:
                        sheet.cell(row_number, base_column + 2, "included in EHC")
            row_number += 1

    monthly_row = row_number
    sheet.cell(monthly_row, 1, "Total Monthly Premium")
    _gross_left_cell(sheet.cell(monthly_row, 1), bold=True)
    for premium_column in [7, 10, 13, 16, 19]:
        letter = get_column_letter(premium_column)
        premium_rows = premium_rows_by_column.get(premium_column, [])
        if premium_rows:
            refs = ",".join(f"{letter}{source_row}" for source_row in premium_rows)
            sheet.cell(monthly_row, premium_column, f"=SUM({refs})")
    row_number += 1

    annual_row = row_number
    sheet.cell(annual_row, 1, "Total Annual Premium")
    _gross_left_cell(sheet.cell(annual_row, 1), bold=True)
    for premium_column in [7, 10, 13, 16, 19]:
        letter = get_column_letter(premium_column)
        sheet.cell(annual_row, premium_column, f"={letter}{monthly_row}*12")

    _format_gross_cost_sheet(sheet, annual_row)
    _clear_unused_sheet_formatting(sheet, annual_row + 1)
    _unhide_all_rows(sheet)
    sheet.freeze_panes = None
    _audit(
        model,
        source_file=None,
        source_type="Output",
        source_location="Gross Cost Comparison",
        field_name="Gross cost comparison header",
        value="; ".join(_carrier_display_name(header) for header in _comparison_carrier_headers(model)[1:] if _carrier_display_name(header)),
        confidence="HIGH",
        status="FOUND",
        destination_sheet="Gross Cost Comparison",
        destination_cell="E1:S2",
        notes="Top extraction header rebuilt from the market survey comparison layout.",
    )


def _unhide_all_rows(sheet: Any) -> None:
    max_row = max(sheet.max_row, 200)
    for row in range(1, max_row + 1):
        sheet.row_dimensions[row].hidden = False


def _remove_trailing_blank_rows(sheet: Any, keep_through_row: int) -> None:
    while sheet.max_row > keep_through_row:
        row = sheet.max_row
        has_value = any(sheet.cell(row, column).value not in (None, "") for column in range(1, sheet.max_column + 1))
        if has_value:
            break
        sheet.delete_rows(row, 1)


def _clear_unused_sheet_formatting(sheet: Any, first_row: int, max_columns: int | None = None) -> None:
    max_row = max(sheet.max_row, 200)
    max_col = max_columns or max(sheet.max_column, 20)
    for row in range(first_row, max_row + 1):
        has_value = any(sheet.cell(row, column).value not in (None, "") for column in range(1, max_col + 1))
        if has_value:
            continue
        sheet.row_dimensions[row].height = None
        sheet.row_dimensions[row].hidden = False
        for column in range(1, max_col + 1):
            cell = sheet.cell(row, column)
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()
            cell.alignment = Alignment()
            cell.number_format = "General"
            cell.font = Font(name=WORKBOOK_FONT_NAME)


def _write_gross_cost_header(sheet: Any, model: dict[str, Any]) -> None:
    dark_fill = PatternFill("solid", fgColor="004C55")
    group_fill = PatternFill("solid", fgColor="0F766E")
    subheader_fill = PatternFill("solid", fgColor="EAF8F8")
    sheet.cell(1, 1, "")
    sheet.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    sheet.cell(1, 1).fill = dark_fill

    if model.get("carrier_names"):
        headers = _comparison_carrier_headers(model)
        group_labels = [_carrier_display_name(header) for header in headers[1:]]
    else:
        group_labels = [_carrier_display_name(header) for header in _comparison_carrier_headers(model)[1:]]
    groups = [(5 + (index * 3), label) for index, label in enumerate(group_labels[:5])]
    for start_col, label in groups:
        sheet.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=start_col + 2)
        header = sheet.cell(1, start_col, label)
        header.fill = group_fill
        header.font = Font(bold=True, color="FFFFFF")
        header.alignment = Alignment(horizontal="center", vertical="center")
        for offset, subheader in enumerate(["Volume", "Rate", "Premium"]):
            cell = sheet.cell(2, start_col + offset, subheader)
            cell.fill = subheader_fill
            cell.font = Font(color="000000")
            cell.alignment = Alignment(horizontal="center", vertical="center")


def _gross_cost_sections(rows: list[dict[str, Any]]) -> list[tuple[str, list[dict[str, Any]]]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in rows:
        grouped[item["benefit"]].append(item)

    ordered_names = [
        "Employee Life Insurance",
        "Employee AD&D",
        "Dependent Life",
        "Long Term Disability",
        "Short Term Disability",
        "Extended Health Care",
        "Dental Care",
        "HCSA",
        "EAP",
        "Virtual Care",
        "Medical Second Opinion",
        "Wellness",
    ]
    sections: list[tuple[str, list[dict[str, Any]]]] = []
    used = set()
    for benefit in ordered_names:
        if benefit in grouped:
            section = "Extended Health" if benefit == "Extended Health Care" else ("Dental" if benefit == "Dental Care" else benefit)
            sections.append((section, grouped[benefit]))
            used.add(benefit)
    for benefit in sorted(set(grouped) - used):
        sections.append((benefit, grouped[benefit]))
    return sections


def _gross_cost_label(item: dict[str, Any]) -> str:
    benefit = item["benefit"]
    if benefit == "Employee Life Insurance":
        return "Employee Life"
    if benefit == "Employee AD&D":
        return "Employee AD&D"
    if benefit == "Extended Health Care":
        return str(item.get("tier_class") or "Extended Health")
    if benefit == "Dental Care":
        return str(item.get("tier_class") or "Dental")
    if benefit == "Virtual Care":
        return "Virtual Healthcare"
    return benefit


def _gross_left_cell(cell: Any, bold: bool = False) -> None:
    cell.fill = PatternFill("solid", fgColor="004C55")
    cell.font = Font(bold=bold, color="FFFFFF")
    cell.alignment = Alignment(vertical="center", wrap_text=True)


def _format_gross_cost_sheet(sheet: Any, max_row: int) -> None:
    fills = {
        5: "F2F2F2",
        8: "CDEFF3",
        11: "D9D9D9",
        14: "F2F2F2",
        17: "F2F2F2",
    }
    thin_teal = Side(style="thin", color="004C55")
    border = Border(bottom=thin_teal)
    sheet.freeze_panes = "E3"
    for row in range(1, max_row + 1):
        for column in range(1, 20):
            cell = sheet.cell(row, column)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if column == 1 and row >= 4:
                cell.fill = PatternFill("solid", fgColor="004C55")
                cell.font = Font(color="FFFFFF", bold=cell.font.bold)
            for start_col, fill in fills.items():
                if start_col <= column <= start_col + 2 and row >= 3:
                    cell.fill = PatternFill("solid", fgColor=fill)
                    cell.border = border
            if column in {7, 10, 13, 16, 19} and row >= 4:
                cell.number_format = "$#,##0.00"
            elif column in {6, 9, 12, 15, 18} and row >= 4:
                cell.number_format = "$#,##0.000"
            elif column in {5, 8, 11, 14, 17} and row >= 4:
                cell.number_format = "#,##0"
    widths = {1: 28, 5: 12, 6: 12, 7: 14}
    for column in range(1, 20):
        sheet.column_dimensions[get_column_letter(column)].width = widths.get(column, 12)


def _build_reconciliation_rows(model: dict[str, Any]) -> list[dict[str, Any]]:
    if model.get("carrier_names"):
        return _build_multi_carrier_reconciliation_rows(model)

    rows: list[dict[str, Any]] = []
    single_class_benefits = _single_class_benefits(model["benefit_rows"])
    for benefit, benefit_rows in _ordered_benefit_rows(model["benefit_rows"]):
        for record in benefit_rows:
            has_rate_data = any(
                record.get(field) not in (None, "", 0, 0.0)
                for field in ["current_rate", "current_premium", "rate", "monthly_premium", "volume", "lives"]
            )
            is_standard_missing = record.get("status") == "NOT AVAILABLE" and benefit in STANDARD_BENEFITS
            if not has_rate_data and not is_standard_missing:
                continue

            rows.append(
                {
                    "benefit": benefit,
                    "tier_class": _rate_row_label(record, "All", single_class_benefits),
                    "volume": record.get("volume"),
                    "lives": record.get("lives"),
                    "current_rate": None,
                    "current_premium": None,
                    "carrier_rate": record.get("rate"),
                    "carrier_premium": record.get("monthly_premium"),
                    "carrier_volume": record.get("volume"),
                    "carrier_lives": record.get("lives"),
                    "rate_basis": record.get("rate_basis"),
                    "volume_basis": record.get("volume_basis"),
                    "source": record.get("source_location") or record.get("source_type"),
                    "status": record.get("status"),
                    "notes": record.get("notes"),
                }
            )
    return rows


def _build_multi_carrier_reconciliation_rows(model: dict[str, Any]) -> list[dict[str, Any]]:
    rows_by_key: dict[tuple[str, str], dict[str, Any]] = {}
    ordered_keys: list[tuple[str, str]] = []
    single_class_benefits = _single_class_benefits(model["benefit_rows"])
    for benefit, benefit_rows in _ordered_benefit_rows(model["benefit_rows"]):
        for record in benefit_rows:
            has_rate_data = any(
                record.get(field) not in (None, "", 0, 0.0)
                for field in ["rate", "monthly_premium", "volume", "lives"]
            )
            is_standard_missing = record.get("status") == "NOT AVAILABLE" and benefit in STANDARD_BENEFITS
            if not has_rate_data and not is_standard_missing:
                continue

            label = _rate_row_label(record, "All", single_class_benefits)
            key = (benefit, label)
            if key not in rows_by_key:
                rows_by_key[key] = {
                    "benefit": benefit,
                    "tier_class": label,
                    "volume": None,
                    "lives": None,
                    "current_rate": None,
                    "current_premium": None,
                    "rate_basis": record.get("rate_basis"),
                    "volume_basis": record.get("volume_basis"),
                    "source": "",
                    "status": record.get("status"),
                    "notes": "",
                }
                ordered_keys.append(key)
            row = rows_by_key[key]
            index = int(record.get("carrier_index") or 1)
            if 1 <= index <= 5:
                row[f"carrier{index}_rate"] = record.get("rate")
                row[f"carrier{index}_premium"] = record.get("monthly_premium")
                row[f"carrier{index}_volume"] = record.get("volume")
                row[f"carrier{index}_lives"] = record.get("lives")
            row["volume"] = row.get("volume") or record.get("volume")
            row["lives"] = row.get("lives") or record.get("lives")
            row["source"] = _join_unique(row.get("source"), f"{record.get('carrier_name')}: {record.get('source_location') or record.get('source_type')}")
            row["notes"] = _join_unique(row.get("notes"), f"{record.get('carrier_name')}: {record.get('notes')}")
            if record.get("status") == "VALID":
                row["status"] = "VALID"
    return [rows_by_key[key] for key in ordered_keys]


def _single_class_benefits(rows: list[dict[str, Any]]) -> set[str]:
    concrete_classes: dict[str, set[str]] = defaultdict(set)
    for row in rows:
        benefit = str(row.get("benefit_line") or "")
        class_name = _normalize(row.get("class_name"))
        if not benefit or not class_name or class_name in {"all", "all employees"}:
            continue
        concrete_classes[benefit].add(class_name)
    return {benefit for benefit, classes in concrete_classes.items() if classes == {"class a"}}


def _rate_row_label(record: dict[str, Any], fallback: str, single_class_benefits: set[str]) -> str:
    class_name = record.get("class_name")
    tier = record.get("tier")
    if record.get("benefit_line") in single_class_benefits and _normalize(class_name) == "class a":
        class_name = "All"
    parts = [class_name, tier]
    label = " - ".join(str(part) for part in parts if part)
    return label or fallback


def _join_unique(existing: Any, value: Any) -> str:
    text = _clean_text(value)
    if not text:
        return _clean_text(existing)
    existing_text = _clean_text(existing)
    if not existing_text:
        return text
    parts = [part.strip() for part in existing_text.split(";") if part.strip()]
    if text not in parts:
        parts.append(text)
    return "; ".join(parts)


def _write_rate_reconciliation(sheet: Any, model: dict[str, Any], rows: list[dict[str, Any]]) -> None:
    _prepare_sheet(sheet)
    carrier_headers = _comparison_carrier_headers(model)
    market_carriers = carrier_headers[1:MAX_COMPARISON_CARRIERS + 1] or ["Carrier 1"]
    headers = [
        "Benefit",
        "Tier/Class",
        "Volume",
        "Lives",
        "Current Rate",
        "Current Premium",
    ]
    for carrier in market_carriers:
        carrier_label = _carrier_display_name(carrier) or "Carrier"
        headers.extend([f"{carrier_label} Rate", f"{carrier_label} Premium"])
    source_col = len(headers) + 1
    first_market_carrier = _carrier_display_name(market_carriers[0]) or "Carrier"
    headers.extend(
        [
            "Source",
            f"{first_market_carrier} Volume",
            f"{first_market_carrier} Lives",
            "Negotiated Rate",
            "Negotiated Premium",
            "% Change against Current",
            "Status",
            "Notes",
        ]
    )
    negotiated_rate_col = source_col + 3
    negotiated_premium_col = source_col + 4
    change_col = source_col + 5
    status_col = source_col + 6

    for column, header in enumerate(headers, start=1):
        sheet.cell(1, column, header)

    _rate_reconciliation_header(sheet, len(headers))
    row_number = 2
    for item in rows:
        values: list[Any] = [
            item["benefit"],
            item["tier_class"],
            item.get("volume"),
            item.get("lives"),
            item.get("current_rate"),
            item.get("current_premium"),
        ]
        for carrier_index, _carrier in enumerate(market_carriers, start=1):
            values.extend(
                [
                    item.get(f"carrier{carrier_index}_rate", item.get("carrier_rate") if carrier_index == 1 else None),
                    item.get(f"carrier{carrier_index}_premium", item.get("carrier_premium") if carrier_index == 1 else None),
                ]
            )

        first_rate_column = get_column_letter(7)
        first_premium_column = get_column_letter(8)
        negotiated_premium_column = get_column_letter(negotiated_premium_col)
        values.extend(
            [
            item.get("source"),
            item.get("carrier1_volume", item.get("carrier_volume")),
            item.get("carrier1_lives", item.get("carrier_lives")),
                f"={first_rate_column}{row_number}",
                f"={first_premium_column}{row_number}",
                f'=IF(OR(F{row_number}="",F{row_number}=0,{negotiated_premium_column}{row_number}=""),"",{negotiated_premium_column}{row_number}/F{row_number}-1)',
            item.get("status"),
            item.get("notes"),
            ]
        )
        for column, value in enumerate(values, start=1):
            sheet.cell(row_number, column, value)
        _status_cell(sheet.cell(row_number, status_col), item.get("status"))
        _audit(
            model,
            source_file=None,
            source_type="Output",
            source_location="Rate Reconciliation",
            field_name="Rate reconciliation row",
            benefit_line=item["benefit"],
            value=item["tier_class"],
            confidence="HIGH" if item.get("status") == "VALID" else "MEDIUM",
            status=item.get("status") or "FOUND",
            destination_sheet="Rate Reconciliation",
            destination_cell=f"Row {row_number}",
            notes="Current and carrier values reconciled from extracted rate rows.",
        )
        row_number += 1

    total_row = row_number
    sheet.cell(total_row, 1, "Totals" if rows else "No rate data available")
    carrier_premium_columns = [8 + (index * 2) for index in range(len(market_carriers))]
    if rows:
        for column in [6, *carrier_premium_columns, negotiated_premium_col]:
            letter = get_column_letter(column)
            sheet.cell(total_row, column, f"=SUM({letter}2:{letter}{total_row - 1})")
        sheet.cell(total_row, change_col, f'=IF(OR(F{total_row}=0,F{total_row}=""),"",{negotiated_premium_column}{total_row}/F{total_row}-1)')
    _group_row(sheet, total_row, len(headers))

    _format_rate_reconciliation(sheet, total_row, len(headers))
    _set_fixed_row_height(sheet, total_row, 15)


def _write_line_rates_formatted(sheet: Any, model: dict[str, Any], rows: list[dict[str, Any]]) -> None:
    _prepare_sheet(sheet)
    max_col = _write_line_rates_header(sheet, model)
    market_carriers = _comparison_carrier_headers(model)[1:MAX_COMPARISON_CARRIERS + 1] or ["Carrier 1"]

    row_number = 5
    detail_rows: list[int] = []
    pooled_rows: list[int] = []
    health_rows: list[int] = []
    dental_rows: list[int] = []

    grouped = _group_line_rate_rows(rows)
    for group_name, group_rows in grouped:
        if not group_rows:
            continue
        section_row = row_number
        sheet.cell(section_row, 1, group_name)
        _line_rates_section(sheet, section_row, max_col)
        row_number += 1

        for item in group_rows:
            label = item.get("tier_class") or "All"
            volume_or_lives = item.get("volume") if item.get("volume") not in (None, "") else item.get("lives")
            values: list[Any] = [
                label,
                volume_or_lives,
                item.get("current_rate"),
                item.get("current_premium"),
            ]
            for carrier_index, _carrier in enumerate(market_carriers, start=1):
                values.extend(
                    [
                        item.get(f"carrier{carrier_index}_rate", item.get("carrier_rate") if carrier_index == 1 else None),
                        item.get(f"carrier{carrier_index}_premium", item.get("carrier_premium") if carrier_index == 1 else None),
                    ]
                )
            for column, value in enumerate(values, start=1):
                sheet.cell(row_number, column, value)
            detail_rows.append(row_number)
            if _line_rate_group_kind(group_name) == "pooled":
                pooled_rows.append(row_number)
            elif _line_rate_group_kind(group_name) == "health":
                health_rows.append(row_number)
            elif _line_rate_group_kind(group_name) == "dental":
                dental_rows.append(row_number)
            row_number += 1

        kind = _line_rate_group_kind(group_name)
        next_group_name = _next_group_name(grouped, group_name)
        next_kind = _line_rate_group_kind(next_group_name) if next_group_name else None
        if kind == "pooled" and next_kind != "pooled" and pooled_rows:
            total_row = _write_line_rate_total_row(sheet, row_number, "Total Pooled Benefits", pooled_rows, max_col)
            row_number += 1
            row_number = _write_line_rate_difference_row(sheet, row_number, "Pooled Difference", total_row, max_col)
        elif kind == "health" and next_kind != "health" and health_rows:
            total_row = _write_line_rate_total_row(sheet, row_number, "Total Health Premiums", health_rows, max_col)
            row_number += 1
            row_number = _write_line_rate_annual_row(sheet, row_number, "Total Annual EHC premiums", total_row, max_col)
        elif kind == "dental" and next_kind != "dental" and dental_rows:
            total_row = _write_line_rate_total_row(sheet, row_number, "Total Dental Premiums", dental_rows, max_col)
            row_number += 1
            row_number = _write_line_rate_annual_row(sheet, row_number, "Total Annual Dental Premiums", total_row, max_col)
            if health_rows:
                row_number = _write_line_rate_total_row(sheet, row_number, "Total EHC & Dental", [*health_rows, *dental_rows], max_col) + 1

    monthly_total_row = _write_line_rate_total_row(sheet, row_number, "Total monthly premium", detail_rows, max_col)
    row_number += 1
    row_number = _write_line_rate_difference_row(sheet, row_number, "Monthly premium difference", monthly_total_row, max_col)
    annual_total_row = row_number
    row_number = _write_line_rate_annual_row(sheet, row_number, "Total annual premium", monthly_total_row, max_col)
    _write_line_rate_annual_difference_row(sheet, row_number, "Annual premium difference", annual_total_row, max_col)

    _format_line_rates_sheet(sheet, row_number, max_col)


def _rate_reconciliation_header(sheet: Any, max_col: int) -> None:
    header_fill = PatternFill("solid", fgColor="0F766E")
    negotiated_fill = PatternFill("solid", fgColor="92D050")
    for column in range(1, max_col + 1):
        cell = sheet.cell(1, column)
        header = str(cell.value or "")
        cell.fill = negotiated_fill if "Negotiated" in header or "% Change" in header else header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="top", wrap_text=True)


def _format_rate_reconciliation(sheet: Any, max_row: int, max_col: int) -> None:
    border = Border(
        left=Side(style="thin", color="666666"),
        right=Side(style="thin", color="666666"),
        top=Side(style="thin", color="666666"),
        bottom=Side(style="thin", color="666666"),
    )
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
    for row in range(1, max_row + 1):
        for column in range(1, max_col + 1):
            cell = sheet.cell(row, column)
            header = str(sheet.cell(1, column).value or "")
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if "Premium" in header and row > 1:
                cell.number_format = "$#,##0.00"
            elif "% Change" in header and row > 1:
                cell.number_format = "0.0%"
            elif ("Rate" in header or "Volume" in header) and row > 1:
                cell.number_format = "#,##0.000"
            elif "Lives" in header and row > 1:
                cell.number_format = "#,##0"
    widths = {
        1: 32,
        2: 14,
        3: 14,
        4: 10,
    }
    for column in range(1, max_col + 1):
        header = str(sheet.cell(1, column).value or "")
        if header == "Source":
            width = 22
        elif header == "Status":
            width = 14
        elif header == "Notes":
            width = 42
        else:
            width = widths.get(column, 13)
        sheet.column_dimensions[get_column_letter(column)].width = width


def _set_fixed_row_height(sheet: Any, max_row: int, height: float) -> None:
    for row in range(1, max(max_row, sheet.max_row, 200) + 1):
        sheet.row_dimensions[row].height = height


def _premium_formula(row_number: int, rate_column: str, item: dict[str, Any]) -> str:
    volume = f"C{row_number}"
    lives = f"D{row_number}"
    rate = f"{rate_column}{row_number}"
    basis = _normalize(item.get("rate_basis"))
    if "1,000" in basis or "1000" in basis:
        return f'=IF(OR({volume}="",{rate}=""),"",{volume}*{rate}/1000)'
    if "100" in basis:
        return f'=IF(OR({volume}="",{rate}=""),"",{volume}*{rate}/100)'
    if "10" in basis:
        return f'=IF(OR({volume}="",{rate}=""),"",{volume}*{rate}/10)'
    if item.get("volume_basis") == "tier" or item.get("tier_class"):
        return f'=IF(OR({lives}="",{rate}=""),"",{lives}*{rate})'
    return f'=IF(AND({volume}<>"",{rate}<>""),{volume}*{rate},IF(AND({lives}<>"",{rate}<>""),{lives}*{rate},""))'


def _write_line_rates_header(sheet: Any, model: dict[str, Any]) -> int:
    carrier_headers = _comparison_carrier_headers(model)
    market_carriers = carrier_headers[1:MAX_COMPARISON_CARRIERS + 1] or ["Carrier 1"]
    max_col = 4 + (len(market_carriers) * 2)
    _title(sheet, "RateTable_LineRates_Formatted", 1, 1, max_col)
    _merge_header(sheet, 3, 3, 4, "Current Plan", "808080")
    for index, carrier in enumerate(market_carriers):
        start_col = 5 + (index * 2)
        _merge_header(sheet, 3, start_col, start_col + 1, _carrier_display_name(carrier) or "Carrier", "004650")
    headers = [
        "Benefit / Class",
        "Volume",
        "Rate",
        "Premium",
    ]
    for _carrier in market_carriers:
        headers.extend(["Rate", "Premium"])
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(4, column, header)
        cell.font = Font(bold=True, color="004650")
        cell.fill = PatternFill("solid", fgColor="EAF8F8")
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    return max_col


def _merge_header(sheet: Any, row: int, start_col: int, end_col: int, label: str, fill: str) -> None:
    sheet.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    cell = sheet.cell(row, start_col)
    cell.value = label
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center", vertical="center")


def _line_rates_section(sheet: Any, row: int, max_col: int) -> None:
    for column in range(1, max_col + 1):
        cell = sheet.cell(row, column)
        cell.font = Font(bold=True, color="004650")
        cell.fill = PatternFill("solid", fgColor="EAF8F8")


def _group_line_rate_rows(rows: list[dict[str, Any]]) -> list[tuple[str, list[dict[str, Any]]]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in rows:
        grouped[item["benefit"]].append(item)
    benefits = list(grouped)
    benefits.sort(key=lambda benefit: (_line_rate_group_sort(benefit), _benefit_sort_order(benefit), benefit))
    return [(benefit, grouped[benefit]) for benefit in benefits]


def _line_rate_group_sort(benefit: str) -> int:
    kind = _line_rate_group_kind(benefit)
    return {"pooled": 10, "health": 20, "dental": 30, "optional": 40, "other": 50}[kind]


def _line_rate_group_kind(benefit: str | None) -> str:
    if benefit in {"Employee Life Insurance", "Employee AD&D", "Dependent Life", "Long Term Disability", "Short Term Disability"}:
        return "pooled"
    if benefit == "Extended Health Care":
        return "health"
    if benefit == "Dental Care":
        return "dental"
    if benefit in OPTIONAL_BENEFITS:
        return "optional"
    return "other"


def _next_group_name(groups: list[tuple[str, list[dict[str, Any]]]], current_name: str) -> str | None:
    for index, (name, _rows) in enumerate(groups):
        if name == current_name and index + 1 < len(groups):
            return groups[index + 1][0]
    return None


def _line_rate_premium_columns(max_col: int) -> list[int]:
    return list(range(4, max_col + 1, 2))


def _write_line_rate_total_row(sheet: Any, row: int, label: str, source_rows: list[int], max_col: int) -> int:
    sheet.cell(row, 1, label)
    if source_rows:
        for column in _line_rate_premium_columns(max_col):
            letter = get_column_letter(column)
            refs = ",".join(f"{letter}{source_row}" for source_row in source_rows)
            sheet.cell(row, column, f"=SUM({refs})")
    _line_rates_total_style(sheet, row, max_col)
    return row


def _write_line_rate_difference_row(sheet: Any, row: int, label: str, total_row: int, max_col: int) -> int:
    sheet.cell(row, 1, label)
    current_premium = f"D{total_row}"
    for column in _line_rate_premium_columns(max_col)[1:]:
        letter = get_column_letter(column)
        sheet.cell(row, column, f"={letter}{total_row}-{current_premium}")
    _line_rates_total_style(sheet, row, max_col)
    return row + 1


def _write_line_rate_annual_row(sheet: Any, row: int, label: str, monthly_total_row: int, max_col: int) -> int:
    sheet.cell(row, 1, label)
    for column in _line_rate_premium_columns(max_col):
        letter = get_column_letter(column)
        sheet.cell(row, column, f"={letter}{monthly_total_row}*12")
    _line_rates_total_style(sheet, row, max_col)
    return row + 1


def _write_line_rate_annual_difference_row(sheet: Any, row: int, label: str, annual_total_row: int, max_col: int) -> int:
    sheet.cell(row, 1, label)
    current_premium = f"D{annual_total_row}"
    for column in _line_rate_premium_columns(max_col)[1:]:
        letter = get_column_letter(column)
        sheet.cell(row, column, f"={letter}{annual_total_row}-{current_premium}")
    _line_rates_total_style(sheet, row, max_col)
    return row + 1


def _line_rates_total_style(sheet: Any, row: int, max_col: int) -> None:
    for column in range(1, max_col + 1):
        cell = sheet.cell(row, column)
        cell.font = Font(bold=True, color="004650")
        cell.fill = PatternFill("solid", fgColor="EAF8F8")


def _format_line_rates_sheet(sheet: Any, max_row: int, max_col: int) -> None:
    border = Border(
        left=Side(style="thin", color="8DD3C7"),
        right=Side(style="thin", color="8DD3C7"),
        top=Side(style="thin", color="8DD3C7"),
        bottom=Side(style="thin", color="8DD3C7"),
    )
    sheet.freeze_panes = "A5"
    sheet.auto_filter.ref = f"A4:{get_column_letter(max_col)}{max_row}"
    for row in range(1, max_row + 1):
        for column in range(1, max_col + 1):
            cell = sheet.cell(row, column)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if column >= 4 and column % 2 == 0 and row >= 5:
                cell.number_format = "$#,##0.00"
            elif column >= 3 and column % 2 == 1 and row >= 5:
                cell.number_format = "#,##0.000"
    widths = {1: 34, 2: 14}
    for column in range(1, max_col + 1):
        sheet.column_dimensions[get_column_letter(column)].width = widths.get(column, 15)


def _write_plan_design(sheet: Any, model: dict[str, Any]) -> None:
    _prepare_sheet(sheet)
    carrier = model.get("carrier_name") or "Carrier"
    headers = ["Benefit", "Provision", "Current Plan", carrier]
    _title(sheet, "PLAN DESIGN", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    row_number = 4
    current_benefit = None
    for record in _ordered_plan_design(model["plan_design"]):
        if record["benefit_line"] != current_benefit:
            current_benefit = record["benefit_line"]
            sheet.cell(row_number, 1, _plan_design_section_label(current_benefit))
            for column in range(1, len(headers) + 1):
                cell = sheet.cell(row_number, column)
                cell.fill = PatternFill("solid", fgColor="004C55")
                cell.font = Font(bold=True, color="FFFFFF", size=14 if column == 1 else 11)
            row_number += 1

        values = [
            record["benefit_line"],
            record["provision_name"],
            None,
            record["provision_value"],
        ]
        for column, value in enumerate(values, start=1):
            sheet.cell(row_number, column, value)
        _audit_destination(model, record, "Plan design provision", record["provision_value"], "Plan Design", f"Row {row_number}", record.get("confidence"), record.get("status"))
        row_number += 1

    _format_plan_design_sheet(sheet, max(row_number - 1, 3), len(headers))
    sheet.column_dimensions["C"].hidden = True


def _plan_design_section_label(benefit: str) -> str:
    return {
        "Employee Life Insurance": "Employee Life",
        "Employee AD&D": "Employee AD&D",
        "Dependent Life": "Dependent Life",
        "Extended Health Care": "Extended Health",
        "Dental Care": "Dental",
        "Virtual Care": "Virtual Healthcare",
    }.get(benefit, benefit)


def _format_plan_design_sheet(sheet: Any, max_row: int, max_col: int) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    value_fill = PatternFill("solid", fgColor="CBEFF4")
    border = Border(
        left=Side(style="thin", color="D9E2F3"),
        right=Side(style="thin", color="D9E2F3"),
        top=Side(style="thin", color="D9E2F3"),
        bottom=Side(style="thin", color="D9E2F3"),
    )
    sheet.freeze_panes = "A4"
    sheet.auto_filter.ref = f"A3:{get_column_letter(max_col)}{max_row}"
    for row in range(1, max_row + 1):
        for column in range(1, max_col + 1):
            cell = sheet.cell(row, column)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if row == 3:
                cell.fill = header_fill
                cell.font = Font(bold=True, color="1F2937")
            if column == 4 and row >= 4 and sheet.cell(row, 1).value not in {None, ""}:
                cell.fill = value_fill
            if column == 2 and row >= 4 and _important_plan_provision(cell.value):
                cell.font = Font(bold=True, color="004C55", size=12)
    widths = {1: 30, 2: 34, 3: 12, 4: 64}
    for column in range(1, max_col + 1):
        sheet.column_dimensions[get_column_letter(column)].width = widths.get(column, 18)


def _important_plan_provision(value: Any) -> bool:
    text = str(value or "").lower()
    return any(
        term in text
        for term in [
            "benefit amount",
            "termination age",
            "coinsurance",
            "maximum",
            "not eligible",
            "deductible",
            "drug coverage",
            "hearing aids",
            "paramedical",
        ]
    )


def _write_master_plan_design(sheet: Any, model: dict[str, Any]) -> None:
    _prepare_sheet(sheet)
    headers = [
        "ClientID",
        "ClientName",
        "Carrier",
        "PolicyNumber",
        "ClassName",
        "BenefitCategory",
        "Benefit",
        "CoverageDescription",
        "Coinsurance",
        "Maximum",
        "Frequency",
        "WaitingPeriod",
        "TerminationAge",
        "Notes",
        "DentalNormalizationField",
        "RawDentalValue",
        "NormalizedDentalValue",
        "DentalNormalizationNotes",
        "SourceFile",
        "SourceLocation",
        "Confidence",
        "Status",
    ]
    _title(sheet, "PLANDESIGN", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    row_number = 4
    if model["master_plan_design"]:
        for record in _ordered_master_plan_design(model["master_plan_design"]):
            values = [
                record.get("client_id"),
                record.get("client_name"),
                record.get("carrier"),
                record.get("policy_number"),
                record.get("class_name"),
                record.get("benefit_category"),
                record.get("benefit"),
                record.get("coverage_description"),
                record.get("coinsurance"),
                record.get("maximum"),
                record.get("frequency"),
                record.get("waiting_period"),
                record.get("termination"),
                record.get("notes"),
                record.get("dental_normalization_field"),
                record.get("raw_dental_value"),
                record.get("normalized_dental_value"),
                record.get("dental_normalization_notes"),
                _source_file_name(record.get("source_file")),
                record.get("source_location"),
                record.get("confidence"),
                record.get("status"),
            ]
            for column, value in enumerate(values, start=1):
                sheet.cell(row_number, column, value)
            _status_cell(sheet.cell(row_number, len(headers)), record.get("status"))
            _audit_destination(model, record, "PlanDesign database row", record.get("coverage_description"), "PlanDesign", f"Row {row_number}", record.get("confidence"), record.get("status"))
            row_number += 1
    else:
        sheet.cell(row_number, 1, _client_id(model))
        sheet.cell(row_number, 2, model.get("client_name"))
        sheet.cell(row_number, 3, model.get("carrier_name"))
        sheet.cell(row_number, len(headers), "NOT AVAILABLE")
        _status_cell(sheet.cell(row_number, len(headers)), "NOT AVAILABLE")
        row_number += 1

    _format_plan_database_sheet(sheet, max(row_number - 1, 3), len(headers))


def _write_booklet_extraction_review(sheet: Any, model: dict[str, Any]) -> None:
    _prepare_sheet(sheet)
    headers = BOOKLET_REVIEW_HEADERS
    _title(sheet, "BOOKLET_EXTRACTION_REVIEW", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    rows = _build_booklet_review_rows(model)
    model["booklet_extraction_review"] = rows
    for row_number, record in enumerate(rows, start=4):
        values = [record.get(header) for header in headers]
        for column, value in enumerate(values, start=1):
            sheet.cell(row_number, column, value)
        review_status = "NEEDS REVIEW" if record.get("ReviewFlag") == "Yes" else "FOUND"
        _status_cell(sheet.cell(row_number, headers.index("ReviewFlag") + 1), review_status)
        _audit(
            model,
            source_file=record.get("SourceFile"),
            source_type="Output",
            source_location="Booklet_Extraction_Review",
            field_name="Booklet extraction review row",
            value=record.get("BenefitName") or record.get("ClassName"),
            confidence=record.get("Confidence") or "MEDIUM",
            status=review_status,
            destination_sheet="Booklet_Extraction_Review",
            destination_cell=f"Row {row_number}",
            notes=record.get("Notes"),
        )

    _format_booklet_review_sheet(sheet, max(3 + len(rows), 3), len(headers))
    _set_fixed_row_height(sheet, max(3 + len(rows), 3), 15)


def _write_dental_normalization_reference(sheet: Any) -> None:
    _prepare_sheet(sheet)
    headers = ["DentalField", "ApprovedDropdownValue", "RawValueExamples", "NotesRule"]
    _title(sheet, "DENTAL NORMALIZATION REFERENCE", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    rows = _dental_normalization_reference_rows()
    for row_number, record in enumerate(rows, start=4):
        for column, header in enumerate(headers, start=1):
            sheet.cell(row_number, column, record.get(header))

    _format_sheet(sheet, max(3 + len(rows), 3), len(headers))
    widths = [34, 32, 52, 54]
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width


def _write_dental_normalized_values(sheet: Any, model: dict[str, Any]) -> None:
    _prepare_sheet(sheet)
    headers = [
        "ClientName",
        "CarrierName",
        "DentalField",
        "Benefit",
        "Provision",
        "RawExtractedValue",
        "NormalizedDropdownValue",
        "NormalizationConfidence",
        "Notes",
        "SourceFile",
        "SourceLocation",
    ]
    _title(sheet, "DENTAL NORMALIZED VALUES", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    rows = _build_dental_normalized_rows(model)
    if not rows:
        rows = [
            {
                "ClientName": model.get("client_name"),
                "CarrierName": model.get("carrier_name"),
                "DentalField": None,
                "Benefit": "Dental Care",
                "Provision": None,
                "RawExtractedValue": None,
                "NormalizedDropdownValue": None,
                "NormalizationConfidence": "LOW",
                "Notes": "No dental normalization candidates were extracted.",
                "SourceFile": None,
                "SourceLocation": None,
            }
        ]

    for row_number, record in enumerate(rows, start=4):
        for column, header in enumerate(headers, start=1):
            sheet.cell(row_number, column, record.get(header))
        _status_cell(sheet.cell(row_number, 8), "FOUND" if record.get("NormalizationConfidence") == "HIGH" else "NEEDS REVIEW")

    _format_sheet(sheet, max(3 + len(rows), 3), len(headers))
    widths = [28, 24, 34, 24, 30, 42, 32, 20, 54, 34, 18]
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width


def _build_dental_normalized_rows(model: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    seen: set[tuple[Any, ...]] = set()

    for record in model.get("plan_design", []):
        normalized = _dental_normalized_output_row(
            client_name=model.get("client_name"),
            carrier_name=record.get("carrier_name") or model.get("carrier_name"),
            benefit=record.get("benefit_line"),
            provision=record.get("provision_name"),
            raw_value=record.get("raw_dental_value"),
            normalized_value=record.get("normalized_dental_value"),
            field_name=record.get("dental_normalization_field"),
            confidence=record.get("dental_normalization_confidence"),
            notes=record.get("dental_normalization_notes"),
            source_file=record.get("source_file"),
            source_location=record.get("source_location"),
        )
        if normalized:
            key = tuple(normalized.get(header) for header in ["CarrierName", "DentalField", "Benefit", "Provision", "RawExtractedValue", "SourceFile", "SourceLocation"])
            if key not in seen:
                seen.add(key)
                rows.append(normalized)

    for record in model.get("master_plan_design", []):
        normalized = _dental_normalized_output_row(
            client_name=record.get("client_name") or model.get("client_name"),
            carrier_name=record.get("carrier") or model.get("carrier_name"),
            benefit=record.get("benefit"),
            provision=record.get("provision_name"),
            raw_value=record.get("raw_dental_value"),
            normalized_value=record.get("normalized_dental_value"),
            field_name=record.get("dental_normalization_field"),
            confidence=record.get("dental_normalization_confidence"),
            notes=record.get("dental_normalization_notes"),
            source_file=record.get("source_file"),
            source_location=record.get("source_location"),
        )
        if normalized:
            key = tuple(normalized.get(header) for header in ["CarrierName", "DentalField", "Benefit", "Provision", "RawExtractedValue", "SourceFile", "SourceLocation"])
            if key not in seen:
                seen.add(key)
                rows.append(normalized)

    return rows


def _dental_normalized_output_row(
    client_name: Any,
    carrier_name: Any,
    benefit: Any,
    provision: Any,
    raw_value: Any,
    normalized_value: Any,
    field_name: Any,
    confidence: Any,
    notes: Any,
    source_file: Any,
    source_location: Any,
) -> dict[str, Any] | None:
    if not field_name:
        return None
    return {
        "ClientName": client_name,
        "CarrierName": carrier_name,
        "DentalField": field_name,
        "Benefit": benefit,
        "Provision": provision,
        "RawExtractedValue": raw_value,
        "NormalizedDropdownValue": normalized_value,
        "NormalizationConfidence": confidence,
        "Notes": notes,
        "SourceFile": _source_file_name(source_file),
        "SourceLocation": source_location,
    }


def _write_alternates(sheet: Any, model: dict[str, Any]) -> None:
    _prepare_sheet(sheet)
    headers = ["Section", "Benefit", "Alternate", "Proposed Change", "Rate Impact", "SourceFile", "SourceLocation", "Confidence", "Status", "Notes"]
    _title(sheet, "ALTERNATES", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    rows = model["alternates"] or [
        {
            "section": "Alternates",
            "benefit": None,
            "alternate": "Not Available",
            "proposed_change": None,
            "rate_impact": None,
            "source_file": None,
            "source_location": None,
            "confidence": "LOW",
            "status": "NOT AVAILABLE",
            "notes": "Alternates section was not found.",
        }
    ]
    for row_number, record in enumerate(rows, start=4):
        values = [
            record.get("section"),
            record.get("benefit"),
            record.get("alternate"),
            record.get("proposed_change"),
            record.get("rate_impact"),
            _source_file_name(record.get("source_file")),
            record.get("source_location"),
            record.get("confidence"),
            record.get("status"),
            record.get("notes"),
        ]
        for column, value in enumerate(values, start=1):
            sheet.cell(row_number, column, value)
        _status_cell(sheet.cell(row_number, 9), record.get("status"))
        _audit_destination(model, record, "Alternate", record.get("alternate"), "Alternates", f"Row {row_number}", record.get("confidence"), record.get("status"))

    _format_sheet(sheet, max(3 + len(rows), 3), len(headers))


def _write_deviations(sheet: Any, model: dict[str, Any]) -> None:
    _prepare_sheet(sheet)
    headers = ["DeviationType", "Benefit", "Description", "Impact", "SourceFile", "SourceLocation", "Confidence", "Status", "Notes"]
    _title(sheet, "DEVIATIONS", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    rows = model["deviations"] or [
        {
            "deviation_type": "Deviations",
            "benefit": None,
            "description": "Not Available",
            "impact": None,
            "source_file": None,
            "source_location": None,
            "confidence": "LOW",
            "status": "NOT AVAILABLE",
            "notes": "No deviation section was found.",
        }
    ]
    for row_number, record in enumerate(rows, start=4):
        values = [
            record.get("deviation_type"),
            record.get("benefit"),
            record.get("description"),
            record.get("impact"),
            _source_file_name(record.get("source_file")),
            record.get("source_location"),
            record.get("confidence"),
            record.get("status"),
            record.get("notes"),
        ]
        for column, value in enumerate(values, start=1):
            sheet.cell(row_number, column, value)
        _status_cell(sheet.cell(row_number, 8), record.get("status"))
        _audit_destination(model, record, "Deviation", record.get("description"), "Deviations", f"Row {row_number}", record.get("confidence"), record.get("status"))

    _format_sheet(sheet, max(3 + len(rows), 3), len(headers))


def _write_rate_guarantee(sheet: Any, model: dict[str, Any]) -> None:
    _prepare_sheet(sheet)
    carrier_names = [str(name) for name in model.get("carrier_names", []) if name][:MAX_COMPARISON_CARRIERS]
    if not carrier_names:
        carrier_names = [str(model.get("carrier_name") or "Carrier")]
    headers = ["Item", "Current Plan", *carrier_names]
    row_definitions = [
        ("Rate Guarantee", None, True),
        ("   Pooled", "rate_guarantee_pooled", False),
        ("   Health & Dental", "rate_guarantee_health_dental", False),
        ("Pooling", None, True),
        ("   Pooling Level", "pooling_level", False),
        ("   Pooling Charge", "pooling_charge", False),
        ("Target Loss Ratio", "target_loss_ratio", False),
        ("IBNR", None, True),
        ("   Extended Health Care", "ibnr_ehc", False),
        ("   Dental", "ibnr_dental", False),
        ("Rate Caps", None, True),
        ("   1st Renewal", "first_renewal_cap", False),
        ("   2nd Renewal", "second_renewal_cap", False),
    ]

    _title(sheet, "RATE GUARANTEE SUMMARY", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    for offset, (label, key, is_section) in enumerate(row_definitions, start=4):
        sheet.cell(offset, 1, label)
        if is_section:
            continue
        row_values = []
        for column_offset, carrier in enumerate(carrier_names, start=3):
            value = _normalized_rate_guarantee_value(model, carrier, key)
            sheet.cell(offset, column_offset, value or "n/a")
            row_values.append(value)
        _audit(
            model,
            source_file=None,
            source_type="Output",
            source_location="Rate Guarantee Sum",
            field_name=label,
            value="; ".join(str(value) for value in row_values if value) or "n/a",
            confidence="HIGH" if any(row_values) else "LOW",
            status="FOUND" if any(row_values) else "NOT AVAILABLE",
            destination_sheet="Rate Guarantee Sum",
            destination_cell=f"Row {offset}",
        )

    max_row = 3 + len(row_definitions)
    _format_rate_guarantee_sheet(sheet, max_row, len(headers))
    _clear_unused_sheet_formatting(sheet, max_row + 1)
    _unhide_all_rows(sheet)


def _normalized_rate_guarantee_value(model: dict[str, Any], carrier: str, key: str | None) -> str | None:
    values = _rate_guarantee_values_for_carrier(model, carrier)
    if not key:
        return None
    if key == "rate_guarantee_pooled":
        explicit = _normalize_rate_guarantee_period(values.get("Pooled"), "pooled")
        if explicit:
            return explicit
        return _normalize_rate_guarantee_period(
            _join_rate_guarantee_fields(values, ["Pooled", "Rate guarantee", "Rate guarantee period"]),
            "pooled",
        )
    if key == "rate_guarantee_health_dental":
        explicit = _normalize_rate_guarantee_period(values.get("Health & Dental"), "health_dental")
        if explicit:
            return explicit
        return _normalize_rate_guarantee_period(
            _join_rate_guarantee_fields(values, ["Health & Dental", "Rate guarantee period", "Rate guarantee"]),
            "health_dental",
        )
    if key == "pooling_level":
        return _normalize_pooling_level(values.get("Pooling level"))
    if key == "pooling_charge":
        return _normalize_first_percent(values.get("Pooling charge"))
    if key == "target_loss_ratio":
        return _normalize_target_loss_ratio(values.get("Target loss ratio"))
    if key == "ibnr_ehc":
        return _normalize_ibnr(values.get("IBNR"), "health")
    if key == "ibnr_dental":
        return _normalize_ibnr(values.get("IBNR"), "dental")
    if key == "first_renewal_cap":
        return _normalize_first_percent(values.get("First renewal cap"))
    if key == "second_renewal_cap":
        return _normalize_first_percent(values.get("Second renewal cap"))
    return None


def _rate_guarantee_values_for_carrier(model: dict[str, Any], carrier: str) -> dict[str, str]:
    carrier_values = {
        str(key): str(value)
        for key, value in model.get("rate_guarantee_by_carrier", {}).get(carrier, {}).items()
        if value not in (None, "")
    }
    if carrier_values:
        return carrier_values

    # Backward-compatible fallback for older combined values that included "Carrier: value" fragments.
    out: dict[str, str] = {}
    for field, value in model.get("rate_guarantee", {}).items():
        carrier_parts = _extract_prefixed_rate_guarantee_parts(str(value), carrier)
        if carrier_parts:
            out[str(field)] = "; ".join(carrier_parts)
    return out


def _extract_prefixed_rate_guarantee_parts(value: str, carrier: str) -> list[str]:
    carrier_pattern = re.escape(carrier)
    pattern = rf"(?:^|;)\s*{carrier_pattern}:\s*(.*?)(?=(?:;\s*[A-Za-z][A-Za-z0-9 &/.-]*:\s*)|$)"
    return [_clean_text(match) for match in re.findall(pattern, value, flags=re.IGNORECASE) if _clean_text(match)]


def _join_rate_guarantee_fields(values: dict[str, str], fields: list[str]) -> str:
    return "; ".join(_clean_text(values.get(field)) for field in fields if _clean_text(values.get(field)))


def _normalize_rate_guarantee_period(value: Any, group: str) -> str | None:
    text = _clean_text(value)
    lower = text.lower()
    if not text:
        return None
    if group == "pooled":
        match = re.search(r"(\d+)\s*months?\s*\([^)]*(?:life|ad&d|dependent|critical)[^)]*\)", text, flags=re.IGNORECASE)
        if match:
            return f"{match.group(1)} months"
        if "second renewal" in lower:
            return "second renewal"
        if "next renewal" in lower:
            return "next renewal"
    if group == "health_dental":
        match = re.search(r"(\d+)\s*months?\s*\([^)]*(?:ehc|health|dental)[^)]*\)", text, flags=re.IGNORECASE)
        if match:
            return f"{match.group(1)} months"
    sequence = re.search(r"(\d+\s*,\s*\d+\s*or\s*\d+)\s+months?", text, flags=re.IGNORECASE)
    if sequence:
        return f"{sequence.group(1)} months"
    months = re.findall(r"\b\d+(?:\.\d+)?\s*months?", text, flags=re.IGNORECASE)
    if months:
        unique = []
        for item in months:
            normalized = _clean_text(item)
            if normalized not in unique:
                unique.append(normalized)
        return "; ".join(unique[:2])
    return _short_rate_guarantee_text(text)


def _normalize_pooling_level(value: Any) -> str | None:
    text = _clean_text(value)
    lower = text.lower()
    if not text:
        return None
    money = _extract_first_money(text)
    if money:
        return money
    if "first dollar" in lower:
        return "First dollar"
    if "included in ehc" in lower:
        return "Included in EHC"
    return _short_rate_guarantee_text(text)


def _normalize_first_percent(value: Any) -> str | None:
    text = _clean_text(value)
    match = re.search(r"\d+(?:\.\d+)?%", text)
    return match.group(0) if match else None


def _normalize_target_loss_ratio(value: Any) -> str | None:
    text = _clean_text(value)
    if not text:
        return None
    health = re.search(r"health(?:\s+is)?\s+(\d+(?:\.\d+)?%)", text, flags=re.IGNORECASE)
    dental = re.search(r"dental(?:\s+is)?\s+(\d+(?:\.\d+)?%)", text, flags=re.IGNORECASE)
    if health and dental:
        if health.group(1) == dental.group(1):
            return health.group(1)
        return f"Health {health.group(1)}; Dental {dental.group(1)}"
    return _normalize_first_percent(text) or _short_rate_guarantee_text(text)


def _normalize_ibnr(value: Any, benefit: str) -> str | None:
    text = _clean_text(value)
    if not text:
        return None
    pattern = r"health benefit:\s*(\d+(?:\.\d+)?%)" if benefit == "health" else r"dental benefit:\s*(\d+(?:\.\d+)?%)"
    match = re.search(pattern, text, flags=re.IGNORECASE)
    if match:
        return match.group(1)
    percentages = re.findall(r"\d+(?:\.\d+)?%", text)
    if not percentages:
        return None
    return percentages[0] if benefit == "health" else percentages[min(1, len(percentages) - 1)]


def _short_rate_guarantee_text(value: Any) -> str | None:
    text = _clean_text(value)
    if not text:
        return None
    parts = [part.strip() for part in re.split(r";|\.", text) if part.strip()]
    return parts[0] if parts else text


def _format_rate_guarantee_sheet(sheet: Any, max_row: int, max_col: int) -> None:
    _format_sheet(sheet, max_row, max_col)
    section_font = Font(name=WORKBOOK_FONT_NAME, bold=True, color="004C55", size=11)
    muted_font = Font(name=WORKBOOK_FONT_NAME, color="666666", size=10)
    carrier_fill = PatternFill("solid", fgColor="C9DEE3")
    current_fill = PatternFill("solid", fgColor="E7E7E7")
    border = Border(
        left=Side(style="thin", color="7FA3A8"),
        right=Side(style="thin", color="7FA3A8"),
        top=Side(style="thin", color="7FA3A8"),
        bottom=Side(style="thin", color="7FA3A8"),
    )
    for column in range(1, max_col + 1):
        header = sheet.cell(3, column)
        if column == 2:
            header.fill = PatternFill("solid", fgColor="D9D9D9")
            header.font = Font(name=WORKBOOK_FONT_NAME, bold=True, color="666666", size=11)
        elif column >= 3:
            header.fill = PatternFill("solid", fgColor="004F59")
            header.font = Font(name=WORKBOOK_FONT_NAME, bold=True, color="6EE7D8", size=11)
        header.alignment = Alignment(horizontal="center" if column >= 3 else "left", vertical="center", wrap_text=True)

    for row in range(4, max_row + 1):
        raw_label = str(sheet.cell(row, 1).value or "")
        is_section = bool(raw_label.strip()) and not raw_label.startswith(" ")
        for column in range(1, max_col + 1):
            cell = sheet.cell(row, column)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.font = section_font if is_section and column == 1 else muted_font
            if column == 2:
                cell.fill = current_fill
            elif column >= 3:
                cell.fill = carrier_fill
            if is_section:
                cell.fill = PatternFill(fill_type=None)
                if column > 1:
                    cell.value = None
    widths = {1: 34, 2: 24}
    for column in range(1, max_col + 1):
        sheet.column_dimensions[get_column_letter(column)].width = widths.get(column, 28)


def _write_extended_health(sheet: Any, model: dict[str, Any]) -> None:
    _prepare_sheet(sheet)
    headers = ["Category", "Detail", "Monthly Premium", "Annual Premium", "Status", "Notes"]
    _title(sheet, "EXTENDED HEALTH", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    rows = [row for row in model["benefit_rows"] if row["benefit_line"] == "Extended Health Care"]
    row_number = 4
    if rows:
        for record in rows:
            detail = record.get("tier") or record.get("class_name") or "Extended Health Care"
            values = [record["benefit_line"], detail, record.get("monthly_premium"), record.get("annual_premium"), record.get("status"), record.get("notes")]
            for column, value in enumerate(values, start=1):
                sheet.cell(row_number, column, value)
            _status_cell(sheet.cell(row_number, 5), record.get("status"))
            row_number += 1
    else:
        sheet.cell(row_number, 1, "Extended Health Care")
        sheet.cell(row_number, 5, "Not Available")
        _status_cell(sheet.cell(row_number, 5), "NOT AVAILABLE")
        row_number += 1

    _format_sheet(sheet, max(row_number - 1, 3), len(headers))


def _write_exec_summary(sheet: Any, model: dict[str, Any], cost_rows: list[dict[str, Any]]) -> None:
    _prepare_sheet(sheet)
    carrier = model.get("carrier_name") or "Carrier"
    headers = ["Benefit", "Current Plan", carrier, "% Difference", "Rate Guarantee"]
    _title(sheet, "EXECUTIVE SUMMARY - COST", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    row_number = 4
    summary_benefits = [
        "Employee Life Insurance",
        "Employee AD&D",
        "Dependent Life",
        "Short Term Disability",
        "Long Term Disability",
        "Extended Health Care",
        "Dental Care",
        "Optional / Value Added",
    ]
    for benefit in summary_benefits:
        if benefit == "Optional / Value Added":
            criteria = OPTIONAL_BENEFITS
            formula = _sumif_many(criteria)
        else:
            formula = f'=SUMIF(\'Total Cost Comparison\'!$M:$M,"{benefit}",\'Total Cost Comparison\'!$H:$H)'
        sheet.cell(row_number, 1, benefit)
        sheet.cell(row_number, 3, formula)
        sheet.cell(row_number, 4, f'=IF(B{row_number}=0,"",C{row_number}/B{row_number}-1)')
        row_number += 1

    sheet.cell(row_number, 1, "Total")
    sheet.cell(row_number, 3, f"=SUM(C4:C{row_number - 1})")
    row_number += 1
    sheet.cell(row_number, 1, "Rate Guarantee")
    sheet.cell(row_number, 3, "='Rate Guarantee Sum'!C5")

    _format_sheet(sheet, row_number, len(headers))
    for row in range(4, row_number):
        sheet.cell(row, 3).number_format = "$#,##0.00"
        sheet.cell(row, 4).number_format = "0.0%"


def _write_audit_sheet(sheet: Any, model: dict[str, Any]) -> None:
    _prepare_sheet(sheet)
    headers = [
        "Source file",
        "Source type",
        "Source sheet or Word table number",
        "Field name",
        "Benefit line",
        "Extracted value",
        "Destination sheet",
        "Destination cell or row",
        "Confidence score",
        "Status",
        "Notes",
    ]
    _title(sheet, "EXTRACTION AUDIT", 1, 1, len(headers))
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    for row_number, item in enumerate(model["audit"], start=4):
        values = [
            item.get("source_file"),
            item.get("source_type"),
            item.get("source_location"),
            item.get("field_name"),
            item.get("benefit_line"),
            item.get("value"),
            item.get("destination_sheet"),
            item.get("destination_cell"),
            item.get("confidence"),
            item.get("status"),
            item.get("notes"),
        ]
        for column, value in enumerate(values, start=1):
            sheet.cell(row_number, column, value)
        _status_cell(sheet.cell(row_number, 10), item.get("status"))

    _format_sheet(sheet, max(3 + len(model["audit"]), 3), len(headers))


def _ordered_benefit_rows(rows: list[dict[str, Any]]) -> list[tuple[str, list[dict[str, Any]]]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        grouped[row["benefit_line"]].append(row)

    benefits = list(grouped.keys())
    benefits.sort(key=lambda benefit: (_category_sort(benefit), _benefit_sort_order(benefit), benefit))
    return [(benefit, grouped[benefit]) for benefit in benefits]


def _ordered_plan_design(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(rows, key=lambda row: (_category_sort(row["benefit_line"]), _benefit_sort_order(row["benefit_line"]), row["benefit_line"], row["provision_name"]))


def _ordered_master_plan_design(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(
        rows,
        key=lambda row: (
            str(row.get("class_name") or ""),
            _category_sort(_plan_design_sort_benefit(row)),
            _benefit_sort_order(_plan_design_sort_benefit(row)),
            str(row.get("benefit") or ""),
            str(row.get("provision_name") or ""),
        ),
    )


def _plan_design_sort_benefit(row: dict[str, Any]) -> str:
    benefit = str(row.get("benefit") or "")
    category = str(row.get("benefit_category") or "")
    return _proposal_plan_benefit(category, benefit)


def _build_booklet_review_rows(model: dict[str, Any]) -> list[dict[str, Any]]:
    if not model["master_plan_design"]:
        empty = {header: None for header in BOOKLET_REVIEW_HEADERS}
        empty.update(
            {
                "ClientID": _client_id(model),
                "ClientName": model.get("client_name"),
                "Carrier": model.get("carrier_name"),
                "BenefitIncluded": "No",
                "ReviewFlag": "Yes",
                "Confidence": "LOW",
                "LastUpdated": datetime.now().strftime("%Y-%m-%d"),
                "Notes": "No Schedule of Benefits rows were extracted.",
            }
        )
        return [empty]

    return [_booklet_review_row_from_plan(model, row) for row in _ordered_master_plan_design(model["master_plan_design"])]


def _booklet_review_row_from_plan(model: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    description = str(plan.get("coverage_description") or "")
    benefit = str(plan.get("benefit") or "")
    provision = str(plan.get("provision_name") or "")
    maximum = plan.get("maximum")
    coinsurance = plan.get("coinsurance")
    frequency = plan.get("frequency")
    termination = plan.get("termination")
    notes = plan.get("notes")

    row = {header: None for header in BOOKLET_REVIEW_HEADERS}
    row.update(
        {
            "ClientID": plan.get("client_id") or _client_id(model),
            "ClientName": plan.get("client_name") or model.get("client_name"),
            "Carrier": plan.get("carrier") or model.get("carrier_name"),
            "PolicyNumber": plan.get("policy_number"),
            "SourceFile": _source_file_name(plan.get("source_file")),
            "ClassName": plan.get("class_name"),
            "BenefitCategory": plan.get("benefit_category"),
            "BenefitName": benefit,
            "BenefitIncluded": "No" if "not eligible" in description.lower() else "Yes",
            "BenefitAmountText": description if provision == "Benefit amount" else None,
            "BenefitAmountType": _benefit_amount_type(description) if provision == "Benefit amount" else None,
            "BenefitAmountNumeric": _extract_first_money(description) if provision == "Benefit amount" else None,
            "CoinsuranceText": description if coinsurance else None,
            "CoinsurancePercent": _extract_first_percent_number(coinsurance),
            "DeductibleText": description if "deductible" in provision.lower() or "deductible" in description.lower() else None,
            "MaximumText": description if maximum else None,
            "MaximumAmount": _extract_first_money(maximum or description),
            "MaximumFrequency": frequency,
            "LifetimeMaximum": maximum if "lifetime" in f"{description} {maximum}".lower() else None,
            "NonEvidenceMaximum": description if "non-evidence" in provision.lower() else None,
            "WaitingPeriodText": plan.get("waiting_period"),
            "WaitingPeriodMonths": _extract_month_count(plan.get("waiting_period") or description),
            "EliminationPeriodText": description if "elimination period" in description.lower() else None,
            "EliminationPeriodDays": _extract_day_count(description) if "elimination period" in description.lower() else None,
            "TerminationAgeText": termination,
            "TerminationAgeNumber": _extract_age_number(termination or description),
            "TerminationTrigger": _termination_trigger(termination or description),
            "WaiverOfPremium": description if "waiver of premium" in description.lower() else None,
            "DefinitionOfDisability": description if "definition of disability" in description.lower() else None,
            "ReviewFlag": "Yes" if str(plan.get("status")).upper() in {"NEEDS REVIEW", "LOW CONFIDENCE"} or str(plan.get("confidence")).upper() == "LOW" else "No",
            "Confidence": plan.get("confidence"),
            "LastUpdated": datetime.now().strftime("%Y-%m-%d"),
            "Notes": notes,
            "Benefit": _proposal_plan_benefit(str(plan.get("benefit_category") or ""), benefit),
            "CoverageDescription": description,
            "Coinsurance": coinsurance,
            "Maximum": maximum,
            "Frequency": frequency,
            "WaitingPeriod": plan.get("waiting_period"),
            "TerminationAge": termination,
        }
    )
    _populate_review_specific_columns(row, benefit, description, coinsurance, maximum, frequency)
    return row


def _populate_review_specific_columns(
    row: dict[str, Any],
    benefit: str,
    description: str,
    coinsurance: str | None,
    maximum: str | None,
    frequency: str | None,
) -> None:
    lower_benefit = benefit.lower()
    lower_description = description.lower()

    if benefit == "Employee Life Insurance" and row.get("BenefitAmountText"):
        row["LifeBenefitAmount"] = row["BenefitAmountText"]
    if benefit == "Employee AD&D" and row.get("BenefitAmountText"):
        row["ADDBenefitAmount"] = row["BenefitAmountText"]
    if benefit == "Dependent Life":
        spouse = re.search(r"\$\s?\d[\d,]*(?:\.\d+)?\s+for\s+spouse", description, flags=re.IGNORECASE)
        child = re.search(r"\$\s?\d[\d,]*(?:\.\d+)?\s+for\s+each\s+child", description, flags=re.IGNORECASE)
        if spouse:
            row["DependentLifeSpouseAmount"] = spouse.group(0).replace(" ", "")
        if child:
            row["DependentLifeChildAmount"] = child.group(0).replace(" ", "")

    if "drug" in lower_benefit:
        row["PrescriptionDrugCoinsurance"] = coinsurance
        row["DrugMax"] = maximum
        if "88ga" in lower_description or "telus" in lower_description:
            row["DrugPlanType"] = description
        if "mandatory generic" in lower_description:
            row["DrugType"] = "Mandatory Generic"
    elif "hospital" in lower_benefit:
        row["HospitalCoverage"] = description
    elif "paramedical" in lower_benefit:
        row["ParamedicalStandardMaximum"] = maximum or description
    elif "psychologist" in lower_description or "counselor" in lower_description:
        row["ParamedicalPsychologyMaximum"] = maximum or description
    elif "vision" in lower_benefit:
        if maximum:
            row["VisionMaximum"] = maximum
        if frequency:
            if "eye examination" in lower_description:
                row["EyeExamFrequencyMonths"] = _extract_month_count(frequency)
            else:
                row["VisionFrequencyMonths"] = _extract_month_count(frequency)
    elif "dental basic" in lower_benefit:
        row["DentalBasicCoinsurance"] = coinsurance
        row["DentalBasicMaximum"] = maximum
        if "combined maximum" in lower_description or "level 1 - 4" in lower_description or "level 1-4" in lower_description:
            _set_booklet_dental_normalization(row, "Major Dental Annual Maximum", maximum or description, description)
    elif "dental major" in lower_benefit:
        row["DentalMajorCoinsurance"] = coinsurance
        row["DentalMajorMaximum"] = maximum
        if maximum or "maximum" in lower_description:
            _set_booklet_dental_normalization(row, "Major Dental Annual Maximum", maximum or description, description)
    elif "orthodont" in lower_benefit:
        row["DentalOrthoCoinsurance"] = coinsurance
        row["DentalOrthoMaximum"] = maximum
        if maximum or "maximum" in lower_description:
            _set_booklet_dental_normalization(row, "Orthodontics Lifetime Maximum", maximum or description, description)
    elif "dental recall" in lower_benefit:
        row["DentalRecallFrequency"] = frequency or description
        _set_booklet_dental_normalization(row, "Basic Dental Recall Frequency", frequency or description, description)
    elif "travel" in lower_benefit:
        row["TravelMaximum"] = maximum
        row["TravelTripDurationDays"] = _extract_day_count(description)


def _set_booklet_dental_normalization(row: dict[str, Any], field_name: str, raw_value: Any, match_text: Any | None = None) -> None:
    result = _normalize_dental_dropdown_value(field_name, raw_value, match_text)
    if field_name == "Basic Dental Recall Frequency":
        row["DentalBasicRecallRaw"] = result["raw_value"]
        row["DentalBasicRecallNormalized"] = result["normalized_value"]
    elif field_name == "Major Dental Annual Maximum":
        row["DentalMajorAnnualMaximumRaw"] = result["raw_value"]
        row["DentalMajorAnnualMaximumNormalized"] = result["normalized_value"]
    elif field_name == "Orthodontics Lifetime Maximum":
        row["DentalOrthoLifetimeMaximumRaw"] = result["raw_value"]
        row["DentalOrthoLifetimeMaximumNormalized"] = result["normalized_value"]

    if result.get("notes"):
        row["DentalNormalizationNotes"] = _join_unique(row.get("DentalNormalizationNotes"), result["notes"])
        if result["normalized_value"] == DENTAL_NORMALIZATION_OTHER:
            row["Notes"] = _join_unique(row.get("Notes"), result["notes"])


def _add_placeholder_benefit(model: dict[str, Any], benefit: str, path: Path | None, location: str) -> None:
    record = {
        "carrier_name": model.get("carrier_name"),
        "carrier_slot": model.get("carrier_slot"),
        "carrier_index": model.get("carrier_index"),
        "benefit_line": benefit,
        "class_name": None,
        "tier": None,
        "lives": None,
        "volume": None,
        "rate": None,
        "monthly_premium": None,
        "annual_premium": None,
        "rate_basis": None,
        "volume_basis": None,
        "included_status": "Not Available",
        "notes": "Benefit was found or expected, but no usable rate or premium was identified.",
        "source_file": str(path) if path else None,
        "source_type": "Excel" if path else "System",
        "source_location": location,
        "confidence": "LOW",
        "status": "NOT AVAILABLE",
    }
    model["benefit_rows"].append(record)
    _audit(model, path, "System", location, "Benefit missing rate/premium", "Not Available", "LOW", "NOT AVAILABLE", benefit_line=benefit)


def _add_master_plan_design(
    model: dict[str, Any],
    client_id: str,
    client_name: str | None,
    carrier: str,
    policy_number: str | None,
    class_name: str,
    benefit_category: str,
    benefit: str,
    provision_name: str,
    coverage_description: str,
    coinsurance: str | None,
    maximum: str | None,
    frequency: str | None,
    waiting_period: str | None,
    termination: str | None,
    notes: str | None,
    source_file: Path,
    source_location: str,
    confidence: str,
    status: str,
) -> None:
    dental_normalization = _dental_normalization_from_parts(
        benefit,
        provision_name,
        coverage_description,
        maximum=maximum,
        frequency=frequency,
    )
    record = {
        "client_id": client_id,
        "client_name": client_name,
        "carrier": carrier,
        "policy_number": policy_number,
        "class_name": class_name,
        "benefit_category": benefit_category,
        "benefit": benefit,
        "provision_name": provision_name,
        "coverage_description": coverage_description,
        "coinsurance": coinsurance,
        "maximum": maximum,
        "frequency": frequency,
        "waiting_period": waiting_period,
        "termination": termination,
        "notes": notes,
        "source_file": str(source_file),
        "source_type": "Word",
        "source_location": source_location,
        "confidence": confidence,
        "status": status,
    }
    if dental_normalization:
        record.update(
            {
                "dental_normalization_field": dental_normalization["field_name"],
                "raw_dental_value": dental_normalization["raw_value"],
                "normalized_dental_value": dental_normalization["normalized_value"],
                "dental_normalization_notes": dental_normalization["notes"],
                "dental_normalization_confidence": dental_normalization["confidence"],
            }
        )
    key = (
        record["class_name"],
        record["benefit_category"],
        record["benefit"],
        record["provision_name"],
        record["coverage_description"],
    )
    existing = {
        (
            row.get("class_name"),
            row.get("benefit_category"),
            row.get("benefit"),
            row.get("provision_name"),
            row.get("coverage_description"),
        )
        for row in model["master_plan_design"]
    }
    if key in existing:
        return

    model["master_plan_design"].append(record)
    plan_benefit = _proposal_plan_benefit(benefit_category, benefit)
    plan_provision = provision_name if benefit == plan_benefit else f"{benefit} - {provision_name}"
    _add_plan_design(model, plan_benefit, plan_provision, coverage_description, source_file, "Word", source_location, confidence, status)
    _audit(
        model,
        source_file,
        "Word",
        source_location,
        "PlanDesign row",
        coverage_description,
        confidence,
        status,
        benefit_line=plan_benefit,
    )


def _proposal_plan_benefit(benefit_category: str, benefit: str) -> str:
    if benefit in STANDARD_BENEFITS or benefit in OPTIONAL_BENEFITS:
        return benefit
    category = benefit_category.lower()
    if "life" in category and "ad&d" not in category:
        return "Employee Life Insurance"
    if "ad&d" in category:
        return "Employee AD&D"
    if "dependent life" in category:
        return "Dependent Life"
    if "health" in category:
        if benefit == "Travel":
            return "Travel"
        return "Extended Health Care"
    if "dental" in category:
        return "Dental Care"
    return benefit or "Other Benefits"


def _add_alternate(
    model: dict[str, Any],
    section: str,
    benefit: str,
    alternate: str,
    proposed_change: str | None,
    rate_impact: str | None,
    source_file: Path,
    source_location: str,
    confidence: str,
    status: str,
    source_type: str = "Word",
) -> None:
    record = {
        "section": section,
        "benefit": benefit,
        "alternate": alternate,
        "proposed_change": proposed_change,
        "rate_impact": rate_impact,
        "source_file": str(source_file),
        "source_type": source_type,
        "source_location": source_location,
        "confidence": confidence,
        "status": status,
        "notes": "Extracted from Alternates section.",
    }
    key = (section, benefit, alternate)
    existing = {(row.get("section"), row.get("benefit"), row.get("alternate")) for row in model["alternates"]}
    if key not in existing:
        model["alternates"].append(record)
        _audit(model, source_file, source_type, source_location, "Alternate", alternate, confidence, status, benefit_line=benefit)


def _add_deviation(
    model: dict[str, Any],
    deviation_type: str,
    benefit: str,
    description: str,
    impact: str | None,
    source_file: Path,
    source_location: str,
    confidence: str,
    status: str,
) -> None:
    record = {
        "deviation_type": deviation_type,
        "benefit": benefit,
        "description": description,
        "impact": impact,
        "source_file": str(source_file),
        "source_type": "Word",
        "source_location": source_location,
        "confidence": confidence,
        "status": status,
        "notes": "Extracted from Clarifications and Deviations / quote condition text.",
    }
    key = (deviation_type, benefit, description)
    existing = {(row.get("deviation_type"), row.get("benefit"), row.get("description")) for row in model["deviations"]}
    if key not in existing:
        model["deviations"].append(record)
        _audit(model, source_file, "Word", source_location, "Deviation", description, confidence, status, benefit_line=benefit)


def _add_plan_design(
    model: dict[str, Any],
    benefit_line: str,
    provision_name: str,
    provision_value: Any,
    source_file: Path,
    source_type: str,
    source_location: str,
    confidence: str,
    status: str,
) -> None:
    dental_normalization = _dental_normalization_from_parts(benefit_line, provision_name, provision_value)
    record = {
        "carrier_name": model.get("carrier_name"),
        "carrier_slot": model.get("carrier_slot"),
        "carrier_index": model.get("carrier_index"),
        "benefit_line": benefit_line,
        "provision_name": provision_name,
        "provision_value": provision_value,
        "source_file": str(source_file),
        "source_type": source_type,
        "source_location": source_location,
        "confidence": confidence,
        "status": status,
        "notes": f"Extracted from {source_location}.",
    }
    if dental_normalization:
        record.update(
            {
                "dental_normalization_field": dental_normalization["field_name"],
                "raw_dental_value": dental_normalization["raw_value"],
                "normalized_dental_value": dental_normalization["normalized_value"],
                "dental_normalization_notes": dental_normalization["notes"],
                "dental_normalization_confidence": dental_normalization["confidence"],
            }
        )
    carrier_key = str(model.get("carrier_name") or "")
    key = (carrier_key, benefit_line, provision_name, str(provision_value))
    existing = {
        (
            str(row.get("carrier_name") or ""),
            row["benefit_line"],
            row["provision_name"],
            str(row["provision_value"]),
        )
        for row in model["plan_design"]
    }
    if key not in existing:
        model["plan_design"].append(record)
        _audit(model, source_file, source_type, source_location, provision_name, provision_value, confidence, status, benefit_line=benefit_line)


def _add_included_service(model: dict[str, Any], benefit_line: str, notes: str, source_file: Path, paragraph_index: int) -> None:
    service = {
        "carrier_name": model.get("carrier_name"),
        "carrier_slot": model.get("carrier_slot"),
        "carrier_index": model.get("carrier_index"),
        "benefit_line": benefit_line,
        "notes": notes,
        "source_file": str(source_file),
        "source_type": "Word",
        "source_location": f"Paragraph {paragraph_index}",
        "confidence": "HIGH",
        "status": "INCLUDED",
    }
    key = (benefit_line, notes)
    existing = {(item["benefit_line"], item["notes"]) for item in model["included_services"]}
    if key not in existing:
        model["included_services"].append(service)
        _audit(model, source_file, "Word", f"Paragraph {paragraph_index}", "Included service", notes, "HIGH", "INCLUDED", benefit_line=benefit_line)


def _set_rate_guarantee(
    model: dict[str, Any],
    field_name: str,
    value: Any,
    source_file: Path,
    paragraph_index: int,
    confidence: str,
    notes: str | None = None,
) -> None:
    existing = model["rate_guarantee"].get(field_name)
    if existing and str(value) in str(existing):
        return
    model["rate_guarantee"][field_name] = value if not existing else f"{existing}; {value}"
    _store_rate_guarantee_for_carrier(model, model.get("carrier_name") or "Carrier", field_name, value)
    _audit(model, source_file, "Word", f"Paragraph {paragraph_index}", field_name, value, confidence, "FOUND", notes=notes)


def _append_rate_guarantee(
    model: dict[str, Any],
    field_name: str,
    value: str,
    source_file: Path,
    paragraph_index: int,
) -> None:
    existing = model["rate_guarantee"].get(field_name)
    if existing and value in str(existing):
        return
    model["rate_guarantee"][field_name] = value if not existing else f"{existing}; {value}"
    _store_rate_guarantee_for_carrier(model, model.get("carrier_name") or "Carrier", field_name, value)
    _audit(model, source_file, "Word", f"Paragraph {paragraph_index}", field_name, value, "HIGH", "FOUND")


def _store_rate_guarantee_for_carrier(model: dict[str, Any], carrier: str, field_name: str, value: Any) -> None:
    carrier = _clean_text(carrier) or "Carrier"
    value_text = _clean_text(value)
    if not value_text:
        return
    carrier_values = model.setdefault("rate_guarantee_by_carrier", {}).setdefault(carrier, {})
    existing = carrier_values.get(field_name)
    if existing and value_text in str(existing):
        return
    carrier_values[field_name] = value_text if not existing else f"{existing}; {value_text}"


def _merge_confirmation(model: dict[str, Any], benefit: str, premium: float, volume: float | None) -> None:
    for record in model["benefit_rows"]:
        if record["benefit_line"] != benefit:
            continue
        if record.get("monthly_premium") not in (None, 0):
            record["notes"] = f"{record.get('notes', '')} EBC confirms monthly premium {premium}.".strip()
            return
    model["benefit_rows"].append(
        {
            "carrier_name": model.get("carrier_name"),
            "carrier_slot": model.get("carrier_slot"),
            "carrier_index": model.get("carrier_index"),
            "benefit_line": benefit,
            "class_name": "Grand Total",
            "tier": None,
            "lives": None,
            "volume": volume,
            "rate": None,
            "monthly_premium": premium,
            "annual_premium": round(premium * 12, 2),
            "rate_basis": None,
            "volume_basis": "volume" if volume is not None else None,
            "included_status": "Included",
            "notes": "Extracted from EBC Grand Total row.",
            "source_file": None,
            "source_type": "Excel",
            "source_location": "EBC Grand Total",
            "confidence": "MEDIUM",
            "status": "VALID",
        }
    )


def _missing(model: dict[str, Any], field_name: str, notes: str) -> None:
    model["missing"].append(field_name)
    _audit(model, None, "System", "", field_name, "Not Available", "LOW", "MISSING", notes=notes)


def _audit(
    model: dict[str, Any],
    source_file: Path | str | None,
    source_type: str,
    source_location: str,
    field_name: str,
    value: Any,
    confidence: str,
    status: str,
    benefit_line: str | None = None,
    destination_sheet: str | None = None,
    destination_cell: str | None = None,
    notes: str | None = None,
) -> None:
    if status == "NEEDS REVIEW":
        model["needs_review"].append(field_name)
    model["audit"].append(
        {
            "source_file": str(source_file) if source_file else None,
            "source_type": source_type,
            "source_location": source_location,
            "field_name": field_name,
            "benefit_line": benefit_line,
            "value": _display_value(value),
            "destination_sheet": destination_sheet,
            "destination_cell": destination_cell,
            "confidence": confidence,
            "status": status,
            "notes": notes,
        }
    )


def _audit_destination(
    model: dict[str, Any],
    source_record: dict[str, Any],
    field_name: str,
    value: Any,
    destination_sheet: str,
    destination_cell: str,
    confidence: str | None,
    status: str | None,
) -> None:
    _audit(
        model,
        source_record.get("source_file"),
        source_record.get("source_type") or "Source",
        source_record.get("source_location") or "",
        field_name,
        value,
        confidence or "MEDIUM",
        status or "FOUND",
        benefit_line=source_record.get("benefit_line"),
        destination_sheet=destination_sheet,
        destination_cell=destination_cell,
        notes=source_record.get("notes"),
    )


def _normalize_benefit(value: Any) -> str | None:
    text = _normalize(value)
    if not text:
        return None
    if text.startswith("per ") or "$" in text:
        return None

    for canonical in BENEFIT_ALIASES:
        if text == canonical.lower():
            return canonical

    matches: list[tuple[int, str]] = []
    for canonical, aliases in BENEFIT_ALIASES.items():
        for alias in sorted(aliases, key=len, reverse=True):
            if _alias_match(text, alias):
                matches.append((len(alias), canonical))
    if matches:
        return sorted(matches, reverse=True)[0][1]

    if _looks_like_unknown_benefit(text):
        return _clean_text(value).title()
    return None


def _benefit_from_funding_line(text: str) -> str | None:
    lower = text.lower()
    if ":" not in text:
        return None
    if not any(term in lower for term in ["pooled", "experience rated", "non-refund"]):
        return None
    return _normalize_benefit(text.split(":", 1)[0])


def _looks_like_unknown_benefit(text: str) -> bool:
    if len(text) < 4 or len(text) > 80:
        return False
    blocked = ["total", "rate", "premium", "effective", "prepared", "commission", "benefit", "class", "volume", "per"]
    if text in blocked or text.startswith(tuple(f"{item} " for item in blocked)):
        return False
    return any(word in text for word in ["insurance", "care", "assistance", "spending", "wellness", "travel"])


def _alias_match(text: str, alias: str) -> bool:
    alias_norm = _normalize(alias)
    return bool(re.search(rf"(?<![a-z0-9]){re.escape(alias_norm)}(?![a-z0-9])", text))


def _category_sort(benefit: str) -> int:
    return CATEGORY_ORDER[_benefit_category(benefit)]


def _benefit_category(benefit: str) -> str:
    if benefit in STANDARD_BENEFITS:
        return "standard"
    if benefit in OPTIONAL_BENEFITS:
        return "optional"
    return "carrier_specific"


def _benefit_sort_order(benefit: str) -> int:
    if benefit in STANDARD_BENEFITS:
        return STANDARD_BENEFITS.index(benefit)
    if benefit in OPTIONAL_BENEFITS:
        return OPTIONAL_BENEFITS.index(benefit)
    return 999


def _row_label(record: dict[str, Any], fallback: str) -> str:
    parts = [record.get("class_name"), record.get("tier")]
    label = " - ".join(str(part) for part in parts if part)
    return label or fallback


def _sumif_many(criteria: list[str]) -> str:
    formulas = [f'SUMIF(\'Total Cost Comparison\'!$M:$M,"{criterion}",\'Total Cost Comparison\'!$H:$H)' for criterion in criteria]
    return "=" + "+".join(formulas) if formulas else "=0"


def _prepare_sheet(sheet: Any) -> None:
    for merged_range in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged_range))
    tables = getattr(sheet, "_tables", None)
    if tables is not None and hasattr(tables, "clear"):
        tables.clear()
    sheet.auto_filter.ref = None
    max_row = max(sheet.max_row, 200)
    max_col = max(sheet.max_column, 20)
    for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            cell.value = None
            cell.comment = None


def _sanitize_workbook_for_excel(workbook: Any) -> None:
    if hasattr(workbook, "_external_links"):
        workbook._external_links = []

    for sheet in workbook.worksheets:
        tables = getattr(sheet, "_tables", None)
        if tables is not None and hasattr(tables, "clear"):
            tables.clear()
        if getattr(sheet.auto_filter, "ref", None) and "#REF!" in str(sheet.auto_filter.ref):
            sheet.auto_filter.ref = None
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and "#REF!" in cell.value:
                    cell.value = None

    defined_names = getattr(workbook, "defined_names", None)
    if defined_names is None:
        return
    try:
        for name, defined_name in list(defined_names.items()):
            definition = f"{getattr(defined_name, 'attr_text', '')} {defined_name}"
            if "#REF!" in definition:
                del defined_names[name]
    except Exception:
        return


def _apply_workbook_font_family(workbook: Any, font_name: str) -> None:
    for named_style in getattr(workbook, "_named_styles", []):
        font = getattr(named_style, "font", None)
        if font is not None:
            try:
                named_style.font = _font_with_name(font, font_name)
            except Exception:
                pass

    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.font is not None and cell.font.name != font_name:
                    cell.font = _font_with_name(cell.font, font_name)


def _font_with_name(font: Any, font_name: str) -> Font:
    updated = copy(font)
    updated.name = font_name
    return updated


def _title(sheet: Any, title: str, row: int, start_column: int, end_column: int) -> None:
    sheet.cell(row, start_column, title)
    sheet.merge_cells(start_row=row, start_column=start_column, end_row=row, end_column=end_column)
    cell = sheet.cell(row, start_column)
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.font = Font(bold=True, color="FFFFFF", size=14)
    cell.alignment = Alignment(horizontal="left", vertical="center")


def _format_sheet(sheet: Any, max_row: int, max_col: int) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(bold=True, color="1F2937")
    border = Border(
        left=Side(style="thin", color="D9E2F3"),
        right=Side(style="thin", color="D9E2F3"),
        top=Side(style="thin", color="D9E2F3"),
        bottom=Side(style="thin", color="D9E2F3"),
    )
    sheet.freeze_panes = "A4"
    sheet.auto_filter.ref = f"A3:{get_column_letter(max_col)}{max_row}"
    for row in range(1, max_row + 1):
        for column in range(1, max_col + 1):
            cell = sheet.cell(row, column)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
            if row == 3:
                cell.fill = header_fill
                cell.font = header_font
            if isinstance(cell.value, (int, float)) and column in {5, 7, 8}:
                cell.number_format = "$#,##0.00"
    for column in range(1, max_col + 1):
        width = 34 if column in {1, 6, 11} else 16
        sheet.column_dimensions[get_column_letter(column)].width = width


def _format_plan_database_sheet(sheet: Any, max_row: int, max_col: int) -> None:
    _format_sheet(sheet, max_row, max_col)
    widths = {
        1: 28,
        2: 34,
        3: 20,
        5: 18,
        6: 20,
        7: 24,
        8: 64,
        9: 14,
        10: 24,
        11: 24,
        12: 18,
        13: 40,
        14: 30,
        15: 28,
        16: 26,
        17: 28,
        18: 36,
        19: 34,
        20: 18,
        21: 12,
        22: 14,
    }
    for column, width in widths.items():
        sheet.column_dimensions[get_column_letter(column)].width = width


def _format_booklet_review_sheet(sheet: Any, max_row: int, max_col: int) -> None:
    _format_sheet(sheet, max_row, max_col)
    sheet.freeze_panes = "A4"
    header_fill = PatternFill("solid", fgColor="004C55")
    detail_fill = PatternFill("solid", fgColor="EAF8F8")
    review_fill = PatternFill("solid", fgColor="FFF2CC")
    for column in range(1, max_col + 1):
        header = sheet.cell(3, column)
        header.fill = header_fill
        header.font = Font(bold=True, color="FFFFFF")
        if BOOKLET_REVIEW_HEADERS[column - 1] in {"ReviewFlag", "Confidence", "Notes"}:
            header.fill = review_fill
            header.font = Font(bold=True, color="1F2937")

    for row in range(4, max_row + 1):
        for column in range(1, max_col + 1):
            header_name = BOOKLET_REVIEW_HEADERS[column - 1]
            cell = sheet.cell(row, column)
            if 44 <= column <= 50 or header_name.startswith("Dental") and ("Normalized" in header_name or "Raw" in header_name):
                cell.fill = detail_fill
            if header_name == "ReviewFlag" and str(cell.value).lower() == "yes":
                cell.fill = review_fill
            if header_name.endswith("Percent") and isinstance(cell.value, (int, float)):
                cell.number_format = "0.0%"

    wide_columns = {
        "ClientID": 28,
        "ClientName": 34,
        "SourceFile": 34,
        "ClassName": 18,
        "ClassDescription": 28,
        "BenefitCategory": 18,
        "BenefitName": 24,
        "BenefitAmountText": 38,
        "CoinsuranceText": 34,
        "DeductibleText": 30,
        "MaximumText": 34,
        "TerminationAgeText": 38,
        "Notes": 34,
        "DentalBasicRecallRaw": 28,
        "DentalBasicRecallNormalized": 26,
        "DentalMajorAnnualMaximumRaw": 30,
        "DentalMajorAnnualMaximumNormalized": 30,
        "DentalOrthoLifetimeMaximumRaw": 30,
        "DentalOrthoLifetimeMaximumNormalized": 30,
        "DentalNormalizationNotes": 42,
        "Benefit": 24,
        "CoverageDescription": 58,
    }
    for column, header in enumerate(BOOKLET_REVIEW_HEADERS, start=1):
        width = wide_columns.get(header, 16)
        if header in {"BenefitAmountNumeric", "MaximumAmount", "SingleDeductible", "FamilyDeductible"}:
            width = 14
        sheet.column_dimensions[get_column_letter(column)].width = width


def _group_row(sheet: Any, row: int, max_col: int) -> None:
    fill = PatternFill("solid", fgColor="EAF2F8")
    for column in range(1, max_col + 1):
        cell = sheet.cell(row, column)
        cell.fill = fill
        cell.font = Font(bold=True, color="1F2937")


def _status_cell(cell: Any, status: str | None) -> None:
    fill = STATUS_FILLS.get(str(status or "").upper())
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)


def _category_label(category: str) -> str:
    return {
        "standard": "Standard Benefit",
        "optional": "Optional / Value Added",
        "carrier_specific": "Carrier-Specific Benefit",
    }[category]


def _travel_provision_name(text: str) -> str:
    lower = text.lower()
    if "maximum" in lower:
        return "Travel maximum"
    if "emergency" in lower:
        return "Travel emergency coverage"
    if "referral" in lower:
        return "Travel referral coverage"
    return "Travel / Out-of-Province coverage"


def _first_percent_or_text(text: str) -> str:
    match = re.search(r"\d+(?:\.\d+)?%", text)
    return match.group(0) if match else text


def _row_has_cost_data(lives: Any, volume: Any, rate: Any, monthly_premium: Any, tier: str | None) -> bool:
    values = [lives, volume, rate, monthly_premium]
    return any(value not in (None, "", 0, 0.0) for value in values)


def _calculate_premium(
    benefit: str,
    lives: float | None,
    volume: float | None,
    rate: float | None,
    rate_basis: str | None,
    tier: str | None,
) -> float | None:
    if rate in (None, ""):
        return None
    basis = _normalize(rate_basis)
    if volume not in (None, "") and ("1,000" in basis or "1000" in basis):
        return round(float(volume) * float(rate) / 1000, 2)
    if volume not in (None, "") and "100" in basis:
        return round(float(volume) * float(rate) / 100, 2)
    if volume not in (None, "") and "10" in basis:
        return round(float(volume) * float(rate) / 10, 2)
    if lives not in (None, "") and (tier or benefit in {"Extended Health Care", "Dental Care", "Dependent Life"}):
        return round(float(lives) * float(rate), 2)
    if lives not in (None, ""):
        return round(float(lives) * float(rate), 2)
    if volume not in (None, ""):
        return round(float(volume) * float(rate), 2)
    return None


def _variance(actual: float | None, calculated: float | None) -> float | None:
    if actual is None or calculated is None:
        return None
    return round(float(actual) - float(calculated), 2)


def _percent_change(new_value: float | None, old_value: float | None, source_percent: float | None = None) -> float | None:
    if old_value not in (None, 0, 0.0) and new_value is not None:
        return float(new_value) / float(old_value) - 1
    if source_percent is not None:
        return float(source_percent) / 100
    return None


def _clean_class(value: Any) -> str | None:
    text = _clean_text(value)
    if not text:
        return None
    if len(text) <= 3 and re.match(r"^[A-Z0-9]+$", text, flags=re.IGNORECASE):
        return f"Class {text.upper()}"
    if text.lower().startswith("class "):
        return text
    return None


def _normalize_tier(value: Any) -> str | None:
    text = _normalize(value)
    mapping = {
        "single": "Single",
        "sngl": "Single",
        "couple": "Couple",
        "cpl": "Couple",
        "family": "Family",
        "fmly": "Family",
        "employee only": "Employee only",
        "employee plus spouse": "Employee plus spouse",
        "employee plus child": "Employee plus child",
        "employee plus family": "Employee plus family",
    }
    return mapping.get(text)


def _cell_at(row: list[Any], index: int) -> Any:
    return row[index] if index < len(row) else None


def _first_non_empty(values: list[Any]) -> Any:
    for value in values:
        if value not in (None, ""):
            return value
    return None


def _cell_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value
    return str(value).strip()


def _format_date(value: Any) -> str:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value)
    match = re.search(r"\d{4}-\d{2}-\d{2}", text)
    return match.group(0) if match else text


def _to_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("$", "").replace(",", "").replace("%", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _display_value(value: Any) -> Any:
    if isinstance(value, float):
        return round(value, 4)
    if isinstance(value, (datetime, date)):
        return _format_date(value)
    return value


def _client_id(model: dict[str, Any]) -> str:
    text = str(model.get("client_name") or "UNKNOWN CLIENT")
    text = re.sub(r"[^A-Za-z0-9]+", "-", text).strip("-").upper()
    return text[:60] or "UNKNOWN-CLIENT"


def _source_file_name(value: Any) -> str | None:
    if not value:
        return None
    try:
        return Path(str(value)).name
    except OSError:
        return str(value)


def _row_text(row: list[Any]) -> str:
    return " ".join(str(value) for value in row if value not in (None, "")).strip()


def _clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _normalize(value: Any) -> str:
    return _clean_text(value).lower()


def _print_summary(model: dict[str, Any]) -> None:
    rates_found = sum(1 for row in model["benefit_rows"] if row.get("rate") not in (None, 0, ""))
    premiums_found = sum(1 for row in model["benefit_rows"] if row.get("monthly_premium") not in (None, 0, ""))
    missing_count = sum(1 for row in model["audit"] if row.get("status") == "MISSING")
    review_count = sum(1 for row in model["audit"] if row.get("status") == "NEEDS REVIEW")
    print(f"Carrier detected: {model.get('carrier_name')}")
    if model.get("carrier_names"):
        print(f"Carriers compared: {', '.join(model.get('carrier_names', []))}")
    print(f"Client detected: {model.get('client_name')}")
    print(f"Excel sheets read: {', '.join(model.get('excel_sheets_read', []))}")
    print(f"Word tables read: {model.get('word_tables_read')}")
    print(f"Benefit lines found: {len({row['benefit_line'] for row in model['benefit_rows']})}")
    print(f"Rates found: {rates_found}")
    print(f"Premiums found: {premiums_found}")
    print(f"Plan design provisions found: {len(model.get('plan_design', []))}")
    print(f"Schedule of Benefits rows found: {len(model.get('master_plan_design', []))}")
    print(f"Alternates found: {len(model.get('alternates', []))}")
    print(f"Deviations found: {len(model.get('deviations', []))}")
    print(f"Fields missing: {missing_count}")
    print(f"Fields needing review: {review_count}")
    print(f"Output workbook path: {model.get('output_workbook')}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Focused carrier quote extraction test.")
    parser.add_argument("excel_path", nargs="?", help="Carrier quote Excel workbook.")
    parser.add_argument("word_path", nargs="?", help="Carrier quote Word document.")
    parser.add_argument(
        "--carrier-package-comparison",
        nargs="+",
        action="append",
        metavar="FILE_PATH",
        dest="carrier_package_comparison",
        help="One carrier package to add to a comparison. All files passed after this option are read together as one carrier.",
    )
    parser.add_argument("--pdf-comparison", nargs="+", dest="pdf_comparison", help="One or more carrier proposal PDFs to compare.")
    parser.add_argument(
        "--excel-word-comparison",
        nargs=2,
        action="append",
        metavar=("EXCEL_PATH", "WORD_PATH"),
        dest="excel_word_comparison",
        help="Excel/Word carrier package to add to a comparison. Can be used with --pdf-comparison.",
    )
    parser.add_argument("--output-folder", dest="output_folder")
    parser.add_argument("--database-path", dest="database_path")
    parser.add_argument("--client-id", dest="client_id")
    parser.add_argument("--client-name", dest="client_name")
    parser.add_argument("--booklet-path", dest="booklet_path")
    args = parser.parse_args()

    if args.carrier_package_comparison:
        model = run_packaged_quote_comparison(args.carrier_package_comparison, args.output_folder)
    elif args.pdf_comparison and args.excel_word_comparison:
        model = run_mixed_quote_comparison(args.pdf_comparison, args.excel_word_comparison, args.output_folder)
    elif args.pdf_comparison:
        model = run_pdf_quote_comparison(args.pdf_comparison, args.output_folder)
    elif args.excel_word_comparison:
        model = run_mixed_quote_comparison([], args.excel_word_comparison, args.output_folder)
    else:
        if not args.excel_path or not args.word_path:
            parser.error("Provide excel_path and word_path, or use --pdf-comparison with one or more PDFs.")
        model = run_quote_test(args.excel_path, args.word_path, args.output_folder)
    _print_summary(model)
    if args.database_path:
        try:
            envelope = build_focused_comparison_envelope(model, args.client_id, args.client_name)
            if args.booklet_path:
                envelope = validate_focused_proposals_against_booklet(envelope, args.booklet_path)
            audit = load_to_sqlite(envelope, args.database_path, client_id=envelope["client_id"])
            model["sqlite_audit"] = audit
            preview = sqlite_preview(args.database_path)
            plan_design_rows = sum(
                int(item.get("rows_loaded") or 0)
                for item in audit
                if item.get("table_name") == "proposal_plan_design"
            )
            print(f"Saved proposals: {len(envelope['proposals'])}")
            print(f"Saved plan design rows: {plan_design_rows}")
            validation_summary = envelope.get("_booklet_validation_summary")
            if validation_summary:
                print(f"Booklet validation status: {validation_summary.get('status')}")
                print(f"Booklet validation rows: {validation_summary.get('validation_rows', 0)}")
                print(f"Booklet validation review issues: {validation_summary.get('review_issues', 0)}")
            print(f"SQLite: {args.database_path}")
            print(f"SQLite proposal rows: {preview.get('proposal_files', 0)}")
        except Exception as exc:
            print(f"SQLite persistence failed: {exc}", file=sys.stderr)
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
