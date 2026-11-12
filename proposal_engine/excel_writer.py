from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from proposal_engine import DEFAULT_OUTPUT_DIR, DEFAULT_TEMPLATE_PATH
from proposal_engine.benefit_library import (
    clean_benefit_name,
    contains_error_token,
    load_benefit_library,
    normalize_benefit_name,
)

TARGET_SHEETS = [
    "Total Cost Comparison",
    "Plan Design",
    "Extended Health",
    "Rate Guarantee Sum",
    "ExecSummary - Cost",
]

CATEGORY_ORDER = {
    "standard": 0,
    "optional": 1,
    "carrier_specific": 2,
}

STATUS_FILLS = {
    "MATCH": "D9EAD3",
    "DIFFERENT": "FCE4D6",
    "MISSING": "FCE4D6",
    "NEEDS REVIEW": "FFF2CC",
    "NOT AVAILABLE": "E7E6E6",
    "LOW CONFIDENCE": "EADCF8",
    "CONFLICT": "F4CCCC",
}

CORE_EXEC_SUMMARY = [
    ("Life", ["employee_life"]),
    ("AD&D", ["employee_ad_d"]),
    ("Dep Life", ["dependent_life"]),
    ("LTD", ["long_term_disability"]),
    ("EHC", ["extended_health_care"]),
    ("Dental", ["dental_care"]),
]


def write_to_excel(
    validated_data: dict[str, Any],
    output_path: str | Path | None = None,
    test_mode: bool = False
) -> str:
    output = Path(output_path) if output_path else DEFAULT_OUTPUT_DIR / (
        "Carrier_Extraction_Test_Output.xlsx" if test_mode else "Carrier_Proposal_Comparison.xlsx"
    )
    output.parent.mkdir(parents=True, exist_ok=True)

    workbook = _load_template_workbook()
    _ensure_target_sheets(workbook)
    model = _build_market_survey_model(validated_data)

    stats = {
        "total_cost_rows": _write_total_cost_comparison(workbook["Total Cost Comparison"], model),
        "plan_design_rows": _write_plan_design(workbook["Plan Design"], model),
        "rate_guarantee_rows": _write_rate_guarantee(workbook["Rate Guarantee Sum"], model),
        "extended_health_rows": _write_extended_health(workbook["Extended Health"], model),
        "exec_summary_rows": _write_exec_summary(workbook["ExecSummary - Cost"], model),
        "carrier_count": len(model["carriers"]),
        "benefit_count": len(model["benefits"]),
        "manual_review_items": len(validated_data.get("exceptions", [])),
    }
    validated_data["_market_survey_stats"] = stats

    output = _save_workbook_safely(workbook, output)
    return str(output)


def _load_template_workbook() -> Any:
    template = _resolve_template_path()
    if template and template.exists():
        workbook = load_workbook(template, keep_links=False)
        _strip_external_links(workbook)
        return workbook

    workbook = Workbook()
    workbook.remove(workbook.active)
    for sheet_name in TARGET_SHEETS:
        workbook.create_sheet(sheet_name)
    return workbook


def _resolve_template_path() -> Path | None:
    if DEFAULT_TEMPLATE_PATH.exists():
        return DEFAULT_TEMPLATE_PATH

    return None


def _ensure_target_sheets(workbook: Any) -> None:
    for sheet_name in TARGET_SHEETS:
        if sheet_name not in workbook.sheetnames:
            workbook.create_sheet(sheet_name)


def _strip_external_links(workbook: Any) -> None:
    if hasattr(workbook, "_external_links"):
        workbook._external_links = []


def _save_workbook_safely(workbook: Any, output: Path) -> Path:
    try:
        workbook.save(output)
        return output
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = output.with_name(f"{output.stem}_{timestamp}{output.suffix}")
        workbook.save(fallback)
        return fallback


def _build_market_survey_model(data: dict[str, Any]) -> dict[str, Any]:
    library = load_benefit_library()
    carriers = _carrier_names(data)
    benefit_map = _collect_benefits(data, library)
    benefits = sorted(
        benefit_map.values(),
        key=lambda item: (
            CATEGORY_ORDER.get(item.get("benefit_category"), 9),
            item.get("sort_order") if item.get("sort_order") is not None else 999,
            str(item.get("display_name") or ""),
        ),
    )

    detail_records = [*data.get("benefit_lines", []), *data.get("rates", []), *data.get("premiums", [])]
    row_groups = _build_cost_rows(benefits, detail_records, library)
    plan_rows = _build_plan_rows(benefits, data.get("plan_design", []), library)
    rate_guarantee = _aggregate_rate_guarantee(data.get("pooling_retention", []), carriers)

    return {
        "data": data,
        "carriers": carriers,
        "benefits": benefits,
        "cost_rows": row_groups,
        "plan_rows": plan_rows,
        "rate_guarantee": rate_guarantee,
        "exceptions": data.get("exceptions", []),
    }


def _carrier_names(data: dict[str, Any]) -> list[str]:
    names: list[str] = []
    summaries = data.get("proposal_summaries") or [data.get("proposal_summary", {})]
    for summary in summaries:
        name = summary.get("carrier_name") or summary.get("detected_carrier")
        if not name or str(name).strip().lower() in {"default generic", "default_generic", "unknown"}:
            name = _carrier_from_source(summary.get("source_file"))
        _append_unique(names, str(name or "Unknown Carrier"))

    for collection in ["benefits", "benefit_lines", "rates", "premiums", "plan_design", "pooling_retention"]:
        for record in data.get(collection, []):
            name = record.get("carrier_name") or _carrier_from_source(record.get("source_file"))
            _append_unique(names, str(name or "Unknown Carrier"))

    return names or ["Unknown Carrier"]


def _collect_benefits(data: dict[str, Any], library: dict[str, Any]) -> dict[str, dict[str, Any]]:
    benefit_map: dict[str, dict[str, Any]] = {}

    for record in data.get("benefits", []):
        if _is_noisy_market_record(record):
            continue
        benefit = _normalize_benefit_record(record, library)
        if benefit:
            benefit_map.setdefault(benefit["benefit_key"], benefit)

    for collection in ["benefit_lines", "rates", "premiums", "plan_design"]:
        for record in data.get(collection, []):
            if _is_noisy_market_record(record):
                continue
            benefit = normalize_benefit_name(record.get("benefit_line"), str(record), library)
            if benefit:
                benefit_map.setdefault(benefit["benefit_key"], benefit)

    return benefit_map


def _normalize_benefit_record(record: dict[str, Any], library: dict[str, Any]) -> dict[str, Any] | None:
    name = record.get("display_name") or record.get("normalized_benefit_name") or record.get("benefit_name")
    if _is_support_label(name):
        return None
    benefit = normalize_benefit_name(name, str(record), library)
    if not benefit:
        return None
    benefit["carrier_name"] = record.get("carrier_name")
    benefit["source_file"] = record.get("source_file")
    return benefit


def _is_noisy_market_record(record: dict[str, Any]) -> bool:
    if contains_error_token(record):
        return True
    label = (
        record.get("display_name")
        or record.get("normalized_benefit_name")
        or record.get("benefit_name")
        or record.get("benefit_line")
        or record.get("provision_name")
    )
    return _is_support_label(label)


def _is_support_label(value: Any) -> bool:
    label = _normalize_text(clean_benefit_name(value))
    if not label:
        return True
    exact = {
        "current plan",
        "carrier",
        "rate",
        "premium",
        "monthly premium",
        "annual premium",
        "total",
        "total cost",
        "total premium",
        "difference",
        "percentage difference",
        "percent difference",
        "% difference",
        "avg comp",
        "average comp",
        "estimated total",
        "union annual premiums",
        "benefit amount",
        "waiver of premium",
        "reduction of coverage",
        "termination of coverage",
        "covered illnesses",
        "coverage coinsurance",
        "deductible",
        "deductibles",
        "coinsurance",
        "maximum",
        "maximums",
        "recall frequency",
        "fee guide year",
        "eye exams",
        "spouse",
        "child",
        "child(ren)",
        "alternates",
        "drug max",
        "grouphealth",
        "rbc",
        "union premiums",
        "paramedical practitioners",
        "medical services & supplies",
        "medical services and supplies",
        "plan type",
    }
    prefixes = (
        "total ",
        "difference ",
        "percentage difference",
        "percent difference",
        "% difference",
        "avg comp",
        "average comp",
        "estimated total",
        "comp this year",
    )
    provision_terms = (
        " overall maximum",
        " maximum",
        " elimination period",
        " recall",
        " deductible",
        " coinsurance",
        " fee guide",
        " eye exam",
        " massage therapist",
        " chiropractor",
        " naturopath",
        " physiotherapist",
        " audiologist",
        " dietician",
        " osteopath",
        " podiatrist",
        " chiropodist",
        " occupational therapist",
        " out of province",
        " out-of-country",
        " credit carry forward",
        " benefit amount",
        " admin fee",
        " renewal",
        " do not copy",
        " add ",
        " vision care",
        " drug max",
        " grouphealth",
        " plan type",
        " pay direct",
        " combined with basic",
        " paramedical",
        " practitioner",
        " therapist",
        " psychologist",
        " social worker",
        " clinical counsellor",
        " psychotherapist",
        " taxability",
        " taxabitlity",
        " union premiums",
        " premiums",
    )
    return (
        label in exact
        or label.startswith(prefixes)
        or label[:1].isdigit()
        or label.startswith("$")
        or any(term in f" {label}" for term in provision_terms)
    )


def _clean_market_label(value: Any, fallback: str) -> str:
    label = clean_benefit_name(value)
    if _is_support_label(label):
        return fallback
    return label[:120] if label else fallback


def _build_cost_rows(
    benefits: list[dict[str, Any]],
    detail_records: list[dict[str, Any]],
    library: dict[str, Any]
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for record in detail_records:
        if _is_noisy_market_record(record):
            continue
        benefit = normalize_benefit_name(record.get("benefit_line"), str(record), library)
        if benefit:
            grouped[benefit["benefit_key"]].append(record)

    for benefit in benefits:
        benefit_key = benefit["benefit_key"]
        records = grouped.get(benefit_key, [])
        row_keys = _unique_cost_row_keys(records)
        if not row_keys:
            row_keys = [("", "")]

        for class_name, tier in row_keys:
            row_records = [
                record
                for record in records
                if _same_text(record.get("class_name"), class_name) and _same_text(record.get("tier"), tier)
            ]
            if not row_records and class_name == "" and tier == "":
                row_records = records

            rows.append(
                {
                    "benefit": benefit,
                    "class_name": class_name or None,
                    "tier": tier or None,
                    "records": row_records,
                    "values_by_carrier": _carrier_cost_values(row_records),
                }
            )

    return rows


def _unique_cost_row_keys(records: list[dict[str, Any]]) -> list[tuple[str, str]]:
    keys: list[tuple[str, str]] = []
    for record in records:
        class_name = str(record.get("class_name") or "").strip()
        tier = str(record.get("tier") or "").strip()
        key = (class_name, tier)
        if key not in keys and (class_name or tier):
            keys.append(key)
    return keys


def _carrier_cost_values(records: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    values: dict[str, dict[str, Any]] = {}
    for record in records:
        carrier = str(record.get("carrier_name") or _carrier_from_source(record.get("source_file")) or "Unknown Carrier")
        current = values.setdefault(
            carrier,
            {
                "lives": None,
                "volume": None,
                "rate": None,
                "monthly_premium": None,
                "annual_premium": None,
                "status": None,
                "notes": [],
            },
        )
        for field in ["lives", "volume", "rate", "monthly_premium", "annual_premium"]:
            if current[field] is None and record.get(field) not in (None, ""):
                current[field] = record.get(field)
        if record.get("status"):
            current["status"] = record.get("status")
        if record.get("notes"):
            current["notes"].append(record.get("notes"))
    return values


def _build_plan_rows(
    benefits: list[dict[str, Any]],
    records: list[dict[str, Any]],
    library: dict[str, Any]
) -> list[dict[str, Any]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        if _is_noisy_market_record(record):
            continue
        benefit = normalize_benefit_name(record.get("benefit_line"), str(record), library)
        if benefit:
            grouped[benefit["benefit_key"]].append(record)

    rows: list[dict[str, Any]] = []
    for benefit in benefits:
        benefit_records = grouped.get(benefit["benefit_key"], [])
        provisions = _unique_values(record.get("provision_name") for record in benefit_records)
        if not provisions:
            provisions = ["Included / Availability"]

        for provision in provisions:
            matching = [record for record in benefit_records if _same_text(record.get("provision_name"), provision)]
            rows.append(
                {
                    "benefit": benefit,
                    "provision": provision,
                    "values_by_carrier": _carrier_plan_values(matching),
                    "current_plan": _first_non_empty(record.get("booklet_value") for record in matching),
                    "validation_status": _combined_validation_status(matching),
                }
            )
    return rows


def _carrier_plan_values(records: list[dict[str, Any]]) -> dict[str, Any]:
    values: dict[str, Any] = {}
    for record in records:
        carrier = str(record.get("carrier_name") or _carrier_from_source(record.get("source_file")) or "Unknown Carrier")
        value = record.get("provision_value") or record.get("status") or "NEEDS REVIEW"
        values.setdefault(carrier, value)
    return values


def _aggregate_rate_guarantee(records: list[dict[str, Any]], carriers: list[str]) -> dict[str, dict[str, Any]]:
    fields = [
        "rate_guarantee",
        "pooled_benefits",
        "experience_rated_benefits",
        "pooling_level",
        "pooling_charge",
        "target_loss_ratio",
        "ibnr",
        "rate_caps",
        "first_renewal_cap",
        "second_renewal_cap",
    ]
    result = {carrier: {field: None for field in fields} for carrier in carriers}
    for record in records:
        carrier = str(record.get("carrier_name") or _carrier_from_source(record.get("source_file")) or "Unknown Carrier")
        current = result.setdefault(carrier, {field: None for field in fields})
        for field in fields:
            if current[field] is None and record.get(field) not in (None, ""):
                current[field] = record.get(field)
    return result


def _write_total_cost_comparison(sheet: Any, model: dict[str, Any]) -> int:
    _prepare_sheet(sheet)
    carriers = model["carriers"]
    header_columns = _carrier_columns(carriers, start_col=6)

    _write_title(sheet, "TOTAL COST COMPARISON", 1, 1, max(7, 5 + len(carriers) * 2))
    sheet.cell(3, 1, "Benefit / Class / Tier")
    sheet.cell(3, 2, "Lives")
    sheet.cell(3, 3, "Volume")
    sheet.cell(2, 4, "Current Plan")
    sheet.merge_cells(start_row=2, start_column=4, end_row=2, end_column=5)
    sheet.cell(3, 4, "Rate")
    sheet.cell(3, 5, "Premium")

    for carrier, rate_col, premium_col in header_columns:
        sheet.cell(2, rate_col, carrier)
        sheet.merge_cells(start_row=2, start_column=rate_col, end_row=2, end_column=premium_col)
        sheet.cell(3, rate_col, "Rate")
        sheet.cell(3, premium_col, "Premium")

    row = 4
    stats_rows = 0
    last_benefit_key = None
    for item in model["cost_rows"]:
        benefit = item["benefit"]
        if benefit["benefit_key"] != last_benefit_key:
            sheet.cell(row, 1, benefit["display_name"])
            sheet.cell(row, 2, _category_label(benefit.get("benefit_category")))
            _style_group_row(sheet, row, max(7, 5 + len(carriers) * 2))
            row += 1
            last_benefit_key = benefit["benefit_key"]

        label = _clean_market_label(item.get("tier") or item.get("class_name") or benefit["display_name"], benefit["display_name"])
        sheet.cell(row, 1, label)
        sheet.cell(row, 2, _first_non_empty(value.get("lives") for value in item["values_by_carrier"].values()))
        sheet.cell(row, 3, _first_non_empty(value.get("volume") for value in item["values_by_carrier"].values()))
        sheet.cell(row, 4, None)
        sheet.cell(row, 5, None)

        for carrier, rate_col, premium_col in header_columns:
            value = item["values_by_carrier"].get(carrier)
            if value:
                sheet.cell(row, rate_col, value.get("rate"))
                sheet.cell(row, premium_col, _premium_value(value))
                _apply_value_status(sheet.cell(row, premium_col), value.get("status"))
            else:
                sheet.cell(row, rate_col, None)
                sheet.cell(row, premium_col, "Not Available")
                _apply_value_status(sheet.cell(row, premium_col), "NOT AVAILABLE")
        row += 1
        stats_rows += 1

    _format_market_sheet(sheet, max_row=max(row - 1, 3), max_col=max(7, 5 + len(carriers) * 2))
    return stats_rows


def _write_plan_design(sheet: Any, model: dict[str, Any]) -> int:
    _prepare_sheet(sheet)
    carriers = model["carriers"]
    status_col = 4 + len(carriers)
    _write_title(sheet, "PLAN DESIGN", 1, 1, status_col)
    headers = ["Benefit", "Provision", "Current Plan", *carriers, "Validation Status"]
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    row = 4
    current_benefit_key = None
    for item in model["plan_rows"]:
        benefit = item["benefit"]
        if benefit["benefit_key"] != current_benefit_key:
            sheet.cell(row, 1, benefit["display_name"])
            sheet.cell(row, 2, _category_label(benefit.get("benefit_category")))
            _style_group_row(sheet, row, status_col)
            row += 1
            current_benefit_key = benefit["benefit_key"]

        sheet.cell(row, 1, benefit["display_name"])
        sheet.cell(row, 2, item["provision"])
        sheet.cell(row, 3, item.get("current_plan"))
        for offset, carrier in enumerate(carriers, start=4):
            sheet.cell(row, offset, item["values_by_carrier"].get(carrier, "Not Available"))
        status = item.get("validation_status") or "NEEDS REVIEW"
        sheet.cell(row, status_col, status)
        _apply_value_status(sheet.cell(row, status_col), status)
        row += 1

    _format_market_sheet(sheet, max_row=max(row - 1, 3), max_col=status_col)
    return max(0, row - 4)


def _write_rate_guarantee(sheet: Any, model: dict[str, Any]) -> int:
    _prepare_sheet(sheet)
    carriers = model["carriers"]
    max_col = 2 + len(carriers)
    _write_title(sheet, "RATE GUARANTEE SUMMARY", 1, 1, max_col)
    headers = ["", "Current Plan", *carriers]
    for column, header in enumerate(headers, start=1):
        sheet.cell(2, column, header)

    rows = [
        ("Rate Guarantee", "rate_guarantee"),
        ("Pooled benefits", "pooled_benefits"),
        ("Experience-rated benefits", "experience_rated_benefits"),
        ("Pooling level", "pooling_level"),
        ("Pooling charge", "pooling_charge"),
        ("Target loss ratio", "target_loss_ratio"),
        ("IBNR", "ibnr"),
        ("Rate caps", "rate_caps"),
        ("First renewal cap", "first_renewal_cap"),
        ("Second renewal cap", "second_renewal_cap"),
    ]

    for row_offset, (label, field) in enumerate(rows, start=3):
        sheet.cell(row_offset, 1, label)
        sheet.cell(row_offset, 2, None)
        for offset, carrier in enumerate(carriers, start=3):
            value = model["rate_guarantee"].get(carrier, {}).get(field)
            sheet.cell(row_offset, offset, value if value not in (None, "") else "Not Available")
            if value in (None, ""):
                _apply_value_status(sheet.cell(row_offset, offset), "NOT AVAILABLE")

    _format_market_sheet(sheet, max_row=2 + len(rows), max_col=max_col)
    return len(rows)


def _write_extended_health(sheet: Any, model: dict[str, Any]) -> int:
    _prepare_sheet(sheet)
    _write_title(sheet, "EXTENDED HEALTH AND DENTAL CLAIMS EXPERIENCE", 1, 1, 5)
    headers = ["Period", "Paid Premium", "Paid Claims", "Paid Loss Ratio (%)", "Status"]
    for column, header in enumerate(headers, start=1):
        sheet.cell(3, column, header)

    sheet.cell(4, 1, "Claims experience")
    sheet.cell(4, 5, "Not Available")
    _apply_value_status(sheet.cell(4, 5), "NOT AVAILABLE")
    sheet.cell(5, 1, "Dental claims experience")
    sheet.cell(5, 5, "Not Available")
    _apply_value_status(sheet.cell(5, 5), "NOT AVAILABLE")
    _format_market_sheet(sheet, max_row=5, max_col=5)
    return 0


def _write_exec_summary(sheet: Any, model: dict[str, Any]) -> int:
    _prepare_sheet(sheet)
    carriers = model["carriers"]
    max_col = 4 + len(carriers)
    _write_title(sheet, "EXECUTIVE SUMMARY - COST", 1, 1, max_col)
    headers = ["", "", "Benefit", "Current Plan", *carriers]
    for column, header in enumerate(headers, start=1):
        sheet.cell(2, column, header)

    total_cost_rows = _total_cost_row_lookup(model["cost_rows"])
    premium_columns = _premium_columns_for_summary(carriers, start_col=6)
    row = 3
    for label, benefit_keys in CORE_EXEC_SUMMARY:
        sheet.cell(row, 3, label)
        sheet.cell(row, 4, _sum_formula("Total Cost Comparison", 5, total_cost_rows, benefit_keys))
        for offset, carrier in enumerate(carriers, start=5):
            premium_col = premium_columns[carrier]
            sheet.cell(row, offset, _sum_formula("Total Cost Comparison", premium_col, total_cost_rows, benefit_keys))
        row += 1

    optional_keys = [
        item["benefit"]["benefit_key"]
        for item in model["cost_rows"]
        if item["benefit"].get("benefit_category") in {"optional", "carrier_specific"}
    ]
    if optional_keys:
        sheet.cell(row, 3, "Optional / Carrier-Specific")
        sheet.cell(row, 4, _sum_formula("Total Cost Comparison", 5, total_cost_rows, optional_keys))
        for offset, carrier in enumerate(carriers, start=5):
            premium_col = premium_columns[carrier]
            sheet.cell(row, offset, _sum_formula("Total Cost Comparison", premium_col, total_cost_rows, optional_keys))
        row += 1

    total_row = row
    sheet.cell(total_row, 3, "Total")
    for column in range(4, max_col + 1):
        sheet.cell(total_row, column, f"=SUM({get_column_letter(column)}3:{get_column_letter(column)}{total_row - 1})")
    row += 1

    percent_row = row
    sheet.cell(percent_row, 3, "% Difference")
    for column in range(5, max_col + 1):
        sheet.cell(percent_row, column, f"=IF($D${total_row}=0,\"\",{get_column_letter(column)}{total_row}/$D${total_row}-1)")
    row += 1

    guarantee_row = row
    sheet.cell(guarantee_row, 3, "Rate Guarantee")
    for offset, carrier in enumerate(carriers, start=5):
        rate_guarantee_col = 3 + carriers.index(carrier)
        sheet.cell(guarantee_row, offset, f"='Rate Guarantee Sum'!{get_column_letter(rate_guarantee_col)}3")

    _format_market_sheet(sheet, max_row=guarantee_row, max_col=max_col)
    for column in range(4, max_col + 1):
        for row_index in range(3, total_row + 1):
            sheet.cell(row_index, column).number_format = "$#,##0.00"
        sheet.cell(percent_row, column).number_format = "0.0%"
    return guarantee_row - 2


def _prepare_sheet(sheet: Any) -> None:
    for merged_range in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged_range))
    _clear_sheet_tables(sheet)
    sheet.auto_filter.ref = None
    max_row = max(sheet.max_row, 120)
    max_col = max(sheet.max_column, 20)
    for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            cell.value = None
            cell.comment = None


def _clear_sheet_tables(sheet: Any) -> None:
    tables = getattr(sheet, "_tables", None)
    if tables is None:
        return
    if hasattr(tables, "clear"):
        tables.clear()
    else:
        sheet._tables = {}


def _write_title(sheet: Any, title: str, row: int, start_col: int, end_col: int) -> None:
    sheet.cell(row, start_col, title)
    sheet.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    cell = sheet.cell(row, start_col)
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.font = Font(bold=True, size=14, color="FFFFFF")
    cell.alignment = Alignment(horizontal="left", vertical="center")


def _format_market_sheet(sheet: Any, max_row: int, max_col: int) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(bold=True, color="1F2937")
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    sheet.freeze_panes = "A4"
    sheet.auto_filter.ref = f"A3:{get_column_letter(max_col)}{max_row}"

    for row in range(1, max_row + 1):
        for column in range(1, max_col + 1):
            cell = sheet.cell(row, column)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
            if row in {2, 3}:
                cell.fill = header_fill
                cell.font = header_font
            if isinstance(cell.value, (int, float)) and column >= 4:
                cell.number_format = "$#,##0.00" if column % 2 == 1 else "0.0000"

    widths = {
        1: 34,
        2: 12,
        3: 14,
    }
    for column in range(1, max_col + 1):
        width = widths.get(column, 15)
        sheet.column_dimensions[get_column_letter(column)].width = width


def _style_group_row(sheet: Any, row: int, max_col: int) -> None:
    fill = PatternFill("solid", fgColor="EAF2F8")
    for column in range(1, max_col + 1):
        cell = sheet.cell(row, column)
        cell.fill = fill
        cell.font = Font(bold=True, color="1F2937")


def _apply_value_status(cell: Any, status: str | None) -> None:
    if not status:
        return
    fill = STATUS_FILLS.get(str(status).upper())
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)


def _carrier_columns(carriers: list[str], start_col: int) -> list[tuple[str, int, int]]:
    columns = []
    for index, carrier in enumerate(carriers):
        rate_col = start_col + index * 2
        columns.append((carrier, rate_col, rate_col + 1))
    return columns


def _premium_columns_for_summary(carriers: list[str], start_col: int) -> dict[str, int]:
    return {carrier: start_col + index * 2 + 1 for index, carrier in enumerate(carriers)}


def _total_cost_row_lookup(cost_rows: list[dict[str, Any]]) -> list[tuple[str, int]]:
    lookup = []
    row_number = 4
    last_benefit_key = None
    for item in cost_rows:
        benefit_key = item["benefit"]["benefit_key"]
        if benefit_key != last_benefit_key:
            row_number += 1
            last_benefit_key = benefit_key
        lookup.append((benefit_key, row_number))
        row_number += 1
    return lookup


def _sum_formula(sheet_name: str, column: int, row_lookup: list[tuple[str, int]], benefit_keys: list[str]) -> str:
    rows = [row for benefit_key, row in row_lookup if benefit_key in set(benefit_keys)]
    if not rows:
        return "=0"
    col = get_column_letter(column)
    refs = [f"'{sheet_name}'!{col}{row}" for row in rows]
    return "=SUM(" + ",".join(refs) + ")"


def _premium_value(value: dict[str, Any]) -> Any:
    if value.get("monthly_premium") not in (None, ""):
        return value.get("monthly_premium")
    if value.get("annual_premium") not in (None, ""):
        return value.get("annual_premium")
    status = str(value.get("status") or "").upper()
    if status in {"DECLINED", "NOT AVAILABLE", "MISSING"}:
        return status.title()
    notes = " ".join(str(note) for note in value.get("notes", []))
    if "included" in notes.lower():
        return "Included"
    return "Needs Review"


def _combined_validation_status(records: list[dict[str, Any]]) -> str:
    statuses = [str(record.get("validation_status") or record.get("status") or "").upper() for record in records]
    for status in ["DIFFERENT", "MISSING", "NEEDS REVIEW", "MATCH", "NOT AVAILABLE"]:
        if status in statuses:
            return status
    return "NOT AVAILABLE"


def _category_label(category: str | None) -> str:
    return {
        "standard": "Standard Benefit",
        "optional": "Optional Benefit",
        "carrier_specific": "Carrier-Specific Benefit",
    }.get(category or "", "Benefit")


def _carrier_from_source(source_file: Any) -> str:
    if not source_file:
        return "Unknown Carrier"
    stem = Path(str(source_file)).stem
    return stem[:80] if stem else "Unknown Carrier"


def _append_unique(items: list[str], value: str) -> None:
    if not value:
        return
    if not any(existing.lower() == value.lower() for existing in items):
        items.append(value)


def _unique_values(values: Any) -> list[str]:
    result: list[str] = []
    for value in values:
        text = str(value or "").strip()
        if text and not any(existing.lower() == text.lower() for existing in result):
            result.append(text)
    return result


def _first_non_empty(values: Any) -> Any:
    for value in values:
        if value not in (None, ""):
            return value
    return None


def _same_text(left: Any, right: Any) -> bool:
    return str(left or "").strip().lower() == str(right or "").strip().lower()


def _normalize_text(value: Any) -> str:
    return " ".join(str(value or "").strip().lower().split())
