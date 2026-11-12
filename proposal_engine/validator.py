from __future__ import annotations

from collections import defaultdict
from typing import Any


VALID_STATUSES = {
    "VALID",
    "NEEDS REVIEW",
    "MISSING",
    "CONFLICT",
    "LOW CONFIDENCE",
    "NOT AVAILABLE",
}


def validate_proposal_data(normalized_data: dict[str, Any]) -> dict[str, Any]:
    """Validate normalized proposal data and add exception records instead of failing."""

    data = normalized_data
    summary = data.get("proposal_summary", {})
    carrier_name = summary.get("carrier_name")
    source_file = summary.get("source_file")

    if not carrier_name:
        _add_exception(data, "MISSING", "carrier_name", "Carrier name was not found.", source_file)
    if not source_file:
        _add_exception(data, "MISSING", "source_file", "Source file was not captured.", source_file)
    if summary.get("detected_carrier") == "default_generic":
        _add_exception(data, "NEEDS REVIEW", "detected_carrier", "Carrier was not confidently matched to a config.", source_file)

    if not data.get("benefit_lines"):
        _add_exception(data, "MISSING", "benefit_lines", "No benefit lines were extracted.", source_file)
    if not data.get("benefits"):
        _add_exception(data, "MISSING", "benefits", "No benefit records were found.", source_file)
    if not data.get("rates"):
        _add_exception(data, "MISSING", "rates", "No rates were extracted.", source_file)
    if not data.get("plan_design"):
        _add_exception(data, "MISSING", "plan_design", "No plan design rows were extracted.", source_file)

    _validate_premiums(data)
    _flag_low_confidence(data)
    _flag_conflicts(data)
    _normalize_statuses(data)
    return data


def _validate_premiums(data: dict[str, Any]) -> None:
    for record in data.get("premiums", []):
        monthly = _number(record.get("monthly_premium"))
        annual = _number(record.get("annual_premium"))
        if monthly is None or annual is None:
            continue
        expected = round(monthly * 12, 2)
        if abs(expected - annual) > 1:
            record["status"] = "CONFLICT"
            record["notes"] = _append_note(record.get("notes"), f"Annual premium does not equal monthly x 12. Expected {expected}.")
            _add_exception(
                data,
                "CONFLICT",
                "annual_premium",
                "Annual premium does not equal monthly premium x 12.",
                record.get("source_file"),
                record.get("source_page"),
                record.get("source_sheet"),
            )


def _flag_low_confidence(data: dict[str, Any]) -> None:
    for key in [
        "benefit_lines",
        "benefits",
        "rates",
        "premiums",
        "plan_design",
        "demographics",
        "underwriting_assumptions",
        "pooling_retention",
    ]:
        for record in data.get(key, []):
            if str(record.get("confidence_score", "")).upper() == "LOW":
                record["status"] = "LOW CONFIDENCE"
                _add_exception(
                    data,
                    "LOW CONFIDENCE",
                    key,
                    f"Low confidence extraction in {key}.",
                    record.get("source_file"),
                    record.get("source_page"),
                    record.get("source_sheet"),
                )


def _flag_conflicts(data: dict[str, Any]) -> None:
    seen: dict[tuple[str, str, str], set[str]] = defaultdict(set)
    for record in data.get("rates", []):
        key = (
            str(record.get("benefit_line") or ""),
            str(record.get("class_name") or ""),
            str(record.get("tier") or ""),
        )
        if record.get("rate") is not None:
            seen[key].add(str(record.get("rate")))

    for key, values in seen.items():
        if len(values) > 1:
            _add_exception(
                data,
                "CONFLICT",
                "rate",
                f"Conflicting rates found for benefit/class/tier: {' / '.join(key)}.",
                data.get("proposal_summary", {}).get("source_file"),
            )


def _normalize_statuses(data: dict[str, Any]) -> None:
    for key, value in data.items():
        if not isinstance(value, list):
            continue
        for record in value:
            if not isinstance(record, dict):
                continue
            status = str(record.get("status") or "NEEDS REVIEW").upper()
            record["status"] = status if status in VALID_STATUSES else "NEEDS REVIEW"


def _add_exception(
    data: dict[str, Any],
    status: str,
    field_name: str,
    message: str,
    source_file: str | None,
    source_page: Any = None,
    source_sheet: Any = None,
) -> None:
    data.setdefault("exceptions", []).append(
        {
            "carrier_name": data.get("proposal_summary", {}).get("carrier_name"),
            "exception_type": status,
            "field_name": field_name,
            "message": message,
            "source_file": source_file,
            "source_page": source_page,
            "source_sheet": source_sheet,
            "confidence_score": "LOW",
            "status": status,
            "notes": "Validation flag.",
        }
    )


def _number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _append_note(existing: str | None, addition: str) -> str:
    return f"{existing} {addition}".strip() if existing else addition
