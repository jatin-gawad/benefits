from __future__ import annotations

import hashlib
import re
from pathlib import Path
from typing import Any


def build_focused_comparison_envelope(
    model: dict[str, Any],
    client_id: str | None = None,
    client_name: str | None = None,
) -> dict[str, Any]:
    """Map the focused workbook model to the existing proposal persistence contract."""

    display_client_name = _first_text(client_name, model.get("client_name"), "Unknown Client")
    effective_client_id = _first_text(client_id, _client_key(display_client_name))
    carrier_slots = _carrier_slots(model)
    proposals = [
        _build_proposal(model, effective_client_id, display_client_name, carrier_index, carrier_name, len(carrier_slots))
        for carrier_index, carrier_name in carrier_slots
    ]
    return {
        "client_id": effective_client_id,
        "client_name": display_client_name,
        "proposals": proposals,
    }


def _build_proposal(
    model: dict[str, Any],
    client_id: str,
    client_name: str,
    carrier_index: int,
    carrier_name: str,
    carrier_count: int,
) -> dict[str, Any]:
    benefit_rows = [
        row
        for row in model.get("benefit_rows", [])
        if _belongs_to_carrier(row, carrier_index, carrier_name, carrier_count)
    ]
    plan_rows = [
        row
        for row in model.get("plan_design", [])
        if _belongs_to_carrier(row, carrier_index, carrier_name, carrier_count)
    ]
    source_files = _source_files(model, benefit_rows, plan_rows, carrier_index, carrier_name, carrier_count)
    primary_source = source_files[0] if source_files else ""
    package_hash = _package_hash(source_files or [f"{client_id}|{carrier_name}"])
    has_quote_values = bool(plan_rows) or any(_has_rate_or_premium(row) for row in benefit_rows)

    proposal = {
        "proposal_summary": {
            "client_name": client_name,
            "carrier_name": carrier_name,
            "detected_carrier": _carrier_key(carrier_name),
            "detection_confidence": "MEDIUM",
            "proposal_date": model.get("proposal_date"),
            "effective_date": model.get("effective_date"),
            "source_file": primary_source,
            "source_files": source_files,
            "source_file_name": Path(primary_source).name if primary_source else f"{carrier_name} proposal",
            "source_file_hash": package_hash,
            "status": "VALID" if has_quote_values else "NEEDS REVIEW",
            "extraction_method": "Focused Carrier Comparison",
        },
        "benefits": _benefits(benefit_rows),
        "benefit_lines": [_benefit_line(row) for row in benefit_rows],
        "rates": [_rate(row) for row in benefit_rows if _has_rate_data(row)],
        "premiums": [_premium(row) for row in benefit_rows if _has_premium_data(row)],
        "plan_design": [_plan_design(row) for row in plan_rows],
        "booklet_validation": [],
        "demographics": [_demographic(row) for row in benefit_rows if _has_demographic_data(row)],
        "underwriting_assumptions": [],
        "pooling_retention": _rate_guarantee_rows(model, carrier_name),
        "exceptions": _exceptions(model, source_files),
    }
    return proposal


def _carrier_slots(model: dict[str, Any]) -> list[tuple[int, str]]:
    slots: dict[int, str] = {}
    for index, carrier in enumerate(model.get("carrier_names") or [], start=1):
        name = _clean_text(carrier)
        if name:
            slots[index] = name

    for collection_name in ("benefit_rows", "plan_design", "included_services"):
        for row in model.get(collection_name, []):
            name = _clean_text(row.get("carrier_name"))
            index = _positive_int(row.get("carrier_index"))
            if name and index:
                slots.setdefault(index, name)

    if not slots:
        carrier_name = _first_text(model.get("carrier_name"), "Unknown Carrier")
        slots[1] = carrier_name

    return sorted(slots.items())


def _belongs_to_carrier(
    row: dict[str, Any],
    carrier_index: int,
    carrier_name: str,
    carrier_count: int,
) -> bool:
    row_index = _positive_int(row.get("carrier_index"))
    if row_index:
        return row_index == carrier_index

    row_carrier = _clean_text(row.get("carrier_name"))
    if row_carrier:
        return row_carrier.casefold() == carrier_name.casefold()

    return carrier_count == 1


def _source_files(
    model: dict[str, Any],
    benefit_rows: list[dict[str, Any]],
    plan_rows: list[dict[str, Any]],
    carrier_index: int,
    carrier_name: str,
    carrier_count: int,
) -> list[str]:
    package_sources = _unique_paths(
        (model.get("carrier_source_packages") or {}).get(str(carrier_index)) or []
    )
    if package_sources:
        return package_sources

    selected = _unique_paths(
        row.get("source_file")
        for row in [*benefit_rows, *plan_rows]
        if _clean_text(row.get("source_file"))
    )
    if selected:
        return selected

    all_sources = _unique_paths(model.get("source_files") or [])
    if carrier_count == 1:
        return all_sources

    carrier_token = re.sub(r"[^a-z0-9]+", " ", carrier_name.casefold()).strip()
    matched = [
        path
        for path in all_sources
        if carrier_token and carrier_token in re.sub(r"[^a-z0-9]+", " ", Path(path).stem.casefold())
    ]
    if matched:
        return matched

    if 0 < carrier_index <= len(all_sources):
        return [all_sources[carrier_index - 1]]

    return all_sources


def _benefits(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    benefits: list[dict[str, Any]] = []
    seen: set[str] = set()
    for row in rows:
        name = _clean_text(row.get("benefit_line"))
        key = _carrier_key(name)
        if not name or key in seen:
            continue
        seen.add(key)
        benefits.append(
            {
                "benefit_key": key,
                "benefit_name": name,
                "normalized_benefit_name": name,
                "display_name": name,
                "benefit_category": _benefit_category(name),
                "sort_order": len(benefits) + 1,
                "is_standard": 1,
                "is_known": 1,
                "status": "VALID",
            }
        )
    return benefits


def _benefit_line(source: dict[str, Any]) -> dict[str, Any]:
    return {
        **_evidence(source),
        "benefit_line": source.get("benefit_line"),
        "class_name": source.get("class_name"),
        "tier": source.get("tier"),
        "lives": source.get("lives"),
        "volume": source.get("volume"),
        "rate": source.get("rate"),
        "monthly_premium": source.get("monthly_premium"),
        "annual_premium": source.get("annual_premium"),
    }


def _rate(source: dict[str, Any]) -> dict[str, Any]:
    return {
        **_evidence(source),
        "benefit_line": source.get("benefit_line"),
        "class_name": source.get("class_name"),
        "tier": source.get("tier"),
        "age_band": source.get("age_band"),
        "gender": source.get("gender"),
        "province": source.get("province"),
        "lives": source.get("lives"),
        "volume": source.get("volume"),
        "rate": source.get("rate"),
        "rate_basis": source.get("rate_basis"),
        "monthly_premium": source.get("monthly_premium"),
        "annual_premium": source.get("annual_premium"),
    }


def _premium(source: dict[str, Any]) -> dict[str, Any]:
    return {
        **_evidence(source),
        "benefit_line": source.get("benefit_line"),
        "class_name": source.get("class_name"),
        "tier": source.get("tier"),
        "monthly_premium": source.get("monthly_premium"),
        "annual_premium": source.get("annual_premium"),
        "employer_cost": source.get("employer_cost"),
        "employee_cost": source.get("employee_cost"),
        "total_cost": source.get("total_cost"),
    }


def _plan_design(source: dict[str, Any]) -> dict[str, Any]:
    return {
        **_evidence(source),
        "benefit_line": source.get("benefit_line"),
        "provision_name": source.get("provision_name"),
        "provision_value": source.get("provision_value"),
        "booklet_value": None,
        "variance": None,
        "validation_status": None,
    }


def _demographic(source: dict[str, Any]) -> dict[str, Any]:
    return {
        **_evidence(source),
        "benefit_line": source.get("benefit_line"),
        "class_name": source.get("class_name"),
        "tier": source.get("tier"),
        "lives": source.get("lives"),
        "volume": source.get("volume"),
    }


def _rate_guarantee_rows(model: dict[str, Any], carrier_name: str) -> list[dict[str, Any]]:
    values = (model.get("rate_guarantee_by_carrier") or {}).get(carrier_name) or {}
    if not values:
        return []
    guarantee = "; ".join(f"{key}: {value}" for key, value in sorted(values.items()) if _clean_text(value))
    if not guarantee:
        return []
    return [
        {
            "benefit_line": "All Benefits",
            "rate_guarantee": guarantee,
            "confidence_score": "MEDIUM",
            "status": "VALID",
            "source_sheet": "Rate Guarantee Summary",
        }
    ]


def _exceptions(model: dict[str, Any], source_files: list[str]) -> list[dict[str, Any]]:
    selected = {str(Path(path)).casefold() for path in source_files}
    rows: list[dict[str, Any]] = []
    for source in model.get("audit", []):
        if _clean_text(source.get("status")).upper() not in {"NEEDS REVIEW", "MISSING"}:
            continue
        source_file = _clean_text(source.get("source_file"))
        if source_file and selected and str(Path(source_file)).casefold() not in selected:
            continue
        rows.append(
            {
                **_evidence(source),
                "exception_type": source.get("status"),
                "field_name": source.get("field_name"),
                "message": _first_text(source.get("notes"), source.get("value"), "Review extracted proposal value."),
            }
        )
    return rows


def _evidence(source: dict[str, Any]) -> dict[str, Any]:
    location = _clean_text(source.get("source_location"))
    page_match = re.search(r"\bpage\s+(\d+)\b", location, re.IGNORECASE)
    return {
        "source_file": source.get("source_file"),
        "source_page": int(page_match.group(1)) if page_match else None,
        "source_sheet": None if page_match else location or None,
        "confidence_score": _first_text(source.get("confidence_score"), source.get("confidence")),
        "status": _first_text(source.get("status"), "VALID"),
        "source_type": source.get("source_type"),
        "source_location": location or None,
        "notes": source.get("notes"),
    }


def _has_rate_or_premium(row: dict[str, Any]) -> bool:
    return _has_rate_data(row) or _has_premium_data(row)


def _has_rate_data(row: dict[str, Any]) -> bool:
    return any(row.get(key) not in (None, "") for key in ("rate", "lives", "volume"))


def _has_premium_data(row: dict[str, Any]) -> bool:
    return any(row.get(key) not in (None, "") for key in ("monthly_premium", "annual_premium"))


def _has_demographic_data(row: dict[str, Any]) -> bool:
    return any(row.get(key) not in (None, "") for key in ("lives", "volume"))


def _benefit_category(value: str) -> str:
    text = value.casefold()
    if "dental" in text or "orthodont" in text:
        return "Dental"
    if "life" in text or "ad&d" in text:
        return "Life"
    if "disability" in text or text in {"ltd", "std"}:
        return "Disability"
    if any(token in text for token in ("health", "drug", "vision", "medical", "paramedical")):
        return "Health"
    return "Other"


def _package_hash(source_files: list[str]) -> str:
    hasher = hashlib.sha256()
    for value in sorted({_clean_text(item) for item in source_files if _clean_text(item)}, key=str.casefold):
        path = Path(value)
        hasher.update(path.name.casefold().encode("utf-8"))
        hasher.update(b"\0")
        if path.exists() and path.is_file():
            with path.open("rb") as source:
                for chunk in iter(lambda: source.read(1024 * 1024), b""):
                    hasher.update(chunk)
        else:
            hasher.update(value.encode("utf-8"))
        hasher.update(b"\0")
    return hasher.hexdigest()


def _unique_paths(values: Any) -> list[str]:
    paths: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = _clean_text(value)
        if not text:
            continue
        key = str(Path(text)).casefold()
        if key in seen:
            continue
        seen.add(key)
        paths.append(text)
    return sorted(paths, key=str.casefold)


def _carrier_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "_", _clean_text(value).casefold()).strip("_") or "unknown_carrier"


def _client_key(value: Any) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "-", _clean_text(value)).strip("-").upper()[:60] or "UNKNOWN-CLIENT"


def _positive_int(value: Any) -> int | None:
    try:
        numeric = int(value)
    except (TypeError, ValueError):
        return None
    return numeric if numeric > 0 else None


def _clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _first_text(*values: Any) -> str:
    for value in values:
        text = _clean_text(value)
        if text:
            return text
    return ""
