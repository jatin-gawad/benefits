from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

from proposal_engine.benefit_library import (
    clean_benefit_name,
    contains_error_token,
    load_benefit_library,
    looks_like_benefit_heading,
    normalize_benefit_name,
)


MONEY_RE = re.compile(r"(?<!\w)-?\$?\s*\d[\d,]*(?:\.\d{1,4})?\s*%?")
DATE_RE = re.compile(
    r"\b(?:\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|"
    r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4})\b",
    re.IGNORECASE,
)


def extract_fields(sections: list[dict[str, Any]], carrier_config: dict[str, Any]) -> dict[str, Any]:
    """Extract generic proposal fields from classified sections using config mappings."""

    source_file = _first_value(section.get("source_file") for section in sections)
    carrier_name = carrier_config.get("carrier_name")
    raw: dict[str, Any] = {
        "proposal_summary": {
            "source_file": source_file,
            "carrier_name": carrier_name,
        },
        "benefits": [],
        "benefit_lines": [],
        "rates": [],
        "premiums": [],
        "plan_design": [],
        "demographics": [],
        "underwriting_assumptions": [],
        "pooling_retention": [],
        "exceptions": [],
    }

    _extract_summary(raw, sections, carrier_config)
    _extract_tables(raw, sections, carrier_config)
    _extract_assumptions(raw, sections, carrier_config)
    _add_missing_section_exceptions(raw, sections)
    return raw


def _extract_summary(raw: dict[str, Any], sections: list[dict[str, Any]], carrier_config: dict[str, Any]) -> None:
    field_mappings = carrier_config.get("field_mappings", {})
    source_file = raw["proposal_summary"].get("source_file")

    for field_name in ["client_name", "proposal_date", "effective_date", "monthly_premium", "annual_premium"]:
        labels = field_mappings.get(field_name, [])
        value, confidence, location = _find_label_value(sections, labels, prefer=field_name)
        if value is not None:
            raw["proposal_summary"][field_name] = value
            if field_name in {"monthly_premium", "annual_premium"}:
                premium = _premium_record(
                    carrier_name=raw["proposal_summary"].get("carrier_name"),
                    benefit_line="Total",
                    monthly_premium=value if field_name == "monthly_premium" else None,
                    annual_premium=value if field_name == "annual_premium" else None,
                    source_file=source_file,
                    location=location,
                    confidence=confidence,
                    notes=f"Extracted from summary label {field_name}."
                )
                raw["premiums"].append(premium)
        else:
            raw["exceptions"].append(
                _exception(
                    raw["proposal_summary"].get("carrier_name"),
                    "MISSING",
                    field_name,
                    f"{field_name.replace('_', ' ').title()} was not found.",
                    source_file,
                )
            )


def _extract_tables(raw: dict[str, Any], sections: list[dict[str, Any]], carrier_config: dict[str, Any]) -> None:
    benefit_library = load_benefit_library()
    for section in sections:
        section_name = section.get("section_name", "unknown")
        if section_name in {"underwriting", "pooling_retention"}:
            continue
        for table in section.get("tables", []):
            current_benefit: dict[str, Any] | None = None
            for row_index, row in enumerate(table):
                cells = [str(cell).strip() for cell in row if str(cell).strip()]
                if not cells:
                    continue
                row_text = " ".join(cells)
                mapped_benefit = _detect_mapping(row_text, carrier_config.get("benefit_mappings", {}))
                heading = looks_like_benefit_heading(cells, row_text, benefit_library)
                normalized_benefit = normalize_benefit_name(mapped_benefit or heading, row_text, benefit_library)
                row_has_formula_error = contains_error_token(row_text)

                if heading and normalized_benefit:
                    current_benefit = normalized_benefit
                    _add_benefit_record(raw, normalized_benefit, section, "MEDIUM", "Benefit heading found in proposal content.")
                elif normalized_benefit and normalized_benefit.get("is_known"):
                    current_benefit = normalized_benefit
                    _add_benefit_record(raw, normalized_benefit, section, "MEDIUM", "Known benefit found in row content.")
                elif current_benefit and _row_belongs_to_current_benefit(cells, row_text):
                    normalized_benefit = current_benefit

                if row_has_formula_error:
                    if normalized_benefit and normalized_benefit.get("is_known"):
                        _add_benefit_record(raw, normalized_benefit, section, "LOW", "Benefit retained, but formula error values were skipped.")
                    continue

                if _is_support_summary_row(cells, row_text) and not (
                    normalized_benefit and normalized_benefit.get("is_known")
                ):
                    continue

                benefit_line = normalized_benefit.get("display_name") if normalized_benefit else None
                tier = _detect_mapping(row_text, carrier_config.get("tier_mappings", {}))
                class_name = _detect_class_from_label(cells[0]) or _detect_class(cells[0]) or _detect_class(row_text)
                tier = tier or _detect_tier_from_label(cells[0])
                numbers = _extract_numbers(row_text)
                source_file = section.get("source_file")
                location = {
                    "source_page": section.get("page_number"),
                    "source_sheet": section.get("sheet_name"),
                }

                if section_name == "plan_design" and benefit_line:
                    record = _plan_design_record(
                        raw["proposal_summary"].get("carrier_name"),
                        benefit_line,
                        cells,
                        source_file,
                        location,
                        "MEDIUM",
                    )
                    if record:
                        raw["plan_design"].append(record)
                    continue

                if section_name == "plan_design" and current_benefit:
                    record = _plan_design_record(
                        raw["proposal_summary"].get("carrier_name"),
                        current_benefit.get("display_name"),
                        cells,
                        source_file,
                        location,
                        "LOW",
                    )
                    if record:
                        raw["plan_design"].append(record)
                    continue

                if not benefit_line:
                    continue

                _add_benefit_record(raw, normalized_benefit, section, "MEDIUM", "Benefit row retained for market survey output.")
                metrics = _infer_row_metrics(cells, numbers, row_text)
                lives = metrics.get("lives")
                volume = metrics.get("volume")
                rate = metrics.get("rate")
                monthly_premium = metrics.get("monthly_premium")
                annual_premium = metrics.get("annual_premium")
                confidence = "MEDIUM" if any(value is not None for value in [rate, monthly_premium, annual_premium]) else "LOW"
                status = "VALID" if confidence == "MEDIUM" else "NEEDS REVIEW"

                line_record = {
                    "carrier_name": raw["proposal_summary"].get("carrier_name"),
                    "benefit_line": benefit_line,
                    "class_name": class_name,
                    "tier": tier,
                    "lives": lives,
                    "volume": volume,
                    "rate": rate,
                    "monthly_premium": monthly_premium,
                    "annual_premium": annual_premium,
                    "rate_basis": _detect_rate_basis(row_text),
                    "volume_basis": "volume" if volume is not None else None,
                    "source_file": source_file,
                    "source_page": location["source_page"],
                    "source_sheet": location["source_sheet"],
                    "confidence_score": confidence,
                    "status": status,
                    "notes": f"Generic table extraction from {section_name}, row {row_index + 1}.",
                }
                raw["benefit_lines"].append(line_record)

                if rate is not None or monthly_premium is not None or annual_premium is not None:
                    raw["rates"].append(
                        {
                            "carrier_name": line_record["carrier_name"],
                            "benefit_line": benefit_line,
                            "class_name": class_name,
                            "tier": tier,
                            "age_band": None,
                            "gender": None,
                            "province": None,
                            "lives": lives,
                            "volume": volume,
                            "rate": rate,
                            "rate_basis": line_record["rate_basis"],
                            "monthly_premium": monthly_premium,
                            "annual_premium": annual_premium,
                            "source_file": source_file,
                            "source_page": location["source_page"],
                            "source_sheet": location["source_sheet"],
                            "confidence_score": confidence,
                            "status": status,
                            "notes": line_record["notes"],
                        }
                    )

                if monthly_premium is not None or annual_premium is not None:
                    raw["premiums"].append(
                        _premium_record(
                            carrier_name=line_record["carrier_name"],
                            benefit_line=benefit_line,
                            monthly_premium=monthly_premium,
                            annual_premium=annual_premium,
                            source_file=source_file,
                            location=location,
                            confidence=confidence,
                            notes=line_record["notes"],
                            class_name=class_name,
                            tier=tier,
                        )
                    )

                if lives is not None or volume is not None:
                    raw["demographics"].append(
                        {
                            "carrier_name": line_record["carrier_name"],
                            "benefit_line": benefit_line,
                            "class_name": class_name,
                            "tier": tier,
                            "lives": lives,
                            "volume": volume,
                            "source_file": source_file,
                            "source_page": location["source_page"],
                            "source_sheet": location["source_sheet"],
                            "confidence_score": confidence,
                            "status": status,
                            "notes": line_record["notes"],
                        }
                    )


def _extract_assumptions(raw: dict[str, Any], sections: list[dict[str, Any]], carrier_config: dict[str, Any]) -> None:
    field_mappings = carrier_config.get("field_mappings", {})
    carrier_name = raw["proposal_summary"].get("carrier_name")

    assumption_fields = [
        "pooling_level",
        "pooling_charge",
        "commission",
        "retention",
        "target_loss_ratio",
        "rate_guarantee",
        "pooled_benefits",
        "experience_rated_benefits",
        "ibnr",
        "rate_caps",
        "first_renewal_cap",
        "second_renewal_cap",
    ]
    for field_name in assumption_fields:
        value, confidence, location = _find_label_value(sections, field_mappings.get(field_name, []), prefer=field_name)
        if value is None:
            continue
        record = {
            "carrier_name": carrier_name,
            "benefit_line": None,
            "pooling_level": None,
            "pooling_charge": None,
            "commission": None,
            "retention": None,
            "target_loss_ratio": None,
            "rate_guarantee": None,
            "pooled_benefits": None,
            "experience_rated_benefits": None,
            "ibnr": None,
            "rate_caps": None,
            "first_renewal_cap": None,
            "second_renewal_cap": None,
            "source_file": location.get("source_file"),
            "source_page": location.get("source_page"),
            "source_sheet": location.get("source_sheet"),
            "confidence_score": confidence,
            "status": "VALID" if confidence != "LOW" else "NEEDS REVIEW",
            "notes": f"Extracted from {field_name} label.",
        }
        record[field_name] = value
        raw["pooling_retention"].append(record)

    for section in sections:
        if section.get("section_name") != "underwriting":
            continue
        lines = [line.strip() for line in section.get("text", "").splitlines() if line.strip()]
        for line in lines[:25]:
            if len(line) < 8:
                continue
            raw["underwriting_assumptions"].append(
                {
                    "carrier_name": carrier_name,
                    "assumption_name": _first_words(line, 6),
                    "assumption_value": line,
                    "source_file": section.get("source_file"),
                    "source_page": section.get("page_number"),
                    "source_sheet": section.get("sheet_name"),
                    "confidence_score": "LOW",
                    "status": "NEEDS REVIEW",
                    "notes": "Generic underwriting text captured for review.",
                }
            )


def _find_label_value(
    sections: list[dict[str, Any]],
    labels: list[str],
    prefer: str = ""
) -> tuple[Any, str, dict[str, Any]]:
    label_terms = [_normalize(label) for label in labels if label]
    for section in sections:
        location = {
            "source_file": section.get("source_file"),
            "source_page": section.get("page_number"),
            "source_sheet": section.get("sheet_name"),
        }
        for table in section.get("tables", []):
            for row in table:
                cells = [str(cell).strip() for cell in row]
                row_text = _normalize(" ".join(cells))
                if not any(label in row_text for label in label_terms):
                    continue
                value = _value_from_cells(cells, labels, prefer)
                if value is not None:
                    return value, "HIGH", location

        for line in section.get("text", "").splitlines():
            line_norm = _normalize(line)
            if not any(label in line_norm for label in label_terms):
                continue
            value = _value_from_line(line, labels, prefer)
            if value is not None:
                return value, "MEDIUM", location

    return None, "LOW", {}


def _add_benefit_record(
    raw: dict[str, Any],
    benefit: dict[str, Any] | None,
    section: dict[str, Any],
    confidence: str,
    notes: str,
) -> None:
    if not benefit:
        return

    record = {
        "carrier_name": raw.get("proposal_summary", {}).get("carrier_name"),
        "benefit_key": benefit.get("benefit_key"),
        "benefit_name": benefit.get("benefit_name"),
        "normalized_benefit_name": benefit.get("normalized_benefit_name"),
        "display_name": benefit.get("display_name"),
        "benefit_category": benefit.get("benefit_category"),
        "sort_order": benefit.get("sort_order"),
        "is_standard": benefit.get("is_standard"),
        "is_known": benefit.get("is_known"),
        "source_file": section.get("source_file"),
        "source_page": section.get("page_number"),
        "source_sheet": section.get("sheet_name"),
        "confidence_score": confidence,
        "status": "VALID" if confidence in {"HIGH", "MEDIUM"} else "NEEDS REVIEW",
        "notes": notes,
    }

    key = (
        record["carrier_name"],
        record["benefit_key"],
        record["source_file"],
        record["source_page"],
        record["source_sheet"],
    )
    existing_keys = {
        (
            item.get("carrier_name"),
            item.get("benefit_key"),
            item.get("source_file"),
            item.get("source_page"),
            item.get("source_sheet"),
        )
        for item in raw.get("benefits", [])
    }
    if key not in existing_keys:
        raw.setdefault("benefits", []).append(record)


def _row_belongs_to_current_benefit(cells: list[str], row_text: str) -> bool:
    first = cells[0].strip() if cells else ""
    if not first:
        return False
    if _detect_class(first) or _detect_tier_from_label(first):
        return True
    if _extract_numbers(row_text) and len(first) <= 80:
        return True
    if re.search(r"\b(rate|premium|included|not available|declined|pooled|guarantee)\b", row_text, flags=re.IGNORECASE):
        return True
    return False


def _is_support_summary_row(cells: list[str], row_text: str) -> bool:
    first = _normalize(clean_benefit_name(cells[0] if cells else ""))
    if not first:
        return True
    support_exact = {
        "current plan",
        "carrier",
        "rate",
        "premium",
        "monthly premium",
        "annual premium",
        "total",
        "total premium",
        "total cost",
        "difference",
        "percentage difference",
        "percent difference",
        "% difference",
        "avg comp",
        "average comp",
        "estimated total",
        "union annual premiums",
    }
    support_prefixes = (
        "total ",
        "difference ",
        "percentage difference",
        "percent difference",
        "% difference",
        "avg comp",
        "average comp",
        "estimated total",
        "comp this year",
        "current plan ",
    )
    if first in support_exact or first.startswith(support_prefixes):
        return True
    header_words = {"benefit", "class", "lives", "volume", "rate", "premium", "notes"}
    row_words = set(re.findall(r"[a-z]+", _normalize(row_text)))
    return len(row_words) >= 3 and row_words.issubset(header_words)


def _infer_row_metrics(cells: list[str], numbers: list[dict[str, Any]], row_text: str) -> dict[str, Any]:
    lives = _pick_lives(numbers, row_text)
    volume = _pick_volume(numbers, row_text)
    rate = _pick_rate(numbers, row_text)
    monthly_premium = _pick_premium(numbers, row_text, annual=False)
    annual_premium = _pick_premium(numbers, row_text, annual=True)

    numeric_values = [item["value"] for item in numbers if not item["is_percent"]]
    if rate is None and numeric_values:
        rate_candidates = [value for value in numeric_values if 0 <= value <= 10000]
        if rate_candidates:
            rate = rate_candidates[0]

    if monthly_premium is None:
        currency_values = [item["value"] for item in numbers if item["is_currency"] and not item["is_percent"]]
        if currency_values:
            monthly_premium = currency_values[0]

    if annual_premium is None and monthly_premium is not None and "annual" in _normalize(row_text):
        annual_premium = monthly_premium
        monthly_premium = None

    if lives is None and re.search(r"\b(lives|employees|certificates)\b", row_text, flags=re.IGNORECASE):
        lives_candidates = [value for value in numeric_values if 0 <= value <= 100000]
        if lives_candidates:
            lives = int(lives_candidates[0])

    if volume is None and re.search(r"\b(volume|benefit amount|amount)\b", row_text, flags=re.IGNORECASE):
        volume_candidates = [value for value in numeric_values if value >= 1000]
        if volume_candidates:
            volume = volume_candidates[0]

    return {
        "lives": lives,
        "volume": volume,
        "rate": rate,
        "monthly_premium": monthly_premium,
        "annual_premium": annual_premium,
    }


def _value_from_cells(cells: list[str], labels: list[str], prefer: str) -> Any:
    for index, cell in enumerate(cells):
        if any(_normalize(label) in _normalize(cell) for label in labels):
            nearby = cells[index + 1:index + 5]
            for candidate in nearby:
                value = _coerce_value(candidate, prefer)
                if value is not None:
                    return value
            value = _coerce_value(cell, prefer)
            if value is not None:
                return value
    return None


def _value_from_line(line: str, labels: list[str], prefer: str) -> Any:
    for label in labels:
        match = re.search(re.escape(label), line, flags=re.IGNORECASE)
        if not match:
            continue
        tail = line[match.end():]
        value = _coerce_value(tail, prefer)
        if value is not None:
            return value
    return _coerce_value(line, prefer)


def _coerce_value(text: str, prefer: str) -> Any:
    text = (text or "").strip(" :-\t")
    if not text:
        return None
    if "date" in prefer:
        match = DATE_RE.search(text)
        return match.group(0) if match else None
    if prefer in {"monthly_premium", "annual_premium", "pooling_charge", "commission", "retention", "target_loss_ratio"}:
        number = _first_number(text)
        return number
    if prefer in {"lives"}:
        number = _first_number(text)
        return int(number) if number is not None else None
    if prefer in {"volume", "pooling_level"}:
        return _first_number(text) or text[:120]
    cleaned = re.sub(r"\s+", " ", text)
    return cleaned[:160] if cleaned else None


def _detect_mapping(text: str, mappings: dict[str, list[str]]) -> str | None:
    normalized = _normalize(text)
    matches: list[tuple[int, str]] = []
    for canonical, aliases in mappings.items():
        for alias in aliases:
            alias_norm = _normalize(alias)
            if alias_norm and re.search(rf"(?<![a-z0-9]){re.escape(alias_norm)}(?![a-z0-9])", normalized):
                matches.append((len(alias_norm), canonical))
    if not matches:
        return None
    return sorted(matches, reverse=True)[0][1]


def _detect_class(text: str) -> str | None:
    match = re.search(
        r"\b(?:class|division)\s+[A-Za-z0-9](?:\s*(?:to|and|&|-|/)\s*[A-Za-z0-9])?",
        text,
        flags=re.IGNORECASE,
    )
    return match.group(0).strip() if match else None


def _detect_class_from_label(text: str) -> str | None:
    text = clean_benefit_name(text)
    normalized = _normalize(text)
    if normalized.startswith("class "):
        return _detect_class(text) or text.strip()
    class_labels = [
        "executive",
        "union",
        "non union",
        "non-union",
        "hourly",
        "salaried",
        "management",
        "mgmt",
    ]
    for label in class_labels:
        if normalized == label or normalized.startswith(label + " "):
            return text.strip()
    return None


def _detect_tier_from_label(text: str) -> str | None:
    normalized = _normalize(clean_benefit_name(text))
    tier_map = {
        "single": "Single",
        "couple": "Couple",
        "family": "Family",
        "employee only": "Employee only",
        "employee plus spouse": "Employee plus spouse",
        "employee + spouse": "Employee plus spouse",
        "employee plus child": "Employee plus child",
        "employee + child": "Employee plus child",
        "employee plus family": "Employee plus family",
        "employee + family": "Employee plus family",
        "ee": "Employee only",
        "ee + spouse": "Employee plus spouse",
        "ee + child": "Employee plus child",
        "ee + family": "Employee plus family",
    }
    return tier_map.get(normalized)


def _extract_numbers(text: str) -> list[dict[str, Any]]:
    values: list[dict[str, Any]] = []
    for match in MONEY_RE.finditer(text):
        raw = match.group(0).strip()
        number = _parse_number(raw)
        if number is None:
            continue
        values.append(
            {
                "raw": raw,
                "value": number,
                "is_currency": "$" in raw or "," in raw,
                "is_percent": "%" in raw,
            }
        )
    return values


def _parse_number(raw: str) -> float | None:
    cleaned = raw.replace("$", "").replace(",", "").replace("%", "").strip()
    try:
        return float(cleaned)
    except ValueError:
        return None


def _first_number(text: str) -> float | None:
    numbers = _extract_numbers(text)
    return numbers[0]["value"] if numbers else None


def _pick_lives(numbers: list[dict[str, Any]], text: str) -> int | None:
    if not re.search(r"\b(lives|employees|certificates|count)\b", text, flags=re.IGNORECASE):
        return None
    candidates = [item["value"] for item in numbers if 0 <= item["value"] <= 100000 and not item["is_percent"]]
    return int(candidates[0]) if candidates else None


def _pick_volume(numbers: list[dict[str, Any]], text: str) -> float | None:
    if not re.search(r"\b(volume|benefit amount|insured)\b", text, flags=re.IGNORECASE):
        return None
    candidates = [item["value"] for item in numbers if item["value"] >= 1000 and not item["is_percent"]]
    return candidates[0] if candidates else None


def _pick_rate(numbers: list[dict[str, Any]], text: str) -> float | None:
    if not re.search(r"\b(rate|rates)\b", text, flags=re.IGNORECASE):
        return None
    candidates = [item["value"] for item in numbers if 0 <= item["value"] < 100 and not item["is_currency"]]
    return candidates[0] if candidates else None


def _pick_premium(numbers: list[dict[str, Any]], text: str, annual: bool) -> float | None:
    target = "annual" if annual else "monthly"
    if target not in _normalize(text) and "premium" not in _normalize(text) and "cost" not in _normalize(text):
        return None
    candidates = [item["value"] for item in numbers if item["value"] >= 1 and not item["is_percent"]]
    if not candidates:
        return None
    return max(candidates) if annual else candidates[0]


def _detect_rate_basis(text: str) -> str | None:
    normalized = _normalize(text)
    if "1000" in normalized or "1,000" in normalized:
        return "per $1,000 volume"
    if "month" in normalized:
        return "monthly"
    if "annual" in normalized:
        return "annual"
    return None


def _plan_design_record(
    carrier_name: str | None,
    benefit_line: str,
    cells: list[str],
    source_file: str | None,
    location: dict[str, Any],
    confidence: str,
) -> dict[str, Any] | None:
    if len(cells) < 2:
        return None
    provision_name = cells[0]
    provision_value = " | ".join(cells[1:])[:500]
    return {
        "carrier_name": carrier_name,
        "benefit_line": benefit_line,
        "provision_name": provision_name,
        "provision_value": provision_value,
        "booklet_value": None,
        "variance": None,
        "validation_status": "NOT AVAILABLE",
        "source_file": source_file,
        "source_page": location.get("source_page"),
        "source_sheet": location.get("source_sheet"),
        "confidence_score": confidence,
        "status": "NEEDS REVIEW",
        "notes": "Generic plan design row captured for review.",
    }


def _premium_record(
    carrier_name: str | None,
    benefit_line: str | None,
    monthly_premium: Any,
    annual_premium: Any,
    source_file: str | None,
    location: dict[str, Any],
    confidence: str,
    notes: str,
    class_name: str | None = None,
    tier: str | None = None,
) -> dict[str, Any]:
    return {
        "carrier_name": carrier_name,
        "benefit_line": benefit_line,
        "class_name": class_name,
        "tier": tier,
        "monthly_premium": monthly_premium,
        "annual_premium": annual_premium,
        "employer_cost": None,
        "employee_cost": None,
        "total_cost": annual_premium or monthly_premium,
        "source_file": source_file,
        "source_page": location.get("source_page"),
        "source_sheet": location.get("source_sheet"),
        "confidence_score": confidence,
        "status": "VALID" if confidence in {"HIGH", "MEDIUM"} else "NEEDS REVIEW",
        "notes": notes,
    }


def _add_missing_section_exceptions(raw: dict[str, Any], sections: list[dict[str, Any]]) -> None:
    found = {section.get("section_name") for section in sections if section.get("section_name") != "unknown"}
    for section_name in ["premium_summary", "rate_summary", "plan_design"]:
        if section_name not in found:
            raw["exceptions"].append(
                _exception(
                    raw["proposal_summary"].get("carrier_name"),
                    "MISSING",
                    section_name,
                    f"No confident {section_name.replace('_', ' ')} section was found.",
                    raw["proposal_summary"].get("source_file"),
                )
            )


def _exception(
    carrier_name: str | None,
    exception_type: str,
    field_name: str,
    message: str,
    source_file: str | None,
) -> dict[str, Any]:
    return {
        "carrier_name": carrier_name,
        "exception_type": exception_type,
        "field_name": field_name,
        "message": message,
        "source_file": source_file,
        "source_page": None,
        "source_sheet": None,
        "confidence_score": "LOW",
        "status": exception_type if exception_type in {"MISSING", "CONFLICT", "LOW CONFIDENCE"} else "NEEDS REVIEW",
        "notes": "Created by generic proposal extraction.",
    }


def _first_value(values: Any) -> Any:
    for value in values:
        if value:
            return value
    return None


def _first_words(text: str, count: int) -> str:
    words = re.findall(r"\w+", text)
    return " ".join(words[:count])


def _normalize(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip().lower()
