from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from proposal_engine import DEFAULT_BENEFIT_LIBRARY_PATH

ERROR_TOKENS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!", "#NUM!")


def load_benefit_library(path: str | Path | None = None) -> dict[str, Any]:
    library_path = Path(path) if path else DEFAULT_BENEFIT_LIBRARY_PATH
    with library_path.open("r", encoding="utf-8") as file:
        data = json.load(file)

    benefits: list[dict[str, Any]] = []
    for group_name in ["standard_benefits", "optional_benefits"]:
        for item in data.get(group_name, []):
            benefits.append(item)

    data["_benefits"] = benefits
    return data


def normalize_benefit_name(name: str | None, text: str = "", library: dict[str, Any] | None = None) -> dict[str, Any] | None:
    if not name and not text:
        return None

    source_name = clean_benefit_name(name or "")
    if not source_name and contains_error_token(text):
        return None

    search_text = f"{source_name} {text}".strip()
    normalized_text = _normalize(search_text)
    library = library or load_benefit_library()

    best_match: tuple[int, dict[str, Any]] | None = None
    for benefit in library.get("_benefits", []):
        terms = [benefit.get("canonical_name", ""), benefit.get("display_name", ""), *benefit.get("aliases", [])]
        for term in terms:
            term_norm = _normalize(term)
            if not term_norm:
                continue
            if re.search(rf"(?<![a-z0-9]){re.escape(term_norm)}(?![a-z0-9])", normalized_text):
                score = len(term_norm)
                if best_match is None or score > best_match[0]:
                    best_match = (score, benefit)

    if best_match is not None:
        benefit = best_match[1]
        return {
            "benefit_key": benefit.get("benefit_key"),
            "benefit_name": source_name or benefit.get("display_name"),
            "normalized_benefit_name": benefit.get("canonical_name") or benefit.get("display_name"),
            "display_name": benefit.get("display_name") or benefit.get("canonical_name"),
            "benefit_category": benefit.get("category", "standard"),
            "sort_order": benefit.get("sort_order", 999),
            "is_standard": benefit.get("category") == "standard",
            "is_known": True,
        }

    if not source_name:
        source_name = clean_benefit_name(text)
    if not source_name:
        return None
    if _is_non_benefit_label(source_name):
        return None

    key = re.sub(r"[^a-z0-9]+", "_", _normalize(source_name)).strip("_") or "carrier_specific_benefit"
    return {
        "benefit_key": f"carrier_specific_{key}",
        "benefit_name": source_name,
        "normalized_benefit_name": source_name,
        "display_name": source_name,
        "benefit_category": "carrier_specific",
        "sort_order": 900,
        "is_standard": False,
        "is_known": False,
    }


def looks_like_benefit_heading(cells: list[str], text: str, library: dict[str, Any] | None = None) -> str | None:
    library = library or load_benefit_library()
    if contains_error_token(text) and not any(
        (normalize_benefit_name(cell, "", library) or {}).get("is_known") for cell in cells[:2]
    ):
        return None

    cleaned_cells = [clean_benefit_name(cell) for cell in cells if clean_benefit_name(cell)]
    if not cleaned_cells:
        return None

    first = cleaned_cells[0]
    normalized = normalize_benefit_name(first, text, library)
    if normalized and normalized.get("is_known"):
        return normalized["benefit_name"] or normalized["display_name"]

    if _is_non_benefit_label(first):
        return None

    if len(first) < 3 or len(first) > 80:
        return None
    if _number_density(first) > 0.25:
        return None

    words = re.findall(r"[A-Za-z][A-Za-z&/-]*", first)
    if len(words) < 2:
        return None
    if _contains_benefit_signal(first):
        return first
    return None


def clean_benefit_name(value: str | None) -> str:
    value = re.sub(r"\s+", " ", str(value or "")).strip()
    for token in ERROR_TOKENS:
        value = value.replace(token, " ")
    value = re.sub(r"\s+", " ", value).strip()
    value = value.strip(":-")
    return value


def contains_error_token(value: Any) -> bool:
    text = str(value or "").upper()
    return any(token in text for token in ERROR_TOKENS)


def _contains_benefit_signal(value: str) -> bool:
    normalized = _normalize(value)
    signals = [
        "insurance",
        "life",
        "disability",
        "health",
        "dental",
        "care",
        "assistance",
        "spending",
        "illness",
        "opinion",
        "virtual",
        "program",
        "account",
    ]
    return any(signal in normalized for signal in signals)


def _is_non_benefit_label(value: str) -> bool:
    normalized = _normalize(value)
    blocked = {
        "",
        "single",
        "couple",
        "family",
        "employee only",
        "employee plus spouse",
        "employee plus child",
        "employee plus family",
        "class",
        "benefit",
        "total",
        "total cost",
        "total premium",
        "monthly premium",
        "annual total",
        "annual premium",
        "monthly premium difference",
        "annual premium difference",
        "difference",
        "% difference",
        "percent difference",
        "monthly premium difference",
        "percentage difference",
        "avg comp",
        "average comp",
        "comp this year",
        "estimated total",
        "union annual premiums",
        "rate guarantee",
        "pooling level",
        "pooling charge",
        "target loss ratio",
        "ibnr",
        "rate caps",
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
    if normalized in blocked:
        return True
    if re.match(r"^[\d$]", normalized):
        return True
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
    if any(term in f" {normalized}" for term in provision_terms):
        return True
    blocked_prefixes = (
        "total ",
        "difference ",
        "% difference",
        "percentage difference",
        "avg comp",
        "average comp",
        "comp this year",
        "estimated total",
        "union annual premiums",
        "current plan",
    )
    return normalized.startswith("class ") or normalized.startswith(blocked_prefixes)


def _number_density(value: str) -> float:
    if not value:
        return 0
    return sum(1 for char in value if char.isdigit()) / len(value)


def _normalize(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip().lower()
