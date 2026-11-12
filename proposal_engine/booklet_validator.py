from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from proposal_engine.booklet_plan_design.labeled_dental_rate_mapper import (
    canonical_dental_provision_code,
    dental_coinsurance_provision_code,
    dental_maximum_provision_code,
    dental_provision_name,
)
from proposal_engine.document_reader import read_document


_BOOKLET_FIELD_MAP = (
    ("coverage", "Coverage Description", "coverage_description"),
    ("coinsurance", "Coinsurance", "coinsurance"),
    ("maximum", "Maximum", "maximum"),
    ("frequency", "Frequency", "frequency"),
    ("waiting_period", "Waiting Period", "waiting_period"),
    ("termination_age", "Termination Age", "termination_age"),
)

_BENEFIT_ALIASES = (
    ("dependent_life", ("dependent life", "spouse life", "child life")),
    ("adnd", ("ad&d", "a d & d", "accidental death", "accidental death and dismemberment")),
    ("critical_illness", ("critical illness",)),
    ("ltd", ("long term disability", "long-term disability", "ltd")),
    ("std", ("short term disability", "short-term disability", "weekly indemnity", "std")),
    ("paramedical", ("paramedical", "professional services", "practitioner")),
    ("vision", ("vision", "eye exam", "eyeglass")),
    ("drugs", ("prescription drug", "prescription drugs", "drug plan", "drugs")),
    ("hospital", ("hospital", "hospital room")),
    ("travel", ("travel", "out of country", "out-of-country")),
    ("dental", ("dental", "orthodont", "recall", "scaling")),
    ("life", ("employee life", "basic life", "life insurance", "life")),
    ("health", ("extended health care", "extended health", "health care", "ehc")),
)

_HEALTH_BENEFITS = {"health", "drugs", "paramedical", "vision", "hospital", "travel"}
_STRUCTURED_PROVISIONS = {"coinsurance", "maximum", "frequency", "waiting_period", "termination_age"}
_VALIDATOR_SOURCE = "canonical_proposal_booklet_validation"


def validate_against_booklet(normalized_data: dict[str, Any], booklet_path: str | Path | None) -> dict[str, Any]:
    """Compare proposal plan design rows to a benefit booklet when one is provided."""

    if not booklet_path:
        for record in normalized_data.get("plan_design", []):
            record["validation_status"] = "BOOKLET NOT PROVIDED"
        return normalized_data

    try:
        booklet_document = read_document(booklet_path)
        booklet_text = _normalize(booklet_document.get("all_text", ""))
    except Exception as exc:
        normalized_data.setdefault("exceptions", []).append(
            {
                "carrier_name": normalized_data.get("proposal_summary", {}).get("carrier_name"),
                "exception_type": "NEEDS REVIEW",
                "field_name": "booklet",
                "message": f"Benefit booklet could not be read: {exc}",
                "source_file": str(booklet_path),
                "source_page": None,
                "source_sheet": None,
                "confidence_score": "LOW",
                "status": "NEEDS REVIEW",
                "notes": "Booklet validation skipped after read failure.",
            }
        )
        for record in normalized_data.get("plan_design", []):
            record["validation_status"] = "NEEDS REVIEW"
        return normalized_data

    if not normalized_data.get("plan_design"):
        normalized_data.setdefault("exceptions", []).append(
            {
                "carrier_name": normalized_data.get("proposal_summary", {}).get("carrier_name"),
                "exception_type": "MISSING",
                "field_name": "plan_design",
                "message": "Booklet was provided, but proposal plan design rows were not available to compare.",
                "source_file": normalized_data.get("proposal_summary", {}).get("source_file"),
                "source_page": None,
                "source_sheet": None,
                "confidence_score": "LOW",
                "status": "MISSING",
                "notes": "Booklet validation could not compare missing proposal rows.",
            }
        )
        return normalized_data

    for record in normalized_data.get("plan_design", []):
        provision_name = _normalize(str(record.get("provision_name") or ""))
        provision_value = _normalize(str(record.get("provision_value") or ""))
        if not provision_name:
            record["validation_status"] = "NEEDS REVIEW"
            continue

        if provision_name not in booklet_text:
            record["validation_status"] = "MISSING FROM BOOKLET"
            record["notes"] = _append_note(record.get("notes"), "Provision label was not found in the booklet.")
            continue

        if provision_value and provision_value in booklet_text:
            record["booklet_value"] = record.get("provision_value")
            record["variance"] = None
            record["validation_status"] = "MATCH"
        else:
            record["booklet_value"] = _find_booklet_snippet(booklet_text, provision_name)
            record["variance"] = "Proposal value differs from or could not be matched to booklet text."
            record["validation_status"] = "DIFFERENT" if provision_value else "NEEDS REVIEW"

    return normalized_data


def validate_focused_proposals_against_booklet(
    envelope: dict[str, Any],
    booklet_path: str | Path | None,
) -> dict[str, Any]:
    """Validate focused proposal facts using the established structured booklet extractor."""

    if not booklet_path:
        return envelope

    try:
        booklet_facts = extract_structured_booklet_facts(booklet_path)
    except Exception as exc:
        return _record_booklet_extraction_failure(envelope, booklet_path, str(exc))

    return validate_focused_proposals_against_booklet_facts(
        envelope,
        booklet_facts,
        booklet_path=booklet_path,
    )


def extract_structured_booklet_facts(booklet_path: str | Path) -> list[dict[str, Any]]:
    """Adapt booklet ExtractionRow values to validation facts without writing a workbook or database."""

    from proposal_engine.booklet_plan_design import booklet_plan_design_extractor as extractor

    path = Path(booklet_path)
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(f"Benefits booklet was not found: {path}")
    if path.suffix.casefold() != ".pdf":
        raise ValueError("Canonical proposal validation currently requires a benefits booklet PDF.")

    rules = extractor.load_extraction_rules("")
    text = extractor.extract_pdf_text(path)
    if not extractor.clean_text(text):
        raise ValueError("No readable text was found in the benefits booklet PDF.")

    client_info = extractor.detect_client_info(text, path, rules)
    class_names = extractor.detect_booklet_class_names(text, path, client_info)
    benefit_text = extractor.remove_front_matter_for_benefits(text, rules)
    plan_design_text = text if extractor.is_schedule_first_booklet(text, client_info) else benefit_text

    rows = []
    for class_name in class_names:
        rows.extend(
            extractor.extract_plan_design_rows(
                plan_design_text,
                path,
                client_info,
                rules,
                class_name,
                document_text=text,
            )
        )

    facts: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str, str, str, str]] = set()
    for row in rows:
        for provision_code, provision_name, attribute_name in _BOOKLET_FIELD_MAP:
            value = str(getattr(row, attribute_name, "") or "").strip()
            if not value:
                continue
            fact_provision_code = provision_code
            fact_provision_name = provision_name
            if provision_code == "coinsurance":
                dental_code = canonical_dental_provision_code(row.benefit)
                if dental_code:
                    fact_provision_code = dental_coinsurance_provision_code(dental_code)
                    fact_provision_name = f"{dental_provision_name(dental_code)} Coinsurance"
            elif provision_code == "maximum":
                dental_code = canonical_dental_provision_code(row.benefit)
                if dental_code:
                    fact_provision_code = dental_maximum_provision_code(dental_code)
                    fact_provision_name = f"{dental_provision_name(dental_code)} Maximum"
            fact = {
                "benefit_line": row.benefit,
                "benefit_code": _benefit_code(row.benefit),
                "provision_code": fact_provision_code,
                "provision_name": fact_provision_name,
                "provision_value": value,
                "class_name": row.class_name,
                "client_name": row.client_name,
                "carrier": row.carrier,
                "source_file": str(path),
                "source_page": _positive_int(row.source_page),
                "source_section": row.source_section,
                "confidence_score": row.confidence,
                "review_flag": row.review_flag,
            }
            key = (
                _normalize(row.benefit),
                fact_provision_code,
                _comparable_value(value),
                _normalize(row.class_name),
                str(row.source_page or ""),
                _normalize(row.source_section),
            )
            if key in seen:
                continue
            seen.add(key)
            facts.append(fact)

    return facts


def validate_focused_proposals_against_booklet_facts(
    envelope: dict[str, Any],
    booklet_facts: list[dict[str, Any]],
    booklet_path: str | Path = "",
) -> dict[str, Any]:
    """Compare canonical proposal rows to already extracted canonical booklet facts."""

    proposals = envelope.get("proposals") or []
    totals = {
        "status": "COMPLETED",
        "booklet_path": str(booklet_path or ""),
        "booklet_facts": len(booklet_facts),
        "proposal_facts": 0,
        "validation_rows": 0,
        "matches": 0,
        "differences": 0,
        "review_issues": 0,
    }

    if not booklet_facts:
        return _record_booklet_extraction_failure(
            envelope,
            booklet_path,
            "The booklet extractor did not produce any structured plan design facts.",
        )

    for proposal in proposals:
        proposal["booklet_validation"] = []
        proposal["exceptions"] = [
            row
            for row in proposal.get("exceptions", [])
            if row.get("validation_source") != _VALIDATOR_SOURCE
        ]
        plan_rows = proposal.get("plan_design") or []
        totals["proposal_facts"] += len(plan_rows)

        if not plan_rows:
            _append_validation_exception(
                proposal,
                None,
                "No proposal plan design facts were available for booklet validation.",
            )
            totals["review_issues"] += 1
            continue

        for proposal_fact in plan_rows:
            outcome = _compare_proposal_fact(proposal_fact, booklet_facts)
            validation = outcome.get("validation")
            if validation is not None:
                proposal["booklet_validation"].append(validation)
                totals["validation_rows"] += 1
                if validation["validation_status"] == "MATCH":
                    totals["matches"] += 1
                else:
                    totals["differences"] += 1
                continue

            _append_validation_exception(proposal, proposal_fact, str(outcome["review_message"]))
            totals["review_issues"] += 1

    if totals["validation_rows"] == 0:
        totals["status"] = "NEEDS REVIEW"
    elif totals["differences"] or totals["review_issues"]:
        totals["status"] = "COMPLETED WITH REVIEW"
    envelope["_booklet_validation_summary"] = totals
    return envelope


def _compare_proposal_fact(
    proposal_fact: dict[str, Any],
    booklet_facts: list[dict[str, Any]],
) -> dict[str, Any]:
    provision_name = str(proposal_fact.get("provision_name") or "")
    provision_code = _provision_code(
        provision_name,
        str(proposal_fact.get("provision_value") or ""),
    )
    if not provision_code:
        return {
            "review_message": (
                f"Booklet validation has no canonical provision mapping for '{provision_name or 'unnamed provision'}'."
            )
        }

    proposal_benefit = _proposal_benefit_code(
        str(proposal_fact.get("benefit_line") or ""),
        provision_name,
    )
    candidates = [
        fact
        for fact in booklet_facts
        if fact.get("provision_code") == provision_code
        and _benefits_compatible(proposal_benefit, str(fact.get("benefit_code") or ""))
        and str(fact.get("provision_value") or "").strip()
    ]
    exact_candidates = [
        fact for fact in candidates if proposal_benefit and fact.get("benefit_code") == proposal_benefit
    ]
    if exact_candidates:
        candidates = exact_candidates

    if not candidates:
        return {
            "review_message": (
                f"No structured booklet value was found for "
                f"'{proposal_fact.get('benefit_line') or 'Plan Design'} - {provision_name or 'Provision'}'."
            )
        }

    comparable_values: dict[str, list[dict[str, Any]]] = {}
    for fact in candidates:
        comparable_values.setdefault(_comparable_value(fact.get("provision_value")), []).append(fact)
    comparable_values.pop("", None)
    if len(comparable_values) != 1:
        return {
            "review_message": (
                f"The booklet contains multiple class or benefit values for "
                f"'{proposal_fact.get('benefit_line') or 'Plan Design'} - {provision_name or 'Provision'}'; "
                "a single comparison would be ambiguous."
            )
        }

    matched_facts = next(iter(comparable_values.values()))
    booklet_fact = matched_facts[0]
    proposal_value = str(proposal_fact.get("provision_value") or "").strip()
    booklet_value = str(booklet_fact.get("provision_value") or "").strip()
    values_match = _values_match(proposal_value, booklet_value)

    if not values_match and not _is_structured_provision(provision_code):
        return {
            "review_message": (
                f"The proposal and booklet contain different free-text coverage wording for "
                f"'{proposal_fact.get('benefit_line') or 'Plan Design'} - {provision_name or 'Provision'}'; "
                "manual review is required."
            )
        }

    booklet_classes = sorted(
        {
            str(fact.get("class_name") or "").strip()
            for fact in matched_facts
            if str(fact.get("class_name") or "").strip()
        },
        key=str.casefold,
    )
    validation_status = "MATCH" if values_match else "DIFFERENT"
    validation = {
        **proposal_fact,
        "benefit_line": proposal_fact.get("benefit_line"),
        "provision_name": proposal_fact.get("provision_name"),
        "provision_value": proposal_value,
        "booklet_value": booklet_value,
        "variance": None if values_match else "Proposal value differs from the extracted booklet value.",
        "validation_status": validation_status,
        "status": "VALID" if values_match else "NEEDS REVIEW",
        "confidence_score": _lower_confidence(
            proposal_fact.get("confidence_score"),
            booklet_fact.get("confidence_score"),
        ),
        "booklet_source_file": booklet_fact.get("source_file"),
        "booklet_source_page": booklet_fact.get("source_page"),
        "booklet_source_section": booklet_fact.get("source_section"),
        "booklet_class_names": booklet_classes,
        "booklet_carrier": booklet_fact.get("carrier"),
        "booklet_confidence_score": booklet_fact.get("confidence_score"),
        "comparison_method": "Structured booklet plan design field",
        "validation_source": _VALIDATOR_SOURCE,
    }
    return {"validation": validation}


def _record_booklet_extraction_failure(
    envelope: dict[str, Any],
    booklet_path: str | Path | None,
    message: str,
) -> dict[str, Any]:
    proposals = envelope.get("proposals") or []
    for proposal in proposals:
        proposal["booklet_validation"] = []
        proposal["exceptions"] = [
            row
            for row in proposal.get("exceptions", [])
            if row.get("validation_source") != _VALIDATOR_SOURCE
        ]
        _append_validation_exception(
            proposal,
            None,
            f"Benefits booklet validation could not run: {message}",
            source_file=str(booklet_path or ""),
        )
    envelope["_booklet_validation_summary"] = {
        "status": "NEEDS REVIEW",
        "booklet_path": str(booklet_path or ""),
        "booklet_facts": 0,
        "proposal_facts": sum(len(proposal.get("plan_design") or []) for proposal in proposals),
        "validation_rows": 0,
        "matches": 0,
        "differences": 0,
        "review_issues": len(proposals),
        "message": message,
    }
    return envelope


def _append_validation_exception(
    proposal: dict[str, Any],
    proposal_fact: dict[str, Any] | None,
    message: str,
    source_file: str = "",
) -> None:
    fact = proposal_fact or {}
    summary = proposal.get("proposal_summary") or {}
    proposal.setdefault("exceptions", []).append(
        {
            "exception_type": "NEEDS REVIEW",
            "field_name": "proposal_booklet_validation",
            "message": message,
            "benefit_line": fact.get("benefit_line"),
            "provision_name": fact.get("provision_name"),
            "source_file": source_file or fact.get("source_file") or summary.get("source_file"),
            "source_page": fact.get("source_page"),
            "source_sheet": fact.get("source_sheet"),
            "confidence_score": fact.get("confidence_score") or "LOW",
            "status": "NEEDS REVIEW",
            "validation_source": _VALIDATOR_SOURCE,
        }
    )


def _proposal_benefit_code(benefit_line: str, provision_name: str) -> str:
    provision_benefit = _benefit_code(provision_name)
    if provision_benefit and provision_benefit not in {"life", "health"}:
        return provision_benefit
    return _benefit_code(benefit_line) or provision_benefit


def _benefit_code(value: Any) -> str:
    normalized = _normalize(str(value or "")).replace("&", " and ")
    for code, aliases in _BENEFIT_ALIASES:
        if any(_contains_phrase(normalized, alias.replace("&", " and ")) for alias in aliases):
            return code
    return ""


def _benefits_compatible(proposal_code: str, booklet_code: str) -> bool:
    if not proposal_code or not booklet_code:
        return False
    if proposal_code == booklet_code:
        return True
    return proposal_code == "health" and booklet_code in _HEALTH_BENEFITS


def _provision_code(value: str, provision_value: str = "") -> str:
    normalized = _normalize(value)
    dental_code = canonical_dental_provision_code(value)
    if dental_code and (
        _has_percentage(provision_value)
        or any(token in normalized for token in ("coinsurance", "reimbursement"))
    ):
        return dental_coinsurance_provision_code(dental_code)
    if dental_code and any(token in normalized for token in ("maximum", "plan max", "benefit max")):
        return dental_maximum_provision_code(dental_code)
    if any(token in normalized for token in ("coinsurance", "reimbursement")):
        return "coinsurance"
    if any(token in normalized for token in ("termination", "coverage ends", "expiry age")):
        return "termination_age"
    if any(token in normalized for token in ("waiting period", "elimination period", "qualification period")):
        return "waiting_period"
    if any(token in normalized for token in ("frequency", "recall")):
        return "frequency"
    if any(token in normalized for token in ("maximum", "plan max", "benefit max", "non evidence max")):
        return "maximum"
    if any(token in normalized for token in ("coverage", "benefit amount", "benefit schedule", "plan type", "drug type")):
        return "coverage"
    return ""


def _is_structured_provision(provision_code: str) -> bool:
    return (
        provision_code in _STRUCTURED_PROVISIONS
        or provision_code.startswith("coinsurance.dental_")
        or provision_code.startswith("maximum.dental_")
    )


def _has_percentage(value: str) -> bool:
    return re.search(
        r"(?<![\d.])(?:100(?:\.0+)?|\d{1,2}(?:\.\d+)?)\s*(?:%|percent|per\s+cent)(?![a-z0-9])",
        value or "",
        flags=re.IGNORECASE,
    ) is not None


def _values_match(left: Any, right: Any) -> bool:
    left_value = _comparable_value(left)
    right_value = _comparable_value(right)
    return bool(left_value) and left_value == right_value


def _comparable_value(value: Any) -> str:
    text = _normalize(str(value or ""))
    text = text.replace("per cent", "%").replace("percent", "%")
    text = text.replace("calendar year", "year").replace("per year", "/year")
    text = text.replace("annually", "/year").replace("each year", "/year")
    text = re.sub(r"\$\s*([0-9])", r"\1", text)
    text = re.sub(r"(?<=\d),(?=\d)", "", text)
    text = re.sub(r"\s+%", "%", text)
    text = re.sub(r"\s*/\s*", "/", text)
    return re.sub(r"[^a-z0-9%/]+", " ", text).strip()


def _contains_phrase(text: str, phrase: str) -> bool:
    return re.search(rf"(?<![a-z0-9]){re.escape(phrase)}(?![a-z0-9])", text) is not None


def _lower_confidence(left: Any, right: Any) -> str:
    ranks = {"LOW": 1, "MEDIUM": 2, "HIGH": 3}
    values = [str(value or "").strip().upper() for value in (left, right)]
    known = [value for value in values if value in ranks]
    return min(known, key=lambda value: ranks[value]) if known else "LOW"


def _positive_int(value: Any) -> int | None:
    try:
        number = int(value)
        return number if number > 0 else None
    except (TypeError, ValueError):
        return None


def _find_booklet_snippet(booklet_text: str, provision_name: str) -> str | None:
    index = booklet_text.find(provision_name)
    if index < 0:
        return None
    return booklet_text[index:index + 240]


def _normalize(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip().lower()


def _append_note(existing: str | None, addition: str) -> str:
    return f"{existing} {addition}".strip() if existing else addition
