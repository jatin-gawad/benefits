from __future__ import annotations

from copy import deepcopy
from typing import Any


SUMMARY_FIELDS = [
    "client_name",
    "carrier_name",
    "proposal_date",
    "effective_date",
    "source_file",
    "detected_carrier",
    "detection_confidence",
]

BENEFIT_LINE_FIELDS = [
    "carrier_name",
    "benefit_line",
    "class_name",
    "tier",
    "lives",
    "volume",
    "rate",
    "monthly_premium",
    "annual_premium",
    "rate_basis",
    "volume_basis",
    "source_file",
    "source_page",
    "source_sheet",
    "confidence_score",
    "status",
    "notes",
]

BENEFIT_RECORD_FIELDS = [
    "carrier_name",
    "benefit_key",
    "benefit_name",
    "normalized_benefit_name",
    "display_name",
    "benefit_category",
    "sort_order",
    "is_standard",
    "is_known",
    "source_file",
    "source_page",
    "source_sheet",
    "confidence_score",
    "status",
    "notes",
]

RATE_RECORD_FIELDS = [
    "carrier_name",
    "benefit_line",
    "class_name",
    "tier",
    "age_band",
    "gender",
    "province",
    "lives",
    "volume",
    "rate",
    "rate_basis",
    "monthly_premium",
    "annual_premium",
    "source_file",
    "source_page",
    "source_sheet",
    "confidence_score",
    "status",
    "notes",
]

PREMIUM_RECORD_FIELDS = [
    "carrier_name",
    "benefit_line",
    "class_name",
    "tier",
    "monthly_premium",
    "annual_premium",
    "employer_cost",
    "employee_cost",
    "total_cost",
    "source_file",
    "source_page",
    "source_sheet",
    "confidence_score",
    "status",
    "notes",
]

PLAN_DESIGN_FIELDS = [
    "carrier_name",
    "benefit_line",
    "provision_name",
    "provision_value",
    "booklet_value",
    "variance",
    "validation_status",
    "source_file",
    "source_page",
    "source_sheet",
    "confidence_score",
    "status",
    "notes",
]

DEMOGRAPHIC_FIELDS = [
    "carrier_name",
    "benefit_line",
    "class_name",
    "tier",
    "lives",
    "volume",
    "source_file",
    "source_page",
    "source_sheet",
    "confidence_score",
    "status",
    "notes",
]

ASSUMPTION_FIELDS = [
    "carrier_name",
    "assumption_name",
    "assumption_value",
    "source_file",
    "source_page",
    "source_sheet",
    "confidence_score",
    "status",
    "notes",
]

POOLING_RETENTION_FIELDS = [
    "carrier_name",
    "benefit_line",
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
    "source_file",
    "source_page",
    "source_sheet",
    "confidence_score",
    "status",
    "notes",
]

EXCEPTION_FIELDS = [
    "carrier_name",
    "exception_type",
    "field_name",
    "message",
    "source_file",
    "source_page",
    "source_sheet",
    "confidence_score",
    "status",
    "notes",
]


def empty_normalized_proposal() -> dict[str, Any]:
    return {
        "proposal_summary": {field: None for field in SUMMARY_FIELDS},
        "benefits": [],
        "benefit_lines": [],
        "rates": [],
        "premiums": [],
        "plan_design": [],
        "demographics": [],
        "underwriting_assumptions": [],
        "pooling_retention": [],
        "exceptions": [],
        "sqlite_audit": [],
    }


def normalize_proposal_data(raw_data: dict[str, Any], carrier_config: dict[str, Any]) -> dict[str, Any]:
    normalized = empty_normalized_proposal()
    summary = deepcopy(raw_data.get("proposal_summary", {}))
    summary.setdefault("carrier_name", carrier_config.get("carrier_name"))
    summary.setdefault("detected_carrier", carrier_config.get("carrier_key"))
    normalized["proposal_summary"] = _normalize_record(summary, SUMMARY_FIELDS)

    list_mappings = {
        "benefits": BENEFIT_RECORD_FIELDS,
        "benefit_lines": BENEFIT_LINE_FIELDS,
        "rates": RATE_RECORD_FIELDS,
        "premiums": PREMIUM_RECORD_FIELDS,
        "plan_design": PLAN_DESIGN_FIELDS,
        "demographics": DEMOGRAPHIC_FIELDS,
        "underwriting_assumptions": ASSUMPTION_FIELDS,
        "pooling_retention": POOLING_RETENTION_FIELDS,
        "exceptions": EXCEPTION_FIELDS,
    }

    for key, fields in list_mappings.items():
        normalized[key] = [_normalize_record(item, fields) for item in raw_data.get(key, [])]

    return normalized


def combine_proposals(proposals: list[dict[str, Any]], client_id: str | None = None) -> dict[str, Any]:
    combined = empty_normalized_proposal()
    combined["client_id"] = client_id
    combined["proposals"] = proposals
    combined["proposal_summaries"] = [proposal.get("proposal_summary", {}) for proposal in proposals]

    for proposal in proposals:
        for key in [
            "benefits",
            "benefit_lines",
            "rates",
            "premiums",
            "plan_design",
            "demographics",
            "underwriting_assumptions",
            "pooling_retention",
            "exceptions",
            "sqlite_audit",
        ]:
            combined[key].extend(deepcopy(proposal.get(key, [])))

    if proposals:
        combined["proposal_summary"] = deepcopy(proposals[0].get("proposal_summary", {}))
    return combined


def _normalize_record(record: dict[str, Any], fields: list[str]) -> dict[str, Any]:
    result = {field: record.get(field) for field in fields}
    result.setdefault("status", "NEEDS REVIEW")
    result.setdefault("confidence_score", "LOW")
    return result
