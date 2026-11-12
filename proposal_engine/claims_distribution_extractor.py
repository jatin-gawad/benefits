"""Carrier-agnostic claims distribution table extraction helper.

This module is intentionally small and structure-based. The desktop app uses the
C# extractor for the live Renewal Review workflow; this Python helper mirrors the
same table shape for future scripting or spot checks.
"""

from __future__ import annotations

from dataclasses import dataclass, asdict
import re
from pathlib import Path
from typing import Iterable


TITLE_CLUES = (
    "Claims Distribution by Service",
    "Claims Distribution by Service Summary",
    "Medical Claims Distribution",
    "Service Summary",
    "Health Service Utilization",
    "Benefit Utilization Summary",
    "Claims Breakdown by Category",
    "Claims by Service Category",
)

ROW_RE = re.compile(
    r"^(?P<service>[A-Za-z][A-Za-z ,/&().+-]+?)\s+"
    r"\$?(?P<amount>-?[\d,]+(?:\.\d+)?)\s+"
    r"(?P<paid_pct>-?[\d.]+)%\s+"
    r"(?P<norm_paid_pct>-?[\d.]+)%\s+"
    r"(?P<occurrences>\d+)\s+"
    r"(?P<occ_pct>-?[\d.]+)%\s+"
    r"(?P<norm_occ_pct>-?[\d.]+)%\s+"
    r"\$?(?P<avg>-?[\d,]+(?:\.\d+)?)$",
    re.IGNORECASE,
)


@dataclass
class ClaimsDistributionRow:
    benefit: str
    service_category: str
    service_description_original: str
    amount_paid: float | None
    percent_paid: float | None
    norm_percent_paid: float | None
    occurrences: int | None
    percent_occurrences: float | None
    norm_percent_occurrences: float | None
    avg_claim: float | None
    source: str
    carrier: str = ""
    policy_number: str = ""
    report_period: str = ""
    division: str = "ALL"
    class_name: str = "ALL"


def extract_claims_distribution_from_lines(
    pages: Iterable[tuple[int, Iterable[str]]],
    source_file: str,
    carrier_name: str = "",
) -> list[dict[str, object | None]]:
    rows: list[ClaimsDistributionRow] = []
    for page_number, page_lines in pages:
        lines = [_clean(line) for line in page_lines if _clean(line)]
        page_text = " ".join(lines)
        if not _looks_like_claims_distribution(page_text):
            continue

        metadata = _metadata(lines)
        for line in lines:
            match = ROW_RE.match(line)
            if not match:
                continue

            original = _clean(match.group("service"))
            rows.append(
                ClaimsDistributionRow(
                    benefit=metadata.get("Benefit", "Medical"),
                    service_category=normalize_service_category(original),
                    service_description_original=original,
                    amount_paid=_number(match.group("amount")),
                    percent_paid=_number(match.group("paid_pct")),
                    norm_percent_paid=_number(match.group("norm_paid_pct")),
                    occurrences=_integer(match.group("occurrences")),
                    percent_occurrences=_number(match.group("occ_pct")),
                    norm_percent_occurrences=_number(match.group("norm_occ_pct")),
                    avg_claim=_number(match.group("avg")),
                    source=f"{carrier_name or Path(source_file).stem} PDF page {page_number}",
                    carrier=carrier_name,
                    policy_number=metadata.get("Policy(s)", metadata.get("Policy", "")),
                    report_period=metadata.get("Period", ""),
                    division=metadata.get("Division(s)", metadata.get("Division", "ALL")),
                    class_name=metadata.get("Class(s)", metadata.get("Class", "ALL")),
                )
            )

    return [asdict(row) for row in rows]


def normalize_service_category(value: str) -> str:
    key = re.sub(r"[^a-z0-9]", "", value.lower())
    if key == "total":
        return "Total"
    if key in {"drug", "drugs", "prescriptiondrugs", "prescriptionmedication"}:
        return "Drugs"
    if key in {
        "paramedical",
        "paramedicalservices",
        "practitionerservices",
        "healthpractitioners",
        "professionalservices",
        "alliedhealth",
        "massagetherapy",
        "physiotherapy",
        "chiropractic",
        "psychology",
        "naturopathic",
        "acupuncture",
    }:
        return "Paramedical"
    if key in {"serviceandsupplies", "medicalsupplies", "supplies", "medicalservicesandsupplies"}:
        return "Medical Supplies"
    if key in {"outofcountry", "travel", "emergencytravel"}:
        return "Out of Country"
    if key in {"vision", "visioncare", "eyeglasses", "eyeexams"}:
        return "Vision"
    if key in {"dental", "basicdental", "majordental", "orthodontics"}:
        return "Dental"
    if key == "hospital":
        return "Hospital"
    if key == "ambulance":
        return "Ambulance"
    if key == "outofprovince":
        return "Out of Province"
    if key in {"other", "miscellaneous"}:
        return "Other"
    return _clean(value)


def _looks_like_claims_distribution(text: str) -> bool:
    return any(clue.lower() in text.lower() for clue in TITLE_CLUES) or (
        "amount paid" in text.lower()
        and "number of occurrences" in text.lower()
        and "average amount paid" in text.lower()
    )


def _metadata(lines: Iterable[str]) -> dict[str, str]:
    result: dict[str, str] = {"Benefit": "Medical", "Division(s)": "ALL", "Class(s)": "ALL"}
    for line in lines:
        match = re.match(r"^(Policy\(s\)|Policy|Period|Benefit|Division\(s\)|Division|Class\(s\)|Class)\s+(.+)$", line, re.I)
        if match:
            result[_clean(match.group(1))] = _clean(match.group(2))
    return result


def _number(value: str) -> float | None:
    cleaned = value.replace("$", "").replace(",", "").replace("%", "").strip()
    return float(cleaned) if cleaned else None


def _integer(value: str) -> int | None:
    cleaned = value.replace(",", "").strip()
    return int(cleaned) if cleaned else None


def _clean(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()
