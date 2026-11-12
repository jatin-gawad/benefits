from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from collections import defaultdict
from dataclasses import asdict, dataclass
from decimal import Decimal, InvalidOperation
from datetime import datetime
from pathlib import Path
from typing import Iterable

try:
    from proposal_engine.booklet_plan_design.labeled_dental_rate_mapper import (
        CANONICAL_LABELED_DENTAL_RATE_MAPPER_VERSION,
        canonical_dental_provision_code,
        map_labeled_dental_coverage_statuses,
        map_labeled_dental_maximums,
        map_labeled_dental_rates,
    )
except ModuleNotFoundError:  # Direct command-file execution starts in this script's folder.
    module_folder = str(Path(__file__).resolve().parent)
    if module_folder not in sys.path:
        sys.path.insert(0, module_folder)
    from labeled_dental_rate_mapper import (
        CANONICAL_LABELED_DENTAL_RATE_MAPPER_VERSION,
        canonical_dental_provision_code,
        map_labeled_dental_coverage_statuses,
        map_labeled_dental_maximums,
        map_labeled_dental_rates,
    )

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.drawing.image import Image as OpenPyxlImage
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError as exc:  # pragma: no cover - handled for user-friendly CLI output
    raise SystemExit(
        "Missing dependency: openpyxl is required to update Benefits_Master_Database.xlsx."
    ) from exc


MODULE_NAME = "Benefit Booklet Summary Builder"
PLAN_SHEET = "PlanDesign"
LOG_SHEET = "ProcessingLog"
REVIEW_SHEET = "Booklet_Extraction_Review"
SUMMARY_SHEET = "Employee_Benefits_Summary"
SCHEDULE_SHEET = "Schedule_of_Benefits"
PPT_EXPORT_SHEET = "Schedule_PPT_Export"
PPT_EXPORT_TABLE = "tblPPT_ScheduleOfBenefits"
SUMMARY_LOGO_FILE_NAME = "WestlandBenefitsHorizontalColour.png"
SUMMARY_LOGO_START_ROW = 6
DEFAULT_RULES_FILE = "booklet_plan_design_rules.json"
PAGE_MARKER_PATTERN = re.compile(r"\[\[PAGE\s+(\d+)\]\]")
MONEY_AMOUNT_PATTERN = r"\$\s?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?"
LUMP_SUM_BENEFITS = {"Life", "Life / AD&D", "AD&D", "Dependent Life", "Critical Illness"}
FULL_PAYOUT_BENEFITS = {"Life", "Life / AD&D", "AD&D", "Dependent Life"}
SCHEDULE_MAX_DETAIL_LINES = 24
SUMMARY_DISPLAY_ORDER = [
    "Carrier & Policy",
    "Eligibility Requirements",
    "Class Breakdown",
    "Life / AD&D",
    "Dependent Life",
    "Long-Term Disability",
    "Critical Illness",
    "Prescription Drugs",
    "Medical Services & Supplies",
    "Paramedicals",
    "Mental Health Practitioners",
    "Hospital",
    "Vision Care",
    "Emergency Travel",
    "Dental Care",
    "Survivor Benefit",
    "Health Care Spending Account",
    "Lifestyle Spending Account",
    "Employee Assistance Program",
    "Virtual Care",
    "Medical Second Opinion",
]
SUMMARY_ORDER_INDEX = {name: index for index, name in enumerate(SUMMARY_DISPLAY_ORDER)}

BENEFIT_PROVIDER_PATTERNS: dict[str, tuple[tuple[str, str], ...]] = {
    "Critical Illness": (
        ("Allstate", r"\bAllstate\b"),
        ("Manulife", r"\bManulife\b"),
        ("Canada Life", r"\bCanada Life\b"),
        ("Sun Life", r"\bSun Life\b"),
        ("Equitable", r"\bEquitable(?: Life)?\b"),
        ("Industrial Alliance", r"\b(?:Industrial Alliance|iA Financial Group)\b"),
    ),
    "Employee Assistance Program": (
        ("Homewood Health", r"\bHomewood Health\b"),
        ("TELUS Health", r"\bTELUS Health\b"),
        ("LifeWorks", r"\bLifeWorks\b"),
        ("Dialogue", r"\bDialogue(?: Health Technologies Inc\.)?\b"),
        ("GreenShield", r"\b(?:GreenShield|Green Shield)\b"),
    ),
    "Health Care Spending Account": (
        ("National HealthClaim", r"\b(?:National HealthClaim|National Health Claim|NHC)\b"),
    ),
    "Lifestyle Spending Account": (
        ("National HealthClaim", r"\b(?:National HealthClaim|National Health Claim|NHC)\b"),
    ),
    "Medical Second Opinion": (
        ("WorldCare", r"\bWorld\s*Care\b"),
        ("Advica Health", r"\bAdvica(?:\s+Health)?\b"),
        ("Teladoc Health", r"\bTeladoc(?:\s+(?:Health|Medical Experts))?\b"),
        ("Dialogue", r"\bDialogue(?:\s+(?:Health(?:\s+Technologies\s+Inc\.?)?|Medical\s+Second\s+Opinion))?\b"),
        ("Sparrow", r"\bSparrow\b"),
    ),
}

PLAN_COLUMNS = [
    "ClientID",
    "ClientName",
    "PolicyNumber",
    "Carrier",
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
    "SourceFile",
    "SourceSection",
    "SourcePage",
    "LastUpdated",
]

LOG_COLUMNS = [
    "ProcessedAt",
    "ClientName",
    "Carrier",
    "SourceFile",
    "Module",
    "Status",
    "RowsAdded",
    "RowsUpdated",
    "ErrorMessage",
]

REQUESTED_REVIEW_COLUMNS = [
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
]

LEGACY_REVIEW_COLUMNS = [
    "Benefit",
    "CoverageDescription",
    "Coinsurance",
    "Maximum",
    "Frequency",
    "WaitingPeriod",
    "TerminationAge",
]

BENCHMARK_COLUMNS = [
    "LifeBenefitAmount",
    "LifeEarningsMultiplier",
    "ADDBenefitAmount",
    "DependentLifeSpouseAmount",
    "DependentLifeChildAmount",
    "CriticalIllnessAmount",
    "CriticalIllnessProvider",
    "CriticalIllnessPlanType",
    "CriticalIllnessEmployeeAmount",
    "CriticalIllnessSpouseAmount",
    "CriticalIllnessChildAmount",
    "CriticalIllnessEvidenceRequirement",
    "CriticalIllnessEvidenceThreshold",
    "CriticalIllnessOptionalBenefits",
    "CriticalIllnessRiders",
    "CriticalIllnessSingleMonthlyRate",
    "CriticalIllnessFamilyMonthlyRate",
    "CriticalIllnessRateGuaranteeDate",
    "CriticalIllnessPreEx",
    "CriticalIllnessTerminationAge",
    "PolicyholderLegalName",
    "PolicyEffectiveDate",
    "PolicyAnniversaryDate",
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
    "TravelMaximum",
    "TravelPerTripMaximum",
    "TravelTripDurationDays",
    "EAPIncluded",
    "VirtualCareIncluded",
    "SecondOpinionIncluded",
]

REVIEW_COLUMNS = (
    REQUESTED_REVIEW_COLUMNS
    + [column for column in LEGACY_REVIEW_COLUMNS if column not in REQUESTED_REVIEW_COLUMNS]
    + [column for column in BENCHMARK_COLUMNS if column not in REQUESTED_REVIEW_COLUMNS]
)


@dataclass(frozen=True)
class BenefitDefinition:
    name: str
    keywords: tuple[str, ...]
    headings: tuple[str, ...] = ()
    page_span: int = 2


@dataclass(frozen=True)
class SummaryFieldDefinition:
    label: str
    patterns: tuple[str, ...]
    max_results: int = 1


@dataclass(frozen=True)
class SummarySectionDefinition:
    name: str
    keywords: tuple[str, ...]
    headings: tuple[str, ...]
    fields: tuple[SummaryFieldDefinition, ...]
    source: str = "section"
    page_span: int = 2
    optional: bool = False


@dataclass(frozen=True)
class ExtractionRules:
    benefits: tuple[BenefitDefinition, ...]
    carrier_patterns: tuple[str, ...]
    client_rejected_names: tuple[str, ...]
    front_matter_markers: tuple[str, ...]
    field_patterns: dict[str, tuple[str, ...]]
    summary_sections: tuple[SummarySectionDefinition, ...]


@dataclass
class ClientInfo:
    client_id: str
    client_name: str
    carrier: str = ""
    group_number: str = ""
    policy_number: str = ""


@dataclass
class ExtractionRow:
    client_id: str
    client_name: str
    carrier: str
    class_name: str
    benefit: str
    coverage_description: str
    coinsurance: str
    maximum: str
    frequency: str
    waiting_period: str
    termination_age: str
    notes: str
    source_file: str
    source_section: str
    confidence: str
    review_flag: str
    last_updated: str
    policy_number: str = ""
    source_page: int | str = ""
    benefit_category: str = ""


@dataclass(frozen=True)
class ScheduleLifeDetails:
    """Source-traceable Group Life values shared by extraction and summaries."""

    benefit_amount: str
    maximum_benefit: str
    benefit_reduction: str
    waiver_of_premium: str
    termination_age: str
    adnd_amount: str
    source_page: int


@dataclass(frozen=True)
class ScheduleDependentLifeDetails:
    """Source-traceable Dependent Life values from a Schedule of Benefits."""

    spouse_amount: str
    dependent_child_amount: str
    termination_age: str
    source_page: int
    dependent_child_eligibility: str = ""
    stillbirth_benefit: str = ""


@dataclass(frozen=True)
class ScheduleAsiDetails:
    """Source-traceable GroupHEALTH Accident & Serious Illness values."""

    benefit_amount: str
    critical_disease_benefit: str
    serious_illness_benefit: str
    termination_age: str
    source_page: int


@dataclass(frozen=True)
class ScheduleCriticalIllnessDetails:
    """Source-traceable Group Critical Illness values from the schedule."""

    plan_type: str
    benefit_amount: str
    guaranteed_issue_limit: str
    insured_illnesses: str
    tax_status: str
    termination_age: str
    provider: str
    policy_number: str
    source_page: int


@dataclass(frozen=True)
class ScheduleLtdDetails:
    """Source-traceable LTD values shared by extraction and summaries."""

    benefit_schedule: str
    maximum_benefit: str
    non_evidence_maximum: str
    tax_status: str
    qualification_period: str
    disability_definition: str
    benefit_period: str
    termination_age: str
    source_page: int


@dataclass(frozen=True)
class ScheduleStdDetails:
    """Source-traceable STD values shared by extraction and summaries."""

    benefit_schedule: str
    maximum_benefit: str
    non_evidence_maximum: str
    tax_status: str
    qualification_period: str
    disability_definition: str
    benefit_period: str
    termination_age: str
    source_page: int


@dataclass(frozen=True)
class ScheduleEhcDetails:
    """Source-traceable EHC schedule values shared by extraction and summaries."""

    deductible: str
    prescription_drugs_listed: bool
    prescription_drug_coinsurance: str
    prescription_drug_plan_type: str
    prescription_drug_maximum: str
    prescription_drug_type: str
    other_medical_listed: bool
    other_medical_coinsurance: str
    health_care_practitioners_listed: bool
    health_care_practitioners_coinsurance: str
    health_care_practitioners_maximum: str
    health_care_practitioners: tuple[str, ...]
    vision_listed: bool
    vision_coinsurance: str
    vision_coverage_status: str
    vision_maximum: str
    eye_examination_benefit: str
    hearing_aid_benefit: str
    orthotics_benefit: str
    travel_listed: bool
    travel_coinsurance: str
    travel_referral_maximum: str
    travel_maximum: str
    travel_duration: str
    travel_termination_age: str
    travel_details_source_page: int
    travel_cancellation_listed: bool
    travel_cancellation_coinsurance: str
    travel_cancellation_maximum: str
    source_page: int


@dataclass(frozen=True)
class DesjardinsPlanDesignItem:
    """A source-traceable Desjardins table item used by both workbook outputs."""

    summary_section: str
    benefit: str
    coverage_description: str
    coinsurance: str
    maximum: str
    frequency: str
    source_section: str
    source_page: int


@dataclass(frozen=True)
class SunLifePlanDesignItem:
    """A source-traceable Sun Life Benefit Summary item shared by both outputs."""

    summary_section: str
    benefit: str
    coverage_description: str
    coinsurance: str
    maximum: str
    frequency: str
    waiting_period: str
    termination_age: str
    source_section: str
    source_page: int


@dataclass
class ProcessingError:
    source_file: str
    client_name: str = ""
    carrier: str = ""
    error_message: str = ""


@dataclass
class ProcessedSource:
    source_file: str
    client_name: str
    carrier: str
    status: str


@dataclass
class WorkbookUpdateResult:
    rows_added: int
    rows_updated: int
    review_path: str
    master_path: str


@dataclass
class EmployeeBenefitsSummary:
    source_file: str
    client_name: str
    carrier: str
    class_name: str
    rows: list[tuple[str, str]]


@dataclass
class BenefitAdministrationExtraction:
    source_file: str
    client_name: str
    benefit_display_name: str
    insurer_display_name: str
    service_provider_display_names: list[str]
    policy_number: str
    product_name: str
    is_optional: bool
    source_page_number: int
    source_text: str
    extraction_confidence: float


@dataclass
class ScheduleBenefitPage:
    summary: EmployeeBenefitsSummary
    page_number: int
    page_total: int
    sections: list[tuple[str, str]]


BENEFITS = [
    BenefitDefinition("Life", ("Life Insurance", "Employee Life", "Basic Life")),
    BenefitDefinition("AD&D", ("Accidental Death and Dismemberment", "AD&D", "Accidental Death")),
    BenefitDefinition("Critical Illness", ("Critical Illness", "Critical Illness (CI)", "CI")),
    BenefitDefinition("Dependent Life", ("Dependent Life", "Spouse Life", "Child Life")),
    BenefitDefinition("LTD", ("Long Term Disability", "Long-Term Disability", "LTD")),
    BenefitDefinition("STD", ("Short Term Disability", "Short-Term Disability", "STD", "Weekly Indemnity")),
    BenefitDefinition(
        "Drugs",
        (
            "Prescription Drugs",
            "EHC - Drugs",
            "Direct Drugs",
            "Drug Benefit",
            "Pay Direct Drug Card",
            "Generic Substitution",
            "Dispensing Fee",
        ),
    ),
    BenefitDefinition(
        "Medical Services & Supplies",
        (
            "Medical Services & Supplies",
            "Medical Supplies & Services",
            "Other Medical Expenses",
            "Other Medical",
        ),
    ),
    BenefitDefinition(
        "Paramedical",
        (
            "Paramedical",
            "Physiotherapist",
            "Massage Therapist",
            "Chiropractor",
            "Psychologist",
            "Social Worker",
            "Counsellor",
            "Dietician",
            "Speech Therapist",
            "Audiologist",
            "Naturopath",
            "Acupuncturist",
        ),
    ),
    BenefitDefinition("Vision", ("Vision Care", "Eye Exam", "Eyeglasses", "Contact Lenses")),
    BenefitDefinition("Dental Basic", ("Dental Care", "Basic Services", "Preventive", "Restorative")),
    BenefitDefinition("Dental Major", ("Major Services", "Crowns", "Dentures", "Bridgework")),
    BenefitDefinition("Dental Orthodontic", ("Orthodontics", "Orthodontic")),
    BenefitDefinition("Dental Recall", ("Recall Exam", "Recall", "Scaling", "Root Planing")),
    BenefitDefinition("Travel", ("Out of Country", "Emergency Travel Assistance", "Travel Assistance")),
    BenefitDefinition(
        "Survivor Benefit",
        ("Survivor Benefit", "Survivor Benefits", "Premium Waived"),
    ),
    BenefitDefinition(
        "Health Care Spending Account",
        (
            "Health Care Spending Account",
            "Health Spending Account",
            "Healthcare Spending Account",
            "HCSA",
            "HSCA",
        ),
    ),
    BenefitDefinition(
        "Lifestyle Spending Account",
        (
            "Lifestyle Spending Account",
            "Life Style Spending Account",
            "Wellness Spending Account",
            "LSA",
        ),
    ),
    BenefitDefinition("EAP", ("Employee Assistance Program", "EAP")),
    BenefitDefinition(
        "Virtual Care",
        ("Virtual Health Care", "Virtual Care", "Online Doctor", "Telemedicine", "Teladoc", "Maple"),
    ),
    BenefitDefinition(
        "Medical Second Opinion",
        (
            "Medical Second Opinion",
            "Expert Medical Opinion",
            "Best Doctors",
            "WorldCare",
            "Advica Health",
            "Teladoc Medical Experts",
            "Dialogue Medical Second Opinion",
        ),
    ),
]

CARRIER_PATTERNS = (
    "Allstate",
    "National HealthClaim",
    "Benepac",
    "BBD",
    "Manulife",
    "Sun Life",
    "Canada Life",
    "Green Shield",
    "GSC",
    "Blue Cross",
    "Pacific Blue Cross",
    "Desjardins",
    "Equitable",
    "Empire Life",
    "Beneva",
    "RBC Insurance",
    "Industrial Alliance",
    "iA Financial",
)

CLIENT_REJECTED_NAMES = (
    "manulife",
    "manulife financial",
    "benefits booklet",
    "table of contents",
    "introduction",
    "planning for financial security",
    "a message from your plan sponsor",
    "all employees",
    "all full-time employees",
)

FRONT_MATTER_MARKERS = (
    r"Your plan sponsor has chosen to offer the following benefits",
    r"Benefit Details\s+Your Plan's Coverage",
    r"Extended Health Care Benefit",
)

FIELD_PATTERNS = {
    "Coinsurance": (
        r"\b(?:50|60|70|75|80|85|90|100)\s*%",
    ),
    "Maximum": (
        r"\bBenefit Amount\s+(.{1,220}?)(?=\s+(?:Qualifying Period|Definition of Disability|Coverage Ends|Coverage ends|Waiver|Exclusions|$))",
        r"\bCombined Maximum\s+(.{1,220}?)(?=\s+(?:Maximum applies to|Level I|Coverage Ends|Coverage ends|Exclusions|$))",
        r"\bMaximum applies to\s+(.{1,260}?)(?=\s+(?:Level I|Level II|Level III|Level IV|Level V|Coverage Ends|Coverage ends|Exclusions|$))",
        r"\bMaximum\s+(.{1,180}?)(?=\s+(?:Deductible|Co-insurance|Coverage Ends|Coverage ends|Exclusions|$))",
        r"\$\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?\s+per\s+\d+\s+calendar\s+year\(s\)",
        r"\$\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?\s+per\s+calendar\s+year",
        r"\$\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?\s+per\s+lifetime",
        r"\$\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?\s+lifetime\s+maximum",
        r"\$\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?",
    ),
    "Frequency": (
        r"\bevery\s+(?:6|9|12|24)\s+months?\b",
        r"\bevery\s+\d+\s+calendar\s+years?\b",
        r"\bonce\s+every\s+(?:6|9|12|24)\s+months?\b",
        r"\bper\s+\d+\s+calendar\s+years?\b",
        r"\bper\s+calendar\s+year\b",
        r"\blifetime\b",
    ),
    "WaitingPeriod": (
        r"\bEligibility Period\s+(.{1,220}?)(?=\s+(?:Effective Date|Effective Date of Insurance|Eligibility Date|Hours Worked|Class Breakdown|Life|Group Life|Dependent Life|Critical Illness|Prescription Drugs|$))",
        r"\bWaiting Period\s+(.{1,80}?)(?=\s+(?:Deductible|Benefit Amount|Maximum|Coverage Ends|Coverage ends|$))",
        r"\bdate\s+of\s+hire\b",
        r"\bfirst\s+day\s+following\s+3\s+months\b",
        r"\b(?:3|6)\s+months?\b",
    ),
    "TerminationAge": (
        r"\bCoverage Ends?\s+(.{1,120}?)(?=\s+(?:Exclusions|Benefit Details|Waiver|$))",
        r"\bearlier\s+of\s+age\s+(?:65|70|85)\s+or\s+(?:your\s+)?retirement\b",
        r"\bage\s+(?:65|70|85)\b",
        r"\bretirement\b",
    ),
}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Extract plan design details from benefit booklet PDFs.")
    parser.add_argument(
        "--input",
        nargs="+",
        required=True,
        help="One or more booklet PDF files or folders containing booklet PDFs.",
    )
    parser.add_argument(
        "--master",
        default="",
        help="Path to Benefits_Master_Database.xlsx or .xlsm.",
    )
    parser.add_argument(
        "--review",
        default="",
        help="Optional path for the separate review workbook. Defaults beside the master workbook.",
    )
    parser.add_argument(
        "--rules",
        default="",
        help="Optional JSON extraction rules file. Defaults to booklet_plan_design_rules.json beside this script.",
    )
    parser.add_argument(
        "--flexible-benefit-summaries",
        default="",
        help="Optional local Windows OCR flexible-benefit summaries supplied by the desktop app.",
    )
    parser.add_argument(
        "--identity-only",
        action="store_true",
        help="Read only enough booklet text to identify the client for flat-folder discovery.",
    )
    parser.add_argument("--json", action="store_true", help="Print machine-readable run summary.")
    parser.add_argument("--progress", action="store_true", help="Print progress messages to stderr for the desktop app.")
    args = parser.parse_args(argv)

    try:
        if args.identity_only:
            summary = discover_booklet_client_identities(args.input, args.rules)
        else:
            if not args.master:
                parser.error("--master is required unless --identity-only is used.")
            summary = run_extraction(
                args.input,
                args.master,
                args.review,
                args.rules,
                flexible_benefit_summaries_path=args.flexible_benefit_summaries,
                progress=args.progress,
            )
    except Exception as exc:
        if args.json:
            print(json.dumps({"status": "Error", "error_message": str(exc)}, indent=2))
        else:
            print(f"Booklet extraction failed: {exc}", file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps(summary, indent=2))
    elif args.identity_only:
        print(f"Booklet client identification complete. PDFs checked: {summary['pdfs_found']}")
    else:
        print("Benefit booklet extraction complete.")
        print(f"Master workbook: {summary['master_path']}")
        print(f"Review workbook: {summary['review_path']}")
        print(f"PDFs processed: {summary['pdfs_processed']}")
        print(f"Rows added: {summary['rows_added']}")
        print(f"Rows updated: {summary['rows_updated']}")
        if summary["errors"]:
            print(f"Errors: {len(summary['errors'])}")

    return 0 if not summary["errors"] else 2


def discover_booklet_client_identities(
    input_paths: Iterable[str],
    rules_path: str = "",
) -> dict:
    pdf_paths = collect_pdf_paths(input_paths)
    rules = load_extraction_rules(rules_path)
    identities: list[dict[str, object]] = []
    errors: list[dict[str, str]] = []

    for pdf_path in pdf_paths:
        identity = {
            "sourcePath": str(pdf_path.resolve()),
            "sourceFile": pdf_path.name,
            "clientName": "",
            "documentClientName": "",
            "fileNameClientName": "",
            "carrier": "",
            "groupNumber": "",
            "policyNumber": "",
            "confidenceScore": 0.0,
            "identitySource": "Unconfirmed",
            "identityAgreement": "Not available",
            "errorMessage": "",
        }
        try:
            text = extract_pdf_text(pdf_path, max_pages=3, prefer_pypdf=True)
            searchable = text[:6000]
            direct_client_name = detect_client_name_for_identity(searchable, rules)
            structured_file_client_name = detect_structured_booklet_filename_client_name(pdf_path)
            identity["documentClientName"] = direct_client_name
            identity["fileNameClientName"] = structured_file_client_name
            client_info = detect_client_info(searchable, pdf_path, rules)
            identity["carrier"] = client_info.carrier
            identity["groupNumber"] = client_info.group_number
            identity["policyNumber"] = client_info.policy_number
            if direct_client_name and structured_file_client_name:
                if client_names_likely_match(direct_client_name, structured_file_client_name):
                    identity["clientName"] = direct_client_name
                    identity["confidenceScore"] = 0.98
                    identity["identitySource"] = "Document and structured file name"
                    identity["identityAgreement"] = "Matched"
                else:
                    identity["identitySource"] = "Conflict"
                    identity["identityAgreement"] = "Conflict"
                    identity["errorMessage"] = "The client name in the filename does not match the client name found inside the PDF."
            elif direct_client_name:
                identity["clientName"] = direct_client_name
                identity["confidenceScore"] = 0.95
                identity["identitySource"] = "Document"
                identity["identityAgreement"] = "PDF only"
            elif structured_file_client_name:
                identity["clientName"] = structured_file_client_name
                identity["confidenceScore"] = 0.90
                identity["identitySource"] = "Structured file name"
                identity["identityAgreement"] = "File name only"
            elif client_info.client_name:
                identity["clientName"] = client_info.client_name
                identity["confidenceScore"] = 0.40
                identity["identitySource"] = "File name fallback"
                identity["identityAgreement"] = "Unstructured file name only"
            else:
                identity["errorMessage"] = "No reliable client name was found in the booklet."
        except Exception as exc:
            message = str(exc)
            identity["errorMessage"] = message
            errors.append({"sourceFile": pdf_path.name, "message": message})
        identities.append(identity)

    return {
        "status": "Completed",
        "pdfs_found": len(pdf_paths),
        "identities": identities,
        "errors": errors,
    }


def detect_client_name_for_identity(text: str, rules: ExtractionRules) -> str:
    """Keep a labelled policyholder name intact for client grouping."""
    participant_name = extract_participant_cover_client_name(text, rules)
    if participant_name:
        return participant_name

    policyholder = first_regex(text, [r"\bPOLICYHOLDER\s*:\s*([^\n\r]{3,180})"])
    if policyholder and is_likely_client_name(policyholder, rules):
        return policyholder
    return detect_client_name(text, rules)


def detect_structured_booklet_filename_client_name(pdf_path: Path) -> str:
    parts = [part.strip() for part in re.split(r"__+", pdf_path.stem) if part.strip()]
    client_name = ""
    if len(parts) >= 5:
        client_code = parts[0]
        compact_client_name = parts[1]
        if re.fullmatch(r"[A-Za-z0-9]{4,12}", client_code) and any(
            part.lower() == "booklet" for part in parts[2:]
        ):
            client_name = humanize_compact_client_name(compact_client_name)

    if not client_name:
        policy_class_match = match_policy_class_filename(pdf_path)
        if policy_class_match:
            client_name = clean_client_name(policy_class_match.group("client"))

    if len(client_name) < 3 or not any(character.isalpha() for character in client_name):
        return ""
    return client_name


def match_policy_class_filename(pdf_path: Path) -> re.Match[str] | None:
    """Recognize carrier exports named '<client> <policy>_<class>.pdf'."""
    return re.fullmatch(
        r"(?P<client>.+?\S)\s+(?P<policy>\d{4,})(?:\s*)_(?P<class>[^_].+)",
        pdf_path.stem.strip(),
    )


def client_names_likely_match(left: str, right: str) -> bool:
    left_key = re.sub(r"[^a-z0-9]+", "", left.lower().replace("&", " and "))
    right_key = re.sub(r"[^a-z0-9]+", "", right.lower().replace("&", " and "))
    for left_candidate in {left_key, left_key.replace("and", "")}:
        for right_candidate in {right_key, right_key.replace("and", "")}:
            if left_candidate == right_candidate:
                return True

            shared_minimum_length = min(len(left_candidate), len(right_candidate))
            if shared_minimum_length >= 12 and (
                left_candidate in right_candidate or right_candidate in left_candidate
            ):
                return True

    return False


def run_extraction(
    input_paths: Iterable[str],
    master_path: str,
    review_path: str = "",
    rules_path: str = "",
    flexible_benefit_summaries_path: str = "",
    progress: bool = False,
) -> dict:
    pdf_paths = collect_pdf_paths(input_paths)
    if not pdf_paths:
        raise ValueError("No PDF files were found for the booklet extraction.")
    emit_progress(progress, f"BOOKLET_PROGRESS|TOTAL|{len(pdf_paths)}")

    rules = load_extraction_rules(rules_path)
    all_rows: list[ExtractionRow] = []
    benefit_administration_records: list[BenefitAdministrationExtraction] = []
    errors: list[ProcessingError] = []
    processed_sources: list[ProcessedSource] = []
    summaries: list[EmployeeBenefitsSummary] = []
    pdf_results: dict[str, dict[str, object]] = {}
    seen_pdf_hashes: dict[str, str] = {}
    flexible_benefit_summaries = load_flexible_benefit_summaries(flexible_benefit_summaries_path)

    for index, pdf_path in enumerate(pdf_paths, start=1):
        try:
            emit_progress(progress, f"BOOKLET_PROGRESS|START|{index}|{len(pdf_paths)}|{pdf_path.name}")
            pdf_hash = file_sha256(pdf_path)
            duplicate_of = seen_pdf_hashes.get(pdf_hash)
            if duplicate_of:
                processed_sources.append(
                    ProcessedSource(
                        source_file=pdf_path.name,
                        client_name="",
                        carrier="",
                        status=f"Skipped duplicate PDF of {duplicate_of}",
                    )
                )
                pdf_results[pdf_path.name] = {
                    "status": "Skipped duplicate PDF",
                    "duplicate_of": duplicate_of,
                    "rows_found": 0,
                }
                emit_progress(progress, f"BOOKLET_PROGRESS|SUCCESS|{index}|{len(pdf_paths)}|{pdf_path.name}")
                continue
            seen_pdf_hashes[pdf_hash] = pdf_path.name

            flexible_benefit_summary = flexible_benefit_summaries.get(pdf_path.name.lower())
            text = extract_pdf_text(pdf_path)
            if not clean_text(text):
                scanned_summary = build_scanned_flexible_benefits_summary(pdf_path, flexible_benefit_summary)
                if scanned_summary is None:
                    raise ValueError("No readable text was found in the PDF.")

                scanned_status = flexible_benefit_processing_status(flexible_benefit_summary)

                summaries.append(scanned_summary)
                processed_sources.append(
                    ProcessedSource(
                        source_file=pdf_path.name,
                        client_name="",
                        carrier="National HealthClaim",
                        status=scanned_status,
                    )
                )
                pdf_results[pdf_path.name] = {
                    "client_name": "",
                    "carrier": "National HealthClaim",
                    "class_name": scanned_summary.class_name,
                    "rows_found": 0,
                }
                if flexible_benefit_summary is None:
                    pdf_results[pdf_path.name]["warning"] = (
                        "Image-only source - local OCR review required before plan-design values can be saved."
                    )
                emit_progress(progress, f"BOOKLET_PROGRESS|SUCCESS|{index}|{len(pdf_paths)}|{pdf_path.name}")
                continue

            client_info = detect_client_info(text, pdf_path, rules)
            benefit_administration_records.extend(
                extract_grouphealth_benefit_administration(text, pdf_path, client_info)
            )
            class_names = detect_booklet_class_names(text, pdf_path, client_info)
            benefit_text = remove_front_matter_for_benefits(text, rules)
            plan_design_source_text = text if is_schedule_first_booklet(text, client_info) else benefit_text
            source_rows: list[ExtractionRow] = []
            for class_name in class_names:
                rows = extract_plan_design_rows(
                    plan_design_source_text,
                    pdf_path,
                    client_info,
                    rules,
                    class_name,
                    document_text=text,
                )
                summaries.append(build_employee_benefits_summary(text, benefit_text, pdf_path, client_info, rules, class_name))
                source_rows.extend(rows)

            scanned_summary = build_scanned_flexible_benefits_summary(pdf_path, flexible_benefit_summary)
            scanned_warning = ""
            if not source_rows and scanned_summary is not None:
                summaries.append(scanned_summary)
                scanned_warning = flexible_benefit_processing_status(flexible_benefit_summary)

            all_rows.extend(source_rows)
            processed_sources.append(
                ProcessedSource(
                    source_file=pdf_path.name,
                    client_name="" if scanned_warning else client_info.client_name,
                    carrier="National HealthClaim" if scanned_warning else client_info.carrier,
                    status=scanned_warning or ("Success" if source_rows else "No benefit rows found"),
                )
            )
            pdf_results[pdf_path.name] = {
                "client_name": "" if scanned_warning else client_info.client_name,
                "carrier": "National HealthClaim" if scanned_warning else client_info.carrier,
                "class_name": scanned_summary.class_name if scanned_warning else "; ".join(class_names),
                "rows_found": len(source_rows),
            }
            if scanned_warning and flexible_benefit_summary is None:
                pdf_results[pdf_path.name]["warning"] = (
                    "Image-only source - local OCR review required before plan-design values can be saved."
                )
            emit_progress(progress, f"BOOKLET_PROGRESS|SUCCESS|{index}|{len(pdf_paths)}|{pdf_path.name}")
        except Exception as exc:
            errors.append(ProcessingError(source_file=pdf_path.name, error_message=str(exc)))
            emit_progress(progress, f"BOOKLET_PROGRESS|ERROR|{index}|{len(pdf_paths)}|{pdf_path.name}|{exc}")

    result = update_workbooks(
        master_path=Path(master_path),
        review_path=Path(review_path) if review_path else None,
        rows=all_rows,
        summaries=summaries,
        processed_sources=processed_sources,
        errors=errors,
    )

    return {
        "status": "Complete" if not errors else "CompleteWithErrors",
        "module": MODULE_NAME,
        "rules_path": str(resolve_rules_path(rules_path)) if resolve_rules_path(rules_path).exists() else "",
        "master_path": result.master_path,
        "review_path": result.review_path,
        "pdfs_found": len(pdf_paths),
        "pdfs_processed": len(pdf_paths) - len(errors),
        "benefit_rows_found": len(all_rows),
        "benefit_administration_records": [asdict(record) for record in benefit_administration_records],
        "summary_sheets_created": len(summaries),
        "rows_added": result.rows_added,
        "rows_updated": result.rows_updated,
        "pdf_results": pdf_results,
        "errors": [asdict(error) for error in errors],
    }


def emit_progress(enabled: bool, message: str) -> None:
    if enabled:
        print(message, file=sys.stderr, flush=True)


def file_sha256(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def load_extraction_rules(rules_path: str = "") -> ExtractionRules:
    path = resolve_rules_path(rules_path)
    if not path.exists():
        return default_extraction_rules()

    with path.open("r", encoding="utf-8") as rules_file:
        data = json.load(rules_file)

    benefits = []
    for item in data.get("benefits", []):
        name = clean_text(str(item.get("name", "")))
        if not name:
            continue
        headings = tuple(clean_text(str(value)) for value in item.get("headings", []) if clean_text(str(value)))
        keywords = tuple(clean_text(str(value)) for value in item.get("keywords", []) if clean_text(str(value)))
        benefits.append(
            BenefitDefinition(
                name=name,
                keywords=keywords or headings,
                headings=headings,
                page_span=int(item.get("pageSpan", 2) or 2),
            )
        )

    field_patterns = {}
    for field_name, patterns in data.get("fieldPatterns", {}).items():
        field_patterns[str(field_name)] = tuple(str(pattern) for pattern in patterns if str(pattern).strip())

    summary_sections = []
    for section_data in data.get("summarySections", []):
        name = clean_text(str(section_data.get("name", "")))
        if not name:
            continue

        fields = []
        for field_data in section_data.get("fields", []):
            label = clean_text(str(field_data.get("label", "")))
            patterns = tuple(str(pattern) for pattern in field_data.get("patterns", []) if str(pattern).strip())
            max_results = int(field_data.get("maxResults", 1) or 1)
            if label and patterns:
                fields.append(SummaryFieldDefinition(label, patterns, max_results))

        summary_sections.append(
            SummarySectionDefinition(
                name=name,
                keywords=tuple(clean_text(str(value)) for value in section_data.get("keywords", []) if clean_text(str(value))),
                headings=tuple(clean_text(str(value)) for value in section_data.get("headings", []) if clean_text(str(value))),
                fields=tuple(fields),
                source=clean_text(str(section_data.get("source", "section"))) or "section",
                page_span=int(section_data.get("pageSpan", 2) or 2),
                optional=bool(section_data.get("optional", False)),
            )
        )

    client_data = data.get("clientName", {})
    return ExtractionRules(
        benefits=tuple(benefits) or tuple(BENEFITS),
        carrier_patterns=tuple(data.get("carrierPatterns", [])) or CARRIER_PATTERNS,
        client_rejected_names=tuple(client_data.get("rejectedNames", [])) or CLIENT_REJECTED_NAMES,
        front_matter_markers=tuple(data.get("frontMatterMarkers", [])) or FRONT_MATTER_MARKERS,
        field_patterns=field_patterns or FIELD_PATTERNS,
        summary_sections=tuple(summary_sections),
    )


def resolve_rules_path(rules_path: str = "") -> Path:
    if rules_path:
        return Path(rules_path).expanduser()
    return Path(__file__).with_name(DEFAULT_RULES_FILE)


def default_extraction_rules() -> ExtractionRules:
    return ExtractionRules(
        benefits=tuple(BENEFITS),
        carrier_patterns=CARRIER_PATTERNS,
        client_rejected_names=CLIENT_REJECTED_NAMES,
        front_matter_markers=FRONT_MATTER_MARKERS,
        field_patterns=FIELD_PATTERNS,
        summary_sections=(),
    )


def collect_pdf_paths(input_paths: Iterable[str]) -> list[Path]:
    pdf_paths: list[Path] = []
    seen: set[str] = set()

    for raw_path in input_paths:
        path = Path(raw_path).expanduser()
        if path.is_file() and path.suffix.lower() == ".pdf":
            candidates = [path]
        elif path.is_dir():
            candidates = sorted(path.rglob("*.pdf"))
        else:
            raise FileNotFoundError(f"Input path was not found or is not a PDF/folder: {path}")

        for candidate in candidates:
            resolved = str(candidate.resolve()).lower()
            if resolved not in seen:
                seen.add(resolved)
                pdf_paths.append(candidate)

    return sorted(pdf_paths, key=pdf_path_sort_key)


def pdf_path_sort_key(path: Path) -> tuple[int, str]:
    name = path.name.lower()
    return (0 if "booklet" in name else 1, name)


def extract_pdf_text(
    pdf_path: Path,
    max_pages: int | None = None,
    prefer_pypdf: bool = False,
) -> str:
    """Read PDF text with available project/runtime libraries, preferring pdfplumber when installed."""
    pypdf_text = ""
    if prefer_pypdf:
        pypdf_text = try_extract_with_pypdf(pdf_path, max_pages)
        if has_meaningful_pdf_text(pypdf_text):
            return pypdf_text

    pdfplumber_text = try_extract_with_pdfplumber(pdf_path, max_pages)
    if has_meaningful_pdf_text(pdfplumber_text):
        return pdfplumber_text

    if not pypdf_text:
        pypdf_text = try_extract_with_pypdf(pdf_path, max_pages)
    if has_meaningful_pdf_text(pypdf_text):
        return pypdf_text

    # Image-only PDFs can legitimately yield page markers from both readers.
    # Return those markers after both attempts so one supplement does not fail a batch.
    if pypdf_text or pdfplumber_text:
        return pypdf_text or pdfplumber_text

    raise RuntimeError(
        "No PDF text extraction library is available. Install pdfplumber or pypdf for this module."
    )


def has_meaningful_pdf_text(text: str) -> bool:
    """Ignore page markers so an empty extractor can fall back to another PDF reader."""
    without_page_markers = PAGE_MARKER_PATTERN.sub(" ", text or "")
    return bool(re.search(r"[A-Za-z0-9]", without_page_markers))


def try_extract_with_pdfplumber(pdf_path: Path, max_pages: int | None = None) -> str:
    try:
        import pdfplumber  # type: ignore
    except ImportError:
        return ""

    chunks: list[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        pages = pdf.pages[:max_pages] if max_pages is not None else pdf.pages
        for page_number, page in enumerate(pages, start=1):
            chunks.append(f"\n\n[[PAGE {page_number}]]\n{page.extract_text() or ''}")
    return "\n".join(chunks)


def try_extract_with_pypdf(pdf_path: Path, max_pages: int | None = None) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:
        return ""

    reader = PdfReader(str(pdf_path))
    chunks = []
    pages = reader.pages[:max_pages] if max_pages is not None else reader.pages
    for page_number, page in enumerate(pages, start=1):
        chunks.append(f"\n\n[[PAGE {page_number}]]\n{page.extract_text() or ''}")
    return "\n".join(chunks)


def detect_client_info(text: str, pdf_path: Path, rules: ExtractionRules) -> ClientInfo:
    searchable = text[:6000]
    client_name = detect_client_name(searchable, rules) or clean_client_name_from_pdf_stem(pdf_path.stem)
    # Carrier branding can appear after the cover and class pages. Keep client
    # detection tight, but inspect enough front matter to identify the carrier.
    carrier = detect_carrier(text[:20000], rules)
    group_number = first_regex(
        searchable,
        [
            r"\bPlan\s*(?:Number|No\.?|#)\s*[:\-]?\s*([A-Za-z0-9\-]+)",
            r"\bGroup\s*(?:Number|No\.?|#)\s*[:\-]?\s*([A-Za-z0-9\-]+)",
            r"\bAccount\s*(?:Number|No\.?|#)\s*[:\-]?\s*([A-Za-z0-9\-]+)",
        ],
    )
    policy_number = first_regex(
        searchable,
        [
            r"\bPolicy\s*(?:Number\b|No\.?|#)\s*[:\-]?\s*([A-Za-z0-9\-]+)",
            r"\bContract\s*(?:Number|No\.?|#)\s*[:\-]?\s*([A-Za-z0-9\-]+)",
        ],
    )
    if re.search(r"\bAllstate\b", text, flags=re.IGNORECASE) and re.search(
        r"\bPOLICY SPECIFICATIONS\b",
        text,
        flags=re.IGNORECASE,
    ):
        policy_number = first_regex(text, [r"\bPOLICY NUMBER\s*:\s*([A-Za-z0-9\-]+)"]) or policy_number
        carrier = "Allstate"
    client_id_seed = group_number or policy_number
    client_id = create_client_id(client_name, client_id_seed)
    return ClientInfo(client_id, client_name, carrier, group_number, policy_number)


def clean_client_name_from_pdf_stem(stem: str) -> str:
    policy_class_match = match_policy_class_filename(Path(f"{stem}.pdf"))
    if policy_class_match:
        return clean_client_name(policy_class_match.group("client"))

    parts = [part for part in re.split(r"__+", stem) if part]
    candidate = parts[1] if len(parts) > 1 else stem
    candidate = re.sub(r"[_-]+", " ", candidate)
    candidate = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", candidate)
    candidate = re.sub(r"\bCentre\s*for\s*the\b", "Centre for the", candidate, flags=re.IGNORECASE)
    candidate = re.sub(r"\bforthe\b", "for the", candidate, flags=re.IGNORECASE)
    return clean_client_name(candidate)


def combine_file_class_with_detected_class(class_name: str, pdf_path: Path) -> str:
    file_class = detect_file_class_label(pdf_path)
    cleaned_class = clean_class_name(class_name)
    if not file_class:
        return cleaned_class
    if not cleaned_class:
        return file_class
    if cleaned_class.lower().startswith(file_class.lower()):
        return cleaned_class
    return f"{file_class} - {cleaned_class}"


def detect_file_class_label(pdf_path: Path) -> str:
    match = re.search(r"\bClass\s+([A-Za-z0-9]+)\b", pdf_path.stem, flags=re.IGNORECASE)
    if match:
        return f"Class {match.group(1).upper()}"

    policy_class_match = match_policy_class_filename(pdf_path)
    if policy_class_match:
        return clean_class_name(policy_class_match.group("class"))

    parts = [part.strip() for part in re.split(r"__+", pdf_path.stem) if part.strip()]
    if len(parts) >= 5 and any(part.lower() == "booklet" for part in parts[3:]):
        class_code = parts[2]
        if re.fullmatch(r"[A-Za-z0-9-]{1,12}", class_code) and class_code.lower() not in {
            "en",
            "fr",
            "booklet",
        }:
            return f"Class {class_code.upper()}"

    return ""


def detect_booklet_class_names(text: str, pdf_path: Path, client_info: ClientInfo) -> list[str]:
    if is_allstate_critical_illness_policy(text, client_info):
        return ["All Classes"]

    if is_equitable_booklet(text, client_info):
        equitable_classes = extract_equitable_class_names(text)
        if equitable_classes:
            return equitable_classes

    if is_sunlife_benefit_summary_booklet(text, client_info):
        sunlife_class = detect_sunlife_cover_class_name(text)
        if sunlife_class:
            return [sunlife_class]

    class_name = combine_file_class_with_detected_class(detect_class_name(text), pdf_path)
    return [class_name]


def detect_client_name(text: str, rules: ExtractionRules) -> str:
    participant_name = extract_participant_cover_client_name(text, rules)
    if participant_name:
        return participant_name

    lines = [clean_text(line) for line in text.splitlines()]
    lines = [line for line in lines if line and not re.fullmatch(r"\[\[PAGE\s+\d+\]\]", line, flags=re.IGNORECASE)]

    policyholder = first_regex(
        text,
        [
            r"\bPOLICYHOLDER\s*:\s*([^\n\r]{3,180})",
            r"(?im)^\s*Policyholder\s{2,}([^\n\r]{3,180})\s*$",
        ],
    )
    if policyholder:
        cleaned_policyholder = clean_client_name(policyholder)
        if is_likely_client_name(cleaned_policyholder, rules):
            return cleaned_policyholder

    for index, line in enumerate(lines[:80]):
        if re.fullmatch(r"Group\s+Name\s+and\s+Number\(s\)", line, flags=re.IGNORECASE):
            for candidate in lines[index + 1 : index + 6]:
                if re.fullmatch(r"Employee\s+Classification|Group\s+Number\(s\)|Effective\s+Date", candidate, flags=re.IGNORECASE):
                    break
                cleaned = clean_client_name(candidate)
                if is_likely_client_name(cleaned, rules):
                    return cleaned

    for index, line in enumerate(lines[:80]):
        if re.fullmatch(r"Benefits\s+Booklet", line, flags=re.IGNORECASE):
            for candidate in lines[index + 1 : index + 6]:
                cleaned = clean_client_name(candidate)
                if is_likely_client_name(cleaned, rules):
                    return cleaned

    for index, line in enumerate(lines[:80]):
        if re.match(r"^(?:Plan|Group|Policy|Contract)\s*(?:Number|No\.?|#)", line, flags=re.IGNORECASE):
            for candidate in reversed(lines[max(0, index - 5) : index]):
                cleaned = clean_client_name(candidate)
                if is_likely_client_name(cleaned, rules):
                    return cleaned

            # Sun Life cover pages place the legal client name below Contract
            # Number, Effective, and Issued metadata rather than above it.
            for candidate in lines[index + 1 : index + 12]:
                if re.match(
                    r"^(?:Effective|Issued|Benefit Summary|Your Group Benefits|Contract)\b",
                    candidate,
                    flags=re.IGNORECASE,
                ):
                    continue
                if re.search(r"\b(?:Management )?Employees?\b", candidate, flags=re.IGNORECASE):
                    continue
                cleaned = clean_client_name(candidate)
                if is_likely_client_name(cleaned, rules):
                    return cleaned

    patterns = [
        r"(?:Prepared\s+for|Client\s+Name|Plan\s+Sponsor|Policyholder|Company\s+Name)\s*[:\-]\s*([^\n\r]{3,90})",
        r"(?:Benefits\s+Booklet\s+for)\s+([^\n\r]{3,90})",
    ]
    value = first_regex(text, patterns)
    return clean_client_name(value)


def detect_sunlife_cover_class_name(text: str) -> str:
    """Read the employee class printed beside the contract metadata on page one."""
    pages = split_marked_pages(text)
    first_page = next((page_text for page_number, page_text in pages if page_number == 1), pages[0][1] if pages else text)
    lines = [clean_text(line) for line in first_page.splitlines()]
    lines = [line for line in lines if line and not re.fullmatch(r"\[\[PAGE\s+\d+\]\]", line, flags=re.IGNORECASE)]

    for index, line in enumerate(lines[:40]):
        if not re.match(r"^Contract\s*(?:Number|No\.?)\s*[:\-]?", line, flags=re.IGNORECASE):
            continue

        candidates = [
            candidate
            for candidate in lines[index + 1 : index + 14]
            if not re.match(r"^(?:Effective|Issued|Contract|Your Group Benefits)\b", candidate, flags=re.IGNORECASE)
        ]
        if len(candidates) >= 2:
            return clean_class_name(candidates[1])

        preceding = [
            candidate
            for candidate in lines[max(0, index - 5) : index]
            if not re.match(r"^(?:Effective|Issued|Your Group Benefits)\b", candidate, flags=re.IGNORECASE)
        ]
        if len(preceding) >= 2:
            return clean_class_name(preceding[-1])

    return ""


def extract_participant_cover_client_name(text: str, rules: ExtractionRules) -> str:
    """Read the participant name from page one without crossing a PDF page marker."""
    pages = split_marked_pages(text)
    first_page = next((page_text for page_number, page_text in pages if page_number == 1), pages[0][1] if pages else text)
    lines = [clean_text(line) for line in first_page.splitlines()]
    lines = [line for line in lines if line]

    for index, line in enumerate(lines[:80]):
        if not re.fullmatch(r"For\s+Participants\s+of", line, flags=re.IGNORECASE):
            continue

        client_lines: list[str] = []
        for candidate in lines[index + 1 : index + 6]:
            if re.match(
                r"^(?:\[\[PAGE\s+\d+\]\]|Introduction\b|Policyholder\b|Your\s+Benefits\s+Handbook\b|"
                r"Plan\s+Effective\s+Date\b|Effective\s+Date\s+of\s+this\s+Document\b|Employee\s+Class\b)",
                candidate,
                flags=re.IGNORECASE,
            ):
                break
            client_lines.append(candidate)

        cleaned_name = clean_client_name(" ".join(client_lines))
        if is_likely_client_name(cleaned_name, rules):
            return cleaned_name

    return ""


def is_likely_client_name(value: str, rules: ExtractionRules) -> bool:
    if len(value) < 3:
        return False

    normalized = normalize_for_search(value)
    rejected = {normalize_for_search(name) for name in rules.client_rejected_names}
    if normalized in rejected:
        return False

    if re.match(r"^(?:plan|group|policy|contract|class)\b", value, flags=re.IGNORECASE):
        return False

    return any(character.isalpha() for character in value)


def detect_carrier(text: str, rules: ExtractionRules) -> str:
    if is_benepac_booklet(text):
        return "Benepac"

    normalized = normalize_for_search(text)
    if re.search(r"\b(?:national\s*health\s*claim|nhc)\b|lifehealth\+", text, flags=re.IGNORECASE):
        return "National HealthClaim"
    for carrier in rules.carrier_patterns:
        if normalize_for_search(carrier) in normalized:
            return carrier
    return ""


def first_regex(text: str, patterns: Iterable[str]) -> str:
    for pattern in patterns:
        pattern = normalize_regex_pattern(pattern)
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return clean_text(match.group(1)).strip(" :-")
    return ""


def normalize_regex_pattern(pattern: str) -> str:
    """Correct accidentally doubled regex escapes in extractor rule snippets."""
    if "\\\\" not in pattern:
        return pattern

    for token in ("b", "B", "s", "S", "d", "D", "w", "W", "A", "Z"):
        pattern = pattern.replace("\\\\" + token, "\\" + token)
    for token in ("(", ")", "[", "]", "{", "}", "-", ".", "+", "*", "?", "|", "^", "$"):
        pattern = pattern.replace("\\\\" + token, "\\" + token)
    pattern = pattern.replace("\\\\u", "\\u")
    return pattern


def create_client_id(client_name: str, policy_seed: str = "") -> str:
    client_part = re.sub(r"[^A-Za-z0-9]+", "-", client_name.upper()).strip("-")
    seed_part = re.sub(r"[^A-Za-z0-9]+", "-", policy_seed.upper()).strip("-")
    if seed_part:
        return f"{client_part}-{seed_part}"
    return client_part


def extract_plan_design_rows(
    text: str,
    pdf_path: Path,
    client_info: ClientInfo,
    rules: ExtractionRules,
    class_name: str = "",
    document_text: str = "",
) -> list[ExtractionRow]:
    """Find each requested benefit and convert nearby booklet text into PlanDesign rows.

    ``text`` may exclude front matter to keep broad keyword matching focused. The optional
    full document remains available to the authoritative GroupHEALTH schedule parsers.
    """
    if is_allstate_critical_illness_policy(text, client_info):
        allstate_rows = extract_allstate_critical_illness_plan_design_rows(text, pdf_path, client_info, class_name)
        if allstate_rows:
            return merge_labeled_dental_rate_rows(allstate_rows, text, pdf_path, client_info, class_name)

    if is_benepac_booklet(text):
        benepac_rows = extract_benepac_schedule_plan_design_rows(text, pdf_path, client_info, class_name)
        if benepac_rows:
            return merge_labeled_dental_rate_rows(benepac_rows, text, pdf_path, client_info, class_name)
    if is_equitable_booklet(text, client_info):
        equitable_rows = extract_equitable_schedule_plan_design_rows(text, pdf_path, client_info, class_name)
        if equitable_rows:
            return merge_labeled_dental_rate_rows(equitable_rows, text, pdf_path, client_info, class_name)

    if is_desjardins_booklet(text, client_info):
        desjardins_rows = extract_desjardins_plan_design_rows(text, pdf_path, client_info, class_name)
        if desjardins_rows:
            return merge_labeled_dental_rate_rows(desjardins_rows, text, pdf_path, client_info, class_name)

    if is_sunlife_benefit_summary_booklet(document_text or text, client_info):
        sunlife_rows = extract_sunlife_plan_design_rows(document_text or text, pdf_path, client_info, class_name)
        if sunlife_rows:
            return merge_labeled_dental_rate_rows(sunlife_rows, text, pdf_path, client_info, class_name)

    rows: list[ExtractionRow] = []
    normalized_lines = build_search_lines(text)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    schedule_source_text = document_text or text
    is_grouphealth = is_grouphealth_booklet(schedule_source_text, client_info)
    schedule_life = extract_schedule_life_details(schedule_source_text)
    schedule_dependent_life = extract_schedule_dependent_life_details(schedule_source_text)
    schedule_asi = extract_schedule_asi_details(schedule_source_text) if is_grouphealth else None
    schedule_critical_illness = (
        extract_schedule_critical_illness_details(schedule_source_text)
        if is_grouphealth
        else None
    )
    schedule_std = extract_schedule_std_details(schedule_source_text) if is_grouphealth else None
    schedule_ltd = extract_schedule_ltd_details(schedule_source_text) if is_grouphealth else None
    schedule_ehc = extract_schedule_ehc_details(schedule_source_text) if is_grouphealth else None
    schedule_dental_page, schedule_dental_text = (
        find_grouphealth_dental_schedule_page(schedule_source_text)
        if is_grouphealth
        else (0, "")
    )
    schedule_dental_recall = extract_recall_frequency(schedule_dental_text)

    for definition in rules.benefits:
        is_prescription_drugs = definition.name in {"Drugs", "Prescription Drugs"}
        is_medical_services = definition.name == "Medical Services & Supplies"
        is_paramedical = definition.name == "Paramedical"
        is_vision = definition.name == "Vision"
        is_travel = definition.name == "Travel"
        is_dental = definition.name.startswith("Dental ")
        if is_grouphealth and definition.name == "Life" and not schedule_life:
            continue
        if is_grouphealth and is_prescription_drugs and (
            not schedule_ehc or not schedule_ehc.prescription_drugs_listed
        ):
            continue
        if is_grouphealth and is_medical_services and (
            not schedule_ehc or not schedule_ehc.other_medical_listed
        ):
            continue
        if is_grouphealth and is_dental and (
            not schedule_dental_text
            or definition.name != "Dental Recall"
            or not schedule_dental_recall
        ):
            continue

        section_text, matched_keywords = find_benefit_section(text, normalized_lines, definition)
        if is_grouphealth and definition.name == "Life" and schedule_life:
            section_text = "\n".join(build_schedule_life_summary_lines(schedule_source_text))
            matched_keywords = ["Schedule of Benefits - Group Life Insurance"]
        if is_grouphealth and definition.name == "AD&D" and schedule_asi:
            section_text = "Accident & Serious Illness (ASI)"
            matched_keywords = ["Schedule of Benefits - Accident & Serious Illness (ASI)"]
        if is_grouphealth and definition.name == "Critical Illness" and schedule_critical_illness:
            section_text = "\n".join(
                build_schedule_critical_illness_summary_lines(schedule_source_text)
            )
            matched_keywords = ["Schedule of Benefits - Group Critical Illness Insurance"]
        if is_grouphealth and definition.name == "Dependent Life":
            if schedule_dependent_life:
                section_text = "Dependent Life Insurance"
                matched_keywords = ["Dependent Life Insurance"]
            else:
                section_text = "Dependent Life Insurance: Not covered"
                matched_keywords = ["Schedule of Benefits"]
        if is_grouphealth and definition.name == "STD":
            if schedule_std:
                section_text = "Short-Term Disability (STD)"
                matched_keywords = ["Short-Term Disability (STD)"]
            else:
                section_text = "Short-Term Disability (STD): Not covered"
                matched_keywords = ["Schedule of Benefits"]
        if is_grouphealth and definition.name == "LTD":
            section_text = (
                "Long-Term Disability (LTD)"
                if schedule_ltd
                else "Long-Term Disability (LTD): Not covered"
            )
            matched_keywords = ["Schedule of Benefits - Long-Term Disability (LTD)"]
        if is_grouphealth and is_prescription_drugs and schedule_ehc:
            section_text = (
                "\n".join(build_schedule_prescription_drug_summary_lines(schedule_source_text))
                or "Prescription Drugs: Included"
            )
            matched_keywords = ["Schedule of Benefits - Extended Health Care - Prescription Drugs"]
        if is_grouphealth and is_medical_services and schedule_ehc:
            section_text = (
                "\n".join(build_schedule_ehc_summary_lines(schedule_source_text))
                or "Other Medical Expenses: Included"
            )
            matched_keywords = ["Schedule of Benefits - Extended Health Care - Other Medical Expenses"]
        if (
            is_grouphealth
            and is_paramedical
            and schedule_ehc
            and schedule_ehc.health_care_practitioners_listed
        ):
            section_text = (
                "\n".join(build_schedule_paramedical_summary_lines(schedule_source_text))
                or "Health Care Practitioners: Included"
            )
            matched_keywords = [
                "Schedule of Benefits - Extended Health Care - Health Care Practitioners"
            ]
        if is_grouphealth and is_vision and schedule_ehc and schedule_ehc.vision_listed:
            section_text = (
                "\n".join(build_schedule_vision_summary_lines(schedule_source_text))
                or "Vision Care: Included"
            )
            matched_keywords = ["Schedule of Benefits - Extended Health Care - Vision Care"]
        if is_grouphealth and is_travel and schedule_ehc and schedule_ehc.travel_listed:
            section_text = (
                "\n".join(build_schedule_travel_summary_lines(schedule_source_text))
                or "Travel Insurance and Assistance: Included"
            )
            matched_keywords = [
                "Schedule of Benefits - Extended Health Care - Travel Insurance and Assistance"
            ]
        if is_grouphealth and definition.name == "Dental Recall":
            section_text = f"Recall: {schedule_dental_recall}"
            matched_keywords = ["Schedule of Benefits - Dental Care - Recall"]
        if is_grouphealth and definition.name in {"EAP", "Virtual Care"}:
            section_text = extract_grouphealth_sparrow_section(text) or section_text or "Sparrow"
            matched_keywords = unique_in_order([*matched_keywords, "Sparrow"])
        if not section_text:
            continue

        fields = extract_fields(section_text, rules, definition.name)
        coverage_description = build_plan_coverage_description(definition.name, section_text, is_grouphealth)
        source_section = ", ".join(matched_keywords[:4])
        source_page: int | str = ""
        if definition.name == "Life":
            fields["Coinsurance"] = ""
            benefit_amount = extract_lump_sum_benefit_amount(section_text)
            if schedule_life:
                benefit_amount = schedule_life.benefit_amount or benefit_amount
                fields["Maximum"] = schedule_life.maximum_benefit or fields["Maximum"]
                fields["TerminationAge"] = schedule_life.termination_age or fields["TerminationAge"]
                source_section = "Schedule of Benefits - Group Life Insurance"
                source_page = schedule_life.source_page
            if benefit_amount:
                coverage_description = f"Benefit Amount: {benefit_amount}"

        if is_grouphealth and definition.name == "AD&D" and schedule_asi:
            coverage_description = "\n".join(
                compact_detail_lines(
                    [
                        f"Benefit Amount: {schedule_asi.benefit_amount}",
                        f"Critical Disease Benefit: {schedule_asi.critical_disease_benefit}",
                        f"Serious Illness Benefit: {schedule_asi.serious_illness_benefit}",
                    ]
                )
            )
            fields["Coinsurance"] = ""
            fields["Maximum"] = ""
            fields["Frequency"] = ""
            fields["WaitingPeriod"] = ""
            fields["TerminationAge"] = schedule_asi.termination_age
            source_section = "Schedule of Benefits - Accident & Serious Illness (ASI)"
            source_page = schedule_asi.source_page

        if is_grouphealth and definition.name == "Critical Illness" and schedule_critical_illness:
            coverage_description = "\n".join(
                build_schedule_critical_illness_summary_lines(schedule_source_text)
            )
            fields["Coinsurance"] = ""
            fields["Maximum"] = schedule_critical_illness.benefit_amount
            fields["Frequency"] = ""
            fields["WaitingPeriod"] = ""
            fields["TerminationAge"] = schedule_critical_illness.termination_age
            source_section = "Schedule of Benefits - Group Critical Illness Insurance"
            source_page = schedule_critical_illness.source_page

        if is_grouphealth and definition.name == "Dependent Life":
            fields["Coinsurance"] = ""
            fields["Maximum"] = ""
            fields["Frequency"] = ""
            fields["WaitingPeriod"] = ""
            if schedule_dependent_life:
                coverage_description = "\n".join(
                    compact_detail_lines(
                        [
                            f"Spouse: {schedule_dependent_life.spouse_amount}",
                            f"Dependent: {schedule_dependent_life.dependent_child_amount}",
                            f"Dependent Child Eligibility: {schedule_dependent_life.dependent_child_eligibility}",
                            f"Stillbirth Benefit: {schedule_dependent_life.stillbirth_benefit}",
                        ]
                    )
                )
                fields["TerminationAge"] = schedule_dependent_life.termination_age
                source_section = "Schedule of Benefits - Dependent Life Insurance"
                source_page = schedule_dependent_life.source_page
            else:
                coverage_description = "Coverage: Not covered"
                fields["TerminationAge"] = ""
                source_section = "Schedule of Benefits - Dependent Life Insurance not listed"
                source_page = schedule_life.source_page if schedule_life else ""

        if is_grouphealth and definition.name == "STD":
            fields["Coinsurance"] = ""
            fields["Frequency"] = ""
            if schedule_std:
                coverage_description = "\n".join(format_schedule_std_summary_lines(schedule_std))
                fields["Maximum"] = schedule_std.maximum_benefit
                fields["WaitingPeriod"] = schedule_std.qualification_period
                fields["TerminationAge"] = schedule_std.termination_age
                source_section = "Schedule of Benefits - Short-Term Disability (STD)"
                source_page = schedule_std.source_page
            else:
                coverage_description = "Coverage: Not covered"
                fields["Maximum"] = ""
                fields["WaitingPeriod"] = ""
                fields["TerminationAge"] = ""
                source_section = "Schedule of Benefits - Short-Term Disability (STD) not listed"
                source_page = schedule_life.source_page if schedule_life else ""

        if is_grouphealth and definition.name == "LTD":
            fields["Coinsurance"] = ""
            fields["Frequency"] = ""
            if schedule_ltd:
                coverage_description = "\n".join(format_schedule_ltd_summary_lines(schedule_ltd))
                fields["Maximum"] = schedule_ltd.maximum_benefit
                fields["WaitingPeriod"] = schedule_ltd.qualification_period
                fields["TerminationAge"] = schedule_ltd.termination_age
                source_section = "Schedule of Benefits - Long-Term Disability (LTD)"
                source_page = schedule_ltd.source_page
            else:
                coverage_description = "Coverage: Not covered"
                fields["Maximum"] = ""
                fields["WaitingPeriod"] = ""
                fields["TerminationAge"] = ""
                source_section = "Schedule of Benefits - Long-Term Disability (LTD) not listed"
                source_page = schedule_life.source_page if schedule_life else ""

        if is_grouphealth and definition.name in {"EAP", "Virtual Care"}:
            fields["Coinsurance"] = ""
            fields["Maximum"] = ""
            fields["Frequency"] = ""
            fields["WaitingPeriod"] = ""
            fields["TerminationAge"] = ""
            if re.search(r"\bSparrow\b", section_text, flags=re.IGNORECASE):
                coverage_description = "Coverage: Sparrow"

        if is_grouphealth and definition.name == "Medical Second Opinion":
            fields["Coinsurance"] = ""
            fields["Maximum"] = ""
            fields["Frequency"] = ""
            fields["WaitingPeriod"] = ""
            fields["TerminationAge"] = ""
            coverage_description = "Coverage: Included"
            source_section = "Medical Second Opinion"
            source_page = next(
                (
                    page_number
                    for page_number, page_text in split_marked_pages(schedule_source_text)
                    if re.search(r"\bMedical\s+Second\s+Opinion\b", page_text, flags=re.IGNORECASE)
                ),
                "",
            )

        if is_grouphealth and is_prescription_drugs and schedule_ehc:
            coverage_description = "\n".join(
                build_schedule_prescription_drug_summary_lines(schedule_source_text)
            )
            fields["Coinsurance"] = schedule_ehc.prescription_drug_coinsurance
            fields["Maximum"] = schedule_ehc.prescription_drug_maximum
            fields["Frequency"] = ""
            fields["WaitingPeriod"] = ""
            fields["TerminationAge"] = ""
            source_section = "Schedule of Benefits - Extended Health Care - Prescription Drugs"
            source_page = schedule_ehc.source_page

        if is_grouphealth and is_medical_services and schedule_ehc:
            coverage_description = "\n".join(build_schedule_ehc_summary_lines(schedule_source_text))
            fields["Coinsurance"] = schedule_ehc.other_medical_coinsurance
            fields["Maximum"] = ""
            fields["Frequency"] = ""
            fields["WaitingPeriod"] = ""
            fields["TerminationAge"] = ""
            source_section = "Schedule of Benefits - Extended Health Care - Other Medical Expenses"
            source_page = schedule_ehc.source_page

        if (
            is_grouphealth
            and is_paramedical
            and schedule_ehc
            and schedule_ehc.health_care_practitioners_listed
        ):
            coverage_description = "\n".join(
                build_schedule_paramedical_summary_lines(schedule_source_text)
            )
            fields["Coinsurance"] = schedule_ehc.health_care_practitioners_coinsurance
            fields["Maximum"] = schedule_ehc.health_care_practitioners_maximum
            fields["WaitingPeriod"] = ""
            fields["TerminationAge"] = ""
            source_section = (
                "Schedule of Benefits - Extended Health Care - Health Care Practitioners"
            )
            source_page = schedule_ehc.source_page

        if is_grouphealth and is_vision and schedule_ehc and schedule_ehc.vision_listed:
            coverage_description = "\n".join(
                build_schedule_vision_summary_lines(schedule_source_text)
            )
            fields["Coinsurance"] = schedule_ehc.vision_coinsurance
            fields["Maximum"] = schedule_ehc.vision_maximum or schedule_ehc.eye_examination_benefit
            fields["Frequency"] = ""
            fields["WaitingPeriod"] = ""
            fields["TerminationAge"] = ""
            source_section = "Schedule of Benefits - Extended Health Care - Vision Care"
            source_page = schedule_ehc.source_page

        if is_grouphealth and is_travel and schedule_ehc and schedule_ehc.travel_listed:
            coverage_description = "\n".join(
                build_schedule_travel_summary_lines(schedule_source_text)
            )
            fields["Coinsurance"] = schedule_ehc.travel_coinsurance
            fields["Maximum"] = schedule_ehc.travel_maximum
            fields["Frequency"] = schedule_ehc.travel_duration
            fields["WaitingPeriod"] = ""
            fields["TerminationAge"] = schedule_ehc.travel_termination_age
            source_section = (
                "Schedule of Benefits - Extended Health Care - Travel Insurance and Assistance"
            )
            source_page = schedule_ehc.travel_details_source_page or schedule_ehc.source_page

        if is_grouphealth and definition.name == "Dental Recall":
            coverage_description = f"Frequency: {schedule_dental_recall}"
            fields["Coinsurance"] = ""
            fields["Maximum"] = ""
            fields["Frequency"] = schedule_dental_recall
            fields["WaitingPeriod"] = ""
            fields["TerminationAge"] = ""
            source_section = "Schedule of Benefits - Dental Care - Recall"
            source_page = schedule_dental_page

        if should_skip_plan_design_benefit(definition.name, section_text, fields, text, client_info):
            continue
        confidence = (
            "High"
            if is_grouphealth and (
                definition.name in {"AD&D", "Critical Illness", "Dependent Life", "STD", "LTD"}
                or definition.name == "Medical Second Opinion"
                or is_prescription_drugs
                or is_medical_services
                or (
                    is_paramedical
                    and schedule_ehc is not None
                    and schedule_ehc.health_care_practitioners_listed
                )
                or (is_vision and schedule_ehc is not None and schedule_ehc.vision_listed)
                or (is_travel and schedule_ehc is not None and schedule_ehc.travel_listed)
                or (definition.name == "Dental Recall" and bool(schedule_dental_text))
            )
            else score_confidence(section_text, matched_keywords, fields)
        )
        notes = confidence_note(confidence)
        rows.append(
            ExtractionRow(
                client_id=client_info.client_id,
                client_name=client_info.client_name,
                carrier=client_info.carrier,
                class_name=class_name,
                benefit=definition.name,
                coverage_description=coverage_description,
                coinsurance=fields["Coinsurance"],
                maximum=fields["Maximum"],
                frequency=fields["Frequency"],
                waiting_period=fields["WaitingPeriod"],
                termination_age=fields["TerminationAge"],
                notes=notes,
                source_file=pdf_path.name,
                source_section=source_section,
                confidence=confidence,
                review_flag="Review" if confidence in {"Medium", "Low"} else "",
                last_updated=now,
                source_page=source_page,
            )
        )

        if (
            is_grouphealth
            and is_travel
            and schedule_ehc
            and schedule_ehc.travel_cancellation_listed
        ):
            rows.append(
                ExtractionRow(
                    client_id=client_info.client_id,
                    client_name=client_info.client_name,
                    carrier=client_info.carrier,
                    class_name=class_name,
                    benefit="Travel Cancellation Insurance",
                    benefit_category="Travel",
                    coverage_description="Coverage: Included",
                    coinsurance=schedule_ehc.travel_cancellation_coinsurance,
                    maximum=schedule_ehc.travel_cancellation_maximum,
                    frequency="",
                    waiting_period="",
                    termination_age=schedule_ehc.travel_termination_age,
                    notes="",
                    source_file=pdf_path.name,
                    source_section=(
                        "Schedule of Benefits - Extended Health Care - "
                        "Travel Cancellation Insurance"
                    ),
                    confidence="High",
                    review_flag="",
                    last_updated=now,
                    source_page=(
                        schedule_ehc.travel_details_source_page or schedule_ehc.source_page
                    ),
                )
            )

    if is_grouphealth and schedule_dental_text:
        dental_deductible = extract_grouphealth_dental_deductible(schedule_dental_text)
        dental_fee_guide = extract_grouphealth_dental_fee_guide(schedule_dental_text)
        for benefit, label, value, source_section in (
            (
                "Dental Deductible",
                "Deductible",
                dental_deductible,
                "Schedule of Benefits - Dental Care - Deductible",
            ),
            (
                "Dental Fee Guide",
                "Fee Guide",
                dental_fee_guide,
                "Schedule of Benefits - Dental Care - Fee Guide",
            ),
        ):
            if not value:
                continue
            rows.append(
                ExtractionRow(
                    client_id=client_info.client_id,
                    client_name=client_info.client_name,
                    carrier=client_info.carrier,
                    class_name=class_name,
                    benefit=benefit,
                    benefit_category="Dental Care",
                    coverage_description=f"{label}: {value}",
                    coinsurance="",
                    maximum="",
                    frequency="",
                    waiting_period="",
                    termination_age="",
                    notes="",
                    source_file=pdf_path.name,
                    source_section=source_section,
                    confidence="High",
                    review_flag="",
                    last_updated=now,
                    source_page=schedule_dental_page,
                )
            )

    dental_source_text = schedule_source_text if is_grouphealth else text
    merged_rows = merge_labeled_dental_rate_rows(
        rows,
        dental_source_text,
        pdf_path,
        client_info,
        class_name,
    )
    if is_grouphealth:
        administration_records = extract_grouphealth_benefit_administration(
            schedule_source_text,
            pdf_path,
            client_info,
        )
        apply_grouphealth_administration_to_plan_rows(merged_rows, administration_records)
    return merged_rows


def _source_page_supported_dental_maximum(
    maximum: str,
    source_page: int | str,
    full_text: str,
) -> str:
    """Keep a fallback maximum only when its value appears on the mapped dental page."""

    maximum = clean_text(maximum)
    if not maximum or source_page in (None, ""):
        return ""

    try:
        page_number = int(str(source_page).strip())
    except (TypeError, ValueError):
        return ""

    page_text = next(
        (page_text for page, page_text in split_marked_pages(full_text) if page == page_number),
        "",
    )
    if not page_text:
        return ""

    if re.search(r"\bno\s+maximum\b", maximum, flags=re.IGNORECASE):
        return maximum if re.search(r"\bno\s+maximum\b", page_text, flags=re.IGNORECASE) else ""

    if re.search(r"\bunlimited\b", maximum, flags=re.IGNORECASE):
        return maximum if re.search(r"\bunlimited\b", page_text, flags=re.IGNORECASE) else ""

    amounts = re.findall(r"\$\s?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?", maximum)
    if amounts:
        page_amounts = {
            re.sub(r"[^0-9.]", "", amount)
            for amount in re.findall(
                r"\$\s?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?",
                page_text,
            )
        }
        expected_amounts = {re.sub(r"[^0-9.]", "", amount) for amount in amounts}
        return maximum if expected_amounts.issubset(page_amounts) else ""

    normalized_maximum = normalize_for_search(maximum)
    return maximum if normalized_maximum and normalized_maximum in normalize_for_search(page_text) else ""


def merge_labeled_dental_rate_rows(
    rows: list[ExtractionRow],
    full_text: str,
    pdf_path: Path,
    client_info: ClientInfo,
    class_name: str = "",
) -> list[ExtractionRow]:
    """Replace ambiguous dental percentages only when an explicit label-rate mapping exists."""

    excluded_output_codes = {"routine"}
    mapped_statuses = [
        status
        for status in map_labeled_dental_coverage_statuses(full_text)
        if status.provision_code not in excluded_output_codes
    ]
    status_by_code = {status.provision_code: status for status in mapped_statuses}
    mapped_rates = [
        rate
        for rate in map_labeled_dental_rates(full_text)
        if rate.provision_code not in status_by_code
        and rate.provision_code not in excluded_output_codes
    ]
    mapped_maximums = [
        maximum
        for maximum in map_labeled_dental_maximums(full_text)
        if maximum.provision_code not in status_by_code
        and maximum.provision_code not in excluded_output_codes
    ]
    if not mapped_rates and not mapped_maximums and not mapped_statuses:
        return rows

    mapped_by_code = {rate.provision_code: rate for rate in mapped_rates}
    maximum_by_code = {maximum.provision_code: maximum for maximum in mapped_maximums}
    mapped_codes = set(mapped_by_code) | set(maximum_by_code) | set(status_by_code)
    mapped_values = {rate.reimbursement_percent for rate in mapped_rates}
    existing_by_code: dict[str, list[ExtractionRow]] = defaultdict(list)
    retained: list[ExtractionRow] = []

    for row in rows:
        provision_code = canonical_dental_provision_code(row.benefit)
        if provision_code in excluded_output_codes:
            continue
        if provision_code in mapped_codes:
            existing_by_code[provision_code].append(row)
            continue

        row_percentages = _valid_percentage_values(row.coinsurance)
        is_redundant_dental_summary = (
            not provision_code
            and "dental" in row.benefit.lower()
            and len(row_percentages) > 1
            and set(row_percentages).issubset(mapped_values)
        )
        if not is_redundant_dental_summary:
            retained.append(row)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for rate in mapped_rates:
        compatible_existing = next(
            (
                row
                for row in existing_by_code.get(rate.provision_code, [])
                if _valid_percentage_values(row.coinsurance) == [rate.reimbursement_percent]
            ),
            None,
        )
        eligibility_note = f" Eligibility: {rate.eligibility}." if rate.eligibility else ""
        existing_notes = f" {compatible_existing.notes}" if compatible_existing and compatible_existing.notes else ""
        labeled_maximum = maximum_by_code.get(rate.provision_code)
        existing_maximum = (
            _source_page_supported_dental_maximum(
                compatible_existing.maximum,
                rate.source_page,
                full_text,
            )
            if compatible_existing
            else ""
        )
        resolved_maximum = (
            rate.maximum
            or (labeled_maximum.maximum if labeled_maximum else "")
            or existing_maximum
        )
        frequency = canonical_labeled_dental_frequency(
            rate.provision_code,
            resolved_maximum,
            " ".join(
                value
                for value in [rate.source_text, labeled_maximum.source_text if labeled_maximum else ""]
                if value
            ),
        )
        retained.append(
            ExtractionRow(
                client_id=client_info.client_id,
                client_name=client_info.client_name,
                carrier=client_info.carrier,
                class_name=class_name,
                benefit=rate.benefit_name,
                coverage_description=rate.source_text,
                coinsurance=rate.reimbursement_text,
                maximum=resolved_maximum,
                frequency=frequency,
                waiting_period=compatible_existing.waiting_period if compatible_existing else "",
                termination_age=compatible_existing.termination_age if compatible_existing else "",
                notes=(
                    f"{CANONICAL_LABELED_DENTAL_RATE_MAPPER_VERSION}: explicit label and percentage."
                    f"{eligibility_note}{existing_notes}"
                ).strip(),
                source_file=pdf_path.name,
                source_section=f"Dental Care - {rate.provision_name}",
                confidence=rate.confidence,
                review_flag="",
                last_updated=now,
                policy_number=client_info.policy_number,
                source_page=rate.source_page,
            )
        )

    for maximum in mapped_maximums:
        if maximum.provision_code in mapped_by_code:
            continue
        compatible_existing = next(iter(existing_by_code.get(maximum.provision_code, [])), None)
        frequency = canonical_labeled_dental_frequency(
            maximum.provision_code,
            maximum.maximum,
            maximum.source_text,
        )
        retained.append(
            ExtractionRow(
                client_id=client_info.client_id,
                client_name=client_info.client_name,
                carrier=client_info.carrier,
                class_name=class_name,
                benefit=maximum.benefit_name,
                coverage_description=maximum.source_text,
                coinsurance="",
                maximum=maximum.maximum,
                frequency=frequency,
                waiting_period=compatible_existing.waiting_period if compatible_existing else "",
                termination_age=compatible_existing.termination_age if compatible_existing else "",
                notes=f"{CANONICAL_LABELED_DENTAL_RATE_MAPPER_VERSION}: explicit label and maximum.",
                source_file=pdf_path.name,
                source_section=f"Dental Care - {maximum.provision_name}",
                confidence=maximum.confidence,
                review_flag="",
                last_updated=now,
                policy_number=client_info.policy_number,
                source_page=maximum.source_page,
            )
        )

    for status in mapped_statuses:
        retained.append(
            ExtractionRow(
                client_id=client_info.client_id,
                client_name=client_info.client_name,
                carrier=client_info.carrier,
                class_name=class_name,
                benefit=status.benefit_name,
                coverage_description=f"Coverage: {status.coverage_status}",
                coinsurance="",
                maximum="",
                frequency="",
                waiting_period="",
                termination_age="",
                notes=(
                    f"{CANONICAL_LABELED_DENTAL_RATE_MAPPER_VERSION}: "
                    "explicit label and coverage status."
                ),
                source_file=pdf_path.name,
                source_section=f"Dental Care - {status.provision_name}",
                confidence=status.confidence,
                review_flag="",
                last_updated=now,
                policy_number=client_info.policy_number,
                source_page=status.source_page,
            )
        )

    return retained


def canonical_labeled_dental_frequency(
    provision_code: str,
    maximum: str,
    source_text: str,
) -> str:
    if provision_code == "orthodontic" and re.search(
        r"\blifetime\b",
        f"{maximum} {source_text}",
        flags=re.IGNORECASE,
    ):
        return "per lifetime"
    return ""


def _valid_percentage_values(value: str) -> list[Decimal]:
    percentages: list[Decimal] = []
    for match in re.finditer(r"(?<![\d.])(100(?:\.0+)?|\d{1,2}(?:\.\d+)?)\s*%(?![a-z0-9])", value or "", flags=re.IGNORECASE):
        try:
            percentage = Decimal(match.group(1))
        except InvalidOperation:
            continue
        if 0 <= percentage <= 100 and percentage not in percentages:
            percentages.append(percentage)
    return percentages


def should_skip_plan_design_benefit(
    benefit_name: str,
    section_text: str,
    fields: dict[str, str],
    full_text: str = "",
    client_info: ClientInfo | None = None,
) -> bool:
    if benefit_name == "Digital Wellness" and is_grouphealth_booklet(full_text or section_text, client_info):
        return True

    if benefit_name != "Critical Illness":
        return False

    has_schedule_amount = re.search(
        rf"\b(?:Principal Sum|Benefit Amount|Benefit amount)\s*-?\s*{MONEY_AMOUNT_PATTERN}",
        section_text,
        flags=re.IGNORECASE,
    )
    has_clean_amount = bool(re.fullmatch(MONEY_AMOUNT_PATTERN, clean_text(fields.get("Maximum", ""))))

    if re.search(
        r"\bOptional\s+(?:Personal\s+)?Critical\s+Illness\b|\bOptional\s+.*?\bCI\b|\bOptional\s+Member\s+Advantage\b",
        section_text,
        flags=re.IGNORECASE,
    ) and not has_schedule_amount:
        return True

    return not (has_schedule_amount or has_clean_amount)


def build_plan_coverage_description(benefit_name: str, section_text: str, is_grouphealth: bool) -> str:
    if is_grouphealth and benefit_name in {"EAP", "Virtual Care"}:
        return "Sparrow"
    return build_coverage_description(section_text)


def is_grouphealth_booklet(text: str, client_info: ClientInfo | None = None) -> bool:
    if (client_info and client_info.carrier.lower() == "benepac") or is_benepac_booklet(text):
        return False

    carrier_text = client_info.carrier if client_info else ""
    searchable = f"{carrier_text} {text or ''}"
    return bool(re.search(r"\b(?:Group\s*HEALTH|GroupHEALTH|Sparrow)\b", searchable, flags=re.IGNORECASE))


def extract_grouphealth_benefit_administration(
    text: str,
    pdf_path: Path,
    client_info: ClientInfo,
) -> list[BenefitAdministrationExtraction]:
    """Extract the GroupHEALTH insurer/provider table without mixing it into plan provisions."""
    if not is_grouphealth_booklet(text, client_info):
        return []

    benefit_patterns = [
        (
            "Protector Series Optional Accident & Serious Illness (ASI)",
            r"(?:Industrial\s+Alliance\s+)?Protector\s+Series(?:\s*(?:TM|™|�))?\s+(?:\d{4,10}\s+)?Optional\s+Accident\s*(?:&|and)\s*Serious\s+Illness\s*\(ASI\)",
            "Protector Series",
            True,
        ),
        (
            "Protector Series Optional Critical Illness",
            r"(?:Temple\s+Insurance\s+Company\s+)?Protector\s+Series(?:\s*(?:TM|™|�))?\s+(?:\d{4,10}\s+)?Optional\s+Critical\s+Illness",
            "Protector Series",
            True,
        ),
        (
            "Protector Series Optional Life Insurance",
            r"Protector\s+Series(?:\s*(?:TM|™|�))?\s+(?:The\s+Co-operators\s+\d{4,10}\s+)?Optional\s+Life\s+Insurance",
            "Protector Series",
            True,
        ),
        (
            "Travel Insurance and Assistance / Travel Cancellation Insurance",
            r"Travel\s+Insurance\s+and\s+Assistance",
            "",
            False,
        ),
        (
            "Employee & Family Assistance Program (EFAP)",
            r"Employee\s*(?:&|and)\s*Family\s+Assistance\s+Program\s*\(EFAP\)",
            "",
            False,
        ),
        (
            "Accidental Death and Dismemberment (AD&D)",
            r"Accidental\s+Death\s+and\s+Dismemberment\s*\(AD\s*&\s*D\)",
            "",
            False,
        ),
        ("Dependent Life Insurance", r"Dependent\s+Life\s+Insurance", "", False),
        (
            "Accident & Serious Illness (ASI)",
            r"Accident\s*(?:&|and)\s*Serious\s+Illness(?:\s*\(ASI\))?",
            "",
            False,
        ),
        ("Group Critical Illness", r"Group\s+Critical\s+Illness", "", False),
        ("Group Life Insurance", r"Group\s+Life\s+Insurance", "", False),
        ("Long-Term Disability", r"Long[\s-]*Term\s+Disability\s*\*?", "", False),
        ("Extended Health Care", r"Extended\s+Health\s+Care", "", False),
        ("Medical Second Opinion", r"Medical\s+Second\s+Opinion", "", False),
        ("Virtual Healthcare (VHC)", r"Virtual\s+Health\s*care\s*\(VHC\)", "", False),
        ("Dental Care", r"Dental\s+Care", "", False),
    ]
    insurer_patterns = [
        ("SSQ, Life Insurance Company Inc.", r"SSQ\s*,?\s*Life\s+Insurance\s+Company\s+Inc\.?"),
        ("Industrial Alliance", r"(?:Industrial\s+Alliance|iA\s+Financial\s+Group)"),
        ("The Co-operators", r"The\s+Co-operators"),
        ("Temple Insurance Company", r"Temple\s+Insurance\s+Company"),
    ]
    provider_patterns = [
        ("Disability Management Institute", r"Disability\s+Management\s+Institute(?:\s*\(DMI\))?"),
        ("GroupHEALTH", r"Group\s*HEALTH"),
        ("Trident Global Assistance", r"Trident\s+Global\s+Assistance"),
        ("TELUS Health Virtual Care", r"TELUS\s+Health\s+Virtual\s+Care"),
        ("TELUS Health", r"TELUS\s+Health(?!\s+Virtual\s+Care)"),
        *BENEFIT_PROVIDER_PATTERNS["Medical Second Opinion"],
    ]

    records: list[BenefitAdministrationExtraction] = []
    for page_number, page_text in split_marked_pages(text):
        if not re.search(
            r"Your\s+Insurers\s*,?\s*Service\s+Providers\s+and\s+Policy\s+Numbers",
            page_text,
            flags=re.IGNORECASE,
        ):
            continue

        table_text = re.split(r"\*\s*This\s+benefit\s+is\s+Insured", page_text, maxsplit=1, flags=re.IGNORECASE)[0]
        searchable = clean_text(table_text)
        benefit_matches: list[tuple[int, int, str, str, bool]] = []
        for display_name, pattern, product_name, is_optional in benefit_patterns:
            for match in re.finditer(pattern, searchable, flags=re.IGNORECASE):
                benefit_matches.append((match.start(), match.end(), display_name, product_name, is_optional))

        benefit_matches.sort(key=lambda item: (item[0], -(item[1] - item[0])))
        non_overlapping: list[tuple[int, int, str, str, bool]] = []
        for match in benefit_matches:
            if any(match[0] < existing[1] and match[1] > existing[0] for existing in non_overlapping):
                continue
            non_overlapping.append(match)
        non_overlapping.sort(key=lambda item: item[0])
        if not non_overlapping:
            continue

        insurer_occurrences = _grouphealth_organization_occurrences(searchable, insurer_patterns)
        provider_occurrences = _grouphealth_organization_occurrences(searchable, provider_patterns)
        policy_occurrences = [
            (match.start(), match.end(), match.group(0))
            for match in re.finditer(
                r"\b(?:[A-Za-z0-9]*\d[A-Za-z0-9]*-[A-Za-z0-9-]+|\d{4,10})\b",
                searchable,
            )
        ]

        used_policy_indexes: set[int] = set()
        previous_policy_end = 0
        for index, (start, end, display_name, product_name, is_optional) in enumerate(non_overlapping):
            next_benefit_start = non_overlapping[index + 1][0] if index + 1 < len(non_overlapping) else len(searchable)
            policy_index = next(
                (
                    candidate
                    for candidate, occurrence in enumerate(policy_occurrences)
                    if candidate not in used_policy_indexes and occurrence[0] >= start and occurrence[0] < next_benefit_start
                ),
                None,
            )
            policy_occurrence = policy_occurrences[policy_index] if policy_index is not None else None
            if policy_index is not None:
                used_policy_indexes.add(policy_index)

            policy_start = policy_occurrence[0] if policy_occurrence else next_benefit_start
            policy_end = policy_occurrence[1] if policy_occurrence else end
            insurer_candidates = [
                occurrence
                for occurrence in insurer_occurrences
                if occurrence[0] >= previous_policy_end and occurrence[0] < policy_start
            ]
            insurer_occurrence = min(
                insurer_candidates,
                key=lambda occurrence: abs(((occurrence[0] + occurrence[1]) // 2) - ((start + end) // 2)),
                default=None,
            )
            if display_name in {"Medical Second Opinion", "Virtual Healthcare (VHC)"}:
                insurer_occurrence = None
            insurer = insurer_occurrence[2] if insurer_occurrence else ""

            provider_candidates = [
                occurrence
                for occurrence in provider_occurrences
                if occurrence[0] >= end and occurrence[0] < next_benefit_start
            ]
            providers = unique_in_order([occurrence[2] for occurrence in provider_candidates])
            policy = policy_occurrence[2] if policy_occurrence else ""
            if not insurer and not providers and not policy:
                continue

            window_start = min(start, insurer_occurrence[0] if insurer_occurrence else start)
            window_end = next_benefit_start
            source_text = searchable[window_start:window_end].strip()
            confidence = 0.95 if policy and (insurer or providers) else 0.80 if insurer or providers else 0.65
            records.append(
                BenefitAdministrationExtraction(
                    source_file=pdf_path.name,
                    client_name=client_info.client_name,
                    benefit_display_name=display_name,
                    insurer_display_name=insurer,
                    service_provider_display_names=providers,
                    policy_number=policy,
                    product_name=product_name,
                    is_optional=is_optional,
                    source_page_number=page_number,
                    source_text=source_text,
                    extraction_confidence=confidence,
                )
            )
            previous_policy_end = policy_end

    return records


def _grouphealth_organization_occurrences(
    text: str,
    patterns: list[tuple[str, str]],
) -> list[tuple[int, int, str]]:
    occurrences: list[tuple[int, int, str]] = []
    for display_name, pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            occurrences.append((match.start(), match.end(), display_name))
    return sorted(occurrences, key=lambda item: item[0])


def build_grouphealth_carrier_policy_summary_lines(
    full_text: str,
    client_info: ClientInfo,
) -> list[str]:
    """Build carrier/policy lines from the booklet's administration table."""
    records = extract_grouphealth_benefit_administration(
        full_text,
        Path(client_info.client_name or "grouphealth-booklet"),
        client_info,
    )
    benefit_labels = {
        "Group Life Insurance": "Life",
        "Dependent Life Insurance": "Dependent Life",
        "Accident & Serious Illness (ASI)": "ASI",
        "Group Critical Illness": "Group CI",
        "Extended Health Care": "Extended Health",
        "Travel Insurance and Assistance / Travel Cancellation Insurance": "Travel",
        "Dental Care": "Dental",
        "Protector Series Optional Life Insurance": "Optional Life",
        "Protector Series Optional Accident & Serious Illness (ASI)": "Optional ASI",
        "Protector Series Optional Critical Illness": "Optional CI",
    }
    grouped: dict[tuple[str, str], list[str]] = {}
    for record in records:
        if not record.insurer_display_name:
            continue
        insurer = "SSQ" if record.insurer_display_name.startswith("SSQ") else record.insurer_display_name
        key = (insurer, format_policy_number(record.policy_number))
        grouped.setdefault(key, [])
        benefit_label = benefit_labels.get(record.benefit_display_name, record.benefit_display_name)
        if benefit_label not in grouped[key]:
            grouped[key].append(benefit_label)

    lines: list[str] = []
    for (insurer, policy), benefits in grouped.items():
        policy_text = f": {policy}" if policy else ""
        benefit_text = f" ({', '.join(benefits)})" if benefits else ""
        lines.append(f"{insurer}{policy_text}{benefit_text}")
    return lines


def apply_grouphealth_administration_to_plan_rows(
    rows: list[ExtractionRow],
    records: list[BenefitAdministrationExtraction],
) -> None:
    """Attach the booklet's benefit-specific carrier and policy to each plan row."""
    record_by_benefit = {record.benefit_display_name: record for record in records}
    for row in rows:
        if row.benefit == "Life":
            record_name = "Group Life Insurance"
        elif row.benefit == "Dependent Life":
            record_name = "Dependent Life Insurance"
        elif row.benefit == "AD&D":
            record_name = "Accident & Serious Illness (ASI)"
        elif row.benefit == "Critical Illness":
            record_name = "Group Critical Illness"
        elif row.benefit in {
            "Drugs",
            "Prescription Drugs",
            "Medical Services & Supplies",
            "Paramedical",
            "Vision",
            "Survivor Benefit",
        }:
            record_name = "Extended Health Care"
        elif row.benefit.startswith("Dental "):
            record_name = "Dental Care"
        elif row.benefit in {"Travel", "Travel Cancellation Insurance"}:
            record_name = "Travel Insurance and Assistance / Travel Cancellation Insurance"
        elif row.benefit == "Medical Second Opinion":
            record_name = "Medical Second Opinion"
        elif row.benefit == "EAP":
            record_name = "Employee & Family Assistance Program (EFAP)"
        elif row.benefit == "Virtual Care":
            record_name = "Virtual Healthcare (VHC)"
        else:
            continue

        record = record_by_benefit.get(record_name)
        if not record:
            continue
        administrator = record.insurer_display_name or first_non_empty(
            *record.service_provider_display_names
        )
        if administrator:
            row.carrier = administrator
        row.policy_number = format_policy_number(record.policy_number)


def extract_grouphealth_sparrow_section(text: str) -> str:
    cleaned = clean_text(text)
    match = re.search(r"\bSparrow\b", cleaned, flags=re.IGNORECASE)
    if not match:
        return ""
    start = max(0, match.start() - 300)
    end = min(len(cleaned), match.end() + 1800)
    return cleaned[start:end]


def build_employee_benefits_summary(
    full_text: str,
    benefit_text: str,
    pdf_path: Path,
    client_info: ClientInfo,
    rules: ExtractionRules,
    class_name_override: str = "",
) -> EmployeeBenefitsSummary:
    class_name = class_name_override or combine_file_class_with_detected_class(detect_class_name(full_text), pdf_path)
    if is_allstate_critical_illness_policy(full_text, client_info):
        return EmployeeBenefitsSummary(
            source_file=pdf_path.name,
            client_name=client_info.client_name,
            carrier="Allstate",
            class_name=class_name,
            rows=build_allstate_critical_illness_summary_rows(full_text, client_info),
        )

    if is_benepac_booklet(full_text):
        return EmployeeBenefitsSummary(
            source_file=pdf_path.name,
            client_name=client_info.client_name,
            carrier=client_info.carrier,
            class_name=class_name,
            rows=build_benepac_summary_rows(full_text),
        )
    if is_equitable_booklet(full_text, client_info):
        return EmployeeBenefitsSummary(
            source_file=pdf_path.name,
            client_name=client_info.client_name,
            carrier=client_info.carrier,
            class_name=class_name,
            rows=build_equitable_summary_rows(full_text),
        )

    if is_desjardins_booklet(full_text, client_info):
        return EmployeeBenefitsSummary(
            source_file=pdf_path.name,
            client_name=client_info.client_name,
            carrier=client_info.carrier,
            class_name=class_name,
            rows=build_desjardins_summary_rows(full_text, client_info),
        )

    if is_sunlife_benefit_summary_booklet(full_text, client_info):
        return EmployeeBenefitsSummary(
            source_file=pdf_path.name,
            client_name=client_info.client_name,
            carrier=client_info.carrier or "Sun Life",
            class_name=class_name,
            rows=build_sunlife_summary_rows(full_text, client_info),
        )

    summary_text = extract_schedule_pages_for_summary(benefit_text) or benefit_text
    search_lines = build_search_lines(summary_text)
    summary_rows: list[tuple[str, str]] = []
    is_grouphealth = is_grouphealth_booklet(full_text, client_info)

    for section in rules.summary_sections:
        if is_grouphealth and section.name == "Digital Wellness Program":
            continue

        section_text = resolve_summary_section_text(section, full_text, summary_text, search_lines)
        detail_lines = build_summary_detail_lines(section, section_text, full_text, client_info)
        detail_lines = apply_flex_plan_summary_rules(section.name, detail_lines, full_text)
        detail_lines = apply_grouphealth_summary_rules(
            section.name,
            detail_lines,
            full_text,
            is_grouphealth,
            client_info,
        )
        if should_skip_summary_section(section, section_text, detail_lines):
            continue
        if section.optional and not detail_lines:
            continue
        summary_rows.append((section.name, "\n".join(detail_lines)))

    return EmployeeBenefitsSummary(
        source_file=pdf_path.name,
        client_name=client_info.client_name,
        carrier=client_info.carrier,
        class_name=class_name,
        rows=summary_rows,
    )


def apply_flex_plan_summary_rules(
    section_name: str,
    detail_lines: list[str],
    full_text: str,
) -> list[str]:
    if not is_flex_plan_booklet(full_text):
        return detail_lines

    flex_builders = {
        "Life / AD&D": build_flex_life_summary_lines,
        "Long-Term Disability": build_flex_ltd_summary_lines,
        "Prescription Drugs": build_flex_prescription_drug_summary_lines,
        "Medical Services & Supplies": build_flex_ehc_summary_lines,
        "Paramedicals": build_flex_paramedical_summary_lines,
        "Hospital": build_flex_hospital_summary_lines,
        "Vision Care": build_flex_vision_summary_lines,
        "Dental Care": build_flex_dental_summary_lines,
    }
    builder = flex_builders.get(section_name)
    if not builder:
        return detail_lines

    flex_lines = builder(full_text)
    return flex_lines or detail_lines


def is_flex_plan_booklet(text: str) -> bool:
    cleaned = clean_text(text)
    return bool(
        re.search(r"\bFlex Plan\b", cleaned, flags=re.IGNORECASE)
        or (
            re.search(r"\bAll Options\b", cleaned, flags=re.IGNORECASE)
            and re.search(r"\bOption\s+1\b", cleaned, flags=re.IGNORECASE)
            and re.search(r"\bOption\s+2\b", cleaned, flags=re.IGNORECASE)
            and re.search(r"\bOption\s+3\b", cleaned, flags=re.IGNORECASE)
        )
    )


def build_flex_life_summary_lines(full_text: str) -> list[str]:
    page = find_schedule_page_text(full_text, "Group Life Insurance", "All Options", "Benefit Amount")
    if not page:
        return []

    benefit_amount = first_regex(page, [r"\bBenefit Amount\s+(.{1,260}?)(?=\s+Maximum Benefit\b)"])
    max_with = first_regex(page, [rf"\bMaximum Benefit\s+({MONEY_AMOUNT_PATTERN})\s+with Evidence of Insurability"])
    max_without = first_regex(page, [rf"\bMaximum Benefit\s+({MONEY_AMOUNT_PATTERN})\s+without Evidence of Insurability"])
    benefit_reduction = first_regex(page, [r"\bBenefit Reduction\s+(.{1,120}?)(?=\s+Waiver of Premium\b)"])
    waiver = first_regex(page, [r"\bWaiver of Premium\s+(.{1,360}?)(?=\s+Definition of Total Disability\b)"])
    termination = first_regex(page, [r"\bTermination Age\s+(.{1,180}?)(?=$|\s+Accidental Death|\s+Schedule of Benefits)"])

    return compact_detail_lines(
        [
            "Applies To: All Options",
            f"Life Benefit Amount: {benefit_amount}",
            f"Maximum With Evidence: {max_with}",
            f"Maximum Without Evidence: {max_without}",
            f"Benefit Reduction: {benefit_reduction}",
            f"Waiver of Premium: {waiver}",
            f"Termination Age: {termination}",
        ]
    )


def build_flex_ltd_summary_lines(full_text: str) -> list[str]:
    page = find_schedule_page_text(full_text, "Long-Term Disability", "Maximum From All Sources")
    if not page:
        return []

    benefit_amount = first_regex(page, [r"\bBenefit Amount\s+(.{1,160}?)(?=\s+Maximum Benefit\b)"])
    max_with = first_regex(page, [rf"\bwith Evidence of Insurability\s+({MONEY_AMOUNT_PATTERN}\s+per\s+month)"])
    max_without = first_regex(page, [rf"\bwithout Evidence of Insurability\s+({MONEY_AMOUNT_PATTERN}\s+per\s+month)"])
    all_sources = first_regex(page, [r"\bMaximum From All Sources\s+(.{1,180}?)(?=\s+Elimination Period\b)"])
    benefit_period = first_regex(page, [r"\bMaximum Benefit Period\s+(.{1,80}?)(?=\s+Definition of Total Disability\b)"])
    disability_definition = first_regex(page, [r"\bDefinition of Total Disability\s+(.{1,180}?)(?=\s+Taxability of Benefits\b)"])
    tax_status = first_regex(page, [r"\bTaxability of Benefits\s+(.{1,80}?)(?=\s+Work Re-entry Program\b)"])
    termination = first_regex(page, [r"\bTermination Age\s+(.{1,180}?)(?=$|\s+Schedule of Benefits)"])
    elimination_periods = unique_regex_group_matches(
        page,
        [r"\b(\d+\s+day\(s\)\s+for\s+(?:Accident|Hospitalization|Sickness))"],
    )

    lines = compact_detail_lines(
        [
            "Applies To: All Options",
            f"Benefit Schedule: {benefit_amount}",
            f"Maximum With Evidence: {max_with}",
            f"Maximum Without Evidence: {max_without}",
            f"Maximum From All Sources: {all_sources}",
        ]
    )
    if elimination_periods:
        lines.append("Elimination Period:")
        lines.extend(f"- {period}" for period in elimination_periods)
    lines.extend(
        compact_detail_lines(
            [
                f"Benefit Period: {benefit_period}",
                f"Definition: {disability_definition}",
                f"Tax Status: {tax_status}",
                f"Termination Age: {termination}",
            ]
        )
    )
    return lines


def build_flex_prescription_drug_summary_lines(full_text: str) -> list[str]:
    ehc_page = find_schedule_page_text(full_text, "Extended Health Care", "% Reimbursement")
    if not ehc_page:
        return []

    options = extract_flex_ehc_option_rows(ehc_page)
    drug_row = next((row for row in options if row[0] == "Prescription Drug"), None)
    deductible = first_regex(ehc_page, [r"\bCalendar Year Deductible\s+(.{1,80}?)(?=\s+Deductible Does Not Apply To\b)"])
    dispensing_fee = first_regex(ehc_page, [r"\bDispensing Fee Cap\s+(.{1,140}?)(?=$|\s+Hospital\b|\s+Family Services)"])
    drug_max = extract_drug_maximum(ehc_page)
    drug_type = extract_drug_type(ehc_page)

    lines = compact_detail_lines(
        [
            f"Deductible: {deductible}",
            "Plan Type: Pay Direct Drug Card",
            f"Drug Max: {drug_max}",
            f"Drug Type: {drug_type}",
        ]
    )
    if drug_row:
        lines.extend(["Reimbursement Options:", format_option_row_header(("Option 1", "Option 2", "Option 3"))])
        lines.append(format_option_row(drug_row))
    if dispensing_fee:
        lines.append(f"Dispensing Fee Cap: {dispensing_fee}")
    return lines


def build_flex_ehc_summary_lines(full_text: str) -> list[str]:
    ehc_page = find_schedule_page_text(full_text, "Extended Health Care", "% Reimbursement")
    if not ehc_page:
        return []

    deductible = first_regex(ehc_page, [r"\bCalendar Year Deductible\s+(.{1,80}?)(?=\s+Deductible Does Not Apply To\b)"])
    rows = extract_flex_ehc_option_rows(ehc_page)
    lines = compact_detail_lines(
        [
            "Applies To: All Options",
            f"Deductible: {deductible}",
        ]
    )
    if rows:
        lines.extend(["Reimbursement Options:", format_option_row_header(("Option 1", "Option 2", "Option 3"))])
        lines.extend(format_option_row(row) for row in rows)
    lines.extend(
        [
            "Travel/Referral/Travel Cancellation: 100% for all options",
            "Note: Eye exams and nursing are reimbursed at 100%",
        ]
    )
    return lines


def build_flex_hospital_summary_lines(full_text: str) -> list[str]:
    page = find_schedule_page_text(full_text, "Hospital", "Health Care Practitioners")
    if not page:
        return []
    room = first_regex(page, [r"\bHospital Room\s+(.{1,80}?)(?=\s+Hospital Indemnity\b)"])
    indemnity = first_regex(page, [r"\bHospital Indemnity\s+(.{1,220}?)(?=\s+Convalescent Care\b)"])
    convalescent = first_regex(page, [r"\bConvalescent Care\s+(.{1,180}?)(?=\s+Health Care Practitioners\b)"])
    return compact_detail_lines(
        [
            "Applies To: All Options",
            "Coinsurance: 100%",
            f"Coverage: {room}",
            f"Hospital Indemnity: {indemnity}",
            f"Convalescent Care: {convalescent}",
        ]
    )


def build_flex_paramedical_summary_lines(full_text: str) -> list[str]:
    page = find_schedule_page_text(full_text, "Health Care Practitioners", "Acupuncturist", "Option 1")
    if not page:
        return []

    rows = extract_flex_paramedical_option_rows(page)
    if not rows:
        return []
    return [
        "Maximums are per person per calendar year.",
        "Practitioner Options:",
        format_option_row_header(("Option 1", "Option 2")),
        *[format_option_row(row) for row in rows],
    ]


def build_flex_vision_summary_lines(full_text: str) -> list[str]:
    page = find_schedule_page_text(full_text, "Vision Care", "Visual Training")
    if not page:
        return []
    adult_max = first_regex(page, [rf"\bAdults\b.{{0,80}}?\bmaximum\s+({MONEY_AMOUNT_PATTERN}\s+every\s+\d+\s+consecutive\s+months)"])
    child_max = first_regex(page, [rf"\bDependent Children\b.{{0,80}}?\bmaximum\s+({MONEY_AMOUNT_PATTERN}\s+every\s+\d+\s+consecutive\s+months)"])
    visual_training = first_regex(page, [rf"\bVisual Training\b.{{0,120}}?(Lifetime maximum\s+{MONEY_AMOUNT_PATTERN})"])
    return compact_detail_lines(
        [
            "Option: Option 1 only",
            "Coverage: Eyeglasses, contact lenses and laser vision correction",
            f"Adult Maximum: {adult_max}",
            f"Dependent Children Maximum: {child_max}",
            f"Visual Training: {visual_training}",
        ]
    )


def build_flex_dental_summary_lines(full_text: str) -> list[str]:
    pages = find_schedule_pages_text(full_text, "Dental Care Coverage", "% Payment", max_pages=2)
    if not pages:
        return []
    dental_text = clean_text(" ".join(pages))

    deductible = first_regex(dental_text, [r"\bCalendar Year Deductible\s+All Options\s+(No Deductible)", r"\bCalendar Year Deductible\s+(No Deductible)"])
    recall = extract_recall_frequency(dental_text)
    basic = extract_flex_three_option_values(
        dental_text,
        r"Basic Dental Care\s+&\s+Routine Dental Care",
    )
    major = extract_flex_three_option_values(
        dental_text,
        r"Dental Restorative\s+Services\s+\(subject to\s+certain limits\)",
    )
    ortho = extract_flex_three_option_values(
        dental_text,
        r"Orthodontic Care\s+\(subject to certain\s+limits\)",
    )
    basic_max = extract_flex_dental_basic_maximums(dental_text)
    ortho_max = extract_flex_dental_ortho_maximums(dental_text)

    lines = compact_detail_lines([f"Deductible: {deductible}"])
    if basic:
        lines.extend(["Basic/Routine Dental Care:", format_option_row(("Basic/Routine", *basic))])
    if major:
        lines.extend(["Major Restorative Services:", format_option_row(("Major Restorative", *major))])
    if ortho:
        lines.extend(["Orthodontic Care:", format_option_row(("Orthodontics", *ortho)), "- Dependent children under age 26 only"])
    if recall:
        lines.append(f"Recall: {recall}")
    if basic_max:
        lines.extend(["Basic/Routine/Major Maximums:", format_option_row_header(("Option 1", "Option 2", "Option 3")), format_option_row(("Maximum", *basic_max))])
    if ortho_max:
        lines.extend(["Orthodontic Maximums:", format_option_row(("Orthodontics", *ortho_max))])
    return lines


def find_schedule_page_text(full_text: str, *terms: str) -> str:
    _page_number, page_text = find_schedule_page(full_text, *terms)
    return page_text


def find_schedule_page(full_text: str, *terms: str) -> tuple[int, str]:
    pages = split_marked_pages(full_text)
    normalized_terms = [term.lower() for term in terms if term]
    for page_number, page_text in pages:
        cleaned_page = clean_text(page_text)
        normalized_page = cleaned_page.lower()
        if not is_schedule_heading_page(cleaned_page):
            continue
        if normalized_terms and not all(term in normalized_page for term in normalized_terms):
            continue
        return page_number, cleaned_page

    # GroupHEALTH schedules frequently continue across pages without repeating
    # the Schedule of Benefits heading. Limit the fallback to the pages directly
    # following a confirmed schedule heading so later policy wording is ignored.
    schedule_starts = [
        index
        for index, (_page_number, page_text) in enumerate(pages)
        if is_schedule_heading_page(page_text)
    ]
    for start_index in schedule_starts:
        end_index = min(len(pages), start_index + 15)
        for page_number, page_text in pages[start_index + 1 : end_index]:
            cleaned_page = clean_text(page_text)
            normalized_page = cleaned_page.lower()
            if normalized_terms and not all(term in normalized_page for term in normalized_terms):
                continue
            return page_number, cleaned_page
    return 0, ""


def find_grouphealth_dental_schedule_page(full_text: str) -> tuple[int, str]:
    for terms in (
        ("Dental Care Coverage", "% Payment"),
        ("Dental Care Coverage", "Basic Dental Care"),
    ):
        page_number, page_text = find_schedule_page(full_text, *terms)
        if page_text and re.search(
            r"\b(?:Basic\s+Dental\s+Care|Dental\s+Restorative\s+Services|Orthodontic\s+Care)\b",
            page_text,
            flags=re.IGNORECASE,
        ):
            return page_number, page_text
    return 0, ""


def is_schedule_heading_page(page_text: str) -> bool:
    lines = [clean_text(line) for line in (page_text or "").splitlines() if clean_text(line)]
    if not lines:
        return False
    if lines[0].lower().startswith("table of contents"):
        return False
    return any(
        re.match(r"^schedule of benefits\b", line, flags=re.IGNORECASE)
        for line in lines[:8]
    )


def find_schedule_pages_text(full_text: str, *terms: str, max_pages: int = 1) -> list[str]:
    pages = split_marked_pages(full_text)
    normalized_terms = [term.lower() for term in terms if term]
    for index, (_page_number, page_text) in enumerate(pages):
        cleaned_page = clean_text(page_text)
        normalized_page = cleaned_page.lower()
        if not is_schedule_heading_page(cleaned_page):
            continue
        if normalized_terms and not all(term in normalized_page for term in normalized_terms):
            continue
        return [clean_text(pages[selected_index][1]) for selected_index in range(index, min(len(pages), index + max_pages))]
    return []


def extract_flex_ehc_option_rows(page: str) -> list[tuple[str, str, str, str]]:
    option_rows: list[tuple[str, str, str, str]] = []
    labels = [
        "Prescription Drug",
        "Hospital",
        "Health Care Practitioners",
        "Vision Care",
        "Other Medical Expenses",
    ]
    for label in labels:
        pattern = rf"\b{re.escape(label)}\s+(\d{{2,3}}%\*?|N/A)\s+(\d{{2,3}}%\*?|N/A)\s+(\d{{2,3}}%\*?|N/A)"
        match = re.search(pattern, page, flags=re.IGNORECASE)
        if match:
            option_rows.append((label, match.group(1), match.group(2), match.group(3)))
    return option_rows


def extract_flex_paramedical_option_rows(page: str) -> list[tuple[str, str, str]]:
    practitioner_patterns = [
        ("Acupuncturist", r"Acupuncturist"),
        ("Chiropractor", r"Chiropractor(?:\s+\(includes x-rays\))?"),
        ("Massage Therapist/Orthotherapist", r"Massage Therapist/Orthotherapist"),
        ("Naturopath", r"Naturopath"),
        ("Occupational Therapist", r"Occupational Therapist"),
        ("Osteopath", r"Osteopath(?:\s+\(includes x-rays\))?"),
        ("Physiotherapist/Physical Rehabilitation Therapist", r"Physiotherapist/Physical Rehabilitation Therapist"),
        ("Podiatrist/Chiropodist", r"Podiatrist/Chiropodist(?:\s+\(includes x-rays\))?"),
        ("Psychologist/Social Worker/Registered Clinical Counselor/Psychotherapist", r"Psychologist/Social Worker/\s*Registered Clinical Counsell?or/Psychotherapist"),
        ("Speech Therapist", r"Speech Therapist"),
    ]
    rows: list[tuple[str, str, str]] = []
    for display, pattern in practitioner_patterns:
        match = re.search(rf"\b{pattern}\s+({MONEY_AMOUNT_PATTERN}|N/A)\s+({MONEY_AMOUNT_PATTERN}|N/A)", page, flags=re.IGNORECASE)
        if match:
            rows.append((display, clean_money(match.group(1)) if "$" in match.group(1) else match.group(1), clean_money(match.group(2)) if "$" in match.group(2) else match.group(2)))
    return rows


def extract_flex_three_option_values(text: str, label_pattern: str) -> tuple[str, str, str]:
    match = re.search(
        rf"\b{label_pattern}\s+Option\s+1\s+(\d{{2,3}}%|N/A)\s+Option\s+2\s+(\d{{2,3}}%|N/A)\s+Option(?:\s+3)?\s+(\d{{2,3}}%|N/A)",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        return ("", "", "")
    return (match.group(1), match.group(2), match.group(3))


def extract_flex_dental_basic_maximums(text: str) -> tuple[str, str, str]:
    match = re.search(
        rf"\bBasic Dental Care/Routine Dental Care/Major Restorative Services\s+Option\s+1\s+Option\s+2\s+Option\s+3\s+"
        rf"(combined maximum of\s+{MONEY_AMOUNT_PATTERN}.{{1,90}}?calendar year)\s+"
        rf"(combined maximum of\s+{MONEY_AMOUNT_PATTERN}.{{1,90}}?calendar year)\s+"
        rf"(combined maximum of\s+{MONEY_AMOUNT_PATTERN}.{{1,90}}?calendar year)",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        return ("", "", "")
    return tuple(clean_text(match.group(index)) for index in range(1, 4))  # type: ignore[return-value]


def extract_flex_dental_ortho_maximums(text: str) -> tuple[str, str, str]:
    match = re.search(
        rf"\bOrthodontic Care\s+Option\s+1\s+Option\s+2\s+Option\s+3\s+({MONEY_AMOUNT_PATTERN}.{{1,80}}?lifetime)\s+(N/A)\s+(N/A)",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        return ("", "", "")
    return (normalize_money_phrase(match.group(1)), match.group(2), match.group(3))


def format_option_row_header(options: tuple[str, ...]) -> str:
    return "Item | " + " | ".join(options)


def format_option_row(row: tuple[str, ...]) -> str:
    return " | ".join(clean_text(part) for part in row)


def compact_detail_lines(lines: list[str]) -> list[str]:
    compacted: list[str] = []
    for line in lines:
        if ":" in line:
            label, value = line.split(":", 1)
            if not clean_text(value):
                continue
            compacted.append(f"{clean_text(label)}: {clean_flex_detail_value(value)}")
        elif clean_text(line):
            compacted.append(clean_flex_detail_value(line))
    return unique_in_order(compacted)


def clean_flex_detail_value(value: str) -> str:
    cleaned = clean_text(value)
    cleaned = re.sub(r"[\uf0a1▪•]\s*", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip(" -;:")


def apply_grouphealth_summary_rules(
    section_name: str,
    detail_lines: list[str],
    full_text: str,
    is_grouphealth: bool,
    client_info: ClientInfo,
) -> list[str]:
    if not is_grouphealth:
        return detail_lines

    if section_name == "Carrier & Policy":
        return build_grouphealth_carrier_policy_summary_lines(full_text, client_info) or detail_lines

    if section_name in {"Employee Assistance Program", "Virtual Care"}:
        return ["Coverage: Sparrow"]

    if section_name == "Medical Second Opinion":
        lines = unique_in_order(detail_lines)
        has_provider = any(line.lower().startswith("provider:") for line in lines)
        has_coverage = any(line.lower().startswith("coverage:") for line in lines)
        if has_provider and not has_coverage:
            lines.append("Coverage: Included")
        return lines or ["Coverage: Included"]

    if section_name == "Digital Wellness Program":
        return []

    if section_name == "Dental Care":
        return build_grouphealth_schedule_dental_summary_lines(full_text)

    return detail_lines


def build_grouphealth_schedule_dental_summary_lines(full_text: str) -> list[str]:
    page_number, page_text = find_grouphealth_dental_schedule_page(full_text)
    if not page_text:
        return []

    marked_page = f"[[PAGE {page_number}]]\n{page_text}"
    rates = {
        rate.provision_code: rate
        for rate in map_labeled_dental_rates(marked_page)
        if rate.provision_code != "routine"
    }
    maximums = {
        maximum.provision_code: maximum
        for maximum in map_labeled_dental_maximums(marked_page)
        if maximum.provision_code != "routine"
    }
    statuses = {
        status.provision_code: status
        for status in map_labeled_dental_coverage_statuses(marked_page)
        if status.provision_code != "routine"
    }
    provision_labels = {
        "preventive": "Preventive Services",
        "basic": "Basic Services",
        "endodontic_periodontic": "Endodontic / Periodontic Services",
        "dentures": "Dentures",
        "major": "Major Restorative Services",
        "orthodontic": "Orthodontics",
        "accidental_injury": "Accidental Dental Injury",
        "other_treatment": "Other Dental Treatment",
    }
    provision_order = tuple(provision_labels)
    lines: list[str] = []

    for code in provision_order:
        status = statuses.get(code)
        rate = rates.get(code)
        if status:
            lines.append(f"{provision_labels[code]}: {status.coverage_status}")
        elif rate:
            lines.append(f"{provision_labels[code]}: {rate.reimbursement_text}")

    deductible = extract_grouphealth_dental_deductible(page_text)
    if deductible:
        lines.append(f"Deductible: {deductible}")

    maximum_by_code = {
        code: rates[code].maximum
        for code in rates
        if rates[code].maximum and code not in statuses
    }
    for code, maximum in maximums.items():
        if code not in statuses and code not in maximum_by_code:
            maximum_by_code[code] = maximum.maximum
    distinct_maximums = unique_in_order(maximum_by_code.values())
    if len(distinct_maximums) == 1:
        lines.append(f"Maximum: {distinct_maximums[0]}")
    elif len(distinct_maximums) > 1:
        for code in provision_order:
            if code in maximum_by_code:
                lines.append(f"{provision_labels[code]} Maximum: {maximum_by_code[code]}")

    recall = extract_recall_frequency(page_text)
    if recall:
        lines.append(f"Recall: {recall}")

    fee_guide = extract_grouphealth_dental_fee_guide(page_text)
    if fee_guide:
        lines.append(f"Fee Guide: {fee_guide}")
    termination_age = extract_schedule_termination_age_after(full_text, page_number)
    if termination_age:
        lines.append(f"Termination Age: {termination_age}")
    return unique_in_order(lines)


def extract_grouphealth_dental_deductible(page_text: str) -> str:
    cleaned = clean_text(page_text)
    if re.search(
        r"\bCalendar\s+Year\b.{0,100}\bNo\s+Deductible\b|"
        r"\bNo\s+Deductible\b.{0,100}\bDeductible\b",
        cleaned,
        flags=re.IGNORECASE,
    ):
        return "None"

    deductible = first_regex(
        cleaned,
        [
            rf"\bCalendar\s+Year\s+Deductible\b.{{0,120}}?({MONEY_AMOUNT_PATTERN}(?:\s+(?:Single|Family))?(?:\s*/\s*{MONEY_AMOUNT_PATTERN}(?:\s+(?:Single|Family))?)?)",
            rf"\bCalendar\s+Year\b.{{0,80}}?({MONEY_AMOUNT_PATTERN}(?:\s+(?:Single|Family))?(?:\s*/\s*{MONEY_AMOUNT_PATTERN}(?:\s+(?:Single|Family))?)?).{{0,80}}?\bDeductible\b",
        ],
    )
    return clean_text(deductible)


def extract_grouphealth_dental_fee_guide(page_text: str) -> str:
    cleaned = clean_text(page_text)
    if re.search(
        r"\bCurrent\s+fee\s+guide\s+for\s+general\s+practitioners\b",
        cleaned,
        flags=re.IGNORECASE,
    ):
        return "Current provincial fee guide for general practitioners"

    return first_regex(
        cleaned,
        [
            r"\bDental\s+(?:Procedure\s+)?Fee\s+Guide\s*:?\s*(current\s+year|one\s+year\s+prior|two\s+years?\s+prior)",
            r"\bFee\s+Guide\s*:?\s*(current\s+year|one\s+year\s+prior|two\s+years?\s+prior)",
        ],
    )


def should_skip_summary_section(
    section: SummarySectionDefinition,
    section_text: str,
    detail_lines: list[str],
) -> bool:
    if section.name != "Critical Illness":
        return False

    joined_details = clean_text(" ".join(detail_lines))
    if not joined_details:
        return True

    optional_only = re.search(
        r"\bOptional\s+(?:Personal\s+)?Critical\s+Illness\b|\bOptional\s+.*?\bCI\b",
        section_text,
        flags=re.IGNORECASE,
    )
    has_core_ci_value = re.search(
        r"\b(?:Principal Sum|Benefit Amount|Coverage):\s*(?:\$|\d+\s+(?:insured\s+)?illnesses|Comprehensive Conditions)",
        joined_details,
        flags=re.IGNORECASE,
    )

    if optional_only and not re.search(r"\bGroup Critical Illness\b", section_text, flags=re.IGNORECASE):
        return True

    return not bool(has_core_ci_value)


def extract_schedule_pages_for_summary(text: str) -> str:
    pages = split_marked_pages(text)
    if not pages:
        return ""

    selected_pages: list[str] = []
    for _page_number, page_text in pages:
        lines = [clean_text(line) for line in page_text.splitlines() if clean_text(line)]
        top_lines = " ".join(lines[:5])
        if re.search(r"\bSchedule of Benefits\b", top_lines, flags=re.IGNORECASE) or re.search(
            r"\bClass\s+\d+\s+-", top_lines, flags=re.IGNORECASE
        ):
            selected_pages.append(page_text)

    return "\n".join(selected_pages)


def resolve_summary_section_text(
    section: SummarySectionDefinition,
    full_text: str,
    benefit_text: str,
    search_lines: list[tuple[str, str]],
) -> str:
    if section.source == "document":
        return full_text
    if section.name == "Paramedicals":
        paramedical_text = choose_best_paramedical_summary_section(benefit_text, full_text)
        if paramedical_text:
            return paramedical_text

    definition = BenefitDefinition(section.name, section.keywords, section.headings)
    section_search_text = full_text if section.name == "Medical Second Opinion" else benefit_text
    section_text, _matched = find_benefit_section_by_page_heading(
        section_search_text,
        definition,
        section.page_span,
    )
    if not section_text:
        section_search_lines = build_search_lines(section_search_text)
        section_text, _matched = find_benefit_section(
            section_search_text,
            section_search_lines,
            definition,
        )
    return section_text


def extract_paramedical_summary_section(text: str) -> str:
    cleaned = clean_text(text)
    if not cleaned:
        return ""

    marker = re.search(
        r"\bHealth Care Practitioners\s*(?:[-\u2013\u2014]\s*)?Included\b",
        cleaned,
        flags=re.IGNORECASE,
    )
    if not marker:
        marker = re.search(r"\bPractitioner\s+Maximum\b", cleaned, flags=re.IGNORECASE)
    if not marker:
        return ""

    start = max(0, marker.start() - 700)
    remaining = cleaned[marker.end() :]
    end_match = re.search(
        r"\b(?:Vision Care|Other Medical Expenses|Hospital|Dental Care|Emergency Travel|Travel Insurance)\s*(?:[-\u2013\u2014]\s*)?(?:Included|Coverage)?\b",
        remaining,
        flags=re.IGNORECASE,
    )
    end = marker.end() + end_match.start() if end_match else min(len(cleaned), marker.end() + 2200)
    return cleaned[start:end]


def choose_best_paramedical_summary_section(summary_text: str, full_text: str) -> str:
    summary_section = extract_paramedical_summary_section(summary_text)
    full_section = extract_paramedical_summary_section(full_text)

    if not summary_section:
        return full_section
    if not full_section:
        return summary_section

    summary_count = len(split_semicolon_values(extract_paramedical_practitioners(summary_section)))
    full_count = len(split_semicolon_values(extract_paramedical_practitioners(full_section)))
    return full_section if full_count > summary_count else summary_section


def split_semicolon_values(value: str) -> list[str]:
    return [part.strip() for part in value.split(";") if part.strip()]


def build_summary_detail_lines(
    section: SummarySectionDefinition,
    section_text: str,
    full_text: str,
    client_info: ClientInfo,
) -> list[str]:
    detail_lines: list[str] = []
    source_text = full_text if section.source == "document" else section_text

    if section.name == "Carrier & Policy":
        policy_value = format_policy_number(client_info.group_number or client_info.policy_number)
        if client_info.carrier and policy_value:
            detail_lines.append(f"{client_info.carrier}: {policy_value}")
        elif client_info.carrier:
            detail_lines.append(client_info.carrier)

    schedule_lines = build_schedule_summary_override_lines(section.name, full_text)
    if schedule_lines:
        return add_documented_provider_line(section.name, schedule_lines, section_text or full_text)

    for field in section.fields:
        value = extract_summary_field_value(section, field, source_text)
        if not value:
            continue
        if field.label.lower() == "recall":
            value = normalize_recall_frequency(value)

        values = split_semicolon_values(value)
        if field.max_results > 1 and len(values) > 1:
            detail_lines.append(f"{field.label}:")
            detail_lines.extend(f"- {part}" for part in values)
        else:
            detail_lines.append(f"{field.label}: {value}")

    return add_documented_provider_line(section.name, unique_in_order(detail_lines), section_text)


def add_documented_provider_line(section_name: str, detail_lines: list[str], evidence: str) -> list[str]:
    if any(line.lower().startswith("provider:") for line in detail_lines):
        return detail_lines

    for provider_name, pattern in BENEFIT_PROVIDER_PATTERNS.get(section_name, ()):
        if re.search(pattern, evidence or "", flags=re.IGNORECASE):
            return [f"Provider: {provider_name}", *detail_lines]
    return detail_lines


def cost_sharing_line(items: Iterable[tuple[str, str]]) -> str:
    details: list[str] = []
    for label, percentage_text in items:
        match = re.search(r"\b(\d{1,3})%", percentage_text or "")
        if not match:
            continue
        plan_percent = int(match.group(1))
        if plan_percent < 0 or plan_percent > 100:
            continue
        member_percent = 100 - plan_percent
        prefix = f"{label}: " if label else ""
        details.append(f"{prefix}plan {plan_percent}% / member {member_percent}%")

    return f"Cost Sharing: {'; '.join(details)}" if details else ""


def build_schedule_summary_override_lines(section_name: str, full_text: str) -> list[str]:
    if is_benepac_booklet(full_text):
        return build_benepac_schedule_summary_lines(section_name, full_text)

    if is_flex_plan_booklet(full_text) and section_name != "Dependent Life":
        return []

    schedule_builders = {
        "Life / AD&D": build_schedule_life_summary_lines,
        "Dependent Life": build_schedule_dependent_life_summary_lines,
        "Short-Term Disability": build_schedule_std_summary_lines,
        "Long-Term Disability": build_schedule_ltd_summary_lines,
        "Critical Illness": build_schedule_critical_illness_summary_lines,
        "Prescription Drugs": build_schedule_prescription_drug_summary_lines,
        "Medical Services & Supplies": build_schedule_ehc_summary_lines,
        "Paramedicals": build_schedule_paramedical_summary_lines,
        "Vision Care": build_schedule_vision_summary_lines,
        "Emergency Travel": build_schedule_travel_summary_lines,
    }
    builder = schedule_builders.get(section_name)
    return builder(full_text) if builder else []


def is_allstate_critical_illness_policy(text: str, client_info: ClientInfo | None = None) -> bool:
    carrier = client_info.carrier if client_info else ""
    searchable = clean_text(f"{carrier} {text}")
    return bool(
        re.search(r"\bAllstate\b", searchable, flags=re.IGNORECASE)
        and re.search(r"\bPOLICY SPECIFICATIONS\b", searchable, flags=re.IGNORECASE)
        and re.search(r"\bBASIC BENEFIT AMOUNT\b", searchable, flags=re.IGNORECASE)
        and re.search(r"\bCritical Illness\b|\bGCIPNP\d*\b", searchable, flags=re.IGNORECASE)
    )


def extract_allstate_policy_specification_pages(full_text: str) -> list[tuple[int, str]]:
    pages = [
        (page_number, clean_text(page_text))
        for page_number, page_text in split_marked_pages(full_text)
        if re.search(r"\bPOLICY SPECIFICATIONS\b", page_text, flags=re.IGNORECASE)
    ]
    return pages


def extract_allstate_page_value(
    pages: list[tuple[int, str]],
    pattern: str,
) -> tuple[str, int | str]:
    for page_number, page_text in pages:
        match = re.search(pattern, page_text, flags=re.IGNORECASE)
        if match:
            return clean_text(match.group(1)).strip(" .;:-"), page_number
    return "", ""


def extract_allstate_critical_illness_items(full_text: str) -> list[tuple[str, str, int | str]]:
    pages = extract_allstate_policy_specification_pages(full_text)
    if not pages:
        return []

    items: list[tuple[str, str, int | str]] = []

    def add(label: str, pattern: str, transform=None) -> None:
        value, page_number = extract_allstate_page_value(pages, pattern)
        if not value:
            return
        normalized = transform(value) if transform else value
        if normalized:
            items.append((label, normalized, page_number))

    add("Provider", r"\b(Allstate)\s+Insurance Company of Canada\b")
    if not any(label == "Provider" for label, _value, _page in items):
        provider_page = next(
            (page_number for page_number, page_text in pages if re.search(r"\bAllstate\b", page_text, flags=re.IGNORECASE)),
            pages[0][0],
        )
        items.append(("Provider", "Allstate", provider_page))

    add(
        "Policyholder Legal Name",
        r"\bPOLICYHOLDER\s*:\s*(.{3,220}?)(?=\s+POLICY NUMBER\s*:)",
        lambda value: value.title().replace(" Dba ", " dba ") if value.isupper() else value,
    )
    add("Policy Number", r"\bPOLICY NUMBER\s*:\s*([A-Za-z0-9-]+)")
    add("Policy Effective Date", r"\bPOLICY EFFECTIVE DATE\s*:\s*(.{3,80}?)(?=\s+POLICY ANNIVERSARY DATE\s*:)")
    add(
        "Policy Anniversary Date",
        r"\bPOLICY ANNIVERSARY DATE\s*:\s*(.{3,260}?)(?=\s+GOVERNING JURISDICTION\s*:)",
        lambda value: re.sub(r"\bOctboer\b", "October", value, flags=re.IGNORECASE),
    )
    add(
        "Eligibility Waiting Period",
        r"\bELIGIBILITY WAITING PERIOD\s*:\s*(.{1,100}?)(?=\s+BASIC BENEFIT AMOUNT\s*:)",
    )
    add("Plan Type", r"\bBASIC BENEFIT AMOUNT\s*:\s*(.{1,100}?)(?=\s+\$)")
    add("Employee Benefit Amount", rf"({MONEY_AMOUNT_PATTERN})\s+for\s+Insured Employee or Member", clean_money)
    add("Spouse Benefit Amount", rf"({MONEY_AMOUNT_PATTERN})\s+for\s+Insured Spouse", clean_money)
    add("Child Benefit Amount", rf"({MONEY_AMOUNT_PATTERN})\s+for\s+Insured Child\(ren\)", clean_money)
    add(
        "Evidence of Insurability Requirement",
        r"\bGUARANTEED ISSUE LIMIT\s*:\s*(.{1,320}?)(?=\s+OPTIONAL BENEFIT\(S\)\s*:)",
    )

    evidence = next((value for label, value, _page in items if label == "Evidence of Insurability Requirement"), "")
    evidence_page = next((page for label, _value, page in items if label == "Evidence of Insurability Requirement"), "")
    evidence_threshold = first_regex(evidence, [rf"\bover\s+({MONEY_AMOUNT_PATTERN})"])
    if evidence_threshold:
        items.append(("Evidence Threshold", clean_money(evidence_threshold), evidence_page))

    add(
        "Optional Benefit",
        r"\bOPTIONAL BENEFIT\(S\)\s*:\s*(.{1,180}?)(?=\s+RIDER\(S\)\s*:)",
    )

    rider_text, rider_page = extract_allstate_page_value(pages, r"\bRIDER\(S\)\s*:\s*(.+)$")
    for rider in unique_regex_group_matches(
        rider_text,
        [
            r"((?:Critical Illness Enhancement|Amyotrophic Lateral Sclerosis \(ALS\) Benefit)\s+Rider)",
        ],
    ):
        items.append(("Rider", rider, rider_page))

    add(
        "Single Monthly Rate",
        rf"\bMonthly rate of\s+({MONEY_AMOUNT_PATTERN})\s+per insured employee or member for Single Coverage",
        lambda value: f"{clean_money(value)} per insured employee or member",
    )
    add(
        "Family Monthly Rate",
        rf"(?:;\s*or\s+|\bor\s+)?({MONEY_AMOUNT_PATTERN})\s+per insured employee or member for Family Coverage",
        lambda value: f"{clean_money(value)} per insured employee or member",
    )
    add("Rate Guarantee Date", r"\bRATE GUARANTEE DATE\s*:\s*([^\s]+)")
    return items


def extract_allstate_critical_illness_plan_design_rows(
    full_text: str,
    pdf_path: Path,
    client_info: ClientInfo,
    class_name: str,
) -> list[ExtractionRow]:
    rows: list[ExtractionRow] = []
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for label, value, source_page in extract_allstate_critical_illness_items(full_text):
        maximum = value if label in {"Employee Benefit Amount", "Spouse Benefit Amount", "Child Benefit Amount", "Evidence Threshold"} else ""
        rows.append(
            ExtractionRow(
                client_id=client_info.client_id,
                client_name=client_info.client_name,
                carrier="Allstate",
                class_name=class_name or "All Classes",
                benefit="Critical Illness",
                coverage_description=f"{label}: {value}",
                coinsurance="",
                maximum=maximum,
                frequency="Monthly" if label in {"Single Monthly Rate", "Family Monthly Rate"} else "",
                waiting_period=value if label == "Eligibility Waiting Period" else "",
                termination_age="",
                notes="Allstate Critical Illness policy specifications",
                source_file=pdf_path.name,
                source_section=label,
                confidence="High",
                review_flag="",
                last_updated=now,
                policy_number=client_info.policy_number,
                source_page=source_page,
            )
        )
    return rows


def build_allstate_critical_illness_summary_rows(
    full_text: str,
    client_info: ClientInfo,
) -> list[tuple[str, str]]:
    items = extract_allstate_critical_illness_items(full_text)
    values: dict[str, list[str]] = defaultdict(list)
    for label, value, _page in items:
        values[label].append(value)

    policy_number = first_non_empty(*(values.get("Policy Number", [])), client_info.policy_number)
    carrier_lines = [f"Allstate: {policy_number}" if policy_number else "Allstate"]
    for label in ("Policyholder Legal Name", "Policy Effective Date", "Policy Anniversary Date"):
        if values.get(label):
            carrier_lines.append(f"{label}: {values[label][0]}")

    eligibility_lines = []
    if values.get("Eligibility Waiting Period"):
        eligibility_lines.append(f"Waiting Period: {values['Eligibility Waiting Period'][0]}")

    critical_illness_lines = ["Provider: Allstate"]
    for label in (
        "Plan Type",
        "Employee Benefit Amount",
        "Spouse Benefit Amount",
        "Child Benefit Amount",
        "Evidence of Insurability Requirement",
        "Evidence Threshold",
        "Optional Benefit",
    ):
        if values.get(label):
            critical_illness_lines.append(f"{label}: {values[label][0]}")
    if values.get("Rider"):
        critical_illness_lines.append("Riders:")
        critical_illness_lines.extend(f"- {rider}" for rider in values["Rider"])
    for label in ("Single Monthly Rate", "Family Monthly Rate", "Rate Guarantee Date"):
        if values.get(label):
            critical_illness_lines.append(f"{label}: {values[label][0]}")

    rows = [("Carrier & Policy", "\n".join(carrier_lines))]
    if eligibility_lines:
        rows.append(("Eligibility Requirements", "\n".join(eligibility_lines)))
    rows.append(("Critical Illness", "\n".join(critical_illness_lines)))
    return rows


def is_schedule_first_booklet(text: str, client_info: ClientInfo | None = None) -> bool:
    return (
        is_allstate_critical_illness_policy(text, client_info)
        or is_benepac_booklet(text)
        or is_equitable_booklet(text, client_info)
    )


def is_equitable_booklet(text: str, client_info: ClientInfo | None = None) -> bool:
    carrier = client_info.carrier if client_info else ""
    cleaned = clean_text(f"{carrier} {text}")
    return bool(
        re.search(r"\bEquitable\b", cleaned, flags=re.IGNORECASE)
        and re.search(r"\bSchedule of benefits\b", cleaned, flags=re.IGNORECASE)
    )


def is_desjardins_booklet(text: str, client_info: ClientInfo | None = None) -> bool:
    carrier = clean_text(client_info.carrier if client_info else "")
    searchable = clean_text(f"{carrier} {text}")
    return bool(
        carrier.lower() == "desjardins"
        or re.search(r"\bDesjardins(?: Insurance| Financial Security)?\b", searchable, flags=re.IGNORECASE)
    )


def is_sunlife_benefit_summary_booklet(text: str, client_info: ClientInfo | None = None) -> bool:
    """Identify Sun Life booklets that contain the authoritative Benefit Summary."""
    carrier = clean_text(client_info.carrier if client_info else "")
    searchable = clean_text(f"{carrier} {text}")
    return bool(
        re.search(r"\bSun Life(?: Assurance Company of Canada| Financial)?\b", searchable, flags=re.IGNORECASE)
        and re.search(r"\bBenefit Summary\b", searchable, flags=re.IGNORECASE)
        and re.search(r"\bContract\s+(?:No\.?|Number)\s*:?\s*\d+", searchable, flags=re.IGNORECASE)
    )


def sunlife_page_text(full_text: str, page_number: int) -> str:
    for number, page_text in split_marked_pages(full_text):
        if number == page_number:
            return clean_text(page_text)
    return ""


def sunlife_find_page(full_text: str, *required_patterns: str) -> tuple[int, str]:
    """Find a non-table-of-contents page containing every required pattern."""
    for page_number, page_text in split_marked_pages(full_text):
        cleaned = clean_text(page_text)
        if "TABLE OF CONTENTS" in cleaned.upper():
            continue
        if all(re.search(pattern, cleaned, flags=re.IGNORECASE) for pattern in required_patterns):
            return page_number, cleaned
    return 0, ""


def sunlife_value(text: str, patterns: Iterable[str]) -> str:
    return clean_text(first_regex(text, patterns))


def sunlife_join(*values: str) -> str:
    return "; ".join(unique_in_order([clean_text(value) for value in values if clean_text(value)]))


def build_sunlife_plan_design_items(full_text: str) -> list[SunLifePlanDesignItem]:
    """Extract concise, source-backed values from Sun Life Benefit Summary pages."""
    items: list[SunLifePlanDesignItem] = []

    def add(
        summary_section: str,
        benefit: str,
        description: str = "",
        coinsurance: str = "",
        maximum: str = "",
        frequency: str = "",
        waiting_period: str = "",
        termination_age: str = "",
        source_section: str = "",
        source_page: int = 0,
    ) -> None:
        values = (description, coinsurance, maximum, frequency, waiting_period, termination_age)
        if not any(clean_text(value) for value in values):
            return
        items.append(
            SunLifePlanDesignItem(
                summary_section=summary_section,
                benefit=benefit,
                coverage_description=clean_text(description),
                coinsurance=clean_text(coinsurance),
                maximum=clean_text(maximum),
                frequency=clean_text(frequency),
                waiting_period=clean_text(waiting_period),
                termination_age=clean_text(termination_age),
                source_section=clean_text(source_section),
                source_page=source_page,
            )
        )

    summary_page, summary = sunlife_find_page(
        full_text,
        r"\bBenefit Summary\b",
        r"\bPrescription drugs\b",
        r"\bDeductible\b",
    )
    continuation_page, continuation = sunlife_find_page(
        full_text,
        r"\bParamedical services\b",
        r"\bOut-of-province\b",
    )
    dental_page, dental = sunlife_find_page(
        full_text,
        r"\bPreventive procedures\b",
        r"\bMajor procedures\b",
        r"\bBenefit year maximum\b",
    )
    if not dental:
        dental_page, dental = sunlife_find_page(
            full_text,
            r"\bPreventive procedures\b",
            r"\bEmployee\s*Critical\s*Illness\b",
            r"\bLife\s+Employee\s+Life\b",
        )
    life_page, life = sunlife_find_page(
        full_text,
        r"\bLife\s*-?\s*Contract\s+(?:No\.?|Number)\b",
        r"\bEmployee Basic Life\b",
        r"\bMaximum\b",
    )
    if life:
        life = clean_text(f"{life} {sunlife_page_text(full_text, life_page + 1)}")
    if not life:
        life_page, life = dental_page, dental
    critical_illness_page, critical_illness = sunlife_find_page(
        full_text,
        r"\bCritical Illness\s*-?\s*Contract\s+(?:No\.?|Number)\b",
        r"\bEmployee Basic Critical Illness\b",
    )
    if not critical_illness:
        critical_illness_page, critical_illness = dental_page, dental
    ltd_page, ltd = sunlife_find_page(
        full_text,
        r"\bLong-Term Disability\s*-?\s*Contract\s+(?:No\.?|Number)\b",
        r"\bElimination period\b",
    )
    adnd_page, adnd = sunlife_find_page(
        full_text,
        r"\bAccidental Death and Dismemberment\b",
        r"\bEqual to Employee(?: Basic)? Life coverage\b",
    )
    eligibility_page, eligibility = sunlife_find_page(
        full_text,
        r"\bEligibility\b",
        r"\bpermanent employee working in Canada\b",
        r"\bhours a week\b",
    )
    virtual_page, virtual = sunlife_find_page(
        full_text,
        r"\bLumino Health Virtual\s*Care\b",
        r"\bDialogue\b",
    )
    recall_page, recall = sunlife_find_page(
        full_text,
        r"\bOral examinations\b",
        r"\brecall(?: or specific)? examination\b",
    )
    opinion_page, opinion = sunlife_find_page(
        full_text,
        r"\bTeladoc Medical Experts\b",
        r"\bCritical Illness\b",
    )

    waiting = sunlife_value(
        f"{summary} {eligibility}",
        [r"\bWaiting Period(?: General Information)?\s+(\d+\s*months? of continuous employment)"],
    )
    hours = sunlife_value(eligibility, [r"\bat least\s+(\d+ hours? a week)"])
    eligibility_description = sunlife_join(
        "Permanent employee working in Canada" if re.search(r"\bpermanent employee working in Canada\b", eligibility, flags=re.IGNORECASE) else "",
        hours,
    )
    add(
        "Eligibility Requirements",
        "Eligibility Requirements",
        eligibility_description,
        waiting_period=waiting,
        source_section="General Information - Eligibility",
        source_page=eligibility_page or summary_page,
    )

    life_schedule = sunlife_value(
        life,
        [
            r"\bEmployee Basic Life\s+Amount\s+(.{1,220}?)(?=\s+Maximum\b)",
            rf"\bLife\s+Employee\s+Life\s+Amount\s+({MONEY_AMOUNT_PATTERN})",
        ],
    )
    life_amount = sunlife_value(
        life,
        [
            rf"\bEmployee Basic Life.*?\bMaximum\s+[^$\d]{{0,12}}({MONEY_AMOUNT_PATTERN})",
            rf"\bLife\s+Employee\s+Life\s+Amount\s+({MONEY_AMOUNT_PATTERN})",
        ],
    )
    life_reduction = sunlife_value(
        life,
        [r"\bReduction\s+(Coverage is reduced to \d{1,3}% of the above amount when you reach age \d+)"],
    )
    life_termination = sunlife_value(
        life,
        [r"\bEmployee Basic Life.*?\bTermination\s+(When\s*you\s*retire\s*or\s*reach\s*age\s*\d+,?\s*whichever\s*is\s*earlier)"],
    )
    add(
        "Life / AD&D",
        "Life",
        sunlife_join(f"Benefit Amount: {life_schedule}" if life_schedule else "", life_reduction),
        maximum=life_amount,
        termination_age=life_termination,
        source_section="Benefit Summary - Life",
        source_page=life_page,
    )

    dependent_page, dependent_life = sunlife_find_page(
        full_text,
        r"\bBasic Dependent Life\b",
        r"\bSpouse\b",
        r"\bChild\b",
    )
    dependent_amount = sunlife_value(dependent_life or life, [r"\bBasic Dependent Life\s+Amount\s+(.{1,180}?)(?=\s+Termination\b)"])
    dependent_spouse = sunlife_value(dependent_amount, [rf"\bSpouse\s+[^$\d]{{0,8}}({MONEY_AMOUNT_PATTERN})"])
    dependent_child = sunlife_value(dependent_amount, [rf"\bChild\s+[^$\d]{{0,8}}({MONEY_AMOUNT_PATTERN})"])
    dependent_description = sunlife_join(
        f"Spouse: {dependent_spouse}" if dependent_spouse else "",
        f"Child: {dependent_child}" if dependent_child else "",
    ) or dependent_amount
    dependent_termination = sunlife_value(
        dependent_life or life,
        [r"\bBasic Dependent Life.*?\bTermination\s+(When\s*you\s*retire\s*or\s*reach\s*age\s*\d+,?\s*whichever\s*is\s*earlier)"],
    )
    add(
        "Dependent Life",
        "Dependent Life",
        dependent_description,
        termination_age=dependent_termination,
        source_section="Benefit Summary - Basic Dependent Life",
        source_page=dependent_page or life_page,
    )

    adnd_amount = sunlife_value(
        adnd,
        [r"\bEmployee(?: Basic)? Accidental Death and Dismemberment\s+Amount\s+(Equal to Employee(?: Basic)? Life coverage)"],
    )
    adnd_termination = sunlife_value(
        adnd,
        [r"\bEmployee(?: Basic)? Accidental Death and Dismemberment.*?\bTermination\s+(When you retire or reach age \d+, whichever is earlier)"],
    )
    add(
        "Life / AD&D",
        "AD&D",
        "Employee Accidental Death and Dismemberment",
        maximum=adnd_amount,
        termination_age=adnd_termination,
        source_section="Benefit Summary - Accidental Death and Dismemberment",
        source_page=adnd_page,
    )

    ci_amount = sunlife_value(critical_illness, [rf"\bEmployee(?:\s+Basic)?\s*Critical\s*Illness\s+Amount\s+({MONEY_AMOUNT_PATTERN})"])
    ci_termination = sunlife_value(
        critical_illness,
        [r"\bEmployee(?:\s+Basic)?\s*Critical\s*Illness.*?\bTermination\s+(When\s*you\s*retire\s*or\s*reach\s*age\s*\d+,?\s*whichever\s*is\s*earlier)"],
    )
    add(
        "Critical Illness",
        "Critical Illness",
        "Employee Critical Illness; Teladoc Medical Experts services included",
        maximum=ci_amount,
        termination_age=ci_termination,
        source_section="Benefit Summary - Critical Illness",
        source_page=critical_illness_page,
    )

    ltd_schedule = sunlife_value(
        ltd,
        [r"\bMaximum amount\s+(.{1,520}?)(?=\s+The maximum amount may be reduced|\s+Proof of good health\b)"],
    )
    ltd_maximum = sunlife_value(ltd_schedule, [rf"\bmaximum benefit of\s+({MONEY_AMOUNT_PATTERN})"])
    ltd_non_evidence = sunlife_value(
        ltd,
        [rf"\bProof of good health\s+Approval required for coverage in excess of\s+({MONEY_AMOUNT_PATTERN})"],
    )
    ltd_elimination = sunlife_value(ltd, [r"\bElimination period\s+(\d+\s+days?)"])
    ltd_benefit_period_raw = sunlife_value(
        ltd,
        [
            r"\bMaximum benefit period\s+(.{1,220}?)(?=\s+Benefits may also end|\s+Termination\b)",
            r"\bMaximum benefit\s+(.{1,220}?)\s+period\s+Benefits may also end",
        ],
    )
    ltd_age = re.search(r"\breach age\s+(\d+)\b", ltd_benefit_period_raw, flags=re.IGNORECASE)
    ltd_benefit_period = f"To age {ltd_age.group(1)}" if ltd_age else ltd_benefit_period_raw
    ltd_termination = sunlife_value(ltd, [r"\bTermination\s+(.{1,220}?)(?=\s+Tax status\b|$)"])
    ltd_tax_source = sunlife_value(
        ltd,
        [r"\bTax status\s+(.{1,700}?)(?=\s+Critical Illness\s*-?\s*Contract\b|$)"],
    )
    ltd_tax_status = normalize_ltd_tax_status(ltd_tax_source)
    if ltd_schedule:
        add(
            "Long-Term Disability",
            "LTD",
            sunlife_join(
                f"Benefit Schedule: {ltd_schedule}",
                f"Non-Evidence Maximum: {ltd_non_evidence}" if ltd_non_evidence else "",
                f"Tax Status: {ltd_tax_status}" if ltd_tax_status else "",
                f"Benefit Period: {ltd_benefit_period}" if ltd_benefit_period else "",
            ),
            maximum=f"{ltd_maximum} per month" if ltd_maximum else "",
            waiting_period=ltd_elimination,
            termination_age=ltd_termination,
            source_section="Benefit Summary - Long-Term Disability",
            source_page=ltd_page,
        )

    deductible = sunlife_value(summary, [r"\bDeductible\s+(None|No deductible|\$[\d,]+(?:\.\d{2})?)"])
    drug_coinsurance = sunlife_value(
        summary,
        [
            r"\bReimbursement level Prescription drugs\s+(\d{1,3}%)",
            r"\bPrescription drugs\s+(\d{1,3}%)",
        ],
    )
    drug_maximum = sunlife_value(
        summary,
        [rf"\bPrescription drugs\s+\d{{1,3}}%\s+(Maximum of {MONEY_AMOUNT_PATTERN} per person per benefit year)"],
    )
    if not drug_maximum and drug_coinsurance:
        drug_maximum = "Unlimited"
    drug_card = "Included" if re.search(r"\bDrug card plan Included\b", summary, flags=re.IGNORECASE) else ""
    substitution = sunlife_value(
        summary,
        [r"\bDrug substitution limit\s+(Charges in excess of the lowest priced equivalent drug are not covered unless specifically approved by Sun Life\.)"],
    )
    add(
        "Prescription Drugs",
        "Prescription Drugs",
        sunlife_join(f"Drug card plan {drug_card}" if drug_card else "", f"Deductible {deductible}" if deductible else "", substitution),
        coinsurance=drug_coinsurance,
        maximum=drug_maximum,
        frequency="Benefit year",
        source_section="Benefit Summary - Extended Health Care",
        source_page=summary_page,
    )

    medical_coinsurance = sunlife_value(
        continuation,
        [
            r"\bMedical services and equipment\s+(\d{1,3}%)",
            r"\bMedical services and\s+(\d{1,3}%)\s+equipment\b",
            r"\bAll other eligible expenses\s+(\d{1,3}%)",
        ],
    )
    add(
        "Medical Services & Supplies",
        "Medical Services & Supplies",
        "Medical services and equipment",
        coinsurance=medical_coinsurance,
        source_section="Benefit Summary - Extended Health Care",
        source_page=continuation_page,
    )

    paramedical_coinsurance = sunlife_value(continuation, [r"\bParamedical services\s+(\d{1,3}%)"])
    paramedical_maximum = sunlife_value(
        continuation,
        [
            rf"\bParamedical services\s+\d{{1,3}}%\s+(up to a maximum of {MONEY_AMOUNT_PATTERN} per person per specialty in a benefit year)",
            rf"\bParamedical services\s+\d{{1,3}}%\s+(up to a maximum of {MONEY_AMOUNT_PATTERN} per person per benefit year per specialty)",
        ],
    )
    practitioners = "Acupuncture, audiology, chiropractic, dietetics, massage therapy, naturopathy, occupational therapy, osteopathy, physiotherapy, podiatry/chiropody, speech therapy"
    add(
        "Paramedicals",
        "Paramedicals",
        practitioners,
        coinsurance=paramedical_coinsurance,
        maximum=paramedical_maximum,
        frequency="Benefit year",
        source_section="Benefit Summary - Paramedical Services",
        source_page=continuation_page,
    )
    mental_practitioners = "Psychologists, social workers, psychotherapists, and approved clinical counsellors"
    add(
        "Mental Health Practitioners",
        "Mental Health Practitioners",
        mental_practitioners,
        coinsurance=paramedical_coinsurance,
        maximum=paramedical_maximum,
        frequency="Benefit year",
        source_section="Benefit Summary - Paramedical Services",
        source_page=continuation_page,
    )

    health_pages = f"{summary} {continuation}"
    hospital_coinsurance = sunlife_value(health_pages, [r"\bIn-province hospital\s+(\d{1,3}%)"])
    hospital_description = sunlife_value(
        health_pages,
        [r"\bIn-province hospital\s+\d{1,3}%\s+(of the difference between the cost of a ward and a semi-private hospital room)"],
    )
    convalescent = sunlife_value(
        health_pages,
        [r"\bConvalescent hospital\s+\d{1,3}%\s+(of the difference between the cost of a ward and a semi-private room, up to \$[\d,]+ per day for a maximum of \d+ days)"],
    )
    add(
        "Hospital",
        "Hospital",
        sunlife_join(hospital_description, convalescent),
        coinsurance=hospital_coinsurance,
        source_section="Benefit Summary - Extended Health Care",
        source_page=continuation_page or summary_page,
    )

    travel_coinsurance = sunlife_value(
        continuation,
        [
            r"\bOut-of-province emergency services\s+(\d{1,3}%)",
            r"\bOut-of-province\s+(\d{1,3}%)\s+emergency services\b",
        ],
    )
    travel_duration = sunlife_value(
        continuation,
        [
            r"\bEmergency Travel Assistance included\s+(Maximum of \d+ days per trip)",
            r"\bTime limit\s+[^\d]{0,8}(\d+ days after the date the person leaves the province where the person lives)",
        ],
    )
    travel_amount = sunlife_value(
        continuation,
        [rf"\bLifetime\s*maximum\s*of\s*({MONEY_AMOUNT_PATTERN})\s*per\s*person\s*for\s*out-of-Canada\s*services"],
    )
    travel_maximum = (
        f"Lifetime maximum of {travel_amount} per person for out-of-Canada services"
        if travel_amount
        else ""
    )
    add(
        "Emergency Travel",
        "Emergency Travel",
        "Out-of-province emergency services; Emergency Travel Assistance included",
        coinsurance=travel_coinsurance,
        maximum=travel_maximum,
        frequency=travel_duration,
        source_section="Benefit Summary - Extended Health Care",
        source_page=continuation_page,
    )

    preventive = sunlife_value(dental, [r"\bPreventive procedures\s+(\d{1,3}%)"])
    basic = sunlife_value(dental, [r"\bBasic procedures\s+(\d{1,3}%)"])
    major = sunlife_value(dental, [r"\bMajor procedures\s+(\d{1,3}%)"])
    dental_maximum_amount = sunlife_value(
        dental,
        [rf"\bBenefit year maximum\s+({MONEY_AMOUNT_PATTERN} per person)"],
    )
    dental_maximum = (
        f"Combined Basic and Major maximum of {dental_maximum_amount} per benefit year"
        if dental_maximum_amount
        else ""
    )
    recall_frequency = sunlife_value(
        recall,
        [
            r"\b1 recall examination every (\d+ months)",
            r"\b1 recall or specific examination every (\d+ months)",
            r"\brecall examination once every (\d+ months)",
        ],
    )
    add(
        "Dental Care",
        "Dental Basic Services",
        basic or preventive,
        coinsurance=basic or preventive,
        maximum=dental_maximum,
        source_section="Benefit Summary - Dental Care",
        source_page=dental_page,
    )
    add(
        "Dental Care",
        "Dental Major Restorative Services",
        major,
        coinsurance=major,
        maximum=dental_maximum,
        source_section="Benefit Summary - Dental Care",
        source_page=dental_page,
    )
    add(
        "Dental Care",
        "Dental Recall",
        recall_frequency,
        frequency=recall_frequency,
        source_section="Dental Care - Oral Examinations",
        source_page=recall_page,
    )

    if re.search(
        r"\bLumino Health Virtual(?:\s*Care\s*services\s+Included|\s+Included\s+Care\s*services)\b",
        continuation,
        flags=re.IGNORECASE,
    ):
        provider = "Dialogue" if re.search(r"\bprovided by Dialogue\b", virtual, flags=re.IGNORECASE) else ""
        add(
            "Virtual Care",
            "Virtual Care",
            sunlife_join("Lumino Health Virtual Care included", f"Provided by {provider}" if provider else ""),
            source_section="Lumino Health Virtual Care",
            source_page=virtual_page or continuation_page,
        )

    if opinion:
        add(
            "Medical Second Opinion",
            "Medical Second Opinion",
            "Teladoc Medical Experts included with Critical Illness coverage",
            source_section="Critical Illness - Teladoc Medical Experts",
            source_page=opinion_page,
        )

    return items


def extract_sunlife_plan_design_rows(
    full_text: str,
    pdf_path: Path,
    client_info: ClientInfo,
    class_name: str,
) -> list[ExtractionRow]:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return [
        ExtractionRow(
            client_id=client_info.client_id,
            client_name=client_info.client_name,
            carrier=client_info.carrier or "Sun Life",
            class_name=class_name,
            benefit=item.benefit,
            coverage_description=item.coverage_description,
            coinsurance=item.coinsurance,
            maximum=item.maximum,
            frequency=item.frequency,
            waiting_period=item.waiting_period,
            termination_age=item.termination_age,
            notes="Sun Life Benefit Summary extraction",
            source_file=pdf_path.name,
            source_section=item.source_section,
            confidence="High",
            review_flag="",
            last_updated=now,
            policy_number=client_info.policy_number,
            source_page=item.source_page,
            benefit_category=item.summary_section,
        )
        for item in build_sunlife_plan_design_items(full_text)
    ]


def build_sunlife_summary_rows(full_text: str, client_info: ClientInfo) -> list[tuple[str, str]]:
    items = build_sunlife_plan_design_items(full_text)
    grouped: dict[str, list[SunLifePlanDesignItem]] = {}
    for item in items:
        grouped.setdefault(item.summary_section, []).append(item)

    policy = client_info.policy_number or extract_policy_number(full_text)
    carrier = client_info.carrier or "Sun Life"
    effective_date = sunlife_value(
        full_text,
        [
            r"\bBenefit Summary Effective\s+([A-Za-z]+ \d{1,2}, \d{4})",
            r"\bEffective\s*:?\s*([A-Za-z]+ \d{1,2}, \d{4})",
        ],
    )
    rows: list[tuple[str, str]] = [
        (
            "Carrier & Policy",
            sunlife_join(
                f"Carrier: {carrier}",
                f"Policy: {policy}" if policy else "",
                f"Effective Date: {effective_date}" if effective_date else "",
            ),
        )
    ]

    section_order = (
        "Eligibility Requirements",
        "Life / AD&D",
        "Dependent Life",
        "Long-Term Disability",
        "Critical Illness",
        "Prescription Drugs",
        "Medical Services & Supplies",
        "Paramedicals",
        "Mental Health Practitioners",
        "Hospital",
        "Emergency Travel",
        "Dental Care",
        "Virtual Care",
        "Medical Second Opinion",
    )
    for section in section_order:
        section_items = grouped.get(section, [])
        if not section_items:
            continue
        detail_lines: list[str] = []
        for item in section_items:
            prefix = f"{item.benefit}: " if len(section_items) > 1 else ""
            details = []
            if item.coverage_description:
                details.append(item.coverage_description)
            if item.coinsurance:
                details.append(f"Coinsurance: {item.coinsurance}")
            if item.maximum:
                details.append(f"Maximum: {item.maximum}")
            if item.frequency:
                details.append(f"Frequency: {item.frequency}")
            if item.waiting_period:
                details.append(f"Waiting Period: {item.waiting_period}")
            if item.termination_age:
                details.append(f"Termination: {item.termination_age}")
            if details:
                detail_lines.append(prefix + "; ".join(details))
        if detail_lines:
            rows.append((section, "\n".join(detail_lines)))
    return rows


def desjardins_page_block(
    full_text: str,
    heading_pattern: str,
    page_count: int = 1,
    required_pattern: str = "",
) -> tuple[int, str]:
    """Return a booklet table page and any directly following continuation pages."""
    pages = split_marked_pages(full_text)
    for index, (page_number, page_text) in enumerate(pages):
        cleaned_page = clean_text(page_text)
        # Desjardins repeats every benefit heading in its table of contents.  A
        # table-of-contents hit has no plan values and must never win over the
        # actual benefit table later in the booklet.
        if "TABLE OF CONTENTS" in cleaned_page.upper():
            continue
        if not re.search(heading_pattern, cleaned_page, flags=re.IGNORECASE):
            continue
        if required_pattern and not re.search(required_pattern, cleaned_page, flags=re.IGNORECASE):
            continue
        selected = " ".join(page[1] for page in pages[index : index + max(1, page_count)])
        return page_number, clean_text(selected)
    return 0, ""


def desjardins_text_after_label(text: str, label_pattern: str, length: int = 700) -> str:
    match = re.search(label_pattern, text, flags=re.IGNORECASE)
    if not match:
        return ""
    return clean_text(text[match.end() : match.end() + length])


def desjardins_join(*values: str) -> str:
    return "; ".join(unique_in_order([clean_text(value) for value in values if clean_text(value)]))


def build_desjardins_plan_design_items(full_text: str) -> list[DesjardinsPlanDesignItem]:
    """Extract the paired service/limit rows used in Desjardins EHC and dental booklets."""
    items: list[DesjardinsPlanDesignItem] = []

    def add(
        summary_section: str,
        benefit: str,
        description: str,
        coinsurance: str = "",
        maximum: str = "",
        frequency: str = "",
        source_section: str = "",
        source_page: int = 0,
    ) -> None:
        values = [description, coinsurance, maximum, frequency]
        if not any(clean_text(value) for value in values):
            return
        items.append(
            DesjardinsPlanDesignItem(
                summary_section=summary_section,
                benefit=benefit,
                coverage_description=clean_text(description),
                coinsurance=clean_text(coinsurance),
                maximum=clean_text(maximum),
                frequency=clean_text(frequency),
                source_section=clean_text(source_section),
                source_page=source_page,
            )
        )

    life_page, life_text = desjardins_page_block(full_text, r"\bLIFE BENEFIT\s+SUMMARY OF BENEFITS\b", required_pattern=r"\bBASIC LIFE BENEFIT\b")
    life_amount = first_regex(life_text, [rf"\bAmount of Insurance\s+({MONEY_AMOUNT_PATTERN})"])
    life_reduction = first_regex(life_text, [r"\b(The Amount of Insurance is reduced by \d{1,3}% on the Participant's \d+(?:st|nd|rd|th) birthday)"])
    add("Life / AD&D", "Life", desjardins_join("Basic Life Benefit", life_reduction), "100%", life_amount, source_section="Life Benefit", source_page=life_page)

    add_page, add_text = desjardins_page_block(full_text, r"\bACCIDENTAL DEATH AND DISMEMBERMENT BENEFIT\s+SUMMARY OF BENEFITS\b", required_pattern=r"\bBASIC ACCIDENTAL DEATH")
    add_amount = first_regex(add_text, [r"\bParticipant\s+Amount of Insurance\s+(Equal to the Basic Life Benefit)"])
    add("Life / AD&D", "AD&D", "Basic Accidental Death and Dismemberment Benefit; reduction equals Basic Life Benefit", "100%", add_amount, source_section="Accidental Death and Dismemberment Benefit", source_page=add_page)

    ehc_page, ehc_text = desjardins_page_block(full_text, r"\bEXTENDED HEALTH CARE BENEFIT\s+SUMMARY OF BENEFITS\b", required_pattern=r"\bDeductible\b")
    generic_drugs = first_regex(ehc_text, [r"\b(Generic drugs:\s*\d{1,3}%.*?market)(?=\s+2\))"])
    brand_drugs = first_regex(ehc_text, [r"\b(Brand name drugs:\s*\d{1,3}%.*?market)(?=\s+\*)"])
    costco = first_regex(ehc_text, [r"\b(For all Employees.*?increased by \d{1,3}%)"])
    drug_percent = first_regex(ehc_text, [r"\bGeneric drugs:\s*(\d{1,3}%)"])
    add("Prescription Drugs", "Drugs", desjardins_join(generic_drugs, brand_drugs, costco), drug_percent, source_section="Extended Health Care Benefit - Drugs", source_page=ehc_page)

    markup_page, markup_text = desjardins_page_block(full_text, r"\bMARK-UP AND DISPENSING FEE\b", required_pattern=r"\bDispensing fee\b")
    markup = first_regex(markup_text, [r"\bMark-up\s+(Reasonable and Customary Charges)"])
    dispensing = first_regex(markup_text, [r"\bDispensing fee\s+(Reasonable and Customary Charges)"])
    add(
        "Prescription Drugs",
        "Drug Mark-up and Dispensing Fee",
        "Carrier limits for eligible drug expenses",
        "80%",
        desjardins_join(markup, dispensing),
        source_section="Eligible Expenses in Canada - Mark-up and Dispensing Fee",
        source_page=markup_page,
    )

    payment_page, payment_text = desjardins_page_block(full_text, r"\bThis provision applies to Quebec residents only\b")
    for benefit, label, section in (
        ("Hospitalization", "Hospitalization", "Hospitalization"),
        ("Dental Treatment Due to an Accident", "Dental Treatment due to an accident", "Dental Treatment Due to an Accident"),
        ("Vision Care - Health Benefit", "Vision Care", "Vision Care"),
        ("Travel Insurance - Health Benefit", "Travel Insurance", "Travel Insurance"),
        ("Extended Health Care - Other Expenses", "All other expenses", "Extended Health Care Benefit"),
    ):
        percentage = first_regex(payment_text, [rf"\b{re.escape(label)}\s+(\d{{1,3}}%)"])
        add("Medical Services & Supplies", benefit, f"{label} reimbursement under Extended Health Care", percentage, source_section=section, source_page=payment_page)

    fertility_page, fertility_text = desjardins_page_block(full_text, r"\bFertility Treatment and Surrogacy\b", page_count=2)
    fertility_max = first_regex(fertility_text, [rf"\bFertility treatment.*?({MONEY_AMOUNT_PATTERN}\s+combined lifetime(?:, including genetic testing)?)"])
    surrogacy_max = first_regex(fertility_text, [rf"\bSurrogacy.*?({MONEY_AMOUNT_PATTERN}\s+combined lifetime(?:, including genetic testing, prenatal screening tests, ultrasounds and laboratory tests related to the pregnancy)?)"])
    add("Prescription Drugs", "Fertility Treatment", "Eligible drugs and treatment, including genetic testing", maximum=fertility_max, source_section="Fertility Treatment", source_page=fertility_page)
    add("Prescription Drugs", "Surrogacy", "Eligible drugs and treatment for the surrogate pregnancy", maximum=surrogacy_max, source_section="Surrogacy", source_page=fertility_page)
    vaccines = first_regex(fertility_text, [r"\bVaccines\s+(Reasonable and Customary Charges)"])
    sclerotherapy = first_regex(fertility_text, [rf"\bSclerotherapy products.*?({MONEY_AMOUNT_PATTERN}\s+per visit)"])
    smoking = first_regex(fertility_text, [rf"\bSmoking cessation aids.*?({MONEY_AMOUNT_PATTERN}\s+lifetime)"])
    drug_overall = first_regex(fertility_text, [rf"\bOverall Maximum Benefit.*?({MONEY_AMOUNT_PATTERN}\s+per calendar year)"])
    add("Prescription Drugs", "Vaccines", "Eligible vaccine expenses", maximum=vaccines, source_section="Other Eligible Drug Expenses", source_page=fertility_page)
    add("Prescription Drugs", "Sclerotherapy", "Sclerotherapy products to treat varicose veins", maximum=sclerotherapy, source_section="Other Eligible Drug Expenses", source_page=fertility_page)
    add("Prescription Drugs", "Smoking Cessation", "Smoking cessation aids (products only)", maximum=smoking, source_section="Other Eligible Drug Expenses", source_page=fertility_page)
    add("Prescription Drugs", "Other Eligible Drug Expenses", "Overall eligible drug expense maximum", maximum=drug_overall, source_section="Other Eligible Drug Expenses", source_page=fertility_page)

    hospital_page, hospital_text = desjardins_page_block(full_text, r"\bHOSPITALIZATION\b", required_pattern=r"\bPalliative Care Establishment\b")
    ward = first_regex(hospital_text, [r"\bHospital\s+(The difference between the cost of a ward and a semi-private room)"])
    palliative_amount = first_regex(hospital_text, [rf"({MONEY_AMOUNT_PATTERN}\s+per day and a lifetime maximum of)\s+Care Establishment\s+(\d+ days)"])
    palliative_days = first_regex(hospital_text, [r"\blifetime maximum of\s+Care Establishment\s+(\d+ days)"])
    palliative = desjardins_join(palliative_amount, palliative_days).replace("; ", " ")
    rehabilitation_amount = first_regex(hospital_text, [rf"({MONEY_AMOUNT_PATTERN}\s+per day and a maximum of).*?\b\d+ days per period of confinement"])
    rehabilitation_days = first_regex(hospital_text, [r"\b(\d+ days per period of confinement)"])
    rehabilitation = desjardins_join(rehabilitation_amount, rehabilitation_days).replace("; ", " ")
    add("Hospital", "Hospitalization", "Acute care hospital confinement", "100%", ward, source_section="Hospitalization", source_page=hospital_page)
    add("Hospital", "Palliative Care", "Palliative care establishment confinement", "100%", palliative, source_section="Hospitalization", source_page=hospital_page)
    add("Hospital", "Convalescent/Rehabilitation Centre", "Eligible confinement following hospital discharge", "100%", rehabilitation, source_section="Hospitalization", source_page=hospital_page)

    paramedical_page, paramedical_text = desjardins_page_block(full_text, r"\bHEALTH PROFESSIONALS\b", required_pattern=r"\bParamedical Services\b")
    paramedical_max = first_regex(paramedical_text, [rf"({MONEY_AMOUNT_PATTERN}\s+per calendar year for each type of professional(?:\s+plus\s+{MONEY_AMOUNT_PATTERN}\s+for x-rays)?)"])
    one_visit = first_regex(paramedical_text, [r"\b(For each type of professional, only one visit per day is eligible)"])
    add("Paramedicals", "Paramedical Services", desjardins_join("Eligible professionals include acupuncture, chiropractic, massage therapy, physiotherapy, psychology/counselling, speech therapy, and other listed professionals", one_visit), "80%", paramedical_max, source_section="Health Professionals - Paramedical Services", source_page=paramedical_page)

    home_page, home_text = desjardins_page_block(full_text, r"\bHome Nursing Care\b", required_pattern=r"\bGENDER AFFIRMATION\b")
    home_nursing = first_regex(home_text, [rf"\bHome Nursing Care.*?({MONEY_AMOUNT_PATTERN}\s+per calendar year)"])
    gender_max = first_regex(home_text, [rf"\bGENDER AFFIRMATION.*?({MONEY_AMOUNT_PATTERN}\s+lifetime)"])
    add("Medical Services & Supplies", "Home Nursing Care", "Registered nurse or licensed practical nurse services provided at home", "80%", home_nursing, source_section="Home Nursing Care", source_page=home_page)
    add("Medical Services & Supplies", "Gender Affirmation", "Eligible surgeries and treatments not covered under the provincial medical plan", "80%", gender_max, source_section="Gender Affirmation", source_page=home_page)

    mobility_page, mobility_text = desjardins_page_block(full_text, r"\bMOBILITY AIDS\b", required_pattern=r"\bWheelchairs\b")
    wheelchair_frequency = first_regex(mobility_text, [r"\b(One in any \d+-month period)"])
    wheelchair_max = first_regex(mobility_text, [rf"\bOne in any \d+-month period, up to a lifetime maximum of ({MONEY_AMOUNT_PATTERN})"])
    walkers = first_regex(mobility_text, [r"\bWalkers, canes or crutches.*?(Reasonable and Customary Charges)"])
    add("Medical Services & Supplies", "Walkers, Canes or Crutches", "Purchase or rental at Desjardins' option", "80%", walkers, source_section="Medical Equipment or Supplies - Mobility Aids", source_page=mobility_page)
    add("Medical Services & Supplies", "Wheelchairs", "Purchase, repair or rental at Desjardins' option", "80%", wheelchair_max, wheelchair_frequency, "Medical Equipment or Supplies - Mobility Aids", mobility_page)

    orthopaedic_page, orthopaedic_text = desjardins_page_block(full_text, r"\bORTHOPAEDIC SUPPLIES\b", required_pattern=r"\bFoot orthoses\b")
    shoes_text = orthopaedic_text.split("Foot orthoses", 1)[0]
    shoes_adult = first_regex(
        shoes_text,
        [
            rf"\bOne pair per calendar year, up to\s+prefabricated shoes\s+({MONEY_AMOUNT_PATTERN})\s+for adults",
            rf"\bOne pair per calendar year, up to\s+({MONEY_AMOUNT_PATTERN})\s+for adults",
        ],
    )
    shoes_child = first_regex(
        shoes_text,
        [
            rf"\b2 pairs per calendar year, up to\s+prefabricated shoes\s+({MONEY_AMOUNT_PATTERN})\s+for Children under age 18",
            rf"\b2 pairs per calendar year, up to\s+({MONEY_AMOUNT_PATTERN})\s+for Children under age 18",
        ],
    )
    shoes = desjardins_join(
        f"One pair per calendar year, up to {shoes_adult} for adults" if shoes_adult else "",
        f"2 pairs per calendar year, up to {shoes_child} for Children under age 18" if shoes_child else "",
    )
    foot_block = desjardins_text_after_label(orthopaedic_text, r"\bFoot orthoses\b")
    foot_orthoses = desjardins_join(
        first_regex(foot_block, [rf"\b(One pair per calendar year, up to {MONEY_AMOUNT_PATTERN} for adults)"]),
        first_regex(foot_block, [rf"\b(2 pairs per calendar year, up to {MONEY_AMOUNT_PATTERN} for Children under age 18)"]),
    )
    add("Medical Services & Supplies", "Orthopaedic Shoes", "Custom-made, open-toed, modified or adjusted shoes and qualifying Denis Browne brace shoes", "80%", shoes, source_section="Medical Equipment or Supplies - Orthopaedic Supplies", source_page=orthopaedic_page)
    add("Medical Services & Supplies", "Foot Orthoses", "Manufactured and billed by a centre recognized by Desjardins", "80%", foot_orthoses, source_section="Medical Equipment or Supplies - Orthopaedic Supplies", source_page=orthopaedic_page)

    prostheses_page, prostheses_text = desjardins_page_block(full_text, r"\bPROSTHESES\b", required_pattern=r"\bHearing aids\b")
    hearing = first_regex(prostheses_text, [rf"({MONEY_AMOUNT_PATTERN}\s+in any \d+-month period, including)\s+Hearing aids\s+initial batteries"])
    if hearing:
        hearing = f"{hearing} initial batteries"
    wigs = first_regex(prostheses_text, [rf"\bWigs.*?({MONEY_AMOUNT_PATTERN}\s+lifetime)"])
    breast = first_regex(prostheses_text, [rf"\bBreast prostheses.*?({MONEY_AMOUNT_PATTERN}\s+in any \d+-month period)"])
    limb_limits = desjardins_join(
        *unique_regex_group_matches(
            prostheses_text,
            [rf"({MONEY_AMOUNT_PATTERN}\s+per (?:prosthesis|repair))"],
        )
    )
    artificial_eyes = first_regex(prostheses_text, [r"\bArtificial eyes\s+Purchase and repair\s+(Reasonable and Customary Charges)"])
    add("Medical Services & Supplies", "Hearing Aids", "Including initial batteries", "80%", hearing, source_section="Medical Equipment or Supplies - Prostheses", source_page=prostheses_page)
    add("Medical Services & Supplies", "Wigs", "When required for alopecia or temporary hair loss due to chemotherapy or radiotherapy", "80%", wigs, source_section="Medical Equipment or Supplies - Prostheses", source_page=prostheses_page)
    add("Medical Services & Supplies", "Breast Prostheses", "External prostheses and qualifying mastectomy brassieres", "80%", breast, source_section="Medical Equipment or Supplies - Prostheses", source_page=prostheses_page)
    add("Medical Services & Supplies", "Artificial Limbs and Myoelectric Prosthetics", "Purchase, repair and replacement as specified", "80%", limb_limits, source_section="Medical Equipment or Supplies - Prostheses", source_page=prostheses_page)
    add("Medical Services & Supplies", "Artificial Eyes", "Purchase and repair", "80%", artificial_eyes, source_section="Medical Equipment or Supplies - Prostheses", source_page=prostheses_page)

    mental_page, mental_text = desjardins_page_block(full_text, r"\bMENTAL HEALTH AND SUBSTANCE ABUSE\b", required_pattern=r"\bsubstance abuse\b")
    mental_max = first_regex(mental_text, [rf"\b(?:substance abuse|behavioral addictions|eating disorders).*?({MONEY_AMOUNT_PATTERN}\s+per day, up to {MONEY_AMOUNT_PATTERN}\s+lifetime)"])
    add("Paramedicals", "Mental Health and Substance Abuse", "Eligible treatment in a recognized centre for substance abuse, behavioral addictions, or eating disorders", "80%", mental_max, source_section="Mental Health and Substance Abuse", source_page=mental_page)

    vision_page, vision_text = desjardins_page_block(full_text, r"\bVISION CARE\b", required_pattern=r"\bEye exam\b")
    eye_exam = first_regex(vision_text, [r"\b(One in any period of .*?age 18\.)"])
    intraocular = first_regex(vision_text, [rf"\bIntraocular lenses.*?({MONEY_AMOUNT_PATTERN}\s+per calendar year)"])
    add("Vision Care", "Vision Care - Eye Exam", "Reasonable and Customary Charges", "100%", frequency=eye_exam, source_section="Vision Care", source_page=vision_page)
    add("Vision Care", "Vision Care - Intraocular Lenses", "Replacement for natural crystalline lenses in case of cataracts", "100%", intraocular, source_section="Vision Care", source_page=vision_page)

    referral_page, referral_text = desjardins_page_block(full_text, r"\bREFERRAL TREATMENT\b", page_count=3)
    referral_outside_canada = first_regex(
        referral_text,
        [
            rf"\bExpenses incurred outside Canada\s+({MONEY_AMOUNT_PATTERN}\s+per calendar year per Covered Person)",
            rf"({MONEY_AMOUNT_PATTERN}\s+per calendar year per)\s+Expenses incurred outside Canada\s+Covered Person",
        ],
    )
    if referral_outside_canada.endswith("per"):
        referral_outside_canada = f"{referral_outside_canada} Covered Person"
    referral_canada = first_regex(referral_text, [r"\b(In Canada: same coverage as provided for under the In Canada provision of this Benefit)"])
    add("Referral Treatment", "Referral Treatment - In Canada", "Referred treatment outside the province of residence but within Canada", "100%", referral_canada, source_section="Referral Treatment", source_page=referral_page)
    add("Referral Treatment", "Referral Treatment - Outside Canada", "Referred treatment outside Canada; prior approval and referral conditions apply", "100%", referral_outside_canada, source_section="Referral Treatment", source_page=referral_page)

    travel_page, travel_text = desjardins_page_block(full_text, r"\bTRAVEL INSURANCE\b", page_count=5, required_pattern=r"\bMedical Emergency\b")
    travel_duration = first_regex(travel_text, [r"\b(first \d+ days of a stay outside their province of residence)"])
    travel_max = first_regex(travel_text, [rf"\bAll Eligible Expenses\s+({MONEY_AMOUNT_PATTERN}\s+lifetime per Covered Person)"])
    add("Travel Insurance", "Travel Insurance", "Medical emergency expenses while outside the province of residence", "100%", travel_max, travel_duration, "Travel Insurance", travel_page)

    dental_page, dental_text = desjardins_page_block(full_text, r"\bDENTAL CARE BENEFIT\s+SUMMARY OF BENEFITS\b", required_pattern=r"\bMajor Restorative Services\b")
    preventive = first_regex(dental_text, [r"\bPreventive Services\s+(\d{1,3}%)"])
    basic = first_regex(dental_text, [r"\bBasic Services\s+(\d{1,3}%)"])
    major = first_regex(dental_text, [r"\bMajor Restorative Services\s+(\d{1,3}%)"])
    dental_max_amount = first_regex(
        dental_text,
        [
            rf"\bPreventive, Basic and Major\s+Combined maximum of ({MONEY_AMOUNT_PATTERN})\s+per",
            rf"\bPreventive, Basic and Major Restorative Services\s+Combined maximum of ({MONEY_AMOUNT_PATTERN})\s+per",
        ],
    )
    dental_max = f"Combined maximum of {dental_max_amount} per calendar year per Covered Person" if dental_max_amount else ""
    dental_coinsurance = desjardins_join(f"Preventive {preventive}" if preventive else "", f"Basic {basic}" if basic else "", f"Major Restorative {major}" if major else "")
    add("Dental Care", "Dental Care", "Preventive, Basic and Major Restorative Services", dental_coinsurance, dental_max, source_section="Dental Care Benefit - Summary of Benefits", source_page=dental_page)
    add("Dental Care", "Dental Basic Services", "Restorations, endodontics, periodontics and oral surgery", basic, dental_max, source_section="Dental Care Benefit - Basic Services", source_page=dental_page)
    add("Dental Care", "Dental Major Restorative Services", "Dentures, bridgework, crowns and other listed restorations", major, dental_max, source_section="Dental Care Benefit - Major Restorative Services", source_page=dental_page)

    dental_recall_page, dental_recall_text = desjardins_page_block(full_text, r"\bPREVENTIVE SERVICES\b", required_pattern=r"\bPreventive or recall oral\b")
    recall_frequency = first_regex(dental_recall_text, [r"\bPreventive or recall oral\s+(One in any \d+-month period)\s+examination"])
    add("Dental Care", "Dental Recall", "Preventive or recall oral examination", preventive, frequency=recall_frequency, source_section="Dental Care Benefit - Preventive Services", source_page=dental_recall_page)

    return items


def extract_desjardins_plan_design_rows(
    full_text: str,
    pdf_path: Path,
    client_info: ClientInfo,
    class_name: str,
) -> list[ExtractionRow]:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return [
        ExtractionRow(
            client_id=client_info.client_id,
            client_name=client_info.client_name,
            carrier=client_info.carrier,
            class_name=class_name,
            benefit=item.benefit,
            coverage_description=item.coverage_description,
            coinsurance=item.coinsurance,
            maximum=item.maximum,
            frequency=item.frequency,
            waiting_period="",
            termination_age="",
            notes="Desjardins booklet table extraction",
            source_file=pdf_path.name,
            source_section=item.source_section,
            confidence="High",
            review_flag="",
            last_updated=now,
            policy_number=client_info.policy_number,
            source_page=item.source_page,
        )
        for item in build_desjardins_plan_design_items(full_text)
    ]


def build_desjardins_summary_rows(full_text: str, client_info: ClientInfo) -> list[tuple[str, str]]:
    items = build_desjardins_plan_design_items(full_text)
    rows: list[tuple[str, str]] = []
    policy = client_info.policy_number or extract_policy_number(full_text)
    rows.append(("Carrier & Policy", desjardins_join(client_info.carrier or "Desjardins", f"Policy No. {policy}" if policy else "")))

    # The PlanDesign sheet and SQLite retain every extracted table row.  This
    # summary intentionally keeps only the limits and plan-level facts a
    # consultant needs in a client-facing benefits overview.
    by_benefit = {item.benefit: item for item in items}

    def item(benefit: str) -> DesjardinsPlanDesignItem | None:
        return by_benefit.get(benefit)

    def value(benefit: str, field_name: str) -> str:
        source = item(benefit)
        return clean_text(getattr(source, field_name, "")) if source else ""

    def lines(*entries: tuple[str, str]) -> str:
        return "\n".join(f"{label}: {detail}" for label, detail in entries if clean_text(detail))

    life = item("Life")
    life_reduction = first_regex(
        life.coverage_description if life else "",
        [r"\breduced by\s+(\d{1,3}%\s+on the Participant's\s+\d+(?:st|nd|rd|th) birthday)"],
    )
    life_details = lines(
        ("Life Benefit Amount", value("Life", "maximum")),
        ("Life Reduction", life_reduction),
        ("AD&D Benefit Amount", value("AD&D", "maximum")),
    )
    if life_details:
        rows.append(("Life / AD&D", life_details))

    fertility_limit = desjardins_join(
        first_regex(value("Fertility Treatment", "maximum"), [rf"({MONEY_AMOUNT_PATTERN}\s+combined lifetime)"]),
        first_regex(value("Surrogacy", "maximum"), [rf"({MONEY_AMOUNT_PATTERN}\s+combined lifetime)"]),
    )
    drug_details = lines(
        ("Coinsurance", value("Drugs", "coinsurance")),
        ("Mark-up and Dispensing Fees", value("Drug Mark-up and Dispensing Fee", "maximum")),
        ("Fertility and Surrogacy", fertility_limit),
        ("Other Eligible Drug Expenses", value("Other Eligible Drug Expenses", "maximum")),
    )
    if drug_details:
        rows.append(("Prescription Drugs", drug_details))

    medical_details = lines(
        ("Coinsurance", value("Extended Health Care - Other Expenses", "coinsurance")),
        ("Home Nursing Care", value("Home Nursing Care", "maximum")),
        ("Gender Affirmation", value("Gender Affirmation", "maximum")),
        ("Medical Equipment and Supplies", "Detailed limits are available in Plan Design" if any(item(name) for name in ("Wheelchairs", "Orthopaedic Shoes", "Hearing Aids")) else ""),
    )
    if medical_details:
        rows.append(("Medical Services & Supplies", medical_details))

    hospital_details = lines(
        ("Hospital Coverage", value("Hospitalization", "coverage_description")),
        ("Palliative Care", value("Palliative Care", "maximum")),
        ("Rehabilitation", value("Convalescent/Rehabilitation Centre", "maximum")),
    )
    if hospital_details:
        rows.append(("Hospital", hospital_details))

    paramedical_frequency = first_regex(
        value("Paramedical Services", "coverage_description"),
        [r"\b(For each type of professional, only one visit per day is eligible)"],
    )
    paramedical_details = lines(
        ("Benefit Amount", value("Paramedical Services", "maximum")),
        ("Frequency", paramedical_frequency),
        ("Mental Health and Substance Abuse", value("Mental Health and Substance Abuse", "maximum")),
    )
    if paramedical_details:
        rows.append(("Paramedicals", paramedical_details))

    eye_exam_frequency = value("Vision Care - Eye Exam", "frequency")
    adult_eye_exam_frequency = first_regex(eye_exam_frequency, [r"\b(\d+ months)\s+for adults"])
    child_eye_exam_frequency = first_regex(eye_exam_frequency, [r"\b(\d+ months)\s+for Children under age 18"])
    normalized_eye_exam_frequency = desjardins_join(
        f"Adults: {adult_eye_exam_frequency}" if adult_eye_exam_frequency else "",
        f"Children under age 18: {child_eye_exam_frequency}" if child_eye_exam_frequency else "",
    ) or eye_exam_frequency
    vision_details = lines(
        ("Eye Exam", normalized_eye_exam_frequency),
        ("Intraocular Lenses", value("Vision Care - Intraocular Lenses", "maximum")),
    )
    if vision_details:
        rows.append(("Vision Care", vision_details))

    referral_details = lines(
        ("Outside Canada", value("Referral Treatment - Outside Canada", "maximum")),
    )
    if referral_details:
        rows.append(("Referral Treatment", referral_details))

    travel_details = lines(
        ("Trip Duration", value("Travel Insurance", "frequency")),
        ("Maximum", value("Travel Insurance", "maximum")),
    )
    if travel_details:
        rows.append(("Travel Insurance", travel_details))

    dental_details = lines(
        ("Coinsurance", value("Dental Care", "coinsurance")),
        ("Maximum", value("Dental Care", "maximum")),
        ("Recall Exam", value("Dental Recall", "frequency")),
    )
    if dental_details:
        rows.append(("Dental Care", dental_details))

    return rows


def extract_equitable_schedule_text(full_text: str) -> str:
    cleaned = clean_text(full_text)
    start = re.search(
        r"\bSchedule of benefits\s+The Plan described in this booklet\b",
        cleaned,
        flags=re.IGNORECASE,
    )
    if not start:
        matches = list(re.finditer(r"\bSchedule of benefits\b", cleaned, flags=re.IGNORECASE))
        start = matches[-1] if matches else None
    if not start:
        return ""

    schedule_text = cleaned[start.start() :]
    end = re.search(
        r"\bNote:\s+The following pages are standard descriptive pages\b|\bHomewood Health employee and family assistance program\b",
        schedule_text,
        flags=re.IGNORECASE,
    )
    if end:
        schedule_text = schedule_text[: end.start()]
    return schedule_text


def extract_equitable_class_names(full_text: str) -> list[str]:
    schedule_text = extract_equitable_schedule_text(full_text)
    class_block = first_regex(
        schedule_text,
        [r"\bClassification\(s\)\s+(.{1,900}?)(?=\s+THE EQUITABLE|\s+General information\b|\s+Maximum age\b|$)"],
    )
    classes: list[str] = []
    for match in re.finditer(
        r"\bClass\s+([A-Za-z0-9]+)\s*:\s*(.{1,220}?)(?=\s+Class\s+[A-Za-z0-9]+\s*:|$)",
        class_block,
        flags=re.IGNORECASE,
    ):
        class_code = match.group(1).upper()
        description = clean_text(match.group(2)).strip(" .•*-")
        classes.append(f"Class {class_code} - {description}" if description else f"Class {class_code}")

    if classes:
        return unique_in_order(classes)

    header = clean_text(full_text[:2500])
    class_codes = first_regex(header, [r"\bCLASSES?\s+([A-Za-z0-9,\sand]+)"])
    codes = re.findall(r"\b[A-Za-z]\b", class_codes)
    return [f"Class {code.upper()}" for code in unique_in_order(codes)]


def extract_equitable_schedule_plan_design_rows(
    full_text: str,
    pdf_path: Path,
    client_info: ClientInfo,
    class_name: str,
) -> list[ExtractionRow]:
    rows: list[ExtractionRow] = []
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for section_name, detail_text in build_equitable_summary_rows(full_text):
        if section_name in {"Carrier & Policy", "Eligibility Requirements"}:
            continue

        for detail in schedule_detail_lines(detail_text):
            if not detail or detail.startswith("- "):
                continue

            rows.append(
                ExtractionRow(
                    client_id=client_info.client_id,
                    client_name=client_info.client_name,
                    carrier=client_info.carrier,
                    class_name=class_name,
                    benefit=equitable_plan_design_benefit_name(section_name),
                    coverage_description=detail,
                    coinsurance="" if section_name == "Life / AD&D" else extract_coinsurance(detail),
                    maximum=extract_maximum(detail),
                    frequency=extract_frequency(detail),
                    waiting_period="",
                    termination_age=extract_termination_age(detail),
                    notes="Equitable Schedule of Benefits",
                    source_file=pdf_path.name,
                    source_section=section_name,
                    confidence="High",
                    review_flag="",
                    last_updated=now,
                )
            )
    return rows


def build_equitable_summary_rows(full_text: str) -> list[tuple[str, str]]:
    schedule_text = extract_equitable_schedule_text(full_text)
    if not schedule_text:
        return []

    builders = [
        ("Carrier & Policy", build_equitable_carrier_policy_lines),
        ("Eligibility Requirements", build_equitable_eligibility_lines),
        ("Life / AD&D", build_equitable_life_adnd_lines),
        ("Long-Term Disability", build_equitable_ltd_lines),
        ("Prescription Drugs", build_equitable_prescription_drug_lines),
        ("Medical Services & Supplies", build_equitable_ehc_lines),
        ("Paramedicals", build_equitable_paramedical_lines),
        ("Hospital", build_equitable_hospital_lines),
        ("Vision Care", build_equitable_vision_lines),
        ("Emergency Travel", build_equitable_travel_lines),
        ("Dental Care", build_equitable_dental_lines),
        ("Survivor Benefit", build_equitable_survivor_lines),
        ("Health Care Spending Account", build_equitable_hcsa_lines),
        ("Lifestyle Spending Account", build_equitable_lsa_lines),
        ("Virtual Care", build_equitable_virtual_care_lines),
        ("Employee Assistance Program", build_equitable_eap_lines),
    ]

    rows: list[tuple[str, str]] = []
    for section_name, builder in builders:
        detail_lines = compact_detail_lines(builder(full_text, schedule_text))
        if detail_lines:
            rows.append((section_name, "\n".join(detail_lines)))
    return rows


def equitable_block(schedule_text: str, start_patterns: list[str], end_patterns: list[str]) -> str:
    start_match = None
    for pattern in start_patterns:
        match = re.search(pattern, schedule_text, flags=re.IGNORECASE)
        if match and (start_match is None or match.start() < start_match.start()):
            start_match = match
    if not start_match:
        return ""

    block = schedule_text[start_match.start() :]
    end_positions = []
    for pattern in end_patterns:
        match = re.search(pattern, block[start_match.end() - start_match.start() :], flags=re.IGNORECASE)
        if match:
            end_positions.append(start_match.end() - start_match.start() + match.start())
    if end_positions:
        block = block[: min(end_positions)]
    return block


def build_equitable_carrier_policy_lines(full_text: str, _schedule_text: str) -> list[str]:
    policy = first_regex(full_text[:6000], [r"\bPOLICY#\s*([A-Za-z0-9\-]+)", r"\bPolicy\s+#\s*([A-Za-z0-9\-]+)"])
    return [f"Equitable: {policy}" if policy else "Equitable"]


def build_equitable_eligibility_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = equitable_block(schedule_text, [r"\bGeneral information\b"], [r"\bEmployee life insurance\b"])
    waiting = first_regex(
        block,
        [
            r"\bWaiting period:\s*(?:\([^)]*\)\s*)?[•\-\s]*(\d+\s+months?)",
            r"\bWaiting period:\s*(?:\([^)]*\)\s*)?(.{1,80}?)(?=\s+Definition of earnings\b)",
        ],
    )
    hours = first_regex(
        block,
        [r"\bMinimum number of hours per week employees must work to be eligible for coverage:\s*(\d+\s+hours\s+per\s+week)"],
    )
    termination = first_regex(block, [r"\bAll other benefits terminate on your\s+(.{1,60}?)(?=\s+Minimum number\b)"])
    return [
        f"Waiting Period: {waiting}",
        f"Hours Worked: {hours}",
        f"General Termination: {termination}",
    ]


def build_equitable_life_adnd_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = equitable_block(schedule_text, [r"\bEmployee life insurance\b"], [r"\bEmployee long term disability\b"])
    amount = first_regex(
        block,
        [
            rf"\bA\s+flat\s+amount\s+of\s+({MONEY_AMOUNT_PATTERN})",
            r"\b(1\s+times\s+annual\s+earnings.*?maximum\s+benefit\s+of\s+\$\s?\d[\d,]*)",
        ],
    )
    no_evidence = first_regex(block, [r"\bNo-evidence limit:\s+(.{1,120}?)(?=\s+Reduction:)"])
    reduction = first_regex(block, [r"\bReduction:\s+(.{1,160}?)(?=\s+AD&D waiver\b)"])
    adnd_waiver = first_regex(
        block,
        [r"\bAD&D waiver of premium:\s+(.{1,180}?)(?=\s+Employee (?:optional life|long term disability|and dependent (?:health|dental) benefits)\b|$)"],
    )
    optional_life = first_regex(block, [rf"\bEmployee optional life / Spousal optional life insurance\s+(.{1,160}?{MONEY_AMOUNT_PATTERN})"])
    return [
        f"Benefit Amount: {amount}",
        f"No-Evidence Limit: {no_evidence}",
        f"Benefit Reduction: {reduction}",
        f"Waiver of Premium: {adnd_waiver}",
        f"Optional Life: {optional_life}",
    ]


def build_equitable_ltd_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = equitable_block(schedule_text, [r"\bEmployee long term disability"], [r"\bEmployee and dependent health benefits\b"])
    benefit = first_regex(
        block,
        [r"\b(66\.7%\s+of\s+your\s+monthly\s+earnings.*?maximum\s+benefit\s+of\s+\$\s?\d[\d,]*\s+per\s+month)"],
    )
    no_evidence = first_regex(block, [r"\bNo-evidence limit:\s+(.{1,140}?\$\s?\d[\d,]*)"])
    all_source = first_regex(block, [r"\bAll source maximum:\s+(.{1,260}?)(?=\s+Elimination period:)"])
    elimination = first_regex(block, [r"\bElimination period:.{0,160}?\s+(\d+\s+consecutive\s+days\s+of\s+disability)"])
    commencement = first_regex(block, [r"\bBenefit commencement date:\s+(.{1,80}?)(?=\s+Maximum benefit period:)"])
    benefit_period = first_regex(block, [r"\bMaximum benefit period:\s+(.{1,80}?)(?=\s+Partial disability:)"])
    tax_status = first_regex(block, [r"\bTax status:\s+(.{1,260}?)(?=$)"])
    return [
        f"Benefit Schedule: {benefit}",
        f"Non-Evidence Limit: {no_evidence}",
        f"All Source Maximum: {all_source}",
        f"Elimination Period: {elimination}",
        f"Benefit Commencement Date: {commencement}",
        f"Benefit Period: {benefit_period}",
        f"Tax Status: {tax_status}",
    ]


def build_equitable_prescription_drug_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = equitable_block(schedule_text, [r"\bEmployee and dependent health benefits\b"], [r"\bSummary of health benefit maximums\b"])
    deductible = first_regex(block, [r"\bDeductible amount per prescription for the drug plan:\s+(.{1,40}?)(?=\s+Deductible amount per calendar year)"])
    plan_type = first_regex(block, [r"\b(Pay-direct drug plan\s+#?[A-Za-z0-9]+)"])
    maximum = first_regex(block, [rf"\bMaximum per insured person per calendar year:\s+({MONEY_AMOUNT_PATTERN})"])
    generic = first_regex(block, [r"\b(This is a mandatory generic pricing plan\.)"])
    supply = first_regex(block, [r"\bMaximum supply eligible:\s+(.{1,120}?)(?=\s+Co-ordination of benefits:)"])
    sdmp = first_regex(block, [r"\b(\d{1,3}%)\s+for\s+SDMP\s+Drugs"])
    costco = first_regex(block, [r"\b(\d{1,3}%)\s+for\s+all\s+other\s+drugs\s+purchased\s+at\s+Costco\s+Pharmacies"])
    other = first_regex(block, [r"\b(\d{1,3}%)\s+for\s+all\s+other\s+drugs\s+purchased\s+at\s+other\s+pharmacies"])
    coinsurance_parts = [
        value
        for value in (
            f"{sdmp} for SDMP drugs" if sdmp else "",
            f"{costco} at Costco Pharmacies" if costco else "",
            f"{other} at other pharmacies" if other else "",
        )
        if value
    ]
    return [
        f"Coinsurance: {'; '.join(coinsurance_parts)}",
        cost_sharing_line((("SDMP drugs", sdmp), ("Costco Pharmacies", costco), ("Other pharmacies", other))),
        f"Deductible: {deductible}",
        f"Plan Type: {plan_type}",
        f"Drug Maximum: {maximum} per calendar year",
        f"Drug Type: {generic}",
        f"Supply Limit: {supply}",
    ]


def build_equitable_ehc_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = equitable_block(schedule_text, [r"\bEmployee and dependent health benefits\b"], [r"\bHospital services:"])
    deductible = first_regex(block, [r"\bDeductible amount per calendar year for all other benefits:\s+(.{1,40}?)(?=\s+Reimbursement percentage)"])
    major = first_regex(block, [r"\bMajor services:\s+(\d{2,3}%)"])
    travel = first_regex(block, [r"\bTravel assist:\s+(\d{2,3}%)"])
    hospital = first_regex(block, [r"\bHospital services:\s+(\d{2,3}%)"])
    vision = first_regex(block, [r"\bVision care services:\s+(\d{2,3}%)"])
    return [
        f"Deductible: {deductible}",
        f"Major Services Coinsurance: {major}",
        cost_sharing_line((("Major services", major),)),
        f"Travel Assist Coinsurance: {travel}",
        f"Hospital Coinsurance: {hospital}",
        f"Vision Care Coinsurance: {vision}",
    ]


def build_equitable_paramedical_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = equitable_block(
        schedule_text,
        [r"#6\s+Maximums per insured person per calendar year for Paramedical services:"],
        [r"#7\s+Maximum in respect of Gender Affirmation Coverage"],
    )
    providers = [
        ("Chiropractor", r"Chiropractor \(including x-rays\)\s+(\$\s?\d[\d,]*)"),
        ("Registered massage therapist", r"Registered massage therapist\s+(\$\s?\d[\d,]*)"),
        ("Naturopath and Dietician combined", r"Naturopath .*? Dietician combined\s+(\$\s?\d[\d,]*)"),
        ("Osteopath", r"Osteopath \(including x-rays\)\s+(\$\s?\d[\d,]*)"),
        ("Physiotherapist and Athletic therapist combined", r"Physiotherapist and Athletic therapist combined\s+(\$\s?\d[\d,]*)"),
        ("Podiatrist/chiropodist", r"Podiatrist/chiropodist \(including x-rays\)\s+(\$\s?\d[\d,]*)"),
        ("Psychologist including MSW / Clinical counsellors", r"Psychologist \(including MSW / Clinical counsellors\)\s+(\$\s?\d[\d,]*)"),
        ("Specialist in acupuncture", r"Specialist in acupuncture\s+(\$\s?\d[\d,]*)"),
        ("Speech therapist and Audiologist combined", r"Speech therapist and Audiologist combined\s+(\$\s?\d[\d,]*)"),
    ]
    lines = [f"{label}: {first_regex(block, [pattern])}" for label, pattern in providers]
    referral = first_regex(block, [r"\b(A physician's prescription \(referral\) is not required[^.]*\.)"])
    if referral:
        lines.append(f"Referral: {referral}")
    return lines


def build_equitable_hospital_lines(_full_text: str, schedule_text: str) -> list[str]:
    coinsurance = first_regex(schedule_text, [r"\bHospital services:\s*(\d{2,3}%)"])
    coverage = "Semi-private hospital" if re.search(r"\bSemi-private hospital\b", schedule_text, flags=re.IGNORECASE) else ""
    return [f"Coinsurance: {coinsurance}", f"Coverage: {coverage}"]


def build_equitable_vision_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = equitable_block(schedule_text, [r"\bVision care\b"], [r"\bSummary of health benefit maximums\b"])
    eyewear = first_regex(block, [r"\bEye glasses or contact lenses or laser eye surgery:\s+maximum\s+(\$\s?\d[\d,]*(?:.{1,120}?24\s+months[^.]*)?)"])
    contacts = first_regex(block, [r"\bSpecial contact lenses.*?maximum:\s+(\$\s?\d[\d,]*(?:.{1,80}?lifetime maximum[^.]*)?)"])
    exams = first_regex(block, [r"\bEye examinations:\s+maximum:\s+(.{1,160}?24\s+months[^.]*)"])
    return [
        f"Vision Maximum: {eyewear}",
        f"Special Contact Lenses: {contacts}",
        f"Eye Exams: {exams}",
    ]


def build_equitable_travel_lines(full_text: str, schedule_text: str) -> list[str]:
    block = equitable_block(schedule_text, [r"#7 Services outside the province:"], [r"\bEmployee and dependent dental benefits\b"])
    time_limit = first_regex(block, [r"\bTime limit for commencement of emergency treatment.*?:\s+(.{1,260}?)(?=\s+Lifetime maximum amount)"])
    under_70 = first_regex(block, [rf"({MONEY_AMOUNT_PATTERN}\s+for actively-working employees under age 70[^,.;]*)"])
    over_70 = first_regex(block, [rf"({MONEY_AMOUNT_PATTERN}\s+for actively-working employees age 70 and older[^,.;]*)"])
    travel_percent = first_regex(schedule_text, [r"\bTravel assist:\s+(\d{1,3}%)"])
    per_trip_family = first_regex(
        full_text,
        [rf"\bmaximum of\s+({MONEY_AMOUNT_PATTERN})\s+for all such expenses for any one trip"],
    )
    return [
        f"Coinsurance: {travel_percent}",
        cost_sharing_line((("Eligible emergency travel expenses", travel_percent),)),
        f"Trip Duration: {time_limit}",
        f"Lifetime Maximum: {under_70}",
        f"Age 70+ Lifetime Maximum: {over_70}",
        f"Per-Trip Family Benefits Maximum: {per_trip_family} for all listed family expenses per trip",
    ]


def build_equitable_dental_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = equitable_block(schedule_text, [r"\bEmployee and dependent dental benefits\b"], [r"\bSurvivor benefit\b"])
    deductible = first_regex(block, [r"\bDeductible amount per calendar year:\s+(.{1,40}?)(?=\s+Type A)"])
    recall = first_regex(block, [r"\bRecall examination period:\s+(.{1,120}?)(?=\s+This Dental plan includes)"])
    type_a = first_regex(block, [r"\bType A:\s+(\d{2,3}%)"])
    type_b = first_regex(block, [r"\bType B:\s+(\d{2,3}%)"])
    type_c = first_regex(block, [r"\bType C:\s+(\d{2,3}%)"])
    basic_major_max = first_regex(block, [rf"\bAnnual calendar year maximum for Type A and type B combined:\s+({MONEY_AMOUNT_PATTERN})"])
    ortho_max = first_regex(block, [rf"\bLifetime maximum for Type C:\s+({MONEY_AMOUNT_PATTERN})"])
    ortho_eligibility = first_regex(block, [r"\b(Both adults and dependent children are eligible for Orthodontic services\.)"])
    return [
        f"Deductible: {deductible}",
        f"Recall: {recall}",
        f"Basic Services: {type_a}",
        f"Major Restorative Services: {type_b}",
        f"Orthodontics: {type_c}",
        cost_sharing_line((("Basic", type_a), ("Major", type_b), ("Orthodontics", type_c))),
        f"Basic and Major Maximum: {basic_major_max} per calendar year" if basic_major_max else "",
        f"Orthodontic Maximum: {ortho_max} lifetime" if ortho_max else "",
        f"Orthodontic Eligibility: {ortho_eligibility}",
    ]


def build_equitable_survivor_lines(full_text: str, schedule_text: str) -> list[str]:
    block = equitable_block(
        schedule_text,
        [r"\bSurvivor benefit\b"],
        [r"\bVirtual Healthcare\b", r"\bMaximum age for dependent children\b"],
    )
    included = first_regex(block, [r"\bFor the following benefits only:\s*(Health and dental)"])
    maximum_period = first_regex(
        block,
        [
            r"\bMaximum period for survivor benefit is the earlier of:\s*(?:[\u2022\-]\s*)?(24 months,?\s+or\s+(?:[\u2022\-]\s*)?the date that the dependent[\u2019']s coverage would have terminated had you lived)"
        ],
    )
    premium_waiver = "Premiums are waived" if re.search(
        r"\bPremiums are [\"']?waived[\"']?\s*\(are not payable\)",
        full_text,
        flags=re.IGNORECASE,
    ) else ""
    return [
        f"Included Benefits: {included}",
        f"Maximum Period: {maximum_period}",
        f"Premiums: {premium_waiver}",
    ]


def build_equitable_hcsa_lines(full_text: str, _schedule_text: str) -> list[str]:
    return build_spending_account_lines(full_text, "Health Care Spending Account")


def build_equitable_lsa_lines(full_text: str, _schedule_text: str) -> list[str]:
    return build_spending_account_lines(full_text, "Lifestyle Spending Account")


def build_spending_account_lines(full_text: str, account_name: str) -> list[str]:
    aliases = (
        r"(?:Health Care Spending Account|Health Spending Account|Healthcare Spending Account|HCSA|HSCA)"
        if account_name == "Health Care Spending Account"
        else r"(?:Lifestyle Spending Account|Life Style Spending Account|Wellness Spending Account|LSA)"
    )
    marker = re.search(rf"\b{aliases}\b", full_text, flags=re.IGNORECASE)
    if not marker:
        return []
    evidence = full_text[max(0, marker.start() - 300) : min(len(full_text), marker.end() + 1800)]
    provider_lines = add_documented_provider_line(account_name, [], evidence)
    amount = first_regex(
        evidence,
        [
            rf"\b(?:annual\s+)?(?:allocation|credit|maximum|benefit amount|account amount)\s*[:\-]?\s*({MONEY_AMOUNT_PATTERN})",
            rf"\b({MONEY_AMOUNT_PATTERN}\s+(?:per|each)\s+(?:calendar\s+year|plan\s+year|year))",
        ],
    )
    frequency = first_regex(evidence, [r"\b(per\s+(?:calendar\s+year|plan\s+year|year))\b"])
    return [*provider_lines, f"Amount: {amount}", f"Frequency: {frequency}"]


def build_equitable_virtual_care_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = equitable_block(schedule_text, [r"\bVirtual Healthcare"], [r"\bHomewood Health employee and family assistance program"])
    service = first_regex(block, [r"\b(Virtual Healthcare \(from Dialogue Health Technologies Inc\.\))"])
    termination = first_regex(schedule_text, [r"\bVirtual Healthcare \(from Dialogue Health Technologies Inc\.\) terminates on the\s+(.{1,50}?)(?=\s+Homewood Health)"])
    waiting = "3 months" if re.search(r"\b3\s+months\s+in\s+respect\s+of\s+Virtual Healthcare", schedule_text, flags=re.IGNORECASE) else ""
    return [f"Coverage: {service}", f"Waiting Period: {waiting}", f"Termination Age: {termination}"]


def build_equitable_eap_lines(full_text: str, schedule_text: str) -> list[str]:
    service = (
        "Homewood Health employee and family assistance program (EFAP)"
        if re.search(r"\bHomewood Health employee and family assistance program \(EFAP\)", full_text, flags=re.IGNORECASE)
        else ""
    )
    waiting = "3 months" if re.search(r"\b3\s+months\s+in\s+respect\s+of\s+Homewood Health employee and family assistance program", full_text, flags=re.IGNORECASE) else ""
    termination = first_regex(full_text, [r"\bHomewood Health employee and family assistance program \(EFAP\) terminates on the\s+(.{1,50}?)(?=\s+Services for your spouse)"])
    contact = first_regex(full_text, [r"\b(call to speak with a Homewood consultant anytime at\s+1\.888\.707\.2115[^.]*\.)"])
    return ["Provider: Homewood Health", f"Coverage: {service}", f"Waiting Period: {waiting}", f"Termination Age: {termination}", f"Contact: {contact}"]


def equitable_plan_design_benefit_name(section_name: str) -> str:
    return {
        "Life / AD&D": "Life / AD&D",
        "Long-Term Disability": "LTD",
        "Medical Services & Supplies": "Medical Services & Supplies",
        "Emergency Travel": "Travel",
        "Employee Assistance Program": "EAP",
    }.get(section_name, section_name)


def is_benepac_booklet(text: str) -> bool:
    cleaned = clean_text(text)
    return bool(
        re.search(r"\bBenepac\b", cleaned, flags=re.IGNORECASE)
        and re.search(r"\bSchedule of Benefits\b", cleaned, flags=re.IGNORECASE)
    )


def extract_benepac_schedule_plan_design_rows(
    full_text: str,
    pdf_path: Path,
    client_info: ClientInfo,
    class_name: str,
) -> list[ExtractionRow]:
    rows: list[ExtractionRow] = []
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for section_name, detail_text in build_benepac_summary_rows(full_text):
        if section_name in {"Carrier & Policy", "Eligibility Requirements"}:
            continue

        for detail in schedule_detail_lines(detail_text):
            if not detail or detail.startswith("- "):
                continue

            rows.append(
                ExtractionRow(
                    client_id=client_info.client_id,
                    client_name=client_info.client_name,
                    carrier=client_info.carrier,
                    class_name=class_name,
                    benefit=benepac_plan_design_benefit_name(section_name),
                    coverage_description=detail,
                    coinsurance=extract_benepac_coinsurance(detail),
                    maximum=extract_benepac_maximum(detail),
                    frequency=extract_benepac_frequency(detail),
                    waiting_period="",
                    termination_age=extract_benepac_termination_age(detail),
                    notes="Benepac Schedule of Benefits",
                    source_file=pdf_path.name,
                    source_section=section_name,
                    confidence="High",
                    review_flag="",
                    last_updated=now,
                )
            )
    return rows


def build_benepac_schedule_summary_lines(section_name: str, full_text: str) -> list[str]:
    rows = dict(build_benepac_summary_rows(full_text))
    detail_text = rows.get(section_name, "")
    return schedule_detail_lines(detail_text)


def build_benepac_summary_rows(full_text: str) -> list[tuple[str, str]]:
    schedule_text = extract_benepac_schedule_text(full_text)
    if not schedule_text:
        return []

    builders = [
        ("Carrier & Policy", build_benepac_carrier_policy_lines),
        ("Eligibility Requirements", build_benepac_eligibility_lines),
        ("Life / AD&D", build_benepac_life_adnd_lines),
        ("Medical Second Opinion", build_benepac_second_opinion_lines),
        ("Dependent Life", build_benepac_dependent_life_lines),
        ("Long-Term Disability", build_benepac_ltd_lines),
        ("Critical Illness", build_benepac_critical_illness_lines),
        ("Prescription Drugs", build_benepac_prescription_drug_lines),
        ("Medical Services & Supplies", build_benepac_ehs_lines),
        ("Paramedicals", build_benepac_paramedical_lines),
        ("Mental Health Practitioners", build_benepac_mental_health_lines),
        ("Hospital", build_benepac_hospital_lines),
        ("Vision Care", build_benepac_vision_lines),
        ("Emergency Travel", build_benepac_travel_lines),
        ("Dental Care", build_benepac_dental_lines),
        ("Employee Assistance Program", build_benepac_eap_lines),
    ]

    rows: list[tuple[str, str]] = []
    for section_name, builder in builders:
        lines = builder(full_text, schedule_text)
        if lines:
            rows.append((section_name, "\n".join(lines)))
    return rows


def extract_benepac_schedule_text(full_text: str) -> str:
    selected_pages: list[str] = []
    in_schedule = False
    for _page_number, page_text in split_marked_pages(full_text):
        lines = [clean_text(line) for line in page_text.splitlines() if clean_text(line)]
        if not lines:
            continue

        top_lines = " ".join(lines[:4])
        page_clean = clean_text(page_text)
        is_start = bool(
            re.search(r"\bSchedule of Benefits\b", top_lines, flags=re.IGNORECASE)
            and re.search(r"\bBenefit Details\b", page_clean, flags=re.IGNORECASE)
        )
        if is_start:
            in_schedule = True
        elif in_schedule and re.search(
            r"\b(?:Empire Life Insurance Company\s+This section|GreenShield\s+This section|General Provisions|General Information|Definitions)\b",
            top_lines,
            flags=re.IGNORECASE,
        ):
            break

        if in_schedule:
            selected_pages.append(page_text)

    return clean_text(" ".join(selected_pages))


def build_benepac_carrier_policy_lines(full_text: str, _schedule_text: str) -> list[str]:
    block = extract_benepac_group_info_block(full_text)
    lines = []
    group_lines = extract_benepac_group_number_lines(block)
    for line in group_lines:
        lines.append(line)
    return unique_in_order(lines)


def build_benepac_eligibility_lines(full_text: str, _schedule_text: str) -> list[str]:
    block = extract_benepac_group_info_block(full_text)
    group_name = extract_benepac_value_after_label(block, "Group Name and Number(s)")
    employee_class = extract_benepac_value_after_label(block, "Employee Classification")
    effective_dates = extract_benepac_effective_dates(block)
    return compact_detail_lines(
        [
            f"Group Name: {group_name}",
            f"Employee Classification: {employee_class}",
            f"Effective Date: {'; '.join(effective_dates)}",
        ]
    )


def build_benepac_life_adnd_lines(_full_text: str, schedule_text: str) -> list[str]:
    life = benepac_block(schedule_text, [r"\bLife Insurance\b"], [r"\bPeople Connect Carepath\b", r"\bAccidental Death"])
    adnd = benepac_block(schedule_text, [r"\bAccidental Death,\s*Disease\b", r"\bAccidental Death"], [r"\bDependent Life Insurance\b", r"\bLong Term Disability\b", r"\bGroup Critical Illness\b", r"\bExtended Health Services\b"])
    life_termination = first_regex(life, [r"\bTermination Age:\s*(Insurance terminates[^.]*?(?:age\s+\d+|retirement|Disabled))"])
    adnd_termination = first_regex(adnd, [r"\bTermination Age:\s*(Insurance terminates[^.]*?(?:age\s+\d+|Disabled))"])
    return compact_detail_lines(
        [
            f"Life Benefit Amount: {first_regex(life, [rf'\\bBenefit Amount:\\s*({MONEY_AMOUNT_PATTERN})'])}",
            f"Life Non-Evidence Maximum: {first_regex(life, [rf'\\bNon-Evidence Maximum:\\s*({MONEY_AMOUNT_PATTERN})'])}",
            f"Life Reduction: {first_regex(life, [r'\\b(Insurance reduces by\\s+.{1,120}?)(?=\\s+(?:null|Termination Age|People Connect|Accidental|$))'])}",
            f"Life Termination Age: {life_termination}",
            f"AD&D Benefit Amount: {first_regex(adnd, [rf'\\bAccidental Death\\s*-\\s*({MONEY_AMOUNT_PATTERN})', rf'\\bBenefit Amount:\\s*({MONEY_AMOUNT_PATTERN})'])}",
            f"AD&D Non-Evidence Maximum: {first_regex(adnd, [rf'\\bNon-Evidence Maximum:\\s*({MONEY_AMOUNT_PATTERN})'])}",
            f"AD&D Dismemberment Schedule: {first_regex(adnd, [r'\\b(Accidental/Disease Dismemberment\\s*[\\-\\u2013\\u2014]\\s*loss schedule\\s+.{1,220}?)(?=\\s+(?:Critical Disease|Reduction|Termination Age|Dependent Life|$))'])}",
            f"AD&D Termination Age: {adnd_termination}",
        ]
    )


def build_benepac_second_opinion_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = benepac_block(schedule_text, [r"\bPeople Connect Carepath\b"], [r"\bAccidental Death"])
    return compact_detail_lines(
        [
            f"Coverage: {first_regex(block, [r'\\b(Included with Life Insurance benefit at no additional cost)'])}",
            f"Coverage For: {first_regex(block, [r'\\bCoverage For:\\s*(.{1,90}?)(?=\\s+Details:)'])}",
            f"Details: {first_regex(block, [r'\\bDetails:\\s*(.{1,420}?)(?=\\s+Termination Age:)'])}",
            f"Termination Age: {first_regex(block, [r'\\bTermination Age:\\s*(Insurance terminates[^.]*?age\\s+\\d+)'])}",
        ]
    )


def build_benepac_dependent_life_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = benepac_block(schedule_text, [r"\bDependent Life Insurance\b"], [r"\bLong Term Disability\b", r"\bGroup Critical Illness\b", r"\bExtended Health Services\b"])
    return compact_detail_lines(
        [
            f"Spouse Benefit: {first_regex(block, [rf'\\bSpouse:\\s*({MONEY_AMOUNT_PATTERN})'])}",
            f"Child Benefit: {first_regex(block, [rf'\\bEach Child:\\s*({MONEY_AMOUNT_PATTERN})'])}",
            f"Dependent Children: {first_regex(block, [r'\\bDependent Children:\\s*(.{1,220}?)(?=\\s+Termination Age:)'])}",
            f"Termination Age: {first_regex(block, [r'\\bTermination Age:\\s*(Insurance terminates[^.]*?(?:age\\s+\\d+|$))'])}",
        ]
    )


def build_benepac_ltd_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = benepac_block(schedule_text, [r"\bLong Term Disability\b"], [r"\bGroup Critical Illness\b", r"\bExtended Health Services\b"])
    return compact_detail_lines(
        [
            f"Benefit Schedule: {first_regex(block, [r'\\bBenefit Amount:\\s*(.{1,240}?)(?=\\s+Non-Evidence Maximum:)'])}",
            f"Non-Evidence Maximum: {first_regex(block, [rf'\\bNon-Evidence Maximum:\\s*({MONEY_AMOUNT_PATTERN})'])}",
            f"Benefit Taxability: {first_regex(block, [r'\\bBenefit Taxability:\\s*(.{1,50}?)(?=\\s+Maximum Income)'])}",
            f"Maximum Income From All Sources: {first_regex(block, [r'\\bMaximum Income From All Sources:\\s*(.{1,140}?)(?=\\s+Elimination Period:)'])}",
            f"Elimination Period: {first_regex(block, [r'\\bElimination Period:\\s*(.{1,50}?)(?=\\s+End of)'])}",
            f"Own Occupation Period: {first_regex(block, [r'\\bEnd of .own occupation. period for definition of total disability:\\s*(.{1,80}?)(?=\\s+Maximum Benefit Period:)'])}",
            f"Benefit Period: {first_regex(block, [r'\\bMaximum Benefit Period:\\s*(.{1,50}?)(?=\\s+Termination Age:)'])}",
            f"Termination Age: {first_regex(block, [r'\\bTermination Age:\\s*(Insurance terminates[^.]*?(?:age\\s+\\d+|$))'])}",
        ]
    )


def build_benepac_critical_illness_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = benepac_block(schedule_text, [r"\bGroup Critical Illness\b"], [r"\bExtended Health Services\b"])
    return compact_detail_lines(
        [
            f"Benefit Amount: {first_regex(block, [rf'({MONEY_AMOUNT_PATTERN}\\s+per\\s+employee)', rf'\\bBenefit Amount:\\s*({MONEY_AMOUNT_PATTERN})'])}",
            f"Coverage: {first_regex(block, [r'\\bCoverage:\\s*(.{1,220}?)(?=\\s+Termination Age:)'])}",
            f"Termination Age: {first_regex(block, [r'\\bTermination Age:\\s*(Insurance terminates[^.]*?(?:age\\s+\\d+|$))'])}",
        ]
    )


def build_benepac_prescription_drug_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = benepac_ehs_block(schedule_text)
    return compact_detail_lines(
        [
            f"Deductible: {first_regex(block, [r'\\bPrescription Drugs Deductible:\\s*(.{1,160}?)(?=\\s+EHS Co-insurance:)'])}",
            f"Coinsurance: {first_regex(block, [r'\\bPrescription Drugs Co-insurance:\\s*(\\d{1,3}%)'])}",
            f"Plan Type: {first_regex(block, [r'\\bPrescription Drug Plan:\\s*(.{1,80}?)(?=\\s+Drug Maximum:)'])}",
            f"Drug Max: {first_regex(block, [r'\\bDrug Maximum:\\s*(Unlimited|' + MONEY_AMOUNT_PATTERN + r')'])}",
        ]
    )


def build_benepac_ehs_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = benepac_ehs_block(schedule_text)
    return compact_detail_lines(
        [
            f"Single Deductible: {first_regex(block, [r'\\bSingle\\s*-\\s*(.{1,80}?)(?=\\s+(?:Green Shield Canada\\s+)?Family\\s*-)'])}",
            f"Family Deductible: {first_regex(block, [r'\\bFamily\\s*-\\s*(.{1,80}?)(?=\\s+Prescription Drugs Deductible:)'])}",
            f"EHS Coinsurance: {first_regex(block, [r'\\bEHS Co-insurance:\\s*(\\d{1,3}%)'])}",
            f"Overall Maximum: {first_regex(block, [r'\\bOverall Maximum:\\s*(Unlimited|' + MONEY_AMOUNT_PATTERN + r')'])}",
            f"Dependent Children: {first_regex(block, [r'\\bDependent Children:\\s*(.{1,220}?)(?=\\s+Termination Age:)'])}",
            f"Termination Age: {first_regex(block, [r'\\bTermination Age:\\s*(Insurance terminates[^.]*?(?:age\\s+\\d+|$))'])}",
        ]
    )


def build_benepac_paramedical_lines(full_text: str, schedule_text: str) -> list[str]:
    block = benepac_ehs_block(schedule_text)
    services = first_regex(block, [r'\\bParamedical Services:\\s*(.{1,220}?)(?=\\s+Hospital Accommodation:)'])
    if re.search(r"\bNone\b", services, flags=re.IGNORECASE):
        return compact_detail_lines(
            [
                f"Coinsurance: {first_regex(block, [r'\\bParamedical Co-insurance:\\s*(\\d{1,3}%)'])}",
                "Services: None",
            ]
        )

    amount = first_regex(services, [rf"({MONEY_AMOUNT_PATTERN}\s+per\s+practitioner\s+per\s+year)"])
    if not amount:
        amount = first_regex(extract_benepac_professional_services_block(full_text), [rf"({MONEY_AMOUNT_PATTERN}\s+maximum\s+per\s+practitioner\s+per\s+calendar\s+year)"])
    amount = re.sub(r"\s+maximum\b", "", clean_text(amount), flags=re.IGNORECASE).replace("per year", "per calendar year")

    lines = compact_detail_lines(
        [
            f"Coinsurance: {first_regex(block, [r'\\bParamedical Co-insurance:\\s*(\\d{1,3}%)'])}",
            f"Benefit Amount: {amount}",
        ]
    )
    practitioners = extract_benepac_standard_practitioners(full_text)
    if practitioners:
        lines.append("Practitioners")
        lines.extend(practitioners)
    return unique_in_order(lines)


def build_benepac_mental_health_lines(full_text: str, schedule_text: str) -> list[str]:
    block = benepac_ehs_block(schedule_text)
    services = first_regex(block, [r'\\bParamedical Services:\\s*(.{1,220}?)(?=\\s+Hospital Accommodation:)'])
    practitioners = extract_benepac_mental_health_practitioners(full_text)
    if not practitioners or re.search(r"\bNone\b", services, flags=re.IGNORECASE):
        return []

    amount = first_regex(extract_benepac_professional_services_block(full_text), [rf"({MONEY_AMOUNT_PATTERN}\s+maximum\s+per\s+practitioner\s+per\s+calendar\s+year)"])
    amount = re.sub(r"\s+maximum\b", "", clean_text(amount), flags=re.IGNORECASE)
    lines = compact_detail_lines(
        [
            f"Coinsurance: {first_regex(block, [r'\\bParamedical Co-insurance:\\s*(\\d{1,3}%)'])}",
            f"Benefit Amount: {amount}",
        ]
    )
    lines.append("Practitioners")
    lines.extend(practitioners)
    return unique_in_order(lines)


def extract_benepac_professional_services_block(full_text: str) -> str:
    cleaned = clean_text(full_text)
    marker = None
    for candidate in re.finditer(r"\bProfessional Services\b", cleaned, flags=re.IGNORECASE):
        preview = cleaned[candidate.start() : candidate.start() + 2600]
        if re.search(r"\bmaximum\s+per\s+practitioner\b|\bPhysiotherapist\b|\bChiropractor\b", preview, flags=re.IGNORECASE):
            marker = candidate
            break
    if marker is None:
        return ""

    remaining = cleaned[marker.start() :]
    end_match = re.search(r"\bPrivate Duty Nursing\b", remaining, flags=re.IGNORECASE)
    if end_match:
        remaining = remaining[: end_match.start()]
    return remaining


def extract_benepac_standard_practitioners(full_text: str) -> list[str]:
    block = extract_benepac_professional_services_block(full_text)
    if not block:
        return []

    practitioner_patterns = [
        ("Acupuncturist", r"\bAcupuncturist\b"),
        ("Chiropractor", r"\bChiropractor\b"),
        ("Dietician", r"\bDieticians?\b"),
        ("Massage Therapist", r"\bRegistered Massage Therapist\b|\bMassage Therapist\b"),
        ("Naturopath", r"\bNaturopath\b"),
        ("Osteopath", r"\bOsteopath\b"),
        ("Physiotherapist / Athletic Therapist", r"\bPhysiotherapist\b|\bCertified Athletic Therapist\b"),
        ("Podiatrist / Chiropodist", r"\bPodiatrist/Chiropodist\b|\bPodiatrist\b|\bChiropodist\b"),
        ("Speech Therapist", r"\bSpeech Therapist\b|\bSpeech Therapist/Pathologist\b"),
    ]
    return [
        display
        for display, pattern in practitioner_patterns
        if re.search(pattern, block, flags=re.IGNORECASE)
    ]


def extract_benepac_mental_health_practitioners(full_text: str) -> list[str]:
    block = extract_benepac_professional_services_block(full_text)
    if not re.search(r"\bMental Health practitioners\b|\bpsychologist\b", block, flags=re.IGNORECASE):
        return []

    # Benepac/GreenShield groups these under "Mental Health practitioners" in the booklet.
    return [
        "Psychologist",
        "Social Worker",
        "Registered Clinical Counsellor",
        "Master of Social Work",
        "Psychotherapist",
    ]


def build_benepac_hospital_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = benepac_ehs_block(schedule_text)
    return compact_detail_lines(
        [
            f"Hospital Accommodation: {first_regex(block, [r'\\bHospital Accommodation:\\s*(.{1,90}?)(?=\\s+(?:Gender Affirmation|Optometric Eye|Travel Benefits|Overall Maximum|$))'])}",
        ]
    )


def build_benepac_vision_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = benepac_ehs_block(schedule_text)
    return compact_detail_lines(
        [
            f"Eye Exams: {first_regex(block, [r'\\bOptometric Eye Examinations:\\s*(.{1,180}?)(?=\\s+(?:Vision Benefit|Travel Benefits|Overall Maximum|$))'])}",
            f"Vision Benefit: {first_regex(block, [r'\\bVision Benefit:\\s*(.{1,180}?)(?=\\s+(?:Travel Benefits|Overall Maximum|$))'])}",
        ]
    )


def build_benepac_travel_lines(_full_text: str, schedule_text: str) -> list[str]:
    travel = benepac_block(schedule_text, [r"\bTravel Plan\b"], [r"\bDental Care\b", r"\bEmployee Assistance\b"])
    if travel:
        return compact_detail_lines(
            [
                f"Deductible: {first_regex(travel, [r'\\bDeductible:\\s*(.{1,60}?)(?=\\s+Co-insurance:)'])}",
                f"Coinsurance: {first_regex(travel, [r'\\bCo-insurance:\\s*(\\d{1,3}%)'])}",
                f"Maximum Number of Days per Trip: {first_regex(travel, [r'\\bMaximum Number of Days per Trip:\\s*(.{1,50}?)(?=\\s+Emergency Services:)'])}",
                f"Emergency Services: {first_regex(travel, [rf'\\bEmergency Services:\\s*({MONEY_AMOUNT_PATTERN}\\s+per\\s+covered\\s+person)'])}",
                f"Referral Services: {first_regex(travel, [rf'\\bReferral Services:\\s*({MONEY_AMOUNT_PATTERN}\\s+per\\s+covered\\s+person\\s+per\\s+calendar\\s+year)'])}",
            ]
        )

    ehs = benepac_ehs_block(schedule_text)
    return compact_detail_lines(
        [
            f"Travel Benefits: {first_regex(ehs, [r'\\bTravel Benefits:\\s*(Covered|Not covered|.{1,80}?)(?=\\s+Overall Maximum:)'])}",
        ]
    )


def build_benepac_dental_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = benepac_block(schedule_text, [r"\bDental Care\b"], [r"\bEmployee Assistance\b"])
    return compact_detail_lines(
        [
            f"Single Deductible: {first_regex(block, [r'\\bSingle\\s*-\\s*(.{1,80}?)(?=\\s+(?:Green Shield Canada\\s+)?Family\\s*-)'])}",
            f"Family Deductible: {first_regex(block, [r'\\bFamily\\s*-\\s*(.{1,80}?)(?=\\s+Basic Services)'])}",
            f"Basic Services: {first_regex(block, [rf'\\bBasic Services \\(includes Endo/Perio\\):\\s*(\\d{{1,3}}%\\s*\\({MONEY_AMOUNT_PATTERN}\\s+per\\s+calendar\\s+year\\))'])}",
            f"Recall Exams: {first_regex(block, [r'\\bRecall Exams:\\s*(.{1,80}?)(?=\\s+Dependent Children:)'])}",
            f"Dependent Children: {first_regex(block, [r'\\bDependent Children:\\s*(.{1,220}?)(?=\\s+Termination Age:)'])}",
            f"Termination Age: {first_regex(block, [r'\\bTermination Age:\\s*(Insurance terminates[^.]*?(?:age\\s+\\d+|$))'])}",
        ]
    )


def build_benepac_eap_lines(_full_text: str, schedule_text: str) -> list[str]:
    block = benepac_block(schedule_text, [r"\bEmployee Assistance\b"], [r"$"])
    return compact_detail_lines(
        [
            f"Coverage: {first_regex(block, [r'\\bConfidential\\s+(.{1,180}?)(?=\\s+Clinical counselling)'])}",
            f"Coverage For: {first_regex(block, [r'\\bCoverage For:\\s*(.{1,90}?)(?=\\s+Coverage:)'])}",
            f"Support Hours: {first_regex(block, [r'\\bCoverage:\\s*(up to\\s+\\d+\\s+hours\\s+.{1,180}?)(?=\\s+Termination Age:)'])}",
            f"Termination Age: {first_regex(block, [r'\\bTermination Age:\\s*(Insurance terminates[^.]*?(?:age\\s+\\d+|$))'])}",
        ]
    )


def benepac_ehs_block(schedule_text: str) -> str:
    return benepac_block(schedule_text, [r"\bExtended Health Services\s*\(EHS\)"], [r"\bTravel Plan\b", r"\bDental Care\b", r"\bEmployee Assistance\b"])


def benepac_block(schedule_text: str, start_patterns: list[str], end_patterns: list[str]) -> str:
    start_positions: list[tuple[int, int]] = []
    for pattern in start_patterns:
        match = re.search(pattern, schedule_text, flags=re.IGNORECASE)
        if match:
            start_positions.append((match.start(), match.end()))
    if not start_positions:
        return ""

    start, _start_end = min(start_positions, key=lambda item: item[0])
    end = len(schedule_text)
    for pattern in end_patterns:
        if pattern == r"$":
            continue
        match = re.search(pattern, schedule_text[start + 1 :], flags=re.IGNORECASE)
        if match:
            end = min(end, start + 1 + match.start())
    return clean_text(schedule_text[start:end])


def extract_benepac_group_info_block(full_text: str) -> str:
    pages = split_marked_pages(full_text)
    if not pages:
        return ""
    first_page = pages[0][1]
    end = re.search(r"\b\[\[PAGE\s+\d+\]\]|\bBenepac\b|\bIntroduction\b", first_page, flags=re.IGNORECASE)
    return first_page[: end.start()] if end else first_page


def extract_benepac_group_number_lines(group_info: str) -> list[str]:
    lines = [clean_text(line) for line in group_info.splitlines() if clean_text(line)]
    results: list[str] = []
    in_group_numbers = False
    for line in lines:
        if re.fullmatch(r"Group\s+Number\(s\)", line, flags=re.IGNORECASE):
            in_group_numbers = True
            continue
        if in_group_numbers and re.fullmatch(r"Effective\s+Date", line, flags=re.IGNORECASE):
            break
        if in_group_numbers:
            results.append(line)
    if results:
        return results

    compact = clean_text(group_info)
    match = re.search(r"\bGroup Number\(s\)\s+(.+?)\s+Effective Date\b", compact, flags=re.IGNORECASE)
    if not match:
        return []
    return split_benepac_provider_lines(match.group(1))


def split_benepac_provider_lines(value: str) -> list[str]:
    provider_pattern = r"(?:People Corporation Health Solutions|Industrial Alliance|Green Shield Canada|The Empire Life Insurance Company|Empire Life|GreenShield)"
    parts = re.split(rf"\s+(?={provider_pattern}\b)", clean_text(value), flags=re.IGNORECASE)
    return [clean_text(part) for part in parts if clean_text(part)]


def extract_benepac_value_after_label(group_info: str, label: str) -> str:
    lines = [clean_text(line) for line in group_info.splitlines() if clean_text(line)]
    for index, line in enumerate(lines):
        if not re.fullmatch(re.escape(label), line, flags=re.IGNORECASE):
            continue
        values = []
        for candidate in lines[index + 1 :]:
            if re.fullmatch(r"Employee Classification|Group Number\(s\)|Effective Date", candidate, flags=re.IGNORECASE):
                break
            values.append(candidate.strip(" ,"))
        return clean_text(" ".join(values))

    compact = clean_text(group_info)
    next_labels = r"Employee Classification|Group Number\(s\)|Effective Date"
    match = re.search(rf"\b{re.escape(label)}\s+(.+?)\s+(?:{next_labels})\b", compact, flags=re.IGNORECASE)
    return clean_text(match.group(1)).strip(" ,") if match else ""


def extract_benepac_effective_dates(group_info: str) -> list[str]:
    compact = clean_text(group_info)
    match = re.search(r"\bEffective Date\s+(.+)$", compact, flags=re.IGNORECASE)
    if not match:
        return []
    value = match.group(1)
    dates = re.findall(r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b", value, flags=re.IGNORECASE)
    month_years = re.findall(r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December),\s+\d{4}\b", value, flags=re.IGNORECASE)
    return unique_in_order([*dates, *month_years])


def benepac_plan_design_benefit_name(section_name: str) -> str:
    return {
        "Life / AD&D": "Life / AD&D",
        "Medical Services & Supplies": "Extended Health Services",
        "Emergency Travel": "Travel",
    }.get(section_name, section_name)


def extract_benepac_coinsurance(detail: str) -> str:
    if "coinsurance" not in detail.lower() and "co-insurance" not in detail.lower() and "basic services" not in detail.lower():
        return ""
    return first_regex(detail, [r"(\d{1,3}%)"])


def extract_benepac_maximum(detail: str) -> str:
    if re.search(r"\b(?:maximum|max|benefit amount|benefit|services|vision benefit|emergency services|referral services|drug max)\b", detail, flags=re.IGNORECASE):
        money = first_regex(detail, [rf"({MONEY_AMOUNT_PATTERN}(?:\s+per\s+(?:employee|month|calendar\s+year|covered\s+person|covered\s+person\s+per\s+calendar\s+year|person))?)"])
        if money:
            return money
        unlimited = first_regex(detail, [r"\b(Unlimited)\b"])
        if unlimited:
            return unlimited
    return ""


def extract_benepac_frequency(detail: str) -> str:
    return first_regex(detail, [r"\b(\d+\s+month\s+period|every\s+\d+\s+months?|twice\s+in\s+\d+\s+months|per\s+calendar\s+year|per\s+lifetime|lifetime)\b"])


def extract_benepac_termination_age(detail: str) -> str:
    if "termination age" not in detail.lower():
        return ""
    return first_regex(detail, [r"\b(age\s+\d+)\b"])


def build_schedule_life_summary_lines(full_text: str) -> list[str]:
    details = extract_schedule_life_details(full_text)
    if not details:
        return []

    asi_details = extract_schedule_asi_details(full_text)

    return compact_detail_lines(
        [
            f"Life Benefit Amount: {details.benefit_amount}",
            f"Maximum Benefit: {details.maximum_benefit}",
            f"ASI Benefit Amount: {asi_details.benefit_amount if asi_details else details.adnd_amount}",
            f"ASI Critical Disease Benefit: {asi_details.critical_disease_benefit if asi_details else ''}",
            f"ASI Serious Illness Benefit: {asi_details.serious_illness_benefit if asi_details else ''}",
            f"Benefit Reduction: {details.benefit_reduction}",
            f"Waiver of Premium: {details.waiver_of_premium}",
            f"Termination Age: {details.termination_age}",
        ]
    )


def extract_schedule_life_details(full_text: str) -> ScheduleLifeDetails | None:
    page_number, page = find_group_life_schedule_page(full_text)
    if not page:
        return None

    benefit_amount = first_regex(page, [r"\bBenefit Amount\s+(.{1,220}?)(?=\s+Maximum Benefit\b)"])
    maximum_benefit = first_regex(
        page,
        [
            rf"\bMaximum Benefit\s+({MONEY_AMOUNT_PATTERN})(?=\s+(?:with|without)\s+Evidence of Insurability\b)",
            rf"\bMaximum Benefit\s+with Evidence of Insurability\s+({MONEY_AMOUNT_PATTERN})",
            rf"\bMaximum Benefit\s+without Evidence of Insurability\s+({MONEY_AMOUNT_PATTERN})",
            rf"\bMaximum Benefit\s+({MONEY_AMOUNT_PATTERN})",
        ],
    )
    benefit_reduction = first_regex(page, [r"\bBenefit Reduction\s+(.{1,180}?)(?=\s+Waiver of Premium\b)"])
    waiver = first_regex(page, [r"\bWaiver of Premium\s+(.{1,420}?)(?=\s+Definition of Total Disability\b)"])
    termination = first_regex(
        page,
        [r"\bTermination Age\s+(.{1,180}?)(?=$|\s+(?:Dependent Life|Accident|Schedule of Benefits))"],
    )

    return ScheduleLifeDetails(
        benefit_amount=normalize_schedule_life_benefit_amount(benefit_amount),
        maximum_benefit=clean_money(maximum_benefit),
        benefit_reduction=clean_text(benefit_reduction),
        waiver_of_premium=clean_text(waiver),
        termination_age=normalize_life_termination_age(termination),
        adnd_amount=extract_schedule_adnd_asi_amount(full_text),
        source_page=page_number,
    )


def build_schedule_dependent_life_summary_lines(full_text: str) -> list[str]:
    details = extract_schedule_dependent_life_details(full_text)
    if not details:
        return ["Coverage: Not covered"] if is_grouphealth_booklet(full_text) else []

    return compact_detail_lines(
        [
            f"Spouse: {details.spouse_amount}",
            f"Dependent: {details.dependent_child_amount}",
            f"Dependent Child Eligibility: {details.dependent_child_eligibility}",
            f"Stillbirth Benefit: {details.stillbirth_benefit}",
            f"Termination Age: {details.termination_age}",
        ]
    )


def extract_schedule_dependent_life_details(full_text: str) -> ScheduleDependentLifeDetails | None:
    page_number, page = find_dependent_life_schedule_page(full_text)
    if not page:
        return None

    spouse_amount = first_regex(
        page,
        [rf"\bBenefit Amount for Spouse\s+({MONEY_AMOUNT_PATTERN})"],
    )
    dependent_child_amount = first_regex(
        page,
        [rf"\bBenefit Amount for Each Child\s+({MONEY_AMOUNT_PATTERN})"],
    )
    if not spouse_amount or not dependent_child_amount:
        return None

    termination = first_regex(
        page,
        [r"\bTermination Age\s+(.{1,180}?)(?=$|\s+(?:Accident|Long[- ]Term|Schedule of Benefits))"],
    )
    dependent_child_eligibility = ""
    if re.search(r"\bDependent Child Eligibility\b", page, flags=re.IGNORECASE):
        child_age = first_regex(page, [r"\bage\s+(\d{1,2})\b"])
        student_age = first_regex(
            page,
            [
                r"\bage\s+(\d{1,2})\s+if\s+in\s+full[- ]time\s+attendance\s+as\s+a\s+student\b",
                r"\bage\s+(\d{1,2})\s+if\s+a\s+student\b",
            ],
        )
        if child_age and student_age:
            dependent_child_eligibility = (
                f"Birth to age {child_age}, or age {student_age} if attending a recognized "
                "educational institute full-time"
            )
    stillbirth_benefit = ""
    if re.search(
        r"\bStillbirth Benefit\b.{0,220}?\bpay\b.{0,120}?\bDependent Child Benefit Amount\b",
        page,
        flags=re.IGNORECASE,
    ):
        stillbirth_benefit = "Included; pays the Dependent Child Benefit Amount"
    return ScheduleDependentLifeDetails(
        spouse_amount=clean_money(spouse_amount),
        dependent_child_amount=clean_money(dependent_child_amount),
        termination_age=normalize_life_termination_age(termination),
        source_page=page_number,
        dependent_child_eligibility=dependent_child_eligibility,
        stillbirth_benefit=stillbirth_benefit,
    )


def find_dependent_life_schedule_page(full_text: str) -> tuple[int, str]:
    for page_number, page_text in split_marked_pages(full_text):
        cleaned_page = clean_text(page_text)
        if "schedule of benefits" not in cleaned_page.lower():
            continue

        lines = [clean_text(line) for line in page_text.splitlines() if clean_text(line)]
        for line_index, line in enumerate(lines):
            if not re.fullmatch(r"Dependent Life Insurance", line, flags=re.IGNORECASE):
                continue

            nearby_lines = lines[line_index + 1 : line_index + 10]
            has_spouse_amount = any(
                re.match(r"^Benefit Amount for Spouse\b", candidate, flags=re.IGNORECASE)
                for candidate in nearby_lines
            )
            has_child_amount = any(
                re.match(r"^Benefit Amount for Each Child\b", candidate, flags=re.IGNORECASE)
                for candidate in nearby_lines
            )
            if has_spouse_amount and has_child_amount:
                return page_number, cleaned_page

    return 0, ""


def find_group_life_schedule_page(full_text: str) -> tuple[int, str]:
    for page_number, page_text in split_marked_pages(full_text):
        cleaned_page = clean_text(page_text)
        if "schedule of benefits" not in cleaned_page.lower():
            continue
        lines = [clean_text(line) for line in page_text.splitlines() if clean_text(line)]
        for line_index, line in enumerate(lines):
            inline_heading = re.fullmatch(
                r"Group Life Insurance\s+Benefit Amount(?:\s+.*)?",
                line,
                flags=re.IGNORECASE,
            )
            if inline_heading:
                return page_number, cleaned_page
            if not re.fullmatch(r"Group Life Insurance", line, flags=re.IGNORECASE):
                continue
            if any(
                re.match(r"^Benefit Amount\b", candidate, flags=re.IGNORECASE)
                for candidate in lines[line_index + 1 : line_index + 6]
            ):
                return page_number, cleaned_page
    return 0, ""


def normalize_schedule_life_benefit_amount(value: str) -> str:
    cleaned = normalize_money_phrase(value)
    match = re.match(
        r"^(?P<multiplier>\d+(?:\.\d+)?)\s*(?:x|times?)\s+(?:your\s+)?(?:annual\s+)?(?P<basis>salary|earnings)\b(?P<remainder>.*)$",
        cleaned,
        flags=re.IGNORECASE,
    )
    if not match:
        return cleaned

    multiplier = match.group("multiplier")
    if multiplier.endswith(".0"):
        multiplier = multiplier[:-2]
    basis = match.group("basis").lower()
    remainder = clean_text(match.group("remainder"))
    separator = "" if remainder.startswith((",", ";", ".")) else " "
    return clean_text(f"{multiplier} x annual {basis}{separator}{remainder}")


def normalize_life_termination_age(value: str) -> str:
    cleaned = clean_text(value)
    age = first_regex(cleaned, [r"\b(?:turn|age)\s*(\d{2})\b"])
    has_retirement = bool(re.search(r"\bretir(?:e[sd]?|ement)\b", cleaned, flags=re.IGNORECASE))
    if age and has_retirement:
        return f"Age {age}; retirement"
    if age:
        return f"Age {age}"
    if has_retirement:
        return "Retirement"
    return cleaned


def extract_schedule_adnd_asi_amount(full_text: str) -> str:
    page = find_schedule_page_text(full_text, "Accident & Serious Illness", "Benefit Amount")
    if not page:
        page = find_schedule_page_text(full_text, "Accidental Death", "Benefit Amount")
    if not page:
        return ""

    return first_regex(
        page,
        [
            r"\bBenefit Amount\s+(.{1,160}?)(?=\s+(?:Benefit Reduction|Enhanced Schedule|Paralysis Benefits|Definition|Maximum Benefit|$))",
            r"\bPrincipal Sum\s+(.{1,160}?)(?=\s+(?:Benefit Reduction|Enhanced Schedule|Paralysis Benefits|Definition|Maximum Benefit|$))",
        ],
    )


def extract_schedule_asi_details(full_text: str) -> ScheduleAsiDetails | None:
    page_number, page = find_schedule_page(
        full_text,
        "Accident & Serious Illness",
        "Critical Disease Benefit",
    )
    if not page:
        return None

    pages = dict(split_marked_pages(full_text))
    combined = clean_text(f"{page} {pages.get(page_number + 1, '')}")
    benefit_amount = first_regex(
        page,
        [r"\bBenefit Amount\s+(?:Employee's Principal Sum\s+)?(.{1,180}?)(?=\s+Benefit Reduction\b)"],
    )
    if re.search(r"\bEqual to the Group Life Insurance amount\b", benefit_amount, flags=re.IGNORECASE):
        benefit_amount = "Same as Group Life Insurance"

    critical_section = first_regex(
        page,
        [r"\bCritical Disease Benefit\s+(.{1,900}?)(?=\s+Serious Illness Benefit\b)"],
    )
    serious_section = first_regex(
        page,
        [r"\bSerious Illness Benefit\s+(.{1,700}?)(?=\s+Schedule of Benefits\b|$)"],
    )
    critical_percent = first_regex(critical_section, [r"\b(\d{1,3}%)\s+of\s+their\s+Benefit Amount\b"])
    critical_maximum = first_regex(critical_section, [rf"\bmaximum benefit payable.+?({MONEY_AMOUNT_PATTERN})"])
    serious_percent = first_regex(serious_section, [r"\b(\d{1,3}%)\s+of\s+their\s+Benefit Amount\b"])
    serious_maximum = first_regex(serious_section, [rf"\bmaximum benefit payable.+?({MONEY_AMOUNT_PATTERN})"])
    serious_illnesses = first_regex(
        serious_section,
        [r"\bcovered serious illnesses are\s+(.{1,180}?)(?=\s+Schedule of Benefits\b|$)"],
    )
    termination = first_regex(combined, [r"\bTermination Age\s+(.{1,180}?)(?=$)"])

    return ScheduleAsiDetails(
        benefit_amount=clean_text(benefit_amount),
        critical_disease_benefit=clean_text(
            f"{critical_percent} of ASI Benefit Amount; maximum {clean_money(critical_maximum)}"
        ),
        serious_illness_benefit=clean_text(
            f"{serious_percent} of ASI Benefit Amount; maximum {clean_money(serious_maximum)}"
            + (f"; {serious_illnesses}" if serious_illnesses else "")
        ),
        termination_age=normalize_life_termination_age(termination),
        source_page=page_number,
    )


def build_schedule_critical_illness_summary_lines(full_text: str) -> list[str]:
    details = extract_schedule_critical_illness_details(full_text)
    if not details:
        return []
    return compact_detail_lines(
        [
            f"Provider: {details.provider}",
            f"Policy: {details.policy_number}",
            f"Plan Type: {details.plan_type}",
            f"Benefit Amount: {details.benefit_amount}",
            f"Guaranteed Issue Limit: {details.guaranteed_issue_limit}",
            f"Covered Illnesses: {details.insured_illnesses}",
            f"Tax Status: {details.tax_status}",
            f"Termination Age: {details.termination_age}",
        ]
    )


def extract_schedule_critical_illness_details(
    full_text: str,
) -> ScheduleCriticalIllnessDetails | None:
    page_number, page = find_schedule_page(
        full_text,
        "Group Critical Illness Insurance",
        "Guaranteed Issue Limit",
    )
    if not page:
        return None

    provider = first_regex(
        full_text,
        [r"(SSQ\s*,?\s*Life Insurance Company Inc\.?)\s+Group Critical Illness\s+[A-Za-z0-9-]+"],
    )
    policy_number = first_regex(
        full_text,
        [r"SSQ\s*,?\s*Life Insurance Company Inc\.?\s+Group Critical Illness\s+([A-Za-z0-9-]+)"],
    )
    tax_status = first_regex(
        page,
        [r"\bTax Status\s+(.{1,100}?)(?=\s+Re-Entry Exclusions\b)"],
    )
    return ScheduleCriticalIllnessDetails(
        plan_type=first_regex(page, [r"\b(Employee Only Plan)\b"]),
        benefit_amount=clean_money(first_regex(page, [rf"\bBenefit Amount\s+({MONEY_AMOUNT_PATTERN})"])),
        guaranteed_issue_limit=clean_money(
            first_regex(page, [rf"\bGuaranteed Issue Limit\s+({MONEY_AMOUNT_PATTERN})"])
        ),
        insured_illnesses=first_regex(page, [r"\b(\d+\s+Insured Illnesses)\b"]),
        tax_status="Non-taxable" if re.search(r"\bnon-taxable\b", tax_status, flags=re.IGNORECASE) else clean_text(tax_status),
        termination_age=normalize_life_termination_age(
            first_regex(page, [r"\bTermination Age\s+(.{1,180}?)(?=$)"])
        ),
        provider=clean_text(provider),
        policy_number=format_policy_number(policy_number),
        source_page=page_number,
    )


def build_schedule_ltd_summary_lines(full_text: str) -> list[str]:
    details = extract_schedule_ltd_details(full_text)
    if details:
        return format_schedule_ltd_summary_lines(details)
    return ["Coverage: Not covered"] if is_grouphealth_booklet(full_text) else []


def build_schedule_std_summary_lines(full_text: str) -> list[str]:
    details = extract_schedule_std_details(full_text)
    if details:
        return format_schedule_std_summary_lines(details)
    return ["Coverage: Not covered"] if is_grouphealth_booklet(full_text) else []


def format_schedule_std_summary_lines(details: ScheduleStdDetails) -> list[str]:
    return compact_detail_lines(
        [
            f"Benefit Schedule: {details.benefit_schedule}",
            f"Maximum: {details.maximum_benefit}",
            f"Non-Evidence Maximum: {details.non_evidence_maximum}",
            f"Tax Status: {details.tax_status}",
            f"Qualification Period: {details.qualification_period}",
            f"Definition of Disability: {details.disability_definition}",
            f"Benefit Period: {details.benefit_period}",
            f"Termination Age: {details.termination_age}",
        ]
    )


def extract_schedule_std_details(full_text: str) -> ScheduleStdDetails | None:
    page_number = 0
    page = ""
    for heading in ("Short-Term Disability", "Short Term Disability", "Weekly Indemnity"):
        for anchor in ("Maximum Benefit Period", "Elimination Period", "Qualifying Period", "Benefit Amount"):
            page_number, page = find_schedule_page(full_text, heading, anchor)
            if page:
                break
        if page:
            break
    if not page:
        return None

    schedule_heading = re.search(
        r"\b(?:Short[ -]Term Disability|Weekly Indemnity)(?:\s*\(STD\))?\b.{0,260}?"
        r"\b(?:Benefit Amount|Maximum(?: Weekly)? Benefit|Elimination Period|Qualifying Period|Benefit Period)\b",
        page,
        flags=re.IGNORECASE,
    )
    if not schedule_heading:
        return None

    benefit_amount = first_regex(
        page,
        [r"\bBenefit Amount\s+(.{1,220}?)(?=\s+(?:Maximum(?: Weekly)? Benefit|Elimination Period|Qualifying Period|Benefit Period)\b)"],
    )
    max_with = first_regex(
        page,
        [
            rf"\bMaximum Benefit\s+with Evidence of Insurability\s+({MONEY_AMOUNT_PATTERN}(?:\s+per\s+week)?)",
            rf"\bMaximum Weekly Benefit\s+({MONEY_AMOUNT_PATTERN}(?:\s+per\s+week)?)",
            rf"\bMaximum Benefit\s+({MONEY_AMOUNT_PATTERN}(?:\s+per\s+week)?)",
        ],
    )
    max_without = first_regex(
        page,
        [
            rf"\bMaximum Benefit\s+without Evidence of Insurability\s+({MONEY_AMOUNT_PATTERN}(?:\s+per\s+week)?)",
            rf"\bwithout Evidence of Insurability\s+({MONEY_AMOUNT_PATTERN}(?:\s+per\s+week)?)",
        ],
    )
    elimination_periods = unique_regex_group_matches(
        page,
        [
            r"\b(\d+\s+day(?:\(s\)|s)?\s+for\s+(?:Accident|Hospitalization|Sickness))",
            r"\b(?:Elimination|Qualifying) Period\s+(\d+\s+(?:calendar\s+)?days?)",
        ],
    )
    detailed_elimination_periods = [
        period
        for period in elimination_periods
        if re.search(r"\bfor\s+(?:Accident|Hospitalization|Sickness)\b", period, flags=re.IGNORECASE)
    ]
    if detailed_elimination_periods:
        elimination_periods = detailed_elimination_periods
    benefit_period = first_regex(
        page,
        [r"\b(?:Maximum )?Benefit Period\s+(.{1,100}?)(?=\s+Definition of Total Disability\b|\s+Taxability of Benefits\b|\s+Termination Age\b)"],
    )
    disability_definition = first_regex(
        page,
        [r"\bDefinition of Total Disability\s+(.{1,220}?)(?=\s+(?:Taxability of Benefits|Claims Review|Rehabilitation|Termination Age|$))"],
    )
    tax_status = first_regex(
        page,
        [r"\bTaxability of Benefits\s+(.{1,80}?)(?=\s+(?:Cost of Living Adjustment|Indexation|Work Re-entry|Survivor Benefit|Waiver of Premium|Termination Age|$))"],
    )
    termination = first_regex(page, [r"\bTermination Age\s+(.{1,180}?)(?=$|\s+Schedule of Benefits)"])

    meaningful_values = [benefit_amount, max_with, max_without, *elimination_periods, benefit_period, disability_definition, tax_status, termination]
    if not any(clean_text(value) for value in meaningful_values):
        return None

    return ScheduleStdDetails(
        benefit_schedule=clean_text(benefit_amount),
        maximum_benefit=normalize_money_phrase(max_with),
        non_evidence_maximum=normalize_money_phrase(max_without),
        tax_status=normalize_ltd_tax_status(tax_status),
        qualification_period=normalize_ltd_qualification_period(elimination_periods),
        disability_definition=normalize_ltd_disability_definition(disability_definition),
        benefit_period=clean_text(benefit_period),
        termination_age=normalize_life_termination_age(termination),
        source_page=page_number,
    )


def format_schedule_ltd_summary_lines(details: ScheduleLtdDetails) -> list[str]:
    return compact_detail_lines(
        [
            f"Benefit Schedule: {details.benefit_schedule}",
            f"Maximum: {details.maximum_benefit}",
            f"Non-Evidence Maximum: {details.non_evidence_maximum}",
            f"Tax Status: {details.tax_status}",
            f"Qualification Period: {details.qualification_period}",
            f"Definition of Disability: {details.disability_definition}",
            f"Benefit Period: {details.benefit_period}",
            f"Termination Age: {details.termination_age}",
        ]
    )


def extract_schedule_ltd_details(full_text: str) -> ScheduleLtdDetails | None:
    page_number, page = find_schedule_page(full_text, "Long-Term Disability", "Maximum From All Sources")
    if not page:
        page_number, page = find_schedule_page(full_text, "Long Term Disability", "Maximum From All Sources")
    if not page:
        page_number, page = find_schedule_page(full_text, "Long-Term Disability", "Maximum Benefit Period")
    if not page:
        page_number, page = find_schedule_page(full_text, "Long Term Disability", "Maximum Benefit Period")
    if not page:
        return None

    benefit_amount = first_regex(page, [r"\bBenefit Amount\s+(.{1,220}?)(?=\s+Maximum Benefit\b)"])
    max_with = first_regex(
        page,
        [
            rf"\bMaximum Benefit\s+with Evidence of Insurability\s+({MONEY_AMOUNT_PATTERN}\s+per\s+month)",
            rf"\bwith Evidence of Insurability\s+({MONEY_AMOUNT_PATTERN}\s+per\s+month)",
        ],
    )
    max_without = first_regex(
        page,
        [
            rf"\bMaximum Benefit\s+without Evidence of Insurability\s+({MONEY_AMOUNT_PATTERN}\s+per\s+month)",
            rf"\bwithout Evidence of Insurability\s+({MONEY_AMOUNT_PATTERN}\s+per\s+month)",
        ],
    )
    elimination_periods = unique_regex_group_matches(
        page,
        [
            r"\b(\d+\s+day(?:\(s\)|s)?\s+for\s+(?:Accident|Hospitalization|Sickness))",
            r"\bElimination Period\s+(\d+\s+months?)",
        ],
    )
    benefit_period = first_regex(page, [r"\bMaximum Benefit Period\s+(.{1,90}?)(?=\s+Definition of Total Disability\b)"])
    disability_definition = first_regex(page, [r"\bDefinition of Total Disability\s+(.{1,220}?)(?=\s+(?:Taxability of Benefits|Claims Review|Rehabilitation|$))"])
    tax_status = first_regex(
        page,
        [r"\bTaxability of Benefits\s+(.{1,80}?)(?=\s+(?:Cost of Living Adjustment|Indexation|Work Re-entry|Survivor Benefit|Waiver of Premium|$))"],
    )
    termination = first_regex(page, [r"\bTermination Age\s+(.{1,180}?)(?=$|\s+Schedule of Benefits)"])

    return ScheduleLtdDetails(
        benefit_schedule=clean_text(benefit_amount),
        maximum_benefit=normalize_money_phrase(max_with),
        non_evidence_maximum=normalize_money_phrase(max_without),
        tax_status=normalize_ltd_tax_status(tax_status),
        qualification_period=normalize_ltd_qualification_period(elimination_periods),
        disability_definition=normalize_ltd_disability_definition(disability_definition),
        benefit_period=clean_text(benefit_period),
        termination_age=normalize_life_termination_age(termination),
        source_page=page_number,
    )


def normalize_ltd_qualification_period(periods: list[str]) -> str:
    normalized = [
        re.sub(r"\bday\(s\)", "days", clean_text(period), flags=re.IGNORECASE)
        for period in periods
        if clean_text(period)
    ]
    day_counts = [
        match.group("days")
        for period in normalized
        if (match := re.match(r"^(?P<days>\d+)\s+days?\s+for\s+", period, flags=re.IGNORECASE))
    ]
    if day_counts and len(day_counts) == len(normalized) and len(set(day_counts)) == 1:
        return f"{day_counts[0]} days"
    return "; ".join(normalized)


def normalize_ltd_tax_status(value: str) -> str:
    cleaned = clean_text(value)
    if re.search(r"\bnon[- ]taxable\b|\bnot\s+taxable\b", cleaned, flags=re.IGNORECASE):
        return "Non-taxable"
    if re.search(r"\btaxable\b", cleaned, flags=re.IGNORECASE):
        return "Taxable"
    return cleaned


def normalize_ltd_disability_definition(value: str) -> str:
    cleaned = clean_text(value)
    own_occupation = re.search(r"\b(?P<years>\d+)\s+years?\s+own occupation\b", cleaned, flags=re.IGNORECASE)
    if own_occupation:
        years = int(own_occupation.group("years"))
        unit = "year" if years == 1 else "years"
        return f"{years} {unit} own occupation"
    return cleaned


def build_schedule_prescription_drug_summary_lines(full_text: str) -> list[str]:
    details = extract_schedule_ehc_details(full_text)
    if not details or not details.prescription_drugs_listed:
        return []

    drug_type_label = (
        "Generic Substitution"
        if details.prescription_drug_type.startswith("Covered at the lowest-cost generic")
        else "Drug Type"
    )
    lines = compact_detail_lines(
        [
            f"Coinsurance: {details.prescription_drug_coinsurance}",
            f"Deductible: {details.deductible}",
            f"Plan Type: {details.prescription_drug_plan_type}",
            f"Maximum: {details.prescription_drug_maximum}",
            f"{drug_type_label}: {details.prescription_drug_type}",
        ]
    )
    lines.extend(
        f"Elective Medications: {value}"
        for value in extract_schedule_elective_medication_lines(full_text)
    )
    return lines


def build_schedule_ehc_summary_lines(full_text: str) -> list[str]:
    details = extract_schedule_ehc_details(full_text)
    if not details or not details.other_medical_listed:
        return []

    return compact_detail_lines(
        [
            f"Coinsurance: {details.other_medical_coinsurance}",
            f"Deductible: {details.deductible}",
            f"Hearing Aids: {details.hearing_aid_benefit}",
            f"Orthotics: {details.orthotics_benefit}",
        ]
    )


def build_schedule_paramedical_summary_lines(full_text: str) -> list[str]:
    details = extract_schedule_ehc_details(full_text)
    if not details or not details.health_care_practitioners_listed:
        return []

    lines = compact_detail_lines(
        [
            f"Coinsurance: {details.health_care_practitioners_coinsurance}",
            f"Deductible: {details.deductible}",
            f"Maximum: {details.health_care_practitioners_maximum}",
        ]
    )
    if details.health_care_practitioners:
        lines.append("Practitioners:")
        lines.extend(f"- {practitioner}" for practitioner in details.health_care_practitioners)
    return lines


def build_schedule_vision_summary_lines(full_text: str) -> list[str]:
    details = extract_schedule_ehc_details(full_text)
    if not details or not details.vision_listed:
        return []

    return compact_detail_lines(
        [
            f"Coverage: {details.vision_coverage_status}",
            f"Coinsurance: {details.vision_coinsurance}",
            f"Deductible: {details.deductible}",
            f"Maximum: {details.vision_maximum}",
            f"Eye Exams: {details.eye_examination_benefit}",
        ]
    )


def build_schedule_travel_summary_lines(full_text: str) -> list[str]:
    details = extract_schedule_ehc_details(full_text)
    if not details or not details.travel_listed:
        return []

    return compact_detail_lines(
        [
            f"Coinsurance: {details.travel_coinsurance}",
            f"Out-of-Province Medical Referral: {details.travel_referral_maximum}",
            f"Emergency Travel: {details.travel_maximum}",
            f"Trip Duration: {details.travel_duration}",
            f"Travel Cancellation: {details.travel_cancellation_maximum}",
            f"Termination Age: {details.travel_termination_age}",
        ]
    )


def extract_schedule_ehc_details(full_text: str) -> ScheduleEhcDetails | None:
    page_number, page = find_schedule_page(full_text, "Extended Health Care", "% Reimbursement")
    if not page:
        return None

    prescription_drugs_listed = bool(
        re.search(r"\bPrescription Drugs?\b", page, flags=re.IGNORECASE)
    )
    other_medical_listed = bool(
        re.search(r"\bOther Medical(?: Expenses?)?\b", page, flags=re.IGNORECASE)
    )
    health_care_practitioners_listed = bool(
        re.search(r"\bHealth Care Practitioners\b", page, flags=re.IGNORECASE)
    )
    vision_listed = bool(re.search(r"\bVision Care\b", page, flags=re.IGNORECASE))
    travel_listed = bool(
        re.search(r"\bTravel Insurance and Assistance\b", page, flags=re.IGNORECASE)
    )
    if not any(
        (
            prescription_drugs_listed,
            other_medical_listed,
            health_care_practitioners_listed,
            vision_listed,
            travel_listed,
        )
    ):
        return None

    flex_rows = extract_flex_ehc_option_rows(page) if is_flex_plan_booklet(full_text) else []
    flex_drugs = next((row for row in flex_rows if row[0] == "Prescription Drug"), None)
    flex_other_medical = next((row for row in flex_rows if row[0] == "Other Medical Expenses"), None)
    flex_paramedical = next((row for row in flex_rows if row[0] == "Health Care Practitioners"), None)
    flex_vision = next((row for row in flex_rows if row[0] == "Vision Care"), None)
    travel_details_source_page, travel_details_page = find_schedule_page(
        full_text,
        "Travel Insurance and Assistance",
        "per trip",
    )
    travel_cancellation_coinsurance = extract_schedule_reimbursement_percent(
        page,
        "Travel Cancellation Insurance",
    )
    travel_cancellation_amount = first_regex(
        travel_details_page,
        [
            rf"\bTravel Cancellation Insurance\s+({MONEY_AMOUNT_PATTERN})\s+per\s+Insured\s+Person\s+per\s+trip\b"
        ],
    )
    travel_cancellation_maximum = (
        f"{clean_money(travel_cancellation_amount)} per insured person per trip"
        if travel_cancellation_amount
        else ""
    )
    travel_referral_amount = first_regex(
        travel_details_page,
        [rf"\bNon-Emergency Treatment\s+({MONEY_AMOUNT_PATTERN})\s+per\s+calendar\s+year\b"],
    )
    travel_maximum_amount = first_regex(
        travel_details_page,
        [
            rf"\bTravel Insurance and Assistance\s+({MONEY_AMOUNT_PATTERN})\s+per\s+Insured\s+Person\s+per\s+trip\b"
        ],
    )
    travel_duration = first_regex(
        travel_details_page,
        [
            r"\bmaximum duration of coverage:\s*(up to\s+\d+\s+days\s+per\s+trip(?:\s*\(\d+\s+days\s+for\s+retirees\))?)"
        ],
    )

    return ScheduleEhcDetails(
        deductible=extract_schedule_ehc_deductible(page),
        prescription_drugs_listed=prescription_drugs_listed,
        prescription_drug_coinsurance=format_schedule_option_percentages(flex_drugs)
        if flex_drugs
        else extract_schedule_reimbursement_percent(page, "Prescription Drugs"),
        prescription_drug_plan_type=first_regex(
            page,
            [r"\b(Pay Direct Drug Card)\b", r"\b(Direct Drug Card)\b"],
        ),
        prescription_drug_maximum=extract_drug_maximum(page),
        prescription_drug_type=extract_drug_type(page),
        other_medical_listed=other_medical_listed,
        other_medical_coinsurance=format_schedule_option_percentages(flex_other_medical)
        if flex_other_medical
        else (
            extract_schedule_reimbursement_percent(page, "Other Medical Expenses")
            or extract_schedule_reimbursement_percent(page, "Other Medical")
        ),
        health_care_practitioners_listed=health_care_practitioners_listed,
        health_care_practitioners_coinsurance=format_schedule_option_percentages(flex_paramedical)
        if flex_paramedical
        else extract_schedule_reimbursement_percent(page, "Health Care Practitioners"),
        health_care_practitioners_maximum=extract_schedule_paramedical_maximum(full_text),
        health_care_practitioners=extract_schedule_health_care_practitioners(full_text),
        vision_listed=vision_listed,
        vision_coinsurance=format_schedule_option_percentages(flex_vision)
        if flex_vision
        else extract_schedule_reimbursement_percent(page, "Vision Care"),
        vision_coverage_status=(
            "Not included"
            if re.search(r"\bVision Care\b[\s.]+Not Included\b", page, flags=re.IGNORECASE)
            else "Included" if vision_listed else ""
        ),
        vision_maximum=extract_schedule_vision_maximum(full_text),
        eye_examination_benefit=extract_schedule_eye_examination_benefit(full_text),
        hearing_aid_benefit=extract_schedule_hearing_aid_benefit(full_text),
        orthotics_benefit=extract_schedule_orthotics_benefit(full_text),
        travel_listed=travel_listed,
        travel_coinsurance=extract_schedule_reimbursement_percent(
            page,
            "Travel Insurance and Assistance",
        ),
        travel_referral_maximum=(
            f"{clean_money(travel_referral_amount)} per calendar year"
            if travel_referral_amount
            else ""
        ),
        travel_maximum=(
            f"{clean_money(travel_maximum_amount)} per insured person per trip"
            if travel_maximum_amount
            else first_regex(
                travel_details_page,
                [r"\bTravel Insurance and Assistance\s+(No maximum)\b"],
            )
        ),
        travel_duration=clean_text(travel_duration),
        travel_termination_age=extract_schedule_termination_age_after(
            full_text,
            travel_details_source_page,
        ),
        travel_details_source_page=travel_details_source_page,
        travel_cancellation_listed=bool(
            travel_cancellation_coinsurance or travel_cancellation_maximum
        ),
        travel_cancellation_coinsurance=travel_cancellation_coinsurance,
        travel_cancellation_maximum=travel_cancellation_maximum,
        source_page=page_number,
    )


def extract_schedule_paramedical_maximum(full_text: str) -> str:
    _page_number, page = find_schedule_page(
        full_text,
        "Health Care Practitioners",
        "Acupuncturist",
    )
    return extract_paramedical_maximums(page) if page else ""


def extract_schedule_health_care_practitioners(full_text: str) -> tuple[str, ...]:
    _page_number, page = find_schedule_page(
        full_text,
        "Health Care Practitioners",
        "Acupuncturist",
    )
    if not page:
        return ()
    practitioners: list[str] = []
    for practitioner in split_semicolon_values(extract_paramedical_practitioners(page)):
        if practitioner.startswith("Psychologist/Social Worker/Registered Clinical Counselor"):
            practitioners.extend(
                [
                    "Psychologist",
                    "Social Worker",
                    "Registered Clinical Counselor",
                    "Psychotherapist",
                ]
            )
        else:
            practitioners.append(practitioner)
    return tuple(unique_in_order(practitioners))


def extract_schedule_eye_examination_benefit(full_text: str) -> str:
    _page_number, page = find_schedule_page(full_text, "Eye Examinations", "consecutive months")
    if not page:
        return ""
    amount = first_regex(page, [rf"\bMaximum of\s+({MONEY_AMOUNT_PATTERN})\b"])
    months = first_regex(page, [r"\b(\d+)\s+consecutive\s+months\b"])
    return f"{clean_money(amount)} every {months} consecutive months" if amount and months else ""


def extract_schedule_hearing_aid_benefit(full_text: str) -> str:
    _page_number, page = find_schedule_page(full_text, "Hearing Aids", "once every")
    if not page:
        return ""
    match = re.search(
        rf"\bHearing Aids\*?\s+Once every\s+(?P<months>\d+)\s+months[^$]{{0,20}}(?P<amount>{MONEY_AMOUNT_PATTERN})",
        page,
        flags=re.IGNORECASE,
    )
    return (
        f"{clean_money(match.group('amount'))} once every {match.group('months')} months"
        if match
        else ""
    )


def extract_schedule_orthotics_benefit(full_text: str) -> str:
    _page_number, page = find_schedule_page(full_text, "Orthotics", "per calendar year")
    if not page:
        return ""
    amount = first_regex(
        page,
        [rf"\bOrthotics\*?.{{0,180}}?({MONEY_AMOUNT_PATTERN})\s+per\s+calendar\s+year\b"],
    )
    return f"{clean_money(amount)} per calendar year" if amount else ""


def extract_schedule_termination_age_after(full_text: str, source_page: int) -> str:
    if not source_page:
        return ""
    pages = dict(split_marked_pages(full_text))
    for page_number in range(source_page, source_page + 3):
        page = clean_text(pages.get(page_number, ""))
        termination = first_regex(page, [r"\bTermination Age\s+(.{1,180}?)(?=$)"])
        if termination:
            return normalize_life_termination_age(termination)
    return ""


def extract_schedule_vision_maximum(full_text: str) -> str:
    _page_number, page = find_schedule_page(
        full_text,
        "Vision Care",
        "Adults",
        "Dependent Children",
    )
    if not page:
        return ""

    adult = extract_schedule_vision_maximum_for_group(page, "Adults")
    dependent = extract_schedule_vision_maximum_for_group(page, "Dependent Children")
    values = compact_detail_lines(
        [
            f"Adults: {adult}",
            f"Dependent Children: {dependent}",
        ]
    )
    return "; ".join(values)


def extract_schedule_vision_maximum_for_group(page: str, group_label: str) -> str:
    match = re.search(
        rf"\b{re.escape(group_label)}\b.{{0,120}}?\bmaximum\s+"
        rf"(?P<amount>{MONEY_AMOUNT_PATTERN})\s+every\s+(?P<months>\d+)"
        r".{0,100}?\bconsecutive\s+months\b",
        page,
        flags=re.IGNORECASE,
    )
    if not match:
        return ""
    return f"{clean_money(match.group('amount'))} every {match.group('months')} consecutive months"


def extract_schedule_ehc_deductible(page: str) -> str:
    return first_regex(
        page,
        [
            r"\bCalendar Year Deductible\s+(.{1,90}?)(?=\s+(?:Deductible Does Not Apply To\b|%\s*Reimbursement\b|Allowable Expenses\b|Prescription Drugs?\b))"
        ],
    )


def extract_schedule_reimbursement_percent(page: str, label: str) -> str:
    label_pattern = re.escape(label)
    if label == "Prescription Drugs":
        label_pattern = r"Prescription Drugs?"
    elif label in {"Other Medical Expenses", "Other Medical"}:
        label_pattern = r"Other Medical(?: Expenses?)?"
    pattern = rf"\*?\s*{label_pattern}(?:\s*\([^)]{{1,80}}\))?\s*\**\s*[\s.]*?(\d{{2,3}}%)"
    return first_regex(page, [pattern])


def format_schedule_option_percentages(row: tuple[str, str, str, str] | None) -> str:
    if not row:
        return ""
    return "; ".join(
        f"Option {index}: {value.rstrip('*')}"
        for index, value in enumerate(row[1:], start=1)
    )


def extract_drug_maximum(text: str) -> str:
    full_text = clean_text(text)
    source_specific_value = first_regex(
        full_text,
        [
            rf"\bPrescription Drugs? are limited to a maximum of\s+({MONEY_AMOUNT_PATTERN}\s+per\s+insured\s+person\s+per\s+calendar\s+year)"
        ],
    )
    if source_specific_value:
        return normalize_money_phrase(source_specific_value)

    section = isolate_prescription_drug_text(text)
    if not section:
        section = clean_text(text)
    if not section:
        return ""

    value = first_regex(
        section,
        [
            rf"\b(?:Annual\s+)?(?:Prescription\s+)?Drug\s+Maximum\s*[:\-]?\s+(Unlimited|{MONEY_AMOUNT_PATTERN}(?:\s+per\s+(?:calendar\s+year|benefit\s+period|certificate|insured))?)",
            rf"\bAnnual\s+Maximum\s*[:\-]?\s+(Unlimited|{MONEY_AMOUNT_PATTERN}(?:\s+per\s+(?:calendar\s+year|benefit\s+period|certificate|insured))?)",
            rf"\bPrescription\s+Drugs?\s+(?:(?!fertility|smoking).){{0,40}}?\bmaximum(?:\s+of)?\s+(Unlimited|{MONEY_AMOUNT_PATTERN}(?:\s+per\s+(?:calendar\s+year|benefit\s+period|certificate|insured))?)",
            rf"\b({MONEY_AMOUNT_PATTERN})\s+(?:annual\s+)?(?:drug\s+)?maximum\b",
            r"\b(Unlimited)\b",
        ],
    )
    return normalize_money_phrase(value) if value and "$" in value else clean_flex_detail_value(value)


def extract_schedule_elective_medication_lines(text: str) -> list[str]:
    schedule_page = find_schedule_page_text(text, "Extended Health Care", "% Reimbursement") or clean_text(text)
    headings = list(
        re.finditer(
            r"\bPrescription Drugs?\s*[-\u2013\u2014]\s*(?:Included|All Options)\b",
            schedule_page,
            flags=re.IGNORECASE,
        )
    )
    section = schedule_page[headings[-1].start() :] if headings else isolate_prescription_drug_text(schedule_page)
    values: list[str] = []
    specifications = (
        (r"fertility(?:\s+drugs?)?", "fertility drugs"),
        (r"anti[- ]?smoking(?:\s+drugs?)?|smoking cessation(?:\s+(?:aids?|drugs?))?", "anti-smoking prescription drugs"),
    )
    for keyword_pattern, description in specifications:
        match = re.search(
            rf"\b(?:{keyword_pattern})\b(?P<context>[^$;]{{0,80}}?)(?P<amount>{MONEY_AMOUNT_PATTERN})(?P<tail>[^$;]{{0,35}})",
            section,
            flags=re.IGNORECASE,
        )
        if not match:
            match = re.search(
                rf"(?P<amount>{MONEY_AMOUNT_PATTERN})(?P<context>[^$;]{{0,60}}?)\b(?:{keyword_pattern})\b",
                section,
                flags=re.IGNORECASE,
            )
        if not match:
            continue

        amount = clean_money(match.group("amount"))
        context = clean_text(match.group(0))
        if re.search(r"\blifetime\b", context, flags=re.IGNORECASE):
            values.append(f"{amount} lifetime maximum on {description}")
        elif re.search(r"\bper\s+(?:calendar\s+)?year\b|\bannual(?:ly)?\b", context, flags=re.IGNORECASE):
            values.append(f"{amount} per calendar year for {description}")
        else:
            values.append(f"{amount} for {description}")
    return values


def extract_drug_type(text: str) -> str:
    section = isolate_prescription_drug_text(text)
    if not section:
        section = clean_text(text)

    patterns = [
        r"\b(covers the lowest cost generic equivalent product)\b",
        r"\b(excludes drugs not authorized for payment by(?: the)? Director of(?: the)? Pharmacare Program of British Columbia)\b",
        r"\b(excludes drugs not authorized by Pharmacare)\b",
        r"\b(Mandatory Generic)\b",
        r"\b(Generic Substitution)\b",
        r"\b(Generic unless no substitution)\b",
        r"\b(BC Formulary)\b",
        r"\b(British Columbia Formulary)\b",
        r"\b(Brand(?: Name)?\s*(?:and|&)\s*Generic)\b",
        r"\b(Generic\s*(?:and|&)\s*Brand)\b",
        r"\b(Brand)\b",
        r"\b(Generic)\b",
    ]
    values: list[str] = []
    for pattern in patterns:
        values.extend(unique_regex_group_matches(section, [pattern]))

    normalized_values = unique_in_order(normalize_drug_type_value(value) for value in values)
    has_specific_value = any(value.lower() not in {"brand", "generic"} for value in normalized_values)
    if has_specific_value:
        normalized_values = [value for value in normalized_values if value.lower() not in {"brand", "generic"}]
    return "; ".join(normalized_values)


def isolate_prescription_drug_text(text: str) -> str:
    cleaned = clean_text(text)
    if not cleaned:
        return ""

    start = re.search(r"\bPrescription Drugs?\s*[-\u2013\u2014]\s*Included\b", cleaned, flags=re.IGNORECASE)
    if not start:
        start = re.search(r"\bPrescription Drugs?\b", cleaned, flags=re.IGNORECASE)
    if start:
        cleaned = cleaned[start.start() :]

    end = re.search(
        r"\b(?:Hospital\s*[-\u2013\u2014]\s*Included|Health Care Practitioners|Paramedical|Vision Care|Dental Care|Emergency Travel)\b",
        cleaned,
        flags=re.IGNORECASE,
    )
    if end:
        cleaned = cleaned[: end.start()]
    return cleaned


def normalize_drug_type_value(value: str) -> str:
    cleaned = clean_flex_detail_value(value)
    if re.search(r"\bcovers the lowest cost generic equivalent product\b", cleaned, flags=re.IGNORECASE):
        return "Covered at the lowest-cost generic equivalent product"
    if re.search(r"excludes drugs not authorized.*Pharmacare", cleaned, flags=re.IGNORECASE):
        return "Excludes drugs not authorized by Pharmacare"
    return cleaned


def extract_summary_field_value(
    section: SummarySectionDefinition,
    field: SummaryFieldDefinition,
    source_text: str,
) -> str:
    label = field.label.lower()
    if section.name == "Paramedicals" and label in {"maximum", "maximums", "benefit amount"}:
        return extract_paramedical_maximums(source_text) or extract_by_json_patterns(
            source_text, field.patterns, field.max_results
        )
    if section.name == "Paramedicals" and label == "practitioners":
        return extract_paramedical_practitioners(source_text) or extract_by_json_patterns(
            source_text, field.patterns, field.max_results
        )
    if section.name == "Prescription Drugs" and label in {"drug max", "drug maximum", "annual drug max", "annual drug maximum"}:
        return extract_drug_maximum(source_text) or extract_by_json_patterns(
            source_text, field.patterns, field.max_results
        )
    if section.name == "Prescription Drugs" and label == "drug type":
        return extract_drug_type(source_text) or extract_by_json_patterns(
            source_text, field.patterns, field.max_results
        )
    if section.name == "Dental Care" and label in {"maximum", "maximums"}:
        return extract_dental_maximums(source_text) or extract_by_json_patterns(
            source_text, field.patterns, field.max_results
        )
    if section.name == "Dental Care" and label == "orthodontics":
        return extract_orthodontic_coverage(source_text) or extract_by_json_patterns(
            source_text, field.patterns, field.max_results
        )
    if section.name == "Dental Care" and label == "recall":
        return extract_recall_frequency(source_text) or extract_by_json_patterns(
            source_text, field.patterns, field.max_results
        )
    return extract_by_json_patterns(source_text, field.patterns, field.max_results)


def detect_class_name(text: str) -> str:
    cover_class_name = detect_cover_class_name(text)
    if cover_class_name:
        return cover_class_name

    policy_cover_class_name = detect_policy_cover_class_name(text)
    if policy_cover_class_name:
        return policy_cover_class_name

    class_name = first_regex(
        text[:3000],
        [
            r"\bClass\s*[:\-]\s*([^\n\r]{3,90})",
            r"\bClass\s+([A-Za-z0-9][^\n\r]{3,90})",
        ],
    )
    return clean_text(class_name)


def detect_policy_cover_class_name(text: str) -> str:
    """Read the class label printed directly below a policy number on a booklet cover."""
    lines = [clean_text(line) for line in text[:3000].splitlines()]
    lines = [line for line in lines if line and not re.fullmatch(r"\[\[PAGE\s+\d+\]\]", line)]

    for index, line in enumerate(lines[:40]):
        if not re.fullmatch(r"Policy\s*(?:Number|No\.?|#)\s*[:\-]?\s*[A-Za-z0-9\-]+", line, flags=re.IGNORECASE):
            continue

        for candidate in lines[index + 1 : index + 3]:
            if re.search(r"^(?:Your Group Insurance Plan|Benefits Booklet|Table of Contents)$", candidate, flags=re.IGNORECASE):
                break
            if len(candidate) >= 3 and len(candidate) <= 90 and any(character.isalpha() for character in candidate):
                return clean_class_name(candidate)

    return ""


def detect_cover_class_name(text: str) -> str:
    lines = [clean_text(line) for line in text[:6000].splitlines()]
    lines = [line for line in lines if line]

    for index, line in enumerate(lines[:120]):
        if not re.fullmatch(r"Employee\s+(?:Class|Classification)", line, flags=re.IGNORECASE):
            continue

        class_lines: list[str] = []
        for candidate in lines[index + 1 : index + 10]:
            if re.search(
                r"^(?:\[\[PAGE\s+\d+\]\]|For Participants(?:\s+of)?|Your Benefits Handbook|Plan Effective Date|Effective Date|Group Number\(s\)|Group Name and Number\(s\)|CSBT|Community Services Benefits Trust)",
                candidate,
                flags=re.IGNORECASE,
            ):
                break
            if re.search(r"\b(?:Plan Effective Date|Effective Date of this Document)\b", candidate, flags=re.IGNORECASE):
                break
            class_lines.append(candidate)

        class_name = clean_class_name(" ".join(class_lines))
        if class_name:
            return class_name

    return ""


def clean_class_name(value: str) -> str:
    cleaned = clean_text(value)
    cleaned = re.sub(r"\bClass\s+([A-Za-z0-9]+)\s+(?![-:])", r"Class \1 - ", cleaned, count=1, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+-\s+", " - ", cleaned)
    cleaned = re.sub(r"\bPart\s+-\s+time\b", "Part-time", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip(" :-")


def format_policy_number(value: str) -> str:
    cleaned = clean_text(value)
    if len(cleaned) < 3 or cleaned.lower() in {"s", "n/a", "na", "none"}:
        return ""
    match = re.fullmatch(r"G0*(\d+)", cleaned, flags=re.IGNORECASE)
    if match:
        return match.group(1)
    return cleaned


def remove_front_matter_for_benefits(text: str, rules: ExtractionRules) -> str:
    for marker in rules.front_matter_markers:
        match = re.search(marker, text, flags=re.IGNORECASE)
        if match:
            return text[match.start() :]
    return text


def build_search_lines(text: str) -> list[tuple[str, str]]:
    lines = [clean_text(line) for line in text.splitlines()]
    lines = [line for line in lines if line]

    if len(lines) <= 3:
        chunks = re.split(r"(?<=[.!?])\s+", clean_text(text))
        lines = [chunk for chunk in chunks if chunk]

    return [(line, normalize_for_search(line)) for line in lines]


def find_benefit_section(
    full_text: str,
    normalized_lines: list[tuple[str, str]],
    definition: BenefitDefinition,
) -> tuple[str, list[str]]:
    heading_section, heading_keywords = find_benefit_section_by_page_heading(full_text, definition, definition.page_span)
    if heading_section:
        return heading_section, heading_keywords

    matched_indexes: list[int] = []
    matched_keywords: list[str] = []
    normalized_keywords = [(keyword, normalize_for_search(keyword)) for keyword in definition.keywords]

    for index, (_line, normalized_line) in enumerate(normalized_lines):
        for keyword, normalized_keyword in normalized_keywords:
            if normalized_keyword and contains_normalized_term(normalized_line, normalized_keyword):
                matched_indexes.append(index)
                if keyword not in matched_keywords:
                    matched_keywords.append(keyword)

    if matched_indexes:
        selected_lines: list[str] = []
        selected_keys: set[int] = set()
        for index in matched_indexes:
            start = max(0, index - 3)
            end = min(len(normalized_lines), index + 10)
            for line_index in range(start, end):
                if line_index not in selected_keys:
                    selected_keys.add(line_index)
                    selected_lines.append(normalized_lines[line_index][0])
        return clean_text(" ".join(selected_lines)), matched_keywords

    return find_benefit_section_by_character_window(full_text, normalized_keywords)


def find_benefit_section_by_page_heading(
    full_text: str,
    definition: BenefitDefinition,
    page_span: int = 2,
) -> tuple[str, list[str]]:
    pages = split_marked_pages(full_text)
    if not pages:
        return "", []

    heading_terms = definition.headings or definition.keywords
    normalized_keywords = [(keyword, normalize_for_search(keyword)) for keyword in heading_terms]
    for page_index, (_page_number, page_text) in enumerate(pages):
        lines = [clean_text(line) for line in page_text.splitlines() if clean_text(line)]
        normalized_heading_lines = [normalize_for_search(line) for line in lines[:3]]

        matched_keywords = [
            keyword
            for keyword, normalized_keyword in normalized_keywords
            if normalized_keyword
            and any(line_matches_normalized_heading(line, normalized_keyword) for line in normalized_heading_lines)
        ]
        if not matched_keywords:
            continue

        selected_pages = []
        for selected_index in range(page_index, min(len(pages), page_index + max(1, page_span))):
            selected_pages.append(pages[selected_index][1])

        return clean_text(" ".join(selected_pages)), matched_keywords

    return "", []


def split_marked_pages(text: str) -> list[tuple[int, str]]:
    matches = list(PAGE_MARKER_PATTERN.finditer(text))
    pages: list[tuple[int, str]] = []
    if matches and clean_text(text[: matches[0].start()]):
        pages.append((0, text[: matches[0].start()]))

    for index, match in enumerate(matches):
        page_number = int(match.group(1))
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        pages.append((page_number, text[start:end]))
    if not pages and clean_text(text):
        pages.append((1, text))
    return pages


def find_benefit_section_by_character_window(
    full_text: str,
    normalized_keywords: list[tuple[str, str]],
) -> tuple[str, list[str]]:
    cleaned = clean_text(full_text)
    normalized_text = normalize_for_search(cleaned)
    windows: list[str] = []
    matched_keywords: list[str] = []

    for keyword, normalized_keyword in normalized_keywords:
        position = find_normalized_term_position(normalized_text, normalized_keyword)
        if position < 0:
            continue
        start = max(0, position - 350)
        end = min(len(cleaned), position + 1200)
        windows.append(cleaned[start:end])
        matched_keywords.append(keyword)

    return clean_text(" ".join(windows)), matched_keywords


def contains_normalized_term(normalized_text: str, normalized_term: str) -> bool:
    return find_normalized_term_position(normalized_text, normalized_term) >= 0


def line_matches_normalized_heading(normalized_line: str, normalized_heading: str) -> bool:
    return normalized_line == normalized_heading or normalized_line.startswith(f"{normalized_heading} ")


def find_normalized_term_position(normalized_text: str, normalized_term: str) -> int:
    if not normalized_term:
        return -1

    match = re.search(rf"(?<![a-z0-9]){re.escape(normalized_term)}(?![a-z0-9])", normalized_text)
    return match.start() if match else -1


def extract_fields(section_text: str, rules: ExtractionRules, benefit_name: str = "") -> dict[str, str]:
    """Extract only explicit values found near the detected benefit section."""
    field_patterns = rules.field_patterns
    fields = {
        "Coinsurance": extract_by_json_patterns(section_text, field_patterns.get("Coinsurance", ()), 4)
        or extract_coinsurance(section_text),
        "Maximum": extract_by_json_patterns(section_text, field_patterns.get("Maximum", ()), 5)
        or extract_maximum(section_text),
        "Frequency": extract_by_json_patterns(section_text, field_patterns.get("Frequency", ()), 4)
        or extract_frequency(section_text),
        "WaitingPeriod": extract_waiting_period_field(section_text, rules),
        "TerminationAge": extract_by_json_patterns(section_text, field_patterns.get("TerminationAge", ()), 3)
        or extract_termination_age(section_text),
    }

    if benefit_name in LUMP_SUM_BENEFITS:
        if benefit_name in FULL_PAYOUT_BENEFITS:
            fields["Coinsurance"] = ""
            fields["Maximum"] = extract_lump_sum_maximum(section_text) or fields["Maximum"]
        else:
            fields["Maximum"] = extract_lump_sum_benefit_amount(section_text) or fields["Maximum"]

    return fields


def extract_by_json_patterns(text: str, patterns: Iterable[str], max_results: int) -> str:
    values: list[str] = []
    searchable = clean_text(text)
    for pattern in patterns:
        try:
            matches = re.finditer(pattern, searchable, flags=re.IGNORECASE)
        except re.error:
            continue

        for match in matches:
            value = match.group(1) if match.lastindex else match.group(0)
            cleaned_value = clean_extracted_value(value)
            if cleaned_value:
                values.append(cleaned_value)

    return "; ".join(unique_in_order(values)[:max_results])


def clean_extracted_value(value: str) -> str:
    cleaned = clean_text(value)
    cleaned = re.sub(r"^(?:of|to a)\s+", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.strip(" :-;,")
    if cleaned.lower() in {"applies to"}:
        return ""
    return cleaned


def normalize_recall_frequency(value: str) -> str:
    cleaned = clean_text(value)
    replacements = [
        (r"\btwo\s+per\s+(?:calendar\s+)?year\b", "2 per year"),
        (r"\bone\s+per\s+(\d{1,2})\s+months?\b", r"1 per \1 months"),
        (r"\bonce\s+every\s+(\d{1,2})\s+months?\b", r"once every \1 months"),
    ]
    for pattern, replacement in replacements:
        cleaned = re.sub(pattern, replacement, cleaned, flags=re.IGNORECASE)
    return cleaned


def extract_recall_frequency(text: str) -> str:
    cleaned = clean_text(text)
    if not cleaned:
        return ""

    recall_line = first_regex(
        cleaned,
        [
            r"\b(recall or periodic oral examination:\s*[^.;]+)",
            r"\b(recall exams?:\s*[^.;]+)",
        ],
    )
    if recall_line:
        recall_line = re.sub(r"^recall(?:\s+or\s+periodic\s+oral\s+examination| exams?)\s*:\s*", "", recall_line, flags=re.IGNORECASE)
        recall_line = trim_recall_frequency_detail(recall_line)
        return normalize_recall_frequency(recall_line)

    return normalize_recall_frequency(
        extract_by_json_patterns(
            cleaned,
            [
                r"\b(twice\s+per\s+(?:calendar\s+)?year)",
                r"\b(2\s+per\s+year)",
                r"\b(1\s+per\s+(?:5|6|9|12)\s+months?)",
                r"\b(one\s+per\s+(?:5|6|9|12)\s+months?)",
                r"\b(1\s+every\s+(?:5|6|9|12)\s+months?)",
                r"\b(once\s+every\s+(?:5|6|9|12)\s+months?)",
            ],
            1,
        )
    )


def trim_recall_frequency_detail(value: str) -> str:
    cleaned = clean_text(value)
    stop_markers = [
        " complete oral examination",
        " emergency oral examination",
        " complete periodontal examination",
        " radiographs",
        " laboratory tests",
        " preventive services",
        " routine dental care",
    ]
    lowered = cleaned.lower()
    stop_positions = [lowered.find(marker) for marker in stop_markers if lowered.find(marker) > 0]
    if stop_positions:
        cleaned = cleaned[: min(stop_positions)]
    cleaned = re.sub(r"\s*•\s*$", "", cleaned)
    return cleaned.strip(" -;:")


def extract_orthodontic_coverage(text: str) -> str:
    cleaned = clean_text(text)
    if not cleaned:
        return ""

    orthodontic_percent = first_regex(
        cleaned,
        [
            r"\bOrthodontic Care\s*\(subject to certain limits\)\s*\.{2,}\s*(\d{2,3}%)",
            r"\bOrthodontic Care.{1,180}?(\d{2,3}%)",
            r"\bLevel V - Orthodontics.{1,300}?(\d{2,3}%)",
        ],
    )
    coverage_note = first_regex(
        cleaned,
        [
            r"\b(Orthodontics\s*[-:]\s*Adults\s*&\s*Dependent\s*Children)",
            r"\b(Orthodontics\s*[-:]\s*Adults\s+and\s+Dependent\s*Children)",
            r"\b(Adults\s*&\s*Dependent\s*Children)",
        ],
    )

    values = []
    if orthodontic_percent:
        values.append(orthodontic_percent)
    if coverage_note:
        coverage_note = re.sub(r"^Orthodontics\s*[-:]\s*", "", coverage_note, flags=re.IGNORECASE)
        values.append(coverage_note)
    return "; ".join(unique_in_order(values))


def extract_paramedical_maximums(text: str) -> str:
    section = isolate_paramedical_maximum_text(text)
    if not section:
        return ""
    section = remove_paramedical_xray_sublimits(section)

    practitioner_amounts_by_name = extract_practitioner_amounts(section)
    practitioner_amounts = list(practitioner_amounts_by_name.values())
    practitioner_amounts.extend(
        clean_money(match)
        for match in unique_regex_group_matches(section, [rf"({MONEY_AMOUNT_PATTERN})\s+per\s+practitioner\s+per\s+calendar\s+year"])
    )
    standard_amount = "$500" if "$500" in practitioner_amounts else (practitioner_amounts[0] if practitioner_amounts else "")

    psychology_group_amount = first_regex(
        section,
        [
            rf"\bPsychologist/Social Worker/Registered Clinical Counsell?or(?:/\s*Psychotherapist)?\b[^$]{{0,160}}({MONEY_AMOUNT_PATTERN})",
            rf"\bPsychologist\b[^$]{{0,220}}\b(?:Social Worker|Registered Clinical Counsell?or)\b[^$]{{0,160}}({MONEY_AMOUNT_PATTERN})",
        ],
    )
    psychology_group_amount = clean_money(psychology_group_amount) if psychology_group_amount else ""
    psychotherapist_amount = extract_standalone_psychotherapist_amount(section)

    values: list[str] = []
    if standard_amount:
        values.append(f"{standard_amount} per practitioner per calendar year")
    if psychotherapist_amount and psychotherapist_amount != standard_amount:
        values.append(f"{psychotherapist_amount} for Psychotherapist")
    if psychology_group_amount and psychology_group_amount not in {standard_amount, psychotherapist_amount}:
        values.append(f"{psychology_group_amount} for Psychologist/Social Worker/Registered Clinical Counselor")
    for practitioner_name, amount in practitioner_amounts_by_name.items():
        if amount and amount != standard_amount and not re.search(
            r"\b(?:psychologist|social worker|counsel|psychotherapist)\b",
            practitioner_name,
            flags=re.IGNORECASE,
        ):
            values.append(f"{amount} for {practitioner_name}")
    return "; ".join(unique_in_order(values))


def extract_practitioner_amounts(section: str) -> dict[str, str]:
    section = remove_paramedical_xray_sublimits(section)
    practitioner_patterns = [
        ("Acupuncturist", r"\bAcupuncturist\b"),
        ("Audiologist", r"\bAudiologist\b"),
        ("Chiropractor", r"\bChiropractor\b"),
        ("Dietician", r"\bDietician\b"),
        ("Massage Therapist/Orthotherapist", r"\bMassage Therapist/Orthotherapist\b"),
        ("Massage Therapist", r"\bMassage Therapist\b"),
        ("Naturopath", r"\bNaturopath\b"),
        ("Occupational Therapist", r"\bOccupational Therapist\b"),
        ("Osteopath", r"\bOsteopath\b"),
        ("Physiotherapist/Physical Rehabilitation Therapist", r"\bPhysiotherapist/Physical Rehabilitation Therapist\b"),
        ("Physiotherapist", r"\bPhysiotherapist\b"),
        ("Podiatrist/Chiropodist", r"\bPodiatrist/Chiropodist\b"),
        ("Speech Therapist", r"\bSpeech Therapist\b"),
        ("Psychologist/Social Worker/Registered Clinical Counselor/Psychotherapist", r"\bPsychologist/Social Worker/Registered Clinical Counsell?or(?:/| )Psychotherapist\b"),
        ("Psychologist/Social Worker/Registered Clinical Counselor", r"\bPsychologist/Social Worker/Registered Clinical Counsell?or\b"),
        ("Psychologist", r"\bPsychologist\b"),
        ("Psychotherapist", r"\bPsychotherapist\b"),
    ]
    amounts: dict[str, str] = {}
    for practitioner_name, practitioner_pattern in practitioner_patterns:
        match = re.search(
            rf"{practitioner_pattern}[^$]{{0,180}}({MONEY_AMOUNT_PATTERN})",
            section,
            flags=re.IGNORECASE,
        )
        if match:
            amounts[practitioner_name] = clean_money(match.group(1))

    remove_embedded_practitioner_amounts(amounts)
    return amounts


def remove_paramedical_xray_sublimits(section: str) -> str:
    return re.sub(
        rf"\([^)]*x-rays?[^)]*{MONEY_AMOUNT_PATTERN}[^)]*\)",
        "",
        section,
        flags=re.IGNORECASE,
    )


def remove_embedded_practitioner_amounts(amounts: dict[str, str]) -> None:
    embedded_pairs = [
        ("Massage Therapist/Orthotherapist", "Massage Therapist"),
        ("Physiotherapist/Physical Rehabilitation Therapist", "Physiotherapist"),
        (
            "Psychologist/Social Worker/Registered Clinical Counselor/Psychotherapist",
            "Psychologist/Social Worker/Registered Clinical Counselor",
        ),
        ("Psychologist/Social Worker/Registered Clinical Counselor", "Psychologist"),
        ("Psychologist/Social Worker/Registered Clinical Counselor", "Psychotherapist"),
    ]
    for combined_name, embedded_name in embedded_pairs:
        if combined_name in amounts and embedded_name in amounts:
            del amounts[embedded_name]


def extract_standalone_psychotherapist_amount(section: str) -> str:
    for match in re.finditer(rf"\bPsychotherapist\b[^$]{{0,140}}({MONEY_AMOUNT_PATTERN})", section, flags=re.IGNORECASE):
        prefix = section[max(0, match.start() - 90) : match.start()]
        if re.search(r"\bPsychologist/Social Worker/Registered Clinical Counsell?or/\s*$", prefix, flags=re.IGNORECASE):
            continue
        return clean_money(match.group(1))
    return ""


def extract_paramedical_practitioners(text: str) -> str:
    section = isolate_paramedical_maximum_text(text)
    if not section:
        return ""

    practitioner_terms = [
        ("Acupuncturist", r"\bAcupuncturist\b"),
        ("Audiologist", r"\bAudiologist\b"),
        ("Chiropractor", r"\bChiropractor\b"),
        ("Dietician", r"\bDietician\b"),
        ("Massage Therapist/Orthotherapist", r"\bMassage Therapist/Orthotherapist\b"),
        ("Massage Therapist", r"\bMassage Therapist\b"),
        ("Naturopath", r"\bNaturopath\b"),
        ("Occupational Therapist", r"\bOccupational Therapist\b"),
        ("Osteopath", r"\bOsteopath\b"),
        (
            "Physiotherapist/Physical Rehabilitation Therapist",
            r"\bPhysiotherapist/Physical Rehabilitation Therapist\b",
        ),
        ("Physiotherapist", r"\bPhysiotherapist\b"),
        ("Podiatrist/Chiropodist", r"\bPodiatrist/Chiropodist\b"),
        (
            "Psychologist/Social Worker/Registered Clinical Counselor/Psychotherapist",
            r"\bPsychologist/Social Worker/Registered Clinical Counsell?or(?:/| )Psychotherapist\b",
        ),
        (
            "Psychologist/Social Worker/Registered Clinical Counselor/Marriage and Family Therapist/Psychoanalyst/Psychotherapist",
            r"\bPsychologist/Social Worker/Registered Clinical Counsell?or/Marriage and Family Therapist/Psychoanalyst/Psychotherapist\b",
        ),
        ("Psychologist", r"\bPsychologist\b"),
        ("Social Worker", r"\bSocial Worker\b"),
        ("Registered Clinical Counselor", r"\bRegistered Clinical Counsell?or\b"),
        ("Marriage and Family Therapist", r"\bMarriage and Family Therapist\b"),
        ("Psychoanalyst", r"\bPsychoanalyst\b"),
        ("Psychotherapist", r"\bPsychotherapist\b"),
        ("Speech Therapist", r"\bSpeech Therapist\b"),
    ]

    selected: list[str] = []
    for display, pattern in practitioner_terms:
        if re.search(pattern, section, flags=re.IGNORECASE):
            selected.append(display)

    selected_set = {value.lower() for value in selected}
    if "massage therapist/orthotherapist" in selected_set:
        selected = [value for value in selected if value.lower() != "massage therapist"]
    if "physiotherapist/physical rehabilitation therapist" in selected_set:
        selected = [value for value in selected if value.lower() != "physiotherapist"]
    if (
        "psychologist/social worker/registered clinical counselor/marriage and family therapist/psychoanalyst/psychotherapist"
        in selected_set
    ):
        selected = [
            value
            for value in selected
            if value.lower()
            != "psychologist/social worker/registered clinical counselor/psychotherapist"
        ]
    if any(value.lower().startswith("psychologist/social worker/registered clinical counselor") for value in selected):
        selected = [
            value
            for value in selected
            if value.lower()
            not in {
                "psychologist",
                "social worker",
                "registered clinical counselor",
                "marriage and family therapist",
                "psychoanalyst",
                "psychotherapist",
            }
        ]

    return "; ".join(unique_in_order(selected))


def isolate_paramedical_maximum_text(text: str) -> str:
    cleaned = clean_text(text)
    if not cleaned:
        return ""

    marker = re.search(r"\bHealth Care Practitioners\s*(?:[-\u2013\u2014]\s*)?Included\b", cleaned, flags=re.IGNORECASE)
    if not marker:
        marker = re.search(r"\bHealth Care Practitioners\b", cleaned, flags=re.IGNORECASE)
    if marker:
        cleaned = cleaned[marker.start() :]

    end_marker = re.search(
        r"\b(?:Vision Care|Other Medical Expenses|Dental Care|Emergency Travel|Hospital - Included)\b",
        cleaned,
        flags=re.IGNORECASE,
    )
    if end_marker:
        cleaned = cleaned[: end_marker.start()]
    return cleaned


def extract_dental_maximums(text: str) -> str:
    cleaned = clean_text(text)
    if not cleaned:
        return ""

    values: list[str] = []
    no_maximum = first_regex(
        cleaned,
        [
            r"\bBasic Dental Care/Routine Dental Care/Major Restorative Services\b.{0,120}?\b(No maximum)\b",
            r"\b(?:Basic Dental Care|Routine Dental Care|Major Restorative Services)\b.{0,180}?\b(No maximum)\b",
        ],
    )
    if no_maximum:
        values.append(no_maximum)

    orthodontic_maximum = first_regex(
        cleaned,
        [
            rf"\bOrthodontic Care\b.{{0,140}}?({MONEY_AMOUNT_PATTERN}\s+per\s+Insured\s+per\s+lifetime)",
            rf"\bOrthodontic Care\b.{{0,180}}?({MONEY_AMOUNT_PATTERN}.{{0,80}}?\blifetime\b)",
            rf"\bOrthodontics\b.{{0,180}}?({MONEY_AMOUNT_PATTERN}.{{0,80}}?\blifetime\b)",
        ],
    )
    if orthodontic_maximum:
        values.append(normalize_money_phrase(orthodontic_maximum))

    return "; ".join(unique_in_order(values))


def normalize_money_phrase(value: str) -> str:
    return re.sub(MONEY_AMOUNT_PATTERN, lambda match: clean_money(match.group(0)), clean_text(value), count=1)


def extract_coinsurance(text: str) -> str:
    percentages = unique_in_order(re.findall(r"\b(?:50|60|70|75|80|85|90|100)\s*%", text))
    return "; ".join(percentages[:4])


def extract_lump_sum_benefit_amount(text: str) -> str:
    cleaned = clean_text(text)
    benefit_amount = first_regex(
        cleaned,
        [
            rf"\bBenefit Amount\s+({MONEY_AMOUNT_PATTERN})",
            rf"\bBenefit amount\s*-?\s*({MONEY_AMOUNT_PATTERN})",
            rf"\bPrincipal Sum\s+({MONEY_AMOUNT_PATTERN})",
            r"\bBenefit Amount\s+(.{1,180}?)(?=\s+(?:Non[- ]Evidence|Reduction|Qualifying|Waiver|Conversion|Termination|Coverage Ends|Exclusions|$))",
            r"\bBenefit amount\s*-?\s*(.{1,120}?)(?=\s+(?:Non[- ]Evidence|Termination|Waiver|Conversion|Coverage Ends|Exclusions|$))",
            r"\bPrincipal Sum\s+(.{1,120}?)(?=\s+(?:Benefit Amount|Tax Status|Re-Entry|Conversion|$))",
        ],
    )
    benefit_amount = normalize_lump_sum_benefit_amount(benefit_amount, cleaned)
    if benefit_amount:
        return benefit_amount

    maximum = first_regex(
        cleaned,
        [
            rf"\bMaximum\s+({MONEY_AMOUNT_PATTERN})",
            rf"\bto\s+a\s+maximum\s+of\s+({MONEY_AMOUNT_PATTERN})",
        ],
    )
    return clean_money(maximum) if maximum else ""


def extract_lump_sum_maximum(text: str) -> str:
    maximum = first_regex(
        clean_text(text),
        [
            rf"\bMaximum Benefit(?:\s+(?:with|without)\s+Evidence of Insurability)?\s+({MONEY_AMOUNT_PATTERN})",
            rf"\bMaximum(?:\s+Benefit)?\s*[:\-]?\s*({MONEY_AMOUNT_PATTERN})",
            rf"\bto\s+a\s+maximum\s+of\s+({MONEY_AMOUNT_PATTERN})",
        ],
    )
    return clean_money(maximum) if maximum else ""


def normalize_lump_sum_benefit_amount(value: str, source_text: str) -> str:
    cleaned = clean_extracted_value(value)
    if not cleaned:
        return ""

    flat_maximum = first_regex(
        f"{cleaned} {source_text[:300]}",
        [
            rf"\bFlat amount\s+Maximum\s+({MONEY_AMOUNT_PATTERN})",
            rf"\bFlat amount\s+({MONEY_AMOUNT_PATTERN})",
        ],
    )
    if flat_maximum:
        return clean_money(flat_maximum)

    earnings_match = re.search(
        r"\b(\d+(?:\.\d+)?)\s*(?:x|times?)\s+(?:your\s+)?(?:annual\s+)?(?:earnings|salary)\b",
        cleaned,
        flags=re.IGNORECASE,
    )
    if earnings_match:
        multiplier = earnings_match.group(1)
        if multiplier.endswith(".0"):
            multiplier = multiplier[:-2]
        maximum = first_regex(
            cleaned,
            [
                rf"\bto\s+a\s+maximum\s+of\s+({MONEY_AMOUNT_PATTERN})",
                rf"\bmaximum\s+of\s+({MONEY_AMOUNT_PATTERN})",
            ],
        )
        if maximum:
            return f"{multiplier} x annual earnings, maximum {clean_money(maximum)}"
        return f"{multiplier} x annual earnings"

    money = first_regex(cleaned, [rf"({MONEY_AMOUNT_PATTERN})"])
    if money:
        return clean_money(money)

    return cleaned


def extract_maximum(text: str) -> str:
    sentences = split_sentences(text)
    maximums: list[str] = []
    amount_pattern = re.compile(MONEY_AMOUNT_PATTERN)
    context_words = re.compile(r"\b(maximum|max|annual|lifetime|per\s+visit|per\s+person|calendar\s+year)\b", re.I)

    for sentence in sentences:
        amounts = amount_pattern.findall(sentence)
        if not amounts:
            continue

        sentence_clean = clean_text(sentence)
        if context_words.search(sentence_clean):
            maximums.append(shorten(sentence_clean, 320))
        else:
            maximums.extend(clean_money(amount) for amount in amounts)

    return "; ".join(unique_in_order(maximums)[:5])


def extract_frequency(text: str) -> str:
    patterns = [
        r"\bevery\s+(?:6|9|12|24)\s+months?\b",
        r"\bevery\s+\d+\s+calendar\s+years?\b",
        r"\bonce\s+every\s+(?:6|9|12|24)\s+months?\b",
        r"\blifetime\b",
        r"\bper\s+calendar\s+year\b",
    ]
    return "; ".join(unique_regex_matches(text, patterns)[:4])


def extract_waiting_period_field(section_text: str, rules: ExtractionRules) -> str:
    field_patterns = rules.field_patterns.get("WaitingPeriod", ())
    labeled_patterns = field_patterns[:1]
    labeled_value = extract_by_json_patterns(section_text, labeled_patterns, 1)
    if labeled_value:
        return labeled_value

    return extract_waiting_period(section_text)


def extract_waiting_period(text: str) -> str:
    patterns = [
        r"\bWaiting Period\s+(No\s+waiting\s+period|No\s+waiting|None|Nil|Date\s+of\s+hire|First\s+day\s+following\s+\d+\s+months?|\d{1,2}\s+months?)",
        r"\bwaiting\s+period\s+(?:is\s+)?(No\s+waiting\s+period|No\s+waiting|None|Nil|Date\s+of\s+hire|First\s+day\s+following\s+\d+\s+months?|\d{1,2}\s+months?)",
    ]
    return extract_by_json_patterns(text, patterns, 1)


def extract_termination_age(text: str) -> str:
    labelled = first_regex(
        clean_text(text),
        [r"\bTermination Age\s+(.{1,180}?)(?=$|\s+(?:Dependent Life|Accident|Schedule of Benefits))"],
    )
    if labelled:
        return normalize_life_termination_age(labelled)

    patterns = [
        r"\bearlier\s+of\s+age\s+(?:65|70)\s+or\s+retirement\b",
        r"\bage\s+(?:65|70)\b",
        r"\bretirement\b",
    ]
    return "; ".join(unique_regex_matches(text, patterns)[:3])


def score_confidence(section_text: str, matched_keywords: list[str], fields: dict[str, str]) -> str:
    populated_field_count = sum(1 for value in fields.values() if value)
    if populated_field_count >= 2 or (len(matched_keywords) >= 2 and len(section_text) > 220):
        return "High"
    if populated_field_count >= 1 or len(section_text) > 140:
        return "Medium"
    return "Low"


def confidence_note(confidence: str) -> str:
    if confidence == "Medium":
        return "Review recommended"
    if confidence == "Low":
        return "Low confidence extraction - review required"
    return ""


def build_coverage_description(section_text: str) -> str:
    cleaned = clean_text(section_text)
    sentences = split_sentences(cleaned)
    if sentences:
        description = " ".join(sentences[:3])
    else:
        description = cleaned
    return shorten(description, 1600)


def update_workbooks(
    master_path: Path,
    review_path: Path | None,
    rows: list[ExtractionRow],
    summaries: list[EmployeeBenefitsSummary],
    processed_sources: list[ProcessedSource],
    errors: list[ProcessingError],
) -> WorkbookUpdateResult:
    master_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = load_or_create_workbook(master_path)

    plan_sheet = ensure_sheet_with_columns(workbook, PLAN_SHEET, PLAN_COLUMNS)
    log_sheet = ensure_sheet_with_columns(workbook, LOG_SHEET, LOG_COLUMNS)
    review_sheet = ensure_sheet_with_columns(workbook, REVIEW_SHEET, REVIEW_COLUMNS, clear_existing=True)

    rows_added, rows_updated, counts_by_source = update_plan_design_sheet(plan_sheet, rows)
    write_review_sheet(review_sheet, rows, summaries)
    display_summaries = build_benefit_summary_display_summaries(summaries)
    write_employee_benefits_summary_sheet(workbook, display_summaries)
    write_schedule_of_benefits_sheets(workbook, display_summaries)
    append_processing_log(log_sheet, processed_sources, errors, counts_by_source)
    autosize_columns(review_sheet, worksheet_header_map(review_sheet))

    workbook.save(master_path)
    separate_review_path = write_separate_review_workbook(master_path, review_path, rows, summaries)

    return WorkbookUpdateResult(
        rows_added=rows_added,
        rows_updated=rows_updated,
        review_path=str(separate_review_path),
        master_path=str(master_path),
    )


def load_flexible_benefit_summaries(path: str) -> dict[str, dict[str, object]]:
    if not path or not Path(path).is_file():
        return {}

    try:
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}

    documents = payload.get("documents", []) if isinstance(payload, dict) else []
    summaries: dict[str, dict[str, object]] = {}
    for document in documents:
        if not isinstance(document, dict):
            continue
        source_file = clean_text(str(document.get("sourceFile", "")))
        carrier = clean_text(str(document.get("carrier", "")))
        if source_file and carrier.lower() == "national healthclaim":
            summaries[Path(source_file).name.lower()] = document
    return summaries


def build_scanned_flexible_benefits_summary(
    pdf_path: Path,
    flexible_benefit_summary: dict[str, object] | None = None,
) -> EmployeeBenefitsSummary | None:
    """List NHC accounts and use local OCR values when the desktop app provides them."""
    source_key = normalize_for_search(pdf_path.stem)
    if not re.search(r"\b(?:nhc|national healthclaim|national health claim)\b", source_key):
        return None

    sections: list[str] = []
    if summary_flag(flexible_benefit_summary, "hcsaEnabled") or re.search(r"\b(?:hcsa|health care spending account)\b", source_key):
        sections.append("Health Care Spending Account")
    if summary_flag(flexible_benefit_summary, "lsaEnabled") or re.search(r"\b(?:lsa|lifestyle spending account)\b", source_key):
        sections.append("Lifestyle Spending Account")
    if not sections:
        return None

    review_note = build_flexible_benefit_summary_note(pdf_path, flexible_benefit_summary)
    return EmployeeBenefitsSummary(
        source_file=pdf_path.name,
        client_name="",
        carrier="National HealthClaim",
        class_name="All Classes",
        rows=[(section, review_note) for section in sections],
    )


def flexible_benefit_processing_status(flexible_benefit_summary: dict[str, object] | None) -> str:
    return (
        "National HealthClaim account allocation extracted using local Windows OCR"
        if flexible_benefit_summary is not None
        else "Image-only flexible-benefits source listed in summary; local OCR review required"
    )


def build_flexible_benefit_summary_note(
    pdf_path: Path,
    flexible_benefit_summary: dict[str, object] | None,
) -> str:
    lines = ["Provider: National HealthClaim"]
    if flexible_benefit_summary is None:
        lines.extend([
            f"Source: {pdf_path.name}",
            "Status: Image-only document - local OCR review required before values can be saved",
        ])
        return "\n".join(lines)

    annual_allocation = format_flexible_benefit_allocation(flexible_benefit_summary)
    if annual_allocation:
        lines.append(f"Annual Allocation: {annual_allocation}")
    allocation = format_flexible_benefit_default_allocation(flexible_benefit_summary)
    if allocation:
        lines.append(f"Default Allocation: {allocation}")
    source_page = flexible_benefit_source_page(flexible_benefit_summary)
    source = f"Source: {pdf_path.name} (local Windows OCR"
    lines.append(f"{source}, page {source_page})" if source_page else f"{source})")
    return "\n".join(lines)


def summary_flag(summary: dict[str, object] | None, field_name: str) -> bool:
    if not summary:
        return False
    value = summary.get(field_name)
    return value is True or clean_text(str(value)).lower() == "true"


def format_flexible_benefit_allocation(summary: dict[str, object]) -> str:
    class_amounts = summary.get("classAmounts", [])
    if not isinstance(class_amounts, list):
        return ""

    formatted_classes: list[str] = []
    for class_amount in class_amounts:
        if not isinstance(class_amount, dict):
            continue
        amounts = [
            ("Single", decimal_from_summary(class_amount.get("singleAmount"))),
            ("Couple", decimal_from_summary(class_amount.get("coupleAmount"))),
            ("Family", decimal_from_summary(class_amount.get("familyAmount"))),
        ]
        populated = [(label, amount) for label, amount in amounts if amount is not None]
        if not populated:
            continue
        distinct_amounts = unique_in_order([format_summary_currency(amount) for _label, amount in populated])
        if len(distinct_amounts) == 1 and len(populated) == 3:
            formatted_classes.append(f"{distinct_amounts[0]} per calendar year (Single, Couple, and Family)")
        else:
            formatted_classes.append(
                "; ".join(f"{label}: {format_summary_currency(amount)}" for label, amount in populated) + " per calendar year"
            )

    return " | ".join(unique_in_order(formatted_classes))


def format_flexible_benefit_default_allocation(summary: dict[str, object]) -> str:
    hcsa = decimal_from_summary(summary.get("defaultHcsaPercentage"))
    lsa = decimal_from_summary(summary.get("defaultLsaPercentage"))
    if hcsa is None and lsa is None:
        return ""
    values = []
    if hcsa is not None:
        values.append(f"{format_summary_number(hcsa)}% HCSA")
    if lsa is not None:
        values.append(f"{format_summary_number(lsa)}% LSA")
    return " / ".join(values)


def flexible_benefit_source_page(summary: dict[str, object]) -> int | None:
    class_amounts = summary.get("classAmounts", [])
    if not isinstance(class_amounts, list):
        return None
    for class_amount in class_amounts:
        if not isinstance(class_amount, dict):
            continue
        source_page = class_amount.get("sourcePageNumber")
        if isinstance(source_page, int) and source_page > 0:
            return source_page
    return None


def decimal_from_summary(value: object) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        return Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError):
        return None


def format_summary_currency(value: Decimal) -> str:
    return f"${value:,.2f}" if value % 1 else f"${value:,.0f}"


def format_summary_number(value: Decimal) -> str:
    return f"{value:.2f}" if value % 1 else f"{value:.0f}"


def build_benefit_summary_display_summaries(
    summaries: list[EmployeeBenefitsSummary],
) -> list[EmployeeBenefitsSummary]:
    """Carry shared plan documents into each employee-class summary without duplicating database rows."""
    if not summaries:
        return []

    class_summaries = [summary for summary in summaries if not is_all_classes_summary(summary)]
    if not class_summaries:
        class_summaries = [summaries[0]]

    supplemental_summaries = [summary for summary in summaries if summary not in class_summaries]
    display_summaries: list[EmployeeBenefitsSummary] = []
    for class_summary in class_summaries:
        combined_rows = list(class_summary.rows)
        for supplemental_summary in supplemental_summaries:
            if summaries_share_client(class_summary, supplemental_summary):
                combined_rows = merge_summary_rows(combined_rows, supplemental_summary.rows)

        display_summaries.append(
            EmployeeBenefitsSummary(
                source_file=class_summary.source_file,
                client_name=class_summary.client_name,
                carrier=class_summary.carrier,
                class_name=class_summary.class_name,
                rows=combined_rows,
            )
        )
    return display_summaries


def is_all_classes_summary(summary: EmployeeBenefitsSummary) -> bool:
    return normalize_for_search(summary.class_name) in {"", "all classes", "all class"}


def summaries_share_client(left: EmployeeBenefitsSummary, right: EmployeeBenefitsSummary) -> bool:
    left_client = normalize_for_search(left.client_name)
    right_client = normalize_for_search(right.client_name)
    if not left_client or not right_client:
        return True
    return left_client in right_client or right_client in left_client


def merge_summary_rows(
    base_rows: list[tuple[str, str]],
    supplemental_rows: list[tuple[str, str]],
) -> list[tuple[str, str]]:
    metadata_sections = {"Carrier & Policy", "Eligibility Requirements", "Class Breakdown"}
    merged_rows = list(base_rows)
    row_positions = {
        normalize_summary_section_name(section_name): index
        for index, (section_name, _details) in enumerate(merged_rows)
    }

    for section_name, details in supplemental_rows:
        normalized_section = normalize_summary_section_name(section_name)
        if normalized_section in metadata_sections or not clean_text(details):
            continue

        existing_index = row_positions.get(normalized_section)
        if existing_index is None:
            merged_rows.append((normalized_section, details))
            row_positions[normalized_section] = len(merged_rows) - 1
            continue

        existing_name, existing_details = merged_rows[existing_index]
        merged_rows[existing_index] = (
            existing_name,
            "\n".join(unique_in_order([*str(existing_details).splitlines(), *str(details).splitlines()])),
        )
    return merged_rows


def load_or_create_workbook(path: Path):
    if path.exists():
        return load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")

    workbook = Workbook()
    default_sheet = workbook.active
    default_sheet.title = PLAN_SHEET
    return workbook


def ensure_sheet_with_columns(workbook, sheet_name: str, columns: list[str], clear_existing: bool = False):
    if sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
    else:
        sheet = workbook.create_sheet(sheet_name)

    if clear_existing and sheet.max_row:
        sheet.delete_rows(1, sheet.max_row)

    existing_headers = [sheet.cell(row=1, column=column_index).value for column_index in range(1, sheet.max_column + 1)]
    existing_headers = [str(value).strip() for value in existing_headers if value]
    if not existing_headers:
        for column_index, column_name in enumerate(columns, start=1):
            sheet.cell(row=1, column=column_index).value = column_name
        return sheet

    header_set = {header.lower() for header in existing_headers}
    next_column = sheet.max_column + 1
    for column_name in columns:
        if column_name.lower() not in header_set:
            sheet.cell(row=1, column=next_column).value = column_name
            next_column += 1

    return sheet


def update_plan_design_sheet(sheet, rows: list[ExtractionRow]) -> tuple[int, int, dict[str, dict[str, int]]]:
    header_map = worksheet_header_map(sheet)
    remove_superseded_labeled_dental_rows(sheet, header_map, rows)
    existing = existing_plan_design_keys(sheet, header_map)
    existing_by_source = existing_plan_design_source_keys(sheet, header_map)
    rows_added = 0
    rows_updated = 0
    counts_by_source: dict[str, dict[str, int]] = defaultdict(lambda: {"added": 0, "updated": 0})

    for row in rows:
        data = plan_row_to_dict(row)
        key = (
            row.client_id.strip().lower(),
            row.class_name.strip().lower(),
            row.benefit.strip().lower(),
            row.source_section.strip().lower(),
            row.coverage_description.strip().lower(),
        )
        source_key = (
            row.source_file.strip().lower(),
            row.class_name.strip().lower(),
            row.benefit.strip().lower(),
            row.source_section.strip().lower(),
            row.coverage_description.strip().lower(),
        )
        row_index = existing.get(key)
        identity_changed = False
        if row_index is None and source_key in existing_by_source:
            row_index = existing_by_source[source_key]
            identity_changed = True

        if row_index is not None:
            changed = update_existing_plan_row(sheet, header_map, row_index, data, row.confidence, identity_changed)
            if changed:
                rows_updated += 1
                counts_by_source[row.source_file]["updated"] += 1
                existing[key] = row_index
                existing_by_source[source_key] = row_index
            continue

        row_index = sheet.max_row + 1
        for column_name in PLAN_COLUMNS:
            sheet.cell(row=row_index, column=header_map[column_name]).value = data.get(column_name, "")
        existing[key] = row_index
        existing_by_source[source_key] = row_index
        rows_added += 1
        counts_by_source[row.source_file]["added"] += 1

    return rows_added, rows_updated, counts_by_source


def remove_superseded_labeled_dental_rows(
    sheet,
    header_map: dict[str, int],
    incoming_rows: list[ExtractionRow],
) -> None:
    """Remove stale dental interpretations for a source/class before canonical rows are upserted."""

    represented_groups = {
        (row.source_file.strip().lower(), row.class_name.strip().lower())
        for row in incoming_rows
        if row.source_file.strip()
    }
    canonical_groups: dict[tuple[str, str], dict[str, set]] = defaultdict(
        lambda: {"codes": set(), "percentages": set()}
    )
    incoming_canonical_keys: dict[tuple[str, str], set[tuple[str, str, str]]] = defaultdict(set)
    for row in incoming_rows:
        if CANONICAL_LABELED_DENTAL_RATE_MAPPER_VERSION not in row.notes:
            continue
        group = (row.source_file.strip().lower(), row.class_name.strip().lower())
        provision_code = canonical_dental_provision_code(row.benefit)
        if provision_code:
            canonical_groups[group]["codes"].add(provision_code)
        canonical_groups[group]["percentages"].update(_valid_percentage_values(row.coinsurance))
        incoming_canonical_keys[group].add(
            (
                row.benefit.strip().lower(),
                row.source_section.strip().lower(),
                row.coverage_description.strip().lower(),
            )
        )

    if not represented_groups:
        return

    source_column = header_map.get("SourceFile")
    class_column = header_map.get("ClassName")
    benefit_column = header_map.get("Benefit")
    notes_column = header_map.get("Notes")
    coinsurance_column = header_map.get("Coinsurance")
    source_section_column = header_map.get("SourceSection")
    coverage_column = header_map.get("CoverageDescription")
    if not source_column or not benefit_column:
        return

    rows_to_delete: list[int] = []
    for row_index in range(2, sheet.max_row + 1):
        source_file = clean_cell(sheet.cell(row=row_index, column=source_column).value).lower()
        class_name = clean_cell(sheet.cell(row=row_index, column=class_column).value).lower() if class_column else ""
        group = (source_file, class_name)
        if group not in represented_groups:
            continue

        notes = clean_cell(sheet.cell(row=row_index, column=notes_column).value) if notes_column else ""
        if CANONICAL_LABELED_DENTAL_RATE_MAPPER_VERSION in notes:
            existing_key = (
                clean_cell(sheet.cell(row=row_index, column=benefit_column).value).lower(),
                clean_cell(sheet.cell(row=row_index, column=source_section_column).value).lower()
                if source_section_column else "",
                clean_cell(sheet.cell(row=row_index, column=coverage_column).value).lower()
                if coverage_column else "",
            )
            if existing_key not in incoming_canonical_keys.get(group, set()):
                rows_to_delete.append(row_index)
            continue

        canonical_group = canonical_groups.get(group)
        if not canonical_group:
            continue
        benefit = clean_cell(sheet.cell(row=row_index, column=benefit_column).value)
        provision_code = canonical_dental_provision_code(benefit)
        if provision_code and provision_code in canonical_group["codes"]:
            rows_to_delete.append(row_index)
            continue

        normalized_benefit = normalize_for_search(benefit)
        percentages = _valid_percentage_values(
            clean_cell(sheet.cell(row=row_index, column=coinsurance_column).value) if coinsurance_column else ""
        )
        if (
            normalized_benefit in {"dental", "dental care", "dental coverage"}
            and len(percentages) > 1
            and set(percentages).issubset(canonical_group["percentages"])
        ):
            rows_to_delete.append(row_index)

    for row_index in reversed(rows_to_delete):
        sheet.delete_rows(row_index, 1)


def existing_plan_design_keys(sheet, header_map: dict[str, int]) -> dict[tuple[str, str, str, str, str], int]:
    keys: dict[tuple[str, str, str, str, str], int] = {}
    client_col = header_map["ClientID"]
    class_col = header_map.get("ClassName")
    benefit_col = header_map["Benefit"]
    source_section_col = header_map.get("SourceSection")
    coverage_col = header_map.get("CoverageDescription")

    for row_index in range(2, sheet.max_row + 1):
        client_id = clean_cell(sheet.cell(row=row_index, column=client_col).value).lower()
        class_name = clean_cell(sheet.cell(row=row_index, column=class_col).value).lower() if class_col else ""
        benefit = clean_cell(sheet.cell(row=row_index, column=benefit_col).value).lower()
        source_section = clean_cell(sheet.cell(row=row_index, column=source_section_col).value).lower() if source_section_col else ""
        coverage_description = clean_cell(sheet.cell(row=row_index, column=coverage_col).value).lower() if coverage_col else ""
        if client_id and benefit:
            keys[(client_id, class_name, benefit, source_section, coverage_description)] = row_index

    return keys


def existing_plan_design_source_keys(sheet, header_map: dict[str, int]) -> dict[tuple[str, str, str, str, str], int]:
    keys: dict[tuple[str, str, str, str, str], int] = {}
    source_col = header_map["SourceFile"]
    class_col = header_map.get("ClassName")
    benefit_col = header_map["Benefit"]
    source_section_col = header_map.get("SourceSection")
    coverage_col = header_map.get("CoverageDescription")

    for row_index in range(2, sheet.max_row + 1):
        source_file = clean_cell(sheet.cell(row=row_index, column=source_col).value).lower()
        class_name = clean_cell(sheet.cell(row=row_index, column=class_col).value).lower() if class_col else ""
        benefit = clean_cell(sheet.cell(row=row_index, column=benefit_col).value).lower()
        source_section = clean_cell(sheet.cell(row=row_index, column=source_section_col).value).lower() if source_section_col else ""
        coverage_description = clean_cell(sheet.cell(row=row_index, column=coverage_col).value).lower() if coverage_col else ""
        if source_file and benefit:
            keys[(source_file, class_name, benefit, source_section, coverage_description)] = row_index

    return keys


def update_existing_plan_row(
    sheet,
    header_map: dict[str, int],
    row_index: int,
    data: dict[str, str],
    confidence: str,
    update_identity: bool = False,
) -> bool:
    changed = False
    clearable_extracted_columns = {
        "Coinsurance",
        "Maximum",
        "Frequency",
        "WaitingPeriod",
        "TerminationAge",
        "ReviewFlag",
    }
    for column_name in PLAN_COLUMNS:
        if column_name == "Benefit" or (column_name == "ClientID" and not update_identity):
            continue

        new_value = data.get(column_name, "")
        cell = sheet.cell(row=row_index, column=header_map[column_name])
        old_value = clean_cell(cell.value)

        if not new_value:
            if column_name in clearable_extracted_columns and old_value:
                cell.value = ""
                changed = True
            continue

        if column_name == "Notes":
            if not old_value:
                cell.value = new_value
                changed = True
            continue

        if confidence == "Low" and old_value and column_name not in clearable_extracted_columns:
            continue

        if old_value != new_value:
            cell.value = new_value
            changed = True

    return changed


def write_review_sheet(
    sheet,
    rows: list[ExtractionRow],
    summaries: list[EmployeeBenefitsSummary],
) -> None:
    header_map = worksheet_header_map(sheet)
    benchmark_rows = build_benchmark_review_rows(rows, summaries)
    for row_index, data in enumerate(benchmark_rows, start=2):
        for column_name in REVIEW_COLUMNS:
            sheet.cell(row=row_index, column=header_map[column_name]).value = data.get(column_name, "")


def build_benchmark_review_rows(
    rows: list[ExtractionRow],
    summaries: list[EmployeeBenefitsSummary],
) -> list[dict[str, str]]:
    rows_by_key: dict[tuple[str, str], list[ExtractionRow]] = defaultdict(list)
    for row in rows:
        rows_by_key[(row.source_file.lower(), row.class_name.lower())].append(row)

    if summaries:
        return [
            build_benchmark_review_row(summary, rows_by_key.get((summary.source_file.lower(), summary.class_name.lower()), []))
            for summary in summaries
        ]

    grouped_rows: dict[tuple[str, str, str, str], list[ExtractionRow]] = defaultdict(list)
    for row in rows:
        grouped_rows[(row.source_file, row.client_name, row.carrier, row.class_name)].append(row)

    fallback_rows = []
    for (source_file, client_name, carrier, class_name), group_rows in grouped_rows.items():
        fallback_rows.append(
            {
                "SourceFile": source_file,
                "ClientName": client_name,
                "Carrier": carrier,
                "ClassName": class_name,
                "Benefit": "Class Benchmark Summary",
                "CoverageDescription": build_benchmark_coverage_description(group_rows),
                "Confidence": aggregate_confidence(group_rows),
                "ReviewFlag": aggregate_review_flags(group_rows),
            }
        )
    return fallback_rows


def build_benchmark_review_row(
    summary: EmployeeBenefitsSummary,
    rows: list[ExtractionRow],
) -> dict[str, str]:
    sections = extract_summary_sections(summary)
    benchmark_data = extract_benchmark_columns_from_sections(sections)
    base_row = rows[0] if rows else None
    data = {column_name: "" for column_name in REVIEW_COLUMNS}
    data.update(
        {
            "ClientID": base_row.client_id if base_row else "",
            "ClientName": summary.client_name,
            "Carrier": summary.carrier,
            "PolicyNumber": first_non_empty(
                extract_policy_number_from_summary(summary),
                base_row.policy_number if base_row else "",
            ),
            "SourceFile": summary.source_file,
            "ClassName": summary.class_name,
            "ClassDescription": summary.class_name,
            "BenefitCategory": "Class Summary",
            "BenefitName": "Class Benchmark Summary",
            "BenefitIncluded": "Yes" if rows or summary.rows else "",
            "Benefit": "Class Benchmark Summary",
            "CoverageDescription": build_benchmark_coverage_description(rows),
            "Coinsurance": "",
            "Maximum": "",
            "Frequency": "",
            "WaitingPeriod": "",
            "TerminationAge": "",
            "Confidence": aggregate_confidence(rows),
            "ReviewFlag": aggregate_review_flags(rows),
            "LastUpdated": max((row.last_updated for row in rows if row.last_updated), default=""),
        }
    )
    data.update(benchmark_data)
    data.update(build_requested_review_values(sections, benchmark_data))
    return data


def extract_policy_number_from_summary(summary: EmployeeBenefitsSummary) -> str:
    carrier_policy = next((details for name, details in summary.rows if name == "Carrier & Policy"), "")
    carrier_line = next((clean_text(line) for line in str(carrier_policy).splitlines() if ":" in line), "")
    if ":" not in carrier_line:
        return ""

    _carrier, value = carrier_line.split(":", 1)
    cleaned = clean_text(value)
    if len(cleaned) < 3 or cleaned.lower() in {"s", "n/a", "na"}:
        return ""
    return cleaned


def build_benchmark_coverage_description(rows: list[ExtractionRow]) -> str:
    benefits = unique_in_order(row.benefit for row in rows)
    if not benefits:
        return ""
    return f"Extracted benefits: {', '.join(benefits)}"


def aggregate_confidence(rows: list[ExtractionRow]) -> str:
    confidence_values = {row.confidence for row in rows if row.confidence}
    if "Low" in confidence_values:
        return "Low"
    if "Medium" in confidence_values:
        return "Medium"
    if "High" in confidence_values:
        return "High"
    return ""


def aggregate_review_flags(rows: list[ExtractionRow]) -> str:
    flags = unique_in_order(row.review_flag for row in rows if row.review_flag)
    return "; ".join(flags)


def extract_summary_sections(summary: EmployeeBenefitsSummary) -> dict[str, dict[str, str]]:
    return {
        section_name: parse_summary_detail_fields(details)
        for section_name, details in summary.rows
    }


def extract_benchmark_columns_from_summary(summary: EmployeeBenefitsSummary) -> dict[str, str]:
    return extract_benchmark_columns_from_sections(extract_summary_sections(summary))


def extract_benchmark_columns_from_sections(sections: dict[str, dict[str, str]]) -> dict[str, str]:
    data = {column_name: "" for column_name in BENCHMARK_COLUMNS}

    carrier_policy = sections.get("Carrier & Policy", {})
    data["PolicyholderLegalName"] = carrier_policy.get("Policyholder Legal Name", "")
    data["PolicyEffectiveDate"] = carrier_policy.get("Policy Effective Date", "")
    data["PolicyAnniversaryDate"] = carrier_policy.get("Policy Anniversary Date", "")

    life = sections.get("Life / AD&D", {})
    life_benefit_amount = first_non_empty(
        life.get("Life Benefit Amount", ""),
        life.get("Benefit Amount", ""),
        life.get("Maximum Benefit", ""),
    )
    data["LifeBenefitAmount"] = normalize_money_phrase(life_benefit_amount) if life_benefit_amount else ""
    data["LifeEarningsMultiplier"] = extract_earnings_multiplier(life_benefit_amount)
    data["ADDBenefitAmount"] = first_non_empty(
        life.get("AD&D / ASI Benefit Amount", ""),
        life.get("AD&D Benefit Amount", ""),
        life.get("ASI Benefit Amount", ""),
    )

    dependent_life = sections.get("Dependent Life", {})
    data["DependentLifeSpouseAmount"] = normalize_money_phrase(dependent_life.get("Spouse", ""))
    data["DependentLifeChildAmount"] = normalize_money_phrase(dependent_life.get("Child", ""))

    critical_illness = sections.get("Critical Illness", {})
    data["CriticalIllnessAmount"] = normalize_money_phrase(
        first_non_empty(
            critical_illness.get("Employee Benefit Amount", ""),
            critical_illness.get("Principal Sum", ""),
            critical_illness.get("Benefit Amount", ""),
            critical_illness.get("Maximum", ""),
        )
    )
    data["CriticalIllnessProvider"] = critical_illness.get("Provider", "")
    data["CriticalIllnessPlanType"] = first_non_empty(
        critical_illness.get("Plan Type", ""),
        critical_illness.get("Plan", ""),
    )
    data["CriticalIllnessEmployeeAmount"] = normalize_money_phrase(critical_illness.get("Employee Benefit Amount", ""))
    data["CriticalIllnessSpouseAmount"] = normalize_money_phrase(critical_illness.get("Spouse Benefit Amount", ""))
    data["CriticalIllnessChildAmount"] = normalize_money_phrase(critical_illness.get("Child Benefit Amount", ""))
    data["CriticalIllnessEvidenceRequirement"] = first_non_empty(
        critical_illness.get("Evidence of Insurability Requirement", ""),
        critical_illness.get("Evidence of Insurability", ""),
    )
    data["CriticalIllnessEvidenceThreshold"] = normalize_money_phrase(
        first_non_empty(
            critical_illness.get("Evidence Threshold", ""),
            data["CriticalIllnessEvidenceRequirement"],
        )
    )
    data["CriticalIllnessOptionalBenefits"] = first_non_empty(
        critical_illness.get("Optional Benefit", ""),
        critical_illness.get("Optional Benefits", ""),
    )
    data["CriticalIllnessRiders"] = critical_illness.get("Riders", "")
    data["CriticalIllnessSingleMonthlyRate"] = critical_illness.get("Single Monthly Rate", "")
    data["CriticalIllnessFamilyMonthlyRate"] = critical_illness.get("Family Monthly Rate", "")
    data["CriticalIllnessRateGuaranteeDate"] = critical_illness.get("Rate Guarantee Date", "")
    data["CriticalIllnessPreEx"] = first_non_empty(
        critical_illness.get("Pre-Ex", ""),
        critical_illness.get("Pre-Existing Conditions", ""),
        critical_illness.get("Re-Entry Exclusions", ""),
    )
    data["CriticalIllnessTerminationAge"] = first_non_empty(
        critical_illness.get("Termination Age", ""),
        critical_illness.get("Termination", ""),
    )

    std = sections.get("Short-Term Disability", {})
    data["STDBenefitPercent"] = extract_percent(first_non_empty(std.get("Benefit Schedule", ""), std.get("Benefit Amount", "")))
    data["STDMaximum"] = normalize_money_phrase(first_non_empty(std.get("Maximum", ""), std.get("Maximum Benefit", "")))

    ltd = sections.get("Long-Term Disability", {})
    ltd_schedule = first_non_empty(ltd.get("Benefit Schedule", ""), ltd.get("Benefit Amount", ""))
    data["LTDBenefitPercent"] = extract_percent(ltd_schedule)
    data["LTDMaximum"] = normalize_money_phrase(
        first_non_empty(
            ltd.get("Maximum With Evidence", ""),
            ltd.get("Maximum Benefit", ""),
            ltd.get("Maximum", ""),
            ltd.get("Maximum Without Evidence", ""),
        )
    )
    data["LTDEliminationPeriodDays"] = extract_day_count(ltd.get("Elimination Period", ""))

    drugs = sections.get("Prescription Drugs", {})
    data["PrescriptionDrugCoinsurance"] = extract_percent(drugs.get("Coinsurance", "")) or drugs.get("Coinsurance", "")
    data["DrugPlanType"] = drugs.get("Plan Type", "")
    data["DrugMax"] = first_non_empty(drugs.get("Drug Max", ""), drugs.get("Drug Maximum", ""), drugs.get("Maximum", ""))
    data["DrugType"] = first_non_empty(drugs.get("Drug Type", ""), drugs.get("Drug Type / Formulary", ""), drugs.get("Formulary", ""))

    hospital = sections.get("Hospital", {})
    data["HospitalCoverage"] = hospital.get("Coverage", "")

    paramedicals = sections.get("Paramedicals", {})
    paramedical_maximums = paramedicals.get("Maximums", "") or paramedicals.get("Maximum", "")
    data["ParamedicalStandardMaximum"] = extract_paramedical_benchmark_maximum(paramedical_maximums, "standard")
    data["ParamedicalPsychologyMaximum"] = extract_paramedical_benchmark_maximum(paramedical_maximums, "psychology")

    vision = sections.get("Vision Care", {})
    vision_maximum = vision.get("Maximum", "")
    data["VisionMaximum"] = normalize_money_phrase(vision_maximum)
    data["VisionFrequencyMonths"] = extract_month_count(vision_maximum)
    data["EyeExamFrequencyMonths"] = extract_month_count(vision.get("Eye Exams", ""))

    dental = sections.get("Dental Care", {})
    dental_maximums = dental.get("Maximums", "") or dental.get("Maximum", "")
    data["DentalBasicCoinsurance"] = extract_percent(dental.get("Basic Services", "")) or dental.get("Basic Services", "")
    data["DentalMajorCoinsurance"] = extract_percent(dental.get("Major Restorative Services", "")) or dental.get(
        "Major Restorative Services", ""
    )
    data["DentalOrthoCoinsurance"] = extract_percent(dental.get("Orthodontics", "")) or dental.get("Orthodontics", "")
    data["DentalBasicMaximum"] = extract_dental_benchmark_maximum(dental_maximums, "basic")
    data["DentalMajorMaximum"] = extract_dental_benchmark_maximum(dental_maximums, "major")
    data["DentalOrthoMaximum"] = extract_dental_benchmark_maximum(dental_maximums, "ortho")
    data["DentalRecallFrequency"] = dental.get("Recall", "")

    travel = sections.get("Emergency Travel", {})
    data["TravelMaximum"] = normalize_money_phrase(first_non_empty(travel.get("Maximum", ""), travel.get("Lifetime Maximum", "")))
    data["TravelPerTripMaximum"] = normalize_money_phrase(
        first_non_empty(travel.get("Per-Trip Family Benefits Maximum", ""), travel.get("Per Trip Maximum", ""))
    )
    data["TravelTripDurationDays"] = extract_day_count(travel.get("Trip Duration", ""))

    data["EAPIncluded"] = included_flag(sections.get("Employee Assistance Program", {}))
    data["VirtualCareIncluded"] = included_flag(sections.get("Virtual Care", {}))
    data["SecondOpinionIncluded"] = included_flag(sections.get("Medical Second Opinion", {}))

    return data


def build_requested_review_values(
    sections: dict[str, dict[str, str]],
    benchmark_data: dict[str, str],
) -> dict[str, str]:
    eligibility = sections.get("Eligibility Requirements", {})
    life = sections.get("Life / AD&D", {})
    drugs = sections.get("Prescription Drugs", {})
    medical = sections.get("Medical Services & Supplies", {})
    ltd = sections.get("Long-Term Disability", {})
    health_dental_termination = sections.get("Health & Dental Termination", {})

    benefit_amount_text = first_non_empty(
        benchmark_data.get("LifeBenefitAmount", ""),
        benchmark_data.get("CriticalIllnessAmount", ""),
    )
    waiting_period = first_non_empty(
        eligibility.get("Waiting Period", ""),
        eligibility.get("Eligibility Period", ""),
    )
    deductible_text = first_non_empty(
        drugs.get("Deductible", ""),
        medical.get("Deductible", ""),
    )
    ltd_maximum = benchmark_data.get("LTDMaximum", "")
    termination_text = first_non_empty(
        health_dental_termination.get("Termination Age", ""),
        health_dental_termination.get("Termination", ""),
    )

    return {
        "BenefitAmountText": benefit_amount_text,
        "BenefitAmountType": extract_benefit_amount_type(benefit_amount_text),
        "BenefitAmountNumeric": extract_money_number(benefit_amount_text),
        "EarningsMultiplier": benchmark_data.get("LifeEarningsMultiplier", ""),
        "SalaryPercent": first_non_empty(benchmark_data.get("LTDBenefitPercent", ""), benchmark_data.get("STDBenefitPercent", "")),
        "CoinsuranceText": drugs.get("Coinsurance", ""),
        "CoinsurancePercent": extract_percent(drugs.get("Coinsurance", "")),
        "DeductibleText": deductible_text,
        "SingleDeductible": extract_deductible_amount(deductible_text, "single"),
        "FamilyDeductible": extract_deductible_amount(deductible_text, "family"),
        "MaximumText": ltd_maximum,
        "MaximumAmount": extract_money_number(ltd_maximum),
        "MaximumFrequency": extract_maximum_frequency(ltd_maximum),
        "LifetimeMaximum": extract_lifetime_maximum(benchmark_data),
        "NonEvidenceMaximum": life.get("Non-Evidence Maximum", ""),
        "WaitingPeriodText": waiting_period,
        "WaitingPeriodMonths": extract_month_count(waiting_period),
        "EliminationPeriodText": ltd.get("Elimination Period", ""),
        "EliminationPeriodDays": benchmark_data.get("LTDEliminationPeriodDays", ""),
        "BenefitPeriod": ltd.get("Benefit Period", ""),
        "TerminationAgeText": termination_text,
        "TerminationAgeNumber": extract_termination_age_number(termination_text),
        "TerminationTrigger": extract_termination_trigger(termination_text),
        "TaxStatus": first_non_empty(ltd.get("Tax Status", ""), sections.get("Critical Illness", {}).get("Tax Status", "")),
        "WaiverOfPremium": life.get("Waiver of Premium", ""),
        "DefinitionOfDisability": ltd.get("Definition", ""),
    }


def parse_summary_detail_fields(details: str) -> dict[str, str]:
    fields: dict[str, list[str]] = {}
    current_label = ""

    for raw_line in str(details or "").splitlines():
        line = clean_text(raw_line.strip())
        if not line:
            continue

        if line.startswith("- ") and current_label:
            fields.setdefault(current_label, []).append(line[2:].strip())
            continue

        if ":" in line:
            label, value = line.split(":", 1)
            current_label = clean_text(label)
            fields.setdefault(current_label, [])
            value = clean_text(value)
            if value:
                fields[current_label].append(value)
            continue

        if current_label:
            fields.setdefault(current_label, []).append(line)

    return {label: "; ".join(unique_in_order(values)) for label, values in fields.items()}


def first_non_empty(*values: str) -> str:
    for value in values:
        cleaned = clean_text(value)
        if cleaned:
            return cleaned
    return ""


def extract_percent(value: str) -> str:
    match = re.search(r"\b\d{1,3}(?:\.\d+)?\s*%", clean_text(value))
    return match.group(0).replace(" ", "") if match else ""


def extract_earnings_multiplier(value: str) -> str:
    match = re.search(r"\b(\d+(?:\.\d+)?)\s*(?:x|times?)\b", clean_text(value), flags=re.IGNORECASE)
    if not match:
        return ""
    multiplier = match.group(1)
    if multiplier.endswith(".0"):
        multiplier = multiplier[:-2]
    return f"{multiplier}x"


def extract_benefit_amount_type(value: str) -> str:
    cleaned = clean_text(value)
    if extract_earnings_multiplier(cleaned):
        return "Earnings multiple"
    if re.search(MONEY_AMOUNT_PATTERN, cleaned):
        return "Flat amount"
    if extract_percent(cleaned):
        return "Salary percentage"
    return ""


def extract_money_number(value: str) -> str:
    money = first_regex(clean_text(value), [rf"({MONEY_AMOUNT_PATTERN})"])
    if not money:
        return ""
    return re.sub(r"[^\d.]", "", money)


def extract_deductible_amount(value: str, kind: str) -> str:
    cleaned = clean_text(value)
    labeled = first_regex(cleaned, [rf"({MONEY_AMOUNT_PATTERN})\s+{kind}"])
    if labeled:
        return clean_money(labeled)

    amounts = unique_regex_group_matches(cleaned, [rf"({MONEY_AMOUNT_PATTERN})"])
    if kind == "single" and amounts:
        return clean_money(amounts[0])
    if kind == "family" and len(amounts) > 1:
        return clean_money(amounts[1])
    return ""


def extract_maximum_frequency(value: str) -> str:
    cleaned = clean_text(value)
    patterns = [
        r"\bper\s+month\b",
        r"\bper\s+week\b",
        r"\bper\s+calendar\s+year\b",
        r"\bper\s+lifetime\b",
        r"\blifetime\b",
        r"\bevery\s+\d+\s+(?:consecutive\s+)?months?\b",
        r"\bevery\s+\d+\s+years?\b",
    ]
    return first_regex(cleaned, [f"({pattern})" for pattern in patterns])


def extract_lifetime_maximum(benchmark_data: dict[str, str]) -> str:
    candidates = [
        benchmark_data.get("DentalOrthoMaximum", ""),
        benchmark_data.get("VisionMaximum", ""),
        benchmark_data.get("TravelMaximum", ""),
    ]
    for candidate in candidates:
        if re.search(r"\blifetime\b", candidate, flags=re.IGNORECASE):
            return candidate
    return ""


def extract_termination_age_number(value: str) -> str:
    match = re.search(
        r"\bage\s+(\d{1,3})\b|\bturn\s+(\d{1,3})\b|\b(\d{1,3})(?:st|nd|rd|th)?\s+birthday\b",
        clean_text(value),
        flags=re.IGNORECASE,
    )
    if not match:
        return ""
    return next((group for group in match.groups() if group), "")


def extract_termination_trigger(value: str) -> str:
    cleaned = clean_text(value)
    triggers = []
    if re.search(r"\bemployment terminates\b", cleaned, flags=re.IGNORECASE):
        triggers.append("Employment terminates")
    if re.search(r"\bretire|retirement\b", cleaned, flags=re.IGNORECASE):
        triggers.append("Retirement")
    if re.search(r"\bage\s+\d{1,3}|\bturn\s+\d{1,3}|\d{1,3}(?:st|nd|rd|th)?\s+birthday\b", cleaned, flags=re.IGNORECASE):
        triggers.append("Age")
    return "; ".join(unique_in_order(triggers))


def extract_month_count(value: str) -> str:
    cleaned = clean_text(value)
    parenthetical_match = re.search(r"\((\d{1,3})\)\s+months?\b", cleaned, flags=re.IGNORECASE)
    if parenthetical_match:
        return parenthetical_match.group(1)

    match = re.search(r"\b(\d{1,3})\s+(?:consecutive\s+)?months?\b", cleaned, flags=re.IGNORECASE)
    if match:
        return match.group(1)

    year_match = re.search(r"\b(\d{1,2})\s+years?\b", cleaned, flags=re.IGNORECASE)
    if year_match:
        return str(int(year_match.group(1)) * 12)

    return ""


def extract_day_count(value: str) -> str:
    cleaned = clean_text(value)
    day_matches = [int(match) for match in re.findall(r"\b(\d{1,4})\s+day", cleaned, flags=re.IGNORECASE)]
    if day_matches:
        return str(max(day_matches))

    month_match = re.search(r"\b(\d{1,3})\s+months?\b", cleaned, flags=re.IGNORECASE)
    if month_match:
        return str(int(month_match.group(1)) * 30)

    return ""


def extract_paramedical_benchmark_maximum(value: str, kind: str) -> str:
    cleaned = clean_text(value)
    parts = [part.strip() for part in cleaned.split(";") if part.strip()]

    if kind == "psychology":
        for part in parts:
            if re.search(r"\b(?:psychologist|social worker|counsel)\b", part, flags=re.IGNORECASE):
                return normalize_money_phrase(part)
        for part in parts:
            if re.search(r"\bpsychotherapist\b", part, flags=re.IGNORECASE):
                return normalize_money_phrase(part)
        return ""

    for part in parts:
        if not re.search(r"\b(?:psychologist|social worker|counsel|psychotherapist)\b", part, flags=re.IGNORECASE):
            money = first_regex(part, [rf"({MONEY_AMOUNT_PATTERN})"])
            if money:
                return clean_money(money)
    return ""


def extract_dental_benchmark_maximum(value: str, kind: str) -> str:
    cleaned = clean_text(value)
    parts = [part.strip() for part in cleaned.split(";") if part.strip()]

    if kind in {"basic", "major"}:
        for part in parts:
            if re.search(r"\bNo maximum\b", part, flags=re.IGNORECASE):
                return "No maximum"

    if kind == "ortho":
        for part in parts:
            if re.search(r"\b(?:orthodontic|ortho|lifetime)\b", part, flags=re.IGNORECASE):
                return normalize_money_phrase(part)
        if len(parts) == 1 and re.search(MONEY_AMOUNT_PATTERN, parts[0]):
            return normalize_money_phrase(parts[0])

    return ""


def included_flag(fields: dict[str, str]) -> str:
    return "Yes" if any(clean_text(value) for value in fields.values()) else ""


def write_employee_benefits_summary_sheet(workbook, summaries: list[EmployeeBenefitsSummary]) -> None:
    remove_generated_class_summary_sheets(workbook, summaries)
    styles = build_employee_summary_styles()

    used_titles = set(workbook.sheetnames)
    for summary_index, summary in enumerate(summaries):
        class_sheet_title = build_class_summary_sheet_title(summary, used_titles)
        used_titles.add(class_sheet_title)
        class_sheet = get_or_create_summary_sheet(workbook, class_sheet_title, summary_index)
        setup_summary_sheet_layout(class_sheet)
        write_summary_blocks(class_sheet, [summary], styles)
        finalize_summary_sheet_print_layout(class_sheet)


def build_employee_summary_styles() -> dict[str, object]:
    divider = Side(style="thin", color="222222")
    soft_divider = Side(style="thin", color="B7B7B7")
    return {
        "client_font": Font(name="Calibri", size=14, bold=False, color="004C55"),
        "title_font": Font(name="Calibri", size=13, bold=False, color="39B8AD"),
        "section_font": Font(name="Calibri", size=11, bold=True, italic=False, color="004C55"),
        "section_highlight_font": Font(name="Calibri", size=11, bold=True, italic=False, color="FFFFFF"),
        "detail_font": Font(name="Calibri", size=11, color="4F4F4F"),
        "footer_font": Font(name="Calibri", size=9, bold=True, color="004C55"),
        "note_font": Font(name="Calibri", size=8, italic=True, color="004C55"),
        "white_fill": PatternFill("solid", fgColor="FFFFFF"),
        "section_highlight_fill": PatternFill("solid", fgColor="66C5B4"),
        "row_border": Border(top=divider),
        "soft_row_border": Border(top=soft_divider),
    }


def remove_generated_class_summary_sheets(
    workbook,
    summaries: list[EmployeeBenefitsSummary],
) -> None:
    target_titles: set[str] = set()
    for summary in summaries:
        base_title = build_class_summary_sheet_title(summary, set())
        target_titles.add(base_title)
        for suffix in range(2, 1000):
            suffix_text = f"_{suffix}"
            target_titles.add(f"{base_title[:31 - len(suffix_text)]}{suffix_text}")

    for sheet_name in list(workbook.sheetnames):
        if sheet_name in target_titles:
            del workbook[sheet_name]


def get_or_create_summary_sheet(workbook, sheet_name: str, index: int):
    if sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
    else:
        sheet = workbook.create_sheet(sheet_name, index)

    for merged_range in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged_range))
    if sheet.max_row:
        sheet.delete_rows(1, sheet.max_row)
    return sheet


def setup_summary_sheet_layout(sheet) -> None:
    sheet.sheet_view.showGridLines = False
    sheet.sheet_view.view = "normal"
    sheet.column_dimensions["A"].width = 36
    sheet.column_dimensions["B"].width = 92
    sheet.page_setup.orientation = "portrait"
    sheet.page_margins.left = 0.35
    sheet.page_margins.right = 0.35
    sheet.page_margins.top = 0.35
    sheet.page_margins.bottom = 0.35
    for row_index in range(1, SUMMARY_LOGO_START_ROW):
        sheet.row_dimensions[row_index].height = 14


def resolve_summary_logo_path() -> Path | None:
    repository_root = Path(__file__).resolve().parents[2]
    candidates = (
        repository_root / "Assets" / SUMMARY_LOGO_FILE_NAME,
        repository_root / "src" / "RenewalReviewBuilder.App" / "Assets" / SUMMARY_LOGO_FILE_NAME,
        Path.cwd() / "Assets" / SUMMARY_LOGO_FILE_NAME,
        Path.cwd() / "src" / "RenewalReviewBuilder.App" / "Assets" / SUMMARY_LOGO_FILE_NAME,
    )
    return next((candidate for candidate in candidates if candidate.is_file()), None)


def add_summary_logo(sheet) -> None:
    logo_path = resolve_summary_logo_path()
    if not logo_path:
        return

    logo = OpenPyxlImage(str(logo_path))
    logo.height = 47
    logo.width = 253
    sheet.add_image(logo, "A1")
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True


def finalize_summary_sheet_print_layout(sheet) -> None:
    last_content_row = 1
    for row_index in range(sheet.max_row, 0, -1):
        if any(clean_text(str(sheet.cell(row=row_index, column=column_index).value or "")) for column_index in range(1, 3)):
            last_content_row = row_index
            break

    print_end_row = last_content_row + 1
    sheet.print_area = f"A1:B{print_end_row}"


def write_summary_blocks(
    sheet,
    summaries: list[EmployeeBenefitsSummary],
    styles: dict[str, object],
) -> None:
    add_summary_logo(sheet)
    current_row = SUMMARY_LOGO_START_ROW
    for summary_index, summary in enumerate(summaries):
        if summary_index:
            current_row += 2

        normalized_rows = normalized_summary_display_rows(summary)

        client_cell = sheet.cell(row=current_row, column=1)
        client_cell.value = display_summary_client_name(summary)
        client_cell.font = styles["client_font"]
        client_cell.fill = styles["white_fill"]
        client_cell.alignment = Alignment(wrap_text=True, vertical="center")
        sheet.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
        sheet.row_dimensions[current_row].height = 28
        current_row += 1

        sheet.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
        title_cell = sheet.cell(row=current_row, column=1)
        title_cell.value = f"Employee Benefits Summary - Class: {display_summary_class_name(summary)}"
        title_cell.font = styles["title_font"]
        title_cell.fill = styles["white_fill"]
        title_cell.alignment = Alignment(wrap_text=True, vertical="center")
        sheet.row_dimensions[current_row].height = 24
        current_row += 2

        for section_name, details in normalized_rows:
            label = summary_section_label_text(section_name, summary)
            section_cell = sheet.cell(row=current_row, column=1)
            details_cell = sheet.cell(row=current_row, column=2)
            section_cell.value = label
            details_cell.value = details

            is_benefit_section = section_name not in {"Carrier & Policy", "Eligibility Requirements", "Class Breakdown"}
            if is_benefit_section:
                section_cell.fill = styles["section_highlight_fill"]
                section_cell.font = styles["section_highlight_font"]
            else:
                section_cell.font = styles["section_font"]
            details_cell.font = styles["detail_font"]
            section_cell.alignment = Alignment(wrap_text=True, vertical="top")
            details_cell.alignment = Alignment(wrap_text=True, vertical="top")
            for column_index in range(1, 3):
                sheet.cell(row=current_row, column=column_index).border = styles["row_border"]

            sheet.row_dimensions[current_row].height = summary_row_height(label, details)
            current_row += 1

        current_row = write_summary_footer(sheet, current_row, styles)


def normalized_summary_display_rows(summary: EmployeeBenefitsSummary) -> list[tuple[str, str]]:
    rows: list[tuple[int, int, str, str]] = []
    inserted_class_breakdown = False

    for original_index, (section_name, details) in enumerate(summary.rows):
        normalized_name = normalize_summary_section_name(section_name)
        normalized_details = normalize_summary_detail_text(normalized_name, details or "", summary)
        if not clean_text(normalized_details):
            continue

        rows.append((SUMMARY_ORDER_INDEX.get(normalized_name, 999), original_index, normalized_name, normalized_details))

        if normalized_name == "Eligibility Requirements":
            class_detail = build_class_breakdown_detail(summary)
            if class_detail:
                rows.append((SUMMARY_ORDER_INDEX["Class Breakdown"], original_index + 1, "Class Breakdown", class_detail))
                inserted_class_breakdown = True

    if not inserted_class_breakdown:
        class_detail = build_class_breakdown_detail(summary)
        if class_detail:
            rows.append((SUMMARY_ORDER_INDEX["Class Breakdown"], 0, "Class Breakdown", class_detail))

    rows.sort(key=lambda item: (item[0], item[1]))
    return [(section_name, details) for _order, _index, section_name, details in rows]


def normalize_summary_section_name(section_name: str) -> str:
    cleaned = clean_text(section_name)
    aliases = {
        "EAP": "Employee Assistance Program",
        "Emergency Travel": "Emergency Travel",
        "Travel": "Emergency Travel",
        "Extended Health Services": "Medical Services & Supplies",
        "Health Spending Account": "Health Care Spending Account",
        "Healthcare Spending Account": "Health Care Spending Account",
        "HCSA": "Health Care Spending Account",
        "HSCA": "Health Care Spending Account",
        "LSA": "Lifestyle Spending Account",
    }
    return aliases.get(cleaned, cleaned)


def build_class_breakdown_detail(summary: EmployeeBenefitsSummary) -> str:
    class_name = clean_class_name(summary.class_name)
    if not class_name:
        return ""

    match = re.match(r"^(Class\s+[A-Za-z0-9]+)\s*-\s*(.+)$", class_name, flags=re.IGNORECASE)
    if match:
        return f"{match.group(1)}: {match.group(2)}"
    return class_name


def summary_section_label_text(section_name: str, summary: EmployeeBenefitsSummary) -> str:
    if section_name == "Carrier & Policy" and is_benepac_summary(summary):
        return "Carrier & Policy: BBD"
    return section_name


def summary_section_audience(section_name: str, summary: EmployeeBenefitsSummary) -> str:
    if section_name in {"Carrier & Policy", "Eligibility Requirements", "Class Breakdown"}:
        return ""
    if section_name == "Critical Illness":
        critical_illness_details = next(
            (details for name, details in summary.rows if normalize_summary_section_name(name) == "Critical Illness"),
            "",
        )
        if re.search(r"\b(?:Spouse|Child|Dependent)\b", critical_illness_details, flags=re.IGNORECASE):
            return "Employee and Dependents"
        return "Employee only"
    if section_name == "Life / AD&D":
        return "Employee only"
    if section_name == "Long-Term Disability":
        return "Class A only\nEmployee only" if re.search(r"\bClass\s+A\b", summary.class_name, flags=re.IGNORECASE) else "Employee only"
    if section_name in {"Prescription Drugs", "Medical Services & Supplies", "Paramedicals", "Mental Health Practitioners", "Hospital", "Vision Care", "Emergency Travel", "Dental Care", "Employee Assistance Program", "Medical Second Opinion"}:
        return "Employee and Dependents"
    if section_name == "Survivor Benefit":
        return "Eligible Dependents"
    return ""


def normalize_summary_detail_text(section_name: str, details: str, summary: EmployeeBenefitsSummary) -> str:
    if section_name == "Carrier & Policy":
        return normalize_carrier_policy_detail(details, summary)

    lines = []
    for line in str(details or "").splitlines():
        normalized_line = normalize_summary_detail_line(section_name, line)
        if normalized_line:
            lines.append(normalized_line)
    return "\n".join(unique_in_order(lines))


def normalize_carrier_policy_detail(details: str, summary: EmployeeBenefitsSummary) -> str:
    normalized_lines: list[str] = []
    for raw_line in str(details or "").splitlines():
        line = clean_summary_noise(raw_line)
        if not line:
            continue

        if re.search(r"\bEmpire Life\b|\bThe Empire Life Insurance Company\b", line, flags=re.IGNORECASE):
            policy = first_regex(line, [r"\b(TC\d{3,}-\d+)\b"])
            normalized_lines.append(f"Empire Life: {policy} (Life & LTD)" if policy else "Empire Life")
            continue

        if re.search(r"\bIndustrial Alliance\b", line, flags=re.IGNORECASE):
            policy = first_regex(
                line,
                [r"\b(\d{5}-\d{4}-\d{4})\b", r"Industrial Alliance\s*:\s*([A-Za-z0-9-]+)"],
            )
            suffix = ""
            if "8101" in policy:
                suffix = " (AD&D)"
            elif "7153" in policy:
                suffix = " (CI)"
            else:
                documented_benefits = first_regex(line, [r"\(([^)]+)\)"])
                if documented_benefits:
                    suffix = f" ({documented_benefits})"
            normalized_lines.append(f"Industrial Alliance: {policy}{suffix}" if policy else "Industrial Alliance")
            continue

        if re.search(r"\bGreen\s*Shield\b|\bGreenShield\b", line, flags=re.IGNORECASE):
            policy = first_regex(line, [r"\b(\d{5})\b"])
            benefit_note = "Extended Health & Dental" if policy == "30720" else "Extended Health"
            normalized_lines.append(f"GreenShield Canada: {policy} ({benefit_note})" if policy else "GreenShield Canada")
            continue

        if re.search(r"\bPeople Corporation Health Solutions\b", line, flags=re.IGNORECASE):
            policy = first_regex(line, [r"\b(HC-\d+)\b"])
            normalized_lines.append(f"People Corporation Health Solutions: {policy}" if policy else "People Corporation Health Solutions")
            continue

        normalized_lines.append(line)

    if not normalized_lines and summary.carrier:
        return summary.carrier
    return "\n".join(unique_in_order(normalized_lines))


def normalize_summary_detail_line(section_name: str, line: str) -> str:
    cleaned = clean_summary_noise(line)
    if not cleaned:
        return ""

    if ":" not in cleaned:
        return cleaned

    label, value = cleaned.split(":", 1)
    raw_label = clean_text(label)
    raw_value = clean_summary_noise(value)
    if raw_label.casefold() == "cost sharing":
        return ""
    if section_name == "Paramedicals" and raw_label == "Services" and re.fullmatch(r"None", raw_value, flags=re.IGNORECASE):
        return "Services: None"

    label = normalize_summary_detail_label(section_name, label)
    value = normalize_summary_detail_value(section_name, label, value)
    if section_name == "Life / AD&D" and label == "Benefit Amount":
        return value
    return f"{label}: {value}" if value else ""


def clean_summary_noise(value: str) -> str:
    cleaned = clean_text(value)
    cleaned = re.sub(r"\b\d+\s+Benepac(?:®|\(R\))?\s+Benefit Details\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\bBenepac(?:®|\(R\))?\s+Benefit Details\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip(" -;:")


def normalize_summary_detail_label(section_name: str, label: str) -> str:
    cleaned = clean_text(label)
    replacements = {
        "Life Benefit Amount": "Benefit Amount",
        "Life Non-Evidence Maximum": "Non-Evidence Maximum",
        "Life Reduction": "Benefit Reduction",
        "Life Termination Age": "Benefit Period",
        "AD&D Termination Age": "Benefit Period",
        "EHS Coinsurance": "Coinsurance",
        "Overall Maximum": "Maximum",
        "Drug Max": "Maximum",
        "Maximum Number of Days per Trip": "Trip Duration",
        "Benefit Taxability": "Tax Status",
        "Own Occupation Period": "Definition",
        "Benefit Period": "Benefit Period",
    }
    if section_name == "Paramedicals" and cleaned == "Services":
        return "Benefit Amount"
    if section_name == "Vision Care" and cleaned == "Vision Benefit":
        return "Maximum"
    return replacements.get(cleaned, cleaned)


def normalize_summary_detail_value(section_name: str, label: str, value: str) -> str:
    cleaned = clean_summary_noise(value)
    cleaned = re.sub(r"\bGreen Shield Canada\b", "", cleaned, flags=re.IGNORECASE).strip(" -;:")
    cleaned = re.sub(r"\s+Empire Life$", "", cleaned, flags=re.IGNORECASE).strip()
    cleaned = cleaned.replace("per year", "per calendar year")

    if label in {"Deductible", "Single Deductible", "Family Deductible"}:
        if re.fullmatch(r"No deductible", cleaned, flags=re.IGNORECASE):
            return "None"
        return cleaned

    if label == "Provider" and re.fullmatch(r"NHC|National Health Claim", cleaned, flags=re.IGNORECASE):
        return "National HealthClaim"

    if section_name == "Paramedicals" and label == "Benefit Amount":
        amount = first_regex(cleaned, [rf"({MONEY_AMOUNT_PATTERN}\s+per\s+practitioner\s+per\s+calendar\s+year)"])
        return amount or cleaned

    if section_name == "Employee Assistance Program" and label == "Coverage":
        cleaned = re.sub(r"\s+with access to:\s*Program$", "", cleaned, flags=re.IGNORECASE).strip()
        if cleaned:
            cleaned = cleaned[0].upper() + cleaned[1:]

    return cleaned


def is_benepac_summary(summary: EmployeeBenefitsSummary) -> bool:
    if summary.carrier.lower() == "benepac":
        return True
    return any("Benepac Schedule of Benefits" in detail for _name, detail in summary.rows)


def summary_row_height(label: str, details: str) -> float:
    label_lines = label.count("\n") + 1
    detail_lines = details.count("\n") + 1
    wrapped_detail_lines = sum(max(0, len(line) // 110) for line in details.splitlines())
    return max(20, min(360, (max(label_lines, detail_lines) + wrapped_detail_lines) * 14 + 8))


def write_summary_footer(sheet, current_row: int, styles: dict[str, object]) -> int:
    current_row += 1
    day_format = "%#d" if sys.platform.startswith("win") else "%-d"
    updated_text = datetime.now().strftime(f"UPDATED: %B {day_format}, %Y").upper()

    sheet.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
    updated_cell = sheet.cell(row=current_row, column=1)
    updated_cell.value = updated_text
    updated_cell.font = styles["footer_font"]
    updated_cell.alignment = Alignment(vertical="center")
    sheet.row_dimensions[current_row].height = 16
    current_row += 1

    sheet.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
    note_cell = sheet.cell(row=current_row, column=1)
    note_cell.value = "E&OE: should there be a discrepancy between these pages and the group contract, the group contract will be considered correct and legally binding."
    note_cell.font = styles["note_font"]
    note_cell.alignment = Alignment(wrap_text=True, vertical="top")
    sheet.row_dimensions[current_row].height = 18
    return current_row + 1


def build_class_summary_sheet_title(summary: EmployeeBenefitsSummary, used_titles: set[str]) -> str:
    class_name = clean_text(summary.class_name)
    client_code = extract_source_client_code(summary.source_file)
    class_match = re.search(r"\bClass\s+([A-Za-z0-9]+)", class_name, flags=re.IGNORECASE)
    if class_match:
        base_title = f"Summary_{client_code}_{class_match.group(1)}" if client_code else f"Summary_Class_{class_match.group(1)}"
    else:
        fallback = class_name or Path(summary.source_file).stem or "Class"
        fallback = re.sub(r"[^A-Za-z0-9]+", "_", fallback).strip("_") or "Class"
        base_title = f"Summary_{client_code}_{fallback}" if client_code else f"Summary_{fallback}"

    base_title = base_title[:31].rstrip("_")
    title = base_title
    suffix = 2
    while title in used_titles:
        suffix_text = f"_{suffix}"
        title = f"{base_title[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    return title


def display_summary_client_name(summary: EmployeeBenefitsSummary) -> str:
    client_name = clean_client_name(summary.client_name)
    if client_name and "__" not in client_name:
        return client_name

    source_name = extract_source_client_name(summary.source_file)
    if source_name:
        return source_name

    if "__" in client_name:
        return humanize_compact_client_name(client_name.split("__", 2)[1])
    return client_name


def display_summary_class_name(summary: EmployeeBenefitsSummary) -> str:
    class_name = clean_class_name(summary.class_name)
    match = re.match(r"^Class\s+(.+)$", class_name, flags=re.IGNORECASE)
    return match.group(1) if match else (class_name or "All Employees")


def extract_source_client_name(source_file: str) -> str:
    stem = Path(source_file).stem
    parts = stem.split("__")
    if len(parts) < 2:
        return ""
    return humanize_compact_client_name(parts[1])


def humanize_compact_client_name(value: str) -> str:
    cleaned = re.sub(r"\s*\(\d+\)$", "", clean_text(value))
    if not cleaned:
        return ""

    replacements = {
        "forthe": " for the ",
        "ofthe": " of the ",
        "andthe": " and the ",
    }
    for compact, spaced in replacements.items():
        cleaned = re.sub(compact, spaced, cleaned, flags=re.IGNORECASE)

    cleaned = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", cleaned)
    cleaned = re.sub(r"\bLtd\b", "Ltd.", cleaned)
    cleaned = re.sub(r"\bInc\b", "Inc.", cleaned)
    cleaned = re.sub(r"\bSociety\b", "Society", cleaned)
    return clean_text(cleaned)


def extract_source_client_code(source_file: str) -> str:
    stem = Path(source_file).stem
    code = stem.split("__", 1)[0]
    code = re.sub(r"[^A-Za-z0-9]+", "_", code).strip("_")
    return code[:12]


def write_schedule_of_benefits_sheets(workbook, summaries: list[EmployeeBenefitsSummary]) -> None:
    remove_generated_schedule_sheets(workbook)
    pages = build_schedule_benefit_pages(summaries)

    schedule_sheet = workbook.create_sheet(SCHEDULE_SHEET)
    setup_schedule_sheet_layout(schedule_sheet)
    write_schedule_benefit_pages(schedule_sheet, pages)

    export_sheet = workbook.create_sheet(PPT_EXPORT_SHEET)
    write_schedule_ppt_export_sheet(export_sheet, pages)


def remove_generated_schedule_sheets(workbook) -> None:
    for sheet_name in (SCHEDULE_SHEET, PPT_EXPORT_SHEET):
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]


def build_schedule_benefit_pages(summaries: list[EmployeeBenefitsSummary]) -> list[ScheduleBenefitPage]:
    pages: list[ScheduleBenefitPage] = []

    for summary in summaries:
        grouped_sections: list[list[tuple[str, str]]] = [[], [], []]
        for section_name, details in summary.rows:
            if not schedule_section_should_be_on_page(section_name, details):
                continue
            grouped_sections[schedule_page_group_index(section_name)].append((section_name, details))

        chunks = [group for group in grouped_sections if group]
        if not chunks:
            chunks = [[("Review Note", "No benefit details were extracted for this booklet.")]]

        page_total = max(1, len(chunks))
        for page_index, sections in enumerate(chunks, start=1):
            pages.append(
                ScheduleBenefitPage(
                    summary=summary,
                    page_number=page_index,
                    page_total=page_total,
                    sections=sections,
                )
            )

    return pages


def schedule_page_group_index(section_name: str) -> int:
    normalized_name = clean_text(section_name).lower()

    if any(
        keyword in normalized_name
        for keyword in (
            "life",
            "a&d",
            "ad&d",
            "accident",
            "critical illness",
            "serious illness",
            "prescription",
            "drug",
            "medical services",
            "medical supplies",
        )
    ):
        return 0

    if any(
        keyword in normalized_name
        for keyword in (
            "hospital",
            "vision",
            "paramed",
            "mental health",
            "travel",
            "employee assistance",
            "virtual care",
            "digital wellness",
            "second opinion",
            "health & dental termination",
        )
    ):
        return 1

    return 2


def schedule_section_should_be_on_page(section_name: str, details: str) -> bool:
    normalized_name = clean_text(section_name).lower()
    if normalized_name in {"carrier & policy", "eligibility requirements"}:
        return False
    return bool(clean_text(details))


def estimate_schedule_section_lines(details: str) -> int:
    lines = schedule_detail_lines(details)
    if not lines:
        return 2
    return 2 + sum(max(1, (len(line) // 95) + 1) for line in lines)


def setup_schedule_sheet_layout(sheet) -> None:
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A1"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True

    column_widths = {
        "A": 3,
        "B": 24,
        "C": 24,
        "D": 24,
        "E": 23,
        "F": 23,
        "G": 23,
        "H": 23,
    }
    for column_letter, width in column_widths.items():
        sheet.column_dimensions[column_letter].width = width


def write_schedule_benefit_pages(sheet, pages: list[ScheduleBenefitPage]) -> None:
    if not pages:
        sheet["A1"] = "Schedule of Benefits"
        sheet["A2"] = "No booklet summaries were available to format."
        return

    styles = build_schedule_styles()
    current_row = 1
    for page in pages:
        current_row = write_schedule_benefit_page(sheet, page, current_row, styles)
        current_row += 3


def build_schedule_styles() -> dict[str, object]:
    thin_teal = Side(style="thin", color="C7E8E1")
    thin_dark = Side(style="thin", color="004B5A")
    return {
        "dark_teal": PatternFill("solid", fgColor="004B5A"),
        "section_teal": PatternFill("solid", fgColor="66C7B5"),
        "soft_teal": PatternFill("solid", fgColor="D7F0EC"),
        "body_grey": PatternFill("solid", fgColor="F2F2F2"),
        "white": PatternFill("solid", fgColor="FFFFFF"),
        "title_font": Font(name="Calibri", size=26, bold=False, color="004B5A"),
        "subtitle_font": Font(name="Calibri", size=15, bold=False, color="4FC3AD"),
        "dark_header_font": Font(name="Calibri", size=10, bold=True, color="FFFFFF"),
        "section_font": Font(name="Calibri", size=10, bold=True, color="FFFFFF"),
        "label_font": Font(name="Calibri", size=9, bold=True, color="004B5A"),
        "detail_font": Font(name="Calibri", size=9, color="004B5A"),
        "quiet_font": Font(name="Calibri", size=9, color="4F4F4F"),
        "section_border": Border(top=thin_dark, bottom=thin_dark),
        "detail_border": Border(bottom=thin_teal),
    }


def write_schedule_benefit_page(
    sheet,
    page: ScheduleBenefitPage,
    start_row: int,
    styles: dict[str, object],
) -> int:
    summary = page.summary
    client_name = display_summary_client_name(summary) or "[Enter Client Name]"
    class_name = clean_text(summary.class_name) or "Class"
    carrier_details = find_summary_section_detail(summary, "Carrier & Policy")
    eligibility_details = find_summary_section_detail(summary, "Eligibility Requirements")

    title_row = start_row
    apply_row_style(sheet, title_row, 1, 8, styles["white"])
    sheet.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=8)
    title_cell = sheet.cell(row=title_row, column=1)
    title_cell.value = client_name
    title_cell.font = styles["title_font"]
    title_cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[title_row].height = 36

    subtitle_row = title_row + 1
    apply_row_style(sheet, subtitle_row, 1, 8, styles["white"])
    sheet.merge_cells(start_row=subtitle_row, start_column=1, end_row=subtitle_row, end_column=8)
    subtitle_cell = sheet.cell(row=subtitle_row, column=1)
    subtitle_cell.value = f"Schedule of Benefits | {page.page_number} of {page.page_total}"
    subtitle_cell.font = styles["subtitle_font"]
    subtitle_cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[subtitle_row].height = 24

    header_row = subtitle_row + 2
    apply_row_style(sheet, header_row, 1, 8, styles["dark_teal"], styles["dark_header_font"])
    sheet.merge_cells(start_row=header_row, start_column=1, end_row=header_row, end_column=8)
    header_cell = sheet.cell(row=header_row, column=1)
    header_cell.value = client_name
    header_cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[header_row].height = 20

    info_row = header_row + 1
    apply_row_style(sheet, info_row, 1, 8, styles["soft_teal"], styles["detail_font"])
    write_merged_cell(sheet, info_row, 1, 2, f"Class: {class_name}", styles["label_font"], styles["soft_teal"])
    write_merged_cell(
        sheet,
        info_row,
        3,
        5,
        f"Eligibility: {clean_schedule_detail_for_one_line(eligibility_details)}",
        styles["detail_font"],
        styles["soft_teal"],
    )
    write_merged_cell(
        sheet,
        info_row,
        6,
        8,
        clean_schedule_detail_for_one_line(carrier_details),
        styles["detail_font"],
        styles["soft_teal"],
    )
    sheet.row_dimensions[info_row].height = 32

    current_row = info_row + 1
    for section_index, (section_name, details) in enumerate(page.sections, start=1):
        current_row = write_schedule_section(
            sheet,
            current_row,
            section_name,
            details,
            styles,
            add_top_gap=section_index > 1,
        )

    return current_row


def write_schedule_section(
    sheet,
    start_row: int,
    section_name: str,
    details: str,
    styles: dict[str, object],
    add_top_gap: bool,
) -> int:
    current_row = start_row
    if add_top_gap:
        current_row += 1

    apply_row_style(sheet, current_row, 1, 8, styles["section_teal"], styles["section_font"], styles["section_border"])
    sheet.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=8)
    section_cell = sheet.cell(row=current_row, column=1)
    section_cell.value = clean_text(section_name)
    section_cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[current_row].height = 20
    current_row += 1

    detail_lines = schedule_detail_lines(details)
    if not detail_lines:
        detail_lines = ["No details extracted."]

    for line in detail_lines:
        current_row = write_schedule_detail_line(sheet, current_row, line, styles)

    return current_row


def write_schedule_detail_line(sheet, row_index: int, line: str, styles: dict[str, object]) -> int:
    apply_row_style(sheet, row_index, 1, 8, styles["body_grey"], styles["detail_font"], styles["detail_border"])

    if line.startswith("- "):
        write_merged_cell(sheet, row_index, 2, 8, line, styles["detail_font"], styles["body_grey"])
    elif ":" in line:
        label, value = line.split(":", 1)
        write_merged_cell(sheet, row_index, 1, 2, f"{clean_text(label)}:", styles["label_font"], styles["body_grey"])
        write_merged_cell(sheet, row_index, 3, 8, clean_text(value), styles["detail_font"], styles["body_grey"])
    else:
        write_merged_cell(sheet, row_index, 1, 8, line, styles["detail_font"], styles["body_grey"])

    sheet.row_dimensions[row_index].height = max(18, min(54, 15 * max(1, (len(line) // 100) + 1)))
    return row_index + 1


def write_schedule_ppt_export_sheet(sheet, pages: list[ScheduleBenefitPage]) -> None:
    headers = [
        "PowerPointTitle",
        "ClientName",
        "SourceFile",
        "Carrier",
        "ClassName",
        "PageNumber",
        "PageTotal",
        "SectionOrder",
        "SectionName",
        "DetailText",
    ]
    for column_index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=column_index)
        cell.value = header
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="004B5A")
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    row_index = 2
    for page in pages:
        client_name = display_summary_client_name(page.summary)
        carrier_details = clean_schedule_detail_for_one_line(find_summary_section_detail(page.summary, "Carrier & Policy"))
        for section_order, (section_name, details) in enumerate(page.sections, start=1):
            values = [
                f"{client_name} - Schedule of Benefits - {page.page_number} of {page.page_total}",
                client_name,
                page.summary.source_file,
                carrier_details,
                page.summary.class_name,
                page.page_number,
                page.page_total,
                section_order,
                clean_text(section_name),
                "\n".join(schedule_detail_lines(details)),
            ]
            for column_index, value in enumerate(values, start=1):
                cell = sheet.cell(row=row_index, column=column_index)
                cell.value = value
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            row_index += 1

    if row_index == 2:
        sheet.cell(row=2, column=1).value = "No schedule rows available"
        row_index = 3

    table_ref = f"A1:{get_column_letter(len(headers))}{row_index - 1}"
    table = Table(displayName=PPT_EXPORT_TABLE, ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium4",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)
    sheet.freeze_panes = "A2"

    widths = [34, 28, 36, 28, 38, 12, 12, 14, 28, 90]
    for column_index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column_index)].width = width


def find_summary_section_detail(summary: EmployeeBenefitsSummary, section_name: str) -> str:
    wanted = clean_text(section_name).lower()
    for current_name, details in summary.rows:
        if clean_text(current_name).lower() == wanted:
            return details
    return ""


def schedule_detail_lines(details: str) -> list[str]:
    lines = [clean_text(line) for line in (details or "").splitlines()]
    return [line for line in lines if line]


def clean_schedule_detail_for_one_line(details: str) -> str:
    return " | ".join(schedule_detail_lines(details)) or ""


def apply_row_style(sheet, row_index: int, start_column: int, end_column: int, fill, font=None, border=None) -> None:
    for column_index in range(start_column, end_column + 1):
        cell = sheet.cell(row=row_index, column=column_index)
        cell.fill = fill
        if font is not None:
            cell.font = font
        if border is not None:
            cell.border = border
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def write_merged_cell(sheet, row_index: int, start_column: int, end_column: int, value: str, font, fill) -> None:
    for column_index in range(start_column, end_column + 1):
        cell = sheet.cell(row=row_index, column=column_index)
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    if end_column > start_column:
        sheet.merge_cells(start_row=row_index, start_column=start_column, end_row=row_index, end_column=end_column)
    cell = sheet.cell(row=row_index, column=start_column)
    cell.value = value
    cell.font = font
    cell.alignment = Alignment(wrap_text=True, vertical="top")


def write_separate_review_workbook(
    master_path: Path,
    review_path: Path | None,
    rows: list[ExtractionRow],
    summaries: list[EmployeeBenefitsSummary],
) -> Path:
    path = review_path or master_path.with_name("Booklet_Extraction_Review.xlsx")
    path.parent.mkdir(parents=True, exist_ok=True)

    workbook = build_review_workbook(rows, summaries)
    try:
        workbook.save(path)
        return path
    except PermissionError:
        fallback_path = path.with_name(f"{path.stem}_{datetime.now():%Y%m%d-%H%M%S}{path.suffix}")
        workbook.save(fallback_path)
        return fallback_path


def build_review_workbook(rows: list[ExtractionRow], summaries: list[EmployeeBenefitsSummary]):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = REVIEW_SHEET
    for column_index, column_name in enumerate(REVIEW_COLUMNS, start=1):
        sheet.cell(row=1, column=column_index).value = column_name

    header_map = worksheet_header_map(sheet)
    write_review_sheet(sheet, rows, summaries)
    autosize_columns(sheet, header_map)
    return workbook


def append_processing_log(
    sheet,
    processed_sources: list[ProcessedSource],
    errors: list[ProcessingError],
    counts_by_source: dict[str, dict[str, int]],
) -> None:
    header_map = worksheet_header_map(sheet)
    processed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for source in processed_sources:
        counts = counts_by_source.get(source.source_file, {"added": 0, "updated": 0})
        append_log_row(
            sheet,
            header_map,
            {
                "ProcessedAt": processed_at,
                "ClientName": source.client_name,
                "Carrier": source.carrier,
                "SourceFile": source.source_file,
                "Module": MODULE_NAME,
                "Status": source.status,
                "RowsAdded": str(counts["added"]),
                "RowsUpdated": str(counts["updated"]),
                "ErrorMessage": "",
            },
        )

    for error in errors:
        append_log_row(
            sheet,
            header_map,
            {
                "ProcessedAt": processed_at,
                "ClientName": error.client_name,
                "Carrier": error.carrier,
                "SourceFile": error.source_file,
                "Module": MODULE_NAME,
                "Status": "Error",
                "RowsAdded": "0",
                "RowsUpdated": "0",
                "ErrorMessage": error.error_message,
            },
        )


def append_log_row(sheet, header_map: dict[str, int], data: dict[str, str]) -> None:
    row_index = sheet.max_row + 1
    for column_name in LOG_COLUMNS:
        sheet.cell(row=row_index, column=header_map[column_name]).value = data.get(column_name, "")


def autosize_columns(sheet, header_map: dict[str, int]) -> None:
    for column_name, column_index in header_map.items():
        max_length = len(column_name)
        for cell in sheet.iter_cols(min_col=column_index, max_col=column_index, min_row=2, values_only=True):
            for value in cell:
                max_length = max(max_length, min(len(clean_cell(value)), 80))
        sheet.column_dimensions[sheet.cell(row=1, column=column_index).column_letter].width = min(max_length + 2, 60)


def worksheet_header_map(sheet) -> dict[str, int]:
    headers = {}
    for column_index in range(1, sheet.max_column + 1):
        value = sheet.cell(row=1, column=column_index).value
        if value:
            headers[str(value).strip()] = column_index
    return headers


def plan_row_to_dict(row: ExtractionRow) -> dict[str, str]:
    return {
        "ClientID": row.client_id,
        "ClientName": row.client_name,
        "PolicyNumber": row.policy_number,
        "Carrier": row.carrier,
        "ClassName": row.class_name,
        "BenefitCategory": row.benefit_category or row.benefit,
        "Benefit": row.benefit,
        "CoverageDescription": row.coverage_description,
        "Coinsurance": row.coinsurance,
        "Maximum": row.maximum,
        "Frequency": row.frequency,
        "WaitingPeriod": row.waiting_period,
        "TerminationAge": row.termination_age,
        "Notes": row.notes,
        "SourceFile": row.source_file,
        "SourceSection": row.source_section,
        "SourcePage": row.source_page,
        "LastUpdated": row.last_updated,
    }


def review_row_to_dict(row: ExtractionRow) -> dict[str, str]:
    return {
        "ClientID": row.client_id,
        "SourceFile": row.source_file,
        "ClientName": row.client_name,
        "PolicyNumber": row.policy_number,
        "Carrier": row.carrier,
        "ClassName": row.class_name,
        "BenefitCategory": row.benefit_category or row.benefit,
        "BenefitName": row.benefit,
        "Benefit": row.benefit,
        "CoverageDescription": row.coverage_description,
        "Coinsurance": row.coinsurance,
        "Maximum": row.maximum,
        "Frequency": row.frequency,
        "WaitingPeriod": row.waiting_period,
        "TerminationAge": row.termination_age,
        "Confidence": row.confidence,
        "ReviewFlag": row.review_flag,
    }


def split_sentences(text: str) -> list[str]:
    return [clean_text(part) for part in re.split(r"(?<=[.!?])\s+", text) if clean_text(part)]


def unique_regex_matches(text: str, patterns: Iterable[str]) -> list[str]:
    matches: list[str] = []
    for pattern in patterns:
        matches.extend(match.group(0) for match in re.finditer(pattern, text, flags=re.IGNORECASE))
    return unique_in_order(clean_text(match) for match in matches)


def unique_regex_group_matches(text: str, patterns: Iterable[str]) -> list[str]:
    matches: list[str] = []
    for pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            value = match.group(1) if match.lastindex else match.group(0)
            matches.append(value)
    return unique_in_order(clean_text(match) for match in matches)


def unique_in_order(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    unique_values: list[str] = []
    for value in values:
        cleaned = clean_text(value)
        key = cleaned.lower()
        if cleaned and key not in seen:
            seen.add(key)
            unique_values.append(cleaned)
    return unique_values


def clean_money(value: str) -> str:
    cleaned = re.sub(r"\s+", "", value or "")
    normalized = cleaned.replace(",", "")
    match = re.fullmatch(r"\$(\d+)(\.\d{2})?", normalized)
    if match:
        dollars = int(match.group(1))
        cents = match.group(2) or ""
        return f"${dollars:,}{cents}"
    return cleaned


def clean_client_name(value: str) -> str:
    cleaned = clean_text(value)
    cleaned = re.split(r"\[\[PAGE\s+\d+\]\]", cleaned, maxsplit=1, flags=re.IGNORECASE)[0]
    cleaned = re.split(r"\bIntroduction\s+Policyholder\b", cleaned, maxsplit=1, flags=re.IGNORECASE)[0]
    cleaned = re.sub(r"^POLICYHOLDER\s*:\s*", "", cleaned, flags=re.IGNORECASE)
    dba_match = re.search(r"\b(?:D\s*/?\s*B\s*/?\s*A|DBA|doing business as)\b\s*(.+)$", cleaned, flags=re.IGNORECASE)
    used_dba_name = bool(dba_match)
    if dba_match:
        cleaned = clean_text(dba_match.group(1))
    cleaned = re.split(
        r"\b(?:Group|Policy|Contract|Benefits|Booklet|Effective|Prepared)\b",
        cleaned,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0]
    if used_dba_name and cleaned.isupper():
        cleaned = cleaned.title()
    return cleaned.strip(" :-,.") or value.strip(" :-,.")


def clean_cell(value) -> str:
    if value is None:
        return ""
    return clean_text(str(value))


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def normalize_for_search(value: str) -> str:
    normalized = value.lower()
    normalized = normalized.replace("&", " and ")
    normalized = normalized.replace("-", " ")
    normalized = re.sub(r"[^a-z0-9%$]+", " ", normalized)
    return re.sub(r"\s+", " ", normalized).strip()


def shorten(value: str, max_length: int) -> str:
    cleaned = clean_text(value)
    if len(cleaned) <= max_length:
        return cleaned
    return cleaned[:max_length].rstrip()


if __name__ == "__main__":
    raise SystemExit(main())
