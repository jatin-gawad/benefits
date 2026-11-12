from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation


CANONICAL_LABELED_DENTAL_RATE_MAPPER_VERSION = "CANONICAL_LABELED_DENTAL_RATE_MAPPER_V1"
_MAX_LABEL_TO_PERCENTAGE_DISTANCE = 600


@dataclass(frozen=True)
class LabeledDentalRate:
    provision_code: str
    provision_name: str
    benefit_name: str
    reimbursement_percent: Decimal
    reimbursement_text: str
    source_label: str
    source_text: str
    source_page: int
    maximum: str = ""
    eligibility: str = ""
    confidence: str = "High"


@dataclass(frozen=True)
class LabeledDentalMaximum:
    provision_code: str
    provision_name: str
    benefit_name: str
    maximum: str
    source_label: str
    source_text: str
    source_page: int
    confidence: str = "High"


@dataclass(frozen=True)
class LabeledDentalCoverageStatus:
    provision_code: str
    provision_name: str
    benefit_name: str
    coverage_status: str
    source_label: str
    source_text: str
    source_page: int
    confidence: str = "High"


@dataclass(frozen=True)
class _DentalProvision:
    code: str
    name: str
    benefit_name: str
    label_patterns: tuple[str, ...]


_PROVISIONS = (
    _DentalProvision(
        "preventive",
        "Preventive Services",
        "Dental Preventive Services",
        (
            r"(?:Dental\s+)?Preventive\s+(?:Services|Procedures|Care)",
            r"Preventive\s+procedures",
        ),
    ),
    _DentalProvision(
        "basic",
        "Basic Services",
        "Dental Basic Services",
        (
            r"Level\s*(?:I|1)\s*[-:]\s*Basic(?:\s+Services)?",
            r"Basic\s+Dental\s+Care",
            r"Basic\s+Dental\s+Services",
            r"Basic\s+Restorative(?:\s*,?\s*Periodontics?\s*/\s*Endodontics?)?",
            r"(?:Dental\s+)?Basic\s+(?:Services|Procedures|Care|Coverage)",
            r"Basic\s+procedures",
        ),
    ),
    _DentalProvision(
        "endodontic_periodontic",
        "Endodontic / Periodontic Services",
        "Dental Endodontic / Periodontic Services",
        (
            r"(?:Endodontic\s*/\s*Periodontic|Periodontic\s*/\s*Endodontic)\s+Services",
        ),
    ),
    _DentalProvision(
        "routine",
        "Routine Dental Care",
        "Dental Routine Care",
        (
            r"Routine\s+Dental\s+Care",
            r"(?:Dental\s+)?Routine\s+(?:Services|Procedures|Care)",
        ),
    ),
    _DentalProvision(
        "dentures",
        "Dentures",
        "Dental Dentures",
        (
            r"Level\s*(?:III|3)\s*[-:]\s*Dentures?",
            r"Dentures?\s+Services",
        ),
    ),
    _DentalProvision(
        "major",
        "Major Restorative Services",
        "Dental Major Restorative Services",
        (
            r"Level\s*(?:IV|4)\s*[-:]\s*(?:Major\s+)?Restorative\s+Services",
            r"Dental\s+Restorative\s+Services",
            r"Major\s+Restorative\s+Services",
            r"Major\s+Restorative",
            r"Major\s+Dental\s+Services",
            r"Major\s+(?:Services|Procedures|Coverage)",
            r"Major\s+procedures",
        ),
    ),
    _DentalProvision(
        "orthodontic",
        "Orthodontic Services",
        "Dental Orthodontic Services",
        (
            r"Level\s*(?:V|5)\s*[-:]\s*Orthodontics?",
            r"Orthodontic\s+(?:Care|Services|Procedures|Coverage|Treatment)",
            r"Orthodontics",
        ),
    ),
    _DentalProvision(
        "accidental_injury",
        "Accidental Dental Injury",
        "Accidental Dental Injury",
        (
            r"Accidental\s+Dental\s+Injury(?:\s+(?:Coverage|Treatment))?",
        ),
    ),
    _DentalProvision(
        "other_treatment",
        "Other Dental Treatment",
        "Dental Other Treatment",
        (
            r"All\s+Other\s+Treatment",
        ),
    ),
)

_PROVISION_BY_CODE = {provision.code: provision for provision in _PROVISIONS}
_PAGE_MARKER = re.compile(r"\[\[PAGE\s+(\d+)\]\]", re.IGNORECASE)
_PERCENTAGE = re.compile(
    r"(?<![\d.])(?P<number>100(?:\.0+)?|\d{1,2}(?:\.\d+)?)\s*(?:%|per\s+cent)(?![a-z0-9])",
    re.IGNORECASE,
)
_COMBINED_MAXIMUM = re.compile(
    r"(?:to\s+a\s+)?(?:combined\s+)?maximum\s+of\s+"
    r"(?P<amount>\$\s?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?)"
    r"(?P<period>\s+per\s+(?:person\s+per\s+)?(?:calendar\s+|benefit\s+)?year|\s+per\s+lifetime|\s+lifetime)?",
    re.IGNORECASE,
)
_MONEY_MAXIMUM = re.compile(
    r"(?P<value>\$\s?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?"
    r"(?:\s+per\s+(?:person\s+per\s+)?(?:calendar\s+|benefit\s+)?year|\s+per\s+lifetime|\s+lifetime)?)",
    re.IGNORECASE,
)
_ELIGIBILITY_PATTERNS = (
    r"only\s+for\s+(?:dependent\s+)?children(?:\s+under\s+age\s+\d{1,2})?",
    r"(?:dependent\s+)?children\s+under\s+age\s+\d{1,2}",
    r"adults?\s+(?:&|and)\s+dependent\s+children",
)
_TYPE_DEFINITION = re.compile(
    r"\bType\s+(?P<type>[A-Z0-9]+)\s*[-\u2013\u2014:]\s*"
    r"(?P<label>Basic\s+services?|Major\s+restorative\s+services?|Orthodontic\s+services?)\b",
    re.IGNORECASE,
)
_TYPE_RATE = re.compile(
    r"\bType\s+(?P<type>[A-Z0-9]+)\s*:?\s*"
    r"(?P<number>100(?:\.0+)?|\d{1,2}(?:\.\d+)?)\s*(?:%|per\s+cent)(?![a-z0-9])",
    re.IGNORECASE,
)
_TYPE_MAXIMUM = re.compile(
    r"(?P<phrase>(?:(?:annual\s+calendar\s+year)|(?:calendar\s+year)|annual|lifetime)\s+maximum\s+for\s+"
    r"(?P<types>Type\s+[A-Z0-9]+(?:\s+(?:and|&)\s+Type\s+[A-Z0-9]+)*(?:\s+combined)?)\s*[:\-]?\s*"
    r"\$\s?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?)",
    re.IGNORECASE,
)
_BENEFIT_PERIOD_MAXIMUM = re.compile(
    r"\bBenefit\s+Period\s+Maximum\s*:?\s*"
    r"(?P<labels>.{1,260}?)\s*"
    r"(?P<amount>\$\s?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?)"
    r"(?P<tail>\s+per\s+(?:insured|person))?",
    re.IGNORECASE | re.DOTALL,
)
_SHARED_MAXIMUM = re.compile(
    r"\bMaximum\s+Amount\s+(?P<labels>.{1,280}?)\s+Covered\W*"
    r"(?P<maximum>No\s+maximum|combined\s+maximum\s+of\s+"
    r"\$\s?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?"
    r"(?:\s+per\s+(?:insured\s+person|person))?(?:\s+per\s+(?:calendar\s+year|benefit\s+period))?)",
    re.IGNORECASE | re.DOTALL,
)
_SHARED_MAXIMUM_COVERED_FIRST = re.compile(
    r"\bMaximum\s+Amount\s+Covered\W*"
    r"(?P<labels>.{1,280}?)\s*"
    r"(?P<maximum>No\s+maximum|combined\s+maximum\s+of\s+"
    r"\$\s?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?"
    r"(?:\s+per\s+(?:insured\s+person|person))?(?:\s+per\s+(?:calendar\s+year|benefit\s+period))?)",
    re.IGNORECASE | re.DOTALL,
)
_NOT_INCLUDED = re.compile(r"\b(?:Not\s+Included|Not\s+Covered|No\s+Coverage)\b", re.IGNORECASE)


def map_labeled_dental_rates(text: str) -> list[LabeledDentalRate]:
    """Map only percentages that have an explicit nearby dental provision label."""

    candidates: list[LabeledDentalRate] = []
    labeled_maximums = {
        maximum.provision_code: maximum
        for maximum in map_labeled_dental_maximums(text)
    }
    for page_number, page_text in _split_pages(text):
        if not _looks_like_dental_page(page_text):
            continue

        labels = _find_labels(page_text)
        for index, (start, end, provision, source_label) in enumerate(labels):
            segment_end = labels[index + 1][0] if index + 1 < len(labels) else len(page_text)
            segment = page_text[start:segment_end]
            percentage_matches = [
                match
                for match in _PERCENTAGE.finditer(segment)
                if match.start() <= _MAX_LABEL_TO_PERCENTAGE_DISTANCE
                and _is_reimbursement_percentage(segment, match)
            ]
            percentages = _distinct_percentages(percentage_matches)
            if len(percentages) != 1:
                continue

            percentage = percentages[0]
            rate_match = next(
                match
                for match in percentage_matches
                if _parse_decimal(match.group("number")) == percentage
            )
            candidates.append(
                LabeledDentalRate(
                    provision_code=provision.code,
                    provision_name=provision.name,
                    benefit_name=provision.benefit_name,
                    reimbursement_percent=percentage,
                    reimbursement_text=format_percentage(percentage),
                    source_label=_clean_text(source_label),
                    source_text=_build_evidence(segment, source_label, rate_match),
                    source_page=page_number,
                    maximum=(
                        _extract_combined_maximum(segment, rate_match)
                        or (labeled_maximums.get(provision.code).maximum if provision.code in labeled_maximums else "")
                    ),
                    eligibility=_extract_eligibility(segment, rate_match),
                )
            )

    candidates.extend(_map_explicit_type_rates(text))

    resolved: list[LabeledDentalRate] = []
    for provision in _PROVISIONS:
        provision_candidates = [candidate for candidate in candidates if candidate.provision_code == provision.code]
        distinct_rates = {candidate.reimbursement_percent for candidate in provision_candidates}
        if len(distinct_rates) == 1:
            resolved.append(provision_candidates[0])
    return resolved


def _is_reimbursement_percentage(segment: str, match: re.Match[str]) -> bool:
    """Exclude treatment-payment percentages that are not reimbursement rates."""

    before = _clean_text(segment[max(0, match.start() - 120) : match.start()]).lower()
    after = _clean_text(segment[match.end() : match.end() + 160]).lower()
    if re.search(r"\bmaximum\s+amount\s+of\s*$", before) and re.match(
        r"^of\s+the\s+(?:total\s+)?(?:estimated\s+)?(?:treatment\s+)?cost\b",
        after,
    ):
        return False
    return True


def map_labeled_dental_maximums(text: str) -> list[LabeledDentalMaximum]:
    """Map maximums only when a dental treatment label and value are explicit."""

    candidates: list[LabeledDentalMaximum] = []
    for page_number, page_text in _split_pages(text):
        if not _looks_like_dental_page(page_text):
            continue
        labels = _find_labels(page_text)
        for index, (start, _end, provision, source_label) in enumerate(labels):
            segment_end = labels[index + 1][0] if index + 1 < len(labels) else len(page_text)
            segment = page_text[start:segment_end]
            maximum = _extract_labeled_maximum_value(segment)
            if not maximum:
                continue
            candidates.append(
                LabeledDentalMaximum(
                    provision_code=provision.code,
                    provision_name=provision.name,
                    benefit_name=provision.benefit_name,
                    maximum=maximum,
                    source_label=_clean_text(source_label),
                    source_text=_clean_text(segment)[:1400],
                    source_page=page_number,
                )
            )

    combined_candidates = [
        *_map_combined_benefit_period_maximums(text),
        *_map_shared_dental_maximums(text),
    ]
    resolved: list[LabeledDentalMaximum] = []
    for provision in _PROVISIONS:
        provision_combined = [
            candidate
            for candidate in combined_candidates
            if candidate.provision_code == provision.code
        ]
        combined_values = {_normalize(candidate.maximum) for candidate in provision_combined}
        if provision_combined:
            if len(combined_values) == 1:
                resolved.append(provision_combined[0])
            continue
        provision_candidates = [candidate for candidate in candidates if candidate.provision_code == provision.code]
        distinct_maximums = {_normalize(candidate.maximum) for candidate in provision_candidates}
        if len(distinct_maximums) == 1:
            resolved.append(provision_candidates[0])
    return resolved


def map_labeled_dental_coverage_statuses(text: str) -> list[LabeledDentalCoverageStatus]:
    """Map explicit nearby exclusions such as ``Orthodontic Care - Not Included``."""

    candidates: list[LabeledDentalCoverageStatus] = []
    for page_number, page_text in _split_pages(text):
        if not _looks_like_dental_page(page_text):
            continue
        labels = _find_labels(page_text)
        for index, (start, _end, provision, source_label) in enumerate(labels):
            segment_end = labels[index + 1][0] if index + 1 < len(labels) else len(page_text)
            segment = page_text[start:segment_end]
            status_match = _NOT_INCLUDED.search(segment[:_MAX_LABEL_TO_PERCENTAGE_DISTANCE])
            if not status_match:
                continue
            candidates.append(
                LabeledDentalCoverageStatus(
                    provision_code=provision.code,
                    provision_name=provision.name,
                    benefit_name=provision.benefit_name,
                    coverage_status="Not Included",
                    source_label=_clean_text(source_label),
                    source_text=_clean_text(segment)[:1400],
                    source_page=page_number,
                )
            )

    resolved: list[LabeledDentalCoverageStatus] = []
    for provision in _PROVISIONS:
        provision_candidates = [
            candidate
            for candidate in candidates
            if candidate.provision_code == provision.code
        ]
        distinct_statuses = {
            _normalize(candidate.coverage_status)
            for candidate in provision_candidates
        }
        if len(distinct_statuses) == 1:
            resolved.append(provision_candidates[0])
    return resolved


def _map_combined_benefit_period_maximums(text: str) -> list[LabeledDentalMaximum]:
    candidates: list[LabeledDentalMaximum] = []
    for page_number, page_text in _split_pages(text):
        for match in _BENEFIT_PERIOD_MAXIMUM.finditer(page_text):
            labels = match.group("labels")
            codes: list[str] = []
            endodontic_is_part_of_basic = (
                re.search(r"\bBasic\s+Restorative\b", labels, flags=re.IGNORECASE) is not None
                and re.search(r"\bEndodont", labels, flags=re.IGNORECASE) is not None
                and re.search(r"\bPeriodont", labels, flags=re.IGNORECASE) is not None
            )
            for provision in _PROVISIONS:
                if provision.code == "endodontic_periodontic" and endodontic_is_part_of_basic:
                    continue
                if any(re.search(pattern, labels, flags=re.IGNORECASE) for pattern in provision.label_patterns):
                    codes.append(provision.code)
            if re.search(r"\bBasic\s+and\s+Major\s+Dental\s+Services\b", labels, flags=re.IGNORECASE):
                codes.extend(["basic", "major"])
            codes = list(dict.fromkeys(codes))
            if len(codes) < 2:
                continue

            amount = _clean_text(f"{match.group('amount')}{match.group('tail') or ''}")
            maximum = f"Combined {amount} per benefit period"
            source_text = _clean_text(match.group(0))[:1400]
            for code in codes:
                provision = _PROVISION_BY_CODE[code]
                candidates.append(
                    LabeledDentalMaximum(
                        provision_code=code,
                        provision_name=provision.name,
                        benefit_name=provision.benefit_name,
                        maximum=maximum,
                        source_label="Benefit Period Maximum",
                        source_text=source_text,
                        source_page=page_number,
                    )
                )
    return candidates


def _map_shared_dental_maximums(text: str) -> list[LabeledDentalMaximum]:
    candidates: list[LabeledDentalMaximum] = []
    for page_number, page_text in _split_pages(text):
        cleaned_page = _clean_text(page_text)
        matches = [
            *list(_SHARED_MAXIMUM.finditer(cleaned_page)),
            *list(_SHARED_MAXIMUM_COVERED_FIRST.finditer(cleaned_page)),
        ]
        for match in matches:
            labels = match.group("labels")
            codes = [
                provision.code
                for provision in _PROVISIONS
                if any(
                    re.search(pattern, labels, flags=re.IGNORECASE)
                    for pattern in provision.label_patterns
                )
            ]
            codes = list(dict.fromkeys(codes))
            if len(codes) < 2:
                continue

            maximum = _clean_text(match.group("maximum"))
            if maximum.lower().startswith("combined"):
                maximum = maximum[0].upper() + maximum[1:]
            source_text = _clean_text(match.group(0))[:1400]
            for code in codes:
                provision = _PROVISION_BY_CODE[code]
                candidates.append(
                    LabeledDentalMaximum(
                        provision_code=code,
                        provision_name=provision.name,
                        benefit_name=provision.benefit_name,
                        maximum=maximum,
                        source_label="Maximum Amount",
                        source_text=source_text,
                        source_page=page_number,
                    )
                )
    return candidates


def canonical_dental_provision_code(value: str) -> str:
    """Return one canonical dental provision code, or blank when wording is mixed/ambiguous."""

    normalized = _normalize(value)
    if not normalized:
        return ""

    level_mappings = (
        (r"\blevel\s*(?:v|5)\b", "orthodontic"),
        (r"\blevel\s*(?:iv|4)\b", "major"),
        (r"\blevel\s*(?:iii|3)\b", "dentures"),
        (r"\blevel\s*(?:i|1)\b", "basic"),
    )
    for pattern, code in level_mappings:
        if re.search(pattern, normalized, flags=re.IGNORECASE):
            return code

    if "basic restorative" in normalized and "endodont" in normalized and "periodont" in normalized:
        return "basic"

    matches: set[str] = set()
    if "orthodont" in normalized or re.search(r"\bortho\b", normalized):
        matches.add("orthodontic")
    if "accidental dental injury" in normalized or "dental accidental injury" in normalized:
        matches.add("accidental_injury")
    if "all other treatment" in normalized or "other dental treatment" in normalized:
        matches.add("other_treatment")
    if "denture" in normalized:
        matches.add("dentures")
    if "endodont" in normalized and "periodont" in normalized:
        matches.add("endodontic_periodontic")
    if "major" in normalized or "dental restorative" in normalized:
        matches.add("major")
    if "routine dental" in normalized or "dental routine" in normalized:
        matches.add("routine")
    if "preventive" in normalized or "preventative" in normalized:
        matches.add("preventive")
    if re.search(r"\bbasic\b", normalized):
        matches.add("basic")
    return next(iter(matches)) if len(matches) == 1 else ""


def dental_provision_name(code: str) -> str:
    provision = _PROVISION_BY_CODE.get(code)
    return provision.name if provision else ""


def dental_coinsurance_provision_code(value: str) -> str:
    code = value if value in _PROVISION_BY_CODE else canonical_dental_provision_code(value)
    return f"coinsurance.dental_{code}" if code else ""


def dental_maximum_provision_code(value: str) -> str:
    code = value if value in _PROVISION_BY_CODE else canonical_dental_provision_code(value)
    return f"maximum.dental_{code}" if code else ""


def format_percentage(value: Decimal) -> str:
    return f"{value.quantize(Decimal('1'))}%" if value == value.to_integral() else f"{value.normalize()}%"


def _split_pages(text: str) -> list[tuple[int, str]]:
    source = text or ""
    matches = list(_PAGE_MARKER.finditer(source))
    if not matches:
        return [(1, source)] if source.strip() else []

    pages: list[tuple[int, str]] = []
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(source)
        pages.append((int(match.group(1)), source[start:end]))
    return pages


def _map_explicit_type_rates(text: str) -> list[LabeledDentalRate]:
    pages = _split_pages(text)
    definitions: dict[str, list[tuple[str, str, str]]] = {}
    for _page_number, page_text in pages:
        matches = list(_TYPE_DEFINITION.finditer(page_text))
        for index, match in enumerate(matches):
            type_code = match.group("type").upper()
            provision_code = canonical_dental_provision_code(match.group("label"))
            if not provision_code:
                continue
            segment_end = matches[index + 1].start() if index + 1 < len(matches) else min(
                len(page_text),
                match.end() + 500,
            )
            definition_segment = page_text[match.start():segment_end]
            definitions.setdefault(type_code, []).append(
                (
                    provision_code,
                    _clean_text(match.group(0)),
                    _extract_eligibility_text(definition_segment),
                )
            )

    resolved_definitions: dict[str, tuple[str, str, str]] = {}
    for type_code, matches in definitions.items():
        provision_codes = {match[0] for match in matches}
        if len(provision_codes) == 1:
            resolved_definitions[type_code] = matches[0]
    if not resolved_definitions:
        return []

    maximums: dict[str, str] = {}
    for _page_number, page_text in pages:
        for match in _TYPE_MAXIMUM.finditer(page_text):
            maximum = _clean_text(match.group("phrase"))
            for type_code in re.findall(r"\bType\s+([A-Z0-9]+)\b", match.group("types"), flags=re.IGNORECASE):
                maximums[type_code.upper()] = maximum

    rates: list[LabeledDentalRate] = []
    for page_number, page_text in pages:
        for match in _TYPE_RATE.finditer(page_text):
            type_code = match.group("type").upper()
            definition = resolved_definitions.get(type_code)
            percentage = _parse_decimal(match.group("number"))
            if not definition or percentage is None or percentage < 0 or percentage > 100:
                continue

            provision_code, definition_label, eligibility = definition
            provision = _PROVISION_BY_CODE[provision_code]
            rate_line = _line_text(page_text, match)
            evidence_parts = [definition_label, rate_line, maximums.get(type_code, "")]
            if eligibility:
                evidence_parts.append(f"Eligibility: {eligibility}")
            rates.append(
                LabeledDentalRate(
                    provision_code=provision.code,
                    provision_name=provision.name,
                    benefit_name=provision.benefit_name,
                    reimbursement_percent=percentage,
                    reimbursement_text=format_percentage(percentage),
                    source_label=definition_label,
                    source_text="; ".join(part for part in evidence_parts if part),
                    source_page=page_number,
                    maximum=maximums.get(type_code, ""),
                    eligibility=eligibility,
                )
            )
    return rates


def _looks_like_dental_page(page_text: str) -> bool:
    normalized = _normalize(page_text)
    return any(
        token in normalized
        for token in (
            "dental",
            "orthodont",
            "level iii dentures",
            "level iv major restorative",
            "major restorative services",
        )
    )


def _find_labels(page_text: str) -> list[tuple[int, int, _DentalProvision, str]]:
    labels: list[tuple[int, int, _DentalProvision, str]] = []
    for provision in _PROVISIONS:
        for pattern in provision.label_patterns:
            for match in re.finditer(rf"\b(?:{pattern})\b", page_text, flags=re.IGNORECASE):
                if _is_heading_candidate(page_text, match):
                    labels.append((match.start(), match.end(), provision, match.group(0)))

    labels.sort(key=lambda item: (item[0], -(item[1] - item[0])))
    deduplicated: list[tuple[int, int, _DentalProvision, str]] = []
    for label in labels:
        if deduplicated and label[0] < deduplicated[-1][1]:
            continue
        deduplicated.append(label)
    return deduplicated


def _is_heading_candidate(page_text: str, match: re.Match[str]) -> bool:
    line_start = page_text.rfind("\n", 0, match.start()) + 1
    line_end = page_text.find("\n", match.end())
    if line_end < 0:
        line_end = len(page_text)

    prefix = _clean_text(page_text[line_start:match.start()]).strip("-*|:.; ")
    tail = _clean_text(page_text[match.end():line_end])
    label = _normalize(match.group(0))
    if re.match(r"^(?:and|or)\b", tail, flags=re.IGNORECASE):
        return False
    if label.startswith("level "):
        return True
    if not prefix:
        return True
    if _PERCENTAGE.search(page_text[match.end():min(line_end, match.end() + 180)]):
        return True
    if any(
        phrase in label
        for phrase in (
            "basic dental care",
            "routine dental care",
            "dental restorative services",
            "orthodontic care",
            "preventive procedures",
            "basic procedures",
            "major procedures",
            "orthodontic procedures",
        )
    ):
        return True
    return False


def _distinct_percentages(matches: list[re.Match[str]]) -> list[Decimal]:
    values: list[Decimal] = []
    for match in matches:
        value = _parse_decimal(match.group("number"))
        if value is None or value < 0 or value > 100 or value in values:
            continue
        values.append(value)
    return values


def _parse_decimal(value: str) -> Decimal | None:
    try:
        return Decimal(value)
    except (InvalidOperation, TypeError, ValueError):
        return None


def _build_evidence(segment: str, source_label: str, rate_match: re.Match[str]) -> str:
    start = max(0, rate_match.start() - 500)
    end = min(len(segment), rate_match.end() + 500)
    evidence = _clean_text(segment[start:end])
    label = _clean_text(source_label)
    if label and label.lower() not in evidence.lower():
        evidence = f"{label}: {evidence}"
    return evidence[:1400].strip()


def _extract_combined_maximum(segment: str, rate_match: re.Match[str]) -> str:
    nearby = segment[rate_match.start():min(len(segment), rate_match.end() + 220)]
    match = _COMBINED_MAXIMUM.search(nearby)
    if not match:
        return ""
    return _clean_text(match.group(0)).capitalize()


def _extract_labeled_maximum_value(segment: str) -> str:
    combined = _COMBINED_MAXIMUM.search(segment)
    if combined:
        return _clean_text(combined.group(0))
    money = _MONEY_MAXIMUM.search(segment)
    if money:
        return _clean_text(money.group("value"))
    unlimited = re.search(r"\bUnlimited\b", segment, flags=re.IGNORECASE)
    if unlimited:
        return _clean_text(unlimited.group(0))
    no_maximum = re.search(r"\bNo\s+maximum\b", segment, flags=re.IGNORECASE)
    return _clean_text(no_maximum.group(0)) if no_maximum else ""


def _extract_eligibility(segment: str, rate_match: re.Match[str]) -> str:
    nearby = segment[max(0, rate_match.start() - 80):min(len(segment), rate_match.end() + 260)]
    return _extract_eligibility_text(nearby)


def _extract_eligibility_text(value: str) -> str:
    for pattern in _ELIGIBILITY_PATTERNS:
        match = re.search(pattern, value, flags=re.IGNORECASE)
        if match:
            return _clean_text(match.group(0))
    return ""


def _line_text(value: str, match: re.Match[str]) -> str:
    line_start = value.rfind("\n", 0, match.start()) + 1
    line_end = value.find("\n", match.end())
    if line_end < 0:
        line_end = len(value)
    return _clean_text(value[line_start:line_end])


def _normalize(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (value or "").lower()).strip()


def _clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()
