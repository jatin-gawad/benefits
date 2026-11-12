# Canonical Claims Model Classification

| Model | Classification | Reason |
|---|---|---|
| `CanonicalClaimsImportEnvelope` | Import DTO | The single payload entering the canonical write boundary. |
| `CanonicalClaimsReport` | Domain model | Aggregate ownership and report identity shared by facts. |
| `CanonicalClaimsPeriod` | Value object | Normalized immutable reporting-period meaning. |
| Financial, monthly, category, distribution, benchmark, and utilization facts | Domain models | Typed carrier-neutral claims facts used by future consumers. |
| `CanonicalClaimsFactRevision` | Domain audit model | Retains both source references and before/after adjudication values when one canonical natural key changes. |
| Drug, indication, provider, large-claim, disability, life facts | Domain models | Specialized contract responsibilities; not persisted in Phase 1. |
| `CanonicalClaimsSourceEvidence` | Value object | Immutable lineage attached to a fact. |
| `CanonicalClaimsValidationResult` | Read model | Reports validation outcome without becoming a second validation engine. |
| `CanonicalClaimsParityResult` | Read model | Reports migration comparison and review status. |
| `CanonicalClaimsImportResult` | Read model | Returns transaction, retry, validation, and parity outcome. |
| `CanonicalClaimsReportReadModel` | Read model | Carrier-neutral output of `CanonicalClaimsQueryService`. |
| `ManulifeClaimsCanonicalMapper` | Compatibility adapter | Converts an existing working parsed result without changing extraction. |
| `CanadaLifeClaimsCanonicalMapper` | Compatibility adapter | Converts `CarrierPdfParseResult` produced by the existing Canada Life parser; it does not parse or persist. |
| `SunLifeClaimsCanonicalMapper` | Compatibility adapter | Converts the existing claims-experience parsed result into up to four content-dated canonical period reports; it does not open PDFs or persist. |
| `CanadaLifeClaimsTerminology` | Compatibility adapter | Holds explicit carrier terminology aliases before the canonical boundary; unknown terms remain ambiguous. |

The corresponding SQLite rows are database entities implemented by the migration/persistence layer, not reused as domain objects.

Phase 2 adds `CanonicalClaimsMappingClassification` as a value object used by evidence and validation read models. It records whether a mapped value is direct, normalized, derived, unavailable, or ambiguous without creating a second canonical model.
