# Canonical Claims Phase 1 Schema Record and Delta

## Baseline recorded

The live desktop database was inspected read-only at `%LOCALAPPDATA%\Westland\RenewalReviewBuilder\data\Benefits_Intelligence.db` before launching the modified application. The inspection confirmed the existing production schema without copying runtime data into the repository.

Existing claims and lineage tables confirmed:

- `ClaimsExperience`
- `MonthlyClaims`
- `claims_categories`
- `claims_distribution`
- `claims_distribution_reports`
- `claims_distribution_rows`
- `benefit_experience_history`
- `DocumentSource`
- `ImportBatch`
- `ImportLog`
- `ValidationIssue`
- `Policy`
- `Benefit`
- `Premium`
- `Pooling`

The baseline showed source metadata columns already present on the primary legacy claims tables. No canonical table existed in this snapshot.

## Additive schema delta

| Table | Classification and justification |
|---|---|
| `CanonicalClaimsReport` | Aggregate database entity; owns client/carrier/policy/source identity and the idempotency key. |
| `CanonicalClaimsPeriod` | Period database entity; keeps normalized period boundaries explicit and queryable. |
| `CanonicalClaimsFinancialFact` | Stable financial fact entity; stores premium, claims, loss-ratio input, and lives without carrier columns. |
| `CanonicalClaimsMonthlyFact` | Stable time-series fact entity; monthly facts require a distinct month key and values. |
| `CanonicalClaimsCategoryFact` | Stable category entity; typed category amount/count/percentage facts. |
| `CanonicalClaimsDistributionFact` | Stable distribution entity; adds dimension and average fields genuinely distinct from a category total. |
| `CanonicalClaimsBenchmarkFact` | Stable benchmark entity; retains the benchmark definition with its typed value. |
| `CanonicalClaimsEvidence` | Lineage entity; links every persisted fact to existing documents/import batches and parser/mapping evidence. |
| `CanonicalClaimsValidation` | Canonical validation history; complements issues also written to existing `ValidationIssue`. |
| `CanonicalClaimsParityResult` | Migration read/audit entity; stores legacy-versus-canonical comparisons and review state. |
| `CanonicalClaimsUtilizationFact` | Additive adjudication fact; stores report-period + benefit + relationship + category + service amounts and source state. |
| `CanonicalClaimsFactRevision` | Additive audit entity; stores the full natural key, both source references, changed values, and review status. |

No specialized drug, indication, provider, large-claim, disability, or life table is created in Phase 1. No legacy table is deleted, renamed, or deprecated.

## Phase 2 Canada Life additive delta

No Canada Life table and no carrier-named column is added. Existing canonical tables are sufficient for Canada Life facts.

The following carrier-neutral columns are added idempotently to existing Phase 1 databases:

- `CanonicalClaimsEvidence`: mapping classification, normalization rule, derivation formula, normalized flag, warning code, and warning message. These are required to distinguish direct, normalized, derived, unavailable, and ambiguous values and retain field-level lineage.
- `CanonicalClaimsValidation`: mapping classification, source reference, and normalization rule. These allow Canada Life warnings to reuse the shared validation history instead of creating a carrier validation table.
- Category, distribution, and benchmark facts: coverage, division, class, measurement type, dimension type/value, currency, and unit of measure where those shared fact tables did not already retain them. These fields are carrier-neutral and are necessary to avoid losing values already present on `CanonicalClaimFact` during persistence.

All changes use `CREATE TABLE IF NOT EXISTS`, column-existence checks, and additive `ALTER TABLE` statements. Every legacy claims table remains unchanged.

## Sun Life content-period extension

Sun Life claims-service PDFs keep their explicit start and end dates as canonical period identity. `Current`, `Prior`, `Prior Prior`, and `Prior Prior Prior` are calculated presentation roles for the four newest unique periods and are not persisted as identity. The optional 6/9/12 selection remains a comparison-window override for compatible exports.

Relationship-level service facts are additive. Relation `All`, category subtotals, relationship totals, and grand totals are retained only as non-additive controls. The canonical natural key combines client, carrier, policy, benefit, period dates, normalized relationship, category, and service. Identical values collapse idempotently; changed values select one deterministic current fact and create a `NeedsReview` revision tied to both source references.

Legacy category worksheets and existing PowerPoint tables remain available. Their grouped labels are produced at the compatibility export boundary, while the new canonical export tables retain source-category detail.

## Idempotency and transaction rules

The unique report key combines client, carrier, policy, reporting period, source document, and mapping version. A retry reuses the report, deletes only its canonical children, and recreates those children inside one transaction. Any database error rolls the transaction back, preserving the previous complete canonical report.
