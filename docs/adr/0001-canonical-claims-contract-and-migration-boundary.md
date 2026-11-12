# Canonical Claims Contract and Migration Boundary

## Context

Claims are stored in several working legacy shapes. Existing parsers, exports, dashboards, source tracking, and underwriting calculations must continue to work while a common claims language is introduced.

## Decision

Phase 1 adds `CanonicalClaimsImportEnvelope` beside the legacy models. Carrier-specific extraction stays in the current parser. `ICarrierClaimsCanonicalMapper` is the boundary that converts parser output into the canonical contract. `CanonicalClaimsImportService` is the only application-level canonical write boundary. UI and parser classes do not write canonical tables.

The proof of concept uses the existing Manulife-compatible `ClaimsExperienceDataModelResult`; it does not implement new Manulife Claims Insight extraction.

## Carrier-specific logic boundary

`CarrierClaimsCanonicalMapperRegistry` selects a mapper before canonical data exists. `ManulifeClaimsCanonicalMapper` may understand Manulife labels. Canonical persistence, queries, parity, underwriting, renewal, and UI code contain no Manulife conditional.

## Canonical ownership

Canonical Claims Experience owns historical claims facts. Document Operations continues to own files, versions, import batches, and primary evidence. Underwriting and Renewal remain consumers and their formulas are unchanged.

## Persistence choice

Ten additive stable tables store report ownership, period, five shared fact shapes, validation, evidence, and parity. Specialized drug, indication, provider, large-claim, disability, and life types are contracts only in Phase 1 and have no tables.

## Compatibility strategy

Option A was selected: the existing transaction saves legacy data first. For the enabled proof-of-concept mapper, the controlled canonical service then maps, validates, compares, and writes canonical data in its own transaction. A canonical failure is reported as needing review without undoing a successful legacy import.

## Source evidence

Canonical report and fact evidence reuse `Documents`, `DocumentSource`, `ImportBatch`, `ImportLog`, and `ValidationIssue`. Evidence records carry parser/mapping versions and the most precise available page, worksheet, section, row, or cell reference. The proof-of-concept result only exposes document-level precision.

## Parity requirements

Parity stores legacy/canonical values, variance, tolerance, classification, explanation, related IDs, and review status. Missing required canonical values and unexpected mismatches require review; they are never counted as complete parity.

## Alternatives considered

- Rewriting parsers around canonical objects was rejected because it risked working PDF/Excel behavior.
- Writing canonical tables directly from each parser was rejected because it creates several uncontrolled write paths.
- A read-only adapter was lower risk but rejected because Phase 1 requires durable evidence, retry testing, and parity history.
- One universal key-value table was rejected because it weakens types, constraints, and query clarity.

## Risks

- A second transaction means legacy data can succeed while canonical persistence needs review.
- Existing parser results sometimes lack page/cell evidence or policy numbers.
- SQLite `NUMERIC` values must continue to be compared with documented decimal tolerance.
- Only the Manulife-compatible workflow is enabled in Phase 1.

## Rollback

Remove the `CanonicalClaimsImportService` invocation from `SaveClaimsExperienceAsync`. All legacy tables and workflows remain intact. Additive canonical tables can remain unused; no destructive rollback is required.

## Future Phase 2

- Add carrier mappers one at a time with golden fixtures.
- Map category/distribution and specialized claims sections where the source supports them.
- Move downstream read consumers to `CanonicalClaimsQueryService` after parity approval.
- Improve evidence precision at parser boundaries.
- Add operational review UI only after canonical data is proven stable.

## Phase 2 amendment: Canada Life

Phase 2 implements the first additional carrier mapper, `CanadaLifeClaimsCanonicalMapper`, against the existing `CarrierPdfParseResult`. `CanadaLifePdfParser` remains the extraction owner. The mapper receives its completed parsed result and contains no PDF, Excel, SQL, UI, or transaction code.

The compatibility decision remains Option A. `SaveRenewalAnalysisAsync` commits all existing Canada Life legacy renewal, experience, and claims-distribution work first. It then invokes `CanonicalClaimsImportService`. Mapping, validation, or canonical transaction failure is logged and reported without removing the successful legacy result.

Canada Life routing uses the existing typed mapper registry. No new dependency-injection system or feature-flag framework was introduced because the desktop workflow does not currently have one. The typed registry prevents a Canada Life parsed model from being passed to the Manulife mapper and keeps carrier checks out of shared engines and UI code.

Phase 2 retains page evidence supplied by the existing parser and marks document-only evidence for review. Explicit aliases live in `CanadaLifeClaimsTerminology`; no broad closest-match or substring fallback is permitted. Missing values remain missing, zero remains `ReportedZero`, and unsupported concepts are recorded as unavailable rather than inferred.

Rollback remains additive: remove Canada Life from the mapper registry or remove the post-legacy canonical invocation. Canonical rows can remain unused and every Canada Life legacy table and workflow continues to function.

## Data ownership matrix

| Area | Owns |
|---|---|
| Document Operations | Files, classification, import batch, versioning, evidence |
| Claims Experience | Canonical historical claims, utilization, categories, benchmarks, claims-derived metrics |
| Underwriting Intelligence | Assumptions, credibility, trend, scenarios, projections, required rates |
| Renewal Intelligence | Carrier request, negotiation strategy, recommended position, final result |
| Plan Design Intelligence | Current provisions, alternative provisions, design scenario inputs |
