# Flexible Benefits Extraction

## Purpose

The desktop application can import HCSA, LSA, wellness, and other flexible-spending plan documents into the existing Benefits Intelligence SQLite database. The signed PDF remains the legal source. Normalized records are operational data and every extracted field retains its document, version, page, bounding box, extraction method, confidence, and review status.

The first carrier implementation supports National HealthClaim LifeHealth+ enrollment and administrative-services indemnity documents. The design separates carrier detection, local PDF/OCR extraction, carrier parsing, normalization, validation, persistence, and presentation so another carrier can be added without copying database or UI logic.

## Reused Architecture

- `ConnectorRegistry` selects the carrier and document-specific connector.
- `IBenefitsConnector` and `NormalizedImportPayload` remain the common ingestion boundary.
- `DatabaseService.InitializeCoreAsync()` remains the additive SQLite initialization hook.
- `Documents`, `DocumentSource`, `ValidationIssue`, `Clients`, and existing client-workspace routing remain the shared source, review, client, and navigation foundations.
- `PlanDesignWorkspaceDataService`, `PlanDesignWorkspaceView`, and the existing Client Overview host the new presentation surfaces.

## New Objects And Why They Exist

- `IFlexibleBenefitDocumentTextExtractor` isolates machine text, local OCR, page layout, selection marks, and signature-mark evidence from carrier business rules.
- `WindowsOcrPdfExtractor` supplies local, in-process PDF and Windows OCR extraction. It does not send documents to an external service.
- `IFlexibleBenefitDocumentParser` defines the reusable carrier parser contract.
- `NationalHealthClaimFlexibleBenefitParser` maps National HealthClaim wording and layout to normalized records without hard-coding the client or expected values.
- `FlexibleBenefitNormalizationService` owns dates, amounts, percentages, ratios, ages, durations, labels, and controlled value normalization.
- `FlexibleBenefitValidationService` owns cross-field and completion-state validation.
- `FlexibleBenefitImportService` orchestrates classification, connector resolution, validation, and one-transaction persistence.
- `FlexibleBenefitRepository` provides the Plan Design and review features with one database-facing boundary.
- `FlexibleBenefitCategoryCatalogue` maps carrier wording to stable category codes from `config/flexible_benefits/lsa_categories.json`.

These objects are separate because their responsibilities change independently: carrier form wording can change without changing schema, and a future approved OCR engine can change without changing carrier normalization.

## Normalized Schema

The additive schema is initialized idempotently and uses foreign keys and indexes:

- `FlexibleBenefitDocumentVersion`: immutable import version, hash, effective period, current/replaced state, and raw extraction audit JSON.
- `FlexibleBenefitPlan`: plan-level dates, account enablement, allocation, year, runoff, carry-forward, ages, funding, fees, status, and version links.
- `FlexibleBenefitClass`: class identity, tier limits, HCSA/LSA amounts and percentages, waiting period, proration, effective dates, and source quality.
- `FlexibleBenefitCategory`: controlled category code plus original document wording, group, account type, checkbox state, and source quality.
- `FlexibleBenefitEligibilityRule`: extensible rules that do not fit fixed plan or class columns.
- `FlexibleBenefitFundingConfiguration`: PAD, pre-funding, pay-as-you-go, funds-clear, and payment-method terms.
- `FlexibleBenefitFee`: setup, administration, claim, employee, annual, amendment, and other fees with fixed/percentage values, tax, and notice terms.
- `FlexibleBenefitAgreement`: agreement dates, legal parties, signers, signature evidence, governing province, and status.
- `FlexibleBenefitContact`: normalized plan roles and extracted contact details. Existing client contacts are matched before a duplicate is created.
- `FlexibleBenefitPlanSourceField`: mandatory field-level lineage for every operational value.
- `FlexibleBenefitCorrectionAudit`: immutable before/after record for permitted manual corrections.
- `FlexibleBenefitExtractionStaging`: raw extraction staging and completion state retained for audit.

Money is stored with SQLite `NUMERIC` affinity and is read/written as `decimal`. Missing values remain `NULL`; they are not converted to zero. The plan, class, category, and version tables preserve effective dates and current/historical records.

## Extraction Flow

1. Classify the PDF and administrator.
2. Confirm the target client selected in the workspace.
3. Calculate the source hash and apply duplicate/version rules.
4. Read machine text first and render pages for local OCR when layout extraction is required.
5. Detect text lines, checkboxes, selection state, and signature marks with page coordinates.
6. Map carrier wording into controlled enums and category codes.
7. Validate individual fields and cross-field consistency.
8. Keep raw extraction in staging for audit.
9. Save the normalized graph, lineage, validation issues, and document version in one transaction.
10. Refresh Client Plan Design and the concise Client Overview only after commit.

Checkboxes are tri-state: `true` means visibly selected, `false` means visibly unselected, and `NULL` means the source could not be determined. Blank employee enrollment sections are recognized but do not create employee records.

## Validation And Completion

Validation covers client and administrator confirmation, document classification, critical dates, meaningful plan data, non-negative limits and durations, percentage bounds, fixed allocation totals, dependent age ordering, fee units, checkbox reconciliation, typed-name versus signature evidence, and required source lineage.

- `Failed`: a critical requirement or transaction fails. No normalized graph is committed.
- `Needs Review`: data is preservable, but one or more warnings or ambiguous values require a person.
- `Completed`: all required checks pass and the transaction commits.

For `Signed_LSA_HCSA_NHC.pdf`, the plan is intentionally `Needs Review` because the postal code is missing/invalid and the plan-administrator email field contains `Hive HR`, which is not an email address. Those values are not silently guessed.

## Duplicate And Version Rules

- The default import mode is `Cancel`.
- Same hash, client, and document type returns `Duplicate` and creates no additional rows.
- Same hash for a different client is stored separately and marked `Needs Review`.
- Same effective period with a different file returns `Existing Version` unless the user explicitly chooses a version action.
- `ImportAsNewVersion` preserves both graphs, makes the new plan current, and links it to the prior plan.
- `ReplaceExistingVersion` also preserves prior rows, marks the prior plan and document version `Replaced`, and makes the replacement current.
- Historical rows are never destructively overwritten.

## Field Lineage And Review

Every stored field can be opened from Plan Design to see its raw and normalized value, legal source document, source page, bounding box, extraction method, confidence, correction status, and review status. `View Source Page` opens the legal PDF at the recorded page. A correction is restricted to supported normalized fields, records analyst/reason/timestamp, and writes an immutable correction audit row. `Mark Reviewed` updates the plan and its field review states without deleting warnings or source evidence.

## User Interface

Client Workspace > Plan Design includes a `Flexible Benefits` tab with:

- plan, class limits, eligible categories, funding, fees, agreement, and contact sections;
- `View Original Document`, `View Source Page`, `View Extraction Details`, `Edit/Confirm`, `Mark Reviewed`, and `Compare to Prior Version` actions;
- review warnings, field lineage, and preserved version history.

Client Overview only shows concise HCSA/LSA amount summaries, administrator, effective date, and review state. It does not expose fees, category details, signatures, or employee/dependent personal information.

## Privacy And Security

- OCR runs locally in the desktop process.
- The implementation does not send document content to an external service.
- Extracted names, emails, signatures, and dependent details are not written to diagnostic logs.
- Existing client-workspace access controls remain the authorization boundary.
- Blank employee sections do not create census records. Future completed enrollment data must use the existing census/enrollment model.

## Adding Another Carrier

1. Add carrier aliases to the classifier only when shared identification cannot already recognize them.
2. Implement `IFlexibleBenefitDocumentParser` for the carrier wording and layout.
3. Reuse `FlexibleBenefitNormalizationService`, the controlled catalogue, validation service, repository, schema, review workflow, and UI.
4. Add an `IFlexibleBenefitsConnector` and register it once in `ConnectorRegistry`.
5. Add parser fixtures for the carrier document and assertions for checkbox states, normalized values, lineage, validation, and duplicate/version behavior.
6. Do not add carrier-specific tables or client-specific values.

## Verification

The automated coverage includes normalization, tri-state checkbox detection, fuzzy-label regression, signature evidence, category mapping, client-facing amount summaries, real National HealthClaim PDF parsing, transactional rollback, foreign keys, duplicate rejection, version preservation, replacement history, source lineage, correction audit, review persistence, and connector resolution.
