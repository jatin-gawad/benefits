# Canonical Benefit Administration

## Purpose

Benefit Administration stores who insures, administers, manages, or delivers a benefit. It is intentionally separate from the existing `PlanDesign` provision rows that store what a participant receives, such as maximums, deductibles, coinsurance, eligibility, and reimbursement rules.

The initial document profile is GroupHEALTH, but the database, repository, service, and Client Workspace UI are carrier-neutral.

## Existing schema audit

The desktop SQLite schema already contains:

- `Clients`: reused as the confirmed client master.
- `PlanDesign`: retained unchanged for benefit provisions.
- `Benefit`: reused as the existing benefit/source-tracking master. Canonical administration benefit codes are inserted only when absent.
- `Policy`: retained unchanged; it is not suitable for one-benefit-to-many-provider assignments.
- `Documents`: reused for source-document lineage and client ownership validation.
- Source-tracking fields on the existing `Benefit`, `Policy`, `Documents`, and related import tables.

The audited desktop schema did not contain an equivalent canonical `Organization`, provider-assignment, or benefit-administration table. It also did not contain a canonical `PlanDesignVersion` table. Therefore `PlanDesignVersionId` is nullable and indexed, but no invalid foreign key was invented. It is ready to reference a future canonical version table without changing the administration model.

No existing table was renamed, deleted, or repurposed.

## Additive migration

Migration: `BenefitAdministrationDatabaseMigration`

The migration runs from the existing `DatabaseService.InitializeCoreAsync()` path and is idempotent. It creates only:

1. `Organization`
2. `ClientBenefitAdministration`
3. `BenefitServiceProviderAssignment`

`Organization` is used because organizations such as WorldCare, TELUS Health Virtual Care, Trident Global Assistance, and Disability Management Institute are not necessarily insurance carriers. Creating a carrier-only provider table would misclassify them.

`ClientBenefitAdministration` references the existing `Clients`, `Benefit`, and `Documents` tables. It stores insurer identity, product and policy information, effective dating, review state, source lineage, and optional-benefit state. It does not contain benefit provisions.

`BenefitServiceProviderAssignment` supports zero, one, or many providers per benefit and keeps provider role, effective dates, review state, and source lineage.

Partial unique indexes prevent duplicate active administration identities and duplicate active provider-role assignments while preserving historical inactive rows. SQLite null behavior is normalized in the index expressions with `COALESCE`.

## Canonical organizations and aliases

The migration safely inserts the following organization masters only when their normalized name is absent:

- GroupHEALTH
- SSQ, Life Insurance Company Inc.
- Industrial Alliance
- Disability Management Institute
- Trident Global Assistance
- WorldCare
- Advica Health
- Teladoc Health
- Dialogue
- Sparrow
- TELUS Health Virtual Care
- The Co-operators

Alias resolution is owned by `BenefitAdministrationService`, not the UI or repository. Current aliases include SSQ variants, Industrial Alliance / iA Financial Group variants, DMI variants, GroupHEALTH variants, and Medical Second Opinion provider variants for WorldCare, Advica Health, Teladoc Health, Dialogue, and Sparrow. Provider roles are assigned from the benefit context as well as the organization identity so a multi-service organization such as Dialogue or Teladoc is not permanently classified as only a second-opinion provider. Source wording is retained in `InsurerSourceDisplayName`, provider `SourceDisplayName`, and the source-lineage JSON.

## Write boundary

`BenefitAdministrationService` owns benefit normalization, organization resolution, validation, duplicate detection, effective dating, source ownership, review state, and transaction coordination. `BenefitAdministrationRepository` contains parameterized SQL only.

Booklet extraction remains in the existing Python booklet pipeline. The GroupHEALTH-specific function extracts the insurer/provider table into neutral candidates. The existing `BookletExtractorRunner` transports those candidates, and the existing `DatabaseService.SavePlanDesignFromBookletExtractionAsync` saves provisions and administration in one transaction. A database failure rolls back the complete save.

Batch processing now passes the already discovered client identity into the existing save method. It does not identify a client from a policy number.

## Review and history

Extracted and carried-forward rows default to `NEEDS_REVIEW`. Corrections also return a row to review. Changing an active insurer or policy end-dates the old administration and creates a new active row. Provider assignments can be edited, end-dated, or deactivated without deleting history.

`CarryForwardAsync` copies active assignments to a target plan-design version, records `CopiedFromClientBenefitAdministrationId`, retains source lineage, and requires review on the copied records.

## Adding support for another document profile

Add carrier-specific table recognition only in the document parser. Emit `BenefitAdministrationExtractionCandidate` records using the existing neutral fields. Do not add carrier checks to the repository, service, ViewModel, or XAML. New organization aliases belong in `BenefitAdministrationService`; new benefit wording belongs in the existing `BenefitMatcher` normalization boundary.

## Production seed boundary

Only organization masters and canonical benefit-type masters are seeded. Client-specific assignments are created through the application workflow and are never globally seeded into production.
