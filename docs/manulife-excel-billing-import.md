# Manulife Excel And CSV Billing Import

## What is reused

The import uses the existing desktop SQLite framework and keeps the current billing paths intact.

- `Clients` remains the selected client record.
- `Carrier` supplies the Manulife carrier identity.
- `Documents`, `ImportBatch`, `DocumentSource`, `ImportLog`, and `ValidationIssue` retain file provenance and audit history.
- `manulife_billing_statement_versions` is the Carrier Billing Import record. It was extended with source-row counts, employee and benefit counts, reported totals, reconciliation difference, parser version, and warning/error counts.
- Existing Manulife PDF and GroupHEALTH PDF parsers are unchanged.

## New detail tables

- `manulife_billing_employee_records` stores one source row per employee/certificate while preserving certificate, class, plan, division, province raw values, reported amounts, calculated amounts, and source row hash.
- `manulife_billing_benefit_records` stores one row per employee and detected benefit code. Unknown codes are preserved rather than discarded.
- `manulife_billing_exceptions` stores row/benefit validation exceptions and their resolution status.

The database save is one transaction. A persistence failure rolls back the version, employee rows, benefit rows, exceptions, document source, and import audit rows together.

## Header aliases

Approved aliases are in `config/benefits/manulife-billing-header-aliases.json`. Add a new source-header variation to the matching canonical field list. The parser normalizes leading/trailing whitespace, line breaks, repeated spaces, apostrophes, punctuation, and casing before matching.

For example, add an alternate certificate header under `certificate_number`:

```json
"certificate_number": ["Certificate Num", "Certificate Number", "Certificate #", "Member Certificate"]
```

An ambiguous required mapping blocks confirmation. Unmapped headers are retained in the preview for review.

## New benefit codes

No code change is required for a new benefit code that follows the standard pattern:

```text
ABC: Plan Code
ABC: Current Month Premium (1)
ABC: Benefit Total (4)
```

The parser groups every `[Benefit Code]: [Field Name]` section dynamically. An unknown code is saved with the original code and shown as a preview warning. Add a display name to `BenefitName` in `ManulifeExcelBillingImportService.cs` only when a friendly label is desired; data ingestion does not depend on this list.

## Desktop workflow

1. Open the client Billing workspace.
2. Select the client.
3. Choose `Select Billing Files` and select `.xlsx`, `.xlsm`, or `.csv`.
4. Review the confirmation summary with period, counts, totals, warnings, errors, and reconciliation difference.
5. Confirm the save.

The batch screen detects billing workbooks and creates a preview-only workflow. It does not save Excel/CSV billing data because the batch screen does not yet have an explicit client-selection control. This prevents a folder named `Billings` from being mistaken for a client.

## Controls and limits

- Accepted input types: `.xlsx`, `.xlsm`, `.csv`; maximum 50 MB.
- Financial calculations use `decimal`; blanks remain null when the distinction matters.
- Certificate, class, plan, and division identifiers remain text.
- Formula-like CSV values and Excel formulas are blocked before confirmation to mitigate CSV/formula injection.
- Ordinary log messages exclude employee names and certificate numbers. The detailed encrypted/local database records retain the source data needed for billing administration.
- The current desktop application has no central authenticated user or malware-scanning provider. Client selection scopes the import locally; enterprise RBAC, encryption-at-rest policy, retention deletion, and malware scanning should be connected to the organization-wide desktop deployment controls before multi-user rollout.
