# Renewal Review Builder Architecture

## Purpose

Renewal Review Builder is a Windows desktop decision-support application for employee-benefits renewals. It ingests carrier documents, preserves source evidence, normalizes renewal data, performs deterministic calculations, stores canonical history, and produces consultant-facing Excel and PowerPoint deliverables.

The system is deliberately a modular monolith. That is the appropriate deployment model for a no-admin WPF tool, but module boundaries are enforced so extraction, persistence, business rules, and presentation do not collapse into UI code.

## System context

```text
Carrier PDF/XLSX/XLSM + prior documents
                  |
                  v
       discovery and typed extractors
                  |
                  v
       canonical analysis models
          |       |        |
          v       v        v
     rules and  SQLite   source checks
     calculations history  and lineage
          |       |        |
          +-------+--------+
                  |
                  v
        WPF review workspaces
                  |
                  v
       Excel and PowerPoint exports
```

## Architectural boundaries

1. **Views and windows** own interaction state, navigation, and display. They may call application services but must not open SQLite connections, issue SQL, or instantiate document extractors.
2. **Application services** coordinate workflows, identity checks, canonical loading, persistence, and exports. They define the boundary used by the UI.
3. **Extractors and parsers** translate source documents into typed models. They must retain source file, page/sheet/row, original text where available, confidence, and review status.
4. **Domain calculations** are deterministic and do not read the UI or database. Premium, loss-ratio, trend, credibility, and rate-basis decisions belong here.
5. **Persistence services** own SQLite schema and queries. Writes are idempotent where the workflow can be repeated, and identity includes client and carrier where a source package contains multiple proposals.
6. **Export writers** consume analyzed models. They do not silently repair missing evidence. Unknown values remain blank or explicitly marked for review.
7. **Python proposal modules** are a bounded companion component. Their database contract is tested against the shared canonical schema; the WPF layer does not invoke ad-hoc Python SQL.

The architecture guard tests in `CanonicalPipelineArchitectureTests` enforce the most important dependency rules.

## Data integrity rules

- The selected client is authoritative. Source identity conflicts fail closed and require review.
- Missing, explicit zero, and not-applicable values are distinct states.
- Carrier-stated, net, gross, proposed, negotiated, and final rates remain separate.
- Calculated values identify their inputs and rate basis; extracted values retain provenance.
- Re-imports must not duplicate canonical records or merge different carriers by file hash alone.
- Excel formulas are verified as formulas and inputs in automated tests; Excel is responsible for recalculating cached values when a workbook opens.
- Presentation filtering may omit only explicit zero/zero premium rows. Rows with missing premiums remain visible.

## Runtime and deployment

- Platform: Windows desktop, WPF, .NET 10.
- Persistence: local SQLite under `%LOCALAPPDATA%\Westland\RenewalReviewBuilder`.
- Distribution: self-contained portable package produced by `scripts\build.ps1`.
- Local source builds use the repository SDK when present; GitHub Actions restores a clean .NET 10 SDK on a Windows runner.
- Client documents and runtime databases are deliberately excluded from source control.

## Verification strategy

- Unit tests cover parsers, calculations, normalization, identity, and persistence.
- Architecture tests reject forbidden UI-to-database/parser shortcuts.
- Real-document regression tests run when their approved local fixtures are present.
- Export tests inspect workbook formulas, Open XML content, chart caches, and source lineage.
- Python tests validate proposal mapping, booklet comparison, and shared SQLite persistence.
- Portable release verification is a separate delivery gate and includes package parity and startup smoke testing.

## Known debt and controlled roadmap

The largest remaining structural risk is the legacy `MainWindow.xaml` / `MainWindow.xaml.cs` shell. It is functional but too large for efficient ownership and review. New functionality must be built in views and services, not added to the shell. Refactoring should be incremental:

1. Move one cohesive workspace at a time behind an application-service interface.
2. Preserve behavior with characterization tests before moving XAML or event handlers.
3. Replace shell event handlers with commands/view models only where the extracted workspace benefits.
4. Keep database migrations and export contracts backward compatible.
5. Measure progress by reduced shell responsibilities, not by introducing additional frameworks.

The project is not represented as free of technical debt. Its senior-quality standard is that risks are explicit, boundaries are testable, and changes are evidence-backed.

## Decision records and detailed contracts

- `docs/architecture/canonical-data-pipeline-standard.md`
- `docs/architecture/canonical-source-lineage-value-guardrails-v1.md`
- `docs/architecture/focused-proposal-canonical-persistence.md`
- `docs/architecture/canonical-claims-phase1-schema.md`
- `docs/adr/0001-canonical-claims-contract-and-migration-boundary.md`
- `docs/testing/TEST_STRATEGY.md`
