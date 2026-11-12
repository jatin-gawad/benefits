# Class Comparison Workspace

## Scope

The Class Comparison workspace is part of the existing desktop WPF Client Workspace and Plan Design route. It reads saved client data from the shared Benefits Intelligence SQLite database. It does not use proposal-comparison mock data.

## Additive SQLite Migration

`DatabaseService.InitializeCoreAsync()` creates these tables idempotently:

- `class_comparison_preferences`: one saved comparison state per client, including selected classes, baseline, active tab, filters, collapsed sections, and recommended class.
- `class_comparison_notes`: analyst notes linked to the client and optional class.
- `class_comparison_audit`: field-level history for preference and note changes.

The migration preserves existing records and adds client foreign keys and client/date indexes. No existing PlanDesign, rate, premium, billing, or document rows are rewritten.

## Data Sources

- Benefits come from saved `PlanDesign` and normalized plan-design feature rows.
- Rates are shown only when a saved rate class exactly matches an employee class name or code. Coverage tiers such as Single and Family are not treated as employee classes.
- Costs are shown only from class-mapped billing, premium, or census records.
- Missing values display as `Not available`; the workspace does not estimate missing costs or coverage.

## Comparison and Scoring

Comparison values normalize money, percentages, durations, booleans, and missing values before applying explicit directional rules. Each benefit rule declares whether higher, lower, included, or neutral values are preferred. Neutral provisions are not scored.

The reusable value score uses configurable weights for plan richness, employee cost, employer cost, eligibility, coverage breadth, competitiveness, administration complexity, and employee impact. Categories without reliable data are excluded and the remaining weights are rebalanced. A recommendation is applied only when a saved choice exists or the score has a unique leader.

## Export and UI Limits

Excel export is implemented for the selected comparison. PDF and PowerPoint class-comparison actions remain disabled until matching templates and export pipelines are available. The WPF shell has a 720-pixel minimum width, so a 390-pixel mobile layout is outside the desktop application boundary; cards and tables use horizontal scrolling at narrow supported widths.
