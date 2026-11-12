# Test Strategy

## Test layers

1. Deterministic unit tests cover calculations, normalization, benefit identity, source selection, and fail-closed decisions.
2. Persistence tests use disposable SQLite databases and verify idempotency and canonical identity.
3. Architecture tests enforce forbidden dependency rules, especially UI-to-SQL and UI-to-extractor shortcuts.
4. Export tests inspect formulas, Open XML, chart caches, replacement tokens, and source evidence.
5. Real-document regression tests exercise approved carrier fixtures when those fixtures exist on the test workstation.
6. Python tests validate focused proposal mapping, booklet comparison, and shared-database persistence.

## Fixture boundary

Client and carrier documents are intentionally excluded from Git. Tests that depend on locally approved source documents guard their execution with file-existence checks. CI therefore proves the synthetic and repository-contained test surface; a release workstation with the controlled fixture library proves the additional real-document surface.

This distinction must be reported in pull requests and release evidence. A passing CI run must not be described as proof that unavailable real-document fixtures executed.

## Required commands

```powershell
dotnet build .\RenewalReviewBuilder.sln -c Release --no-restore -warnaserror
dotnet test .\tests\RenewalReviewBuilder.Tests\RenewalReviewBuilder.Tests.csproj -c Release --no-build --no-restore
py -3.14 -m unittest discover -s tests\python -p "test_*.py" -v
```

## Release evidence

For source handoff, record SDK version, warning/error count, passed/failed/skipped counts, Python results, and credential scan. For a portable release, also verify source-to-package DLL hashes, ZIP/folder parity, startup smoke, and zero remaining packaged-app processes.
