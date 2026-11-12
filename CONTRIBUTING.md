# Contributing

## Local setup

Use Windows with the .NET 10 SDK. From the repository root:

```powershell
dotnet restore .\RenewalReviewBuilder.sln
dotnet build .\RenewalReviewBuilder.sln -c Release --no-restore -warnaserror
dotnet test .\tests\RenewalReviewBuilder.Tests\RenewalReviewBuilder.Tests.csproj -c Release --no-build --no-restore
py -3.14 -m pip install -r requirements-test.txt
py -3.14 -m unittest discover -s tests\python -p "test_*.py" -v
```

If the bundled SDK is present, `.\.dotnet-sdk\dotnet.exe` can be used instead of `dotnet`. Keep NuGet and build output outside OneDrive when local policy blocks intermediate file writes.

## Change standard

- Keep UI code free of SQL and parser construction; add or extend an application service.
- Preserve source lineage and distinguish missing values from explicit zero.
- Add a focused regression test for every defect.
- Do not commit client documents, databases, generated reports, credentials, or portable binaries.
- Do not weaken identity checks or silently map a source to a different client.
- Prefer readable, incremental changes over framework-heavy rewrites.
- Update architecture or contract documentation when a boundary changes.

## Pull-request evidence

Include:

1. The user-visible or data-contract change.
2. The files and architectural boundary affected.
3. Focused test evidence.
4. Full-suite status and any environment-dependent tests that did not execute.
5. Migration, rollback, and data-compatibility notes when persistence changes.

## Definition of done

A change is complete when Release builds with zero warnings, relevant .NET and Python tests pass, generated Office artifacts are inspected when their layout changed, and any remaining limitation is documented explicitly.
