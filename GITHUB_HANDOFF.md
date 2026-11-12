# GitHub Handoff

This repository is the Windows WPF desktop Renewal Review Builder source. It includes the .NET application and tests, Python proposal helpers, carrier/benefit configuration, application-owned workbook template, documentation, and build/packaging scripts.

It deliberately excludes compiled output, local SDK/cache folders, runtime SQLite data, backups, client documents, generated reports, and test-result files.

## Before the first push

1. Create an empty **private** repository in the approved Westland GitHub organization.
2. Confirm the repository's visibility, owners, branch protection, and security review requirements.
3. Review `git status` before committing. The supplied `.gitignore` is defensive, but it is not a substitute for reviewing the staged file list.
4. Do not add a public/open-source license unless Westland has approved one.

```powershell
git init
git add .
git status
git commit -m "Initial import of Renewal Review Builder desktop source"
git branch -M main
git remote add origin https://github.com/ORGANIZATION/REPOSITORY.git
git push -u origin main
```

## Engineering entry points

- `README.md` - build and product overview
- `ARCHITECTURE.md` - system boundaries, data flow, integrity rules, and known debt
- `CONTRIBUTING.md` - local commands and pull-request standard
- `SECURITY.md` - sensitive-data and vulnerability-handling rules
- `docs/testing/TEST_STRATEGY.md` - CI and controlled real-document fixture boundary
- `.github/workflows/ci.yml` - Windows .NET/Python validation

## Current verified baseline

See `VALIDATION_BASELINE.md`. The source builds in Release with zero warnings and zero errors; the complete .NET run has no failures, and the complete Python suite passes.

## Ownership note

The application is a mature modular monolith with substantial domain coverage. The oversized legacy main window is documented debt and should be reduced incrementally behind tested application-service boundaries. It should not be presented as a completed MVVM rewrite.
