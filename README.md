# benefits-intelligence-platform

# Renewal Review Builder

Windows WPF desktop application and Python proposal engine. Start with [START_HERE.md](START_HERE.md). This repository contains source code; build the application before running it.

For engineering ownership, start with [ARCHITECTURE.md](ARCHITECTURE.md), [CONTRIBUTING.md](CONTRIBUTING.md), and [SECURITY.md](SECURITY.md).

Phase 1 Windows desktop app for building an employee benefits renewal review pack from:

- Manulife renewal PDF
- Manulife XLSM supplementary workbook
- Optional Westland PowerPoint template

The app produces:

- `Renewal_Review.xlsx`
- `Client_Renewal_Draft.pptx`
- an on-screen findings checklist

## For Team Members

No install is required.

1. Unzip `RenewalReviewBuilderPortable.zip`.
2. Open the `RenewalReviewBuilderPortable` folder.
3. Double-click `Start Renewal Review Builder.cmd`.
4. Select the renewal PDF, Manulife XLSM, PowerPoint template, and output folder.
5. Click `Validate`, review the findings, then click `Export Pack`.

The app stores settings and logs in:

```text
%LOCALAPPDATA%\Westland\RenewalReviewBuilder
```

## Build The Portable Package

Build this once on a computer that can restore NuGet packages:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

The script uses the local `.dotnet-sdk` folder when it exists, copies source files to a Windows local temp build folder outside OneDrive, restores packages, runs tests, and publishes the no-admin portable folder under:

```text
dist\RenewalReviewBuilderPortable
```

It also creates:

```text
dist\RenewalReviewBuilderPortable.zip
```

The temporary staging folder is intentionally outside this repository so OneDrive does not sync build-only `bin`, `obj`, test, and publish files. The build script rejects a `-BuildRoot` located inside the repository.

If Windows blocks the build while the project is inside OneDrive Documents, move the project to a normal local folder or allow `.dotnet-sdk\dotnet.exe` in Controlled Folder Access, then rerun the build script. End users do not need this build step.

If restore fails with `NU1301` or `No credentials are available in the security package`, the computer/network is blocking NuGet over TLS. Build the portable package from a machine or network connection that can access `https://api.nuget.org/v3/index.json`, then share only `RenewalReviewBuilderPortable.zip` with the team.

## Sample Files

No client documents or sample client paths are bundled. Select your own authorized source documents with Browse.

## Architecture Standards

The desktop app follows one official canonical data pipeline standard for new and migrated workflows:

```text
docs\architecture\canonical-data-pipeline-standard.md
docs\architecture\plan-design-canonical-profile.md
docs\architecture\current-vs-quote-plan-design.md
docs\architecture\focused-proposal-canonical-persistence.md
```
