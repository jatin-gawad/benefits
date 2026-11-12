# Start here

This repository contains the Windows WPF desktop application and its Python proposal engine.

## Build and test

Use Windows with .NET SDK 10.0.300 or a compatible .NET 10 feature update, plus Python 3.14.

```powershell
dotnet restore RenewalReviewBuilder.sln
dotnet build RenewalReviewBuilder.sln -c Release --no-restore
dotnet test tests/RenewalReviewBuilder.Tests/RenewalReviewBuilder.Tests.csproj -c Release --no-build --no-restore
py -3.14 -m pip install -r requirements-test.txt
py -3.14 -m unittest discover -s tests/python -p "test_*.py" -v
```

See README.md for portable packaging, CONTRIBUTING.md for development guidance, and VALIDATION_BASELINE.md for validation scope.

## Configuration

- Select authorized input documents through Browse.
- WESTLAND_ONEDRIVE_ROOT can point to an authorized local source root.
- WESTLAND_SHAREPOINT_DOCUMENTS_URL supplies the SharePoint Documents button destination. The app explains when this setting is missing.
- RENEWAL_TEST_FIXTURES optionally supplies synthetic fixtures named sample-renewal.pdf and sample-claims.xlsm.
- The workbook template contains six empty report sheets. Exporters populate and format them from supplied data.
- National HealthClaim signer names require an unambiguous Subscriber signer name: or Administrator signer name: label. Missing or ambiguous names require review.

Follow SECURITY.md when handling client documents and application data.
