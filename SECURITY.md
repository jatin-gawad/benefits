# Security and Data Handling

## Reporting

This is an internal Westland application. Report suspected vulnerabilities, exposed credentials, or client-data leakage through the approved internal security channel and notify the repository owner. Do not open a public issue containing sensitive details.

## Repository rules

- Never commit client renewal documents, claims files, employee/member data, generated reports, or runtime databases.
- Never commit passwords, tokens, connection strings containing credentials, or local environment files.
- Use synthetic or formally approved de-identified fixtures in tests.
- Treat extracted document text and file metadata as untrusted input.
- Keep SQL parameterized and keep persistence behind application/data services.
- Preserve fail-closed client, policy, and carrier identity checks.
- Review dependency and workflow updates before merging.

## Local data

Runtime data is stored below `%LOCALAPPDATA%\Westland\RenewalReviewBuilder`. It is outside the repository and is not included in source packages. Follow Westland retention and access-control requirements when copying or backing up that directory.

## Supported version

Security fixes are applied to the current main branch and the latest internally distributed portable build. Older portable copies should be replaced rather than patched in place.
