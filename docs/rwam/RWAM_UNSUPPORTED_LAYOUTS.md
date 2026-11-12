# RWAM Unsupported Layouts and Required Decisions

## Supplied samples

Unsupported supplied rate-sheet layouts: **none**. All eight supplied PDFs match RWAM-RS-2026-01 and are parsed into the existing Rate Reconciliation contract.

No supplied file is an ASO reconciliation statement. A document containing gross/net premium, claims, management or administration fees, taxes, gain/loss, or ending ASO balance without the rate-sheet fingerprint is not routed into Rate Reconciliation.

## Fail-closed boundaries

The following inputs are deliberately unsupported or returned for review instead of guessed:

- Scanned/image-only sheets and OCR-derived variants. No such file was supplied, so OCR support is not certified.
- A layout missing the required paired current/renewal headings or the stable benefit-section anchors.
- A same-class companion sheet whose canonical rate set conflicts with an already merged class sheet.
- A file whose carrier cannot be established by the certified complete fingerprint and source provenance. The supplied pages do not print the literal issuer name `RWAM`.
- A rate line with an unreadable amount, unknown tier, unstated period, or unsupported rate basis.
- A Stop Loss line expressed as `2.5% of Net Paid EHC Claims` when the net-paid-claims base is absent.
- WorldCare, EAP, Critical Illness, or Virtual Care zero rows when the source does not establish a usable premium basis.

These cases produce warnings or `Needs Review`; missing/unreadable values are not silently converted to zero.

## Domain decisions still required for broader support

1. Approve an authoritative claims source and calculation contract before converting conditional Stop Loss percentages into premiums.
2. Supply and independently validate at least one scanned RWAM rate sheet before claiming OCR support.
3. Supply any materially different historical or future RWAM layout before extending the fingerprint/table rules.
4. Provide issuer-authenticated document evidence if classification must be made independently of the supplied-folder provenance when the page itself does not name RWAM.

These limitations do not prevent reconciliation of any supplied rate sheet; they define the safe boundary for future documents.
