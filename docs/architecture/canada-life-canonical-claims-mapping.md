# Canada Life Canonical Claims Mapping Specification

## Boundary

`CanadaLifePdfParser` extracts Canada Life documents into `CarrierPdfParseResult`. `CanadaLifeClaimsCanonicalMapper` interprets that result and produces the existing `CanonicalClaimsImportEnvelope`. The mapper does not open source files, validate shared business rules, or write SQLite.

## Mapping rules

| Canada Life source concept | Canonical target | Classification | Evidence and validation |
|---|---|---|---|
| Carrier name `Canada Life` or legacy `Great-West Life` identity | Report carrier `Canada Life` | Normalized | Parsed carrier identity; a non-Canada-Life result is an error. |
| Policy number | Report policy number | Direct | Existing parsed policy; missing policy is an error. |
| Experience period label | Report start/end and fact reporting-period dimension | Normalized | Original label and page retained; invalid label is an error. The report spans the earliest explicit start through latest explicit end in the source. |
| `Extended Health Care`, `Health`, `Medical`, `EHC` | Benefit `EHC` | Direct or Normalized | Explicit benefit-alias rule and page retained; unknown benefit requires review. |
| `Dental`, `Dental Care`, `DTL` | Benefit `Dental` | Direct or Normalized | Explicit benefit-alias rule and page retained. |
| Life, disability, AD&D, dependent life, critical illness, EAP aliases | Canonical benefit identifier | Direct or Normalized | Explicit benefit-alias rule and page retained. |
| Billed premium, paid claims, incurred claims, incurred loss ratio, average lives | Financial fact | Direct | Raw values and page retained. Missing values are unavailable, not zero; negative amounts require review. |
| Paid loss ratio | Benchmark percentage | Derived | Uses unchanged shared `UnderwritingEngine.CalculateLossRatioPercent`; formula and component values retained. |
| Credibility, target loss ratio, annual trend, fee-guide adjustment | Benchmark percentage | Direct | Parsed value, source label, page, parser and mapping version retained. |
| Claims category amount/percent | Category fact | Direct | Raw benefit/category labels and page retained. |
| Canada Life service-category aliases | Canonical category | Direct or Normalized | Explicit category-alias rule retained; unknown category requires review. |
| Claims distribution amount, percentage, occurrence count, average claim | Distribution fact | Direct | Source page, raw values, benefit/category, division and class retained. Negative counts are invalid; negative amounts require adjustment/reversal review. |
| Carrier norm percentage | Benchmark fact | Direct | Source page and Canada Life benchmark definition retained. |
| `ALL` division or class | `All` | Normalized | Dimension rule retained. Unknown or missing dimension requires review. |
| Explicit alphanumeric division/class | Same identifier | Direct | No value is inferred from position, filename, or another document. |
| Pooling threshold and pooling charge | Benchmark facts | Direct | Existing parser exposes the value but not its page; document evidence is retained and precision warning requires review. |
| Demographic current male/female lives | Category facts with `CoveredLives` measurement | Direct | Source page, benefit, age band, sex, and `Lives` unit retained. Lives are not relabelled as employees or certificates. |
| ASO, insured, refund/retention accounting, pooled/non-pooled funding labels | Canonical funding term | Normalized | Alias rules exist for parser outputs that expose these concepts; current Canada Life parsed model does not expose a funding field. |
| Province names and abbreviations | Canadian province code | Direct or Normalized | Alias rules exist; current Canada Life parsed claims model does not expose province. |
| Employees, certificates, dependants, refund values, expenses, tax, commission, reserves, pooled/non-pooled claim totals | No value | Unavailable | Structured information result; never defaulted to zero or inferred. |
| Unknown terminology or identifiers | Blank canonical identifier with raw evidence | Ambiguous | Structured warning and Needs Review; no closest-looking assignment. |

## Reconciliation and completion

The mapper reuses `ClaimsDistributionValidation`, including the existing $1 claims-distribution reconciliation tolerance. Shared validation then checks identity, period, evidence, duplicate fact IDs, membership counts, and required facts. Parity compares the parsed Canada Life result with canonical financial and distribution facts using the existing canonical tolerance.

- `Failed`: an error exists; canonical persistence is not called.
- `Needs Review`: facts can be retained but warnings exist, including missing totals, failed/unavailable reconciliation, unmapped terms, ambiguous dimensions, or incomplete coordinates.
- `Completed`: validation and required reconciliation pass and the canonical transaction commits.

## Idempotency

The shared canonical report key includes client, normalized carrier, policy, normalized period, source document, and mapping version. The same document retry refreshes only its canonical children in one transaction. A different source document for the same period creates a separately traceable canonical report. Legacy rows are never overwritten by the canonical transaction.
