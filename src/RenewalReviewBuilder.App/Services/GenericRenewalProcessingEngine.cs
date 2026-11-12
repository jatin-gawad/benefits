using System.Text.Json;
using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class GenericRenewalProcessingEngine
{
    private readonly PdfTextExtractor _pdfTextExtractor = new();
    private readonly ReviewPackBuilder _reviewPackBuilder = new();
    private readonly DatabaseService _databaseService;

    public GenericRenewalProcessingEngine(DatabaseService? databaseService = null)
    {
        _databaseService = databaseService ?? new DatabaseService();
    }

    public Task<RenewalProcessingResult> process_renewal_file(string file_path, string output_folder)
    {
        return ProcessRenewalFileAsync(file_path, output_folder);
    }

    public async Task<RenewalProcessingResult> ProcessRenewalFileAsync(string filePath, string outputFolder)
    {
        var result = new RenewalProcessingResult();
        try
        {
            if (!File.Exists(filePath))
            {
                result.Errors.Add($"Renewal file was not found: {filePath}");
                return result;
            }

            var pages = _pdfTextExtractor.Extract(filePath);
            var documentText = ParsingHelpers.Clean(string.Join(" ", pages.Select(page => page.Text)));
            var carrierDetection = GenericRenewalDocumentAnalyzer.DetectCarrier(documentText);
            var sections = GenericRenewalDocumentAnalyzer.DetectSections(pages);
            var planDesignFeatures = GenericRenewalDocumentAnalyzer.ExtractPlanDesignFeatures(pages);
            var snapshot = BuildDocumentSnapshot(filePath, pages);

            var renewalCase = new RenewalCase
            {
                PdfPath = filePath,
                OutputFolder = outputFolder,
                PowerPointTemplatePath = ResolveDefaultPowerPointTemplatePath()
            };

            var analysis = await _reviewPackBuilder.AnalyzeAsync(renewalCase);
            var finalOutputFolder = BuildOutputFolder(outputFolder, analysis);
            Directory.CreateDirectory(finalOutputFolder);
            analysis.RenewalCase.OutputFolder = finalOutputFolder;
            result.OutputFolder = finalOutputFolder;

            var validation = GenericRenewalDocumentAnalyzer.Validate(analysis, carrierDetection, sections);
            result.ValidationSummary.AddRange(validation);
            result.Warnings.AddRange(analysis.Pdf.Warnings);
            result.Warnings.AddRange(analysis.Workbook.Warnings);
            result.Warnings.AddRange(validation.Where(item => item.Status.Equals("Warning", StringComparison.OrdinalIgnoreCase)).Select(item => item.Message));
            result.Errors.AddRange(validation.Where(item => item.Status.Equals("Error", StringComparison.OrdinalIgnoreCase)).Select(item => item.Message));

            var auditEntries = GenericRenewalDocumentAnalyzer.BuildAuditLog(filePath, analysis, carrierDetection, sections, planDesignFeatures);

            var databaseStatus = await _databaseService.SaveRenewalAnalysisAsync(analysis);
            result.DatabaseStatus = databaseStatus;
            var artifactStatus = await _databaseService.SaveGenericRenewalArtifactsAsync(
                analysis,
                databaseStatus.ClientId,
                databaseStatus.DocumentId,
                auditEntries,
                planDesignFeatures);
            if (!artifactStatus.Succeeded)
            {
                result.Warnings.Add(artifactStatus.Message);
            }

            var export = await _reviewPackBuilder.ExportAsync(analysis);
            result.Export = export;
            result.WorkbookOutputPath = export.WorkbookPath;

            var extracted = new RenewalProcessingExtractedJson
            {
                CarrierDetection = carrierDetection,
                RenewalCase = analysis.RenewalCase,
                Pdf = analysis.Pdf,
                Workbook = analysis.Workbook
            };
            extracted.Sections.AddRange(sections);
            extracted.Findings.AddRange(analysis.Findings);
            extracted.SourceChecks.AddRange(analysis.SourceChecks);
            extracted.PlanDesignFeatures.AddRange(planDesignFeatures);
            extracted.NemMaximums.AddRange(analysis.NemMaximums);
            extracted.AuditLog.AddRange(auditEntries);

            result.ExtractedJson = JsonSerializer.Serialize(extracted, new JsonSerializerOptions
            {
                WriteIndented = true
            });

            result.ExtractionJsonPath = Path.Combine(finalOutputFolder, "renewal_extraction.json");
            await File.WriteAllTextAsync(result.ExtractionJsonPath, result.ExtractedJson);

            result.ExtractionAuditLogPath = Path.Combine(finalOutputFolder, "extraction_audit_log.xlsx");
            WriteAuditWorkbook(result.ExtractionAuditLogPath, auditEntries);

            result.ValidationSummaryPath = Path.Combine(finalOutputFolder, "validation_summary.xlsx");
            WriteValidationWorkbook(result.ValidationSummaryPath, validation);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            result.Errors.Add(ex.Message);
        }

        return result;
    }

    private static RenewalDocumentSnapshot BuildDocumentSnapshot(string filePath, IEnumerable<PdfPageText> pages)
    {
        var snapshot = new RenewalDocumentSnapshot
        {
            SourceFileName = Path.GetFileName(filePath),
            SourceFilePath = filePath
        };

        foreach (var page in pages)
        {
            var snapshotPage = new RenewalDocumentPage
            {
                PageNumber = page.PageNumber,
                Text = page.Text
            };
            snapshotPage.Lines.AddRange(page.Lines);
            snapshot.Pages.Add(snapshotPage);
        }

        return snapshot;
    }

    private static string ResolveDefaultPowerPointTemplatePath()
    {
        var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var candidates = new[]
        {
            Path.Combine(userProfile, "OneDrive - WESTLANDCORP", "RenewalAutomation", "templates", "renewal_template.pptx"),
            Path.Combine(userProfile, "OneDrive - WESTLANDCORP", "Documents", "Renewal Review Builder", "Uploads", "renewal_template.pptx")
        };

        return candidates.FirstOrDefault(File.Exists) ?? "";
    }

    private static string BuildOutputFolder(string baseOutputFolder, AnalysisResult analysis)
    {
        var clientName = FirstNonBlank(analysis.RenewalCase.ClientName, analysis.Pdf.ClientName, "Unknown Client");
        var date = analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate;
        var dateFolder = date?.ToString("yyyy-MM-dd") ?? DateTime.Today.ToString("yyyy-MM-dd");
        return Path.Combine(baseOutputFolder, SanitizePathPart(clientName), dateFolder);
    }

    private static void WriteAuditWorkbook(string path, IReadOnlyList<ExtractionAuditEntry> auditEntries)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("extraction_audit_log");
        var headers = new[]
        {
            "Source File",
            "Source Page",
            "Section Type",
            "Field Name",
            "Extracted Value",
            "Source Text",
            "Confidence Score",
            "Validation Status",
            "Validation Message"
        };

        WriteHeaders(sheet, headers);
        for (var index = 0; index < auditEntries.Count; index++)
        {
            var row = auditEntries[index];
            var rowNumber = index + 2;
            sheet.Cell(rowNumber, 1).Value = row.SourceFile;
            sheet.Cell(rowNumber, 2).Value = row.SourcePage;
            sheet.Cell(rowNumber, 3).Value = row.SectionType;
            sheet.Cell(rowNumber, 4).Value = row.FieldName;
            sheet.Cell(rowNumber, 5).Value = row.ExtractedValue;
            sheet.Cell(rowNumber, 6).Value = row.SourceText;
            sheet.Cell(rowNumber, 7).Value = row.ConfidenceScore;
            sheet.Cell(rowNumber, 8).Value = row.ValidationStatus;
            sheet.Cell(rowNumber, 9).Value = row.ValidationMessage;
        }

        sheet.Columns().AdjustToContents();
        workbook.SaveAs(path);
    }

    private static void WriteValidationWorkbook(string path, IReadOnlyList<RenewalValidationItem> validation)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("validation_summary");
        WriteHeaders(sheet, ["Rule", "Status", "Message"]);
        for (var index = 0; index < validation.Count; index++)
        {
            var row = validation[index];
            var rowNumber = index + 2;
            sheet.Cell(rowNumber, 1).Value = row.Rule;
            sheet.Cell(rowNumber, 2).Value = row.Status;
            sheet.Cell(rowNumber, 3).Value = row.Message;
        }

        sheet.Columns().AdjustToContents();
        workbook.SaveAs(path);
    }

    private static void WriteHeaders(IXLWorksheet sheet, IReadOnlyList<string> headers)
    {
        for (var index = 0; index < headers.Count; index++)
        {
            sheet.Cell(1, index + 1).Value = headers[index];
        }

        var range = sheet.Range(1, 1, 1, headers.Count);
        range.Style.Fill.BackgroundColor = XLColor.FromHtml("#0E7C7B");
        range.Style.Font.FontColor = XLColor.White;
        range.Style.Font.Bold = true;
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private static string SanitizePathPart(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(value.Select(ch => invalid.Contains(ch) ? '-' : ch).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? "Unknown Client" : cleaned;
    }
}

public static class GenericRenewalDocumentAnalyzer
{
    private static readonly (string Carrier, string[] Clues)[] CarrierClues =
    [
        ("Canada Life", ["Canada Life", "Great-West Life Assurance Co.", "Summary of Renewal Rates and Premium"]),
        ("Manulife", ["Manulife", "Manulife Financial"]),
        ("Sun Life", ["Sun Life"]),
        ("Equitable", ["Equitable", "The Equitable Life Insurance Company of Canada"]),
        ("RBC", ["RBC Insurance", "RBC"]),
        ("GreenShield", ["GreenShield", "Green Shield"]),
        ("GroupSource", ["Claims Payor ALL", "Ledger #", "The following rates will be reflected", "EXPERIENCE CALCULATION"]),
        ("GroupHEALTH", ["GroupHEALTH", "Group Health", "GroupPLUS", "clienthelp.grouphealth.ca", "Current Underwriter SSQ FINANCIAL"]),
        ("Allstate", ["Allstate", "Allstate Insurance Company of Canada"]),
        ("National HealthClaim", ["National HealthClaim", "National Health Claim", "LifeHealth+", "NHC"]),
        ("Desjardins", ["Desjardins"]),
        ("Beneva", ["Beneva"]),
        ("Empire Life", ["Empire Life"]),
        ("IA", ["Industrial Alliance", "iA Financial", "IA Financial"]),
        ("Victor", ["Victor Insurance Managers", "victorinsurance.ca"])
    ];

    private static readonly (string SectionType, string[] Clues)[] SectionClues =
    [
        ("Client / Policy Information", ["Policy Number", "Effective:", "Presented by:"]),
        ("Executive Summary", ["Executive Data Summary", "Executive Summary"]),
        ("Renewal Rates and Premium", ["Summary of Renewal Rates and Premium", "Monthly Premium Comparison", "Total monthly cost summary", "Summary of rates and estimated monthly premium", "Coverage and estimated monthly premium by employee", "The following rates will be reflected"]),
        ("Experience Analysis", ["Experience Analysis", "Experience History", "Paid Loss Ratio", "Extended Health Care Renewal Factors", "Dental Renewal Factors", "Weighted Adjusted Loss Ratio"]),
        ("Demographics", ["Demographic Summary", "Age Distribution", "Life Demographics", "Demographic distribution"]),
        ("Claims Distribution", ["Claims Distribution by Service", "Claims Distribution by Service Summary", "Medical Claims Distribution", "Service Summary", "Health Service Utilization", "Benefit Utilization Summary", "Claims Breakdown by Category", "Claims by Service Category", "Amount Paid Number of Occurrences"]),
        ("Paramedical Detail", ["Claims Distribution - Paramedical And Related Services - Detail", "Paramedical"]),
        ("Pooling", ["Pooling Arrangement", "Pooling Level", "Pooling Charge", "Large Amount Pooling Premium", "Pooled Claims"]),
        ("Schedule of Benefits", ["Schedule of Benefits", "GROUP INFORMATION REPORT"]),
        ("Plan Design", ["GROUP INFORMATION REPORT", "<<< Categories >>>", "Admin Class Listings"]),
        ("Group Information Report", ["GROUP INFORMATION REPORT"])
    ];

    public static CarrierDetectionResult DetectCarrier(string documentText)
    {
        var bestCarrier = "Unknown";
        var bestHits = 0;
        var bestClueCount = 1;
        foreach (var (carrier, clues) in CarrierClues)
        {
            var hits = clues.Count(clue => documentText.Contains(clue, StringComparison.OrdinalIgnoreCase));
            if (hits <= bestHits)
            {
                continue;
            }

            bestCarrier = carrier;
            bestHits = hits;
            bestClueCount = clues.Length;
        }

        return new CarrierDetectionResult
        {
            CarrierName = bestCarrier,
            ConfidenceScore = bestHits == 0 ? 0m : decimal.Round(Math.Min(1m, bestHits / (decimal)bestClueCount), 2)
        };
    }

    public static List<DetectedSection> DetectSections(IEnumerable<PdfPageText> pages)
    {
        var sections = new List<DetectedSection>();
        foreach (var page in pages)
        {
            foreach (var (sectionType, clues) in SectionClues)
            {
                var hits = clues.Count(clue => page.Text.Contains(clue, StringComparison.OrdinalIgnoreCase));
                if (hits == 0)
                {
                    continue;
                }

                sections.Add(new DetectedSection
                {
                    SectionType = sectionType,
                    SectionName = sectionType,
                    PageNumber = page.PageNumber,
                    ConfidenceScore = decimal.Round(Math.Min(1m, hits / (decimal)clues.Length), 2)
                });
            }
        }

        return sections
            .GroupBy(section => new { section.SectionType, section.PageNumber })
            .Select(group => group.OrderByDescending(section => section.ConfidenceScore).First())
            .OrderBy(section => section.PageNumber)
            .ThenBy(section => section.SectionType)
            .ToList();
    }

    public static List<PlanDesignFeatureExtraction> ExtractPlanDesignFeatures(IEnumerable<PdfPageText> pages)
    {
        var features = new List<PlanDesignFeatureExtraction>();
        foreach (var page in pages.Where(page => page.Text.Contains("GROUP INFORMATION REPORT", StringComparison.OrdinalIgnoreCase) ||
                                                 page.Text.Contains("Schedule of Benefits", StringComparison.OrdinalIgnoreCase)))
        {
            var benefit = "";
            var classNumber = "";
            var className = "";
            foreach (var rawLine in page.Lines.Select(ParsingHelpers.Clean))
            {
                if (string.IsNullOrWhiteSpace(rawLine))
                {
                    continue;
                }

                var classMatch = System.Text.RegularExpressions.Regex.Match(rawLine, @"^Class:\s*(?<number>\d+)\s*-\s*(?<name>.+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (classMatch.Success)
                {
                    classNumber = classMatch.Groups["number"].Value;
                    className = classMatch.Groups["name"].Value;
                    continue;
                }

                if (TryMapPlanDesignBenefit(rawLine, out var mappedBenefit))
                {
                    benefit = mappedBenefit;
                    continue;
                }

                var feature = TryParsePlanDesignFeature(rawLine, benefit, classNumber, className, page.PageNumber);
                if (feature is not null)
                {
                    features.Add(feature);
                }
            }
        }

        return features;
    }

    public static List<RenewalValidationItem> Validate(
        AnalysisResult analysis,
        CarrierDetectionResult carrierDetection,
        IReadOnlyCollection<DetectedSection> sections)
    {
        var rows = new List<RenewalValidationItem>();
        AddRequired(rows, "Carrier must be detected", carrierDetection.CarrierName != "Unknown", $"Detected carrier: {carrierDetection.CarrierName} ({carrierDetection.ConfidenceScore:0.00}).");
        AddRequired(rows, "Policy number must be present", !string.IsNullOrWhiteSpace(analysis.RenewalCase.PolicyNumber), $"Policy number: {analysis.RenewalCase.PolicyNumber}");
        AddRequired(rows, "Renewal effective date must be present", analysis.RenewalCase.RenewalDate.HasValue, $"Renewal date: {analysis.RenewalCase.RenewalDate:yyyy-MM-dd}");
        var additiveRateLines = analysis.Pdf.RateLines.Where(line => line.IncludeInPremiumTotal).ToList();
        AddPremiumReconciliation(rows, "Total current premium must reconcile", analysis.Pdf.CurrentMonthlyPremium, additiveRateLines.Select(line => line.CurrentPremium));
        AddPremiumReconciliation(rows, "Total renewal premium must reconcile", analysis.Pdf.RenewalMonthlyPremium, additiveRateLines.Select(line => line.RenewalPremium));
        AddPercentChangeValidation(rows, analysis);
        AddLossRatioValidation(rows, analysis);
        AddExperiencePeriodValidation(rows, analysis);
        AddClaimsDistributionValidation(rows, analysis);
        AddOptionalSectionWarning(rows, sections, "Claims Distribution");
        AddOptionalSectionWarning(rows, sections, "Paramedical Detail");
        AddOptionalSectionWarning(rows, sections, "Plan Design");
        return rows;
    }

    public static List<ExtractionAuditEntry> BuildAuditLog(
        string filePath,
        AnalysisResult analysis,
        CarrierDetectionResult carrierDetection,
        IReadOnlyList<DetectedSection> sections,
        IReadOnlyList<PlanDesignFeatureExtraction>? planDesignFeatures = null)
    {
        var fileName = Path.GetFileName(filePath);
        var rows = new List<ExtractionAuditEntry>
        {
            Audit(fileName, null, "Carrier Detection", "carrier_name", carrierDetection.CarrierName, carrierDetection.CarrierName, carrierDetection.ConfidenceScore),
            Audit(fileName, null, "Carrier Detection", "confidence_score", carrierDetection.ConfidenceScore.ToString("0.00"), carrierDetection.CarrierName, carrierDetection.ConfidenceScore),
            Audit(fileName, null, "Client / Policy Information", "client_name", analysis.RenewalCase.ClientName, analysis.RenewalCase.ClientName, 0.9m),
            Audit(fileName, null, "Client / Policy Information", "policy_number", analysis.RenewalCase.PolicyNumber, analysis.RenewalCase.PolicyNumber, 0.9m),
            Audit(fileName, null, "Client / Policy Information", "renewal_effective_date", analysis.RenewalCase.RenewalDate?.ToString("yyyy-MM-dd") ?? "", analysis.RenewalCase.RenewalDate?.ToString("yyyy-MM-dd") ?? "", 0.9m)
        };

        rows.AddRange(sections.Select(section => Audit(fileName, section.PageNumber, section.SectionType, "section_detected", section.SectionName, section.SectionName, section.ConfidenceScore)));

        foreach (var line in analysis.Pdf.RateLines)
        {
            var page = ParsePage(line.SourceReference);
            rows.Add(Audit(fileName, page, "Renewal Rates and Premium", $"{line.Benefit} {line.TierOrClass} current_premium", FormatDecimal(line.CurrentPremium), line.SourceReference, 0.85m));
            rows.Add(Audit(fileName, page, "Renewal Rates and Premium", $"{line.Benefit} {line.TierOrClass} renewal_premium", FormatDecimal(line.RenewalPremium), line.SourceReference, 0.85m));
        }

        foreach (var line in analysis.Pdf.ExperienceLines)
        {
            var page = ParsePage(line.SourceReference);
            rows.Add(Audit(fileName, page, "Experience Analysis", $"{line.Benefit} premium", FormatDecimal(line.BilledPremium), line.SourceReference, 0.85m));
            rows.Add(Audit(fileName, page, "Experience Analysis", $"{line.Benefit} paid_claims", FormatDecimal(line.PaidClaims), line.SourceReference, 0.85m));
            rows.Add(Audit(fileName, page, "Experience Analysis", $"{line.Benefit} incurred_claims", FormatDecimal(line.IncurredClaims), line.SourceReference, 0.85m));
            rows.Add(Audit(fileName, page, "Experience Analysis", $"{line.Benefit} incurred_loss_ratio", FormatDecimal(line.IncurredLossRatio), line.SourceReference, 0.85m));
        }

        foreach (var claim in analysis.Workbook.ClaimTrends)
        {
            rows.Add(Audit(fileName, ParsePage(claim.SourceReference), "Claims Distribution", $"{claim.Benefit} {claim.Category}", FormatDecimal(claim.Amount), claim.SourceReference, 0.8m));
        }

        foreach (var claim in analysis.Workbook.ClaimsDistributionRows)
        {
            var page = ParsePage(claim.SourceReference);
            rows.Add(Audit(fileName, page, "Claims Distribution", $"{claim.Benefit} {claim.ServiceCategory} amount_paid", FormatDecimal(claim.AmountPaid), claim.SourceReference, 0.85m));
            rows.Add(Audit(fileName, page, "Claims Distribution", $"{claim.Benefit} {claim.ServiceCategory} percent_paid", FormatDecimal(claim.PercentPaid), claim.SourceReference, 0.85m));
            rows.Add(Audit(fileName, page, "Claims Distribution", $"{claim.Benefit} {claim.ServiceCategory} norm_percent_paid", FormatDecimal(claim.NormPercentPaid), claim.SourceReference, 0.85m));
            rows.Add(Audit(fileName, page, "Claims Distribution", $"{claim.Benefit} {claim.ServiceCategory} occurrences", claim.Occurrences?.ToString() ?? "", claim.SourceReference, 0.85m));
            rows.Add(Audit(fileName, page, "Claims Distribution", $"{claim.Benefit} {claim.ServiceCategory} average_claim", FormatDecimal(claim.AverageClaim), claim.SourceReference, 0.85m));
        }

        foreach (var demographic in analysis.Pdf.LifeDemographics.Concat(analysis.Pdf.LtdDemographics))
        {
            rows.Add(Audit(fileName, ParsePage(demographic.SourceReference), "Demographics", $"{demographic.Benefit} {demographic.AgeBand}", $"{demographic.CurrentMaleLives}/{demographic.CurrentFemaleLives}", demographic.SourceReference, 0.8m));
        }

        foreach (var feature in planDesignFeatures ?? [])
        {
            rows.Add(Audit(fileName, feature.SourcePage, "Plan Design", $"{feature.Benefit} {feature.FeatureName}", feature.FeatureValue, feature.SourceText, feature.ConfidenceScore));
        }

        foreach (var nem in analysis.NemMaximums)
        {
            var extractedValue = string.Join(" | ", new[]
                {
                    string.IsNullOrWhiteSpace(nem.CurrentMaximum) ? "" : $"Current Maximum {nem.CurrentMaximum}",
                    string.IsNullOrWhiteSpace(nem.CurrentNem) ? "" : $"Current NEM {nem.CurrentNem}",
                    string.IsNullOrWhiteSpace(nem.AvailableNem) ? "" : $"Available NEM {nem.AvailableNem}"
                }
                .Where(value => !string.IsNullOrWhiteSpace(value)));
            rows.Add(Audit(
                string.IsNullOrWhiteSpace(nem.SourceFileName) ? fileName : nem.SourceFileName,
                nem.SourcePage,
                "Plan Maximums and NEMs",
                nem.Benefit,
                extractedValue,
                nem.SourceReference,
                nem.ConfidenceScore));
        }

        return rows;
    }

    private static PlanDesignFeatureExtraction? TryParsePlanDesignFeature(
        string line,
        string benefit,
        string classNumber,
        string className,
        int pageNumber)
    {
        if (string.IsNullOrWhiteSpace(benefit) ||
            line.StartsWith("For:", StringComparison.OrdinalIgnoreCase) ||
            line.StartsWith("By:", StringComparison.OrdinalIgnoreCase) ||
            line.StartsWith("Quote:", StringComparison.OrdinalIgnoreCase) ||
            line.Contains("Great-West Life Assurance", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        string featureName;
        string featureValue;
        var arrowParts = line.Split(">", 2, StringSplitOptions.TrimEntries);
        if (arrowParts.Length == 2)
        {
            featureName = arrowParts[0].Replace("-", "", StringComparison.Ordinal).Trim();
            featureValue = arrowParts[1];
        }
        else
        {
            var colonParts = line.Split(":", 2, StringSplitOptions.TrimEntries);
            if (colonParts.Length != 2)
            {
                return null;
            }

            featureName = colonParts[0];
            featureValue = colonParts[1];
        }

        if (string.IsNullOrWhiteSpace(featureName) || string.IsNullOrWhiteSpace(featureValue))
        {
            return null;
        }

        return new PlanDesignFeatureExtraction
        {
            Benefit = benefit,
            Category = benefit,
            FeatureName = featureName,
            FeatureValue = featureValue,
            Coinsurance = ExtractFeatureValue(line, @"Coinsurance:\s*(?<value>[^ ]+%?)"),
            Deductible = line.Contains("Deductible", StringComparison.OrdinalIgnoreCase) ? featureValue : "",
            Maximum = ExtractFeatureValue(line, @"Maximum:\s*(?<value>[^:]+?)(?:\s+[A-Z][A-Za-z]+:|$)"),
            Frequency = ExtractFeatureValue(line, @"(?<value>\d+\s+(?:Visits|Pairs|Units|Months|Years)[^:]*)"),
            TerminationAge = ExtractFeatureValue(line, @"Termination Age:\s*(?<value>\d+)"),
            ClassNumber = classNumber,
            ClassName = className,
            SourcePage = pageNumber,
            SourceText = line,
            ConfidenceScore = 0.7m
        };
    }

    private static bool TryMapPlanDesignBenefit(string line, out string benefit)
    {
        benefit = line.ToUpperInvariant() switch
        {
            "EMPLOYEE TERM LIFE INSURANCE" => "Life",
            "ACCIDENTAL DEATH AND DISMEMBERMENT" => "AD&D",
            "HEALTHCARE" => "EHC",
            "DRUGS" => "Drugs",
            "VISIONCARE" => "Vision",
            "DENTAL" => "Dental",
            "<<< ADMIN CLASS LISTINGS >>>" => "Admin Class",
            _ => ""
        };

        return !string.IsNullOrWhiteSpace(benefit);
    }

    private static string ExtractFeatureValue(string text, string pattern)
    {
        var match = System.Text.RegularExpressions.Regex.Match(text, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return match.Success ? ParsingHelpers.Clean(match.Groups["value"].Value) : "";
    }

    private static ExtractionAuditEntry Audit(
        string fileName,
        int? page,
        string section,
        string field,
        string value,
        string sourceText,
        decimal confidence)
    {
        return new ExtractionAuditEntry
        {
            SourceFile = fileName,
            SourcePage = page,
            SectionType = section,
            FieldName = ParsingHelpers.Clean(field),
            ExtractedValue = value,
            SourceText = sourceText,
            ConfidenceScore = confidence,
            ValidationStatus = string.IsNullOrWhiteSpace(value) ? "Needs review" : "Read",
            ValidationMessage = string.IsNullOrWhiteSpace(value) ? "Value was not extracted." : "Value extracted."
        };
    }

    private static void AddRequired(List<RenewalValidationItem> rows, string rule, bool passed, string message)
    {
        rows.Add(new RenewalValidationItem
        {
            Rule = rule,
            Status = passed ? "Pass" : "Error",
            Message = message
        });
    }

    private static void AddPremiumReconciliation(List<RenewalValidationItem> rows, string rule, decimal? total, IEnumerable<decimal?> values)
    {
        if (!total.HasValue)
        {
            rows.Add(new RenewalValidationItem { Rule = rule, Status = "Warning", Message = "Carrier total was not available." });
            return;
        }

        var sum = values.Where(value => value.HasValue).Sum(value => value!.Value);
        var difference = Math.Abs(total.Value - sum);
        rows.Add(new RenewalValidationItem
        {
            Rule = rule,
            Status = difference <= 0.05m ? "Pass" : "Warning",
            Message = $"Carrier total {total.Value:C2}; extracted row total {sum:C2}; difference {difference:C2}."
        });
    }

    private static void AddPercentChangeValidation(List<RenewalValidationItem> rows, AnalysisResult analysis)
    {
        if (analysis.CurrentMonthlyPremium is not > 0m ||
            !analysis.RenewalMonthlyPremium.HasValue ||
            !analysis.MonthlyPremiumChangePercent.HasValue)
        {
            rows.Add(new RenewalValidationItem { Rule = "Percent change should calculate correctly", Status = "Warning", Message = "Premium values were not complete enough to calculate percent change." });
            return;
        }

        var calculated = UnderwritingEngine.CalculatePercentChange(analysis.CurrentMonthlyPremium, analysis.RenewalMonthlyPremium);
        var difference = Math.Abs((calculated ?? 0m) - analysis.MonthlyPremiumChangePercent.Value);
        rows.Add(new RenewalValidationItem
        {
            Rule = "Percent change should calculate correctly",
            Status = difference <= 0.1m ? "Pass" : "Warning",
            Message = $"Extracted {analysis.MonthlyPremiumChangePercent.Value:0.0}%; calculated {calculated:0.0}%."
        });
    }

    private static void AddLossRatioValidation(List<RenewalValidationItem> rows, AnalysisResult analysis)
    {
        var checkedRows = 0;
        var warnings = 0;
        foreach (var row in analysis.Pdf.ExperienceLines.Where(row => row.BilledPremium is > 0m && row.PaidClaims.HasValue && row.PaidLossRatio.HasValue))
        {
            checkedRows++;
            var calculated = decimal.Round(row.PaidClaims!.Value / row.BilledPremium!.Value * 100m, 1);
            if (Math.Abs(calculated - row.PaidLossRatio!.Value) > 0.2m)
            {
                warnings++;
            }
        }

        rows.Add(new RenewalValidationItem
        {
            Rule = "Paid loss ratio should equal paid claims divided by billed premium where possible",
            Status = warnings == 0 ? "Pass" : "Warning",
            Message = checkedRows == 0 ? "No loss-ratio rows were available to validate." : $"{checkedRows} row(s) checked; {warnings} variance(s) found."
        });
    }

    private static void AddExperiencePeriodValidation(List<RenewalValidationItem> rows, AnalysisResult analysis)
    {
        var invalid = analysis.Pdf.ExperienceLines.Count(row => !IsValidPeriod(row.Period));
        rows.Add(new RenewalValidationItem
        {
            Rule = "Experience periods must be valid date ranges",
            Status = invalid == 0 ? "Pass" : "Warning",
            Message = invalid == 0 ? "All extracted experience periods look valid." : $"{invalid} experience period(s) need review."
        });
    }

    private static void AddClaimsDistributionValidation(List<RenewalValidationItem> rows, AnalysisResult analysis)
    {
        var validation = analysis.Workbook.ClaimsDistributionRows.Count > 0
            ? ClaimsDistributionExtractor.BuildValidation(analysis.Workbook.ClaimsDistributionRows)
            : analysis.Workbook.ClaimsDistributionValidation;

        rows.Add(new RenewalValidationItem
        {
            Rule = "Claims Distribution table found",
            Status = validation.TableFound ? "Pass" : "Warning",
            Message = validation.TableFound ? $"{analysis.Workbook.ClaimsDistributionRows.Count} claims distribution row(s) extracted." : "Claims Distribution table not found."
        });

        if (!validation.TableFound)
        {
            return;
        }

        rows.Add(new RenewalValidationItem
        {
            Rule = "Claims Distribution paramedical row found",
            Status = validation.ParamedicalFound ? "Pass" : "Warning",
            Message = validation.ParamedicalFound ? "Paramedical row was extracted." : "Paramedical row was not extracted."
        });

        rows.Add(new RenewalValidationItem
        {
            Rule = "Claims Distribution total row found",
            Status = validation.TotalRowFound ? "Pass" : "Warning",
            Message = validation.TotalRowFound ? $"Total amount paid: {validation.TotalAmountPaid:C2}." : "Total row was not extracted."
        });

        if (validation.TotalRowFound)
        {
            rows.Add(new RenewalValidationItem
            {
                Rule = "Claims Distribution service amounts balance to total within $1",
                Status = validation.AmountsBalanceToTotal ? "Pass" : "Warning",
                Message = $"Service sum {validation.SumOfServiceAmounts:C2}; total {validation.TotalAmountPaid:C2}."
            });
        }

        AddParamedicalDetailValidation(rows, validation, prior: false);
        AddParamedicalDetailValidation(rows, validation, prior: true);

        foreach (var warning in validation.Warnings)
        {
            rows.Add(new RenewalValidationItem
            {
                Rule = "Claims Distribution warning",
                Status = "Warning",
                Message = warning
            });
        }
    }

    private static void AddParamedicalDetailValidation(
        List<RenewalValidationItem> rows,
        ClaimsDistributionValidation validation,
        bool prior)
    {
        var subtotal = prior
            ? validation.PriorParamedicalSubtotalAmountPaid
            : validation.ParamedicalSubtotalAmountPaid;
        var detailSum = prior
            ? validation.SumOfPriorParamedicalDetailAmounts
            : validation.SumOfParamedicalDetailAmounts;
        var variance = prior
            ? validation.PriorParamedicalDetailVariance
            : validation.ParamedicalDetailVariance;
        var balances = prior
            ? validation.PriorParamedicalDetailsBalanceToTotal
            : validation.ParamedicalDetailsBalanceToTotal;
        if (!subtotal.HasValue)
        {
            return;
        }

        var needsReview = balances != true || validation.UnclassifiedParamedicalDetailFound;
        rows.Add(new RenewalValidationItem
        {
            Rule = $"{(prior ? "Prior" : "Current")} paramedical practitioner amounts balance to carrier subtotal within $1",
            Status = needsReview ? "Warning" : "Pass",
            Message = $"Practitioner sum {detailSum:C2}; carrier subtotal {subtotal:C2}; variance {variance:C2}." +
                      (validation.UnclassifiedParamedicalDetailFound
                          ? " An unclassified amount was retained and needs source review."
                          : "")
        });
    }

    private static void AddOptionalSectionWarning(List<RenewalValidationItem> rows, IReadOnlyCollection<DetectedSection> sections, string sectionType)
    {
        if (sections.Any(section => section.SectionType.Equals(sectionType, StringComparison.OrdinalIgnoreCase)))
        {
            rows.Add(new RenewalValidationItem { Rule = $"{sectionType} section", Status = "Pass", Message = $"{sectionType} section was detected." });
        }
        else
        {
            rows.Add(new RenewalValidationItem { Rule = $"{sectionType} section", Status = "Warning", Message = $"{sectionType} section was not detected. This is optional unless needed for this review." });
        }
    }

    private static bool IsValidPeriod(string period)
    {
        var parts = period.Split(" - ", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        return parts.Length == 2 &&
               DateOnly.TryParse(parts[0], out var start) &&
               DateOnly.TryParse(parts[1], out var end) &&
               end >= start;
    }

    private static int? ParsePage(string sourceReference)
    {
        var match = System.Text.RegularExpressions.Regex.Match(sourceReference, @"page\s+(?<page>\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return match.Success && int.TryParse(match.Groups["page"].Value, out var page) ? page : null;
    }

    private static string FormatDecimal(decimal? value)
    {
        return value?.ToString("0.##") ?? "";
    }

}
