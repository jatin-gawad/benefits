using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.ClientHistory;
using System.Security.Cryptography;

namespace RenewalReviewBuilder.App.Services;

public sealed class ReviewPackBuilder
{
    private const string WorkbookSourceArea = "Carrier / Claims Excel";
    private const string LegacyWorkbookSourceArea = "Manulife XLSM";

    private readonly CarrierPdfParser _pdfParser = new();
    private readonly PreviousRenewalReportParser _previousRenewalParser = new();
    private readonly ManulifeWorkbookParser _workbookParser = new();
    private readonly AnalysisEngine _analysisEngine = new();
    private readonly DatabaseService? _databaseService;
    private readonly PowerPointDeckWriter _deckWriter = new();
    private readonly NemMaximumExtractor _nemMaximumExtractor = new();
    private readonly DocumentDiscoveryService _documentDiscoveryService = new();
    private readonly AssurePlusExperienceSummaryBuilder _assurePlusSummaryBuilder = new();
    private readonly GroupHealthClaimsRangePdfParser _claimsRangeParser = new();
    private readonly EquitableClaimsRangePdfParser _equitableClaimsRangeParser = new();
    private readonly ClaimsRangeComparisonBuilder _claimsRangeComparisonBuilder = new();
    private readonly AsoMonthlyActivityPdfParser _asoMonthlyActivityPdfParser = new();
    private readonly SunLifeAsoInvoicePdfParser _sunLifeAsoInvoicePdfParser = new();
    private readonly SunLifeOperatingAccountsPdfParser _sunLifeOperatingAccountsPdfParser = new();

    public ReviewPackBuilder(DatabaseService? databaseService = null)
    {
        _databaseService = databaseService;
    }

    public static bool SourceInputsChangedSince(RenewalCase renewalCase, DateTime analyzedAt)
    {
        var analyzedAtUtc = analyzedAt.Kind == DateTimeKind.Utc
            ? analyzedAt
            : analyzedAt.ToUniversalTime();

        try
        {
            foreach (var folder in new[] { renewalCase.CurrentYearFolder, renewalCase.PriorYearFolder }
                         .Where(folder => !string.IsNullOrWhiteSpace(folder)))
            {
                if (!Directory.Exists(folder))
                {
                    return true;
                }
            }

            var sourcePaths = new[] { renewalCase.PdfPath, renewalCase.WorkbookPath, renewalCase.PowerPointTemplatePath }
                .Concat(renewalCase.AdditionalPdfPaths)
                .Concat(renewalCase.AdditionalWorkbookPaths)
                .Concat(EnumerateSourceFolderFiles(renewalCase.CurrentYearFolder))
                .Concat(EnumerateSourceFolderFiles(renewalCase.PriorYearFolder))
                .Concat(EnumerateSourceFolderFiles(renewalCase.AsoBillingFolder))
                .Concat(GetNearbySupportingWorkbookPaths(renewalCase.WorkbookPath))
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Where(path => !IsGeneratedOutputPath(path))
                .Distinct(StringComparer.OrdinalIgnoreCase);

            foreach (var sourcePath in sourcePaths)
            {
                if (!File.Exists(sourcePath) ||
                    File.GetLastWriteTimeUtc(sourcePath) > analyzedAtUtc)
                {
                    return true;
                }
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return true;
        }

        return false;
    }

    public Task<AnalysisResult> AnalyzeAsync(RenewalCase renewalCase, IProgress<string>? progress = null)
    {
        return Task.Run(() =>
        {
            var result = new AnalysisResult
            {
                RenewalCase = renewalCase,
                CreatedAt = DateTime.Now
            };
            result.DocumentDiscovery = DiscoverInputDocuments(renewalCase);
            AddDocumentDiscoverySourceChecks(result);

            ValidateInputs(renewalCase, result);
            WorkbookExtractionResult? renewalPdfWorkbook = null;
            string renewalPdfCarrier = "";

            if (File.Exists(renewalCase.PdfPath))
            {
                progress?.Report("Reading renewal PDF...");
                try
                {
                    var parsedPdf = _pdfParser.Parse(renewalCase.PdfPath);
                    result.Pdf = parsedPdf.Pdf;
                    renewalPdfWorkbook = parsedPdf.WorkbookFromPdf;
                    renewalPdfCarrier = parsedPdf.Carrier;
                    if (renewalPdfWorkbook is not null)
                    {
                        result.Workbook = renewalPdfWorkbook;
                    }
                    result.RenewalCase.ClientName = FirstNonBlank(renewalCase.ClientName, result.Pdf.ClientName);
                    result.RenewalCase.PolicyNumber = FirstNonBlank(renewalCase.PolicyNumber, result.Pdf.PolicyNumber);
                    result.RenewalCase.RenewalDate = renewalCase.RenewalDate ?? result.Pdf.RenewalDate;
                    result.SourceChecks.Add(new SourceCheck { Area = "Renewal PDF", Status = "Read", Detail = renewalCase.PdfPath });

                    if (renewalPdfWorkbook is not null && !File.Exists(renewalCase.WorkbookPath))
                    {
                        RemoveMissingWorkbookChecks(result);
                        result.SourceChecks.Add(new SourceCheck
                        {
                            Area = $"{parsedPdf.Carrier} PDF claims detail",
                            Status = "Read",
                            Detail = "Claims categories, top drugs, and paid-loss-ratio detail were mapped from the renewal PDF. No separate Excel workbook was selected."
                        });
                    }
                }
                catch (Exception ex)
                {
                    result.SourceChecks.Add(new SourceCheck { Area = "Renewal PDF", Status = "Failed", Detail = ex.Message });
                    result.Findings.Add(new Finding
                    {
                        Severity = FindingSeverity.Critical,
                        Title = "Renewal PDF could not be read",
                        Evidence = ex.Message,
                        RecommendedAction = "Confirm the PDF is not password-protected or scanned-only.",
                        ClientExplanation = "PDF extraction failed, so renewal illustration values need manual review.",
                        SourceReference = renewalCase.PdfPath
                    });
                }
            }

            progress?.Report("Reading ASO operating accounts...");
            ReadOperatingAccounts(renewalCase, result);

            if (File.Exists(renewalCase.WorkbookPath) && !IsLegacyClaimsWorkbookPath(renewalCase.WorkbookPath))
            {
                progress?.Report("Reading carrier claims workbook...");
                var currentCarrier = FirstNonBlank(renewalPdfCarrier, result.Pdf.Carrier);
                var claimsWorkbookPaths = GetSupportingWorkbookPaths(renewalCase);
                try
                {
                    result.Workbook = _workbookParser.Parse(
                        renewalCase.WorkbookPath,
                        result.RenewalCase.ClientName,
                        claimsWorkbookPaths,
                        renewalCase.CurrentYearFolder,
                        renewalCase.PriorYearFolder,
                        result.RenewalCase.PolicyNumber);
                    result.SourceChecks.Add(new SourceCheck { Area = WorkbookSourceArea, Status = "Read", Detail = renewalCase.WorkbookPath });
                    if (ShouldTreatWorkbookAsPriorCarrierHistory(renewalCase, currentCarrier))
                    {
                        result.SourceChecks.Add(new SourceCheck
                        {
                            Area = "Prior carrier claims history",
                            Status = "Read",
                            Detail = $"Using selected claims-history workbook as prior experience because the current renewal carrier is {currentCarrier}."
                        });

                        var promotedTopDrugs = PromoteClaimsWorkbookTopDrugsToPriorHistory(result.Workbook);
                        if (promotedTopDrugs > 0)
                        {
                            result.SourceChecks.Add(new SourceCheck
                            {
                                Area = "Prior Top DINs",
                                Status = "Mapped",
                                Detail = $"{promotedTopDrugs} claims-history Top DIN row(s) were mapped to the prior-drug PowerPoint table."
                            });
                        }
                    }

                    var mergedRows = MergeCurrentWorkbookData(result.Workbook, renewalPdfWorkbook);
                    if (mergedRows > 0)
                    {
                        result.SourceChecks.Add(new SourceCheck
                        {
                            Area = $"{FirstNonBlank(renewalPdfCarrier, result.Pdf.Carrier)} PDF claims detail",
                            Status = "Merged",
                            Detail = $"{mergedRows} PDF-derived claims row(s) were merged into the renewal workbook output."
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (TryReadCurrentManulifeClientCopyWithoutPolicyMatch(
                            renewalCase,
                            result,
                            renewalPdfWorkbook,
                            claimsWorkbookPaths,
                            currentCarrier,
                            ex,
                            out var currentWorkbookRetryDetail))
                    {
                        result.SourceChecks.Add(new SourceCheck
                        {
                            Area = WorkbookSourceArea,
                            Status = "Read",
                            Detail = currentWorkbookRetryDetail
                        });
                        goto WorkbookReadComplete;
                    }

                    if (TryReadClaimsWorkbookAsPriorHistory(
                            renewalCase,
                            result,
                            renewalPdfWorkbook,
                            claimsWorkbookPaths,
                            currentCarrier,
                            ex,
                            out var retryDetail))
                    {
                        result.SourceChecks.Add(new SourceCheck
                        {
                            Area = "Prior carrier claims history",
                            Status = "Read",
                            Detail = retryDetail
                        });
                        goto WorkbookReadComplete;
                    }

                    result.SourceChecks.Add(new SourceCheck { Area = WorkbookSourceArea, Status = "Failed", Detail = ex.Message });
                    result.Findings.Add(new Finding
                    {
                        Severity = FindingSeverity.Critical,
                        Title = "Carrier claims workbook could not be read",
                        Evidence = ex.Message,
                        RecommendedAction = "Confirm the Excel file is the carrier or claims-history workbook and is not locked by Excel.",
                        ClientExplanation = "Workbook extraction failed, so detailed claims analytics are unavailable.",
                        SourceReference = renewalCase.WorkbookPath
                    });
                }

                WorkbookReadComplete:;
            }

            progress?.Report("Reading GroupHEALTH claims-range reports...");
            ReadClaimsRangeReports(renewalCase, result);

            progress?.Report("Reading ASO monthly billing activity...");
            ReadAsoBillingActivityReports(renewalCase, result);

            progress?.Report("Checking additional PDFs for prior renewal reports...");
            ReadAdditionalRenewalPdfs(renewalCase, result);

            progress?.Report("Checking supporting Excel claims history...");
            ReadSupportingClaimsWorkbooks(renewalCase, result);

            AlignServiceCategoryHistory(result.Workbook);
            BuildAssurePlusExperience(result);

            progress?.Report("Reading plan maximums and NEMs...");
            ReadNemMaximums(renewalCase, result);

            ApplyHcsaCalculation(renewalCase, result);
            ApplyAllstateCriticalIllnessEnrollment(renewalCase, result);

            progress?.Report("Analyzing renewal experience...");
            _analysisEngine.Analyze(result);
            SortFindings(result);
            progress?.Report("Analysis complete.");

            return result;
        });
    }

    private static void AlignServiceCategoryHistory(WorkbookExtractionResult workbook)
    {
        ServiceCategoryHistory.AlignToSharedWindow(workbook.EhcSpendingTrends);
        ServiceCategoryHistory.AlignToSharedWindow(workbook.DentalSpendingTrends);
        ServiceCategoryHistory.AlignToSharedWindow(workbook.ParamedicalServiceTrends);
    }

    private static void ApplyHcsaCalculation(RenewalCase renewalCase, AnalysisResult result)
    {
        new HcsaRateLineService().Apply(result, renewalCase.HcsaCalculation);
    }

    private static void ApplyAllstateCriticalIllnessEnrollment(RenewalCase renewalCase, AnalysisResult result)
    {
        new AllstateCriticalIllnessRateLineService().Apply(
            result,
            renewalCase.AllstateCriticalIllnessEnrollment);
    }

    private void ReadClaimsRangeReports(RenewalCase renewalCase, AnalysisResult result)
    {
        var isEquitable = result.Pdf.Carrier.Equals("Equitable", StringComparison.OrdinalIgnoreCase);
        var discoveredPaths = result.DocumentDiscovery?.Documents
            .Where(document => document.DocumentType.Equals(
                DiscoveryDocumentTypes.GroupHealthClaimsRange,
                StringComparison.OrdinalIgnoreCase))
            .Select(document => document.FilePath)
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];
        var candidatePaths = GetInputPdfPaths(renewalCase)
            .Where(path => isEquitable ||
                           discoveredPaths.Contains(path) ||
                           Path.GetFileNameWithoutExtension(path)
                               .Contains("claims range", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (candidatePaths.Count == 0)
        {
            result.ClaimsRange = null;
            return;
        }

        var reports = new List<ClaimsRangeReport>();
        foreach (var path in candidatePaths)
        {
            try
            {
                var parsedReports = isEquitable
                    ? _equitableClaimsRangeParser.Parse(path)
                    : [_claimsRangeParser.Parse(path)];
                foreach (var report in parsedReports)
                {
                    reports.Add(report);
                    result.SourceChecks.Add(new SourceCheck
                    {
                        Area = $"{report.Carrier} Claims Range",
                        Status = report.IsComplete ? "Read" : "Needs Review",
                        Detail = report.IsComplete
                            ? $"{report.SourceFileName}: {report.RowCount} band rows for {report.PeriodLabel}."
                            : $"{report.SourceFileName}: {string.Join(" ", report.Errors)}"
                    });
                }
            }
            catch (Exception ex)
            {
                result.SourceChecks.Add(new SourceCheck
                {
                    Area = "GroupHEALTH Claims Range",
                    Status = "Failed",
                    Detail = $"{Path.GetFileName(path)}: {ex.Message}"
                });
            }
        }

        reports = reports
            .GroupBy(report =>
                $"{NormalizePolicyMergeKey(report.PolicyNumber)}|{report.PeriodStart:yyyyMMdd}|{report.PeriodEnd:yyyyMMdd}",
                StringComparer.OrdinalIgnoreCase)
            .Select(group => group
                .OrderBy(report => PathsEqual(report.SourcePath, renewalCase.PdfPath) ? 0 : 1)
                .ThenByDescending(report => report.IsComplete)
                .First())
            .ToList();

        result.ClaimsRange = _claimsRangeComparisonBuilder.Build(
            reports,
            renewalCase.RenewalDate ?? result.Pdf.RenewalDate);
        var comparison = result.ClaimsRange;
        result.SourceChecks.Add(new SourceCheck
        {
            Area = "Claims Range comparison",
            Status = comparison.CanUseForClientOutput ? "Validated" : "Needs Review",
            Detail = comparison.PriorReport is null || comparison.CurrentReport is null
                ? $"{reports.Count} report(s) were read, but a complete prior/current pair was not available."
                : $"{comparison.PriorReport.PeriodLabel} compared with {comparison.CurrentReport.PeriodLabel}; " +
                  $"{comparison.Rows.Count} paired Health and Dental band rows."
        });

        if (comparison.CanUseForClientOutput)
        {
            return;
        }

        result.Findings.Add(new Finding
        {
            Severity = FindingSeverity.Warning,
            Title = "Claims Range comparison requires review",
            Evidence = comparison.ValidationIssues.Count == 0
                ? "The Claims Range report pair is not client-ready."
                : string.Join(" ", comparison.ValidationIssues.Select(issue => issue.Message)),
            RecommendedAction = comparison.IdentityConfirmationRequired
                ? "Open the Claims Range tab and confirm that the two client identities belong to this renewal session."
                : "Review the cited source pages and resolve the missing, conflicting, or non-contiguous report data.",
            ClientExplanation = "The Claims Range slide will remain suppressed until the report pair is complete and confirmed.",
            SourceReference = string.Join("; ", reports.Select(report => report.SourcePath))
        });
    }

    private void BuildAssurePlusExperience(AnalysisResult result)
    {
        result.AssurePlusExperience = _assurePlusSummaryBuilder.Build(
            result.Pdf.AssurePlusPricingReports,
            result.RenewalCase.RenewalDate ?? result.Pdf.RenewalDate);
        if (result.AssurePlusExperience is null)
        {
            return;
        }

        var summary = result.AssurePlusExperience;
        var years = summary.AnnualRatings
            .Select(row => row.Year)
            .Distinct()
            .OrderBy(year => year)
            .ToList();
        var yearRange = years.Count == 0
            ? "no annual status rows"
            : years.Count == 1
                ? years[0].ToString()
                : $"{years[0]}-{years[^1]}";
        result.SourceChecks.Add(new SourceCheck
        {
            Area = "AssurePLUS experience",
            Status = summary.Warnings.Count == 0 ? "Read" : "Warning",
            Detail = $"{result.Pdf.AssurePlusPricingReports.Count} pricing report(s) reconciled into {yearRange} experience history."
        });

        if (summary.Warnings.Count == 0)
        {
            return;
        }

        result.Findings.Add(new Finding
        {
            Severity = FindingSeverity.Warning,
            Title = "AssurePLUS experience requires review",
            Evidence = string.Join(" ", summary.Warnings),
            RecommendedAction = "Confirm the AssurePLUS status and pricing values against the cited renewal report pages before presenting the slide.",
            ClientExplanation = "One or more AssurePLUS history or pricing values could not be reconciled automatically.",
            SourceReference = string.Join("; ", result.Pdf.AssurePlusPricingReports
                .Select(report => report.SourceReference)
                .Distinct(StringComparer.OrdinalIgnoreCase))
        });
    }

    private static void RemoveMissingWorkbookChecks(AnalysisResult result)
    {
        result.SourceChecks.RemoveAll(check =>
            (check.Area.Equals(WorkbookSourceArea, StringComparison.OrdinalIgnoreCase) ||
             check.Area.Equals(LegacyWorkbookSourceArea, StringComparison.OrdinalIgnoreCase)) &&
            check.Status.Equals("Missing", StringComparison.OrdinalIgnoreCase));
        result.Findings.RemoveAll(finding =>
            finding.Title.Equals($"{WorkbookSourceArea} is missing", StringComparison.OrdinalIgnoreCase) ||
            finding.Title.Equals($"{LegacyWorkbookSourceArea} is missing", StringComparison.OrdinalIgnoreCase));
    }

    private void ReadSupportingClaimsWorkbooks(RenewalCase renewalCase, AnalysisResult result)
    {
        var selectedWorkbookIsModern = File.Exists(renewalCase.WorkbookPath) &&
                                       !IsLegacyClaimsWorkbookPath(renewalCase.WorkbookPath);
        var hasExplicitSupportingSelection = renewalCase.AdditionalWorkbookPaths.Any(File.Exists) ||
                                             (!string.IsNullOrWhiteSpace(renewalCase.CurrentYearFolder) && Directory.Exists(renewalCase.CurrentYearFolder)) ||
                                             (!string.IsNullOrWhiteSpace(renewalCase.PriorYearFolder) && Directory.Exists(renewalCase.PriorYearFolder));
        if (selectedWorkbookIsModern && !hasExplicitSupportingSelection)
        {
            return;
        }

        try
        {
            var updatedRows = _workbookParser.MergeSupportingClaimsWorkbookHistory(
                result.Workbook,
                renewalCase.PdfPath,
                result.RenewalCase.ClientName,
                GetSupportingWorkbookPaths(renewalCase),
                renewalCase.CurrentYearFolder,
                renewalCase.PriorYearFolder,
                result.RenewalCase.PolicyNumber);
            if (updatedRows <= 0)
            {
                return;
            }

            RemoveMissingWorkbookChecks(result);
            result.SourceChecks.Add(new SourceCheck
            {
                Area = "Supporting Excel claims history",
                Status = "Merged",
                Detail = $"{updatedRows} prior-period claims row(s) were merged from supporting Excel workbooks."
            });
        }
        catch (Exception ex)
        {
            result.SourceChecks.Add(new SourceCheck
            {
                Area = "Supporting Excel claims history",
                Status = "Warning",
                Detail = ex.Message
            });
        }
    }

    private bool TryReadCurrentManulifeClientCopyWithoutPolicyMatch(
        RenewalCase renewalCase,
        AnalysisResult result,
        WorkbookExtractionResult? renewalPdfWorkbook,
        IReadOnlyCollection<string> claimsWorkbookPaths,
        string currentCarrier,
        Exception originalException,
        out string detail)
    {
        detail = "";
        if (!currentCarrier.Contains("Manulife", StringComparison.OrdinalIgnoreCase) ||
            !FileNameLooksLikeClientCopyWorkbook(renewalCase.WorkbookPath) ||
            !originalException.Message.Contains("Claims workbook mismatch", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        try
        {
            result.Workbook = _workbookParser.Parse(
                renewalCase.WorkbookPath,
                result.RenewalCase.ClientName,
                claimsWorkbookPaths,
                renewalCase.CurrentYearFolder,
                renewalCase.PriorYearFolder,
                policyNumber: null);

            RemoveMissingWorkbookChecks(result);
            var mergedRows = MergeCurrentWorkbookData(result.Workbook, renewalPdfWorkbook);
            detail = $"Read current Manulife Client_Copy workbook after accepting its related policy/contract number. " +
                     $"Original validation: {originalException.Message}";
            if (mergedRows > 0)
            {
                detail += $" {mergedRows} current renewal PDF row(s) were also merged.";
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private bool TryReadClaimsWorkbookAsPriorHistory(
        RenewalCase renewalCase,
        AnalysisResult result,
        WorkbookExtractionResult? renewalPdfWorkbook,
        IReadOnlyCollection<string> claimsWorkbookPaths,
        string currentCarrier,
        Exception originalException,
        out string detail)
    {
        detail = "";
        if (!ShouldRetryClaimsWorkbookAsPriorCarrierHistory(renewalCase, originalException))
        {
            return false;
        }

        try
        {
            result.Workbook = _workbookParser.Parse(
                renewalCase.WorkbookPath,
                result.RenewalCase.ClientName,
                claimsWorkbookPaths,
                renewalCase.CurrentYearFolder,
                renewalCase.PriorYearFolder,
                policyNumber: null);

            RemoveMissingWorkbookChecks(result);

            var promotedTopDrugs = PromoteClaimsWorkbookTopDrugsToPriorHistory(result.Workbook);
            var mergedRows = MergeCurrentWorkbookData(result.Workbook, renewalPdfWorkbook);
            var parts = new List<string>
            {
                $"Policy/contract mismatch was accepted for prior-carrier history: {originalException.Message}"
            };

            if (!string.IsNullOrWhiteSpace(currentCarrier))
            {
                parts.Add($"current renewal carrier is {currentCarrier}");
            }

            if (promotedTopDrugs > 0)
            {
                parts.Add($"{promotedTopDrugs} Top DIN row(s) mapped to prior-drug output");
                result.SourceChecks.Add(new SourceCheck
                {
                    Area = "Prior Top DINs",
                    Status = "Mapped",
                    Detail = $"{promotedTopDrugs} claims-history Top DIN row(s) were mapped to the prior-drug PowerPoint table."
                });
            }

            if (mergedRows > 0)
            {
                parts.Add($"{mergedRows} current renewal PDF row(s) merged over prior history");
            }

            detail = string.Join("; ", parts);
            return true;
        }
        catch
        {
            return false;
        }
    }

    public Task<ExportResult> ExportAsync(AnalysisResult analysis, IProgress<string>? progress = null)
    {
        return Task.Run(async () =>
        {
            var presentationRateErrors = ClientPresentationRateSelector.ValidateSelectedVersion(analysis);
            if (presentationRateErrors.Count > 0)
            {
                throw new InvalidOperationException(
                    "Client presentation rates are incomplete:" + Environment.NewLine +
                    string.Join(Environment.NewLine, presentationRateErrors.Select(error => $"- {error}")));
            }

            Directory.CreateDirectory(analysis.RenewalCase.OutputFolder);

            progress?.Report("Copying supporting files...");
            var copiedSupportingPdfs = CopySupportingPdfs(GetSupportingPdfPaths(analysis.RenewalCase), analysis.RenewalCase.OutputFolder);
            var copiedSupportingWorkbooks = CopySupportingWorkbooks(GetSupportingWorkbookPaths(analysis.RenewalCase), analysis.RenewalCase.OutputFolder);
            analysis.SourceChecks.RemoveAll(check => check.Area.Equals("Supporting PDFs", StringComparison.OrdinalIgnoreCase) &&
                                                     check.Status.Equals("Copied", StringComparison.OrdinalIgnoreCase));
            analysis.SourceChecks.RemoveAll(check => check.Area.Equals("Supporting Excel files", StringComparison.OrdinalIgnoreCase) &&
                                                     check.Status.Equals("Copied", StringComparison.OrdinalIgnoreCase));
            if (copiedSupportingPdfs.Count > 0)
            {
                analysis.SourceChecks.Add(new SourceCheck
                {
                    Area = "Supporting PDFs",
                    Status = "Copied",
                    Detail = $"{copiedSupportingPdfs.Count} copied to {Path.Combine(analysis.RenewalCase.OutputFolder, "Supporting PDFs")}"
                });
            }

            if (copiedSupportingWorkbooks.Count > 0)
            {
                analysis.SourceChecks.Add(new SourceCheck
                {
                    Area = "Supporting Excel files",
                    Status = "Copied",
                    Detail = $"{copiedSupportingWorkbooks.Count} copied to {Path.Combine(analysis.RenewalCase.OutputFolder, "Supporting Excel Files")}"
                });
            }

            progress?.Report("Writing Excel review workbook...");
            var underwritingRates = _databaseService is null
                ? UnderwritingReferenceRateCatalog.Defaults
                : await _databaseService.GetActiveUnderwritingReferenceRatesAsync();
            analysis.YearsAsClient ??= await ResolveYearsAsClientAsync(analysis);
            var workbookPath = new ExcelReportWriter(underwritingRates).Write(analysis, analysis.RenewalCase.OutputFolder);

            progress?.Report("Writing PowerPoint draft...");
            var deck = _deckWriter.Write(
                analysis,
                analysis.RenewalCase.OutputFolder,
                analysis.RenewalCase.PowerPointTemplatePath,
                underwritingRates);

            progress?.Report("Export complete.");
            return new ExportResult
            {
                OutputFolder = analysis.RenewalCase.OutputFolder,
                WorkbookPath = workbookPath,
                PowerPointPath = deck.Path,
                DeckMessage = deck.Message,
                SupportingPdfFolder = copiedSupportingPdfs.Count > 0 ? Path.Combine(analysis.RenewalCase.OutputFolder, "Supporting PDFs") : null,
                SupportingPdfCount = copiedSupportingPdfs.Count,
                SupportingWorkbookFolder = copiedSupportingWorkbooks.Count > 0 ? Path.Combine(analysis.RenewalCase.OutputFolder, "Supporting Excel Files") : null,
                SupportingWorkbookCount = copiedSupportingWorkbooks.Count
            };
        });
    }

    private async Task<int?> ResolveYearsAsClientAsync(AnalysisResult analysis)
    {
        if (_databaseService is null || analysis.DatabaseClientId is not > 0)
        {
            return analysis.YearsAsClient;
        }

        try
        {
            var events = await ClientHistoryProjectionBuilder
                .CreateDefault(_databaseService)
                .BuildAsync(analysis.DatabaseClientId.Value);
            return new ClientHistoryTimelineService()
                .CalculateKpis(events, DateOnly.FromDateTime(DateTime.Today))
                .YearsAsClient;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return analysis.YearsAsClient;
        }
    }

    private static void ValidateInputs(RenewalCase renewalCase, AnalysisResult result)
    {
        AddFileCheck(result, "Renewal PDF", renewalCase.PdfPath);
        AddFileCheck(result, WorkbookSourceArea, renewalCase.WorkbookPath);
        AddFolderCheck(result, "Current Year Folder", renewalCase.CurrentYearFolder);
        AddFolderCheck(result, "Prior Year Folder", renewalCase.PriorYearFolder);
        AddFolderCheck(result, "ASO Billing Folder", renewalCase.AsoBillingFolder);

        var additionalPdfs = renewalCase.AdditionalPdfPaths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        for (var index = 0; index < additionalPdfs.Count; index++)
        {
            AddFileCheck(result, $"Additional PDF {index + 1}", additionalPdfs[index], false);
        }

        var additionalWorkbooks = renewalCase.AdditionalWorkbookPaths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        for (var index = 0; index < additionalWorkbooks.Count; index++)
        {
            AddFileCheck(result, $"Additional Excel {index + 1}", additionalWorkbooks[index], false);
        }

        if (!string.IsNullOrWhiteSpace(renewalCase.PowerPointTemplatePath))
        {
            AddFileCheck(result, "PowerPoint template", renewalCase.PowerPointTemplatePath, false);
        }

        if (string.IsNullOrWhiteSpace(renewalCase.OutputFolder))
        {
            result.Findings.Add(new Finding
            {
                Severity = FindingSeverity.Critical,
                Title = "Output folder is required",
                Evidence = "No output folder was selected.",
                RecommendedAction = "Choose a folder where the review pack can be saved.",
                ClientExplanation = "The app needs a place to save the review materials.",
                SourceReference = "Output folder"
            });
        }
    }

    private void ReadNemMaximums(RenewalCase renewalCase, AnalysisResult result)
    {
        var previousRenewalPaths = result.PreviousRenewals
            .Select(snapshot => snapshot.SourcePath)
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var paths = CurrentNemSourcePdfPaths(renewalCase)
            .Where(path => !previousRenewalPaths.Any(previousPath => PathsEqual(previousPath, path)))
            .ToList();

        if (paths.Count == 0)
        {
            return;
        }

        var rows = _nemMaximumExtractor.ExtractFromFiles(paths);
        result.NemMaximums.Clear();
        result.NemMaximums.AddRange(rows);
        if (rows.Count > 0)
        {
            result.SourceChecks.Add(new SourceCheck
            {
                Area = "Plan Maximums and NEMs",
                Status = "Read",
                Detail = $"{rows.Count} Life/LTD maximum row(s) found in current plan-design PDF text."
            });
        }
    }

    private static void AddFileCheck(AnalysisResult result, string area, string path, bool critical = true)
    {
        if (File.Exists(path))
        {
            result.SourceChecks.Add(new SourceCheck { Area = area, Status = "Found", Detail = path });
            return;
        }

        result.SourceChecks.Add(new SourceCheck { Area = area, Status = "Missing", Detail = path });
        result.Findings.Add(new Finding
        {
            Severity = critical ? FindingSeverity.Critical : FindingSeverity.Warning,
            Title = $"{area} is missing",
            Evidence = string.IsNullOrWhiteSpace(path) ? "No file was selected." : path,
            RecommendedAction = $"Select the correct {area} file.",
            ClientExplanation = $"{area} was not available for this review.",
            SourceReference = area
        });
    }

    private static void AddFolderCheck(AnalysisResult result, string area, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        if (Directory.Exists(path))
        {
            result.SourceChecks.Add(new SourceCheck { Area = area, Status = "Found", Detail = path });
            return;
        }

        result.SourceChecks.Add(new SourceCheck { Area = area, Status = "Missing", Detail = path });
        result.Findings.Add(new Finding
        {
            Severity = FindingSeverity.Warning,
            Title = $"{area} is missing",
            Evidence = path,
            RecommendedAction = $"Select the correct {area}.",
            ClientExplanation = $"{area} was not available for this review.",
            SourceReference = area
        });
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private void ReadAsoBillingActivityReports(RenewalCase renewalCase, AnalysisResult result)
    {
        if (string.IsNullOrWhiteSpace(renewalCase.AsoBillingFolder))
        {
            return;
        }

        if (!Directory.Exists(renewalCase.AsoBillingFolder))
        {
            return;
        }

        var groupHealthExtraction = _asoMonthlyActivityPdfParser.ExtractFolder(renewalCase.AsoBillingFolder);
        var sunLifeExtraction = _sunLifeAsoInvoicePdfParser.ExtractFolder(renewalCase.AsoBillingFolder);
        var acceptedSunLifePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var statement in sunLifeExtraction.FinancialStatements)
        {
            var expectedClientNames = new[] { renewalCase.ClientName, result.Pdf.ClientName }
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var identityMatches = expectedClientNames.Count > 0 &&
                                  expectedClientNames.All(expected =>
                                      BookletClientIdentityGuardrail.ClientNamesLikelyMatch(expected, statement.ClientName));
            var carrierMatches = result.Pdf.Carrier.Contains("Sun Life", StringComparison.OrdinalIgnoreCase) &&
                                 statement.Carrier.Equals("Sun Life", StringComparison.OrdinalIgnoreCase);
            if (!identityMatches || !carrierMatches)
            {
                sunLifeExtraction.Warnings.Add(
                    $"{statement.SourceFileName} is a Sun Life ASO invoice, but its carrier/client identity does not match the selected renewal. Nothing was staged for persistence.");
                continue;
            }

            acceptedSunLifePaths.Add(statement.SourceFilePath);
            result.AsoFinancialActivityStatements.Add(statement);
        }

        var extractedRecords = groupHealthExtraction.Records
            .Concat(sunLifeExtraction.Records.Where(record => acceptedSunLifePaths.Contains(Path.GetFullPath(record.SourcePath))))
            .GroupBy(record => record.ServiceMonth)
            .Select(group => group
                .OrderByDescending(record => record.InterestCreditCharge.HasValue)
                .ThenByDescending(record => record.SourcePath, StringComparer.OrdinalIgnoreCase)
                .First())
            .OrderBy(record => record.ServiceMonth)
            .ToList();
        result.AsoBillingActivity.Clear();
        result.AsoBillingActivity.AddRange(extractedRecords);

        foreach (var warning in groupHealthExtraction.Warnings.Concat(sunLifeExtraction.Warnings))
        {
            result.SourceChecks.Add(new SourceCheck
            {
                Area = "ASO billing activity",
                Status = "Skipped",
                Detail = warning
            });
        }

        result.SourceChecks.Add(new SourceCheck
        {
            Area = "ASO billing activity",
            Status = extractedRecords.Count > 0 ? "Read" : "Reviewed",
            Detail = extractedRecords.Count > 0
                ? $"{extractedRecords.Count} monthly dental ASO billing statement(s) read from {renewalCase.AsoBillingFolder}; {acceptedSunLifePaths.Count} Sun Life invoice(s) staged for persistence."
                : $"No readable GroupHEALTH Summary of Monthly Activity or Sun Life ASO Invoice reports were found in {renewalCase.AsoBillingFolder}."
        });
    }

    private void ReadOperatingAccounts(RenewalCase renewalCase, AnalysisResult result)
    {
        var candidatePaths = new[] { renewalCase.PdfPath }
            .Concat(renewalCase.AdditionalPdfPaths)
            .Where(path => !string.IsNullOrWhiteSpace(path) &&
                           path.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) &&
                           File.Exists(path))
            .Where(path => !IsGeneratedOutputPath(path))
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var sourceHashes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var candidatePath in candidatePaths)
        {
            try
            {
                var sourceHash = ComputeSha256(candidatePath);
                if (!sourceHashes.Add(sourceHash))
                {
                    result.SourceChecks.Add(new SourceCheck
                    {
                        Area = "ASO operating accounts",
                        Status = "Skipped",
                        Detail = $"{Path.GetFileName(candidatePath)} duplicates an already scanned operating-account PDF."
                    });
                    continue;
                }

                var statement = _sunLifeOperatingAccountsPdfParser.Extract(candidatePath);
                if (statement is null)
                {
                    continue;
                }

                statement.SourceHash = sourceHash;
                var expectedClientNames = new[]
                    {
                        renewalCase.ClientName,
                        result.Pdf.ClientName
                    }
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var identityMatches = expectedClientNames.Count > 0 &&
                                      expectedClientNames.All(expected =>
                                          BookletClientIdentityGuardrail.ClientNamesLikelyMatch(
                                              expected,
                                              statement.ClientName));
                var primaryCarrierIsSunLife = result.Pdf.Carrier.Contains(
                    "Sun Life",
                    StringComparison.OrdinalIgnoreCase);
                if (!primaryCarrierIsSunLife ||
                    !statement.Carrier.Equals("Sun Life", StringComparison.OrdinalIgnoreCase) ||
                    string.IsNullOrWhiteSpace(statement.ClientName) ||
                    string.IsNullOrWhiteSpace(statement.ContractNumber) ||
                    !identityMatches)
                {
                    result.SourceChecks.Add(new SourceCheck
                    {
                        Area = "ASO operating accounts",
                        Status = "Needs Review",
                        Detail = $"{Path.GetFileName(candidatePath)} contains operating accounts, but its Sun Life/client identity does not match the selected renewal. Nothing was staged for persistence."
                    });
                    continue;
                }

                result.AsoFinancialActivityStatements.Add(statement);
                var sourceRole = PathsEqual(candidatePath, renewalCase.PdfPath)
                    ? "primary renewal PDF"
                    : "additional PDF";
                result.SourceChecks.Add(new SourceCheck
                {
                    Area = "ASO operating accounts",
                    Status = statement.ImportStatus == AsoFinancialImportStatuses.Complete ? "Read" : "Needs Review",
                    Detail = statement.ImportStatus == AsoFinancialImportStatuses.Complete
                        ? $"{Path.GetFileName(candidatePath)} ({sourceRole}) | Contract {statement.ContractNumber}: {statement.Periods.Count} monthly operating-account rows read from {statement.SourceSheet}."
                        : $"{Path.GetFileName(candidatePath)} ({sourceRole}) | Contract {statement.ContractNumber}: operating accounts were found but require review. {statement.ValidationSummary}"
                });
            }
            catch (Exception exception)
            {
                result.SourceChecks.Add(new SourceCheck
                {
                    Area = "ASO operating accounts",
                    Status = "Needs Review",
                    Detail = $"{Path.GetFileName(candidatePath)} contains an unreadable operating-account section: {exception.Message}"
                });
            }
        }
    }

    private static string ComputeSha256(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private void ReadAdditionalRenewalPdfs(RenewalCase renewalCase, AnalysisResult result)
    {
        var primaryRenewalFolder = Path.GetDirectoryName(renewalCase.PdfPath) ?? "";
        var candidatePaths = renewalCase.AdditionalPdfPaths
            .Concat(EnumerateFilesFromFolder(primaryRenewalFolder, "*.pdf")
                .Where(IsLikelyAllstateStandalonePolicy))
            .Concat(EnumerateFilesFromFolder(renewalCase.PriorYearFolder, "*.pdf"))
            .Concat(EnumerateFilesFromFolder(renewalCase.CurrentYearFolder, "*.pdf")
                .Where(IsLikelyAllstateStandalonePolicy))
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Where(path => !PathsEqual(path, renewalCase.PdfPath))
            .Where(path => !IsGeneratedOutputPath(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (candidatePaths.Count == 0)
        {
            return;
        }

        var foundCount = 0;
        var supportingCount = 0;
        foreach (var candidatePath in candidatePaths)
        {
            if (IsDiscoveredClaimsRange(result, candidatePath))
            {
                continue;
            }

            try
            {
                var isPriorFolderSource = IsFileInFolder(candidatePath, renewalCase.PriorYearFolder);
                if (isPriorFolderSource && TryMergePriorFolderPdf(candidatePath, result, out var priorDetail))
                {
                    supportingCount++;
                    result.SourceChecks.Add(new SourceCheck
                    {
                        Area = "Prior year support PDF",
                        Status = priorDetail.StartsWith("Skipped:", StringComparison.OrdinalIgnoreCase) ? "Skipped" : "Read",
                        Detail = $"{Path.GetFileName(candidatePath)}: {priorDetail}"
                    });
                    continue;
                }

                if (!isPriorFolderSource && TryMergeStandaloneBenefitPolicy(candidatePath, result, out var standaloneDetail))
                {
                    supportingCount++;
                    result.SourceChecks.Add(new SourceCheck
                    {
                        Area = "Stand-alone benefit policy",
                        Status = standaloneDetail.StartsWith("Skipped:", StringComparison.OrdinalIgnoreCase)
                            ? "Skipped"
                            : standaloneDetail.StartsWith("Needs review:", StringComparison.OrdinalIgnoreCase)
                                ? "Needs review"
                                : "Read",
                        Detail = $"{Path.GetFileName(candidatePath)}: {standaloneDetail}"
                    });
                    continue;
                }

                if (!isPriorFolderSource && TryMergeCarrierSupportingPdf(candidatePath, result, out var supportingDetail))
                {
                    supportingCount++;
                    result.SourceChecks.Add(new SourceCheck
                    {
                        Area = $"{result.Pdf.Carrier} support PDF",
                        Status = supportingDetail.StartsWith("Skipped:", StringComparison.OrdinalIgnoreCase) ? "Skipped" : "Read",
                        Detail = $"{Path.GetFileName(candidatePath)}: {supportingDetail}"
                    });
                    continue;
                }

                PreviousCarrierRenewalExtraction? carrierExtraction = null;
                var snapshot = _previousRenewalParser.Parse(candidatePath);
                if (snapshot is null)
                {
                    carrierExtraction = TryReadCarrierPreviousRenewal(candidatePath);
                    snapshot = carrierExtraction?.Snapshot;
                }

                if (snapshot is null || !IsPriorRenewalSnapshot(snapshot, result))
                {
                    continue;
                }

                var mergeDetail = carrierExtraction is null
                    ? ""
                    : MergePreviousCarrierPdfData(result, carrierExtraction.Pdf, carrierExtraction.Workbook, snapshot);

                result.PreviousRenewals.Add(snapshot);
                foundCount++;
                result.SourceChecks.Add(new SourceCheck
                {
                    Area = "Previous renewal PDF",
                    Status = "Read",
                    Detail = $"{snapshot.SourceFileName}: {DescribePreviousRenewal(snapshot)}{mergeDetail}"
                });
            }
            catch (Exception ex)
            {
                result.SourceChecks.Add(new SourceCheck
                {
                    Area = "Previous renewal PDF",
                    Status = "Skipped",
                    Detail = $"{Path.GetFileName(candidatePath)} could not be scanned for prior-renewal values: {ex.Message}"
                });
            }
        }

        if (foundCount == 0 && supportingCount == 0)
        {
            result.SourceChecks.Add(new SourceCheck
            {
                Area = "Previous renewal PDF scan",
                Status = "Reviewed",
                Detail = $"No previous-renewal values were detected in {candidatePaths.Count} selected supporting PDF(s)."
            });
        }
    }

    private bool TryMergeStandaloneBenefitPolicy(string candidatePath, AnalysisResult result, out string detail)
    {
        detail = "";
        CarrierPdfParseResult parsed;
        try
        {
            parsed = _pdfParser.Parse(candidatePath);
        }
        catch
        {
            return false;
        }

        if (!parsed.Carrier.Equals(AllstateCriticalIllnessPdfParser.CarrierName, StringComparison.OrdinalIgnoreCase) ||
            parsed.Pdf.RateLines.Count == 0)
        {
            return false;
        }

        var primaryClient = NormalizeStandaloneClientKey(result.Pdf.ClientName);
        var supportingClient = NormalizeStandaloneClientKey(parsed.Pdf.ClientName);
        if (string.IsNullOrWhiteSpace(primaryClient) ||
            string.IsNullOrWhiteSpace(supportingClient) ||
            !primaryClient.Equals(supportingClient, StringComparison.OrdinalIgnoreCase))
        {
            detail = "Skipped: the Allstate client identity does not match the primary renewal.";
            return true;
        }

        var sourceFileName = Path.GetFileName(candidatePath);
        if (result.Pdf.RateLines.Any(line =>
                line.SourceCarrier.Equals(AllstateCriticalIllnessPdfParser.CarrierName, StringComparison.OrdinalIgnoreCase) &&
                line.SourceFileName.Equals(sourceFileName, StringComparison.OrdinalIgnoreCase)))
        {
            detail = "Skipped: this Allstate Critical Illness policy is already present.";
            return true;
        }

        result.Pdf.RateLines.AddRange(parsed.Pdf.RateLines);
        foreach (var warning in parsed.Pdf.Warnings.Where(warning =>
                     !result.Pdf.Warnings.Contains(warning, StringComparer.OrdinalIgnoreCase)))
        {
            result.Pdf.Warnings.Add(warning);
        }

        detail = $"Needs review: {parsed.Pdf.RateLines.Count} Allstate Critical Illness tier rate(s) added. Rates are monthly per enrolled employee; enter EE and Family enrollment counts before including them in consolidated premiums.";
        return true;
    }

    private static bool IsLikelyAllstateStandalonePolicy(string path)
    {
        var fileName = Path.GetFileName(path);
        return fileName.Contains("Allstate", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("AICC", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeStandaloneClientKey(string value) =>
        new(ParsingHelpers.Clean(value)
            .Where(char.IsLetterOrDigit)
            .Select(char.ToUpperInvariant)
            .ToArray());

    private static bool IsDiscoveredClaimsRange(AnalysisResult result, string path)
    {
        return Path.GetFileNameWithoutExtension(path)
                   .Contains("claims range", StringComparison.OrdinalIgnoreCase) ||
               result.DocumentDiscovery?.Documents.Any(document =>
                   PathsEqual(document.FilePath, path) &&
                   document.DocumentType.Equals(
                       DiscoveryDocumentTypes.GroupHealthClaimsRange,
                       StringComparison.OrdinalIgnoreCase)) == true;
    }

    private RenewalDocumentDiscoveryResult DiscoverInputDocuments(RenewalCase renewalCase)
    {
        var discovery = new RenewalDocumentDiscoveryResult
        {
            ClientFolder = FirstNonBlank(
                renewalCase.CurrentYearFolder,
                renewalCase.PriorYearFolder,
                Path.GetDirectoryName(renewalCase.PdfPath) ?? "")
        };

        foreach (var path in GetInputPdfPaths(renewalCase))
        {
            try
            {
                discovery.Documents.Add(_documentDiscoveryService.ClassifyDocument(path));
            }
            catch
            {
                var document = DocumentDiscoveryService.ClassifyFileName(path);
                document.Status = "classification_failed";
                discovery.Documents.Add(document);
            }
        }

        discovery.RenewalDocument = SelectDiscoveryRenewalDocument(discovery.Documents);
        discovery.PlanDesign = DocumentDiscoveryService.SelectPlanDesignSource(discovery.Documents, discovery.RenewalDocument);
        return discovery;
    }

    private static void AddDocumentDiscoverySourceChecks(AnalysisResult result)
    {
        var discovery = result.DocumentDiscovery;
        if (discovery is null || discovery.Documents.Count == 0)
        {
            return;
        }

        var sunLifeDocuments = discovery.Documents
            .Where(document => document.Carrier.Equals("Sun Life", StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (sunLifeDocuments.Count == 0)
        {
            return;
        }

        var hasRenewal = sunLifeDocuments.Any(document =>
            document.DocumentType.Equals(DiscoveryDocumentTypes.SunLifeRenewalReport, StringComparison.OrdinalIgnoreCase));
        var hasExperience = sunLifeDocuments.Any(document =>
            document.DocumentType.Equals(DiscoveryDocumentTypes.SunLifeExperienceAnalysis, StringComparison.OrdinalIgnoreCase));

        result.SourceChecks.Add(new SourceCheck
        {
            Area = "Document discovery",
            Status = "Read",
            Detail = $"{sunLifeDocuments.Count} Sun Life document(s) classified across the selected inputs."
        });

        if (!hasRenewal || !hasExperience)
        {
            result.SourceChecks.Add(new SourceCheck
            {
                Area = "Sun Life companion document",
                Status = "Warning",
                Detail = !hasRenewal
                    ? "Sun Life experience analysis was found, but the renewal report was not found in the selected inputs."
                    : "Sun Life renewal report was found, but the separate experience analysis was not found in the selected inputs."
            });
        }
    }

    private bool TryMergePriorFolderPdf(string candidatePath, AnalysisResult result, out string detail)
    {
        detail = "";
        CarrierPdfParseResult parsed;
        try
        {
            parsed = _pdfParser.Parse(candidatePath);
        }
        catch
        {
            return false;
        }

        if (IsIdentityCheckedTopDrugSupport(parsed) &&
            !CanMergeSupportingTopDrugReport(result.Pdf, parsed.Pdf, out var identityDetail))
        {
            result.Pdf.Warnings.Add($"{parsed.Carrier} Top 10 drug rows from {Path.GetFileName(candidatePath)} were not merged: {identityDetail}");
            detail = $"Skipped: {identityDetail}";
            return true;
        }

        var experienceRowsAdded = MergePreviousExperienceRows(result.Pdf, parsed.Pdf);
        var sunLifeRowsAdded = MergeSunLifeExperienceRows(result.Pdf, parsed.Pdf);
        var assurePlusReportsAdded = MergeAssurePlusPricingReports(result.Pdf, parsed.Pdf);
        var priorDrugCountBefore = result.Workbook.PriorDrugTrends.Count;
        var workbookRowsAdded = MergePriorWorkbookData(result.Workbook, parsed.WorkbookFromPdf);
        var priorDrugRowsAdded = result.Workbook.PriorDrugTrends.Count - priorDrugCountBefore;

        var snapshotAdded = false;
        var snapshot = PreviousRenewalReportParser.FromCarrierExtraction(candidatePath, parsed.Pdf);
        if (snapshot is not null &&
            IsPriorRenewalSnapshot(snapshot, result) &&
            !result.PreviousRenewals.Any(existing => PathsEqual(existing.SourcePath, candidatePath)))
        {
            result.PreviousRenewals.Add(snapshot);
            snapshotAdded = true;
        }

        if (experienceRowsAdded == 0 && sunLifeRowsAdded == 0 && assurePlusReportsAdded == 0 &&
            workbookRowsAdded == 0 && !snapshotAdded)
        {
            return false;
        }

        if (experienceRowsAdded > 0 || sunLifeRowsAdded > 0 || workbookRowsAdded > 0)
        {
            RemoveMissingWorkbookChecks(result);
        }

        var parts = new List<string>();
        if (experienceRowsAdded > 0)
        {
            parts.Add($"{experienceRowsAdded} prior-period premium/claims rows merged");
        }

        if (sunLifeRowsAdded > 0)
        {
            parts.Add($"{sunLifeRowsAdded} prior Sun Life experience rows captured");
        }

        if (assurePlusReportsAdded > 0)
        {
            parts.Add($"{assurePlusReportsAdded} prior AssurePLUS pricing report captured");
        }

        var nonDrugWorkbookRows = workbookRowsAdded - priorDrugRowsAdded;
        if (nonDrugWorkbookRows > 0)
        {
            parts.Add($"{nonDrugWorkbookRows} historical claims/detail rows merged");
        }

        if (priorDrugRowsAdded > 0)
        {
            parts.Add($"{priorDrugRowsAdded} prior Top DIN rows merged");
        }

        if (snapshotAdded)
        {
            parts.Add("previous renewal summary captured");
        }

        detail = $"{FirstNonBlank(parsed.Carrier, "Carrier")}: {string.Join("; ", parts)}";
        return true;
    }

    private bool TryMergeCarrierSupportingPdf(string candidatePath, AnalysisResult result, out string detail)
    {
        detail = "";
        if (!result.Pdf.Carrier.Equals("Equitable", StringComparison.OrdinalIgnoreCase) &&
            !result.Pdf.Carrier.Equals("Sun Life", StringComparison.OrdinalIgnoreCase) &&
            !result.Pdf.Carrier.Equals("Empire Life", StringComparison.OrdinalIgnoreCase) &&
            !result.Pdf.Carrier.Equals("GroupHEALTH", StringComparison.OrdinalIgnoreCase) &&
            !result.Pdf.Carrier.Equals("GroupSource", StringComparison.OrdinalIgnoreCase) &&
            !result.Pdf.Carrier.Equals("Canada Life", StringComparison.OrdinalIgnoreCase) &&
            !result.Pdf.Carrier.Equals(RwamRateSheetPdfParser.CarrierName, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        CarrierPdfParseResult parsed;
        try
        {
            if (result.Pdf.Carrier.Equals("GroupHEALTH", StringComparison.OrdinalIgnoreCase))
            {
                var pages = new PdfTextExtractor().Extract(candidatePath);
                var allText = ParsingHelpers.Clean(string.Join(" ", pages.Select(page => page.Text)));
                parsed = new GroupHealthPdfParser().Parse(candidatePath, pages, allText);
            }
            else
            {
                parsed = _pdfParser.Parse(candidatePath);
            }
        }
        catch
        {
            return false;
        }

        if (!parsed.Carrier.Equals(result.Pdf.Carrier, StringComparison.OrdinalIgnoreCase) ||
            !IsCarrierSupportingPdf(candidatePath, parsed))
        {
            return false;
        }

        if (parsed.Carrier.Equals(RwamRateSheetPdfParser.CarrierName, StringComparison.OrdinalIgnoreCase))
        {
            if (!CanMergeRwamRateSheet(result.Pdf, parsed.Pdf, out var rwamIdentityDetail))
            {
                result.Pdf.Warnings.Add($"RWAM companion rate sheet {Path.GetFileName(candidatePath)} was not merged: {rwamIdentityDetail}.");
                detail = $"Skipped: {rwamIdentityDetail}";
                return true;
            }

            var added = MergeRwamRateSheet(result.Pdf, parsed.Pdf);
            detail = added == 0
                ? "Skipped: this RWAM class/division rate sheet was already present"
                : $"{added} RWAM class/division rate rows merged";
            return true;
        }

        if (parsed.Carrier.Equals("Empire Life", StringComparison.OrdinalIgnoreCase) &&
            !CanMergeEmpireLifeSupport(result.Pdf, parsed.Pdf, out var empireIdentityDetail))
        {
            result.Pdf.Warnings.Add(
                $"Empire Life support report {Path.GetFileName(candidatePath)} was not merged: {empireIdentityDetail}.");
            detail = $"Skipped: {empireIdentityDetail}";
            return true;
        }

        if (IsIdentityCheckedTopDrugSupport(parsed) &&
            !CanMergeSupportingTopDrugReport(result.Pdf, parsed.Pdf, out var identityDetail))
        {
            result.Pdf.Warnings.Add($"{parsed.Carrier} Top 10 drug rows from {Path.GetFileName(candidatePath)} were not merged: {identityDetail}");
            detail = $"Skipped: {identityDetail}";
            return true;
        }

        var experienceRowsAdded = MergeCurrentExperienceRows(result.Pdf, parsed.Pdf);
        var sunLifeRowsAdded = MergeSunLifeExperienceRows(result.Pdf, parsed.Pdf);
        var empireLifeRowsAdded = MergeEmpireLifeExperienceRows(result.Pdf, parsed.Pdf);
        var demographicRowsAdded = MergeDemographicRows(result.Pdf.LifeDemographics, parsed.Pdf.LifeDemographics) +
                                   MergeDemographicRows(result.Pdf.LtdDemographics, parsed.Pdf.LtdDemographics);
        var assurePlusReportsAdded = MergeAssurePlusPricingReports(result.Pdf, parsed.Pdf);
        var categoryRowsUpdated = MergeCurrentWorkbookData(result.Workbook, parsed.WorkbookFromPdf);
        result.Pdf.PoolingThreshold ??= parsed.Pdf.PoolingThreshold;
        result.Pdf.PooledPremium ??= parsed.Pdf.PooledPremium;
        result.Pdf.PoolingCharge ??= parsed.Pdf.PoolingCharge;
        if (sunLifeRowsAdded > 0)
        {
            SunLifePdfParser.ApplySunLifePoolingAssumptions(result.Pdf);
        }

        if (experienceRowsAdded == 0 && sunLifeRowsAdded == 0 && empireLifeRowsAdded == 0 &&
            demographicRowsAdded == 0 && assurePlusReportsAdded == 0 && categoryRowsUpdated == 0)
        {
            detail = "reviewed as a carrier support PDF; no renewal-review rows were required from this file";
            return true;
        }

        result.RenewalCase.ClientName = FirstNonBlank(result.RenewalCase.ClientName, parsed.Pdf.ClientName);
        result.RenewalCase.PolicyNumber = FirstNonBlank(result.RenewalCase.PolicyNumber, parsed.Pdf.PolicyNumber);
        RemoveMissingWorkbookChecks(result);

        var parts = new List<string>();
        if (experienceRowsAdded > 0)
        {
            parts.Add($"{experienceRowsAdded} premium/claims rows merged");
        }

        if (sunLifeRowsAdded > 0)
        {
            parts.Add($"{sunLifeRowsAdded} raw Sun Life experience rows captured");
        }

        if (empireLifeRowsAdded > 0)
        {
            parts.Add($"{empireLifeRowsAdded} raw Empire Life experience rows captured");
        }

        if (demographicRowsAdded > 0)
        {
            parts.Add($"{demographicRowsAdded} demographic rows merged");
        }

        if (assurePlusReportsAdded > 0)
        {
            parts.Add($"{assurePlusReportsAdded} AssurePLUS pricing report captured");
        }

        if (categoryRowsUpdated > 0)
        {
            parts.Add($"{categoryRowsUpdated} claims/detail rows merged");
        }

        detail = string.Join("; ", parts);
        return true;
    }

    private static bool IsCarrierSupportingPdf(string candidatePath, CarrierPdfParseResult parsed)
    {
        var fileName = Path.GetFileName(candidatePath);
        if (parsed.Carrier.Equals(RwamRateSheetPdfParser.CarrierName, StringComparison.OrdinalIgnoreCase))
        {
            return parsed.Pdf.RateLines.Count > 0;
        }

        if (parsed.Carrier.Equals("Canada Life", StringComparison.OrdinalIgnoreCase))
        {
            return IsIdentityCheckedTopDrugSupport(parsed);
        }

        if (parsed.Carrier.Equals("Sun Life", StringComparison.OrdinalIgnoreCase))
        {
            var hasExperienceRows = parsed.Pdf.ExperienceLines.Count > 0 || parsed.Pdf.SunLifeExperienceRows.Count > 0;
            var hasWorkbookRows = parsed.WorkbookFromPdf is not null &&
                                  (parsed.WorkbookFromPdf.ClaimsDistributionRows.Count > 0 ||
                                   parsed.WorkbookFromPdf.EhcSpendingTrends.Count > 0 ||
                                   parsed.WorkbookFromPdf.DentalSpendingTrends.Count > 0 ||
                                   parsed.WorkbookFromPdf.ParamedicalServiceTrends.Count > 0);
            var looksLikeExperienceFile = fileName.Contains("EXP", StringComparison.OrdinalIgnoreCase) ||
                                          fileName.Contains("Experience", StringComparison.OrdinalIgnoreCase) ||
                                          fileName.Contains("Claims", StringComparison.OrdinalIgnoreCase) ||
                                          fileName.Contains("Dental", StringComparison.OrdinalIgnoreCase) ||
                                          fileName.Contains("EHC", StringComparison.OrdinalIgnoreCase);
            return (hasExperienceRows || hasWorkbookRows) && (looksLikeExperienceFile || parsed.Pdf.RateLines.Count == 0);
        }

        if (parsed.Carrier.Equals("Empire Life", StringComparison.OrdinalIgnoreCase))
        {
            var hasWorkbookRows = parsed.WorkbookFromPdf is not null &&
                                  (parsed.WorkbookFromPdf.ClaimsDistributionRows.Count > 0 ||
                                   parsed.WorkbookFromPdf.EhcSpendingTrends.Count > 0 ||
                                   parsed.WorkbookFromPdf.DentalSpendingTrends.Count > 0 ||
                                   parsed.WorkbookFromPdf.ParamedicalServiceTrends.Count > 0);
            return parsed.Pdf.ExperienceLines.Count > 0 ||
                   parsed.Pdf.LifeDemographics.Count > 0 ||
                   parsed.Pdf.LtdDemographics.Count > 0 ||
                   hasWorkbookRows;
        }

        if (parsed.Carrier.Equals("GroupHEALTH", StringComparison.OrdinalIgnoreCase) ||
            parsed.Carrier.Equals("GroupSource", StringComparison.OrdinalIgnoreCase))
        {
            var hasWorkbookRows = parsed.WorkbookFromPdf is not null &&
                                  (parsed.WorkbookFromPdf.EhcSpendingTrends.Count > 0 ||
                                   parsed.WorkbookFromPdf.DentalSpendingTrends.Count > 0 ||
                                   parsed.WorkbookFromPdf.ParamedicalServiceTrends.Count > 0 ||
                                   parsed.WorkbookFromPdf.DrugTrends.Count > 0);
            return parsed.Pdf.ExperienceLines.Count > 0 ||
                   parsed.Pdf.AssurePlusPricingReports.Count > 0 ||
                   hasWorkbookRows ||
                   parsed.Pdf.PoolingThreshold.HasValue;
        }

        if (!parsed.Carrier.Equals("Equitable", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (fileName.StartsWith("RenewalStProc_EL_", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return parsed.WorkbookFromPdf?.DrugTrends.Count > 0 &&
               parsed.Pdf.RateLines.Count == 0 &&
               parsed.Pdf.ExperienceLines.Count == 0;
    }

    internal static bool CanMergeRwamRateSheet(
        PdfExtractionResult primary,
        PdfExtractionResult companion,
        out string detail)
    {
        if (!primary.Carrier.Equals(RwamRateSheetPdfParser.CarrierName, StringComparison.OrdinalIgnoreCase) ||
            !companion.Carrier.Equals(RwamRateSheetPdfParser.CarrierName, StringComparison.OrdinalIgnoreCase))
        {
            detail = "both documents must be classified as RWAM rate sheets";
            return false;
        }

        if (string.IsNullOrWhiteSpace(primary.PolicyNumber) ||
            string.IsNullOrWhiteSpace(companion.PolicyNumber) ||
            !NormalizePolicyMergeKey(primary.PolicyNumber).Equals(
                NormalizePolicyMergeKey(companion.PolicyNumber),
                StringComparison.OrdinalIgnoreCase))
        {
            detail = "group identifiers do not match";
            return false;
        }

        if (string.IsNullOrWhiteSpace(primary.ClientName) ||
            string.IsNullOrWhiteSpace(companion.ClientName) ||
            !NormalizeMergeKey(primary.ClientName).Equals(
                NormalizeMergeKey(companion.ClientName),
                StringComparison.OrdinalIgnoreCase))
        {
            detail = "client identities do not match";
            return false;
        }

        if (!primary.RenewalDate.HasValue ||
            !companion.RenewalDate.HasValue ||
            primary.RenewalDate.Value != companion.RenewalDate.Value)
        {
            detail = "renewal effective dates do not match";
            return false;
        }

        var incomingClass = companion.RateLines
            .Select(line => line.AccountClass)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
        var existingClassRows = primary.RateLines
            .Where(line => line.AccountClass.Equals(incomingClass, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (existingClassRows.Count > 0 && !RateSetsEqual(existingClassRows, companion.RateLines))
        {
            detail = $"class/division {incomingClass} is already present with different values";
            return false;
        }

        detail = "matched client, group and renewal effective date";
        return true;
    }

    private static int MergeRwamRateSheet(PdfExtractionResult primary, PdfExtractionResult companion)
    {
        var incomingClass = companion.RateLines
            .Select(line => line.AccountClass)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
        if (primary.RateLines.Any(line => line.AccountClass.Equals(incomingClass, StringComparison.OrdinalIgnoreCase)))
        {
            return 0;
        }

        primary.RateLines.AddRange(companion.RateLines);
        primary.CurrentMonthlyPremium = Add(primary.CurrentMonthlyPremium, companion.CurrentMonthlyPremium);
        primary.RenewalMonthlyPremium = Add(primary.RenewalMonthlyPremium, companion.RenewalMonthlyPremium);
        primary.CurrentEhcPremium = Add(primary.CurrentEhcPremium, companion.CurrentEhcPremium);
        primary.RenewalEhcPremium = Add(primary.RenewalEhcPremium, companion.RenewalEhcPremium);
        primary.CurrentDentalPremium = Add(primary.CurrentDentalPremium, companion.CurrentDentalPremium);
        primary.RenewalDentalPremium = Add(primary.RenewalDentalPremium, companion.RenewalDentalPremium);
        if (primary.CurrentMonthlyPremium.HasValue && primary.RenewalMonthlyPremium.HasValue)
        {
            primary.MonthlyPremiumChange = primary.RenewalMonthlyPremium.Value - primary.CurrentMonthlyPremium.Value;
            primary.MonthlyPremiumChangePercent = UnderwritingEngine.CalculatePercentChange(
                primary.CurrentMonthlyPremium,
                primary.RenewalMonthlyPremium);
        }

        foreach (var warning in companion.Warnings.Where(warning => !primary.Warnings.Contains(warning, StringComparer.OrdinalIgnoreCase)))
        {
            primary.Warnings.Add(warning);
        }

        return companion.RateLines.Count;
    }

    private static bool RateSetsEqual(IReadOnlyCollection<RateLine> left, IReadOnlyCollection<RateLine> right)
    {
        static string Key(RateLine line) => string.Join("|",
            line.Benefit,
            line.TierOrClass,
            line.Volume,
            line.Lives,
            line.CurrentRate,
            line.CurrentPremium,
            line.RenewalRate,
            line.RenewalPremium);

        return left.Select(Key).OrderBy(value => value, StringComparer.Ordinal).SequenceEqual(
            right.Select(Key).OrderBy(value => value, StringComparer.Ordinal),
            StringComparer.Ordinal);
    }

    private static decimal? Add(decimal? left, decimal? right)
    {
        return left.HasValue || right.HasValue ? (left ?? 0m) + (right ?? 0m) : null;
    }

    private static bool IsIdentityCheckedTopDrugSupport(CarrierPdfParseResult parsed)
    {
        return (parsed.Carrier.Equals("Canada Life", StringComparison.OrdinalIgnoreCase) ||
                parsed.Carrier.Equals("Equitable", StringComparison.OrdinalIgnoreCase) ||
                parsed.Carrier.Equals("GroupSource", StringComparison.OrdinalIgnoreCase)) &&
               parsed.WorkbookFromPdf?.DrugTrends.Count > 0 &&
               parsed.Pdf.RateLines.Count == 0 &&
               parsed.Pdf.ExperienceLines.Count == 0;
    }

    internal static bool CanMergeSupportingTopDrugReport(
        PdfExtractionResult primaryRenewal,
        PdfExtractionResult topDrugReport,
        out string detail)
    {
        var primaryPolicy = NormalizePolicyMergeKey(primaryRenewal.PolicyNumber);
        var reportPolicy = NormalizePolicyMergeKey(topDrugReport.PolicyNumber);
        if (string.IsNullOrWhiteSpace(primaryPolicy))
        {
            detail = "the primary renewal PDF has no policy number to verify";
            return false;
        }

        if (string.IsNullOrWhiteSpace(reportPolicy))
        {
            detail = "the Top 10 drug report has no policy number to verify";
            return false;
        }

        if (!primaryPolicy.Equals(reportPolicy, StringComparison.OrdinalIgnoreCase))
        {
            detail = $"policy {topDrugReport.PolicyNumber} does not match the primary renewal policy {primaryRenewal.PolicyNumber}";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(primaryRenewal.ClientName) &&
            !string.IsNullOrWhiteSpace(topDrugReport.ClientName) &&
            !BookletClientIdentityGuardrail.ClientNamesLikelyMatch(primaryRenewal.ClientName, topDrugReport.ClientName))
        {
            detail = $"client '{topDrugReport.ClientName}' does not match the primary renewal client '{primaryRenewal.ClientName}'";
            return false;
        }

        detail = string.IsNullOrWhiteSpace(primaryRenewal.ClientName)
            ? $"policy {primaryRenewal.PolicyNumber} matches the primary renewal"
            : $"policy {primaryRenewal.PolicyNumber} matches the primary renewal for {primaryRenewal.ClientName}";
        return true;
    }

    private static string NormalizePolicyMergeKey(string value)
    {
        return new string(value.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
    }

    private static bool CanMergeEmpireLifeSupport(
        PdfExtractionResult primaryRenewal,
        PdfExtractionResult support,
        out string detail)
    {
        var primaryPolicy = NormalizePolicyMergeKey(primaryRenewal.PolicyNumber);
        var supportPolicy = NormalizePolicyMergeKey(support.PolicyNumber);
        if (primaryPolicy.Length == 0 || supportPolicy.Length == 0)
        {
            detail = "the policy number required for identity verification is missing";
            return false;
        }

        if (!primaryPolicy.Equals(supportPolicy, StringComparison.OrdinalIgnoreCase))
        {
            detail = $"policy {support.PolicyNumber} does not match primary renewal policy {primaryRenewal.PolicyNumber}";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(primaryRenewal.ClientName) &&
            !string.IsNullOrWhiteSpace(support.ClientName) &&
            !BookletClientIdentityGuardrail.ClientNamesLikelyMatch(primaryRenewal.ClientName, support.ClientName))
        {
            detail = $"client '{support.ClientName}' does not match primary renewal client '{primaryRenewal.ClientName}'";
            return false;
        }

        detail = $"policy {primaryRenewal.PolicyNumber} matches";
        return true;
    }

    private static int MergeCurrentExperienceRows(PdfExtractionResult currentPdf, PdfExtractionResult supportingPdf)
    {
        return MergeExperienceRows(currentPdf, supportingPdf, markAsPrevious: false);
    }

    private static int MergeAssurePlusPricingReports(
        PdfExtractionResult currentPdf,
        PdfExtractionResult supportingPdf)
    {
        if (supportingPdf.AssurePlusPricingReports.Count == 0)
        {
            return 0;
        }

        var policyMismatch = !string.IsNullOrWhiteSpace(currentPdf.PolicyNumber) &&
                             !string.IsNullOrWhiteSpace(supportingPdf.PolicyNumber) &&
                             !NormalizeMergeKey(currentPdf.PolicyNumber).Equals(
                                 NormalizeMergeKey(supportingPdf.PolicyNumber),
                                 StringComparison.OrdinalIgnoreCase);
        var clientMismatch = !string.IsNullOrWhiteSpace(currentPdf.ClientName) &&
                             !string.IsNullOrWhiteSpace(supportingPdf.ClientName) &&
                             !NormalizeMergeKey(currentPdf.ClientName).Equals(
                                 NormalizeMergeKey(supportingPdf.ClientName),
                                 StringComparison.OrdinalIgnoreCase);
        if (policyMismatch || clientMismatch)
        {
            currentPdf.Warnings.Add(
                $"AssurePLUS report {Path.GetFileName(supportingPdf.SourcePath)} was not merged because its client or policy identity does not match the primary renewal PDF.");
            return 0;
        }

        var existingKeys = currentPdf.AssurePlusPricingReports
            .Select(AssurePlusReportKey)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var added = 0;
        foreach (var report in supportingPdf.AssurePlusPricingReports)
        {
            if (!existingKeys.Add(AssurePlusReportKey(report)))
            {
                continue;
            }

            currentPdf.AssurePlusPricingReports.Add(report);
            added++;
        }

        return added;
    }

    private static string AssurePlusReportKey(AssurePlusPricingReport report) =>
        $"{Path.GetFullPath(report.SourcePath)}|{report.SourcePage}|{report.EffectiveDate:yyyy-MM-dd}";

    private static int MergeSunLifeExperienceRows(PdfExtractionResult currentPdf, PdfExtractionResult supportingPdf)
    {
        var existingKeys = currentPdf.SunLifeExperienceRows
            .Select(SunLifeExperienceKey)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var added = 0;
        foreach (var row in supportingPdf.SunLifeExperienceRows)
        {
            if (!existingKeys.Add(SunLifeExperienceKey(row)))
            {
                continue;
            }

            currentPdf.SunLifeExperienceRows.Add(new SunLifeExperienceRow
            {
                Benefit = row.Benefit,
                Period = row.Period,
                AverageCoveredEmployees = row.AverageCoveredEmployees,
                BilledPremium = row.BilledPremium,
                AdjustedPremium = row.AdjustedPremium,
                PaidClaims = row.PaidClaims,
                IncurredClaims = row.IncurredClaims,
                AdjustedLossRatio = row.AdjustedLossRatio,
                Weighting = row.Weighting,
                LargeAmountPoolingPremium = row.LargeAmountPoolingPremium,
                QuebecIndustryPoolingPremium = row.QuebecIndustryPoolingPremium,
                PooledClaims = row.PooledClaims,
                QuebecIndustryPooledClaims = row.QuebecIndustryPooledClaims,
                WeightedAdjustedLossRatio = row.WeightedAdjustedLossRatio,
                Credibility = row.Credibility,
                Ibnr = row.Ibnr,
                TargetLossRatio = row.TargetLossRatio,
                AnnualTrendFactor = row.AnnualTrendFactor,
                RateAdjustment = row.RateAdjustment,
                SourceReference = row.SourceReference
            });
            added++;
        }

        return added;
    }

    private static int MergeEmpireLifeExperienceRows(PdfExtractionResult currentPdf, PdfExtractionResult supportingPdf)
    {
        var existingKeys = currentPdf.EmpireLifeExperienceRows
            .Select(row => $"{NormalizeMergeKey(row.Benefit)}|{NormalizeMergeKey(row.Period)}")
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var added = 0;
        foreach (var row in supportingPdf.EmpireLifeExperienceRows)
        {
            var key = $"{NormalizeMergeKey(row.Benefit)}|{NormalizeMergeKey(row.Period)}";
            if (!existingKeys.Add(key))
            {
                continue;
            }

            currentPdf.EmpireLifeExperienceRows.Add(new EmpireLifeExperienceRow
            {
                Benefit = row.Benefit,
                Period = row.Period,
                NonPooledPremium = row.NonPooledPremium,
                AdjustedPremium = row.AdjustedPremium,
                PaidClaims = row.PaidClaims,
                TrendedClaims = row.TrendedClaims,
                IncurredClaims = row.IncurredClaims,
                IncurredLossRatio = row.IncurredLossRatio,
                WeightedLossRatio = row.WeightedLossRatio,
                TargetLossRatio = row.TargetLossRatio,
                Credibility = row.Credibility,
                SourceReference = row.SourceReference
            });
            added++;
        }

        foreach (var row in supportingPdf.EmpireLifePooledBenefits)
        {
            if (currentPdf.EmpireLifePooledBenefits.Any(existing =>
                    existing.Benefit.Equals(row.Benefit, StringComparison.OrdinalIgnoreCase) &&
                    existing.BilledPremium == row.BilledPremium &&
                    existing.PaidClaims == row.PaidClaims))
            {
                continue;
            }

            currentPdf.EmpireLifePooledBenefits.Add(new EmpireLifePooledBenefit
            {
                Benefit = row.Benefit,
                BilledPremium = row.BilledPremium,
                PaidClaims = row.PaidClaims,
                SourceReference = row.SourceReference
            });
        }

        return added;
    }

    private static int MergeDemographicRows(List<DemographicLine> currentRows, IReadOnlyCollection<DemographicLine> supportingRows)
    {
        var existingKeys = currentRows
            .Select(row => $"{NormalizeMergeKey(row.Benefit)}|{DemographicAgeBandNormalizer.NormalizeLifeAgeBand(row.AgeBand)}")
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var added = 0;
        foreach (var row in supportingRows)
        {
            var key = $"{NormalizeMergeKey(row.Benefit)}|{DemographicAgeBandNormalizer.NormalizeLifeAgeBand(row.AgeBand)}";
            if (!existingKeys.Add(key))
            {
                continue;
            }

            currentRows.Add(new DemographicLine
            {
                Benefit = row.Benefit,
                AgeBand = row.AgeBand,
                CurrentFemaleLives = row.CurrentFemaleLives,
                CurrentFemaleVolume = row.CurrentFemaleVolume,
                CurrentMaleLives = row.CurrentMaleLives,
                CurrentMaleVolume = row.CurrentMaleVolume,
                PriorFemaleLives = row.PriorFemaleLives,
                PriorFemaleVolume = row.PriorFemaleVolume,
                PriorMaleLives = row.PriorMaleLives,
                PriorMaleVolume = row.PriorMaleVolume,
                SourceReference = row.SourceReference
            });
            added++;
        }

        return added;
    }

    private static int MergeCurrentWorkbookData(WorkbookExtractionResult currentWorkbook, WorkbookExtractionResult? supportingWorkbook)
    {
        if (supportingWorkbook is null)
        {
            return 0;
        }

        var updated = MergeWorkbookDataWithoutDrugs(currentWorkbook, supportingWorkbook);
        updated += MergeDrugRows(currentWorkbook.DrugTrends, supportingWorkbook.DrugTrends);
        updated += MergeDrugRows(currentWorkbook.PriorDrugTrends, supportingWorkbook.PriorDrugTrends);
        return updated;
    }

    private static int MergePriorWorkbookData(WorkbookExtractionResult currentWorkbook, WorkbookExtractionResult? supportingWorkbook)
    {
        if (supportingWorkbook is null)
        {
            return 0;
        }

        var updated = MergeWorkbookDataWithoutDrugs(currentWorkbook, supportingWorkbook);
        updated += MergeDrugRows(currentWorkbook.PriorDrugTrends, supportingWorkbook.DrugTrends);
        updated += MergeDrugRows(currentWorkbook.PriorDrugTrends, supportingWorkbook.PriorDrugTrends);
        return updated;
    }

    private static int MergeWorkbookDataWithoutDrugs(WorkbookExtractionResult currentWorkbook, WorkbookExtractionResult supportingWorkbook)
    {
        var updated = 0;
        updated += MergeClaimTrends(currentWorkbook.ClaimTrends, supportingWorkbook.ClaimTrends);
        updated += MergeClaimsDistributionRows(currentWorkbook, supportingWorkbook.ClaimsDistributionRows);
        updated += MergeCurrentServiceCategoryRows(currentWorkbook.EhcSpendingTrends, supportingWorkbook.EhcSpendingTrends);
        updated += MergeCurrentServiceCategoryRows(currentWorkbook.DentalSpendingTrends, supportingWorkbook.DentalSpendingTrends);
        updated += MergeCurrentServiceCategoryRows(currentWorkbook.ParamedicalServiceTrends, supportingWorkbook.ParamedicalServiceTrends);
        updated += SyncParamedicalRollupFromEhc(currentWorkbook);
        return updated;
    }

    private static int MergeClaimTrends(List<ClaimTrend> currentRows, IReadOnlyCollection<ClaimTrend> supportingRows)
    {
        var added = 0;
        foreach (var row in supportingRows)
        {
            var exists = currentRows.Any(current =>
                NormalizeMergeKey(current.Benefit).Equals(NormalizeMergeKey(row.Benefit), StringComparison.OrdinalIgnoreCase) &&
                NormalizeMergeKey(current.Category).Equals(NormalizeMergeKey(row.Category), StringComparison.OrdinalIgnoreCase) &&
                NormalizeMergeKey(current.Period).Equals(NormalizeMergeKey(row.Period), StringComparison.OrdinalIgnoreCase));
            if (exists)
            {
                continue;
            }

            currentRows.Add(row);
            added++;
        }

        return added;
    }

    private static int MergeClaimsDistributionRows(WorkbookExtractionResult currentWorkbook, IReadOnlyCollection<ClaimsDistributionRow> supportingRows)
    {
        if (supportingRows.Count == 0)
        {
            return 0;
        }

        var before = currentWorkbook.ClaimsDistributionRows.Count;
        var uniqueRows = new List<ClaimsDistributionRow>();
        foreach (var row in supportingRows)
        {
            var existing = currentWorkbook.ClaimsDistributionRows.FirstOrDefault(current =>
                SameClaimsDistributionContent(current, row));
            if (existing is null)
            {
                uniqueRows.Add(row);
                continue;
            }

            existing.SourceReference = MergeSourceReference(
                existing.SourceReference,
                string.IsNullOrWhiteSpace(row.SourceFile)
                    ? row.SourceReference
                    : $"{row.SourceFile}: {row.SourceReference}");
            existing.PercentPaid ??= row.PercentPaid;
            existing.PriorPercentPaid ??= row.PriorPercentPaid;
        }

        ClaimsDistributionExtractor.MergeIntoWorkbook(currentWorkbook, uniqueRows, updateServiceTrends: false);
        return currentWorkbook.ClaimsDistributionRows.Count - before;
    }

    private static bool SameClaimsDistributionContent(ClaimsDistributionRow left, ClaimsDistributionRow right)
    {
        return NormalizeMergeKey(left.CarrierName).Equals(NormalizeMergeKey(right.CarrierName), StringComparison.OrdinalIgnoreCase) &&
               NormalizeMergeKey(left.PolicyNumber).Equals(NormalizeMergeKey(right.PolicyNumber), StringComparison.OrdinalIgnoreCase) &&
               NormalizeMergeKey(left.ClientName).Equals(NormalizeMergeKey(right.ClientName), StringComparison.OrdinalIgnoreCase) &&
               NormalizeMergeKey(left.Benefit).Equals(NormalizeMergeKey(right.Benefit), StringComparison.OrdinalIgnoreCase) &&
               NormalizeMergeKey(left.ServiceCategory).Equals(NormalizeMergeKey(right.ServiceCategory), StringComparison.OrdinalIgnoreCase) &&
               NormalizeMergeKey(left.ServiceDescriptionOriginal).Equals(NormalizeMergeKey(right.ServiceDescriptionOriginal), StringComparison.OrdinalIgnoreCase) &&
               NormalizeMergeKey(left.ReportPeriod).Equals(NormalizeMergeKey(right.ReportPeriod), StringComparison.OrdinalIgnoreCase) &&
               NormalizeMergeKey(left.PriorReportPeriod).Equals(NormalizeMergeKey(right.PriorReportPeriod), StringComparison.OrdinalIgnoreCase) &&
               left.AmountPaid == right.AmountPaid &&
               SameOrMissing(left.PercentPaid, right.PercentPaid) &&
               left.PriorAmountPaid == right.PriorAmountPaid &&
               SameOrMissing(left.PriorPercentPaid, right.PriorPercentPaid) &&
               left.Occurrences == right.Occurrences &&
               left.PriorOccurrences == right.PriorOccurrences;
    }

    private static bool SameOrMissing(decimal? left, decimal? right)
    {
        return !left.HasValue || !right.HasValue || left.Value == right.Value;
    }

    private static int MergeCurrentServiceCategoryRows(List<ServiceCategoryTrend> currentRows, IReadOnlyCollection<ServiceCategoryTrend> supportingRows)
    {
        var updated = 0;
        foreach (var supporting in supportingRows)
        {
            var supportingKey = ServiceCategoryMergeKey(supporting);
            var matchingRows = currentRows
                .Where(row => ServiceCategoryMergeKey(row).Equals(supportingKey, StringComparison.OrdinalIgnoreCase))
                .ToList();
            var current = matchingRows.FirstOrDefault();
            if (current is null)
            {
                ServiceCategoryHistory.SynchronizeLegacyWindow(supporting);
                currentRows.Add(supporting);
                updated++;
                continue;
            }

            ConsolidateServiceCategoryMatches(currentRows, current, matchingRows);
            if (ServiceCategoryHistory.Merge(current, supporting))
            {
                updated++;
            }
        }

        return updated;
    }

    private static int SyncParamedicalRollupFromEhc(WorkbookExtractionResult workbook)
    {
        var ehcParamedical = workbook.EhcSpendingTrends.FirstOrDefault(row =>
            NormalizeMergeKey(row.Category).Equals("PARAMEDICAL", StringComparison.OrdinalIgnoreCase));
        if (ehcParamedical is null)
        {
            return 0;
        }

        var rollup = workbook.ParamedicalServiceTrends.FirstOrDefault(row =>
            IsParamedicalRollupCategory(row.Category));
        if (rollup is null)
        {
            workbook.ParamedicalServiceTrends.Add(new ServiceCategoryTrend
            {
                Benefit = "Extended Health Care",
                Category = "Paramedical Practitioners",
                PriorPeriodLabel = ehcParamedical.PriorPeriodLabel,
                PriorAmount = ehcParamedical.PriorAmount,
                CurrentPeriodLabel = ehcParamedical.CurrentPeriodLabel,
                CurrentAmount = ehcParamedical.CurrentAmount,
                SourceReference = ehcParamedical.SourceReference
            });
            return 1;
        }

        var updated = 0;
        if (rollup.PriorAmount == 0m && ehcParamedical.PriorAmount != 0m)
        {
            rollup.PriorPeriodLabel = ehcParamedical.PriorPeriodLabel;
            rollup.PriorAmount = ehcParamedical.PriorAmount;
            updated++;
        }

        if (rollup.CurrentAmount == 0m && ehcParamedical.CurrentAmount != 0m)
        {
            rollup.CurrentPeriodLabel = ehcParamedical.CurrentPeriodLabel;
            rollup.CurrentAmount = ehcParamedical.CurrentAmount;
            updated++;
        }

        if (updated > 0)
        {
            rollup.SourceReference = MergeSourceReference(rollup.SourceReference, ehcParamedical.SourceReference);
        }

        return updated;
    }

    private static void ConsolidateServiceCategoryMatches(
        List<ServiceCategoryTrend> currentRows,
        ServiceCategoryTrend target,
        IReadOnlyList<ServiceCategoryTrend> matchingRows)
    {
        foreach (var duplicate in matchingRows.Skip(1).ToList())
        {
            ServiceCategoryHistory.Merge(target, duplicate);
            currentRows.Remove(duplicate);
        }
    }

    private static bool ShouldPromoteExistingCurrentToPrior(ServiceCategoryTrend current, ServiceCategoryTrend supporting)
    {
        if (current.CurrentAmount == 0m || supporting.CurrentAmount == 0m)
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(current.CurrentPeriodLabel) ||
            string.IsNullOrWhiteSpace(supporting.CurrentPeriodLabel) ||
            NormalizeMergeKey(current.CurrentPeriodLabel).Equals(NormalizeMergeKey(supporting.CurrentPeriodLabel), StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return LooksLikeWorkbookSource(current.SourceReference) &&
               LooksLikeCarrierPdfSource(supporting.SourceReference);
    }

    private static bool LooksLikeWorkbookSource(string sourceReference)
    {
        return !string.IsNullOrWhiteSpace(sourceReference) &&
               !sourceReference.Contains("PDF", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeCarrierPdfSource(string sourceReference)
    {
        return sourceReference.Contains("PDF", StringComparison.OrdinalIgnoreCase) ||
               sourceReference.Contains("page", StringComparison.OrdinalIgnoreCase);
    }

    private static string ServiceCategoryMergeKey(ServiceCategoryTrend row)
    {
        var benefitKey = CompactMergeKey(row.Benefit);
        var categoryKey = CompactMergeKey(row.Category);

        if (benefitKey.Contains("DENTAL", StringComparison.OrdinalIgnoreCase))
        {
            return $"DENTAL|{NormalizeDentalMergeCategoryKey(categoryKey)}";
        }

        if (benefitKey.Contains("HEALTH", StringComparison.OrdinalIgnoreCase) ||
            benefitKey.Contains("EHC", StringComparison.OrdinalIgnoreCase) ||
            categoryKey.Contains("PARAMEDICAL", StringComparison.OrdinalIgnoreCase) ||
            IsParamedicalServiceCategoryKey(categoryKey))
        {
            return $"HEALTH|{NormalizeParamedicalCategoryKey(categoryKey)}";
        }

        return $"{benefitKey}|{categoryKey}";
    }

    private static bool IsParamedicalRollupCategory(string category)
    {
        var key = CompactMergeKey(category);
        return key.Equals("PARAMEDICAL", StringComparison.OrdinalIgnoreCase) ||
               key.Equals("PARAMEDICALPRACTITIONERS", StringComparison.OrdinalIgnoreCase) ||
               key.Equals("PARAMEDICALSERVICES", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeDentalMergeCategoryKey(string categoryKey)
    {
        return categoryKey switch
        {
            "PREVENTIVE" => "PREVENTATIVE",
            "PREVENTIVEDIAGNOSTIC" => "PREVENTATIVEDIAGNOSTIC",
            "MAJOR" or "MAJORRESTORATIVE" => "MAJORRESTORATIVE",
            "ENDODONTIC" => "ENDODONTICS",
            "PERIODONTAL" or "PERIODONTICS" => "PERIODONTIC",
            "ORTHODONTIC" or "ORTHODONTICS" => "ORTHODONTIC",
            "DENTURES" or "DENTURESBRIDGES" => "DENTURESBRIDGES",
            _ => categoryKey
        };
    }

    private static string NormalizeParamedicalCategoryKey(string categoryKey)
    {
        return categoryKey switch
        {
            "REGISTEREDMASSAGETHERAPIST" or "MASSAGETHERAPIST" or "MASSEUR" or "MASSAGE" => "MASSAGE",
            "PHYSIOTHERAPIST" or "PHYSIOTHERAPY" => "PHYSIOTHERAPY",
            "CHIROPRACTOR" or "CHIROPRACTIC" => "CHIROPRACTIC",
            "PSYCHOLOGY" or "PSYCHOLOGIST" => "PSYCHOLOGIST",
            "ACUPUNCTURE" or "ACUPUNCTURIST" => "ACUPUNCTURIST",
            "NATUROPATHIC" or "NATUROPATH" => "NATUROPATH",
            "OSTEOPATHIC" or "OSTEOPATH" => "OSTEOPATH",
            "CLINICALCOUNSELLOR" or "CLINICALCOUNSELLORS" or "COUNSELLOR" => "CLINICALCOUNSELLOR",
            "PARAMEDICAL" or "PARAMEDICALPRACTITIONERS" or "PARAMEDICALSERVICES" => "PARAMEDICAL",
            _ => categoryKey
        };
    }

    private static bool IsParamedicalServiceCategoryKey(string categoryKey)
    {
        var normalized = NormalizeParamedicalCategoryKey(categoryKey);
        return !normalized.Equals(categoryKey, StringComparison.OrdinalIgnoreCase) ||
               normalized.Equals("PARAMEDICAL", StringComparison.OrdinalIgnoreCase);
    }

    private static string CompactMergeKey(string value)
    {
        return new string(NormalizeMergeKey(value)
            .Where(char.IsLetterOrDigit)
            .ToArray());
    }

    private PreviousCarrierRenewalExtraction? TryReadCarrierPreviousRenewal(string candidatePath)
    {
        try
        {
            var parsed = _pdfParser.Parse(candidatePath);
            var snapshot = PreviousRenewalReportParser.FromCarrierExtraction(candidatePath, parsed.Pdf);
            return snapshot is null ? null : new PreviousCarrierRenewalExtraction(snapshot, parsed.Pdf, parsed.WorkbookFromPdf);
        }
        catch
        {
            return null;
        }
    }

    private static string MergePreviousCarrierPdfData(
        AnalysisResult result,
        PdfExtractionResult previousPdf,
        WorkbookExtractionResult? previousWorkbook,
        PreviousRenewalSnapshot snapshot)
    {
        var experienceRowsAdded = MergePreviousExperienceRows(result.Pdf, previousPdf);
        var assurePlusReportsAdded = MergeAssurePlusPricingReports(result.Pdf, previousPdf);
        var categoryRowsUpdated =
            MergePreviousServiceCategoryRows(result.Workbook.EhcSpendingTrends, previousWorkbook?.EhcSpendingTrends, snapshot) +
            MergePreviousServiceCategoryRows(result.Workbook.DentalSpendingTrends, previousWorkbook?.DentalSpendingTrends, snapshot) +
            MergePreviousServiceCategoryRows(result.Workbook.ParamedicalServiceTrends, previousWorkbook?.ParamedicalServiceTrends, snapshot);
        var topDrugRowsAdded =
            MergeDrugRows(result.Workbook.PriorDrugTrends, previousWorkbook?.DrugTrends) +
            MergeDrugRows(result.Workbook.PriorDrugTrends, previousWorkbook?.PriorDrugTrends);

        if (experienceRowsAdded == 0 && assurePlusReportsAdded == 0 && categoryRowsUpdated == 0 && topDrugRowsAdded == 0)
        {
            return "";
        }

        var parts = new List<string>();
        if (experienceRowsAdded > 0)
        {
            parts.Add($"{experienceRowsAdded} premium/claims rows merged");
        }

        if (assurePlusReportsAdded > 0)
        {
            parts.Add($"{assurePlusReportsAdded} AssurePLUS pricing report captured");
        }

        if (categoryRowsUpdated > 0)
        {
            parts.Add($"{categoryRowsUpdated} claims category rows merged");
        }

        if (topDrugRowsAdded > 0)
        {
            parts.Add($"{topDrugRowsAdded} prior top-drug rows merged");
        }

        return $"; {string.Join("; ", parts)}";
    }

    private static int MergeDrugRows(List<DrugTrend> currentRows, IReadOnlyCollection<DrugTrend>? supportingRows)
    {
        if (supportingRows is null || supportingRows.Count == 0)
        {
            return 0;
        }

        var existingKeys = currentRows
            .Select(DrugKey)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var added = 0;
        foreach (var row in supportingRows)
        {
            if (string.IsNullOrWhiteSpace(row.DrugName))
            {
                continue;
            }

            if (!existingKeys.Add(DrugKey(row)))
            {
                continue;
            }

            currentRows.Add(new DrugTrend
            {
                Din = row.Din,
                DrugName = row.DrugName,
                Treats = row.Treats,
                Period = row.Period,
                PrescriptionCount = row.PrescriptionCount,
                SubmittedTotal = row.SubmittedTotal,
                NetAmountPaid = row.NetAmountPaid,
                AllowableTotal = row.AllowableTotal,
                PriorPaid = row.PriorPaid,
                SourceReference = row.SourceReference
            });
            added++;
        }

        return added;
    }

    private static int PromoteClaimsWorkbookTopDrugsToPriorHistory(WorkbookExtractionResult workbook)
    {
        if (workbook.DrugTrends.Count == 0)
        {
            return 0;
        }

        var added = MergeDrugRows(workbook.PriorDrugTrends, workbook.DrugTrends);
        if (added > 0)
        {
            workbook.DrugTrends.Clear();
        }

        return added;
    }

    private static int MergePreviousExperienceRows(PdfExtractionResult currentPdf, PdfExtractionResult previousPdf)
    {
        return MergeExperienceRows(currentPdf, previousPdf, markAsPrevious: true);
    }

    private static int MergeExperienceRows(
        PdfExtractionResult currentPdf,
        PdfExtractionResult supportingPdf,
        bool markAsPrevious)
    {
        var existingRows = currentPdf.ExperienceLines
            .GroupBy(ExperienceKey, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

        var updated = 0;
        foreach (var row in supportingPdf.ExperienceLines.Where(IsMergeablePreviousExperienceRow))
        {
            var key = ExperienceKey(row);
            if (!existingRows.TryGetValue(key, out var existing))
            {
                var clone = CloneExperienceRow(row, markAsPrevious);
                currentPdf.ExperienceLines.Add(clone);
                existingRows[key] = clone;
                updated++;
                continue;
            }

            if (EnrichMissingExperienceValues(existing, row, markAsPrevious))
            {
                updated++;
            }
        }

        return updated;
    }

    private static bool EnrichMissingExperienceValues(
        ExperienceLine target,
        ExperienceLine source,
        bool markAsPrevious)
    {
        var updated = false;
        if (ShouldPreferEquitableExperienceSummary(target, source))
        {
            if (source.AverageLives.HasValue && target.AverageLives != source.AverageLives) { target.AverageLives = source.AverageLives; updated = true; }
            if (source.BilledPremium.HasValue && target.BilledPremium != source.BilledPremium) { target.BilledPremium = source.BilledPremium; updated = true; }
            if (source.PaidClaims.HasValue && target.PaidClaims != source.PaidClaims) { target.PaidClaims = source.PaidClaims; updated = true; }
            if (source.IncurredClaims.HasValue && target.IncurredClaims != source.IncurredClaims) { target.IncurredClaims = source.IncurredClaims; updated = true; }
            if (source.PaidLossRatio.HasValue && target.PaidLossRatio != source.PaidLossRatio) { target.PaidLossRatio = source.PaidLossRatio; updated = true; }
            if (source.IncurredLossRatio.HasValue && target.IncurredLossRatio != source.IncurredLossRatio) { target.IncurredLossRatio = source.IncurredLossRatio; updated = true; }
        }

        if (!target.AverageLives.HasValue && source.AverageLives.HasValue) { target.AverageLives = source.AverageLives; updated = true; }
        if (!target.BilledPremium.HasValue && source.BilledPremium.HasValue) { target.BilledPremium = source.BilledPremium; updated = true; }
        if (!target.PaidClaims.HasValue && source.PaidClaims.HasValue) { target.PaidClaims = source.PaidClaims; updated = true; }
        if (!target.IncurredClaims.HasValue && source.IncurredClaims.HasValue) { target.IncurredClaims = source.IncurredClaims; updated = true; }
        if (!target.PaidLossRatio.HasValue && source.PaidLossRatio.HasValue) { target.PaidLossRatio = source.PaidLossRatio; updated = true; }
        if (!target.IncurredLossRatio.HasValue && source.IncurredLossRatio.HasValue) { target.IncurredLossRatio = source.IncurredLossRatio; updated = true; }
        if (!target.Weighting.HasValue && source.Weighting.HasValue) { target.Weighting = source.Weighting; updated = true; }
        if (!target.ProposedAdjustment.HasValue && source.ProposedAdjustment.HasValue) { target.ProposedAdjustment = source.ProposedAdjustment; updated = true; }
        if (!target.TrendFactor.HasValue && source.TrendFactor.HasValue) { target.TrendFactor = source.TrendFactor; updated = true; }
        if (!target.Credibility.HasValue && source.Credibility.HasValue) { target.Credibility = source.Credibility; updated = true; }
        if (!target.TargetLossRatio.HasValue && source.TargetLossRatio.HasValue) { target.TargetLossRatio = source.TargetLossRatio; updated = true; }

        var sourceReference = ExperienceSourceReference(source.SourceReference, markAsPrevious);
        if (updated && !string.IsNullOrWhiteSpace(sourceReference) &&
            !target.SourceReference.Contains(sourceReference, StringComparison.OrdinalIgnoreCase))
        {
            target.SourceReference = string.IsNullOrWhiteSpace(target.SourceReference)
                ? sourceReference
                : $"{target.SourceReference}; {sourceReference}";
        }

        return updated;
    }

    private static bool ShouldPreferEquitableExperienceSummary(ExperienceLine target, ExperienceLine source)
    {
        return target.SourceReference.Contains("Equitable renewal report", StringComparison.OrdinalIgnoreCase) &&
               source.SourceReference.Contains("Equitable Experience Summary", StringComparison.OrdinalIgnoreCase);
    }

    private static int MergePreviousServiceCategoryRows(
        List<ServiceCategoryTrend> currentRows,
        IReadOnlyCollection<ServiceCategoryTrend>? previousRows,
        PreviousRenewalSnapshot snapshot)
    {
        if (previousRows is null || previousRows.Count == 0)
        {
            return 0;
        }

        var fallbackLabel = PreviousRenewalLabel(snapshot);
        var updated = 0;
        foreach (var previous in previousRows)
        {
            var periods = ServiceCategoryHistory.GetPeriods(previous, int.MaxValue)
                .Where(period => !string.IsNullOrWhiteSpace(period.PeriodLabel) && period.Amount != 0m)
                .ToList();
            if (periods.Count == 0)
            {
                var previousAmount = previous.CurrentAmount != 0m ? previous.CurrentAmount : previous.PriorAmount;
                if (previousAmount == 0m) continue;
                periods.Add(new ServiceCategoryPeriodValue
                {
                    PeriodLabel = fallbackLabel,
                    Amount = previousAmount,
                    SourceReference = previous.SourceReference
                });
            }

            var current = currentRows.FirstOrDefault(row =>
                NormalizeMergeKey(row.Benefit).Equals(NormalizeMergeKey(previous.Benefit), StringComparison.OrdinalIgnoreCase) &&
                NormalizeMergeKey(row.Category).Equals(NormalizeMergeKey(previous.Category), StringComparison.OrdinalIgnoreCase));
            if (current is null)
            {
                current = new ServiceCategoryTrend { Benefit = previous.Benefit, Category = previous.Category };
                currentRows.Add(current);
            }

            var before = ServiceCategoryHistory.GetPeriods(current, int.MaxValue).Count;
            foreach (var period in periods)
            {
                ServiceCategoryHistory.Upsert(current, period.PeriodLabel, period.Amount, $"Previous renewal PDF - {period.SourceReference}");
            }
            current.SourceReference = MergeSourceReference(current.SourceReference, $"Previous renewal PDF - {previous.SourceReference}");
            if (ServiceCategoryHistory.GetPeriods(current, int.MaxValue).Count > before) updated++;
        }

        return updated;
    }

    private static bool IsMergeablePreviousExperienceRow(ExperienceLine row)
    {
        return !string.IsNullOrWhiteSpace(row.Period) &&
               !row.Period.Contains("Weighted renewal calculation", StringComparison.OrdinalIgnoreCase) &&
               (row.BilledPremium.HasValue || row.PaidClaims.HasValue || row.IncurredClaims.HasValue);
    }

    private static string ExperienceKey(ExperienceLine row)
    {
        return $"{NormalizeMergeKey(row.Benefit)}|{NormalizeMergeKey(row.Period)}";
    }

    private static string SunLifeExperienceKey(SunLifeExperienceRow row)
    {
        return $"{NormalizeMergeKey(row.Benefit)}|{NormalizeMergeKey(row.Period)}";
    }

    private static string DrugKey(DrugTrend row)
    {
        return $"{NormalizeMergeKey(row.Period)}|{NormalizeMergeKey(row.Din)}|{NormalizeMergeKey(row.DrugName)}";
    }

    private static ExperienceLine CloneExperienceRow(ExperienceLine row, bool markAsPrevious)
    {
        return new ExperienceLine
        {
            Benefit = row.Benefit,
            Period = row.Period,
            AverageLives = row.AverageLives,
            BilledPremium = row.BilledPremium,
            PaidClaims = row.PaidClaims,
            IncurredClaims = row.IncurredClaims,
            PaidLossRatio = row.PaidLossRatio,
            IncurredLossRatio = row.IncurredLossRatio,
            Weighting = row.Weighting,
            ProposedAdjustment = row.ProposedAdjustment,
            TrendFactor = row.TrendFactor,
            Credibility = row.Credibility,
            TargetLossRatio = row.TargetLossRatio,
            SourceReference = ExperienceSourceReference(row.SourceReference, markAsPrevious)
        };
    }

    private static string ExperienceSourceReference(string sourceReference, bool markAsPrevious)
    {
        if (!markAsPrevious || string.IsNullOrWhiteSpace(sourceReference) ||
            sourceReference.StartsWith("Previous renewal PDF - ", StringComparison.OrdinalIgnoreCase))
        {
            return sourceReference;
        }

        return $"Previous renewal PDF - {sourceReference}";
    }

    private static string PreviousRenewalLabel(PreviousRenewalSnapshot snapshot)
    {
        return snapshot.RenewalDate.HasValue
            ? $"{snapshot.RenewalDate.Value.Year} Renewal"
            : "Previous Renewal";
    }

    private static string MergeSourceReference(string current, string addition)
    {
        if (string.IsNullOrWhiteSpace(current))
        {
            return addition;
        }

        return current.Contains(addition, StringComparison.OrdinalIgnoreCase)
            ? current
            : $"{current}; {addition}";
    }

    private static string NormalizeMergeKey(string value)
    {
        return ParsingHelpers.Clean(value).ToUpperInvariant();
    }

    private static bool IsPriorRenewalSnapshot(PreviousRenewalSnapshot snapshot, AnalysisResult result)
    {
        if (!snapshot.HasUsefulValues)
        {
            return false;
        }

        if (PathsEqual(snapshot.SourcePath, result.RenewalCase.PdfPath))
        {
            return false;
        }

        var hasPriorValues =
            snapshot.RenewalDate.HasValue ||
            snapshot.CurrentMonthlyPremium.HasValue ||
            snapshot.RenewalMonthlyPremium.HasValue ||
            snapshot.MonthlyPremiumChange.HasValue ||
            snapshot.MonthlyPremiumChangePercent.HasValue;
        if (!hasPriorValues)
        {
            return false;
        }

        var currentRenewalDate = result.RenewalCase.RenewalDate ?? result.Pdf.RenewalDate;
        if (snapshot.RenewalDate.HasValue && currentRenewalDate.HasValue)
        {
            return snapshot.RenewalDate.Value < currentRenewalDate.Value;
        }

        return !LooksLikeCurrentRenewalSnapshot(snapshot, result);
    }

    private static bool LooksLikeCurrentRenewalSnapshot(PreviousRenewalSnapshot snapshot, AnalysisResult result)
    {
        return SameTextValue(snapshot.PolicyNumber, result.RenewalCase.PolicyNumber) &&
               SameMoneyValue(snapshot.CurrentMonthlyPremium, result.CurrentMonthlyPremium) &&
               SameMoneyValue(snapshot.RenewalMonthlyPremium, result.RenewalMonthlyPremium);
    }

    private static string DescribePreviousRenewal(PreviousRenewalSnapshot snapshot)
    {
        var parts = new List<string>();
        if (snapshot.RenewalDate.HasValue)
        {
            parts.Add($"renewal {snapshot.RenewalDate.Value:dd-MMM-yyyy}");
        }

        if (snapshot.RenewalMonthlyPremium.HasValue)
        {
            parts.Add($"proposed {ParsingHelpers.Money(snapshot.RenewalMonthlyPremium)} monthly");
        }

        if (snapshot.MonthlyPremiumChangePercent.HasValue)
        {
            parts.Add($"{ParsingHelpers.Percent(snapshot.MonthlyPremiumChangePercent)} change");
        }

        return parts.Count == 0 ? "prior-renewal context found" : string.Join("; ", parts);
    }

    private static bool PathsEqual(string left, string right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
        {
            return false;
        }

        return string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsFileInFolder(string filePath, string folderPath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(folderPath))
        {
            return false;
        }

        try
        {
            var fullFilePath = Path.GetFullPath(filePath);
            var fullFolderPath = Path.GetFullPath(folderPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            return fullFilePath.StartsWith(fullFolderPath, StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return false;
        }
    }

    private static bool SameTextValue(string left, string right)
    {
        return !string.IsNullOrWhiteSpace(left) &&
               !string.IsNullOrWhiteSpace(right) &&
               string.Equals(ParsingHelpers.Clean(left), ParsingHelpers.Clean(right), StringComparison.OrdinalIgnoreCase);
    }

    private static bool SameMoneyValue(decimal? left, decimal? right)
    {
        return left.HasValue &&
               right.HasValue &&
               decimal.Round(left.Value, 2) == decimal.Round(right.Value, 2);
    }

    private static bool ShouldTreatWorkbookAsPriorCarrierHistory(RenewalCase renewalCase, string carrier)
    {
        return !string.IsNullOrWhiteSpace(carrier) &&
               !carrier.Contains("Manulife", StringComparison.OrdinalIgnoreCase) &&
               FileNameLooksLikeSupportingClaimsWorkbook(renewalCase.WorkbookPath);
    }

    private static bool ShouldRetryClaimsWorkbookAsPriorCarrierHistory(RenewalCase renewalCase, Exception exception)
    {
        if (!exception.Message.Contains("Claims workbook mismatch", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return FileNameLooksLikeSupportingClaimsWorkbook(renewalCase.WorkbookPath) ||
               renewalCase.AdditionalWorkbookPaths.Any(FileNameLooksLikeSupportingClaimsWorkbook) ||
               GetNearbySupportingWorkbookPaths(renewalCase.WorkbookPath).Any();
    }

    private static bool IsGeneratedOutputPath(string filePath)
    {
        var parts = Path.GetFullPath(filePath).Split(
            [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
            StringSplitOptions.RemoveEmptyEntries);

        return parts.Any(part =>
            part.Equals("Batch Output", StringComparison.OrdinalIgnoreCase) ||
            part.Equals("Supporting PDFs", StringComparison.OrdinalIgnoreCase) ||
            part.Equals("Supporting Excel Files", StringComparison.OrdinalIgnoreCase));
    }

    private sealed record PreviousCarrierRenewalExtraction(
        PreviousRenewalSnapshot Snapshot,
        PdfExtractionResult Pdf,
        WorkbookExtractionResult? Workbook);

    private static List<string> CopySupportingPdfs(IEnumerable<string> paths, string outputFolder)
    {
        var sourcePaths = paths
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (sourcePaths.Count == 0)
        {
            return [];
        }

        var supportingFolder = Path.Combine(outputFolder, "Supporting PDFs");
        Directory.CreateDirectory(supportingFolder);
        var copied = new List<string>();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var sourcePath in sourcePaths)
        {
            var fileName = SafeFileName(Path.GetFileName(sourcePath));
            var destinationPath = UniqueDestinationPath(supportingFolder, fileName, usedNames);
            File.Copy(sourcePath, destinationPath, overwrite: true);
            copied.Add(destinationPath);
        }

        return copied;
    }

    private static List<string> GetSupportingPdfPaths(RenewalCase renewalCase)
    {
        return renewalCase.AdditionalPdfPaths
            .Concat(EnumerateFilesFromFolder(renewalCase.CurrentYearFolder, "*.pdf"))
            .Concat(EnumerateFilesFromFolder(renewalCase.PriorYearFolder, "*.pdf"))
            .Where(path => !PathsEqual(path, renewalCase.PdfPath))
            .Where(path => !IsGeneratedOutputPath(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> GetInputPdfPaths(RenewalCase renewalCase)
    {
        return new[] { renewalCase.PdfPath }
            .Concat(GetSupportingPdfPaths(renewalCase))
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Where(path => !IsGeneratedOutputPath(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static DiscoveredDocument? SelectDiscoveryRenewalDocument(IReadOnlyList<DiscoveredDocument> documents)
    {
        return documents
            .Where(document =>
                document.DocumentType.Equals(DiscoveryDocumentTypes.SunLifeRenewalReport, StringComparison.OrdinalIgnoreCase) ||
                document.IsRenewal)
            .OrderByDescending(document => document.DocumentType.Equals(DiscoveryDocumentTypes.SunLifeRenewalReport, StringComparison.OrdinalIgnoreCase))
            .ThenByDescending(document => document.RateEffectiveDate ?? document.RenewalDate ?? DateOnly.MinValue)
            .ThenBy(document => document.FileName, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();
    }

    private static List<string> CurrentNemSourcePdfPaths(RenewalCase renewalCase)
    {
        return new[] { renewalCase.PdfPath }
            .Concat(renewalCase.AdditionalPdfPaths)
            .Concat(EnumerateFilesFromFolder(renewalCase.CurrentYearFolder, "*.pdf"))
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Where(path => !IsGeneratedOutputPath(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> GetSupportingWorkbookPaths(RenewalCase renewalCase)
    {
        var selectedLegacyClaimsWorkbook = IsLegacyClaimsWorkbookPath(renewalCase.WorkbookPath)
            ? new[] { renewalCase.WorkbookPath }
            : [];

        return selectedLegacyClaimsWorkbook
            .Concat(renewalCase.AdditionalWorkbookPaths)
            .Concat(EnumerateFilesFromFolder(renewalCase.CurrentYearFolder, "*.xlsx"))
            .Concat(EnumerateFilesFromFolder(renewalCase.CurrentYearFolder, "*.xlsm"))
            .Concat(EnumerateFilesFromFolder(renewalCase.CurrentYearFolder, "*.xls"))
            .Concat(EnumerateFilesFromFolder(renewalCase.PriorYearFolder, "*.xlsx"))
            .Concat(EnumerateFilesFromFolder(renewalCase.PriorYearFolder, "*.xlsm"))
            .Concat(EnumerateFilesFromFolder(renewalCase.PriorYearFolder, "*.xls"))
            .Concat(GetNearbySupportingWorkbookPaths(renewalCase.WorkbookPath))
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Where(path => IsLegacyClaimsWorkbookPath(renewalCase.WorkbookPath) || !PathsEqual(path, renewalCase.WorkbookPath))
            .Where(path => !IsGeneratedOutputPath(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static IEnumerable<string> GetNearbySupportingWorkbookPaths(string workbookPath)
    {
        if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
        {
            return [];
        }

        var folder = Path.GetDirectoryName(workbookPath);
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            return [];
        }

        return Directory.EnumerateFiles(folder, "*.xlsx", SearchOption.TopDirectoryOnly)
            .Concat(Directory.EnumerateFiles(folder, "*.xlsm", SearchOption.TopDirectoryOnly))
            .Concat(Directory.EnumerateFiles(folder, "*.xls", SearchOption.TopDirectoryOnly))
            .Where(FileNameLooksLikeSupportingClaimsWorkbook);
    }

    private static bool FileNameLooksLikeSupportingClaimsWorkbook(string path)
    {
        var fileName = Path.GetFileNameWithoutExtension(path);
        var claimsHistorySignals = new[]
        {
            "Claims",
            "Claim",
            "Experience",
            "Utilization",
            "Client_Copy",
            "Client Copy",
            "Top DIN",
            "TopDIN",
            "DIN",
            "Drug",
            "Drugs",
            "EHC",
            "Health",
            "Dental",
            "Paramed",
            "Paramedical"
        };

        return claimsHistorySignals.Any(signal =>
            fileName.Contains(signal, StringComparison.OrdinalIgnoreCase));
    }

    private static bool FileNameLooksLikeClientCopyWorkbook(string path)
    {
        var fileName = Path.GetFileNameWithoutExtension(path);
        return fileName.Contains("Client_Copy", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("Client Copy", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("Client-Copy", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsLegacyClaimsWorkbookPath(string path)
    {
        return !string.IsNullOrWhiteSpace(path) &&
               File.Exists(path) &&
               Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase) &&
               FileNameLooksLikeSupportingClaimsWorkbook(path);
    }

    private static IEnumerable<string> EnumerateFilesFromFolder(string folder, string pattern)
    {
        return !string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder)
            ? Directory.EnumerateFiles(folder, pattern, SearchOption.TopDirectoryOnly)
            : [];
    }

    private static IEnumerable<string> EnumerateSourceFolderFiles(string folder)
    {
        return EnumerateFilesFromFolder(folder, "*.pdf")
            .Concat(EnumerateFilesFromFolder(folder, "*.xlsx"))
            .Concat(EnumerateFilesFromFolder(folder, "*.xlsm"))
            .Concat(EnumerateFilesFromFolder(folder, "*.xls"));
    }

    private static List<string> CopySupportingWorkbooks(IEnumerable<string> paths, string outputFolder)
    {
        var sourcePaths = paths
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Where(path =>
                path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (sourcePaths.Count == 0)
        {
            return [];
        }

        var targetFolder = Path.Combine(outputFolder, "Supporting Excel Files");
        Directory.CreateDirectory(targetFolder);

        var copied = new List<string>();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sourcePath in sourcePaths)
        {
            var fileName = SafeFileName(Path.GetFileName(sourcePath), "Supporting Workbook.xlsx");
            var targetPath = UniqueDestinationPath(targetFolder, fileName, usedNames);
            File.Copy(sourcePath, targetPath, overwrite: true);
            copied.Add(targetPath);
        }

        return copied;
    }

    private static string SafeFileName(string fileName, string fallback = "Supporting PDF.pdf")
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(fileName.Select(character => invalid.Contains(character) ? '_' : character).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? fallback : cleaned;
    }

    private static string UniqueDestinationPath(string folder, string fileName, HashSet<string> usedNames)
    {
        var stem = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);
        var candidate = fileName;
        var index = 2;

        while (!usedNames.Add(candidate))
        {
            candidate = $"{stem} ({index}){extension}";
            index++;
        }

        return Path.Combine(folder, candidate);
    }

    private static void SortFindings(AnalysisResult result)
    {
        var sorted = result.Findings
            .OrderBy(finding => finding.Severity == FindingSeverity.Critical ? 0 : finding.Severity == FindingSeverity.Warning ? 1 : 2)
            .ThenBy(finding => finding.Title)
            .ToList();

        result.Findings.Clear();
        result.Findings.AddRange(sorted);
    }
}
