using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class DocumentDiscoveryService
{
    private readonly PdfTextExtractor _pdfTextExtractor = new();
    private readonly AsoFinancialActivityWorkbookParser _asoFinancialActivityWorkbookParser = new();
    private static readonly Regex RenewalDateRegex = new(
        @"(?:Renewal Date|Renewal Effective|Effective Date|Effective|Implementation Date)\s*[:\-]?\s*(?<date>[A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4}|\d{1,2}[-/][A-Za-z]{3,9}[-/]\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{1,2}-\d{1,2})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex FileNameYearRegex = new(@"(?<!\d)(?<year>20\d{2})(?!\d)", RegexOptions.Compiled);
    private static readonly Regex GroupInformationReportRegex = new(
        @"\bgroup[^A-Za-z0-9]*information[^A-Za-z0-9]*report\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex GirTokenRegex = new(
        @"(?<![A-Za-z0-9])GIR(?![A-Za-z0-9])",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeContractNumberRegex = new(
        @"(?:Contract\s+number|Re:\s*Contract)\s*:?\s*(?<value>\d+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeContractHolderRegex = new(
        @"Contract\s+holder\s*:?\s*(?<value>[A-Za-z0-9 &'.,\-]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeRateEffectiveDateRegex = new(
        @"Rate\s+effective\s+date\s*:?\s*(?<date>[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}|\d{1,2}-[A-Za-z]{3,9}-\d{2,4})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".pdf",
        ".xlsx",
        ".xlsm",
        ".xls",
        ".docx"
    };

    private static readonly PlanDesignSectionDefinition[] PlanDesignSections =
    [
        new("Group Information Report", 2, ["GROUP INFORMATION REPORT"]),
        new("Schedule of Benefits", 3, ["Schedule of Benefits"]),
        new("Benefit Summary", 4, ["Benefit Summary"]),
        new("Plan Design", 4, ["Plan Design"]),
        new("Coverage Summary", 4, ["Coverage Summary"]),
        new("Benefit Details", 4, ["Benefit Details"]),
        new("Class Information", 4, ["Class Information", "Admin Class Listings"])
    ];

    private static readonly string[] SunLifeRenewalReportSignals =
    [
        "Renewal Report",
        "Summary of rates and estimated monthly premium",
        "Rate effective date",
        "Contract number",
        "Contract holder",
        "Coverage and estimated monthly premium by employee",
        "Demographic distribution",
        "Employee Life",
        "Long-Term Disability",
        "Non-Evidence Maximums",
        "NEM"
    ];

    private static readonly string[] SunLifeExperienceAnalysisSignals =
    [
        "Experience analysis",
        "Extended Health Care Renewal Factors",
        "Dental Renewal Factors",
        "Weighted Adjusted Loss Ratio",
        "Credibility",
        "IBNR (% of premium)",
        "Target Loss Ratio",
        "Annual Trend Factor",
        "Rate Adjustment",
        "Billed Premium",
        "Adjusted Premium",
        "Paid Claims",
        "Incurred Claims",
        "Adjusted Loss Ratio",
        "Large Amount Pooling Premium",
        "Pooled Claims"
    ];
    private static readonly string[] GroupHealthRenewalSignals =
    [
        "Group Benefits Plan Renewal",
        "Premium Comparison",
        "Life Age & Volume Distribution",
        "LTD Age & Volume Distribution",
        "Renewal Analysis",
        "Claims Experience Period"
    ];

    private static readonly string[] GroupHealthClaimsSignals =
    [
        "Billed Premium & Claims",
        "Current Underwriter SSQ FINANCIAL",
        "Expense Information",
        "Paramedical Information",
        "Cost Category"
    ];

    private static readonly string[] GroupHealthSupportSignals =
    [
        "Top 10 DINs by Dollar Amount",
        "Top 50 DINs by Dollar Amount",
        "Claims Range by Certificate",
        "WOP Reserve Report",
        "DLR Reserve Report",
        "EP3 Statement",
        "Pooling Level for prescription drug"
    ];


    private static readonly string[] GroupPolicyScheduleTitleSignals =
    [
        "Group Benefits Schedule",
        "Group Benefits Policy",
        "Policy Schedule",
        "Group Policy"
    ];

    private static readonly string[] GroupPolicyScheduleDetailSignals =
    [
        "Benefit Amount",
        "Non-Evidence Limit",
        "Participation Basis",
        "Waiting Period",
        "Qualifying Period",
        "Termination Age",
        "Waiver of Premium",
        "Maximum Benefit Period",
        "Classification Code"
    ];

    private static readonly string[] GroupPolicyScheduleCoverageSignals =
    [
        "Employee Life Insurance",
        "Dependent Life Insurance",
        "Accidental Death and Dismemberment",
        "Long Term Disability"
    ];

    public RenewalDocumentDiscoveryResult DiscoverClientFolder(string clientFolder)
    {
        var result = new RenewalDocumentDiscoveryResult
        {
            ClientFolder = clientFolder
        };

        if (!Directory.Exists(clientFolder))
        {
            result.PlanDesign = PlanDesignSourceDiscovery.None();
            result.PlanDesign.Detail = $"Client folder was not found: {clientFolder}";
            return result;
        }

        foreach (var documentPath in Directory.EnumerateFiles(clientFolder, "*.*", SearchOption.AllDirectories)
                     .Where(path => SupportedExtensions.Contains(Path.GetExtension(path)))
                     .Where(path => !IsGeneratedOutputPath(clientFolder, path))
                     .OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
        {
            result.Documents.Add(ClassifyDocument(documentPath));
        }

        result.RenewalDocument = SelectRenewalDocument(result.Documents);
        result.PlanDesign = SelectPlanDesignSource(result.Documents, result.RenewalDocument);
        return result;
    }

    public DiscoveredDocument ClassifyDocument(string filePath)
    {
        if (Path.GetExtension(filePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return ClassifyPdf(filePath);
        }

        var document = ClassifyFileName(filePath);
        if (_asoFinancialActivityWorkbookParser.LooksLikeAsoFinancialActivityWorkbook(filePath))
        {
            document.DocumentType = DiscoveryDocumentTypes.AsoFinancialActivity;
            document.DocumentConfidenceScore = 0.98m;
            document.IsAsoFinancialActivity = true;
            document.IsBillingStatement = false;
            document.IsClaimsExperience = false;
            document.MatchedSignals.Add("Financial Summary of Activity");

            try
            {
                var statement = _asoFinancialActivityWorkbookParser.Parse(filePath);
                document.ContractHolder = statement.ClientName;
                document.ContractNumber = statement.ContractNumber;
                document.RateEffectiveDate = statement.PeriodEnd;
                document.Carrier = string.IsNullOrWhiteSpace(statement.Carrier)
                    ? document.Carrier
                    : statement.Carrier;
                document.CarrierConfidenceScore = string.IsNullOrWhiteSpace(statement.Carrier) ? 0m : 0.95m;
                document.Status = statement.ImportStatus == AsoFinancialImportStatuses.NeedsReview
                    ? "needs_review"
                    : "ready_for_extraction";
            }
            catch (Exception exception)
            {
                document.DocumentConfidenceScore = 0.75m;
                document.Status = "needs_review";
                document.MatchedSignals.Add($"ASO workbook requires review: {exception.Message}");
            }

            return document;
        }

        if (ManulifeExcelBillingImportService.LooksLikeManulifeBillingWorkbook(filePath))
        {
            document.DocumentType = DiscoveryDocumentTypes.BillingStatement;
            document.Carrier = "Manulife";
            document.CarrierConfidenceScore = 0.95m;
            document.DocumentConfidenceScore = 0.95m;
            document.IsBillingStatement = true;
            document.IsClaimsExperience = false;
        }

        return document;
    }

    public DiscoveredDocument ClassifyPdf(string pdfPath)
    {
        try
        {
            var pages = _pdfTextExtractor.Extract(pdfPath);
            return ClassifyPdfPages(pdfPath, pages);
        }
        catch
        {
            return ClassifyFileName(pdfPath);
        }
    }

    public static DiscoveredDocument ClassifyFileName(string filePath)
    {
        var fileName = Path.GetFileName(filePath);
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
        var extension = Path.GetExtension(filePath);
        var isPdf = extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase);
        var sunLifeFileNameType = isPdf ? ClassifySunLifeFileName(fileNameWithoutExtension) : "";
        if (!string.IsNullOrWhiteSpace(sunLifeFileNameType))
        {
            return new DiscoveredDocument
            {
                FilePath = filePath,
                FileName = fileName,
                FileExtension = extension,
                DocumentType = sunLifeFileNameType,
                Carrier = "Sun Life",
                CarrierConfidenceScore = 0.6m,
                DocumentConfidenceScore = 0.55m,
                IsRenewal = sunLifeFileNameType.Equals(DiscoveryDocumentTypes.SunLifeRenewalReport, StringComparison.OrdinalIgnoreCase),
                IsClaimsExperience = sunLifeFileNameType.Equals(DiscoveryDocumentTypes.SunLifeExperienceAnalysis, StringComparison.OrdinalIgnoreCase)
            };
        }

        var isRenewal = isPdf && LooksLikeRenewalFileName(fileNameWithoutExtension);
        var isSeparateGir = isPdf && !isRenewal && LooksLikeGroupInformationReport(fileNameWithoutExtension, "");
        var isProposal = !isRenewal && !isSeparateGir && LooksLikeProposal(fileNameWithoutExtension, "");
        var isBilling = isPdf && !isRenewal && !isSeparateGir && LooksLikeBillingStatement(fileNameWithoutExtension, "");
        var policySchedule = FindGroupPolicyScheduleMatch(fileNameWithoutExtension, "");
        var isPolicySchedule = !isRenewal && !isSeparateGir && !isProposal && !isBilling && policySchedule.IsMatch;
        var isBooklet = !isRenewal && !isSeparateGir && !isBilling && !isPolicySchedule && LooksLikeBooklet(fileNameWithoutExtension, "");
        var isClaims = !isRenewal && !isSeparateGir && !isBilling && !isBooklet && LooksLikeClaims(fileNameWithoutExtension, "");

        return new DiscoveredDocument
        {
            FilePath = filePath,
            FileName = fileName,
            FileExtension = extension,
            DocumentType = isSeparateGir
                ? DiscoveryDocumentTypes.SeparateGroupInformationReport
                : isRenewal
                    ? DiscoveryDocumentTypes.Renewal
                : isProposal
                    ? DiscoveryDocumentTypes.Proposal
                    : isBilling
                        ? DiscoveryDocumentTypes.BillingStatement
                    : isClaims
                            ? DiscoveryDocumentTypes.ClaimsExperience
                            : isPolicySchedule
                                ? DiscoveryDocumentTypes.GroupPolicySchedule
                                : isBooklet
                                    ? DiscoveryDocumentTypes.BenefitBooklet
                                    : extension.Equals(".docx", StringComparison.OrdinalIgnoreCase)
                                    ? DiscoveryDocumentTypes.WordDocument
                                    : isPdf
                                        ? DiscoveryDocumentTypes.Other
                                        : DiscoveryDocumentTypes.Workbook,
            Carrier = "Unknown",
            IsRenewal = isRenewal,
            IsSeparateGroupInformationReport = isSeparateGir,
            IsProposal = isProposal,
            IsBillingStatement = isBilling,
            IsClaimsExperience = isClaims,
            IsBenefitBooklet = isBooklet,
            IsGroupPolicySchedule = isPolicySchedule,
            DocumentConfidenceScore = policySchedule.ConfidenceScore,
            Status = policySchedule.RequiresReview ? "needs_review" : "ready_for_extraction",
            EmbeddedPlanDesignSection = isSeparateGir ? "Group Information Report" : ""
        };
    }

    public static DiscoveredDocument ClassifyPdfPages(string pdfPath, IReadOnlyList<PdfPageText> pages)
    {
        var text = ParsingHelpers.Clean(string.Join(" ", pages.Select(page => page.Text)));
        var fileName = Path.GetFileName(pdfPath);
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(pdfPath);
        var section = FindBestPlanDesignSection(pages);
        if (RwamRateSheetPdfParser.IsRwamRateSheet(text))
        {
            var renewalDate = FindRenewalDate(text, fileNameWithoutExtension);
            var rwamDocument = new DiscoveredDocument
            {
                FilePath = pdfPath,
                FileName = fileName,
                FileExtension = Path.GetExtension(pdfPath),
                DocumentType = DiscoveryDocumentTypes.RwamRateSheet,
                Carrier = RwamRateSheetPdfParser.CarrierName,
                CarrierConfidenceScore = 0.9m,
                DocumentConfidenceScore = 1m,
                RenewalDate = renewalDate,
                RateEffectiveDate = renewalDate,
                IsRenewal = true,
                Status = "ready_for_extraction",
                EmbeddedPlanDesignSection = section.SectionType
            };
            rwamDocument.MatchedSignals.Add("RWAM rate-sheet content fingerprint");
            rwamDocument.EmbeddedPlanDesignPages.AddRange(section.PageNumbers);
            return rwamDocument;
        }

        var carrier = DetectCarrier(fileNameWithoutExtension, text);
        var groupSourceDocument = ClassifyGroupSourcePdfPages(pdfPath, pages, text, fileName);
        if (groupSourceDocument is not null)
        {
            groupSourceDocument.EmbeddedPlanDesignSection = section.SectionType;
            groupSourceDocument.EmbeddedPlanDesignPages.AddRange(section.PageNumbers);
            return groupSourceDocument;
        }

        var groupHealthDocument = ClassifyGroupHealthPdfPages(pdfPath, pages, text, fileName, carrier);
        if (groupHealthDocument is not null)
        {
            groupHealthDocument.EmbeddedPlanDesignSection = section.SectionType;
            groupHealthDocument.EmbeddedPlanDesignPages.AddRange(section.PageNumbers);
            return groupHealthDocument;
        }

        var sunLifeDocument = ClassifySunLifePdfPages(pdfPath, pages, text, fileName, fileNameWithoutExtension, carrier);
        if (sunLifeDocument is not null)
        {
            sunLifeDocument.EmbeddedPlanDesignSection = section.SectionType;
            sunLifeDocument.EmbeddedPlanDesignPages.AddRange(section.PageNumbers);
            return sunLifeDocument;
        }

        var looksLikeSeparateGirFileName = LooksLikeGroupInformationReport(fileNameWithoutExtension, "");
        var looksLikeGir = looksLikeSeparateGirFileName || LooksLikeGroupInformationReport("", text);
        var looksLikeRenewal = LooksLikeRenewal(fileNameWithoutExtension, text);
        var isSeparateGir = looksLikeSeparateGirFileName || (looksLikeGir && !looksLikeRenewal);
        var isRenewal = !isSeparateGir && looksLikeRenewal;
        var isProposal = LooksLikeProposal(fileNameWithoutExtension, text);
        var isBilling = !isRenewal && !isSeparateGir && LooksLikeBillingStatement(fileNameWithoutExtension, text);
        var policySchedule = FindGroupPolicyScheduleMatch(fileNameWithoutExtension, text);
        var hasBookletFileNameSignal = LooksLikeBooklet(fileNameWithoutExtension, "");
        var isPolicySchedule = !isRenewal && !isSeparateGir && !isProposal && !isBilling &&
                               !hasBookletFileNameSignal && policySchedule.IsMatch;
        var isBooklet = !isRenewal && !isSeparateGir && !isProposal && !isBilling &&
                        (hasBookletFileNameSignal || (!isPolicySchedule && LooksLikeBooklet("", text)));

        var document = new DiscoveredDocument
        {
            FilePath = pdfPath,
            FileName = fileName,
            FileExtension = Path.GetExtension(pdfPath),
            Carrier = isBilling && LooksLikeGroupHealthBillingStatement(fileNameWithoutExtension, text)
                ? "GroupHEALTH"
                : carrier.CarrierName,
            CarrierConfidenceScore = carrier.ConfidenceScore,
            RenewalDate = isRenewal ? FindRenewalDate(text, fileNameWithoutExtension) : null,
            IsRenewal = isRenewal,
            IsSeparateGroupInformationReport = isSeparateGir,
            IsClaimsExperience = !isRenewal && !isSeparateGir && !isProposal && !isBilling &&
                                 !isPolicySchedule && !isBooklet && LooksLikeClaims(fileNameWithoutExtension, text),
            IsBillingStatement = isBilling,
            IsBenefitBooklet = isBooklet,
            IsGroupPolicySchedule = isPolicySchedule,
            IsProposal = !isRenewal && !isSeparateGir && isProposal,
            EmbeddedPlanDesignSection = section.SectionType,
            DocumentConfidenceScore = policySchedule.ConfidenceScore,
            Status = policySchedule.RequiresReview ? "needs_review" : "ready_for_extraction"
        };
        document.EmbeddedPlanDesignPages.AddRange(section.PageNumbers);
        document.MatchedSignals.AddRange(policySchedule.MatchedSignals);

        document.DocumentType = document.IsSeparateGroupInformationReport
            ? DiscoveryDocumentTypes.SeparateGroupInformationReport
            : document.IsRenewal
            ? DiscoveryDocumentTypes.Renewal
            : document.IsBillingStatement
                ? DiscoveryDocumentTypes.BillingStatement
            : document.IsClaimsExperience
                    ? DiscoveryDocumentTypes.ClaimsExperience
                    : document.IsBenefitBooklet
                        ? DiscoveryDocumentTypes.BenefitBooklet
                        : document.IsGroupPolicySchedule
                            ? DiscoveryDocumentTypes.GroupPolicySchedule
                        : document.IsProposal
                            ? DiscoveryDocumentTypes.Proposal
                            : DiscoveryDocumentTypes.Other;

        return document;
    }

    public static PlanDesignSourceDiscovery SelectPlanDesignSource(
        IReadOnlyList<DiscoveredDocument> documents,
        DiscoveredDocument? renewalDocument)
    {
        var candidates = new List<PlanDesignSourceDiscovery>();

        foreach (var document in documents.Where(document => document.IsSeparateGroupInformationReport))
        {
            candidates.Add(BuildSource(
                document,
                PlanDesignSources.SeparateGir,
                1,
                "Group Information Report",
                document.EmbeddedPlanDesignPages,
                "Separate Group Information Report selected as the preferred plan design source."));
        }

        if (renewalDocument is not null && renewalDocument.EmbeddedPlanDesignPages.Count > 0)
        {
            candidates.Add(BuildSource(
                renewalDocument,
                PlanDesignSources.EmbeddedRenewal,
                PlanDesignPriority(renewalDocument.EmbeddedPlanDesignSection),
                renewalDocument.EmbeddedPlanDesignSection,
                renewalDocument.EmbeddedPlanDesignPages,
                $"{renewalDocument.EmbeddedPlanDesignSection} found inside the renewal PDF."));
        }

        foreach (var document in documents.Where(document => document.IsGroupPolicySchedule))
        {
            candidates.Add(BuildSource(
                document,
                PlanDesignSources.PolicySchedule,
                3,
                "Group Policy Schedule",
                document.EmbeddedPlanDesignPages,
                document.Status.Equals("needs_review", StringComparison.OrdinalIgnoreCase)
                    ? "Group policy schedule classification requires review before extraction."
                    : "Group policy schedule selected as a dedicated plan design source."));
        }

        foreach (var document in documents.Where(document => document.IsBenefitBooklet))
        {
            candidates.Add(BuildSource(
                document,
                PlanDesignSources.Booklet,
                5,
                "Benefit Booklet",
                document.EmbeddedPlanDesignPages,
                "Existing booklet selected as the plan design fallback source."));
        }

        foreach (var document in documents.Where(document => document.IsProposal))
        {
            candidates.Add(BuildSource(
                document,
                PlanDesignSources.Proposal,
                6,
                "Proposal",
                document.EmbeddedPlanDesignPages,
                "Proposal selected as the plan design fallback source."));
        }

        return candidates
            .Where(candidate => !string.IsNullOrWhiteSpace(candidate.FilePath))
            .OrderBy(candidate => candidate.Priority)
            .ThenBy(candidate => candidate.FileName, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault() ?? PlanDesignSourceDiscovery.None();
    }

    public static PlanDesignSectionMatch FindBestPlanDesignSection(IEnumerable<PdfPageText> pages)
    {
        var matches = new List<PlanDesignSectionMatch>();
        foreach (var definition in PlanDesignSections)
        {
            var pageNumbers = pages
                .Where(page => definition.Clues.Any(clue => page.Text.Contains(clue, StringComparison.OrdinalIgnoreCase)))
                .Select(page => page.PageNumber)
                .Distinct()
                .OrderBy(page => page)
                .ToList();

            if (pageNumbers.Count == 0)
            {
                continue;
            }

            matches.Add(new PlanDesignSectionMatch(definition.SectionType, definition.Priority, pageNumbers));
        }

        return matches
            .OrderBy(match => match.Priority)
            .ThenBy(match => match.PageNumbers.FirstOrDefault())
            .FirstOrDefault() ?? PlanDesignSectionMatch.None;
    }

    public static int PlanDesignPriority(string sectionType)
    {
        return PlanDesignSections
            .FirstOrDefault(section => section.SectionType.Equals(sectionType, StringComparison.OrdinalIgnoreCase))
            ?.Priority ?? 4;
    }

    private static DiscoveredDocument? SelectRenewalDocument(IReadOnlyList<DiscoveredDocument> documents)
    {
        return documents
            .Where(document => document.IsRenewal)
            .OrderByDescending(RenewalDocumentSortDate)
            .ThenByDescending(document => document.EmbeddedPlanDesignPages.Count > 0)
            .ThenBy(document => document.FileName, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();
    }

    private static DiscoveredDocument? ClassifyGroupHealthPdfPages(
        string pdfPath,
        IReadOnlyList<PdfPageText> pages,
        string text,
        string fileName,
        CarrierDetectionResult carrier)
    {
        var renewalMatches = FindSignalMatches(pages, GroupHealthRenewalSignals);
        var claimsMatches = FindSignalMatches(pages, GroupHealthClaimsSignals);
        var supportMatches = FindSignalMatches(pages, GroupHealthSupportSignals);
        var hasStrongCarrierSignal = carrier.CarrierName.Equals("GroupHEALTH", StringComparison.OrdinalIgnoreCase) ||
                                     GroupHealthPdfParser.IsGroupHealthDocument(text) ||
                                     GroupHealthPdfParser.IsHybridAsoDentalClaimsExperienceDocument(pdfPath, text) ||
                                     text.Contains("Current Underwriter SSQ FINANCIAL", StringComparison.OrdinalIgnoreCase) ||
                                     (text.Contains("EP3 Statement", StringComparison.OrdinalIgnoreCase) && text.Contains("SSQ", StringComparison.OrdinalIgnoreCase));
        var isRecognizedSupportLayout = supportMatches.Count > 0 &&
                                        (text.Contains("Policy #", StringComparison.OrdinalIgnoreCase) ||
                                         text.Contains("Policy Number", StringComparison.OrdinalIgnoreCase) ||
                                         text.Contains("Group:", StringComparison.OrdinalIgnoreCase));
        if (!hasStrongCarrierSignal && !isRecognizedSupportLayout)
        {
            return null;
        }

        string documentType;
        if (text.Contains("Group Benefits Plan Renewal", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Premium Comparison", StringComparison.OrdinalIgnoreCase))
        {
            documentType = DiscoveryDocumentTypes.GroupHealthRenewalReport;
        }
        else if (text.Contains("Top 10 DINs by Dollar Amount", StringComparison.OrdinalIgnoreCase) ||
                 text.Contains("Top 50 DINs by Dollar Amount", StringComparison.OrdinalIgnoreCase))
        {
            documentType = DiscoveryDocumentTypes.GroupHealthTopDrugs;
        }
        else if (text.Contains("EP3 Statement", StringComparison.OrdinalIgnoreCase) ||
                 text.Contains("Pooling Level for prescription drug", StringComparison.OrdinalIgnoreCase))
        {
            documentType = DiscoveryDocumentTypes.GroupHealthPoolingStatement;
        }
        else if (IsGroupHealthClaimsRangeLayout(text))
        {
            documentType = DiscoveryDocumentTypes.GroupHealthClaimsRange;
        }
        else if (text.Contains("Reserve Report", StringComparison.OrdinalIgnoreCase))
        {
            documentType = DiscoveryDocumentTypes.GroupHealthReserveReport;
        }
        else if (claimsMatches.Count > 0)
        {
            documentType = DiscoveryDocumentTypes.GroupHealthClaimsExperience;
        }
        else
        {
            return null;
        }

        var matches = renewalMatches.Concat(claimsMatches).Concat(supportMatches)
            .GroupBy(match => match.Signal, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();
        var policy = MatchValue(new Regex(@"Policy(?:\s+Number)?\s*[:#]\s*(?<value>\d+)", RegexOptions.IgnoreCase), text);
        if (string.IsNullOrWhiteSpace(policy))
        {
            policy = MatchValue(new Regex(@"Current Group\s+.+?\((?<value>\d+)\)", RegexOptions.IgnoreCase), text);
        }
        if (string.IsNullOrWhiteSpace(policy))
        {
            policy = Regex.Matches(Path.GetFileNameWithoutExtension(pdfPath), @"\d+").Select(match => match.Value).LastOrDefault() ?? "";
        }

        var holder = MatchValue(new Regex(@"Policy #\s*\d+\s*-\s*(?<value>[^\r\n]+)", RegexOptions.IgnoreCase), text);
        if (string.IsNullOrWhiteSpace(holder))
        {
            holder = MatchValue(new Regex(@"Current Group\s+(?<value>.+?)\s*\(\d+\)", RegexOptions.IgnoreCase), text);
        }

        var isRenewal = documentType.Equals(DiscoveryDocumentTypes.GroupHealthRenewalReport, StringComparison.OrdinalIgnoreCase);
        var isClaims = documentType.Equals(DiscoveryDocumentTypes.GroupHealthClaimsExperience, StringComparison.OrdinalIgnoreCase) ||
                       documentType.Equals(DiscoveryDocumentTypes.GroupHealthClaimsRange, StringComparison.OrdinalIgnoreCase);
        var document = new DiscoveredDocument
        {
            FilePath = pdfPath,
            FileName = fileName,
            FileExtension = Path.GetExtension(pdfPath),
            DocumentType = documentType,
            Carrier = "GroupHEALTH",
            CarrierConfidenceScore = hasStrongCarrierSignal ? 0.95m : 0.75m,
            DocumentConfidenceScore = decimal.Round(Math.Min(1m, 0.55m + matches.Count * 0.08m), 2),
            ContractNumber = policy,
            ContractHolder = holder,
            RenewalDate = isRenewal ? FindRenewalDate(text, Path.GetFileNameWithoutExtension(pdfPath)) : null,
            IsRenewal = isRenewal,
            IsClaimsExperience = isClaims,
            Status = "ready_for_extraction"
        };
        document.MatchedSignals.AddRange(matches.Select(match => match.Signal));
        document.PagesDetected.AddRange(matches.SelectMany(match => match.Pages).Distinct().OrderBy(page => page));
        return document;
    }

    private static DiscoveredDocument? ClassifyGroupSourcePdfPages(
        string pdfPath,
        IReadOnlyList<PdfPageText> pages,
        string text,
        string fileName)
    {
        if (!GroupSourcePdfParser.IsGroupSourceDocument(pdfPath, text))
        {
            return null;
        }

        var renewal = text.Contains("The following rates will be reflected", StringComparison.OrdinalIgnoreCase) ||
                      text.Contains("EXPERIENCE CALCULATION", StringComparison.OrdinalIgnoreCase);
        var monthly = text.Contains("Billed Premium & Claims", StringComparison.OrdinalIgnoreCase);
        var topDrugs = text.Contains("Top 10 DINs by Dollar Amount", StringComparison.OrdinalIgnoreCase);
        var detail = ContainsAny(text, "Expense Information", "Paramedical Information") ||
                     GroupSourcePdfParser.IsCostCategoryReport(text);
        var documentType = renewal ? DiscoveryDocumentTypes.GroupSourceRenewalReport :
            monthly ? DiscoveryDocumentTypes.GroupSourceClaimsExperience :
            topDrugs ? DiscoveryDocumentTypes.GroupSourceTopDrugs :
            detail ? DiscoveryDocumentTypes.GroupSourceClaimsDetail : DiscoveryDocumentTypes.Other;

        var parser = new GroupSourcePdfParser().Parse(pdfPath, pages.ToList(), text);
        var matched = new List<string>();
        if (renewal) matched.Add("GroupSource rate and experience presentation");
        if (monthly) matched.Add("GroupSource billed premium and claims");
        if (detail) matched.Add("GroupSource claims category detail");
        if (topDrugs) matched.Add("GroupSource Top DIN report");

        var document = new DiscoveredDocument
        {
            FilePath = pdfPath,
            FileName = fileName,
            FileExtension = Path.GetExtension(pdfPath),
            DocumentType = documentType,
            Carrier = "GroupSource",
            CarrierConfidenceScore = 0.98m,
            DocumentConfidenceScore = renewal || monthly ? 0.98m : 0.92m,
            ContractNumber = parser.Pdf.PolicyNumber,
            ContractHolder = parser.Pdf.ClientName,
            RenewalDate = parser.Pdf.RenewalDate,
            IsRenewal = renewal,
            IsClaimsExperience = monthly,
            Status = "ready_for_extraction"
        };
        document.MatchedSignals.AddRange(matched);
        document.PagesDetected.AddRange(pages
            .Where(page => matched.Any(signal =>
                signal.Contains("rate", StringComparison.OrdinalIgnoreCase) && page.Text.Contains("rates will be reflected", StringComparison.OrdinalIgnoreCase) ||
                signal.Contains("billed", StringComparison.OrdinalIgnoreCase) && page.Text.Contains("Billed Premium & Claims", StringComparison.OrdinalIgnoreCase) ||
                signal.Contains("category", StringComparison.OrdinalIgnoreCase) &&
                    (ContainsAny(page.Text, "Expense Information", "Paramedical Information") || GroupSourcePdfParser.IsCostCategoryReport(page.Text)) ||
                signal.Contains("DIN", StringComparison.OrdinalIgnoreCase) && page.Text.Contains("Top 10 DINs by Dollar Amount", StringComparison.OrdinalIgnoreCase)))
            .Select(page => page.PageNumber)
            .Distinct()
            .OrderBy(page => page));
        return document;
    }

    private static bool IsGroupHealthClaimsRangeLayout(string text)
    {
        var hasSupportedTitle =
            text.Contains("Extended Health Care Claims Range by Certificate", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Dental Care Claims Range by Certificate", StringComparison.OrdinalIgnoreCase);
        return hasSupportedTitle &&
               Regex.IsMatch(text, @"\bGroup\s*:\s*\S+", RegexOptions.IgnoreCase) &&
               Regex.IsMatch(text, @"\bPolicy\s*:\s*\d+", RegexOptions.IgnoreCase) &&
               Regex.IsMatch(text, @"\bStart\s+Date\s*:\s*\S+", RegexOptions.IgnoreCase) &&
               Regex.IsMatch(text, @"\bEnd\s+Date\s*:\s*\S+", RegexOptions.IgnoreCase);
    }

    private static DiscoveredDocument? ClassifySunLifePdfPages(
        string pdfPath,
        IReadOnlyList<PdfPageText> pages,
        string text,
        string fileName,
        string fileNameWithoutExtension,
        CarrierDetectionResult carrier)
    {
        var isSunLife = carrier.CarrierName.Equals("Sun Life", StringComparison.OrdinalIgnoreCase) ||
                        text.Contains("Sun Life", StringComparison.OrdinalIgnoreCase) ||
                        fileNameWithoutExtension.Contains("SunLife", StringComparison.OrdinalIgnoreCase) ||
                        fileNameWithoutExtension.Contains("Sun Life", StringComparison.OrdinalIgnoreCase);
        if (!isSunLife)
        {
            return null;
        }

        var renewalMatches = FindSignalMatches(pages, SunLifeRenewalReportSignals);
        var experienceMatches = FindSignalMatches(pages, SunLifeExperienceAnalysisSignals);
        var fileNameDocumentType = ClassifySunLifeFileName(fileNameWithoutExtension);
        if (renewalMatches.Count == 0 && experienceMatches.Count == 0 && string.IsNullOrWhiteSpace(fileNameDocumentType))
        {
            return null;
        }

        var isExperience = experienceMatches.Count > renewalMatches.Count ||
                           fileNameDocumentType.Equals(DiscoveryDocumentTypes.SunLifeExperienceAnalysis, StringComparison.OrdinalIgnoreCase);
        var documentType = isExperience
            ? DiscoveryDocumentTypes.SunLifeExperienceAnalysis
            : DiscoveryDocumentTypes.SunLifeRenewalReport;
        var matchedSignals = (isExperience ? experienceMatches : renewalMatches)
            .Concat(renewalMatches)
            .Concat(experienceMatches)
            .GroupBy(match => match.Signal, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();

        var confidenceBase = isExperience ? SunLifeExperienceAnalysisSignals.Length : SunLifeRenewalReportSignals.Length;
        var confidence = decimal.Round(Math.Min(1m, (matchedSignals.Count + 1m) / Math.Max(4, confidenceBase)), 2);
        var effectiveDate = FindSunLifeRateEffectiveDate(text);

        var document = new DiscoveredDocument
        {
            FilePath = pdfPath,
            FileName = fileName,
            FileExtension = Path.GetExtension(pdfPath),
            DocumentType = documentType,
            Carrier = "Sun Life",
            CarrierConfidenceScore = Math.Max(carrier.ConfidenceScore, 0.75m),
            DocumentConfidenceScore = confidence,
            ContractNumber = MatchValue(SunLifeContractNumberRegex, text),
            ContractHolder = FindSunLifeContractHolder(pages),
            RateEffectiveDate = effectiveDate,
            RenewalDate = effectiveDate,
            IsRenewal = documentType.Equals(DiscoveryDocumentTypes.SunLifeRenewalReport, StringComparison.OrdinalIgnoreCase),
            IsClaimsExperience = documentType.Equals(DiscoveryDocumentTypes.SunLifeExperienceAnalysis, StringComparison.OrdinalIgnoreCase),
            Status = "ready_for_extraction"
        };

        document.MatchedSignals.AddRange(matchedSignals.Select(match => match.Signal));
        document.PagesDetected.AddRange(matchedSignals.SelectMany(match => match.Pages).Distinct().OrderBy(page => page));
        return document;
    }

    private static string ClassifySunLifeFileName(string fileNameWithoutExtension)
    {
        if (!fileNameWithoutExtension.Contains("SunLife", StringComparison.OrdinalIgnoreCase) &&
            !fileNameWithoutExtension.Contains("Sun Life", StringComparison.OrdinalIgnoreCase))
        {
            return "";
        }

        if (ContainsAny(fileNameWithoutExtension, "experience", "exp"))
        {
            return DiscoveryDocumentTypes.SunLifeExperienceAnalysis;
        }

        if (ContainsAny(fileNameWithoutExtension, "renewal", "renewal report"))
        {
            return DiscoveryDocumentTypes.SunLifeRenewalReport;
        }

        return "";
    }

    private static List<SunLifeSignalMatch> FindSignalMatches(IReadOnlyList<PdfPageText> pages, IReadOnlyList<string> signals)
    {
        var matches = new List<SunLifeSignalMatch>();
        foreach (var signal in signals)
        {
            var pagesDetected = pages
                .Where(page => page.Text.Contains(signal, StringComparison.OrdinalIgnoreCase))
                .Select(page => page.PageNumber)
                .Distinct()
                .OrderBy(page => page)
                .ToList();
            if (pagesDetected.Count > 0)
            {
                matches.Add(new SunLifeSignalMatch(signal, pagesDetected));
            }
        }

        return matches;
    }

    private static DateOnly? FindSunLifeRateEffectiveDate(string text)
    {
        var match = SunLifeRateEffectiveDateRegex.Match(text);
        return match.Success && TryParseDate(match.Groups["date"].Value, out var date) ? date : null;
    }

    private static string FindSunLifeContractHolder(IEnumerable<PdfPageText> pages)
    {
        foreach (var line in pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean))
        {
            var match = SunLifeContractHolderRegex.Match(line);
            if (!match.Success)
            {
                continue;
            }

            var value = ParsingHelpers.Clean(match.Groups["value"].Value);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return "";
    }

    private static string MatchValue(Regex regex, string text)
    {
        var match = regex.Match(text);
        return match.Success ? ParsingHelpers.Clean(match.Groups["value"].Value) : "";
    }

    private static PlanDesignSourceDiscovery BuildSource(
        DiscoveredDocument document,
        string source,
        int priority,
        string sectionType,
        IEnumerable<int> pageNumbers,
        string detail)
    {
        var result = new PlanDesignSourceDiscovery
        {
            Source = source,
            Status = PlanDesignStatuses.Available,
            FilePath = document.FilePath,
            FileName = document.FileName,
            DocumentType = document.DocumentType,
            SectionType = sectionType,
            Detail = detail,
            Priority = priority
        };
        result.PageNumbers.AddRange(pageNumbers.Distinct().OrderBy(page => page));
        return result;
    }

    private static bool LooksLikeRenewal(string fileName, string text)
    {
        var searchable = $"{fileName} {text}";
        if (ContainsAny(searchable,
                "Summary of Renewal Rates and Premium",
                "Monthly Premium Comparison",
                "Total monthly cost summary",
                "The following rates will be reflected",
                "Renewal calculation",
                "Renewal Report"))
        {
            return true;
        }

        return searchable.Contains("renewal", StringComparison.OrdinalIgnoreCase) &&
               ContainsAny(text, "rate change", "renewal rate", "renewal premium", "experience analysis", "paid loss ratio");
    }

    private static bool LooksLikeRenewalFileName(string fileName)
    {
        return fileName.Contains("renewal", StringComparison.OrdinalIgnoreCase) &&
               !LooksLikeGroupInformationReport(fileName, "") &&
               !LooksLikeBooklet(fileName, "");
    }

    private static bool LooksLikeGroupInformationReport(string fileName, string text)
    {
        var searchable = $"{fileName} {text}";
        return GroupInformationReportRegex.IsMatch(searchable) ||
               GirTokenRegex.IsMatch(fileName);
    }

    private static bool LooksLikeBooklet(string fileName, string text)
    {
        return ContainsAny(
            $"{fileName} {text}",
            "benefit booklet",
            "booklet",
            "bookelt",
            "master policy",
            "policy specifications",
            "Schedule of Benefits",
            "Benefit Summary",
            "Coverage Summary",
            "Benefit Details",
            "Class Information",
            "Plan Design",
            "Health Care Spending Account",
            "Lifestyle Spending Account") ||
            Regex.IsMatch(fileName, @"(?<![A-Za-z0-9])(?:HCSA|LSA)(?![A-Za-z0-9])", RegexOptions.IgnoreCase);
    }

    private static GroupPolicyScheduleMatch FindGroupPolicyScheduleMatch(string fileName, string text)
    {
        var searchable = $"{fileName} {text}";
        var titles = GroupPolicyScheduleTitleSignals
            .Where(signal => searchable.Contains(signal, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var details = GroupPolicyScheduleDetailSignals
            .Where(signal => searchable.Contains(signal, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var coverage = GroupPolicyScheduleCoverageSignals
            .Where(signal => searchable.Contains(signal, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var matched = titles.Concat(details).Concat(coverage).ToList();

        var isMatch = titles.Count > 0 && details.Count >= 3 && coverage.Count >= 1;
        var requiresReview = !isMatch &&
                             ((titles.Count > 0 && details.Count >= 2) ||
                              (details.Count >= 5 && coverage.Count >= 1));
        var confidence = isMatch
            ? decimal.Round(Math.Min(0.98m, 0.70m + details.Count * 0.04m + coverage.Count * 0.03m), 2)
            : requiresReview ? 0.45m : 0m;
        return new GroupPolicyScheduleMatch(isMatch, requiresReview, confidence, matched);
    }

    private static CarrierDetectionResult DetectCarrier(string fileName, string text)
    {
        if (EmpireLifePdfParser.IsEmpireLifeDocument(fileName, text))
        {
            return new CarrierDetectionResult
            {
                CarrierName = "Empire Life",
                ConfidenceScore = 0.95m
            };
        }

        var carrierFromFileName = GenericRenewalDocumentAnalyzer.DetectCarrier(fileName);
        return carrierFromFileName.CarrierName.Equals("Unknown", StringComparison.OrdinalIgnoreCase)
            ? GenericRenewalDocumentAnalyzer.DetectCarrier(text)
            : carrierFromFileName;
    }

    private static bool LooksLikeProposal(string fileName, string text)
    {
        return ContainsAny($"{fileName} {text}", "proposal", "quote", "quotation", "proposed benefits");
    }

    private static bool LooksLikeClaims(string fileName, string text)
    {
        return ContainsAny(
            $"{fileName} {text}",
            "claims",
            "experience",
            "premium claims",
            "paid claims",
            "loss ratio",
            "certificates premium claims",
            "claims distribution");
    }

    private static bool LooksLikeBillingStatement(string fileName, string text)
    {
        var searchable = $"{fileName} {text}";
        return ContainsAny(
                   searchable,
                   "benefits billing statement",
                   "benefit billing statement",
                   "summary by billing division",
                   "statement reconciliation",
                   "billing period:",
                   "billing id") ||
               fileName.Contains("_bill", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("billing", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeGroupHealthBillingStatement(string fileName, string text)
    {
        return ContainsAny(
            $"{fileName} {text}",
            "clienthelp.grouphealth.ca",
            "grouphealth",
            "community services benefits trust");
    }

    private static bool ContainsAny(string text, params string[] values)
    {
        return values.Any(value => text.Contains(value, StringComparison.OrdinalIgnoreCase));
    }

    private static DateOnly RenewalDocumentSortDate(DiscoveredDocument document)
    {
        return document.RenewalDate ??
               InferRenewalDateFromFileName(document.FileName) ??
               DateOnly.MinValue;
    }

    private static DateOnly? FindRenewalDate(string text, string fileName)
    {
        var match = RenewalDateRegex.Match(text);
        if (match.Success && TryParseDate(match.Groups["date"].Value, out var date))
        {
            return date;
        }

        return InferRenewalDateFromFileName(fileName);
    }

    private static DateOnly? InferRenewalDateFromFileName(string fileName)
    {
        var match = FileNameYearRegex.Match(fileName);
        return match.Success && int.TryParse(match.Groups["year"].Value, out var year)
            ? new DateOnly(year, 1, 1)
            : null;
    }

    private static bool TryParseDate(string value, out DateOnly date)
    {
        var cleaned = value.Replace(",", "", StringComparison.Ordinal).Replace(".", "", StringComparison.Ordinal);
        var formats = new[]
        {
            "MMMM d yyyy",
            "MMM d yyyy",
            "d-MMMM-yyyy",
            "d-MMM-yyyy",
            "d/MMMM/yyyy",
            "d/MMM/yyyy",
            "M/d/yyyy",
            "MM/dd/yyyy",
            "yyyy-M-d",
            "yyyy-MM-dd"
        };

        foreach (var format in formats)
        {
            if (DateOnly.TryParseExact(cleaned, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                return true;
            }
        }

        return DateOnly.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
    }

    private static bool IsGeneratedOutputPath(string clientFolder, string filePath)
    {
        var relativePath = Path.GetRelativePath(clientFolder, filePath);
        var parts = relativePath.Split(
            [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
            StringSplitOptions.RemoveEmptyEntries);

        return parts.Any(part =>
            part.Equals("Batch Output", StringComparison.OrdinalIgnoreCase) ||
            part.Equals("Supporting PDFs", StringComparison.OrdinalIgnoreCase) ||
            part.Equals("Supporting Excel Files", StringComparison.OrdinalIgnoreCase));
    }

    private sealed record PlanDesignSectionDefinition(string SectionType, int Priority, string[] Clues);
}

public sealed record PlanDesignSectionMatch(string SectionType, int Priority, IReadOnlyList<int> PageNumbers)
{
    public static PlanDesignSectionMatch None { get; } = new("", int.MaxValue, []);
}

internal sealed record SunLifeSignalMatch(string Signal, IReadOnlyList<int> Pages);
internal sealed record GroupPolicyScheduleMatch(bool IsMatch, bool RequiresReview, decimal ConfidenceScore, IReadOnlyList<string> MatchedSignals);
