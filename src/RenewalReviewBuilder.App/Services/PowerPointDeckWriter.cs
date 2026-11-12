using System.IO.Compression;
using System.Globalization;
using System.Security;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using RenewalReviewBuilder.App.Models;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Presentation = DocumentFormat.OpenXml.Presentation;

namespace RenewalReviewBuilder.App.Services;

public sealed class PowerPointDeckWriter
{
    public sealed class DeckOutput
    {
        public string? Path { get; set; }
        public string Message { get; set; } = "";
    }

    public DeckOutput Write(
        AnalysisResult analysis,
        string outputFolder,
        string? templatePath,
        IReadOnlyCollection<UnderwritingReferenceRate>? underwritingReferenceRates = null)
    {
        var presentationErrors = ClientPresentationRateSelector.ValidateSelectedVersion(analysis);
        if (presentationErrors.Count > 0)
        {
            throw new InvalidOperationException(
                "PowerPoint client-presentation rates are incomplete: " +
                string.Join(" ", presentationErrors));
        }

        Directory.CreateDirectory(outputFolder);
        var outputPath = Path.Combine(outputFolder, "Client_Renewal_Draft.pptx");
        DeckOutput? templateResult = null;

        if (!string.IsNullOrWhiteSpace(templatePath) && File.Exists(templatePath))
        {
            templateResult = TryWriteFromTemplate(
                analysis,
                templatePath,
                outputPath,
                underwritingReferenceRates ?? UnderwritingReferenceRateCatalog.Defaults);
            if (templateResult.Path is not null)
            {
                return templateResult;
            }
        }

        WriteGeneratedDeck(analysis, outputPath);
        var message = string.IsNullOrWhiteSpace(templatePath)
            ? "No PowerPoint template was selected, so a clean draft deck was generated."
            : $"The PowerPoint template was locked, unreadable, or did not contain supported placeholders, so a clean draft deck was generated.{(string.IsNullOrWhiteSpace(templateResult?.Message) ? "" : $" Detail: {templateResult.Message}")}";

        return new DeckOutput { Path = outputPath, Message = message };
    }

    private static DeckOutput TryWriteFromTemplate(
        AnalysisResult analysis,
        string templatePath,
        string outputPath,
        IReadOnlyCollection<UnderwritingReferenceRate> underwritingReferenceRates)
    {
        try
        {
            var replacements = BuildTokenMapWithReferenceRates(analysis, underwritingReferenceRates);
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            File.Copy(templatePath, outputPath, overwrite: true);
            var tokenFound = ReplaceTemplateTokens(
                outputPath,
                replacements,
                analysis,
                underwritingReferenceRates);
            var tableFound = ReplaceTemplateTables(outputPath, analysis);
            var demographicChartsUpdated = TryPopulateDemographicCharts(outputPath, analysis);
            var activityChartsUpdated = TryPopulateActivityCharts(outputPath, analysis);
            var spendingTrendChartsUpdated = TryPopulateEhcSpendingTrendChart(outputPath, analysis);
            var dentalSpendingTrendChartsUpdated = TryPopulateDentalSpendingTrendChart(outputPath, analysis);
            var paramedicalServiceTrendChartsUpdated = TryPopulateParamedicalServiceTrendChart(outputPath, analysis);
            var topDrugsUpdated = TryPopulateTopDrugSlides(outputPath, analysis);
            var programClaimsSummaryUpdated = TryPopulateProgramPerformanceClaimsSummaryChart(outputPath, analysis);
            var nemMaximumsUpdated = TryPopulateNemMaximumsSlide(outputPath, analysis);
            var assurePlusExperienceUpdated = TryPopulateAssurePlusExperienceSlide(outputPath, analysis);
            var claimsRangeResult = new ClaimsRangePowerPointUpdater().Update(outputPath, analysis.ClaimsRange);
            var claimsRangeHandled = claimsRangeResult.Status != ClaimsRangePowerPointUpdateStatus.SlideNotFound;

            if (!tokenFound && !tableFound && !demographicChartsUpdated && !activityChartsUpdated && !spendingTrendChartsUpdated && !dentalSpendingTrendChartsUpdated && !paramedicalServiceTrendChartsUpdated && !topDrugsUpdated && !programClaimsSummaryUpdated && !nemMaximumsUpdated && !assurePlusExperienceUpdated && !claimsRangeHandled)
            {
                return new DeckOutput
                {
                    Message = "Template was readable but no supported placeholders or chart data targets were found."
                };
            }

            var messageParts = new List<string>();
            if (tokenFound)
            {
                messageParts.Add("mapped template placeholders");
            }

            if (tableFound)
            {
                messageParts.Add("rate table placeholders");
            }

            if (demographicChartsUpdated)
            {
                messageParts.Add("demographic chart data");
            }

            if (activityChartsUpdated)
            {
                messageParts.Add("EHC and dental activity chart data");
            }

            if (spendingTrendChartsUpdated)
            {
                messageParts.Add("EHC spending trend chart data");
            }

            if (dentalSpendingTrendChartsUpdated)
            {
                messageParts.Add("Dental spending trend chart data");
            }

            if (paramedicalServiceTrendChartsUpdated)
            {
                messageParts.Add("paramedical service trend chart data");
            }

            if (topDrugsUpdated)
            {
                messageParts.Add("top drug table");
            }

            if (programClaimsSummaryUpdated)
            {
                messageParts.Add("program claims summary chart");
            }

            if (nemMaximumsUpdated)
            {
                messageParts.Add("NEM maximums table");
            }

            if (assurePlusExperienceUpdated)
            {
                messageParts.Add("AssurePLUS experience history and renewal outcome");
            }

            if (claimsRangeHandled)
            {
                messageParts.Add(claimsRangeResult.Status == ClaimsRangePowerPointUpdateStatus.Updated
                    ? "Claims Range native chart data"
                    : $"Claims Range slide suppression ({claimsRangeResult.Message})");
            }

            var message = $"PowerPoint draft was created from the {string.Join(", ", messageParts)}.";

            return new DeckOutput
            {
                Path = outputPath,
                Message = message
            };
        }
        catch (Exception ex)
        {
            return new DeckOutput { Message = ex.Message };
        }
    }

    private static bool TryPopulateAssurePlusExperienceSlide(string outputPath, AnalysisResult analysis)
    {
        using var presentation = PresentationDocument.Open(outputPath, true);
        var presentationPart = presentation.PresentationPart;
        if (presentationPart is null)
        {
            return false;
        }

        foreach (var slidePart in presentationPart.SlideParts)
        {
            var shapes = slidePart.Slide?.Descendants<Presentation.Shape>()
                .Where(shape => !string.IsNullOrWhiteSpace(GetShapeName(shape)))
                .GroupBy(shape => GetShapeName(shape), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            if (shapes is null || !shapes.ContainsKey("result-band"))
            {
                continue;
            }

            var values = BuildAssurePlusShapeValues(analysis.AssurePlusExperience);
            if (!values.Keys.All(shapes.ContainsKey))
            {
                continue;
            }

            foreach (var pair in values)
            {
                if (shapes.TryGetValue(pair.Key, out var shape))
                {
                    ReplaceShapeText(shape, pair.Value);
                }
            }

            slidePart.Slide!.Save();
            return true;
        }

        return false;
    }

    private static Dictionary<string, string> BuildAssurePlusShapeValues(AssurePlusExperienceSummary? summary)
    {
        if (summary is null)
        {
            return BuildUnavailableAssurePlusShapeValues(
                "AssurePLUS experience unavailable",
                "No AssurePLUS Health and Dental pricing page was extracted.");
        }

        if (!IsClientReadyAssurePlusSummary(summary))
        {
            return BuildUnavailableAssurePlusShapeValues(
                "AssurePLUS experience requires review",
                "The three-year history is incomplete or contains conflicting source values.");
        }

        var hasHealth = summary.HealthOutcome is not null;
        var hasDental = summary.DentalOutcome is not null;
        var title = hasDental ? summary.DentalCommentary : summary.HealthCommentary;
        var subtitle = hasHealth && hasDental
            ? summary.HealthCommentary
            : hasHealth
                ? "Dental coverage is not included in this renewal."
                : "Health coverage is not included in this renewal.";
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["eyebrow"] = "ASSUREPLUS EXPERIENCE",
            ["title"] = title,
            ["subtitle"] = subtitle,
            ["result-title"] = summary.EffectiveDate.HasValue
                ? $"{summary.EffectiveDate.Value:MMMM yyyy} renewal outcome"
                : "Current renewal outcome",
            ["result-subtitle"] = "Your experience directly influenced the final rate change.",
            ["health-outcome-label"] = "HEALTH",
            ["health-outcome"] = FormatAssurePlusPercent(summary.HealthOutcome?.TotalChangePercent),
            ["health-outcome-note"] = AssurePlusOutcomeNote(summary.HealthOutcome),
            ["dental-outcome-label"] = "DENTAL",
            ["dental-outcome"] = FormatAssurePlusPercent(summary.DentalOutcome?.TotalChangePercent),
            ["dental-outcome-note"] = AssurePlusOutcomeNote(summary.DentalOutcome),
            ["takeaway"] = summary.Takeaway
        };

        var years = summary.AnnualRatings.Select(row => row.Year).Distinct().OrderBy(year => year).Take(3).ToList();
        for (var index = 0; index < 3; index++)
        {
            if (index >= years.Count)
            {
                values[$"year-{index}"] = "";
                values[$"status-text-0-{index}"] = "";
                values[$"status-text-1-{index}"] = "";
                continue;
            }

            var year = years[index];
            values[$"year-{index}"] = year.ToString(CultureInfo.InvariantCulture);
            values[$"status-text-0-{index}"] = AssurePlusRating(summary, "Health", year);
            values[$"status-text-1-{index}"] = AssurePlusRating(summary, "Dental", year);
        }

        return values;
    }

    private static bool IsClientReadyAssurePlusSummary(AssurePlusExperienceSummary summary)
    {
        var includedBenefits = new[]
            {
                (Name: "Health", Outcome: summary.HealthOutcome),
                (Name: "Dental", Outcome: summary.DentalOutcome)
            }
            .Where(item => item.Outcome is not null)
            .ToList();
        if (includedBenefits.Count == 0 || includedBenefits.Any(item => !item.Outcome!.TotalChangePercent.HasValue))
        {
            return false;
        }

        var years = summary.AnnualRatings.Select(row => row.Year).Distinct().OrderBy(year => year).ToList();
        return years.Count == 3 &&
               years[1] == years[0] + 1 &&
               years[2] == years[1] + 1 &&
               includedBenefits.All(benefit =>
                   summary.AnnualRatings.Count(row =>
                       row.Benefit.Equals(benefit.Name, StringComparison.OrdinalIgnoreCase) &&
                       !string.IsNullOrWhiteSpace(row.Rating) &&
                       !row.Conflict) == 3);
    }

    private static Dictionary<string, string> BuildUnavailableAssurePlusShapeValues(string title, string subtitle)
    {
        var unavailable = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["eyebrow"] = "ASSUREPLUS EXPERIENCE",
            ["title"] = "",
            ["subtitle"] = "",
            ["result-title"] = title,
            ["result-subtitle"] = subtitle,
            ["health-outcome-label"] = "HEALTH",
            ["health-outcome"] = "",
            ["health-outcome-note"] = "",
            ["dental-outcome-label"] = "DENTAL",
            ["dental-outcome"] = "",
            ["dental-outcome-note"] = "",
            ["takeaway"] = ""
        };
        for (var index = 0; index < 3; index++)
        {
            unavailable[$"year-{index}"] = "";
            unavailable[$"status-text-0-{index}"] = "";
            unavailable[$"status-text-1-{index}"] = "";
        }
        return unavailable;
    }

    private static string AssurePlusRating(AssurePlusExperienceSummary summary, string benefit, int year) =>
        summary.AnnualRatings.FirstOrDefault(row =>
            row.Year == year && row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))?.Rating ?? "";

    private static string AssurePlusOutcomeNote(AssurePlusBenefitPricing? outcome)
    {
        if (outcome is null)
        {
            return "";
        }

        if (outcome.ExperienceAdjustmentPercent is null or 0m)
        {
            return "Standard change; no experience adjustment";
        }

        var adjustment = Math.Abs(outcome.ExperienceAdjustmentPercent.Value);
        return $"Includes a {adjustment.ToString(adjustment % 1m == 0m ? "0" : "0.##", CultureInfo.InvariantCulture)}% experience adjustment";
    }

    private static string FormatAssurePlusPercent(decimal? value)
    {
        if (!value.HasValue)
        {
            return "";
        }

        var prefix = value.Value > 0m ? "+" : value.Value < 0m ? "−" : "";
        return $"{prefix}{Math.Abs(value.Value).ToString("0.00", CultureInfo.InvariantCulture)}%";
    }

    private static string GetShapeName(Presentation.Shape shape) =>
        shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "";

    private static void ReplaceShapeText(Presentation.Shape shape, string value)
    {
        var textNodes = shape.TextBody?.Descendants<Drawing.Text>().ToList() ?? [];
        if (textNodes.Count == 0)
        {
            return;
        }

        textNodes[0].Text = value;
        foreach (var textNode in textNodes.Skip(1))
        {
            textNode.Text = "";
        }
    }

    private static bool ReplaceTemplateTokens(
        string outputPath,
        Dictionary<string, string> replacements,
        AnalysisResult analysis,
        IReadOnlyCollection<UnderwritingReferenceRate> underwritingReferenceRates)
    {
        var tokenFound = false;
        using var presentation = PresentationDocument.Open(outputPath, true);
        var presentationPart = presentation.PresentationPart;
        if (presentationPart is null)
        {
            return false;
        }

        foreach (var slidePart in presentationPart.SlideParts)
        {
            if (slidePart.Slide is null)
            {
                continue;
            }

            var paragraphs = slidePart.Slide.Descendants<Drawing.Paragraph>().ToList();
            var slideText = ExtractParagraphText(paragraphs);
            var normalizedSlideText = NormalizeSlideLookupText(slideText);
            if (IsLifeRenewalSummarySlide(normalizedSlideText) &&
                FindRateLines(analysis, "Dependent Life").Count == 0)
            {
                tokenFound |= RemoveAbsentDependentLifeNarrative(paragraphs);
                paragraphs = slidePart.Slide.Descendants<Drawing.Paragraph>().ToList();
                slideText = ExtractParagraphText(paragraphs);
            }

            var slideReplacements = BuildSlideTokenMap(
                replacements,
                slideText,
                analysis,
                underwritingReferenceRates);
            tokenFound |= ReplaceParagraphTokens(paragraphs, slideReplacements);
            slidePart.Slide.Save();
        }

        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            if (masterPart.SlideMaster is null)
            {
                continue;
            }

            tokenFound |= ReplaceParagraphTokens(masterPart.SlideMaster.Descendants<Drawing.Paragraph>(), replacements);
            masterPart.SlideMaster.Save();

            foreach (var layoutPart in masterPart.SlideLayoutParts)
            {
                if (layoutPart.SlideLayout is null)
                {
                    continue;
                }

                tokenFound |= ReplaceParagraphTokens(layoutPart.SlideLayout.Descendants<Drawing.Paragraph>(), replacements);
                layoutPart.SlideLayout.Save();
            }
        }

        return tokenFound;
    }

    private static Dictionary<string, string> BuildSlideTokenMap(
        Dictionary<string, string> replacements,
        string slideText,
        AnalysisResult analysis,
        IReadOnlyCollection<UnderwritingReferenceRate> underwritingReferenceRates)
    {
        var slideReplacements = new Dictionary<string, string>(replacements, StringComparer.Ordinal);
        var normalizedSlideText = NormalizeSlideLookupText(slideText);

        if (normalizedSlideText.Contains("LIFE INSURANCE", StringComparison.OrdinalIgnoreCase) &&
            normalizedSlideText.Contains("DEMOGRAPHIC ANALYSIS", StringComparison.OrdinalIgnoreCase))
        {
            AddGenericDemographicAnalysisTokens(
                slideReplacements,
                analysis.Pdf.LifeDemographics,
                underwritingReferenceRates,
                UnderwritingRateTypes.Mortality);
            slideReplacements["members insured under the Long-Term Disability benefit."] = "members insured under the Life Insurance benefit.";
        }
        else if (normalizedSlideText.Contains("LONG TERM DISABILITY DEMOGRAPHIC ANALYSIS", StringComparison.OrdinalIgnoreCase) ||
                 normalizedSlideText.Contains("LTD DEMOGRAPHIC ANALYSIS", StringComparison.OrdinalIgnoreCase))
        {
            AddGenericDemographicAnalysisTokens(
                slideReplacements,
                analysis.Pdf.LtdDemographics,
                underwritingReferenceRates,
                UnderwritingRateTypes.Morbidity);
        }

        if (normalizedSlideText.Contains("EXTENDED HEALTH CARE RENEWAL SUMMARY", StringComparison.OrdinalIgnoreCase) ||
            normalizedSlideText.Contains("PARAMEDICAL AND PRACTITIONER TRENDS", StringComparison.OrdinalIgnoreCase))
        {
            AddParamedicalSummaryTokens(slideReplacements, analysis, includeLiteralTemplateText: true);
        }

        return slideReplacements;
    }

    private static bool IsLifeRenewalSummarySlide(string normalizedSlideText) =>
        normalizedSlideText.Contains("LIFE INSURANCE", StringComparison.OrdinalIgnoreCase) &&
        normalizedSlideText.Contains("RENEWAL SUMMARY", StringComparison.OrdinalIgnoreCase);

    private static bool RemoveAbsentDependentLifeNarrative(IEnumerable<Drawing.Paragraph> paragraphs)
    {
        var removed = false;
        foreach (var paragraph in paragraphs.ToList())
        {
            var paragraphText = string.Concat(paragraph.Descendants<Drawing.Text>().Select(text => text.Text));
            if (!paragraphText.Contains("DependentLife_CurrentRate", StringComparison.OrdinalIgnoreCase) &&
                !paragraphText.Contains("DependentLife_RenewalRate", StringComparison.OrdinalIgnoreCase) &&
                !paragraphText.Contains("Add dependent life rate adjustment", StringComparison.OrdinalIgnoreCase) &&
                !paragraphText.Contains("Dependent Life rate", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            paragraph.Remove();
            removed = true;
        }

        return removed;
    }

    private static void AddParamedicalSummaryTokens(
        Dictionary<string, string> tokens,
        AnalysisResult analysis,
        bool includeLiteralTemplateText = false)
    {
        var shares = BuildParamedicalSharePercentages(analysis);
        AddParamedicalShareToken(tokens, shares, "Paramedical_MassagePercent", ["Massage"], "Massage Therapist (xx%)", "Massage Therapist", includeLiteralTemplateText);
        AddParamedicalShareToken(tokens, shares, "Paramedical_PhysiotherapyPercent", ["Physiotherapy"], "Physiotherapist (xx%)", "Physiotherapist", includeLiteralTemplateText);
        AddParamedicalShareToken(tokens, shares, "Paramedical_AcupuncturePercent", ["Acupuncture", "Acupuncturist"], "Acupuncturist (xx%)", "Acupuncturist", includeLiteralTemplateText);
    }

    private static void AddParamedicalShareToken(
        Dictionary<string, string> tokens,
        IReadOnlyDictionary<string, decimal> shares,
        string tokenName,
        IReadOnlyCollection<string> categories,
        string literalTemplateText,
        string literalLabel,
        bool includeLiteralTemplateText)
    {
        var found = categories
            .Select(category => shares.TryGetValue(category, out var categoryShare)
                ? (Found: true, Share: categoryShare)
                : (Found: false, Share: 0m))
            .FirstOrDefault(item => item.Found);
        if (!found.Found)
        {
            return;
        }

        var formattedShare = found.Share.ToString("0.0%", CultureInfo.CurrentCulture);
        AddTokenVariants(tokens, tokenName, formattedShare);
        if (includeLiteralTemplateText)
        {
            tokens[literalTemplateText] = $"{literalLabel} ({formattedShare})";
        }
    }

    private static Dictionary<string, decimal> BuildParamedicalSharePercentages(AnalysisResult analysis)
    {
        var rows = analysis.Workbook.ParamedicalServiceTrends
            .Select(row => new
            {
                Category = ParamedicalCategoryNormalizer.Normalize(row.Category),
                row.CurrentAmount
            })
            .ToList();
        if (rows.Count == 0)
        {
            return [];
        }

        var rollupTotal = rows
            .Where(row => IsParamedicalRollupCategory(row.Category))
            .Sum(row => row.CurrentAmount);
        var detailRows = rows
            .Where(row => !IsParamedicalRollupCategory(row.Category))
            .ToList();
        var total = rollupTotal > 0m
            ? rollupTotal
            : detailRows.Sum(row => row.CurrentAmount);
        if (total <= 0m)
        {
            return [];
        }

        return detailRows
            .GroupBy(row => row.Category, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                group => group.Key,
                group => group.Sum(row => row.CurrentAmount) / total,
                StringComparer.OrdinalIgnoreCase);
    }

    private static bool IsParamedicalRollupCategory(string category)
    {
        return ParamedicalCategoryNormalizer.IsRollup(category) ||
               category.Equals("Paramedical Practitioners", StringComparison.OrdinalIgnoreCase);
    }

    private static string ExtractParagraphText(IEnumerable<Drawing.Paragraph> paragraphs)
    {
        return string.Concat(paragraphs.Select(paragraph => string.Concat(paragraph.Descendants<Drawing.Text>().Select(text => text.Text))));
    }

    private static bool ReplaceTemplateTables(string outputPath, AnalysisResult analysis)
    {
        var tables = PowerPointRateTableBuilder.BuildPlaceholderTables(analysis);
        var tableFound = false;
        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        var slidePaths = archive.Entries
            .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                            entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .Select(entry => entry.FullName)
            .ToList();

        foreach (var slidePath in slidePaths)
        {
            var entry = archive.GetEntry(slidePath);
            if (entry is null)
            {
                continue;
            }

            var document = XDocument.Parse(ReadZipEntryText(entry));
            if (ReplaceTablesInSlide(document, tables))
            {
                tableFound = true;
                entry.Delete();
                var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
                using var stream = newEntry.Open();
                document.Save(stream);
            }
        }

        return tableFound;
    }

    private static bool ReplaceTablesInSlide(XDocument document, Dictionary<string, PowerPointExportTable> tables)
    {
        var p = XNamespace.Get("http://schemas.openxmlformats.org/presentationml/2006/main");
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var found = false;

        foreach (var shape in document.Descendants(p + "sp").ToList())
        {
            var text = string.Concat(shape.Descendants(a + "t").Select(element => element.Value)).Trim();
            var match = tables.FirstOrDefault(table =>
                string.Equals(text, table.Key, StringComparison.Ordinal) ||
                text.Contains(table.Key, StringComparison.Ordinal));
            if (string.IsNullOrWhiteSpace(match.Key))
            {
                continue;
            }

            var id = NextShapeId(document, p);
            if (IsSimpleRateMasterTable(match.Value))
            {
                var tableGroup = BuildSimpleRateTableGroup(shape, match.Value, id, p, a);
                shape.ReplaceWith(tableGroup);
            }
            else
            {
                var tableFrame = BuildTableGraphicFrame(shape, match.Value, id, p, a);
                shape.ReplaceWith(tableFrame);
            }

            found = true;
        }

        return found;
    }

    private static bool TryPopulateNemMaximumsSlide(string outputPath, AnalysisResult analysis)
    {
        var table = PowerPointRateTableBuilder.BuildNemMaximumsTable(analysis);
        if (!table.Rows.Any(row => row.Skip(1).Any(value => !string.IsNullOrWhiteSpace(value))))
        {
            return false;
        }

        var updated = false;
        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        var slidePaths = archive.Entries
            .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                            entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .Select(entry => entry.FullName)
            .ToList();

        foreach (var slidePath in slidePaths)
        {
            var entry = archive.GetEntry(slidePath);
            if (entry is null)
            {
                continue;
            }

            var document = XDocument.Parse(ReadZipEntryText(entry));
            if (!LooksLikeNemSlide(document) || !PopulateNemTableInSlide(document, table))
            {
                continue;
            }

            updated = true;
            entry.Delete();
            var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
            using var stream = newEntry.Open();
            document.Save(stream);
        }

        return updated;
    }

    private static bool LooksLikeNemSlide(XDocument document)
    {
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var text = string.Join(" ", document.Descendants(a + "t").Select(element => element.Value));
        return text.Contains("Reviewing Your Plan", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Current NEM", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Available NEM", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Non-Evidence Maximum", StringComparison.OrdinalIgnoreCase);
    }

    private static bool PopulateNemTableInSlide(XDocument document, PowerPointExportTable table)
    {
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var updated = false;
        foreach (var drawingTable in document.Descendants(a + "tbl"))
        {
            if (!IsNemMaximumsDrawingTable(drawingTable, a))
            {
                continue;
            }

            NormalizeNemMaximumsDrawingTable(drawingTable, a);
            foreach (var row in drawingTable.Elements(a + "tr").Skip(1))
            {
                var cells = row.Elements(a + "tc").ToList();
                if (cells.Count < 3)
                {
                    continue;
                }

                var benefit = GetTableCellText(cells[0], a);
                var source = table.Rows.FirstOrDefault(tableRow =>
                    NemMaximumExtractor.NormalizeBenefitName(tableRow[0]).Equals(NemMaximumExtractor.NormalizeBenefitName(benefit), StringComparison.OrdinalIgnoreCase));
                if (source is null)
                {
                    continue;
                }

                SetTableCellText(cells[1], source[1], a);
                SetTableCellText(cells[2], source[2], a);
                updated = true;
            }
        }

        return updated;
    }

    private static bool IsNemMaximumsDrawingTable(XElement table, XNamespace a)
    {
        var text = string.Join(" ", table.Descendants(a + "t").Select(element => element.Value));
        return text.Contains("Current NEM", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Available NEM", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Current Maximum", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Non-Evidence Maximum", StringComparison.OrdinalIgnoreCase);
    }

    private static void NormalizeNemMaximumsDrawingTable(XElement table, XNamespace a)
    {
        var rows = table.Elements(a + "tr").ToList();
        if (rows.Count == 0)
        {
            return;
        }

        foreach (var row in rows)
        {
            var cells = row.Elements(a + "tc").ToList();
            while (cells.Count > 3)
            {
                cells[^1].Remove();
                cells.RemoveAt(cells.Count - 1);
            }
        }

        var grid = table.Element(a + "tblGrid");
        if (grid is not null)
        {
            var columns = grid.Elements(a + "gridCol").ToList();
            while (columns.Count > 3)
            {
                columns[^1].Remove();
                columns.RemoveAt(columns.Count - 1);
            }
        }

        var headerCells = rows[0].Elements(a + "tc").ToList();
        if (headerCells.Count >= 3)
        {
            SetTableCellText(headerCells[0], "Benefit", a);
            SetTableCellText(headerCells[1], "Current", a);
            SetTableCellText(headerCells[2], "Available", a);
        }
    }

    private static void SetTableCellText(XElement cell, string value, XNamespace a)
    {
        var textRuns = cell.Descendants(a + "t").ToList();
        if (textRuns.Count == 0)
        {
            return;
        }

        textRuns[0].Value = value;
        for (var index = 1; index < textRuns.Count; index++)
        {
            textRuns[index].Value = "";
        }
    }

    private static uint NextShapeId(XDocument document, XNamespace p)
    {
        var ids = document.Descendants(p + "cNvPr")
            .Select(element => uint.TryParse((string?)element.Attribute("id"), out var id) ? id : 0)
            .ToList();
        var maxId = ids.Count == 0 ? 1u : ids.Max();

        return maxId + 1;
    }

    private static XElement BuildSimpleRateTableGroup(XElement placeholderShape, PowerPointExportTable table, uint id, XNamespace p, XNamespace a)
    {
        var xfrm = placeholderShape.Element(p + "spPr")?.Element(a + "xfrm");
        var off = xfrm?.Element(a + "off");
        var ext = xfrm?.Element(a + "ext");
        var x = ParseLong((string?)off?.Attribute("x"), 720000);
        var y = ParseLong((string?)off?.Attribute("y"), 1200000);
        var cx = ParseLong((string?)ext?.Attribute("cx"), 10400000);
        var cy = ParseLong((string?)ext?.Attribute("cy"), 4300000);
        var tableShapes = BuildSimpleRateTableShapes(placeholderShape, table, id + 1, p, a).ToArray();

        return new XElement(
            p + "grpSp",
            new XElement(
                p + "nvGrpSpPr",
                new XElement(p + "cNvPr", new XAttribute("id", id), new XAttribute("name", "RateTable_Master Group")),
                new XElement(p + "cNvGrpSpPr"),
                new XElement(p + "nvPr")),
            new XElement(
                p + "grpSpPr",
                new XElement(
                    a + "xfrm",
                    new XElement(a + "off", new XAttribute("x", x), new XAttribute("y", y)),
                    new XElement(a + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy)),
                    new XElement(a + "chOff", new XAttribute("x", x), new XAttribute("y", y)),
                    new XElement(a + "chExt", new XAttribute("cx", cx), new XAttribute("cy", cy)))),
            tableShapes);
    }

    private static IEnumerable<XElement> BuildSimpleRateTableShapes(XElement placeholderShape, PowerPointExportTable table, uint startId, XNamespace p, XNamespace a)
    {
        var xfrm = placeholderShape.Element(p + "spPr")?.Element(a + "xfrm");
        var off = xfrm?.Element(a + "off");
        var ext = xfrm?.Element(a + "ext");
        var x = ParseLong((string?)off?.Attribute("x"), 720000);
        var y = ParseLong((string?)off?.Attribute("y"), 1200000);
        var cx = ParseLong((string?)ext?.Attribute("cx"), 10400000);
        var cy = ParseLong((string?)ext?.Attribute("cy"), 4300000);

        var rows = new List<string[]> { table.Headers };
        rows.AddRange(table.Rows);
        var columnCount = table.Headers.Length;
        var columnWidths = CalculateColumnWidths(table, cx);
        var rowWeights = rows
            .Select((row, rowIndex) => rowIndex == 0
                ? 1.15
                : Math.Max(1, row.Take(columnCount).Select(CountTextLines).DefaultIfEmpty(1).Max()))
            .ToList();
        var rowUnit = cy / Math.Max(rowWeights.Sum(), 1);
        var rowHeights = rowWeights
            .Select(weight => Math.Max(90000, (long)Math.Round(rowUnit * weight)))
            .ToArray();
        var totalHeight = rowHeights.Sum();
        var totalWidth = columnWidths.Sum();

        var shapes = new List<XElement>();
        var nextId = startId;
        var currentY = y;
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var currentX = x;
            var row = rows[rowIndex];
            var isHeader = rowIndex == 0;
            var isTotal = !isHeader && row.Length > 0 && row[0].Equals("Monthly Total", StringComparison.OrdinalIgnoreCase);

            for (var column = 0; column < columnCount; column++)
            {
                var value = column < row.Length ? row[column] : "";
                var fillColor = CellFillColor(isHeader, false, isTotal, false, true, false, false, column);
                var fontColor = isHeader ? "FFFFFF" : "004650";
                var bold = isHeader || isTotal;
                var numeric = IsNumericTableColumn(table.Headers, column);

                shapes.Add(BuildTextRectShape(
                    p,
                    a,
                    nextId++,
                    $"RateTable_Master Cell {rowIndex + 1}-{column + 1}",
                    currentX,
                    currentY,
                    columnWidths[column],
                    rowHeights[rowIndex],
                    fillColor,
                    value,
                    numeric,
                    1100,
                    fontColor,
                    bold));

                currentX += columnWidths[column];
            }

            currentY += rowHeights[rowIndex];
        }

        const long normalLine = 6350;
        const long thickLine = 19050;
        var separatorY = y + rowHeights.Take(Math.Max(0, rows.Count - 1)).Sum();

        var boundaryX = x;
        for (var column = 0; column <= columnCount; column++)
        {
            shapes.Add(BuildFilledRectShape(
                p,
                a,
                nextId++,
                $"RateTable_Master Vertical {column + 1}",
                boundaryX - normalLine / 2,
                y,
                normalLine,
                totalHeight,
                "8CD2C3"));

            if (column < columnCount)
            {
                boundaryX += columnWidths[column];
            }
        }

        var boundaryY = y;
        for (var rowIndex = 0; rowIndex <= rows.Count; rowIndex++)
        {
            var isMonthlyTotalSeparator = rowIndex == rows.Count - 1;
            shapes.Add(BuildFilledRectShape(
                p,
                a,
                nextId++,
                $"RateTable_Master Horizontal {rowIndex + 1}",
                x,
                (isMonthlyTotalSeparator ? separatorY : boundaryY) - (isMonthlyTotalSeparator ? thickLine : normalLine) / 2,
                totalWidth,
                isMonthlyTotalSeparator ? thickLine : normalLine,
                isMonthlyTotalSeparator ? "004650" : "8CD2C3"));

            if (rowIndex < rows.Count)
            {
                boundaryY += rowHeights[rowIndex];
            }
        }

        return shapes;
    }

    private static XElement BuildTextRectShape(
        XNamespace p,
        XNamespace a,
        uint id,
        string name,
        long x,
        long y,
        long cx,
        long cy,
        string fillColor,
        string text,
        bool numeric,
        int fontSize,
        string fontColor,
        bool bold)
    {
        var lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        if (lines.Length == 0)
        {
            lines = [""];
        }

        return new XElement(
            p + "sp",
            new XElement(
                p + "nvSpPr",
                new XElement(p + "cNvPr", new XAttribute("id", id), new XAttribute("name", name)),
                new XElement(p + "cNvSpPr"),
                new XElement(p + "nvPr")),
            new XElement(
                p + "spPr",
                new XElement(
                    a + "xfrm",
                    new XElement(a + "off", new XAttribute("x", x), new XAttribute("y", y)),
                    new XElement(a + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy))),
                new XElement(a + "prstGeom", new XAttribute("prst", "rect"), new XElement(a + "avLst")),
                new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", fillColor))),
                new XElement(a + "ln", new XElement(a + "noFill"))),
            new XElement(
                p + "txBody",
                new XElement(
                    a + "bodyPr",
                    new XAttribute("wrap", "square"),
                    new XAttribute("anchor", "ctr"),
                    new XAttribute("lIns", "76200"),
                    new XAttribute("rIns", "76200"),
                    new XAttribute("tIns", "38100"),
                    new XAttribute("bIns", "38100")),
                new XElement(a + "lstStyle"),
                lines.Select(line =>
                    new XElement(
                        a + "p",
                        new XElement(a + "pPr", new XAttribute("algn", numeric ? "r" : "l")),
                        new XElement(
                            a + "r",
                            new XElement(
                                a + "rPr",
                                new XAttribute("lang", "en-US"),
                                new XAttribute("sz", fontSize),
                                new XAttribute("b", bold ? "1" : "0"),
                                new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", fontColor)))),
                            new XElement(a + "t", line)),
                        new XElement(a + "endParaRPr", new XAttribute("lang", "en-US"), new XAttribute("sz", fontSize))))));
    }

    private static XElement BuildFilledRectShape(XNamespace p, XNamespace a, uint id, string name, long x, long y, long cx, long cy, string fillColor)
    {
        return new XElement(
            p + "sp",
            new XElement(
                p + "nvSpPr",
                new XElement(p + "cNvPr", new XAttribute("id", id), new XAttribute("name", name)),
                new XElement(p + "cNvSpPr"),
                new XElement(p + "nvPr")),
            new XElement(
                p + "spPr",
                new XElement(
                    a + "xfrm",
                    new XElement(a + "off", new XAttribute("x", x), new XAttribute("y", y)),
                    new XElement(a + "ext", new XAttribute("cx", Math.Max(1, cx)), new XAttribute("cy", Math.Max(1, cy)))),
                new XElement(a + "prstGeom", new XAttribute("prst", "rect"), new XElement(a + "avLst")),
                new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", fillColor))),
                new XElement(a + "ln", new XElement(a + "noFill"))));
    }

    private static XElement BuildTableGraphicFrame(XElement placeholderShape, PowerPointExportTable table, uint id, XNamespace p, XNamespace a)
    {
        var xfrm = placeholderShape.Element(p + "spPr")?.Element(a + "xfrm");
        var off = xfrm?.Element(a + "off");
        var ext = xfrm?.Element(a + "ext");
        var x = (string?)off?.Attribute("x") ?? "720000";
        var y = (string?)off?.Attribute("y") ?? "1200000";
        var cx = ParseLong((string?)ext?.Attribute("cx"), 10400000);
        var cy = ParseLong((string?)ext?.Attribute("cy"), 4300000);

        return new XElement(
            p + "graphicFrame",
            new XElement(
                p + "nvGraphicFramePr",
                new XElement(p + "cNvPr", new XAttribute("id", id), new XAttribute("name", table.Title)),
                new XElement(p + "cNvGraphicFramePr", new XElement(a + "graphicFrameLocks", new XAttribute("noGrp", "1"))),
                new XElement(p + "nvPr")),
            new XElement(
                p + "xfrm",
                new XElement(a + "off", new XAttribute("x", x), new XAttribute("y", y)),
                new XElement(a + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy))),
            new XElement(
                a + "graphic",
                new XElement(
                    a + "graphicData",
                    new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/table"),
                    BuildDrawingTable(table, cx, cy, a))));
    }

    private static XElement BuildDrawingTable(PowerPointExportTable table, long width, long height, XNamespace a)
    {
        var rows = new List<string[]> { table.Headers };
        rows.AddRange(table.Rows);
        var columnCount = table.Headers.Length;
        var columnWidths = CalculateColumnWidths(table, width);
        var rowWeights = rows
            .Select((row, rowIndex) => rowIndex == 0
                ? 1.15
                : Math.Max(1, row.Take(columnCount).Select(CountTextLines).DefaultIfEmpty(1).Max()))
            .ToList();
        var rowUnit = height / Math.Max(rowWeights.Sum(), 1);
        var isSimpleRateMaster = IsSimpleRateMasterTable(table);
        var isLineRateTable = IsLineRateTable(table);
        var isFormattedLineRateTable = IsFormattedLineRateTable(table);
        var isPremiumIllustrationTable = IsPremiumIllustrationTable(table);
        var fontSize = isSimpleRateMaster ? 1100 : isLineRateTable ? 600 : rows.Count > 34 ? 520 : rows.Count > 26 ? 600 : rows.Count > 18 ? 700 : 800;

        return new XElement(
            a + "tbl",
            new XElement(
                a + "tblPr",
                new XAttribute("firstRow", "1"),
                new XAttribute("bandRow", "1"),
                isSimpleRateMaster
                    ? null
                    : new XElement(a + "tableStyleId", "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")),
            new XElement(a + "tblGrid", columnWidths.Select(value => new XElement(a + "gridCol", new XAttribute("w", value)))),
            rows.Select((row, rowIndex) =>
            {
                var isHeader = rowIndex == 0;
                var isLineRateColumnHeader = isFormattedLineRateTable && IsLineRateColumnHeaderRow(row);
                var isLineRateSection = isFormattedLineRateTable && !isHeader && !isLineRateColumnHeader && IsLineRateSectionRow(row);
                var isTotal = isLineRateSection || (!isHeader && !isLineRateColumnHeader && row.Any(cell =>
                    cell.Contains("Total", StringComparison.OrdinalIgnoreCase) ||
                    cell.Contains("Monthly Difference", StringComparison.OrdinalIgnoreCase) ||
                    cell.Contains("premium difference", StringComparison.OrdinalIgnoreCase) ||
                    cell.Contains("Annual premium difference", StringComparison.OrdinalIgnoreCase)));
                var isMonthlyTotal = isSimpleRateMaster &&
                                     !isHeader &&
                                     row.Length > 0 &&
                                     row[0].Equals("Monthly Total", StringComparison.OrdinalIgnoreCase);
                var isBanded = !isHeader && !isLineRateColumnHeader && !isTotal && rowIndex % 2 == 0;
                return new XElement(
                    a + "tr",
                    new XAttribute("h", Math.Max(90000, (long)Math.Round(rowUnit * (isLineRateColumnHeader ? 0.8 : rowWeights[rowIndex])))),
                    Enumerable.Range(0, columnCount).Select(column =>
                    {
                        var value = column < row.Length ? row[column] : "";
                        var numeric = isFormattedLineRateTable ? column > 0 && !isLineRateColumnHeader : IsNumericTableColumn(table.Headers, column);
                        return BuildTableCell(value, isHeader, isLineRateColumnHeader, isTotal, isBanded, isSimpleRateMaster, isFormattedLineRateTable, isPremiumIllustrationTable, isMonthlyTotal, rowIndex, rows.Count, column, columnCount, numeric, fontSize, a);
                    }));
            }));
    }

    private static XElement BuildTableCell(string value, bool isHeader, bool isLineRateColumnHeader, bool isTotal, bool isBanded, bool isSimpleRateMaster, bool isLineRateTable, bool isPremiumIllustrationTable, bool isMonthlyTotal, int rowIndex, int rowCount, int column, int columnCount, bool numeric, int fontSize, XNamespace a)
    {
        var fillColor = CellFillColor(isHeader, isLineRateColumnHeader, isTotal, isBanded, isSimpleRateMaster, isLineRateTable, isPremiumIllustrationTable, column);
        var fontColor = isHeader
            ? isPremiumIllustrationTable ? "E7F6F9" : "FFFFFF"
            : isPremiumIllustrationTable ? "004751" : "004650";
        var bold = isHeader || isLineRateColumnHeader || isTotal ? "1" : "0";
        var lines = value.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        if (lines.Length == 0)
        {
            lines = [""];
        }

        return new XElement(
            a + "tc",
            new XElement(
                a + "txBody",
                new XElement(a + "bodyPr", new XAttribute("wrap", "square"), new XAttribute("anchor", "ctr")),
                new XElement(a + "lstStyle"),
                lines.Select((line, lineIndex) =>
                {
                    var lineBold = bold == "1" ||
                                   (isPremiumIllustrationTable && (column <= 1 || column == columnCount - 1)) ||
                                   (isPremiumIllustrationTable && column is >= 3 and <= 5 && lines.Length > 1 && lineIndex == lines.Length - 1) ||
                                   (!isPremiumIllustrationTable && numeric && lineIndex > 0)
                        ? "1"
                        : "0";
                    var alignment = isPremiumIllustrationTable
                        ? column == 0 ? "l" : isHeader || column == 1 ? "ctr" : "r"
                        : numeric ? "r" : "l";
                    return new XElement(
                        a + "p",
                        new XElement(a + "pPr", new XAttribute("algn", alignment)),
                        new XElement(
                            a + "r",
                            new XElement(
                                a + "rPr",
                                new XAttribute("lang", "en-US"),
                                new XAttribute("sz", fontSize),
                                new XAttribute("b", lineBold),
                                new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", fontColor))),
                                isPremiumIllustrationTable ? new XElement(a + "latin", new XAttribute("typeface", "Arial")) : null),
                            new XElement(a + "t", line)),
                        new XElement(
                            a + "endParaRPr",
                            new XAttribute("lang", "en-US"),
                            new XAttribute("sz", fontSize),
                            isPremiumIllustrationTable ? new XElement(a + "latin", new XAttribute("typeface", "Arial")) : null));
                })),
            new XElement(
                a + "tcPr",
                new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", fillColor))),
                BuildCellBorder(a, "lnL", isSimpleRateMaster, isLineRateTable, isPremiumIllustrationTable, isTotal, isMonthlyTotal, rowIndex, rowCount, column, columnCount),
                BuildCellBorder(a, "lnR", isSimpleRateMaster, isLineRateTable, isPremiumIllustrationTable, isTotal, isMonthlyTotal, rowIndex, rowCount, column, columnCount),
                BuildCellBorder(a, "lnT", isSimpleRateMaster, isLineRateTable, isPremiumIllustrationTable, isTotal, isMonthlyTotal, rowIndex, rowCount, column, columnCount),
                BuildCellBorder(a, "lnB", isSimpleRateMaster, isLineRateTable, isPremiumIllustrationTable, isTotal, isMonthlyTotal, rowIndex, rowCount, column, columnCount)));
    }

    private static XElement BuildCellBorder(XNamespace a, string name, bool isSimpleRateMaster, bool isLineRateTable, bool isPremiumIllustrationTable, bool isTotal, bool isMonthlyTotal, int rowIndex, int rowCount, int column, int columnCount)
    {
        var isMonthlyTotalTop = isSimpleRateMaster &&
                                isMonthlyTotal &&
                                name.Equals("lnT", StringComparison.Ordinal);
        var isLineRateTotalTop = isLineRateTable &&
                                 isTotal &&
                                 name.Equals("lnT", StringComparison.Ordinal);
        var isPremiumIllustrationTotalTop = isPremiumIllustrationTable && isTotal && name.Equals("lnT", StringComparison.Ordinal);
        var color = isMonthlyTotalTop || isLineRateTotalTop || isPremiumIllustrationTotalTop
            ? "004650"
            : isPremiumIllustrationTable ? "71C9B6" : isSimpleRateMaster || isLineRateTable ? "8DD3C7" : "8FD1D2";
        var width = isMonthlyTotalTop || isLineRateTotalTop || isPremiumIllustrationTotalTop ? "19050" : "6350";

        return new XElement(
            a + name,
            new XAttribute("w", width),
            new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", color))));
    }

    private static string CellFillColor(bool isHeader, bool isLineRateColumnHeader, bool isTotal, bool isBanded, bool isSimpleRateMaster, bool isLineRateTable, bool isPremiumIllustrationTable, int column)
    {
        if (isPremiumIllustrationTable)
        {
            return isHeader ? "004751" : isTotal ? "E7F6F9" : "CDECEF";
        }

        if (isHeader)
        {
            if (isLineRateTable)
            {
                if (column < 2)
                {
                    return "FFFFFF";
                }

                return column < 4 ? "808080" : "004650";
            }

            return isSimpleRateMaster ? "004650" : "005059";
        }

        if (isLineRateTable && isLineRateColumnHeader)
        {
            return "EAF8F8";
        }

        if (isLineRateTable && isTotal)
        {
            return "EAF8F8";
        }

        if (isLineRateTable)
        {
            return column switch
            {
                2 or 3 => "F2F2F2",
                4 or 5 => "CDECEF",
                6 or 7 => "EAF8F8",
                8 or 9 => "CDECEF",
                _ => isBanded ? "CDECEF" : "EAF8F8"
            };
        }

        if (isSimpleRateMaster)
        {
            return column is 0 or 2 ? "CDECEF" : "EAF8F8";
        }

        return isTotal ? "DFF4F5" : isBanded ? "CDECEF" : "EAF8F8";
    }

    private static long[] CalculateColumnWidths(PowerPointExportTable table, long width)
    {
        var weights = IsPremiumIllustrationTable(table)
            ? new[] { 2.45, 0.6, 1.5, 1.2, 1.35, 1.55, 1.05 }
            : IsTopDrugExportTable(table)
            ? new[] { 0.45, 3.3, 2.1, 1.95 }
            : table.Title.Equals("RatesOverview", StringComparison.OrdinalIgnoreCase)
                ? new[] { 2.35, 1.35, 0.9, 1.0, 1.35, 1.35, 1.05 }
            : table.Headers.Length switch
        {
            10 => new[] { 2.35, 0.8, 0.78, 1.05, 0.85, 1.15, 0.85, 1.15, 0.9, 1.35 },
            11 => new[] { 0.95, 0.9, 0.6, 0.78, 0.45, 0.9, 1.15, 0.9, 1.15, 0.95, 1.3 },
            12 => new[] { 0.95, 0.9, 0.6, 0.78, 0.45, 0.9, 1.15, 0.9, 1.15, 0.95, 1.3, 1.55 },
            8 => new[] { 1.35, 1.15, 1.15, 0.9, 1.15, 0.9, 1.3, 0.9 },
            7 => new[] { 1.65, 2.55, 0.8, 1.35, 1.15, 1.15, 0.9 },
            5 => new[] { 1.45, 1.35, 1.15, 1.2, 1.0 },
            4 => new[] { 2.4, 1.35, 1.35, 1.0 },
            _ => new[] { 2.7, 0.85, 1.35, 1.15, 1.15, 0.9 }
        };

        var total = weights.Sum();
        var remaining = width;
        var result = new long[weights.Length];
        for (var index = 0; index < weights.Length; index++)
        {
            result[index] = index == weights.Length - 1
                ? remaining
                : Math.Max(350000, (long)Math.Round(width * weights[index] / total));
            remaining -= result[index];
        }

        return result;
    }

    private static bool IsTopDrugExportTable(PowerPointExportTable table)
    {
        return table.Title.Equals("TopDrugs", StringComparison.OrdinalIgnoreCase) ||
               table.Title.Equals("CurrentTopDrugs", StringComparison.OrdinalIgnoreCase) ||
               table.Title.Equals("PriorTopDrugs", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsNumericTableColumn(string[] headers, int column)
    {
        if (column < 0 || column >= headers.Length)
        {
            return false;
        }

        var header = headers[column];
        return header.Contains("Premium", StringComparison.OrdinalIgnoreCase) ||
               header.Contains("Current", StringComparison.OrdinalIgnoreCase) ||
               header.Contains("Renewal", StringComparison.OrdinalIgnoreCase) ||
               header.Contains("Proposed", StringComparison.OrdinalIgnoreCase) ||
               header.Contains("Rate", StringComparison.OrdinalIgnoreCase) ||
               header.Contains("Change", StringComparison.OrdinalIgnoreCase) ||
               header.Contains("Claims", StringComparison.OrdinalIgnoreCase) ||
               header.Contains("Paid", StringComparison.OrdinalIgnoreCase) ||
               header.Contains("Amount", StringComparison.OrdinalIgnoreCase) ||
               header.Contains("Ratio", StringComparison.OrdinalIgnoreCase) ||
               header.Contains("Volume", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSimpleRateMasterTable(PowerPointExportTable table)
    {
        return table.Headers.Length == 4 &&
               table.Headers[0].Equals("Benefit", StringComparison.OrdinalIgnoreCase) &&
               table.Headers[1].Equals("Current Premium", StringComparison.OrdinalIgnoreCase) &&
               table.Headers[2].Equals("Proposed Premium", StringComparison.OrdinalIgnoreCase) &&
               table.Headers[3].Equals("Change (%)", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsLineRateTable(PowerPointExportTable table)
    {
        return table.Title.Equals("RateTable_LineRates", StringComparison.OrdinalIgnoreCase) ||
               table.Title.Equals("RateTable_LineRates_Formatted", StringComparison.OrdinalIgnoreCase) ||
               table.Headers.Any(header => header.Equals("Negotiated Renewal Premium", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsFormattedLineRateTable(PowerPointExportTable table)
    {
        return table.Title.Equals("RateTable_LineRates_Formatted", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsLineRateColumnHeaderRow(string[] row)
    {
        return row.Length >= 10 &&
               row[0].Equals("Benefit / Class", StringComparison.OrdinalIgnoreCase) &&
               row[1].Equals("Volume", StringComparison.OrdinalIgnoreCase) &&
               row[2].Equals("Rate", StringComparison.OrdinalIgnoreCase) &&
               row[3].Equals("Premium", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsLineRateSectionRow(string[] row)
    {
        return row.Length > 0 &&
               !string.IsNullOrWhiteSpace(row[0]) &&
               row.Skip(1).All(string.IsNullOrWhiteSpace);
    }

    private static int CountTextLines(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return 1;
        }

        return value.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n').Length;
    }

    private static long ParseLong(string? value, long fallback)
    {
        return long.TryParse(value, out var result) ? result : fallback;
    }

    private static bool TryPopulateDemographicCharts(string outputPath, AnalysisResult analysis)
    {
        try
        {
            var updated = false;
            var priorYear = analysis.RenewalCase.RenewalDate.HasValue
                ? (analysis.RenewalCase.RenewalDate.Value.Year - 1).ToString()
                : "Prior";
            var currentYear = analysis.RenewalCase.RenewalDate.HasValue
                ? analysis.RenewalCase.RenewalDate.Value.Year.ToString()
                : "Current";

            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);

            var lifeRows = BuildLifeChartRows(analysis.Pdf.LifeDemographics);
            if (lifeRows.Count > 0)
            {
                updated |= UpdateDemographicChartBySlideTitle(
                    archive,
                    "Life Insurance and ASI Demographic Analysis",
                    priorYear,
                    currentYear,
                    lifeRows);
                updated |= UpdateDemographicChartBySlideTitle(
                    archive,
                    "Life Insurance and AD&D Demographic Analysis",
                    priorYear,
                    currentYear,
                    lifeRows);
            }

            var ltdRows = BuildLtdChartRows(analysis.Pdf.LtdDemographics);
            if (ltdRows.Count > 0)
            {
                updated |= UpdateDemographicChartBySlideTitle(
                    archive,
                    "Long-Term Disability Demographic Analysis",
                    priorYear,
                    currentYear,
                    ltdRows);
            }

            return updated;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryPopulateActivityCharts(string outputPath, AnalysisResult analysis)
    {
        try
        {
            var updated = false;
            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);

            var ehcRows = BuildActivityChartRows(
                analysis,
                "Non Pooled Extended Health Care - Subtotal",
                "Extended Health Care",
                "Extended Healthcare",
                "EHC");
            if (ehcRows.Count > 0)
            {
                updated |= UpdateActivityChartBySlideTitle(
                    archive,
                    "Extended Healthcare Claims Analysis",
                    "EHC Analysis",
                    ehcRows);
            }

            var dentalRows = BuildActivityChartRows(analysis, "Dental");
            if (dentalRows.Count > 0)
            {
                updated |= UpdateActivityChartBySlideTitle(
                    archive,
                    "Dental Analysis",
                    "Dental Analysis",
                    dentalRows);
            }

            return updated;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryPopulateEhcSpendingTrendChart(string outputPath, AnalysisResult analysis)
    {
        try
        {
            var rows = BuildEhcSpendingTrendChartRows(analysis);
            if (rows.Count == 0)
            {
                return false;
            }

            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
            return UpdateEhcSpendingTrendChartBySlideTitle(
                archive,
                rows,
                "Extended Healthcare Spending Trends",
                "Extended Health Spending Trends",
                "Health Spending Trends");
        }
        catch
        {
            return false;
        }
    }

    private static bool TryPopulateDentalSpendingTrendChart(string outputPath, AnalysisResult analysis)
    {
        try
        {
            var rows = BuildDentalSpendingTrendChartRows(analysis);
            if (rows.Count == 0)
            {
                return false;
            }

            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
            return UpdateEhcSpendingTrendChartBySlideTitle(
                archive,
                rows,
                "Dental Claims Analysis",
                "Dental Spending Trends",
                "Dental Trend");
        }
        catch
        {
            return false;
        }
    }

    private static bool TryPopulateParamedicalServiceTrendChart(string outputPath, AnalysisResult analysis)
    {
        try
        {
            var rows = BuildParamedicalServiceTrendChartRows(analysis);
            if (rows.Count == 0)
            {
                return false;
            }

            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
            return UpdateEhcSpendingTrendChartBySlideTitle(
                archive,
                rows,
                "Paramedical Services Utilization Trends",
                "Paramedical Services",
                "Utilization Trends");
        }
        catch
        {
            return false;
        }
    }

    private static bool TryPopulateTopDrugSlides(string outputPath, AnalysisResult analysis)
    {
        try
        {
            var currentTable = PowerPointRateTableBuilder.BuildTopDrugsTable(analysis);
            var priorTable = PowerPointRateTableBuilder.BuildPriorTopDrugsTable(analysis);
            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
            var slidePaths = archive.Entries
                .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                                entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .Select(entry => entry.FullName)
                .ToList();
            var currentTableGeometry = FindCurrentTopDrugTableGeometry(archive, slidePaths);

            var updated = false;
            var topDrugSlideIndex = 0;
            foreach (var slidePath in slidePaths)
            {
                var entry = archive.GetEntry(slidePath);
                if (entry is null)
                {
                    continue;
                }

                var document = XDocument.Parse(ReadZipEntryText(entry));
                var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
                var slideText = string.Concat(document.Descendants(a + "t").Select(element => element.Value));
                if (!SlideContainsTopDrugTarget(document, slideText, a))
                {
                    continue;
                }

                var isPriorSlide = IsExplicitPriorTopDrugSlide(slideText) ||
                                   (!IsExplicitCurrentTopDrugSlide(slideText) && topDrugSlideIndex > 0);
                var table = isPriorSlide ? priorTable : currentTable;
                if (!PopulateTopDrugTableInSlide(document, table, isPriorSlide, isPriorSlide ? currentTableGeometry : null))
                {
                    continue;
                }

                topDrugSlideIndex++;
                RemoveTopDrugTablePlaceholderTokens(document);
                updated = true;
                entry.Delete();
                var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
                using var stream = newEntry.Open();
                document.Save(stream);
            }

            return updated;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryPopulateProgramPerformanceClaimsSummaryChart(string outputPath, AnalysisResult analysis)
    {
        try
        {
            var data = BuildProgramClaimsSummaryChartData(analysis);
            if (data.Rows.Count == 0 || string.IsNullOrWhiteSpace(data.ReportingPeriod))
            {
                return false;
            }

            using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
            var updated = UpdateProgramClaimsSummarySlideText(archive, data);
            var chartPaths = FindChartPathsBySlideTitle(
                archive,
                "Program Performance: Claims Summary",
                "Program Performance: Premium vs Claims Summary");

            foreach (var chartPath in chartPaths)
            {
                UpdateProgramClaimsSummaryChartXml(archive, chartPath, data);

                var embeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, chartPath);
                if (embeddedWorkbookPath is not null)
                {
                    UpdateEmbeddedProgramClaimsSummaryWorkbook(archive, embeddedWorkbookPath, data);
                }

                updated = true;
            }

            return updated;
        }
        catch
        {
            return false;
        }
    }

    private static bool UpdateProgramClaimsSummarySlideText(ZipArchive archive, ProgramClaimsSummaryChartData data)
    {
        var updated = false;
        foreach (var slideEntry in archive.Entries
                     .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                                     entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                     .OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase)
                     .ToList())
        {
            var xml = ReadZipEntryText(slideEntry);
            var slideText = ExtractSlideText(xml);
            var normalizedSlideText = NormalizeSlideLookupText(slideText);
            if (!normalizedSlideText.Contains("PROGRAM PERFORMANCE CLAIMS SUMMARY", StringComparison.OrdinalIgnoreCase) &&
                !normalizedSlideText.Contains("PROGRAM PERFORMANCE PREMIUM VS CLAIMS SUMMARY", StringComparison.OrdinalIgnoreCase) &&
                !ContainsReportingPeriodToken(slideText))
            {
                continue;
            }

            var document = XDocument.Parse(xml);
            var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
            var changed = false;
            foreach (var text in document.Descendants(a + "t"))
            {
                var value = text.Value;
                value = value.Replace("[reporting period]", data.ReportingPeriod, StringComparison.OrdinalIgnoreCase);
                value = value.Replace("[Reporting Period]", data.ReportingPeriod, StringComparison.OrdinalIgnoreCase);
                value = value.Replace("[Enter Reporting Period]", data.ReportingPeriod, StringComparison.OrdinalIgnoreCase);
                value = value.Replace("[enter range]", data.ReportingPeriod, StringComparison.OrdinalIgnoreCase);
                if (!string.Equals(value, text.Value, StringComparison.Ordinal))
                {
                    text.Value = value;
                    changed = true;
                }
            }

            if (!changed)
            {
                continue;
            }

            var slidePath = slideEntry.FullName;
            slideEntry.Delete();
            var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
            using var stream = newEntry.Open();
            document.Save(stream);
            updated = true;
        }

        return updated;
    }

    private static bool ContainsReportingPeriodToken(string slideText)
    {
        return slideText.Contains("[reporting period]", StringComparison.OrdinalIgnoreCase) ||
               slideText.Contains("[enter reporting period]", StringComparison.OrdinalIgnoreCase) ||
               slideText.Contains("[enter range]", StringComparison.OrdinalIgnoreCase);
    }

    private static void UpdateProgramClaimsSummaryChartXml(ZipArchive archive, string chartPath, ProgramClaimsSummaryChartData data)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
        var series = document.Descendants(c + "ser").Take(3).ToList();
        if (series.Count < 3)
        {
            return;
        }

        foreach (var extraSeries in document.Descendants(c + "ser").Skip(3).ToList())
        {
            extraSeries.Remove();
        }

        var categories = data.Rows.Select(row => row.Category).ToList();
        SetStringReference(series[0].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$B$1", ["Paid Premiums"]);
        SetStringReference(series[0].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[0].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$B$2:$B${data.Rows.Count + 1}", data.Rows.Select(row => row.PaidPremiums).ToList(), "$#,##0");

        SetStringReference(series[1].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$C$1", ["Paid Claims"]);
        SetStringReference(series[1].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[1].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$C$2:$C${data.Rows.Count + 1}", data.Rows.Select(row => row.PaidClaims).ToList(), "$#,##0");

        SetStringReference(series[2].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$D$1", ["Loss Ratio"]);
        SetStringReference(series[2].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[2].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$D$2:$D${data.Rows.Count + 1}", data.Rows.Select(row => row.LossRatio).ToList(), "0%");

        entry.Delete();
        var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
    }

    private static void UpdateEmbeddedProgramClaimsSummaryWorkbook(ZipArchive archive, string workbookPath, ProgramClaimsSummaryChartData data)
    {
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = data.ReportingPeriod;
        sheet.Cell(1, 2).Value = "Paid Premiums";
        sheet.Cell(1, 3).Value = "Paid Claims";
        sheet.Cell(1, 4).Value = "Loss Ratio";
        sheet.Range(1, 1, 1, 4).Style.Font.Bold = true;

        for (var index = 0; index < data.Rows.Count; index++)
        {
            var rowNumber = index + 2;
            sheet.Cell(rowNumber, 1).Value = data.Rows[index].Category;
            sheet.Cell(rowNumber, 2).Value = data.Rows[index].PaidPremiums;
            sheet.Cell(rowNumber, 3).Value = data.Rows[index].PaidClaims;
            sheet.Cell(rowNumber, 4).FormulaA1 = $"=IFERROR(C{rowNumber}/B{rowNumber},0)";
        }

        sheet.Range(2, 2, data.Rows.Count + 1, 3).Style.NumberFormat.Format = "$#,##0";
        sheet.Range(2, 4, data.Rows.Count + 1, 4).Style.NumberFormat.Format = "0%";
        sheet.Columns(1, 4).AdjustToContents(8, 28);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;

        entry.Delete();
        var newEntry = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var outputStream = newEntry.Open();
        output.CopyTo(outputStream);
    }

    private static bool PopulateTopDrugTableInSlide(
        XDocument document,
        PowerPointExportTable table,
        bool isPriorSlide,
        TableFrameGeometry? referenceGeometry)
    {
        var p = XNamespace.Get("http://schemas.openxmlformats.org/presentationml/2006/main");
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var slideText = string.Concat(document.Descendants(a + "t").Select(element => element.Value));
        if (!isPriorSlide &&
            (slideText.Contains("Prior_Year", StringComparison.OrdinalIgnoreCase) ||
             slideText.Contains("Prior Year", StringComparison.OrdinalIgnoreCase) ||
             slideText.Contains("PriorTopDrugsTable", StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        if (!slideText.Contains("Top 10 DIN", StringComparison.OrdinalIgnoreCase) &&
            !(slideText.Contains("Drug Name", StringComparison.OrdinalIgnoreCase) &&
              slideText.Contains("Total Amount Paid", StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        foreach (var drawingTable in document.Descendants(a + "tbl"))
        {
            var rows = drawingTable.Elements(a + "tr").ToList();
            if (rows.Count < 2)
            {
                continue;
            }

            var headerTexts = rows[0].Elements(a + "tc")
                .Select(cell => GetTableCellText(cell, a))
                .ToList();
            var drugNameColumn = FindHeaderColumn(headerTexts, "Drug Name");
            var treatmentColumn = FindHeaderColumn(headerTexts, "Treatment");
            if (treatmentColumn < 0)
            {
                treatmentColumn = FindHeaderColumn(headerTexts, "Treats");
            }
            var amountColumn = FindHeaderColumn(headerTexts, "Total Amount Paid");
            if (drugNameColumn < 0 || amountColumn < 0)
            {
                continue;
            }

            if (referenceGeometry is not null)
            {
                ApplyTableFrameGeometry(drawingTable, referenceGeometry, p, a);
            }

            var indexColumn = drugNameColumn > 0 ? drugNameColumn - 1 : -1;
            for (var index = 0; index < table.Rows.Count && index + 1 < rows.Count; index++)
            {
                var cells = rows[index + 1].Elements(a + "tc").ToList();
                var row = table.Rows[index];
                SetTableCellText(cells, indexColumn, row[0], a, "ctr");
                SetTableCellText(cells, drugNameColumn, row[1], a, "l");
                SetTableCellText(cells, treatmentColumn, row[2], a, "l");
                SetTableCellText(cells, amountColumn, row[3], a, "r");
            }

            return true;
        }

        return false;
    }

    private static TableFrameGeometry? FindCurrentTopDrugTableGeometry(ZipArchive archive, IReadOnlyList<string> slidePaths)
    {
        foreach (var slidePath in slidePaths)
        {
            var entry = archive.GetEntry(slidePath);
            if (entry is null)
            {
                continue;
            }

            var document = XDocument.Parse(ReadZipEntryText(entry));
            var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
            var slideText = string.Concat(document.Descendants(a + "t").Select(element => element.Value));
            if (IsExplicitPriorTopDrugSlide(slideText))
            {
                continue;
            }

            foreach (var drawingTable in document.Descendants(a + "tbl"))
            {
                if (!IsTopDrugDrawingTable(drawingTable, a))
                {
                    continue;
                }

                return GetTableFrameGeometry(drawingTable);
            }
        }

        return null;
    }

    private static bool SlideContainsTopDrugTarget(XDocument document, string slideText, XNamespace a)
    {
        if (slideText.Contains("Top 10 DIN", StringComparison.OrdinalIgnoreCase) ||
            slideText.Contains("Top 10 Drug", StringComparison.OrdinalIgnoreCase) ||
            slideText.Contains("CurrentTopDrugs", StringComparison.OrdinalIgnoreCase) ||
            slideText.Contains("PriorTopDrugs", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return document.Descendants(a + "tbl").Any(table => IsTopDrugDrawingTable(table, a));
    }

    private static bool IsTopDrugDrawingTable(XElement drawingTable, XNamespace a)
    {
        var firstRow = drawingTable.Elements(a + "tr").FirstOrDefault();
        if (firstRow is null)
        {
            return false;
        }

        var headerTexts = firstRow.Elements(a + "tc")
            .Select(cell => GetTableCellText(cell, a))
            .ToList();

        return FindHeaderColumn(headerTexts, "Drug Name") >= 0 &&
               FindHeaderColumn(headerTexts, "Total Amount Paid") >= 0;
    }

    private static TableFrameGeometry? GetTableFrameGeometry(XElement drawingTable)
    {
        var p = XNamespace.Get("http://schemas.openxmlformats.org/presentationml/2006/main");
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var graphicFrame = drawingTable.Ancestors(p + "graphicFrame").FirstOrDefault();
        var xfrm = graphicFrame?.Element(p + "xfrm");
        var off = xfrm?.Element(a + "off");
        var ext = xfrm?.Element(a + "ext");
        var x = (string?)off?.Attribute("x");
        var y = (string?)off?.Attribute("y");
        var cx = (string?)ext?.Attribute("cx");
        var cy = (string?)ext?.Attribute("cy");

        return string.IsNullOrWhiteSpace(x) ||
               string.IsNullOrWhiteSpace(y) ||
               string.IsNullOrWhiteSpace(cx) ||
               string.IsNullOrWhiteSpace(cy)
            ? null
            : new TableFrameGeometry(x, y, cx, cy);
    }

    private static void ApplyTableFrameGeometry(XElement drawingTable, TableFrameGeometry geometry, XNamespace p, XNamespace a)
    {
        var graphicFrame = drawingTable.Ancestors(p + "graphicFrame").FirstOrDefault();
        var xfrm = graphicFrame?.Element(p + "xfrm");
        if (xfrm is null)
        {
            return;
        }

        var off = xfrm.Element(a + "off");
        if (off is not null)
        {
            off.SetAttributeValue("x", geometry.X);
            off.SetAttributeValue("y", geometry.Y);
        }

        var ext = xfrm.Element(a + "ext");
        if (ext is not null)
        {
            ext.SetAttributeValue("cx", geometry.Cx);
            ext.SetAttributeValue("cy", geometry.Cy);
        }
    }

    private static bool IsExplicitPriorTopDrugSlide(string slideText)
    {
        return slideText.Contains("PriorTopDrugsTable", StringComparison.OrdinalIgnoreCase) ||
               slideText.Contains("PriorTopDrugs", StringComparison.OrdinalIgnoreCase) ||
               slideText.Contains("Prior Top", StringComparison.OrdinalIgnoreCase) ||
               slideText.Contains("Prior_Year", StringComparison.OrdinalIgnoreCase) ||
               slideText.Contains("Prior Year", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsExplicitCurrentTopDrugSlide(string slideText)
    {
        return slideText.Contains("CurrentTopDrugsTable", StringComparison.OrdinalIgnoreCase) ||
               slideText.Contains("CurrentTopDrugs", StringComparison.OrdinalIgnoreCase) ||
               slideText.Contains("Current Top", StringComparison.OrdinalIgnoreCase) ||
               slideText.Contains("Current_Year", StringComparison.OrdinalIgnoreCase) ||
               slideText.Contains("Current Year", StringComparison.OrdinalIgnoreCase);
    }

    private static void RemoveTopDrugTablePlaceholderTokens(XDocument document)
    {
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var tokens = new[]
        {
            "<<CurrentTopDrugsTable>>",
            "<< CurrentTopDrugsTable >>",
            "{{CurrentTopDrugsTable}}",
            "{{ CurrentTopDrugsTable }}",
            "\u00ABCurrentTopDrugsTable\u00BB",
            "\u00AB CurrentTopDrugsTable \u00BB",
            "<<TopDrugsTable>>",
            "<< TopDrugsTable >>",
            "{{TopDrugsTable}}",
            "{{ TopDrugsTable }}",
            "\u00ABTopDrugsTable\u00BB",
            "\u00AB TopDrugsTable \u00BB",
            "<<PriorTopDrugsTable>>",
            "<< PriorTopDrugsTable >>",
            "{{PriorTopDrugsTable}}",
            "{{ PriorTopDrugsTable }}",
            "\u00ABPriorTopDrugsTable\u00BB",
            "\u00AB PriorTopDrugsTable \u00BB",
            "\u00C2\u00ABPriorTopDrugsTable\u00C2\u00BB"
        };

        foreach (var paragraph in document.Descendants(a + "p"))
        {
            RemoveTokensFromParagraph(paragraph, tokens, a);
        }
    }

    private static void RemoveTokensFromParagraph(XElement paragraph, IReadOnlyCollection<string> tokens, XNamespace a)
    {
        var textRuns = paragraph.Descendants(a + "t").ToList();
        if (textRuns.Count == 0)
        {
            return;
        }

        var runTexts = textRuns.Select(element => element.Value ?? "").ToList();
        var combined = string.Concat(runTexts);
        if (string.IsNullOrEmpty(combined))
        {
            return;
        }

        var delete = new bool[combined.Length];
        foreach (var token in tokens)
        {
            var startIndex = 0;
            while (startIndex < combined.Length)
            {
                var tokenIndex = combined.IndexOf(token, startIndex, StringComparison.OrdinalIgnoreCase);
                if (tokenIndex < 0)
                {
                    break;
                }

                for (var index = tokenIndex; index < tokenIndex + token.Length && index < delete.Length; index++)
                {
                    delete[index] = true;
                }

                startIndex = tokenIndex + token.Length;
            }
        }

        if (!delete.Any(value => value))
        {
            return;
        }

        var offset = 0;
        for (var runIndex = 0; runIndex < textRuns.Count; runIndex++)
        {
            var original = runTexts[runIndex];
            var kept = new char[original.Length];
            var count = 0;
            for (var characterIndex = 0; characterIndex < original.Length; characterIndex++)
            {
                if (!delete[offset + characterIndex])
                {
                    kept[count++] = original[characterIndex];
                }
            }

            textRuns[runIndex].Value = new string(kept, 0, count);
            offset += original.Length;
        }
    }

    private static int FindHeaderColumn(IReadOnlyList<string> headers, string text)
    {
        for (var index = 0; index < headers.Count; index++)
        {
            if (headers[index].Contains(text, StringComparison.OrdinalIgnoreCase))
            {
                return index;
            }
        }

        return -1;
    }

    private static string GetTableCellText(XElement cell, XNamespace a)
    {
        return string.Concat(cell.Descendants(a + "t").Select(element => element.Value)).Trim();
    }

    private static void SetTableCellText(IReadOnlyList<XElement> cells, int column, string value, XNamespace a, string alignment)
    {
        if (column < 0 || column >= cells.Count)
        {
            return;
        }

        var cell = cells[column];
        var textRuns = cell.Descendants(a + "t").ToList();
        if (textRuns.Count == 0)
        {
            var paragraphForRun = cell.Descendants(a + "p").FirstOrDefault();
            if (paragraphForRun is null)
            {
                return;
            }

            var run = BuildTableCellTextRun(paragraphForRun, a);
            var endProperties = paragraphForRun.Element(a + "endParaRPr");
            if (endProperties is not null)
            {
                endProperties.AddBeforeSelf(run);
            }
            else
            {
                paragraphForRun.Add(run);
            }

            textRuns = cell.Descendants(a + "t").ToList();
        }

        textRuns[0].Value = value;
        for (var index = 1; index < textRuns.Count; index++)
        {
            textRuns[index].Value = "";
        }

        var paragraph = cell.Descendants(a + "p").FirstOrDefault();
        if (paragraph is null)
        {
            return;
        }

        var paragraphProperties = paragraph.Element(a + "pPr");
        if (paragraphProperties is null)
        {
            paragraphProperties = new XElement(a + "pPr");
            paragraph.AddFirst(paragraphProperties);
        }

        paragraphProperties.SetAttributeValue("algn", alignment);
    }

    private static XElement BuildTableCellTextRun(XElement paragraph, XNamespace a)
    {
        var endProperties = paragraph.Element(a + "endParaRPr");
        var runProperties = new XElement(a + "rPr");
        if (endProperties is not null)
        {
            foreach (var attribute in endProperties.Attributes())
            {
                runProperties.SetAttributeValue(attribute.Name, attribute.Value);
            }

            foreach (var child in endProperties.Elements())
            {
                runProperties.Add(new XElement(child));
            }
        }
        else
        {
            runProperties.SetAttributeValue("lang", "en-US");
            runProperties.SetAttributeValue("sz", "1200");
        }

        return new XElement(
            a + "r",
            runProperties,
            new XElement(a + "t", ""));
    }

    private static bool UpdateActivityChartBySlideTitle(
        ZipArchive archive,
        string slideTitle,
        string sheetName,
        List<ActivityChartRow> rows)
    {
        var chartPath = FindChartPathBySlideTitle(archive, slideTitle);
        if (chartPath is null)
        {
            return false;
        }

        UpdateActivityChartXml(archive, chartPath, sheetName, rows);

        var embeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, chartPath);
        if (embeddedWorkbookPath is not null)
        {
            UpdateEmbeddedActivityWorkbook(archive, embeddedWorkbookPath, sheetName, rows);
        }

        return true;
    }

    private static void UpdateActivityChartXml(ZipArchive archive, string chartPath, string sheetName, List<ActivityChartRow> rows)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
        var allSeries = document.Descendants(c + "ser").ToList();
        var series = allSeries.Take(3).ToList();
        if (series.Count < 3)
        {
            return;
        }

        foreach (var extraSeries in allSeries.Skip(3).ToList())
        {
            extraSeries.Remove();
        }

        var categories = rows.Select(row => row.Label).ToList();
        UpdateActivitySeries(series[0], c, sheetName, "B", "Net Billed Premiums", categories, rows.Select(row => row.BilledPremium).ToList(), "$#,##0");
        UpdateActivitySeries(series[1], c, sheetName, "C", "Paid Claims", categories, rows.Select(row => row.PaidClaims).ToList(), "$#,##0");
        UpdateActivitySeries(series[2], c, sheetName, "D", "Incurred Claims", categories, rows.Select(row => row.IncurredClaims).ToList(), "$#,##0");

        entry.Delete();
        var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
    }

    private static void UpdateActivitySeries(
        XElement series,
        XNamespace c,
        string sheetName,
        string column,
        string seriesLabel,
        IReadOnlyList<string> categories,
        IReadOnlyList<decimal> values,
        string formatCode)
    {
        var escapedSheetName = $"'{sheetName.Replace("'", "''", StringComparison.Ordinal)}'";
        SetStringReference(series.Element(c + "tx")?.Element(c + "strRef"), c, $"{escapedSheetName}!${column}$2", [seriesLabel]);
        SetStringReference(series.Element(c + "cat")?.Element(c + "strRef"), c, $"{escapedSheetName}!$A$3:$A${categories.Count + 2}", categories);
        SetNumberReference(series.Element(c + "val")?.Element(c + "numRef"), c, $"{escapedSheetName}!${column}$3:${column}${values.Count + 2}", values, formatCode);
    }

    private static void UpdateEmbeddedActivityWorkbook(ZipArchive archive, string workbookPath, string sheetName, List<ActivityChartRow> rows)
    {
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add(sheetName.Length > 31 ? sheetName[..31] : sheetName);

        sheet.Cell(1, 1).Value = sheetName == "EHC Analysis"
            ? "Non Pooled Extended Health Care - Subtotal"
            : "Dental Analysis";
        sheet.Cell(2, 2).Value = "Net Billed Premiums";
        sheet.Cell(2, 3).Value = "Paid Claims";
        sheet.Cell(2, 4).Value = "Incurred Claims";
        sheet.Cell(2, 5).Value = "Paid Loss Ratios";

        for (var index = 0; index < rows.Count; index++)
        {
            var rowNumber = index + 3;
            sheet.Cell(rowNumber, 1).Value = rows[index].Label;
            sheet.Cell(rowNumber, 2).Value = rows[index].BilledPremium;
            sheet.Cell(rowNumber, 3).Value = rows[index].PaidClaims;
            sheet.Cell(rowNumber, 4).Value = rows[index].IncurredClaims;
            sheet.Cell(rowNumber, 5).Value = rows[index].PaidLossRatio;
        }

        sheet.Range("B2:E2").Style.Font.Bold = true;
        sheet.Range(3, 2, rows.Count + 2, 4).Style.NumberFormat.Format = "$#,##0";
        sheet.Range(3, 5, rows.Count + 2, 5).Style.NumberFormat.Format = "0.00%";
        sheet.Columns(1, 5).AdjustToContents(8, 26);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;

        entry.Delete();
        var newEntry = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var outputStream = newEntry.Open();
        output.CopyTo(outputStream);
    }

    private static bool UpdateEhcSpendingTrendChartBySlideTitle(
        ZipArchive archive,
        List<SpendingTrendChartRow> rows,
        params string[] slideTitles)
    {
        var updated = false;
        foreach (var chartPath in FindChartPathsBySlideTitle(archive, slideTitles))
        {
            if (!UpdateEhcSpendingTrendChartXml(archive, chartPath, rows))
            {
                continue;
            }

            var embeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, chartPath);
            if (embeddedWorkbookPath is not null)
            {
                UpdateEmbeddedEhcSpendingTrendWorkbook(archive, embeddedWorkbookPath, rows);
            }

            updated = true;
        }

        return updated;
    }

    private static bool UpdateEhcSpendingTrendChartXml(ZipArchive archive, string chartPath, List<SpendingTrendChartRow> rows)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
        var barChart = document.Descendants(c + "barChart").FirstOrDefault();
        var allSeries = (barChart?.Elements(c + "ser") ?? document.Descendants(c + "ser")).ToList();
        var labels = GetSpendingTrendLabels(rows);
        if (allSeries.Count == 0 || labels.Count == 0)
        {
            return false;
        }

        while (allSeries.Count < labels.Count)
        {
            var clone = new XElement(allSeries.Last());
            allSeries.Last().AddAfterSelf(clone);
            allSeries = (barChart?.Elements(c + "ser") ?? document.Descendants(c + "ser")).ToList();
        }

        for (var index = 0; index < allSeries.Count; index++)
        {
            allSeries[index].Element(c + "idx")?.SetAttributeValue("val", index);
            allSeries[index].Element(c + "order")?.SetAttributeValue("val", index);
        }

        var series = allSeries.Take(labels.Count).ToList();
        foreach (var extraSeries in allSeries.Skip(labels.Count).ToList())
        {
            extraSeries.Remove();
        }

        var chart = document.Descendants(c + "chart").FirstOrDefault();
        chart?.Element(c + "title")?.Remove();
        EnsureChartDataTableWithoutSeparateLegend(document, c);
        var categories = rows.Select(row => row.Category).ToList();
        for (var index = 0; index < labels.Count; index++)
        {
            var column = ((char)('B' + index)).ToString();
            var label = labels[index];
            UpdateSpendingTrendSeries(series[index], c, column, label, categories, rows.Select(row => GetSpendingTrendAmount(row, label)).ToList());
        }

        entry.Delete();
        var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
        return true;
    }

    private static void EnsureChartDataTableWithoutSeparateLegend(XDocument document, XNamespace c)
    {
        var plotArea = document.Descendants(c + "plotArea").FirstOrDefault();
        if (plotArea is not null && plotArea.Element(c + "dTable") is null)
        {
            plotArea.Add(
                new XElement(
                    c + "dTable",
                    new XElement(c + "showHorzBorder", new XAttribute("val", "1")),
                    new XElement(c + "showVertBorder", new XAttribute("val", "1")),
                    new XElement(c + "showOutline", new XAttribute("val", "1")),
                    new XElement(c + "showKeys", new XAttribute("val", "1"))));
        }

        var chart = document.Descendants(c + "chart").FirstOrDefault();
        chart?.Element(c + "legend")?.Remove();
    }

    private static void UpdateSpendingTrendSeries(
        XElement series,
        XNamespace c,
        string column,
        string seriesLabel,
        IReadOnlyList<string> categories,
        IReadOnlyList<decimal> values)
    {
        SetStringReference(series.Element(c + "tx")?.Element(c + "strRef"), c, $"Sheet1!${column}$1", [seriesLabel]);
        SetStringReference(series.Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${categories.Count + 1}", categories);
        SetNumberReference(series.Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!${column}$2:${column}${values.Count + 1}", values, "$#,##0");
    }

    private static void UpdateEmbeddedEhcSpendingTrendWorkbook(ZipArchive archive, string workbookPath, List<SpendingTrendChartRow> rows)
    {
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        var labels = GetSpendingTrendLabels(rows);
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "Service Category";
        for (var index = 0; index < labels.Count; index++)
        {
            sheet.Cell(1, index + 2).Value = labels[index];
        }
        sheet.Range(1, 1, 1, labels.Count + 1).Style.Font.Bold = true;

        for (var index = 0; index < rows.Count; index++)
        {
            var row = index + 2;
            sheet.Cell(row, 1).Value = rows[index].Category;
            for (var periodIndex = 0; periodIndex < labels.Count; periodIndex++)
            {
                sheet.Cell(row, periodIndex + 2).Value = GetSpendingTrendAmount(rows[index], labels[periodIndex]);
            }
        }

        sheet.Range(2, 2, rows.Count + 1, labels.Count + 1).Style.NumberFormat.Format = "$#,##0";
        sheet.Columns(1, labels.Count + 1).AdjustToContents(8, 26);
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;
        entry.Delete();
        var newEntry = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var outputStream = newEntry.Open();
        output.CopyTo(outputStream);
    }

    private static bool UpdateDemographicChartBySlideTitle(
        ZipArchive archive,
        string slideTitle,
        string priorYear,
        string currentYear,
        List<DemographicChartRow> rows)
    {
        var chartPaths = FindChartPathsBySlideTitle(archive, slideTitle);
        if (chartPaths.Count == 0)
        {
            return false;
        }

        foreach (var chartPath in chartPaths)
        {
            UpdateChartXml(archive, chartPath, priorYear, currentYear, rows);

            var embeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, chartPath);
            if (embeddedWorkbookPath is not null)
            {
                UpdateEmbeddedChartWorkbook(archive, embeddedWorkbookPath, priorYear, currentYear, rows);
            }
        }

        return true;
    }

    private static string? FindChartPathBySlideTitle(ZipArchive archive, string slideTitle)
    {
        return FindChartPathsBySlideTitle(archive, slideTitle).FirstOrDefault();
    }

    private static List<string> FindChartPathsBySlideTitle(ZipArchive archive, params string[] slideTitles)
    {
        var normalizedSlideTitles = slideTitles
            .Where(title => !string.IsNullOrWhiteSpace(title))
            .Select(NormalizeSlideLookupText)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (normalizedSlideTitles.Count == 0)
        {
            return [];
        }

        var chartPaths = new List<string>();
        foreach (var slideEntry in archive.Entries
                     .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                                     entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                     .OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase))
        {
            var xml = ReadZipEntryText(slideEntry);
            var slideText = ExtractSlideText(xml);
            var normalizedSlideText = NormalizeSlideLookupText(slideText);
            if (!normalizedSlideTitles.Any(title =>
                    xml.Contains(title, StringComparison.OrdinalIgnoreCase) ||
                    normalizedSlideText.Contains(title, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            var relsPath = $"ppt/slides/_rels/{Path.GetFileName(slideEntry.FullName)}.rels";
            var relsEntry = archive.GetEntry(relsPath);
            if (relsEntry is null)
            {
                continue;
            }

            var rels = XDocument.Parse(ReadZipEntryText(relsEntry));
            foreach (var relationship in rels.Root?.Elements() ?? Enumerable.Empty<XElement>())
            {
                if (((string?)relationship.Attribute("Type"))?.EndsWith("/chart", StringComparison.OrdinalIgnoreCase) != true)
                {
                    continue;
                }

                var target = (string?)relationship.Attribute("Target");
                if (string.IsNullOrWhiteSpace(target))
                {
                    continue;
                }

                chartPaths.Add(ResolvePackagePath(slideEntry.FullName, target));
            }
        }

        return chartPaths
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string? FindChartPathBySlideTitleLegacy(ZipArchive archive, string slideTitle)
    {
        var normalizedSlideTitle = NormalizeSlideLookupText(slideTitle);
        foreach (var slideEntry in archive.Entries
                     .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                                     entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                     .OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase))
        {
            var xml = ReadZipEntryText(slideEntry);
            var slideText = ExtractSlideText(xml);
            if (!xml.Contains(slideTitle, StringComparison.OrdinalIgnoreCase) &&
                !NormalizeSlideLookupText(slideText).Contains(normalizedSlideTitle, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var relsPath = $"ppt/slides/_rels/{Path.GetFileName(slideEntry.FullName)}.rels";
            var relsEntry = archive.GetEntry(relsPath);
            if (relsEntry is null)
            {
                continue;
            }

            var rels = XDocument.Parse(ReadZipEntryText(relsEntry));
            var relationship = rels.Root?
                .Elements()
                .FirstOrDefault(element => ((string?)element.Attribute("Type"))?.EndsWith("/chart", StringComparison.OrdinalIgnoreCase) == true);

            var target = (string?)relationship?.Attribute("Target");
            if (string.IsNullOrWhiteSpace(target))
            {
                continue;
            }

            return ResolvePackagePath(slideEntry.FullName, target);
        }

        return null;
    }

    private static string ExtractSlideText(string xml)
    {
        try
        {
            var document = XDocument.Parse(xml);
            var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
            return string.Join(" ", document.Descendants(a + "t").Select(element => element.Value));
        }
        catch
        {
            return xml;
        }
    }

    private static string NormalizeSlideLookupText(string text)
    {
        var normalized = text
            .Replace("\u2013", "-", StringComparison.Ordinal)
            .Replace("\u2014", "-", StringComparison.Ordinal)
            .Replace("-", " ", StringComparison.Ordinal);

        return Regex.Replace(normalized, @"\s+", " ").Trim().ToUpperInvariant();
    }

    private static string? FindEmbeddedWorkbookPath(ZipArchive archive, string chartPath)
    {
        var chartFileName = Path.GetFileName(chartPath);
        var relsPath = $"ppt/charts/_rels/{chartFileName}.rels";
        var relsEntry = archive.GetEntry(relsPath);
        if (relsEntry is null)
        {
            return null;
        }

        var rels = XDocument.Parse(ReadZipEntryText(relsEntry));
        var relationship = rels.Root?
            .Elements()
            .FirstOrDefault(element =>
            {
                var type = (string?)element.Attribute("Type") ?? "";
                var target = (string?)element.Attribute("Target") ?? "";
                return type.EndsWith("/package", StringComparison.OrdinalIgnoreCase) &&
                       target.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
            });

        var targetPath = (string?)relationship?.Attribute("Target");
        return string.IsNullOrWhiteSpace(targetPath)
            ? null
            : ResolvePackagePath(chartPath, targetPath);
    }

    private static void UpdateChartXml(
        ZipArchive archive,
        string chartPath,
        string priorYear,
        string currentYear,
        List<DemographicChartRow> rows)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
        var series = document.Descendants(c + "ser").Take(2).ToList();
        if (series.Count < 2)
        {
            return;
        }

        var categories = rows.Select(row => row.AgeBand).ToList();
        UpdateChartSeries(series[0], c, "B", priorYear, categories, rows.Select(row => row.PriorVolume).ToList());
        UpdateChartSeries(series[1], c, "C", currentYear, categories, rows.Select(row => row.CurrentVolume).ToList());

        entry.Delete();
        var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
    }

    private static void UpdateChartSeries(
        XElement series,
        XNamespace c,
        string column,
        string seriesLabel,
        IReadOnlyList<string> categories,
        IReadOnlyList<decimal> values)
    {
        SetStringReference(series.Element(c + "tx")?.Element(c + "strRef"), c, $"Sheet1!${column}$1", [seriesLabel]);
        SetStringReference(series.Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${categories.Count + 1}", categories);
        SetNumberReference(series.Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!${column}$2:${column}${values.Count + 1}", values);
    }

    private static void SetStringReference(XElement? stringReference, XNamespace c, string formula, IReadOnlyList<string> values)
    {
        if (stringReference is null)
        {
            return;
        }

        SetFormula(stringReference, c, formula);
        var cache = EnsureChild(stringReference, c + "strCache");
        ReplaceCachePoints(cache, c, values.Select(value => (object)value).ToList());
    }

    private static void SetNumberReference(XElement? numberReference, XNamespace c, string formula, IReadOnlyList<decimal> values)
    {
        SetNumberReference(numberReference, c, formula, values, "\"$\"#,##0");
    }

    private static void SetNumberReference(XElement? numberReference, XNamespace c, string formula, IReadOnlyList<decimal> values, string formatCodeValue)
    {
        if (numberReference is null)
        {
            return;
        }

        SetFormula(numberReference, c, formula);
        var cache = EnsureChild(numberReference, c + "numCache");
        var formatCode = EnsureChild(cache, c + "formatCode");
        formatCode.Value = formatCodeValue;
        ReplaceCachePoints(cache, c, values.Select(value => (object)value).ToList());
    }

    private static void SetFormula(XElement reference, XNamespace c, string formula)
    {
        var formulaElement = reference.Element(c + "f");
        if (formulaElement is null)
        {
            formulaElement = new XElement(c + "f");
            reference.AddFirst(formulaElement);
        }

        formulaElement.Value = formula;
    }

    private static XElement EnsureChild(XElement parent, XName name)
    {
        var child = parent.Element(name);
        if (child is not null)
        {
            return child;
        }

        child = new XElement(name);
        parent.Add(child);
        return child;
    }

    private static void ReplaceCachePoints(XElement cache, XNamespace c, IReadOnlyList<object> values)
    {
        cache.Elements(c + "ptCount").Remove();
        cache.Elements(c + "pt").Remove();
        cache.AddFirst(new XElement(c + "ptCount", new XAttribute("val", values.Count)));

        for (var index = 0; index < values.Count; index++)
        {
            var value = values[index] is decimal decimalValue
                ? decimalValue.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)
                : values[index].ToString() ?? "";

            cache.Add(new XElement(
                c + "pt",
                new XAttribute("idx", index),
                new XElement(c + "v", value)));
        }
    }

    private static void UpdateEmbeddedChartWorkbook(
        ZipArchive archive,
        string workbookPath,
        string priorYear,
        string currentYear,
        List<DemographicChartRow> rows)
    {
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");

        sheet.Cell(1, 1).Value = "Age";
        sheet.Cell(1, 2).Value = priorYear;
        sheet.Cell(1, 3).Value = currentYear;

        for (var row = 2; row <= Math.Max(30, rows.Count + 5); row++)
        {
            for (var column = 1; column <= 3; column++)
            {
                sheet.Cell(row, column).Value = "";
            }
        }

        sheet.Range(1, 2, 1, 3).Style.Font.Bold = true;

        for (var index = 0; index < rows.Count; index++)
        {
            var row = index + 2;
            sheet.Cell(row, 1).Value = rows[index].AgeBand;
            sheet.Cell(row, 2).Value = rows[index].PriorVolume;
            sheet.Cell(row, 3).Value = rows[index].CurrentVolume;
        }

        var lastDataRow = rows.Count + 1;
        var totalRow = rows.Count + 3;
        sheet.Cell(totalRow, 1).Value = "Total";
        sheet.Cell(totalRow, 2).FormulaA1 = $"SUM(B2:B{lastDataRow})";
        sheet.Cell(totalRow, 3).FormulaA1 = $"SUM(C2:C{lastDataRow})";
        sheet.Range(totalRow, 1, totalRow, 3).Style.Font.Bold = true;

        sheet.Range(2, 2, totalRow, 3).Style.NumberFormat.Format = "$#,##0";
        sheet.Columns(1, 3).AdjustToContents(8, 18);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;

        entry.Delete();
        var newEntry = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var outputStream = newEntry.Open();
        output.CopyTo(outputStream);
    }

    private static string ReadZipEntryText(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static string ResolvePackagePath(string sourcePartPath, string target)
    {
        if (target.StartsWith("/", StringComparison.Ordinal))
        {
            return target.TrimStart('/');
        }

        var segments = sourcePartPath.Split('/', StringSplitOptions.RemoveEmptyEntries)
            .SkipLast(1)
            .ToList();

        foreach (var segment in target.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment.Equals(".", StringComparison.Ordinal))
            {
                continue;
            }

            if (segment.Equals("..", StringComparison.Ordinal))
            {
                if (segments.Count > 0)
                {
                    segments.RemoveAt(segments.Count - 1);
                }

                continue;
            }

            segments.Add(segment);
        }

        return string.Join("/", segments);
    }

    private static List<DemographicChartRow> BuildLifeChartRows(IReadOnlyCollection<DemographicLine> demographics)
    {
        var usesCombinedSeniorBucket = DemographicAgeBandNormalizer.UsesCombinedLifeSeniorBucket(demographics);
        var buckets = DemographicAgeBandNormalizer.GetLifeAgeBands(demographics);

        return BuildChartRows(
            demographics,
            buckets,
            ageBand => DemographicAgeBandNormalizer.SelectLifeChartBucket(ageBand, usesCombinedSeniorBucket));
    }

    private static List<ActivityChartRow> BuildActivityChartRows(AnalysisResult analysis, params string[] benefitAliases)
    {
        var rows = analysis.Pdf.ExperienceLines
            .Where(row => MatchesActivityBenefit(row.Benefit, benefitAliases))
            .Where(IsActivityExperiencePeriod)
            .Select(row => new
            {
                Row = row,
                PeriodStart = ParseExperiencePeriodStart(row.Period) ?? DateOnly.MinValue
            })
            .GroupBy(item => NormalizeActivityText(item.Row.Period), StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                // A period can be represented by multiple carrier/supporting rows. Keep the
                // preferred benefit label, then fill any blank metric from its companion row.
                var candidates = group
                    .OrderBy(item =>
                    {
                        var normalizedBenefit = NormalizeActivityText(item.Row.Benefit);
                        var aliasIndex = Array.FindIndex(
                            benefitAliases,
                            alias => normalizedBenefit.Equals(NormalizeActivityText(alias), StringComparison.Ordinal));
                        return aliasIndex < 0 ? benefitAliases.Length : aliasIndex;
                    })
                    .ThenBy(item => IsPreviousRenewalExperience(item.Row) ? 1 : 0)
                    .ToList();
                var billedPremium = candidates.Select(item => item.Row.BilledPremium).FirstOrDefault(value => value.HasValue);
                var paidClaims = candidates.Select(item => item.Row.PaidClaims).FirstOrDefault(value => value.HasValue);
                var incurredClaims = candidates.Select(item => item.Row.IncurredClaims).FirstOrDefault(value => value.HasValue);
                var paidLossRatio = candidates.Select(item => item.Row.PaidLossRatio).FirstOrDefault(value => value.HasValue);

                return new
                {
                    candidates[0].PeriodStart,
                    BilledPremium = billedPremium ?? 0m,
                    PaidClaims = paidClaims ?? 0m,
                    IncurredClaims = incurredClaims ?? 0m,
                    PaidLossRatio = (paidLossRatio ?? CalculateLossRatio(paidClaims ?? 0m, billedPremium ?? 0m) * 100m) / 100m
                };
            })
            .OrderBy(item => item.PeriodStart)
            .TakeLast(3)
            .ToList();
        var labels = new[] { "Prior 2 Periods", "Last Period", "Most Current" };
        var result = new List<ActivityChartRow>();
        var labelOffset = Math.Max(0, labels.Length - rows.Count);

        for (var index = 0; index < rows.Count; index++)
        {
            result.Add(new ActivityChartRow(
                labels[index + labelOffset],
                rows[index].BilledPremium,
                rows[index].PaidClaims,
                rows[index].IncurredClaims,
                rows[index].PaidLossRatio));
        }

        return result;
    }

    private static ProgramClaimsSummaryChartData BuildProgramClaimsSummaryChartData(AnalysisResult analysis)
    {
        var candidateRows = analysis.Pdf.ExperienceLines
            .Where(IsActivityExperiencePeriod)
            .Select(row => new
            {
                Row = row,
                PeriodStart = ParseExperiencePeriodStart(row.Period)
            })
            .Where(item => item.PeriodStart.HasValue)
            .ToList();

        var reportingPeriod = candidateRows
            .OrderByDescending(item => item.PeriodStart!.Value)
            .Select(item => item.Row.Period)
            .FirstOrDefault() ?? "";
        var normalizedReportingPeriod = NormalizeActivityText(reportingPeriod);

        var rows = new List<ProgramClaimsSummaryRow>
        {
            BuildProgramClaimsSummaryRow(
                "Extended Health",
                candidateRows.Select(item => item.Row),
                normalizedReportingPeriod,
                ["Extended Health Care", "Extended Health", "EHC"]),
            BuildProgramClaimsSummaryRow(
                "Dental",
                candidateRows.Select(item => item.Row),
                normalizedReportingPeriod,
                ["Dental", "Dental Care"]),
            BuildProgramClaimsSummaryRow(
                "STD",
                candidateRows.Select(item => item.Row),
                normalizedReportingPeriod,
                ["STD", "Short Term Disability", "Short-Term Disability"])
        };

        return new ProgramClaimsSummaryChartData(reportingPeriod, rows);
    }

    private static ProgramClaimsSummaryRow BuildProgramClaimsSummaryRow(
        string category,
        IEnumerable<ExperienceLine> rows,
        string normalizedReportingPeriod,
        IReadOnlyCollection<string> benefitAliases)
    {
        var matchingRow = rows
            .Where(row => MatchesActivityBenefit(row.Benefit, benefitAliases))
            .Where(row => NormalizeActivityText(row.Period).Equals(normalizedReportingPeriod, StringComparison.Ordinal))
            .OrderByDescending(row => row.BilledPremium ?? 0m)
            .FirstOrDefault();
        var paidPremiums = matchingRow?.BilledPremium ?? 0m;
        var paidClaims = matchingRow?.PaidClaims ?? 0m;

        return new ProgramClaimsSummaryRow(category, paidPremiums, paidClaims, CalculateLossRatio(paidClaims, paidPremiums));
    }

    private static decimal CalculateLossRatio(decimal paidClaims, decimal premium)
    {
        return premium == 0m ? 0m : paidClaims / premium;
    }

    private static bool MatchesActivityBenefit(string benefit, IReadOnlyCollection<string> aliases)
    {
        var normalizedBenefit = NormalizeActivityText(benefit);
        if (string.IsNullOrWhiteSpace(normalizedBenefit))
        {
            return false;
        }

        var normalizedAliases = aliases
            .Select(NormalizeActivityText)
            .Where(alias => !string.IsNullOrWhiteSpace(alias))
            .ToList();
        if (normalizedAliases.Any(alias => normalizedBenefit.Equals(alias, StringComparison.Ordinal)))
        {
            return true;
        }

        if (normalizedAliases.Any(alias => alias.Contains("dental", StringComparison.Ordinal)))
        {
            return normalizedBenefit.Contains("dental", StringComparison.Ordinal);
        }

        if (normalizedAliases.Any(alias => alias.Contains("extendedhealth", StringComparison.Ordinal) || alias.Equals("ehc", StringComparison.Ordinal)))
        {
            return normalizedBenefit.Contains("extendedhealth", StringComparison.Ordinal) ||
                   normalizedBenefit.Contains("ehc", StringComparison.Ordinal);
        }

        if (normalizedAliases.Any(alias => alias.Equals("std", StringComparison.Ordinal) || alias.Contains("shorttermdisability", StringComparison.Ordinal)))
        {
            return normalizedBenefit.Equals("std", StringComparison.Ordinal) ||
                   normalizedBenefit.Contains("shorttermdisability", StringComparison.Ordinal);
        }

        return false;
    }

    private static bool IsActivityExperiencePeriod(ExperienceLine row)
    {
        return !string.IsNullOrWhiteSpace(row.Period) &&
               !row.Period.Contains("Weighted", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsPreviousRenewalExperience(ExperienceLine row)
    {
        return row.SourceReference.Contains("Previous renewal PDF", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeActivityText(string? value)
    {
        return Regex.Replace(value ?? "", "[^a-zA-Z0-9]+", "").ToLowerInvariant();
    }

    private static DateOnly? ParseExperiencePeriodStart(string? period)
    {
        if (string.IsNullOrWhiteSpace(period))
        {
            return null;
        }

        var cleaned = ParsingHelpers.Clean(period);
        var firstRangePart = cleaned.Split(" - ", StringSplitOptions.None).FirstOrDefault();
        var directDate = ParsingHelpers.ParseDateOnly(firstRangePart);
        if (directDate.HasValue)
        {
            return directDate;
        }

        var dateMatch = Regex.Match(
            cleaned,
            @"\b(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Sept(?:ember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\.?\s+\d{1,2},?\s+\d{4}\b",
            RegexOptions.IgnoreCase);
        if (dateMatch.Success &&
            DateTime.TryParse(dateMatch.Value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
        {
            return DateOnly.FromDateTime(parsed);
        }

        return null;
    }

    private static List<SpendingTrendChartRow> BuildEhcSpendingTrendChartRows(AnalysisResult analysis)
    {
        return BuildSpendingTrendChartRows(analysis.Workbook.EhcSpendingTrends);
    }

    private static List<SpendingTrendChartRow> BuildDentalSpendingTrendChartRows(AnalysisResult analysis)
    {
        var rows = analysis.Workbook.DentalSpendingTrends.ToList();
        if (rows.Count == 0)
        {
            return [];
        }

        var priorLabel = rows.FirstOrDefault(row => !string.IsNullOrWhiteSpace(row.PriorPeriodLabel))?.PriorPeriodLabel ?? "Prior Period";
        var currentLabel = rows.FirstOrDefault(row => !string.IsNullOrWhiteSpace(row.CurrentPeriodLabel))?.CurrentPeriodLabel ?? "Current Period";

        if (rows.Any(row => DentalCategoryMatches(row.Category, ["Basic", "Major"])))
        {
            return
            [
                BuildDentalClaimsAnalysisChartRow(
                    "Basic",
                    priorLabel,
                    currentLabel,
                    rows,
                    "Basic"),
                BuildDentalClaimsAnalysisChartRow(
                    "Periodontal & Endodontic",
                    priorLabel,
                    currentLabel,
                    rows,
                    "Periodontic",
                    "Periodontal",
                    "Endodontic",
                    "Endodontics",
                    "Periodontal & Endodontic",
                    "Periodontic & Endodontic"),
                BuildDentalClaimsAnalysisChartRow(
                    "Major",
                    priorLabel,
                    currentLabel,
                    rows,
                    "Major",
                    "Major Restorative"),
                BuildDentalClaimsAnalysisChartRow(
                    "Dentures",
                    priorLabel,
                    currentLabel,
                    rows,
                    "Dentures",
                    "Dentures,Bridges",
                    "Dentures, Bridges",
                    "Dentures / Bridges",
                    "Dentures and Bridges"),
                BuildDentalClaimsAnalysisChartRow(
                    "Orthodontic",
                    priorLabel,
                    currentLabel,
                    rows,
                    "Orthodontic",
                    "Orthodontics")
            ];
        }

        return
        [
            BuildDentalClaimsAnalysisChartRow(
                "Preventative & Diagnostic",
                priorLabel,
                currentLabel,
                rows,
                "Diagnostic",
                "Preventative",
                "Preventive",
                "Preventative & Diagnostic",
                "Preventive & Diagnostic"),
            BuildDentalClaimsAnalysisChartRow(
                "Periodontal & Endodontic",
                priorLabel,
                currentLabel,
                rows,
                "Periodontic",
                "Periodontal",
                "Endodontic",
                "Endodontics",
                "Periodontal & Endodontic",
                "Periodontic & Endodontic"),
            BuildDentalClaimsAnalysisChartRow(
                "Major Restorative",
                priorLabel,
                currentLabel,
                rows,
                "Major Restorative",
                "Dentures,Bridges",
                "Dentures, Bridges",
                "Dentures / Bridges",
                "Dentures and Bridges"),
            BuildDentalClaimsAnalysisChartRow(
                "Surgical & Restorative",
                priorLabel,
                currentLabel,
                rows,
                "Oral Surgery",
                "Minor Restorative",
                "Surgical & Restorative"),
            BuildDentalClaimsAnalysisChartRow(
                "Orthodontic",
                priorLabel,
                currentLabel,
                rows,
                "Orthodontic",
                "Orthodontics")
        ];
    }

    private static SpendingTrendChartRow BuildDentalClaimsAnalysisChartRow(
        string category,
        string priorLabel,
        string currentLabel,
        IReadOnlyCollection<ServiceCategoryTrend> rows,
        params string[] aliases)
    {
        var matchingRows = rows
            .Where(row => DentalCategoryMatches(row.Category, aliases))
            .ToList();

        return new SpendingTrendChartRow(
            category,
            priorLabel,
            matchingRows.Sum(row => row.PriorAmount),
            currentLabel,
            matchingRows.Sum(row => row.CurrentAmount),
                AggregateSpendingTrendPeriods(matchingRows));
    }

    private static bool DentalCategoryMatches(string category, IReadOnlyCollection<string> aliases)
    {
        var normalized = NormalizeDentalCategory(category);
        return aliases
            .Select(NormalizeDentalCategory)
            .Any(alias => normalized.Equals(alias, StringComparison.Ordinal));
    }

    private static string NormalizeDentalCategory(string category)
    {
        return new string(ParsingHelpers.Clean(category)
            .Where(char.IsLetterOrDigit)
            .Select(char.ToUpperInvariant)
            .ToArray());
    }

    private static List<SpendingTrendChartRow> BuildParamedicalServiceTrendChartRows(AnalysisResult analysis)
    {
        return analysis.Workbook.ParamedicalServiceTrends
            .Select(row => new SpendingTrendChartRow(
                MapParamedicalServiceChartLabel(row.Category),
                string.IsNullOrWhiteSpace(row.PriorPeriodLabel) ? "Prior Period" : row.PriorPeriodLabel,
                row.PriorAmount,
                string.IsNullOrWhiteSpace(row.CurrentPeriodLabel) ? "Current Period" : row.CurrentPeriodLabel,
                row.CurrentAmount,
                ServiceCategoryHistory.GetPeriods(row)))
            .ToList();
    }

    private static List<SpendingTrendChartRow> BuildSpendingTrendChartRows(IEnumerable<ServiceCategoryTrend> trends)
    {
        return trends
            .Where(row => !row.Category.Equals("Total", StringComparison.OrdinalIgnoreCase))
            .Select(row => new SpendingTrendChartRow(
                row.Category,
                string.IsNullOrWhiteSpace(row.PriorPeriodLabel) ? "Prior Period" : row.PriorPeriodLabel,
                row.PriorAmount,
                string.IsNullOrWhiteSpace(row.CurrentPeriodLabel) ? "Current Period" : row.CurrentPeriodLabel,
                row.CurrentAmount,
                ServiceCategoryHistory.GetPeriods(row)))
            .ToList();
    }

    private static IReadOnlyList<ServiceCategoryPeriodValue> AggregateSpendingTrendPeriods(IEnumerable<ServiceCategoryTrend> rows)
    {
        return rows.SelectMany(row => ServiceCategoryHistory.GetPeriods(row))
            .GroupBy(period => period.PeriodLabel, StringComparer.OrdinalIgnoreCase)
            .Select(group => new ServiceCategoryPeriodValue
            {
                PeriodLabel = group.Key,
                Amount = group.Sum(period => period.Amount),
                SourceReference = string.Join("; ", group.Select(period => period.SourceReference).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase))
            })
            .OrderBy(period => ServiceCategoryHistory.ParsePeriodStart(period.PeriodLabel) ?? DateOnly.MinValue)
            .TakeLast(3)
            .ToList();
    }

    private static List<string> GetSpendingTrendLabels(IEnumerable<SpendingTrendChartRow> rows)
    {
        return rows.SelectMany(GetSpendingTrendPeriods)
            .GroupBy(period => period.PeriodLabel, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(period => ServiceCategoryHistory.ParsePeriodStart(period.PeriodLabel) ?? DateOnly.MinValue)
            .TakeLast(3)
            .Select(period => period.PeriodLabel)
            .ToList();
    }

    private static IReadOnlyList<ServiceCategoryPeriodValue> GetSpendingTrendPeriods(SpendingTrendChartRow row)
    {
        if (row.History is { Count: > 0 }) return row.History;
        return
        [
            new ServiceCategoryPeriodValue { PeriodLabel = row.PriorLabel, Amount = row.PriorAmount },
            new ServiceCategoryPeriodValue { PeriodLabel = row.CurrentLabel, Amount = row.CurrentAmount }
        ];
    }

    private static decimal GetSpendingTrendAmount(SpendingTrendChartRow row, string label)
    {
        return GetSpendingTrendPeriods(row)
            .FirstOrDefault(period => period.PeriodLabel.Equals(label, StringComparison.OrdinalIgnoreCase))?.Amount ?? 0m;
    }

    private static string MapParamedicalServiceChartLabel(string service)
    {
        if (service.Equals("Registered Massage Therapist", StringComparison.OrdinalIgnoreCase))
        {
            return "Massage Therapist";
        }

        if (service.Equals("Physiotherapist", StringComparison.OrdinalIgnoreCase))
        {
            return "Physiotherapy";
        }

        if (service.Equals("Clinical Counsellor", StringComparison.OrdinalIgnoreCase) ||
            service.Equals("Clinical Counselor", StringComparison.OrdinalIgnoreCase))
        {
            return "Social Worker";
        }

        return service;
    }

    private static List<DemographicChartRow> BuildLtdChartRows(IReadOnlyCollection<DemographicLine> demographics)
    {
        var buckets = demographics.Any(row => row.AgeBand.Equals("60 & over", StringComparison.OrdinalIgnoreCase))
            ? new[] { "Under 25", "25-29", "30-34", "35-39", "40-44", "45-49", "50-54", "55-59", "60 & over" }
            : new[] { "Under 25", "25-29", "30-34", "35-39", "40-44", "45-49", "50-54", "55-59", "60-64", "65-69+" };

        return BuildChartRows(demographics, buckets, ageBand => ageBand is "65-69" or "70 & over" ? "65-69+" : ageBand);
    }

    private static List<DemographicChartRow> BuildChartRows(
        IReadOnlyCollection<DemographicLine> demographics,
        string[] buckets,
        Func<string, string> bucketSelector)
    {
        var results = buckets.ToDictionary(
            bucket => bucket,
            bucket => new DemographicChartRow(bucket, 0m, 0m),
            StringComparer.OrdinalIgnoreCase);

        foreach (var row in demographics)
        {
            var bucket = bucketSelector(row.AgeBand);
            if (!results.TryGetValue(bucket, out var existing))
            {
                continue;
            }

            var current = (row.CurrentFemaleVolume ?? 0m) + (row.CurrentMaleVolume ?? 0m);
            var prior = (row.PriorFemaleVolume ?? 0m) + (row.PriorMaleVolume ?? 0m);
            results[bucket] = existing with
            {
                PriorVolume = existing.PriorVolume + prior,
                CurrentVolume = existing.CurrentVolume + current
            };
        }

        return buckets
            .Select(bucket => results[bucket])
            .Where(row => row.PriorVolume != 0m || row.CurrentVolume != 0m || demographics.Count > 0)
            .ToList();
    }

    private static void AddDemographicAnalysisTokens(
        Dictionary<string, string> tokens,
        string prefix,
        IReadOnlyCollection<DemographicLine> demographics,
        IReadOnlyCollection<UnderwritingReferenceRate> underwritingReferenceRates,
        string underwritingRateType,
        bool addGenericCurrentAverageAgeToken = false)
    {
        var currentFemaleLives = demographics.Sum(row => row.CurrentFemaleLives ?? 0);
        var currentMaleLives = demographics.Sum(row => row.CurrentMaleLives ?? 0);
        var currentTotalLives = currentFemaleLives + currentMaleLives;
        var priorTotalLives = demographics.Sum(row => (row.PriorFemaleLives ?? 0) + (row.PriorMaleLives ?? 0));
        var currentFemalePercent = currentTotalLives == 0
            ? (decimal?)null
            : currentFemaleLives / (decimal)currentTotalLives * 100m;
        var currentMalePercent = currentTotalLives == 0
            ? (decimal?)null
            : currentMaleLives / (decimal)currentTotalLives * 100m;
        var currentAverageAge = CalculateWeightedAverageAge(
            demographics,
            row => (row.CurrentFemaleLives ?? 0) + (row.CurrentMaleLives ?? 0));
        var projectedRenewalAverageAge = currentAverageAge.HasValue
            ? currentAverageAge.Value + 1m
            : (decimal?)null;
        var averageAgeChange = currentAverageAge.HasValue && projectedRenewalAverageAge.HasValue
            ? projectedRenewalAverageAge.Value - currentAverageAge.Value
            : (decimal?)null;
        var riskChange = UnderwritingRiskCalculator.CalculateChange(
            demographics,
            underwritingReferenceRates,
            underwritingRateType);
        var totalLivesText = FormatWholeNumber(currentTotalLives == 0 && priorTotalLives == 0 ? null : currentTotalLives);
        var malePercentText = currentMalePercent.HasValue ? ParsingHelpers.Percent(currentMalePercent) : "";
        var femalePercentText = currentFemalePercent.HasValue ? ParsingHelpers.Percent(currentFemalePercent) : "";

        AddTokenVariants(tokens, $"{prefix} AnalysisCurrent Lives Combined Total", totalLivesText);
        AddTokenVariants(tokens, $"{prefix} Analysis Current Lives Combined Total", totalLivesText);
        AddTokenVariants(tokens, $"{prefix}_Analysis_CurrentLivesCombinedTotal", totalLivesText);

        AddTokenVariants(tokens, $"{prefix} Analysis Current Males %", malePercentText);
        AddTokenVariants(tokens, $"{prefix} Analysis Current Males Percent", malePercentText);
        AddTokenVariants(tokens, $"{prefix}_Analysis_CurrentMalesPercent", malePercentText);

        AddTokenVariants(tokens, $"{prefix} Analysis Current Females %", femalePercentText);
        AddTokenVariants(tokens, $"{prefix} Analysis Current Females Percent", femalePercentText);
        AddTokenVariants(tokens, $"{prefix}_Analysis_CurrentFemalesPercent", femalePercentText);

        AddTokenVariants(tokens, $"{prefix} Analysis CurrentAvgAge", FormatOneDecimal(currentAverageAge));
        AddTokenVariants(tokens, $"{prefix} Analysis Current Avg Age", FormatOneDecimal(currentAverageAge));
        AddTokenVariants(tokens, $"{prefix}_Analysis_CurrentAvgAge", FormatOneDecimal(currentAverageAge));

        AddTokenVariants(tokens, $"{prefix} Analysis_Changeinrisk", FormatRiskChange(riskChange));
        AddTokenVariants(tokens, $"{prefix} Analysis Change in risk", FormatRiskChange(riskChange));
        AddTokenVariants(tokens, $"{prefix}_Analysis_Changeinrisk", FormatRiskChange(riskChange));

        AddTokenVariants(tokens, $"{prefix} Analysis Renewal Avg Age", FormatOneDecimal(projectedRenewalAverageAge));
        AddTokenVariants(tokens, $"{prefix}_Analysis_RenewalAvgAge", FormatOneDecimal(projectedRenewalAverageAge));
        AddTokenVariants(tokens, $"{prefix} Analysis Avg Age Change", FormatOneDecimal(averageAgeChange));
        AddTokenVariants(tokens, $"{prefix}_Analysis_AvgAgeChange", FormatOneDecimal(averageAgeChange));

        if (addGenericCurrentAverageAgeToken)
        {
            AddTokenVariants(tokens, "Current Avg Age", FormatOneDecimal(currentAverageAge));
        }
    }

    private static void AddGenericDemographicAnalysisTokens(
        Dictionary<string, string> tokens,
        IReadOnlyCollection<DemographicLine> demographics,
        IReadOnlyCollection<UnderwritingReferenceRate> underwritingReferenceRates,
        string underwritingRateType)
    {
        var currentFemaleLives = demographics.Sum(row => row.CurrentFemaleLives ?? 0);
        var currentMaleLives = demographics.Sum(row => row.CurrentMaleLives ?? 0);
        var currentTotalLives = currentFemaleLives + currentMaleLives;
        var priorTotalLives = demographics.Sum(row => (row.PriorFemaleLives ?? 0) + (row.PriorMaleLives ?? 0));
        var currentFemalePercent = currentTotalLives == 0
            ? (decimal?)null
            : currentFemaleLives / (decimal)currentTotalLives * 100m;
        var currentMalePercent = currentTotalLives == 0
            ? (decimal?)null
            : currentMaleLives / (decimal)currentTotalLives * 100m;
        var currentAverageAge = CalculateWeightedAverageAge(
            demographics,
            row => (row.CurrentFemaleLives ?? 0) + (row.CurrentMaleLives ?? 0));
        var projectedRenewalAverageAge = currentAverageAge.HasValue
            ? currentAverageAge.Value + 1m
            : (decimal?)null;
        var averageAgeChange = currentAverageAge.HasValue && projectedRenewalAverageAge.HasValue
            ? projectedRenewalAverageAge.Value - currentAverageAge.Value
            : (decimal?)null;
        var riskChange = UnderwritingRiskCalculator.CalculateChange(
            demographics,
            underwritingReferenceRates,
            underwritingRateType);
        var totalLivesText = FormatWholeNumber(currentTotalLives == 0 && priorTotalLives == 0 ? null : currentTotalLives);

        AddTokenVariants(tokens, "Analysis Current Lives Combined Total", totalLivesText);
        AddTokenVariants(tokens, "Current Lives Combined Total", totalLivesText);
        AddTokenVariants(tokens, "Current Males %", currentMalePercent.HasValue ? ParsingHelpers.Percent(currentMalePercent) : "");
        AddTokenVariants(tokens, "Current Males Percent", currentMalePercent.HasValue ? ParsingHelpers.Percent(currentMalePercent) : "");
        AddTokenVariants(tokens, "Current Females %", currentFemalePercent.HasValue ? ParsingHelpers.Percent(currentFemalePercent) : "");
        AddTokenVariants(tokens, "Current Females Percent", currentFemalePercent.HasValue ? ParsingHelpers.Percent(currentFemalePercent) : "");
        AddTokenVariants(tokens, "CurrentAvgAge", FormatOneDecimal(currentAverageAge));
        AddTokenVariants(tokens, "Current Avg Age", FormatOneDecimal(currentAverageAge));
        AddTokenVariants(tokens, "Changeinrisk", FormatRiskChange(riskChange));
        AddTokenVariants(tokens, "Change in risk", FormatRiskChange(riskChange));
        AddTokenVariants(tokens, "Renewal Avg Age", FormatOneDecimal(projectedRenewalAverageAge));
        AddTokenVariants(tokens, "RenewalAvgAge", FormatOneDecimal(projectedRenewalAverageAge));
        AddTokenVariants(tokens, "Avg Age Change", FormatOneDecimal(averageAgeChange));
        AddTokenVariants(tokens, "AvgAgeChange", FormatOneDecimal(averageAgeChange));
    }

    private static decimal? CalculateWeightedAverageAge(
        IReadOnlyCollection<DemographicLine> demographics,
        Func<DemographicLine, int> livesSelector)
    {
        var weightedTotal = 0m;
        var totalLives = 0;

        foreach (var row in demographics)
        {
            var midpoint = GetAgeBandMidpoint(row.AgeBand);
            var lives = livesSelector(row);
            if (!midpoint.HasValue || lives <= 0)
            {
                continue;
            }

            weightedTotal += midpoint.Value * lives;
            totalLives += lives;
        }

        return totalLives == 0
            ? null
            : weightedTotal / totalLives;
    }

    private static decimal? GetAgeBandMidpoint(string ageBand)
    {
        var normalized = ParsingHelpers.Clean(ageBand).ToLowerInvariant();
        return normalized switch
        {
            "under 25" => 22m,
            "25-29" => 27m,
            "30-34" => 32m,
            "35-39" => 37m,
            "40-44" => 42m,
            "45-49" => 47m,
            "50-54" => 52m,
            "55-59" => 57m,
            "60-64" => 62m,
            "65-69" or "65-69+" or "65+" => 67m,
            "70 & over" or "70 and over" => 72m,
            "60 & over" or "60 and over" => 62m,
            _ => 0m
        };
    }

    private static string FormatWholeNumber(int? value)
    {
        return value.HasValue
            ? value.Value.ToString("#,##0", CultureInfo.CurrentCulture)
            : "";
    }

    private static string FormatWholeNumber(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("#,##0", CultureInfo.CurrentCulture)
            : "";
    }

    private static string FormatOneDecimal(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("0.0", CultureInfo.CurrentCulture)
            : "";
    }

    private static bool IsPremiumIllustrationTable(PowerPointExportTable table)
    {
        return table.Title.Equals(
            PowerPointRateTableBuilder.PremiumIllustrationTableTitle,
            StringComparison.OrdinalIgnoreCase);
    }

    private static string FormatRiskChange(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("0.00%", CultureInfo.InvariantCulture)
            : "";
    }

    private static bool ReplaceParagraphTokens(IEnumerable<Drawing.Paragraph> paragraphs, Dictionary<string, string> replacements)
    {
        var tokenFound = false;
        foreach (var paragraph in paragraphs)
        {
            var textRuns = paragraph.Descendants<Drawing.Text>().ToList();
            if (textRuns.Count == 0)
            {
                continue;
            }

            var original = string.Concat(textRuns.Select(text => text.Text));
            var replaced = original;
            foreach (var pair in replacements)
            {
                if (replaced.Contains(pair.Key, StringComparison.Ordinal))
                {
                    tokenFound = true;
                    replaced = replaced.Replace(pair.Key, pair.Value, StringComparison.Ordinal);
                }
            }

            if (replaced.Equals(original, StringComparison.Ordinal))
            {
                continue;
            }

            textRuns[0].Text = replaced;
            for (var i = 1; i < textRuns.Count; i++)
            {
                textRuns[i].Text = "";
            }
        }

        return tokenFound;
    }

    private static string DeterminePriorYear(AnalysisResult analysis)
    {
        var sourceYear = analysis.Workbook.PriorDrugTrends
            .Select(drug => TryExtractYear(drug.SourceReference))
            .FirstOrDefault(year => year.HasValue);
        if (sourceYear.HasValue)
        {
            return sourceYear.Value.ToString(CultureInfo.InvariantCulture);
        }

        var folderYear = TryExtractYear(analysis.RenewalCase.PriorYearFolder);
        if (folderYear.HasValue)
        {
            return folderYear.Value.ToString(CultureInfo.InvariantCulture);
        }

        return analysis.RenewalCase.RenewalDate.HasValue
            ? (analysis.RenewalCase.RenewalDate.Value.Year - 1).ToString(CultureInfo.InvariantCulture)
            : "";
    }

    private static int? TryExtractYear(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var years = Regex.Matches(value, @"(?<!\d)(20\d{2})(?!\d)")
            .Cast<Match>()
            .Select(match => int.Parse(match.Value, CultureInfo.InvariantCulture))
            .Where(year => year >= 2000 && year <= 2099)
            .ToList();

        return years.Count == 0 ? null : years.Min();
    }

    private static string ResolveInsurerName(AnalysisResult analysis)
    {
        if (!string.IsNullOrWhiteSpace(analysis.Pdf.Carrier))
        {
            return analysis.Pdf.Carrier.Trim();
        }

        if (analysis.SourceChecks.Any(check => check.Area.Contains("Victor", StringComparison.OrdinalIgnoreCase)))
        {
            return "Victor";
        }

        if (analysis.SourceChecks.Any(check => check.Area.Contains("Manulife", StringComparison.OrdinalIgnoreCase)) ||
            analysis.RenewalCase.PdfPath.Contains("Manulife", StringComparison.OrdinalIgnoreCase) ||
            analysis.RenewalCase.WorkbookPath.Contains("Manulife", StringComparison.OrdinalIgnoreCase))
        {
            return "Manulife";
        }

        return "Carrier";
    }

    private static Dictionary<string, string> BuildTokenMap(AnalysisResult analysis) =>
        BuildTokenMapWithReferenceRates(analysis, UnderwritingReferenceRateCatalog.Defaults);

    private static Dictionary<string, string> BuildTokenMapWithReferenceRates(
        AnalysisResult analysis,
        IReadOnlyCollection<UnderwritingReferenceRate> underwritingReferenceRates)
    {
        var clientName = analysis.RenewalCase.ClientName;
        var policyNumber = analysis.RenewalCase.PolicyNumber;
        var renewalDate = FormatDeckDate(analysis.RenewalCase.RenewalDate);
        var currentYear = analysis.RenewalCase.RenewalDate?.Year.ToString() ?? "";
        var priorYear = DeterminePriorYear(analysis);
        var insurer = ResolveInsurerName(analysis);

        var tokens = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["{{ClientName}}"] = clientName,
            ["{{PolicyNumber}}"] = policyNumber,
            ["{{RenewalDate}}"] = renewalDate,
            ["{{CurrentYear}}"] = currentYear,
            ["{{Current_Year}}"] = currentYear,
            ["{{Current Year}}"] = currentYear,
            ["{{PriorYear}}"] = priorYear,
            ["{{Prior_Year}}"] = priorYear,
            ["{{prior_year}}"] = priorYear,
            ["{{Prior Year}}"] = priorYear,
            ["{{Carrier}}"] = insurer,
            ["{{Insurer}}"] = insurer,
            ["\u00ABClient_Name\u00BB"] = clientName,
            ["\u00ABRenewal_Effective\u00BB"] = renewalDate,
            ["\u00ABPolicy\u00BB"] = policyNumber,
            ["\u00ABCarrier\u00BB"] = insurer,
            ["\u00ABInsurer\u00BB"] = insurer,
            ["\u00ABCurrent_Year\u00BB"] = currentYear,
            ["\u00ABCurrent Year\u00BB"] = currentYear,
            ["\u00ABPrior_Year\u00BB"] = priorYear,
            ["\u00AB Prior_Year \u00BB"] = priorYear,
            ["\u00ABprior_year\u00BB"] = priorYear,
            ["\u00ABPrior Year\u00BB"] = priorYear,
            ["<<Client_Name>>"] = clientName,
            ["<<Renewal_Effective>>"] = renewalDate,
            ["<<Policy>>"] = policyNumber,
            ["<<Carrier>>"] = insurer,
            ["<<Insurer>>"] = insurer,
            ["<<CurrentYear>>"] = currentYear,
            ["<<Current_Year>>"] = currentYear,
            ["<<Current Year>>"] = currentYear,
            ["<<PriorYear>>"] = priorYear,
            ["<<Prior_Year>>"] = priorYear,
            ["<< Prior_Year >>"] = priorYear,
            ["<<Prior Year>>"] = priorYear,
            ["<<prior_year>>"] = priorYear,
            ["<< prior_year >>"] = priorYear,
            ["<<implementation_date>>"] = renewalDate,
            ["<< implementation_date>>"] = renewalDate,
            ["<< implementation_date >>"] = renewalDate,
            ["{{CurrentMonthlyPremium}}"] = ParsingHelpers.Money(analysis.CurrentMonthlyPremium),
            ["{{RenewalMonthlyPremium}}"] = ParsingHelpers.Money(analysis.RenewalMonthlyPremium),
            ["{{MonthlyPremiumChange}}"] = ParsingHelpers.Money(analysis.MonthlyPremiumChange),
            ["{{MonthlyPremiumChangePercent}}"] = ParsingHelpers.Percent(analysis.MonthlyPremiumChangePercent),
            ["{{TopFinding}}"] = analysis.Findings.FirstOrDefault()?.Title ?? ""
        };

        if (analysis.YearsAsClient.HasValue)
        {
            var yearsAsClient = analysis.YearsAsClient.Value.ToString(CultureInfo.InvariantCulture);
            AddTokenVariants(tokens, "YearsAsClient", yearsAsClient);
            AddTokenVariants(tokens, "Years_As_Client", yearsAsClient);
            tokens["XXX years"] = $"{yearsAsClient} years";
            tokens["XXX Years"] = $"{yearsAsClient} years";
        }

        var lifeLines = FindBasicLifeRateLines(analysis);
        var adAndDLines = FindRateLines(analysis, "Accidental Death and Dismemberment", "AD&D", "ADD");
        var ltdLines = FindRateLines(analysis, "Long Term Disability", "Long-Term Disability", "LTD");

        var dependentLifeLines = FindRateLines(analysis, "Dependent Life");

        AddPremiumTokens(tokens, "Life", lifeLines);
        AddPremiumTokens(tokens, "DependentLife", dependentLifeLines);
        AddPremiumTokens(tokens, "ADandD", adAndDLines);
        AddPremiumTokens(tokens, "AD&D", adAndDLines);
        AddPremiumTokens(tokens, "STD", FindRateLines(analysis, "Short Term Disability", "STD"));
        AddPremiumTokens(tokens, "LTD", ltdLines);
        AddPremiumTokens(tokens, "CriticalIllness", FindRateLines(analysis, "Group Critical Illness", "Critical Illness"));
        AddPremiumTokens(tokens, "EFAP", FindRateLines(analysis, "Employee and Family Assistance"));
        AddPremiumTokens(tokens, "VH", FindRateLines(analysis, "Healthcare Online", "Virtual Healthcare", "Virtual Health Care"));
        AddPremiumTokens(tokens, "VirtualHealth", FindRateLines(analysis, "Healthcare Online", "Virtual Healthcare", "Virtual Health Care"));
        AddPremiumTokens(tokens, "MSO", FindRateLines(analysis, "Medical Second Opinion"));
        AddPremiumTokens(tokens, "HCSA", FindRateLines(analysis, "Health Spending Account", "HCSA", "HSA"));

        AddRateTokens(tokens, lifeLines, "Life", "BasicLife", "Basic_Life");
        AddRateTokens(tokens, dependentLifeLines, "DependentLife", "Dependent_Life");
        AddRateTokens(tokens, adAndDLines, "ADandD", "AD&D", "AD_D", "ADD", "AccidentalDeathandDismemberment", "AccidentalDeathAndDismemberment");
        AddRateTokens(tokens, ltdLines, "LTD", "LongtermDisability", "LongTermDisability", "Long_Term_Disability");
        AddDependentLifeRateAdjustmentToken(tokens, dependentLifeLines);

        var ehcLines = FindRateLines(analysis, "Extended Health Care", "Health", "EHC");
        if (ehcLines.Any(line => line.CurrentPremium.HasValue || line.RenewalPremium.HasValue))
        {
            AddPremiumTokens(tokens, "EHC", ehcLines);
        }
        else
        {
            AddPremiumTokens(tokens, "EHC", analysis.Pdf.CurrentEhcPremium, analysis.Pdf.RenewalEhcPremium, null);
        }

        var dentalLines = FindRateLines(analysis, "Dental");
        if (dentalLines.Any(line => line.CurrentPremium.HasValue || line.RenewalPremium.HasValue))
        {
            AddPremiumTokens(tokens, "Dental", dentalLines);
        }
        else
        {
            AddPremiumTokens(tokens, "Dental", analysis.Pdf.CurrentDentalPremium, analysis.Pdf.RenewalDentalPremium, null);
        }

        AddPremiumTokens(
            tokens,
            "Total",
            analysis.CurrentMonthlyPremium,
            analysis.RenewalMonthlyPremium,
            analysis.MonthlyPremiumChangePercent);

        AddPreviousRenewalTokens(tokens, analysis.MostRelevantPreviousRenewal);
        AddNonPooledEhcClaimsTokens(tokens, analysis);
        AddDentalClaimsTokens(tokens, analysis);
        AddDemographicAnalysisTokens(
            tokens,
            "Life",
            analysis.Pdf.LifeDemographics,
            underwritingReferenceRates,
            UnderwritingRateTypes.Mortality);
        AddDemographicAnalysisTokens(
            tokens,
            "LTD",
            analysis.Pdf.LtdDemographics,
            underwritingReferenceRates,
            UnderwritingRateTypes.Morbidity,
            addGenericCurrentAverageAgeToken: true);
        AddNemMaximumTokens(tokens, analysis);
        AddReportingPeriodTokens(tokens, analysis);
        AddParamedicalSummaryTokens(tokens, analysis);

        return tokens;
    }

    private static void AddReportingPeriodTokens(Dictionary<string, string> tokens, AnalysisResult analysis)
    {
        var reportingPeriod = BuildProgramClaimsSummaryChartData(analysis).ReportingPeriod;
        if (string.IsNullOrWhiteSpace(reportingPeriod))
        {
            return;
        }

        AddTokenVariants(tokens, "ReportingPeriod", reportingPeriod);
        AddTokenVariants(tokens, "Reporting_Period", reportingPeriod);
        AddTokenVariants(tokens, "CurrentExperiencePeriod", reportingPeriod);
        AddTokenVariants(tokens, "Current_Experience_Period", reportingPeriod);
        tokens["[reporting period]"] = reportingPeriod;
        tokens["[Reporting Period]"] = reportingPeriod;
        tokens["[Enter Reporting Period]"] = reportingPeriod;
        tokens["[enter range]"] = reportingPeriod;
    }

    private static void AddNemMaximumTokens(Dictionary<string, string> tokens, AnalysisResult analysis)
    {
        AddNemMaximumTokens(tokens, analysis, "Life Insurance", "Life", "BasicLife", "Basic_Life");
        AddNemMaximumTokens(tokens, analysis, "Long-Term Disability", "LTD", "LongTermDisability", "Long_Term_Disability");
    }

    private static void AddNemMaximumTokens(Dictionary<string, string> tokens, AnalysisResult analysis, string benefit, params string[] prefixes)
    {
        var row = analysis.NemMaximums.FirstOrDefault(item =>
            NemMaximumExtractor.NormalizeBenefitName(item.Benefit).Equals(NemMaximumExtractor.NormalizeBenefitName(benefit), StringComparison.OrdinalIgnoreCase));
        foreach (var prefix in prefixes)
        {
            AddTokenVariants(tokens, $"{prefix}_CurrentMaximum", row?.CurrentMaximum ?? "");
            AddTokenVariants(tokens, $"{prefix}_CurrentNEM", row?.CurrentNem ?? "");
            AddTokenVariants(tokens, $"{prefix}_AvailableNEM", row?.AvailableNem ?? "");
            AddTokenVariants(tokens, $"{prefix} Current Maximum", row?.CurrentMaximum ?? "");
            AddTokenVariants(tokens, $"{prefix} Current NEM", row?.CurrentNem ?? "");
            AddTokenVariants(tokens, $"{prefix} Available NEM", row?.AvailableNem ?? "");
        }
    }

    private static void AddPreviousRenewalTokens(Dictionary<string, string> tokens, PreviousRenewalSnapshot? snapshot)
    {
        var date = FormatDeckDate(snapshot?.RenewalDate);
        var carrier = snapshot?.Carrier ?? "";
        var source = snapshot?.SourceFileName ?? "";
        var currentPremium = ParsingHelpers.Money(snapshot?.CurrentMonthlyPremium);
        var renewalPremium = ParsingHelpers.Money(snapshot?.RenewalMonthlyPremium);
        var change = ParsingHelpers.Money(snapshot?.MonthlyPremiumChange);
        var changePercent = ParsingHelpers.Percent(snapshot?.MonthlyPremiumChangePercent);

        foreach (var prefix in new[] { "PreviousRenewal", "PriorRenewal", "LastRenewal" })
        {
            AddTokenVariants(tokens, $"{prefix}Date", date);
            AddTokenVariants(tokens, $"{prefix}Carrier", carrier);
            AddTokenVariants(tokens, $"{prefix}Source", source);
            AddTokenVariants(tokens, $"{prefix}CurrentMonthlyPremium", currentPremium);
            AddTokenVariants(tokens, $"{prefix}RenewalMonthlyPremium", renewalPremium);
            AddTokenVariants(tokens, $"{prefix}Premium", renewalPremium);
            AddTokenVariants(tokens, $"{prefix}MonthlyChange", change);
            AddTokenVariants(tokens, $"{prefix}Change", change);
            AddTokenVariants(tokens, $"{prefix}ChangePercent", changePercent);
            AddTokenVariants(tokens, $"{prefix}MonthlyChangePercent", changePercent);
        }
    }

    private static List<RateLine> FindRateLines(AnalysisResult analysis, params string[] benefits)
    {
        return SelectedRateLines(analysis)
            .Where(line => benefits.Any(benefit => line.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase)))
            .ToList();
    }

    private static List<RateLine> FindBasicLifeRateLines(AnalysisResult analysis)
    {
        return FindRateLines(
            analysis,
            "Life",
            "Basic Life",
            "Employee Life",
            "Employee Term Life",
            "Employee Term Life Insurance",
            "Life Insurance",
            "Employee Life Insurance");
    }

    private static List<RateLine> SelectedRateLines(AnalysisResult analysis)
    {
        var version = analysis.SelectedRenewalVersion;
        if (version?.RateLines.Count > 0)
        {
            return BenefitDisplayOrder.Sort(
                version.RateLines.Select(ToRateLine),
                line => line.Benefit,
                line => line.TierOrClass,
                line => line.AccountClass);
        }

        return BenefitDisplayOrder.Sort(
            analysis.Pdf.RateLines,
            line => line.Benefit,
            line => line.TierOrClass,
            line => line.AccountClass);
    }

    private static RateLine ToRateLine(RenewalVersionRateLine line)
    {
        var currentPremium = line.CurrentPremium;
        var renewalPremium = line.VersionPremium ?? line.ProposedPremium;
        var currentRate = ClientPresentationRateSelector.CurrentRate(line);
        var proposedRate = ClientPresentationRateSelector.ProposedRate(line);
        var versionRate = ClientPresentationRateSelector.VersionRate(line);

        return new RateLine
        {
            Benefit = line.Benefit,
            TierOrClass = line.TierOrClass,
            AccountClass = line.AccountClass,
            Volume = line.Volume,
            Lives = line.Lives,
            CurrentRate = currentRate,
            CurrentPremium = currentPremium,
            RenewalRate = versionRate,
            RenewalPremium = renewalPremium,
            CarrierProposedRate = proposedRate,
            CarrierProposedPremium = line.ProposedPremium,
            NegotiatedRate = versionRate,
            NegotiatedPremium = renewalPremium,
            PercentChange = line.ChangePercent ?? UnderwritingEngine.CalculatePercentChange(currentPremium, renewalPremium),
            SourceReference = line.SourceReference
        };
    }

    private static void AddPremiumTokens(Dictionary<string, string> tokens, string prefix, List<RateLine> lines)
    {
        if (lines.Count == 0)
        {
            AddPremiumTokens(tokens, prefix, null, null, null);
            return;
        }

        var currentPremium = lines.Sum(line => line.CurrentPremium ?? 0m);
        var renewalPremium = lines.Sum(line => line.RenewalPremium ?? 0m);
        var percentChange = lines.Count == 1
            ? lines[0].PercentChange ?? PercentChange(currentPremium, renewalPremium)
            : PercentChange(currentPremium, renewalPremium);

        AddPremiumTokens(tokens, prefix, currentPremium, renewalPremium, percentChange);
    }

    private static void AddPremiumTokens(
        Dictionary<string, string> tokens,
        string prefix,
        decimal? currentPremium,
        decimal? renewalPremium,
        decimal? percentChange)
    {
        tokens[$"{{{{{prefix}_CurrentPremium}}}}"] = ParsingHelpers.Money(currentPremium);
        tokens[$"{{{{{prefix}_RenewalPremium}}}}"] = ParsingHelpers.Money(renewalPremium);
        tokens[$"{{{{{prefix}_ChangePercent}}}}"] = percentChange.HasValue
            ? ParsingHelpers.Percent(percentChange)
            : PercentChange(currentPremium, renewalPremium).HasValue
                ? ParsingHelpers.Percent(PercentChange(currentPremium, renewalPremium))
                : "";
    }

    private static void AddRateTokens(Dictionary<string, string> tokens, List<RateLine> lines, params string[] prefixes)
    {
        var currentRate = SelectRepresentativeRate(lines, line => line.CurrentRate, line => line.CurrentPremium);
        var renewalRate = SelectRepresentativeRate(lines, line => line.RenewalRate, line => line.RenewalPremium);
        foreach (var prefix in prefixes.Where(prefix => !string.IsNullOrWhiteSpace(prefix)).Distinct(StringComparer.Ordinal))
        {
            AddTokenVariants(tokens, $"{prefix}_CurrentRate", FormatRate(currentRate));
            AddTokenVariants(tokens, $"{prefix}_RenewalRate", FormatRate(renewalRate));
        }
    }

    private static void AddDependentLifeRateAdjustmentToken(
        Dictionary<string, string> tokens,
        List<RateLine> dependentLifeLines)
    {
        var currentRate = SelectRepresentativeRate(dependentLifeLines, line => line.CurrentRate, line => line.CurrentPremium);
        var renewalRate = SelectRepresentativeRate(dependentLifeLines, line => line.RenewalRate, line => line.RenewalPremium);

        var adjustment = currentRate.HasValue && renewalRate.HasValue
            ? currentRate.Value == renewalRate.Value
                ? $"The Dependent Life rate will remain ${FormatRate(renewalRate)} per member."
                : $"The Dependent Life rate will change from ${FormatRate(currentRate)} to ${FormatRate(renewalRate)} per member."
            : "Dependent Life rate was not provided in the renewal source.";

        tokens["Add dependent life rate adjustment"] = adjustment;
    }

    private static void AddNonPooledEhcClaimsTokens(Dictionary<string, string> tokens, AnalysisResult analysis)
    {
        var mostCurrent = GetMostCurrentExperienceLine(
            analysis,
            "Non Pooled Extended Health Care - Subtotal",
            "Extended Health Care",
            "Extended Healthcare",
            "EHC");
        var paidClaims = FormatWholeMoney(mostCurrent?.PaidClaims);

        AddTokenVariants(tokens, "NonPooledExtendedHealthCare_SubtotalPaidClaims", paidClaims);
        AddTokenVariants(tokens, "NonPooledExtendedHealthCare_Subtotal_PaidClaims", paidClaims);
        AddTokenVariants(tokens, "NonPooledEHC_SubtotalPaidClaims", paidClaims);
        AddTokenVariants(tokens, "EHC_NonPooled_MostCurrent_PaidClaims", paidClaims);
        AddTokenVariants(tokens, "EHC_NonPooled_PaidClaims", paidClaims);
    }

    private static void AddDentalClaimsTokens(Dictionary<string, string> tokens, AnalysisResult analysis)
    {
        var mostCurrent = GetMostCurrentExperienceLine(analysis, "Dental");
        var paidClaims = FormatWholeNumber(mostCurrent?.PaidClaims);

        AddTokenVariants(tokens, "Mostcurrent_DentalPaidClaims", paidClaims);
        AddTokenVariants(tokens, "MostCurrent_DentalPaidClaims", paidClaims);
        AddTokenVariants(tokens, "MostCurrent_Dental_PaidClaims", paidClaims);
        AddTokenVariants(tokens, "Dental_MostCurrent_PaidClaims", paidClaims);
    }

    private static ExperienceLine? GetMostCurrentExperienceLine(AnalysisResult analysis, params string[] benefitAliases)
    {
        return analysis.Pdf.ExperienceLines
            .Where(row => MatchesActivityBenefit(row.Benefit, benefitAliases))
            .Where(IsActivityExperiencePeriod)
            .OrderBy(row => ParseExperiencePeriodStart(row.Period) ?? DateOnly.MinValue)
            .LastOrDefault();
    }

    private static string FormatWholeMoney(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("C0", CultureInfo.CurrentCulture)
            : "";
    }

    private static decimal? SelectRepresentativeRate(
        IEnumerable<RateLine> lines,
        Func<RateLine, decimal?> rateSelector,
        Func<RateLine, decimal?> premiumSelector)
    {
        return lines
            .Select(line => new
            {
                Rate = rateSelector(line),
                Premium = premiumSelector(line) ?? 0m,
                Lives = line.Lives ?? 0
            })
            .Where(item => item.Rate.HasValue)
            .GroupBy(item => decimal.Round(item.Rate!.Value, 4))
            .OrderByDescending(group => group.Count())
            .ThenByDescending(group => group.Sum(item => item.Premium))
            .ThenByDescending(group => group.Sum(item => item.Lives))
            .Select(group => (decimal?)group.Key)
            .FirstOrDefault();
    }

    private static void AddTokenVariants(Dictionary<string, string> tokens, string key, string value)
    {
        tokens[$"{{{{{key}}}}}"] = value;
        tokens[$"<<{key}>>"] = value;
        tokens[$"\u00AB{key}\u00BB"] = value;
    }

    private static string FormatDeckDate(DateOnly? date)
    {
        return date.HasValue
            ? date.Value.ToString("MMMM dd, yyyy", CultureInfo.InvariantCulture)
            : "";
    }

    private static string FormatRate(decimal? rate)
    {
        return rate.HasValue
            ? rate.Value.ToString("0.####", CultureInfo.InvariantCulture)
            : "";
    }

    private static decimal? PercentChange(decimal? current, decimal? renewal)
    {
        return UnderwritingEngine.CalculatePercentChange(current, renewal);
    }

    private static void WriteGeneratedDeck(AnalysisResult analysis, string outputPath)
    {
        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        var slides = BuildSlides(analysis);

        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Create);
        AddText(archive, "[Content_Types].xml", ContentTypes(slides.Count));
        AddText(archive, "_rels/.rels", PackageRelationships());
        AddText(archive, "docProps/core.xml", CoreProperties(analysis));
        AddText(archive, "docProps/app.xml", AppProperties(slides.Count));
        AddText(archive, "ppt/presentation.xml", Presentation(slides.Count));
        AddText(archive, "ppt/_rels/presentation.xml.rels", PresentationRelationships(slides.Count));
        AddText(archive, "ppt/theme/theme1.xml", Theme());
        AddText(archive, "ppt/slideMasters/slideMaster1.xml", SlideMaster());
        AddText(archive, "ppt/slideMasters/_rels/slideMaster1.xml.rels", SlideMasterRelationships());
        AddText(archive, "ppt/slideLayouts/slideLayout1.xml", SlideLayout());
        AddText(archive, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", SlideLayoutRelationships());

        for (var i = 0; i < slides.Count; i++)
        {
            AddText(archive, $"ppt/slides/slide{i + 1}.xml", SlideXml(slides[i], i + 1));
            AddText(archive, $"ppt/slides/_rels/slide{i + 1}.xml.rels", SlideRelationships());
        }
    }

    private static List<SlideSpec> BuildSlides(AnalysisResult analysis)
    {
        var topFindings = analysis.Findings.Take(5).ToList();
        var rateIssues = analysis.Pdf.RateLines
            .Where(line => Math.Abs(line.CurrentVariance ?? 0m) > 2m || Math.Abs(line.RenewalVariance ?? 0m) > 2m)
            .Take(5)
            .Select(line => $"{line.Benefit} {line.TierOrClass}: variance review needed")
            .ToList();

        var ehcDrivers = analysis.Workbook.ClaimTrends
            .Where(trend => trend.Benefit.Contains("Health", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(trend => trend.PercentOfTotal ?? 0m)
            .Take(4)
            .Select(trend => $"{trend.Category}: {ParsingHelpers.Percent(trend.PercentOfTotal)}")
            .ToList();

        var dentalDrivers = analysis.Workbook.ClaimTrends
            .Where(trend => trend.Benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(trend => trend.PercentOfTotal ?? 0m)
            .Take(4)
            .Select(trend => $"{trend.Category}: {ParsingHelpers.Percent(trend.PercentOfTotal)}")
            .ToList();

        var drugDrivers = analysis.Workbook.IndicationTrends
            .OrderByDescending(indication => indication.AllowableTotal ?? indication.SubmittedTotal ?? 0m)
            .Take(4)
            .Select(indication => $"{indication.Indication}: {ParsingHelpers.Money(indication.AllowableTotal ?? indication.SubmittedTotal)}")
            .ToList();

        return
        [
            new SlideSpec(
                "Renewal Analysis",
                $"{analysis.RenewalCase.ClientName} | Policy {analysis.RenewalCase.PolicyNumber}",
                [
                    $"Renewal date: {(analysis.RenewalCase.RenewalDate.HasValue ? FormatDeckDate(analysis.RenewalCase.RenewalDate) : "Not extracted")}",
                    $"Current monthly premium: {ParsingHelpers.Money(analysis.CurrentMonthlyPremium)}",
                    $"Renewal monthly premium: {ParsingHelpers.Money(analysis.RenewalMonthlyPremium)}",
                    $"Monthly change: {ParsingHelpers.Money(analysis.MonthlyPremiumChange)} ({ParsingHelpers.Percent(analysis.MonthlyPremiumChangePercent)})"
                ]),
            new SlideSpec(
                "Executive Summary",
                "Main renewal story",
                topFindings.Count == 0
                    ? ["No material findings were detected from the extracted files."]
                    : topFindings.Select(finding => $"{finding.Title}: {finding.ClientExplanation}").ToArray()),
            new SlideSpec(
                "Rate Reconciliation",
                "Premium and rate checks",
                rateIssues.Count == 0
                    ? ["No material premium recalculation variances were detected using Phase 1 rules."]
                    : rateIssues.ToArray()),
            new SlideSpec(
                "EHC And Dental Drivers",
                "Claim categories and utilization mix",
                ehcDrivers.Concat(dentalDrivers).DefaultIfEmpty("No EHC or dental driver detail was extracted.").ToArray()),
            new SlideSpec(
                "Drug And Indication Review",
                "High-cost drug patterns",
                drugDrivers.DefaultIfEmpty("No drug indication detail was extracted.").ToArray()),
            new SlideSpec(
                "Consultant Talking Points",
                "Recommended client narrative",
                topFindings
                    .Where(finding => finding.Severity != FindingSeverity.Info)
                    .Take(5)
                    .Select(finding => finding.RecommendedAction)
                    .DefaultIfEmpty("Confirm source documents and use the workbook as the evidence trail before presenting.")
                    .ToArray())
        ];
    }

    private static void AddText(ZipArchive archive, string path, string text)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream);
        writer.Write(text);
    }

    private static string SlideXml(SlideSpec slide, int slideNumber)
    {
        var shapes = new List<string>
        {
            Rect(2, 0, 0, 12192000, 520000, "0E7C7B"),
            TextBox(3, 520000, 700000, 11100000, 700000, slide.Title, 3600, "17202A", true),
            TextBox(4, 540000, 1360000, 11000000, 420000, slide.Subtitle, 1700, "5F6B7A", false),
            Rect(5, 520000, 1900000, 11150000, 15000, "C69C3F"),
            TextBox(6, 760000, 2200000, 10600000, 3600000, string.Join("\n", slide.Bullets.Select(bullet => "- " + bullet)), 1850, "17202A", false),
            TextBox(7, 10400000, 6200000, 1200000, 240000, $"Slide {slideNumber}", 1000, "5F6B7A", false)
        };

        return $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {string.Join(Environment.NewLine, shapes)}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
""";
    }

    private static string TextBox(int id, long x, long y, long cx, long cy, string text, int fontSize, string color, bool bold)
    {
        var paragraphs = text.Split('\n', StringSplitOptions.None)
            .Select(line => Paragraph(line, fontSize, color, bold))
            .ToArray();

        return $"""
<p:sp>
  <p:nvSpPr><p:cNvPr id="{id}" name="Text {id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>{string.Join("", paragraphs)}</p:txBody>
</p:sp>
""";
    }

    private static string Paragraph(string text, int fontSize, string color, bool bold)
    {
        var boldAttr = bold ? " b=\"1\"" : "";
        return $"""
<a:p><a:pPr/><a:r><a:rPr lang="en-US" sz="{fontSize}"{boldAttr}><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:rPr><a:t>{Escape(text)}</a:t></a:r><a:endParaRPr lang="en-US" sz="{fontSize}"/></a:p>
""";
    }

    private static string Rect(int id, long x, long y, long cx, long cy, string fill)
    {
        return $"""
<p:sp>
  <p:nvSpPr><p:cNvPr id="{id}" name="Band {id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>
</p:sp>
""";
    }

    private static string ContentTypes(int slideCount)
    {
        var slideOverrides = string.Join(Environment.NewLine, Enumerable.Range(1, slideCount)
            .Select(i => $"""  <Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"""));

        return $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
{slideOverrides}
</Types>
""";
    }

    private static string PackageRelationships()
    {
        return """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
""";
    }

    private static string Presentation(int slideCount)
    {
        var slideIds = string.Join(Environment.NewLine, Enumerable.Range(1, slideCount)
            .Select(i => $"""    <p:sldId id="{255 + i}" r:id="rId{i + 1}"/>"""));

        return $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>
{slideIds}
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
""";
    }

    private static string PresentationRelationships(int slideCount)
    {
        var relationships = new List<string>
        {
            """  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>"""
        };

        relationships.AddRange(Enumerable.Range(1, slideCount)
            .Select(i => $"""  <Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>"""));

        return $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{string.Join(Environment.NewLine, relationships)}
</Relationships>
""";
    }

    private static string SlideRelationships()
    {
        return """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>
""";
    }

    private static string SlideMaster()
    {
        return """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>
""";
    }

    private static string SlideMasterRelationships()
    {
        return """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
""";
    }

    private static string SlideLayout()
    {
        return """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
""";
    }

    private static string SlideLayoutRelationships()
    {
        return """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
""";
    }

    private static string Theme()
    {
        return """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Westland Minimal">
  <a:themeElements>
    <a:clrScheme name="Westland">
      <a:dk1><a:srgbClr val="17202A"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="0E7C7B"/></a:dk2><a:lt2><a:srgbClr val="F4F7F9"/></a:lt2>
      <a:accent1><a:srgbClr val="0E7C7B"/></a:accent1><a:accent2><a:srgbClr val="C69C3F"/></a:accent2>
      <a:accent3><a:srgbClr val="5F6B7A"/></a:accent3><a:accent4><a:srgbClr val="2E86AB"/></a:accent4>
      <a:accent5><a:srgbClr val="7D4E57"/></a:accent5><a:accent6><a:srgbClr val="2F5D50"/></a:accent6>
      <a:hlink><a:srgbClr val="2E86AB"/></a:hlink><a:folHlink><a:srgbClr val="7D4E57"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Segoe UI">
      <a:majorFont><a:latin typeface="Segoe UI"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Segoe UI"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Clean">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>
""";
    }

    private static string CoreProperties(AnalysisResult analysis)
    {
        return $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Renewal Review Draft</dc:title>
  <dc:subject>{Escape(analysis.RenewalCase.ClientName)}</dc:subject>
  <dc:creator>Renewal Review Builder</dc:creator>
  <cp:lastModifiedBy>Renewal Review Builder</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ssZ}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{DateTime.UtcNow:yyyy-MM-ddTHH:mm:ssZ}</dcterms:modified>
</cp:coreProperties>
""";
    }

    private static string AppProperties(int slideCount)
    {
        return $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Renewal Review Builder</Application>
  <PresentationFormat>Widescreen</PresentationFormat>
  <Slides>{slideCount}</Slides>
</Properties>
""";
    }

    private static string Escape(string? text)
    {
        return SecurityElement.Escape(text ?? "") ?? "";
    }

    private sealed record SlideSpec(string Title, string Subtitle, string[] Bullets);

    private sealed record DemographicChartRow(string AgeBand, decimal PriorVolume, decimal CurrentVolume);

    private sealed record ActivityChartRow(string Label, decimal BilledPremium, decimal PaidClaims, decimal IncurredClaims, decimal PaidLossRatio);

    private sealed record ProgramClaimsSummaryChartData(string ReportingPeriod, IReadOnlyList<ProgramClaimsSummaryRow> Rows);

    private sealed record ProgramClaimsSummaryRow(string Category, decimal PaidPremiums, decimal PaidClaims, decimal LossRatio);

    private sealed record SpendingTrendChartRow(string Category, string PriorLabel, decimal PriorAmount, string CurrentLabel, decimal CurrentAmount, IReadOnlyList<ServiceCategoryPeriodValue>? History = null);

    private sealed record TableFrameGeometry(string X, string Y, string Cx, string Cy);
}
