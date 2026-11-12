using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public enum ClaimsRangePowerPointUpdateStatus
{
    Updated,
    Removed,
    SlideNotFound,
    ContractInvalid
}

public sealed record ClaimsRangePowerPointUpdateResult(
    ClaimsRangePowerPointUpdateStatus Status,
    string Message)
{
    public bool Updated => Status == ClaimsRangePowerPointUpdateStatus.Updated;
    public bool Removed => Status is ClaimsRangePowerPointUpdateStatus.Removed or ClaimsRangePowerPointUpdateStatus.ContractInvalid;
}

/// <summary>
/// Populates the native Claims Range charts in an already-copied renewal PowerPoint deck.
/// The source template is never modified. If client-ready data or the required slide
/// contract is unavailable, the sample slide is removed so template data cannot leak.
/// </summary>
public sealed class ClaimsRangePowerPointUpdater
{
    internal const string ContractMarkerName = "claims-range-contract-v1";
    internal const string ExtendedHealthChartName = "claims-range-ehc-chart";
    internal const string DentalChartName = "claims-range-dental-chart";

    private const string ClaimsRangeTitle = "Claims Range";
    private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string PackageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";

    private static readonly string[] RequiredLegacyShapeNames =
    [
        "subtitle",
        "prior-legend-swatch",
        "current-legend-swatch",
        "prior-legend",
        "current-legend",
        "ehc-heading",
        "dental-heading",
        "vertical-divider",
        "takeaway-band",
        "takeaway-accent",
        "takeaway-label",
        "takeaway-copy",
        "footnote"
    ];

    public ClaimsRangePowerPointUpdateResult Update(
        string presentationPath,
        ClaimsRangeComparison? comparison)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(presentationPath);
        if (!File.Exists(presentationPath))
        {
            throw new FileNotFoundException("The renewal PowerPoint output was not found.", presentationPath);
        }

        var candidateSlidePaths = FindCandidateSlidePaths(presentationPath);
        if (candidateSlidePaths.Count == 0)
        {
            return new ClaimsRangePowerPointUpdateResult(
                ClaimsRangePowerPointUpdateStatus.SlideNotFound,
                "The Claims Range slide was not present in the output deck.");
        }

        if (comparison is null || !comparison.CanUseForClientOutput)
        {
            RemoveSlides(presentationPath, candidateSlidePaths);
            var reason = comparison is null
                ? "not available"
                : "unconfirmed or has unresolved review issues";
            return new ClaimsRangePowerPointUpdateResult(
                ClaimsRangePowerPointUpdateStatus.Removed,
                $"The Claims Range slide was removed because the comparison is {reason}.");
        }

        if (!TryValidateComparison(comparison, out var validationMessage))
        {
            RemoveSlides(presentationPath, candidateSlidePaths);
            return new ClaimsRangePowerPointUpdateResult(
                ClaimsRangePowerPointUpdateStatus.Removed,
                $"The Claims Range slide was removed because {validationMessage}");
        }

        if (candidateSlidePaths.Count != 1)
        {
            RemoveSlides(presentationPath, candidateSlidePaths);
            return new ClaimsRangePowerPointUpdateResult(
                ClaimsRangePowerPointUpdateStatus.ContractInvalid,
                "The Claims Range slide contract was ambiguous, so every candidate slide was removed.");
        }

        var contractInvalid = false;
        var contractMessage = "";
        try
        {
            using var archive = ZipFile.Open(presentationPath, ZipArchiveMode.Update);
            var slidePath = candidateSlidePaths[0];
            if (!TryLoadContract(archive, slidePath, out var contract, out contractMessage))
            {
                contractInvalid = true;
            }
            else
            {
                PopulateSlide(archive, contract!, comparison);
            }
        }
        catch (Exception ex)
        {
            contractInvalid = true;
            contractMessage = ex.Message;
        }

        if (contractInvalid)
        {
            RemoveSlides(presentationPath, candidateSlidePaths);
            return new ClaimsRangePowerPointUpdateResult(
                ClaimsRangePowerPointUpdateStatus.ContractInvalid,
                $"The Claims Range slide contract was invalid and the sample slide was removed. {contractMessage}".Trim());
        }

        return new ClaimsRangePowerPointUpdateResult(
            ClaimsRangePowerPointUpdateStatus.Updated,
            "The Claims Range slide and both native charts were updated.");
    }

    private static bool TryValidateComparison(
        ClaimsRangeComparison comparison,
        out string message)
    {
        if (comparison.PriorReport?.PeriodStart is null ||
            comparison.PriorReport.PeriodEnd is null ||
            comparison.CurrentReport?.PeriodStart is null ||
            comparison.CurrentReport.PeriodEnd is null)
        {
            message = "the prior/current reporting periods are missing.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(comparison.EffectiveTakeaway))
        {
            message = "the approved client takeaway is missing.";
            return false;
        }

        var isGroupHealth =
            comparison.PriorReport.Carrier.Equals("GroupHEALTH", StringComparison.OrdinalIgnoreCase) &&
            comparison.CurrentReport.Carrier.Equals("GroupHEALTH", StringComparison.OrdinalIgnoreCase);
        if (isGroupHealth &&
            (!HasCompleteBands(comparison.Rows, ClaimsRangeBenefitKeys.Health) ||
             !HasCompleteBands(comparison.Rows, ClaimsRangeBenefitKeys.Dental)))
        {
            message = "both benefits must contain each of the nine claims-range bands exactly once.";
            return false;
        }

        if (!isGroupHealth && !HasDynamicBands(comparison.Rows))
        {
            message = "the extracted claims-banding rows are incomplete or duplicated.";
            return false;
        }

        if (comparison.Rows.Any(row => row.PriorCount < 0 || row.CurrentCount < 0))
        {
            message = "certificate counts cannot be negative.";
            return false;
        }

        message = "";
        return true;
    }

    private static bool HasCompleteBands(
        IReadOnlyCollection<ClaimsRangeComparisonRow> rows,
        string benefit)
    {
        var benefitRows = rows
            .Where(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var canonicalCodes = ClaimsRangeBandCatalog.All
            .Select(band => band.RangeCode)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        return benefitRows.Count == ClaimsRangeBandCatalog.All.Count &&
               benefitRows.Select(row => row.RangeCode).Distinct(StringComparer.OrdinalIgnoreCase).Count() == canonicalCodes.Count &&
               benefitRows.All(row => canonicalCodes.Contains(row.RangeCode));
    }

    private static bool HasDynamicBands(IReadOnlyCollection<ClaimsRangeComparisonRow> rows)
    {
        return rows.Count > 0 && rows
            .GroupBy(row => row.Benefit, StringComparer.OrdinalIgnoreCase)
            .All(group => group.Any() &&
                          group.Select(row => row.RangeCode)
                              .Distinct(StringComparer.OrdinalIgnoreCase).Count() == group.Count());
    }

    private static List<string> FindCandidateSlidePaths(string presentationPath)
    {
        using var archive = ZipFile.OpenRead(presentationPath);
        var candidates = new List<string>();
        foreach (var entry in archive.Entries.Where(IsSlideEntry))
        {
            var document = XDocument.Parse(ReadZipEntryText(entry));
            if (HasNamedObject(document, ContractMarkerName) ||
                HasClaimsRangeTitle(document) ||
                HasLegacyClaimsRangeSignature(document))
            {
                candidates.Add(entry.FullName);
            }
        }

        return candidates;
    }

    private static bool TryLoadContract(
        ZipArchive archive,
        string slidePath,
        out ClaimsRangeSlideContract? contract,
        out string message)
    {
        contract = null;
        var slideEntry = archive.GetEntry(slidePath);
        if (slideEntry is null)
        {
            message = $"Slide part not found: {slidePath}";
            return false;
        }

        var slide = XDocument.Parse(ReadZipEntryText(slideEntry));
        var p = XNamespace.Get(PresentationNamespace);
        var c = XNamespace.Get(ChartNamespace);
        var a = XNamespace.Get(DrawingNamespace);
        var r = XNamespace.Get(RelationshipNamespace);

        foreach (var requiredName in RequiredLegacyShapeNames)
        {
            if (FindObjectsByName(slide, requiredName).Count != 1)
            {
                message = $"Required shape '{requiredName}' was missing or duplicated.";
                return false;
            }
        }

        var titleShapes = slide.Descendants(p + "sp")
            .Where(shape => ConcatenateText(shape, a).Trim().Equals(ClaimsRangeTitle, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var markerObjects = FindObjectsByName(slide, ContractMarkerName);
        var chartFrames = slide.Descendants(p + "graphicFrame")
            .Where(frame => frame.Descendants(c + "chart").Any())
            .ToList();
        if (chartFrames.Count != 2)
        {
            message = "The slide must contain exactly two native chart frames.";
            return false;
        }

        XElement? marker;
        XElement extendedHealthFrame;
        XElement dentalFrame;
        if (markerObjects.Count == 1)
        {
            marker = markerObjects[0];
            var ehcFrames = FindObjectsByName(slide, ExtendedHealthChartName);
            var dentalFrames = FindObjectsByName(slide, DentalChartName);
            if (ehcFrames.Count != 1 || dentalFrames.Count != 1 ||
                ehcFrames[0].Name != p + "graphicFrame" || dentalFrames[0].Name != p + "graphicFrame")
            {
                message = "The versioned slide was missing a uniquely named EHC or Dental chart.";
                return false;
            }

            extendedHealthFrame = ehcFrames[0];
            dentalFrame = dentalFrames[0];
        }
        else
        {
            if (markerObjects.Count != 0 || titleShapes.Count != 1)
            {
                message = "The legacy slide did not contain one exact Claims Range title.";
                return false;
            }

            marker = titleShapes[0];
            var orderedFrames = chartFrames
                .OrderBy(frame => ParseLong(frame.Descendants(p + "xfrm").Elements(a + "off").FirstOrDefault()?.Attribute("x")?.Value))
                .ToList();
            extendedHealthFrame = orderedFrames[0];
            dentalFrame = orderedFrames[1];

            SetObjectName(marker, ContractMarkerName);
            SetObjectName(extendedHealthFrame, ExtendedHealthChartName);
            SetObjectName(dentalFrame, DentalChartName);
        }

        var relationshipsPath = $"ppt/slides/_rels/{Path.GetFileName(slidePath)}.rels";
        var relationshipsEntry = archive.GetEntry(relationshipsPath);
        if (relationshipsEntry is null)
        {
            message = "The Claims Range slide relationships were missing.";
            return false;
        }

        var relationships = XDocument.Parse(ReadZipEntryText(relationshipsEntry));
        var ehcChartPath = ResolveChartPath(slidePath, extendedHealthFrame, relationships, c, r);
        var dentalChartPath = ResolveChartPath(slidePath, dentalFrame, relationships, c, r);
        if (string.IsNullOrWhiteSpace(ehcChartPath) || string.IsNullOrWhiteSpace(dentalChartPath) ||
            ehcChartPath.Equals(dentalChartPath, StringComparison.OrdinalIgnoreCase))
        {
            message = "The named chart frames did not resolve to two distinct chart parts.";
            return false;
        }

        if (!HasSupportedChartStructure(archive, ehcChartPath) ||
            !HasSupportedChartStructure(archive, dentalChartPath))
        {
            message = "Each Claims Range chart must be a two-series bar chart with a value axis.";
            return false;
        }

        contract = new ClaimsRangeSlideContract(
            slidePath,
            slide,
            ehcChartPath,
            dentalChartPath);
        message = "";
        return true;
    }

    private static void PopulateSlide(
        ZipArchive archive,
        ClaimsRangeSlideContract contract,
        ClaimsRangeComparison comparison)
    {
        var priorLabel = BuildLegendLabel("Prior", comparison.PriorReport!);
        var currentLabel = BuildLegendLabel("Current", comparison.CurrentReport!);
        ReplaceNamedShapeText(contract.Slide, "subtitle", "Number of certificates by annual paid-claims range");
        ReplaceNamedShapeText(contract.Slide, "prior-legend", priorLabel);
        ReplaceNamedShapeText(contract.Slide, "current-legend", currentLabel);
        ReplaceNamedShapeText(contract.Slide, "ehc-heading", "Extended Health Care");
        ReplaceNamedShapeText(contract.Slide, "dental-heading", "Dental Care");
        ReplaceNamedShapeText(contract.Slide, "takeaway-label", "CLIENT TAKEAWAY");
        ReplaceNamedShapeText(contract.Slide, "takeaway-copy", comparison.EffectiveTakeaway.Trim());
        ReplaceNamedShapeText(contract.Slide, "footnote", ClaimsRangeCommentaryBuilder.Footnote);

        var ehcRows = comparison.Rows
            .Where(row => row.Benefit.Equals(ClaimsRangeBenefitKeys.Health, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(row => row.SortOrder)
            .ToList();
        var dentalRows = comparison.Rows
            .Where(row => row.Benefit.Equals(ClaimsRangeBenefitKeys.Dental, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(row => row.SortOrder)
            .ToList();
        var ehcCategories = ehcRows.Select(row => FormatRangeLabel(row.RangeLabel)).ToList();
        var dentalCategories = dentalRows.Select(row => FormatRangeLabel(row.RangeLabel)).ToList();
        var ehcPrior = ehcRows.Select(row => row.PriorCount).ToList();
        var ehcCurrent = ehcRows.Select(row => row.CurrentCount).ToList();
        var dentalPrior = dentalRows.Select(row => row.PriorCount).ToList();
        var dentalCurrent = dentalRows.Select(row => row.CurrentCount).ToList();
        EnsureChartData(ehcCategories, ehcPrior, ehcCurrent);
        EnsureChartData(dentalCategories, dentalPrior, dentalCurrent);
        var scale = new ClaimsRangeAxisScaleCalculator().Calculate(
            ehcPrior.Concat(ehcCurrent).Concat(dentalPrior).Concat(dentalCurrent));

        ReplaceZipEntry(archive, contract.SlidePath, contract.Slide);
        UpdateChart(
            archive,
            contract.ExtendedHealthChartPath,
            priorLabel,
            currentLabel,
            ehcCategories,
            ehcPrior,
            ehcCurrent,
            scale);
        UpdateChart(
            archive,
            contract.DentalChartPath,
            priorLabel,
            currentLabel,
            dentalCategories,
            dentalPrior,
            dentalCurrent,
            scale);
        UpdateSourceNotes(archive, contract.SlidePath, comparison);
    }

    private static void EnsureChartData(
        List<string> categories,
        List<int> prior,
        List<int> current)
    {
        if (categories.Count > 0)
        {
            return;
        }

        categories.Add("Not available");
        prior.Add(0);
        current.Add(0);
    }

    private static string BuildLegendLabel(string prefix, ClaimsRangeReport report)
    {
        return $"{prefix} ({report.PeriodStart!.Value.Year}-{report.PeriodEnd!.Value.Year % 100:00})";
    }

    private static void UpdateSourceNotes(
        ZipArchive archive,
        string slidePath,
        ClaimsRangeComparison comparison)
    {
        var relationshipsPath = $"ppt/slides/_rels/{Path.GetFileName(slidePath)}.rels";
        var relationshipsEntry = archive.GetEntry(relationshipsPath);
        if (relationshipsEntry is null)
        {
            return;
        }

        var relationships = XDocument.Parse(ReadZipEntryText(relationshipsEntry));
        var notesRelationship = relationships.Root?.Elements().SingleOrDefault(element =>
            ((string?)element.Attribute("Type"))?.EndsWith("/notesSlide", StringComparison.OrdinalIgnoreCase) == true);
        var notesTarget = (string?)notesRelationship?.Attribute("Target");
        if (string.IsNullOrWhiteSpace(notesTarget))
        {
            return;
        }

        var notesPath = ResolvePackagePath(slidePath, notesTarget);
        var notesEntry = archive.GetEntry(notesPath);
        if (notesEntry is null)
        {
            return;
        }

        var notes = XDocument.Parse(ReadZipEntryText(notesEntry));
        var a = XNamespace.Get(DrawingNamespace);
        var paragraphs = notes.Descendants(a + "p").ToList();
        var startIndex = paragraphs.FindIndex(paragraph =>
            ConcatenateText(paragraph, a).Contains("[Sources]", StringComparison.OrdinalIgnoreCase));
        var endIndex = paragraphs.FindIndex(startIndex < 0 ? 0 : startIndex, paragraph =>
            ConcatenateText(paragraph, a).Contains("[/Sources]", StringComparison.OrdinalIgnoreCase));
        if (startIndex < 0 || endIndex < startIndex)
        {
            return;
        }

        var startParagraph = paragraphs[startIndex];
        var endParagraph = paragraphs[endIndex];
        if (startParagraph.Parent is null || startParagraph.Parent != endParagraph.Parent)
        {
            return;
        }

        var startText = ConcatenateText(startParagraph, a);
        var endText = ConcatenateText(endParagraph, a);
        var prefix = startText[..startText.IndexOf("[Sources]", StringComparison.OrdinalIgnoreCase)].Trim();
        var endMarkerIndex = endText.IndexOf("[/Sources]", StringComparison.OrdinalIgnoreCase);
        var suffix = endText[(endMarkerIndex + "[/Sources]".Length)..].Trim();
        var replacementLines = new List<string>();
        if (prefix.Length > 0)
        {
            replacementLines.Add(prefix);
        }

        replacementLines.Add("[Sources]");
        replacementLines.Add(BuildSourceNoteLine(comparison.PriorReport!));
        replacementLines.Add(BuildSourceNoteLine(comparison.CurrentReport!));
        replacementLines.Add("[/Sources]");
        if (suffix.Length > 0)
        {
            replacementLines.Add(suffix);
        }

        foreach (var line in replacementLines)
        {
            var paragraph = new XElement(startParagraph);
            ReplaceParagraphText(paragraph, a, line);
            startParagraph.AddBeforeSelf(paragraph);
        }

        foreach (var paragraph in paragraphs.Skip(startIndex).Take(endIndex - startIndex + 1))
        {
            paragraph.Remove();
        }

        ReplaceZipEntry(archive, notesPath, notes);
    }

    private static string BuildSourceNoteLine(ClaimsRangeReport report)
    {
        var fileName = !string.IsNullOrWhiteSpace(report.SourceFileName)
            ? report.SourceFileName.Trim()
            : Path.GetFileName(report.SourcePath);
        if (string.IsNullOrWhiteSpace(fileName))
        {
            fileName = "Claims Range report";
        }

        var period = report.PeriodStart.HasValue && report.PeriodEnd.HasValue
            ? $", {report.PeriodStart.Value:MMMM d, yyyy}-{report.PeriodEnd.Value:MMMM d, yyyy}"
            : "";
        return $"- {fileName} (Extended Health Care and Dental Care{period})";
    }

    private static void ReplaceParagraphText(XElement paragraph, XNamespace a, string value)
    {
        var textNodes = paragraph.Descendants(a + "t").ToList();
        if (textNodes.Count == 0)
        {
            throw new InvalidOperationException("A Claims Range notes paragraph did not contain editable text.");
        }

        textNodes[0].Value = value;
        foreach (var extraText in textNodes.Skip(1))
        {
            extraText.Value = "";
        }
    }

    private static string FormatRangeLabel(string label)
    {
        return (label ?? "").Replace("-", "–", StringComparison.Ordinal);
    }

    private static void UpdateChart(
        ZipArchive archive,
        string chartPath,
        string priorLabel,
        string currentLabel,
        IReadOnlyList<string> categories,
        IReadOnlyList<int> priorValues,
        IReadOnlyList<int> currentValues,
        ClaimsRangeAxisScale scale)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get(ChartNamespace);
        var a = XNamespace.Get(DrawingNamespace);
        var barChart = document.Descendants(c + "barChart").Single();
        var series = barChart.Elements(c + "ser").ToList();
        if (series.Count != 2)
        {
            throw new InvalidOperationException("A Claims Range chart did not contain exactly two series.");
        }

        UpdateSeries(series[0], c, a, 0, "B", priorLabel, categories, priorValues, "66C5AF");
        UpdateSeries(series[1], c, a, 1, "C", currentLabel, categories, currentValues, "004751");

        var valueAxis = document.Descendants(c + "valAx").Single();
        var scaling = valueAxis.Element(c + "scaling") ?? throw new InvalidOperationException("The value-axis scaling element was missing.");
        SetRequiredValue(scaling.Element(c + "min"), scale.Minimum);
        SetRequiredValue(scaling.Element(c + "max"), scale.Maximum);
        SetRequiredValue(valueAxis.Element(c + "majorUnit"), scale.MajorUnit);
        var numberFormat = valueAxis.Element(c + "numFmt");
        if (numberFormat is not null)
        {
            numberFormat.SetAttributeValue("formatCode", "0");
            numberFormat.SetAttributeValue("sourceLinked", "0");
        }

        var workbookPath = EnsureEmbeddedWorkbookRelationship(archive, chartPath, document, c);
        ReplaceZipEntry(archive, chartPath, document);
        UpdateEmbeddedWorkbook(archive, workbookPath, priorLabel, currentLabel, categories, priorValues, currentValues);
    }

    private static void UpdateSeries(
        XElement series,
        XNamespace c,
        XNamespace a,
        int index,
        string column,
        string seriesLabel,
        IReadOnlyList<string> categories,
        IReadOnlyList<int> values,
        string fillColor)
    {
        SetRequiredValue(series.Element(c + "idx"), index);
        SetRequiredValue(series.Element(c + "order"), index);
        SetSeriesText(series, c, column, seriesLabel);
        SetSeriesCategories(series, c, categories);
        SetSeriesValues(series, c, column, values);
        SetSeriesFill(series, c, a, fillColor);
        RebuildDataLabels(series, c, a, values);
    }

    private static void SetSeriesText(XElement series, XNamespace c, string column, string value)
    {
        var text = series.Element(c + "tx") ?? throw new InvalidOperationException("A chart series title was missing.");
        text.RemoveNodes();
        var reference = new XElement(c + "strRef");
        SetFormula(reference, c, $"Sheet1!${column}$1");
        ReplaceStringPoints(EnsureChild(reference, c + "strCache"), c, [value]);
        text.Add(reference);
    }

    private static void SetSeriesCategories(XElement series, XNamespace c, IReadOnlyList<string> categories)
    {
        var category = series.Element(c + "cat") ?? throw new InvalidOperationException("Chart categories were missing.");
        category.RemoveNodes();
        var reference = new XElement(c + "strRef");
        SetFormula(reference, c, $"Sheet1!$A$2:$A${categories.Count + 1}");
        ReplaceStringPoints(EnsureChild(reference, c + "strCache"), c, categories);
        category.Add(reference);
    }

    private static void SetSeriesValues(
        XElement series,
        XNamespace c,
        string column,
        IReadOnlyList<int> values)
    {
        var value = series.Element(c + "val") ?? throw new InvalidOperationException("Chart series values were missing.");
        value.RemoveNodes();
        var reference = new XElement(c + "numRef");
        SetFormula(reference, c, $"Sheet1!${column}$2:${column}${values.Count + 1}");
        ReplaceNumberPoints(EnsureChild(reference, c + "numCache"), c, values);
        value.Add(reference);
    }

    private static void SetSeriesFill(XElement series, XNamespace c, XNamespace a, string fillColor)
    {
        var properties = series.Element(c + "spPr") ?? throw new InvalidOperationException("Chart series formatting was missing.");
        properties.Elements(a + "solidFill").Remove();
        properties.AddFirst(new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", fillColor))));
    }

    private static void RebuildDataLabels(
        XElement series,
        XNamespace c,
        XNamespace a,
        IReadOnlyList<int> values)
    {
        series.Elements(c + "dLbls").Remove();
        var labels = new XElement(c + "dLbls");
        for (var index = 0; index < values.Count; index++)
        {
            var label = new XElement(c + "dLbl", new XElement(c + "idx", new XAttribute("val", index)));
            if (values[index] == 0)
            {
                label.Add(new XElement(c + "delete", new XAttribute("val", "1")));
            }
            else
            {
                label.Add(
                    BuildDataLabelTextProperties(c, a),
                    new XElement(c + "dLblPos", new XAttribute("val", "outEnd")),
                    new XElement(c + "showLegendKey", new XAttribute("val", "0")),
                    new XElement(c + "showVal", new XAttribute("val", "1")),
                    new XElement(c + "showCatName", new XAttribute("val", "0")),
                    new XElement(c + "showSerName", new XAttribute("val", "0")),
                    new XElement(c + "showPercent", new XAttribute("val", "0")),
                    new XElement(c + "showBubbleSize", new XAttribute("val", "0")));
            }

            labels.Add(label);
        }

        labels.Add(
            new XElement(c + "showLegendKey", new XAttribute("val", "0")),
            new XElement(c + "showVal", new XAttribute("val", "1")),
            new XElement(c + "showCatName", new XAttribute("val", "0")),
            new XElement(c + "showSerName", new XAttribute("val", "0")),
            new XElement(c + "showPercent", new XAttribute("val", "0")),
            new XElement(c + "showBubbleSize", new XAttribute("val", "0")),
            new XElement(c + "showLeaderLines", new XAttribute("val", "0")));

        var category = series.Element(c + "cat") ?? throw new InvalidOperationException("Chart categories were missing.");
        category.AddBeforeSelf(labels);
    }

    private static XElement BuildDataLabelTextProperties(XNamespace c, XNamespace a)
    {
        return new XElement(
            c + "txPr",
            new XElement(a + "bodyPr", new XAttribute("anchorCtr", "1")),
            new XElement(a + "lstStyle"),
            new XElement(
                a + "p",
                new XElement(
                    a + "pPr",
                    new XElement(
                        a + "defRPr",
                        new XAttribute("sz", "900"),
                        new XAttribute("b", "1"),
                        new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", "004751"))))),
                new XElement(a + "endParaRPr", new XAttribute("lang", "en-US"))));
    }

    private static void ReplaceStringPoints(XElement cache, XNamespace c, IReadOnlyList<string> values)
    {
        cache.RemoveNodes();
        cache.Add(new XElement(c + "ptCount", new XAttribute("val", values.Count)));
        for (var index = 0; index < values.Count; index++)
        {
            cache.Add(
                new XElement(
                    c + "pt",
                    new XAttribute("idx", index),
                    new XElement(c + "v", values[index])));
        }
    }

    private static void ReplaceNumberPoints(XElement cache, XNamespace c, IReadOnlyList<int> values)
    {
        cache.RemoveNodes();
        cache.Add(
            new XElement(c + "formatCode", "0"),
            new XElement(c + "ptCount", new XAttribute("val", values.Count)));
        for (var index = 0; index < values.Count; index++)
        {
            cache.Add(
                new XElement(
                    c + "pt",
                    new XAttribute("idx", index),
                    new XElement(c + "v", values[index].ToString(CultureInfo.InvariantCulture))));
        }
    }

    private static void UpdateEmbeddedWorkbook(
        ZipArchive archive,
        string workbookPath,
        string priorLabel,
        string currentLabel,
        IReadOnlyList<string> categories,
        IReadOnlyList<int> priorValues,
        IReadOnlyList<int> currentValues)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "Claims Range";
        sheet.Cell(1, 2).Value = priorLabel;
        sheet.Cell(1, 3).Value = currentLabel;
        sheet.Range(1, 1, 1, 3).Style.Font.Bold = true;
        for (var index = 0; index < categories.Count; index++)
        {
            var row = index + 2;
            sheet.Cell(row, 1).Value = categories[index];
            sheet.Cell(row, 2).Value = priorValues[index];
            sheet.Cell(row, 3).Value = currentValues[index];
        }

        sheet.Range(2, 2, categories.Count + 1, 3).Style.NumberFormat.Format = "0";
        sheet.Columns(1, 3).AdjustToContents(8, 24);
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;

        archive.GetEntry(workbookPath)?.Delete();
        var replacement = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var stream = replacement.Open();
        output.CopyTo(stream);
    }

    private static string EnsureEmbeddedWorkbookRelationship(
        ZipArchive archive,
        string chartPath,
        XDocument chartDocument,
        XNamespace c)
    {
        var relationshipsPath = $"{Path.GetDirectoryName(chartPath)?.Replace('\\', '/')}/_rels/{Path.GetFileName(chartPath)}.rels";
        var relationshipsEntry = archive.GetEntry(relationshipsPath);
        XDocument relationships;
        if (relationshipsEntry is null)
        {
            relationships = new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(XNamespace.Get(PackageRelationshipNamespace) + "Relationships"));
        }
        else
        {
            relationships = XDocument.Parse(ReadZipEntryText(relationshipsEntry));
        }

        var relationship = relationships.Root?.Elements().FirstOrDefault(element =>
            ((string?)element.Attribute("Type"))?.EndsWith("/package", StringComparison.OrdinalIgnoreCase) == true &&
            ((string?)element.Attribute("Target"))?.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) == true);
        if (relationships.Root is null)
        {
            throw new InvalidOperationException($"Chart relationships were invalid: {relationshipsPath}");
        }

        if (relationship is null)
        {
            var relationshipId = NextRelationshipId(relationships.Root);
            var workbookFileName = $"ClaimsRange_{Path.GetFileNameWithoutExtension(chartPath)}.xlsx";
            var target = $"../embeddings/{workbookFileName}";
            relationship = new XElement(
                XNamespace.Get(PackageRelationshipNamespace) + "Relationship",
                new XAttribute("Id", relationshipId),
                new XAttribute("Type", PackageRelationshipType),
                new XAttribute("Target", target));
            relationships.Root.Add(relationship);
        }

        var relationshipIdentifier = (string?)relationship.Attribute("Id")
            ?? throw new InvalidOperationException("The embedded workbook relationship did not have an id.");
        var relationshipTarget = (string?)relationship.Attribute("Target")
            ?? throw new InvalidOperationException("The embedded workbook relationship did not have a target.");
        ReplaceZipEntry(archive, relationshipsPath, relationships);

        var r = XNamespace.Get(RelationshipNamespace);
        var chartSpace = chartDocument.Root
            ?? throw new InvalidOperationException($"Chart XML was empty: {chartPath}");
        var externalData = chartSpace.Element(c + "externalData");
        if (externalData is null)
        {
            externalData = new XElement(c + "externalData");
            chartSpace.Add(externalData);
        }

        externalData.SetAttributeValue(r + "id", relationshipIdentifier);
        externalData.RemoveNodes();
        externalData.Add(new XElement(c + "autoUpdate", new XAttribute("val", "0")));
        return ResolvePackagePath(chartPath, relationshipTarget);
    }

    private static string NextRelationshipId(XElement relationshipsRoot)
    {
        var usedIds = relationshipsRoot.Elements()
            .Select(element => (string?)element.Attribute("Id"))
            .Where(identifier => !string.IsNullOrWhiteSpace(identifier))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        for (var index = 1; ; index++)
        {
            var candidate = $"rId{index}";
            if (!usedIds.Contains(candidate))
            {
                return candidate;
            }
        }
    }

    private static bool HasSupportedChartStructure(ZipArchive archive, string chartPath)
    {
        var entry = archive.GetEntry(chartPath);
        if (entry is null)
        {
            return false;
        }

        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get(ChartNamespace);
        var barCharts = document.Descendants(c + "barChart").ToList();
        var valueAxes = document.Descendants(c + "valAx").ToList();
        return barCharts.Count == 1 &&
               barCharts[0].Elements(c + "ser").Count() == 2 &&
               valueAxes.Count == 1 &&
               valueAxes[0].Element(c + "scaling")?.Element(c + "min") is not null &&
               valueAxes[0].Element(c + "scaling")?.Element(c + "max") is not null &&
               valueAxes[0].Element(c + "majorUnit") is not null;
    }

    private static string? ResolveChartPath(
        string slidePath,
        XElement chartFrame,
        XDocument relationships,
        XNamespace c,
        XNamespace r)
    {
        var relationshipId = (string?)chartFrame.Descendants(c + "chart").SingleOrDefault()?.Attribute(r + "id");
        if (string.IsNullOrWhiteSpace(relationshipId))
        {
            return null;
        }

        var relationship = relationships.Root?.Elements()
            .SingleOrDefault(element =>
                ((string?)element.Attribute("Id")) == relationshipId &&
                ((string?)element.Attribute("Type"))?.EndsWith("/chart", StringComparison.OrdinalIgnoreCase) == true);
        var target = (string?)relationship?.Attribute("Target");
        return string.IsNullOrWhiteSpace(target) ? null : ResolvePackagePath(slidePath, target);
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
            if (segment == ".")
            {
                continue;
            }

            if (segment == "..")
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

    private static void RemoveSlides(string presentationPath, IReadOnlyCollection<string> slidePaths)
    {
        var normalizedPaths = slidePaths
            .Select(path => "/" + path.TrimStart('/'))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        using var presentation = PresentationDocument.Open(presentationPath, true);
        var presentationPart = presentation.PresentationPart
            ?? throw new InvalidOperationException("The PowerPoint presentation part was missing.");
        var slideIdList = presentationPart.Presentation.SlideIdList;
        if (slideIdList is null)
        {
            return;
        }

        foreach (var slideId in slideIdList.Elements<SlideId>().ToList())
        {
            var relationshipId = slideId.RelationshipId?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId) ||
                presentationPart.GetPartById(relationshipId) is not SlidePart slidePart ||
                !normalizedPaths.Contains(slidePart.Uri.OriginalString))
            {
                continue;
            }

            slideId.Remove();
            presentationPart.DeletePart(slidePart);
        }

        presentationPart.Presentation.Save();
    }

    private static bool IsSlideEntry(ZipArchiveEntry entry)
    {
        return entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
               entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
               !entry.FullName.Contains("/_rels/", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasClaimsRangeTitle(XDocument slide)
    {
        var p = XNamespace.Get(PresentationNamespace);
        var a = XNamespace.Get(DrawingNamespace);
        return slide.Descendants(p + "sp")
            .Any(shape => ConcatenateText(shape, a).Trim().Equals(ClaimsRangeTitle, StringComparison.OrdinalIgnoreCase));
    }

    private static bool HasLegacyClaimsRangeSignature(XDocument slide)
    {
        var distinctiveNames = new[]
        {
            "prior-legend",
            "current-legend",
            "ehc-heading",
            "dental-heading",
            "takeaway-copy",
            "footnote"
        };
        return distinctiveNames.Count(name => HasNamedObject(slide, name)) >= 5 &&
               HasNamedObject(slide, "ehc-heading") &&
               HasNamedObject(slide, "dental-heading");
    }

    private static bool HasNamedObject(XDocument slide, string name)
    {
        return FindObjectsByName(slide, name).Count > 0;
    }

    private static List<XElement> FindObjectsByName(XDocument slide, string name)
    {
        var p = XNamespace.Get(PresentationNamespace);
        return slide.Descendants()
            .Where(element => element.Name == p + "sp" || element.Name == p + "graphicFrame")
            .Where(element => string.Equals(
                (string?)element.Descendants(p + "cNvPr").FirstOrDefault()?.Attribute("name"),
                name,
                StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    private static void SetObjectName(XElement element, string name)
    {
        var p = XNamespace.Get(PresentationNamespace);
        var properties = element.Descendants(p + "cNvPr").FirstOrDefault()
            ?? throw new InvalidOperationException("A slide object was missing non-visual properties.");
        properties.SetAttributeValue("name", name);
    }

    private static void ReplaceNamedShapeText(XDocument slide, string shapeName, string value)
    {
        var shape = FindObjectsByName(slide, shapeName).Single();
        var a = XNamespace.Get(DrawingNamespace);
        var textNodes = shape.Descendants(a + "t").ToList();
        if (textNodes.Count == 0)
        {
            throw new InvalidOperationException($"Shape '{shapeName}' did not contain editable text.");
        }

        textNodes[0].Value = value;
        foreach (var extraText in textNodes.Skip(1))
        {
            extraText.Value = "";
        }
    }

    private static string ConcatenateText(XElement element, XNamespace a)
    {
        return string.Concat(element.Descendants(a + "t").Select(text => text.Value));
    }

    private static long ParseLong(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result)
            ? result
            : long.MaxValue;
    }

    private static void SetFormula(XElement reference, XNamespace c, string formula)
    {
        var element = reference.Element(c + "f");
        if (element is null)
        {
            element = new XElement(c + "f");
            reference.AddFirst(element);
        }

        element.Value = formula;
    }

    private static XElement EnsureChild(XElement parent, XName name)
    {
        var element = parent.Element(name);
        if (element is not null)
        {
            return element;
        }

        element = new XElement(name);
        parent.Add(element);
        return element;
    }

    private static void SetRequiredValue(XElement? element, int value)
    {
        if (element is null)
        {
            throw new InvalidOperationException("A required chart scaling element was missing.");
        }

        element.SetAttributeValue("val", value.ToString(CultureInfo.InvariantCulture));
    }

    private static void SetRequiredValue(XElement? element, double value)
    {
        if (element is null)
        {
            throw new InvalidOperationException("A required chart scaling element was missing.");
        }

        element.SetAttributeValue("val", value.ToString("0.###############", CultureInfo.InvariantCulture));
    }

    private static string ReadZipEntryText(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static void ReplaceZipEntry(ZipArchive archive, string path, XDocument document)
    {
        archive.GetEntry(path)?.Delete();
        var replacement = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = replacement.Open();
        document.Save(stream);
    }

    private sealed record ClaimsRangeSlideContract(
        string SlidePath,
        XDocument Slide,
        string ExtendedHealthChartPath,
        string DentalChartPath);
}
