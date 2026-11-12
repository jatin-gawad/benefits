using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class ClaimsExperienceDataModelBuilder
{
    private sealed record ClaimsSourcePeriodDefinition(
        int Rank,
        int SortOrder,
        string Role,
        string PeriodId,
        string PeriodKey,
        string PeriodLabel,
        DateOnly SourceStart,
        DateOnly SourceEnd,
        DateOnly SelectedEnd,
        int SourceMonthCount,
        int SelectedMonthCount,
        string SourceFiles);

    private sealed record ClaimsFactExportSelection(
        ClaimsServiceAdjudicationFact Fact,
        ClaimsSourcePeriodDefinition Period,
        string RelatedSourceFiles,
        bool HasChangedRevision);

    private static List<object?[]> BuildDynamicPeriodRows(ClaimsExperienceDataModelData data)
    {
        var sourcePeriods = BuildRankedSourcePeriodDefinitions(data, applyComparisonOverride: false);
        if (sourcePeriods.Count == 0)
        {
            return BuildLegacyDynamicPeriodRows(data);
        }

        var rows = new List<object?[]>();
        foreach (var period in sourcePeriods.OrderByDescending(item => item.Rank))
        {
            var windowCounts = new[] { data.ReportMonthCount, period.SourceMonthCount, 6, 9, 12 }
                .Where(count => count > 0 && count <= period.SourceMonthCount)
                .Distinct()
                .OrderByDescending(count => count == data.ReportMonthCount)
                .ThenByDescending(count => count)
                .ToList();

            foreach (var monthCount in windowCounts)
            {
                var end = monthCount < period.SourceMonthCount
                    ? EndOfMonth(period.SourceStart.AddMonths(monthCount - 1))
                    : period.SourceEnd;
                rows.Add(
                [
                    PeriodIdForRank(period.Rank, monthCount),
                    FormatPeriodLabel(period.SourceStart, end),
                    ToDateTime(period.SourceStart),
                    ToDateTime(end),
                    monthCount,
                    period.Role,
                    period.Rank == 0 && monthCount == data.ReportMonthCount,
                    monthCount == period.SourceMonthCount
                        ? $"Reported in {period.SourceFiles}."
                        : $"{monthCount}-month comparison window within the source period reported in {period.SourceFiles}."
                ]);
            }
        }

        rows.Add([PeriodIds.CustomCurrent, "Custom Current", "", "", "", "Current", false, "Manual custom period setup."]);
        rows.Add([PeriodIds.CustomPrior, "Custom Prior", "", "", "", "Prior", false, "Manual custom period setup."]);
        return rows;
    }

    private static List<object?[]> BuildLegacyDynamicPeriodRows(ClaimsExperienceDataModelData data)
    {
        var periodStart = data.PeriodStart ?? new DateOnly(DateTime.Today.Year, 1, 1);
        var sourceMonthCount = data.DetectedReportMonthCount > 0
            ? data.DetectedReportMonthCount
            : data.ReportMonthCount;
        var full12End = EndOfMonth(periodStart.AddMonths(11));
        var current9End = EndOfMonth(periodStart.AddMonths(8));
        var current6End = EndOfMonth(periodStart.AddMonths(5));
        var priorStart = periodStart.AddYears(-1);
        var priorPriorStart = periodStart.AddYears(-2);
        List<object?[]> rows =
        [
            [PeriodIds.PriorPrior12, FormatPeriodLabel(priorPriorStart, full12End.AddYears(-2)), ToDateTime(priorPriorStart), ToDateTime(full12End.AddYears(-2)), 12, "Prior Prior", false, "Full 12 months before the prior period."],
            [PeriodIds.Current12, FormatPeriodLabel(periodStart, full12End), ToDateTime(periodStart), ToDateTime(full12End), 12, "Current", data.ReportMonthCount == 12, "Twelve-month reporting option."],
            [PeriodIds.Prior12, FormatPeriodLabel(priorStart, full12End.AddYears(-1)), ToDateTime(priorStart), ToDateTime(full12End.AddYears(-1)), 12, "Prior", false, "Prior twelve-month comparison option."],
            [PeriodIds.Current9, FormatPeriodLabel(periodStart, current9End), ToDateTime(periodStart), ToDateTime(current9End), 9, "Current", data.ReportMonthCount == 9, "First 9 months of the source period."],
            [PeriodIds.Prior9, FormatPeriodLabel(priorStart, current9End.AddYears(-1)), ToDateTime(priorStart), ToDateTime(current9End.AddYears(-1)), 9, "Prior", false, "Matched prior 9-month period."],
            [PeriodIds.Current6, FormatPeriodLabel(periodStart, current6End), ToDateTime(periodStart), ToDateTime(current6End), 6, "Current", data.ReportMonthCount == 6, "First 6 months of the source period."],
            [PeriodIds.Prior6, FormatPeriodLabel(priorStart, current6End.AddYears(-1)), ToDateTime(priorStart), ToDateTime(current6End.AddYears(-1)), 6, "Prior", false, "Matched prior 6-month period."]
        ];

        if (data.ReportMonthCount is not (6 or 9 or 12))
        {
            var selectedEnd = data.PeriodEnd ?? EndOfMonth(periodStart.AddMonths(data.ReportMonthCount - 1));
            rows.Insert(1,
                [data.CurrentSourcePeriodId, FormatPeriodLabel(periodStart, selectedEnd), ToDateTime(periodStart), ToDateTime(selectedEnd), data.ReportMonthCount, "Current", true, $"Automatically detected from the {sourceMonthCount}-month source report."]);
            rows.Insert(2,
                [data.PriorSourcePeriodId, FormatPeriodLabel(priorStart, selectedEnd.AddYears(-1)), ToDateTime(priorStart), ToDateTime(selectedEnd.AddYears(-1)), data.ReportMonthCount, "Prior", false, "Matched prior-period service claims from the current source set."]);
        }

        rows.Add([PeriodIds.CustomCurrent, "Custom Current", "", "", "", "Current", false, "Manual custom period setup."]);
        rows.Add([PeriodIds.CustomPrior, "Custom Prior", "", "", "", "Prior", false, "Manual custom period setup."]);
        return rows;
    }

    private static List<ClaimsSourcePeriodDefinition> BuildRankedSourcePeriodDefinitions(
        ClaimsExperienceDataModelData data,
        bool applyComparisonOverride)
    {
        var candidates = data.SourcePeriods
            .Select(period => (Start: period.Start, End: period.End, SourceFile: period.SourceFile))
            .Concat(data.ServiceAdjudicationFacts.Select(fact => (Start: fact.PeriodStart, End: fact.PeriodEnd, SourceFile: fact.SourceFile)))
            .Concat(data.ServiceControlTotals.Select(total => (Start: total.PeriodStart, End: total.PeriodEnd, SourceFile: total.SourceFile)))
            .Where(period => period.End >= period.Start)
            .GroupBy(period => PeriodDateKey(period.Start, period.End), StringComparer.OrdinalIgnoreCase)
            .Select(group => new
            {
                Start = group.First().Start,
                End = group.First().End,
                SourceFiles = string.Join("; ", group.Select(item => item.SourceFile)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
            })
            .OrderByDescending(period => period.End)
            .ThenByDescending(period => period.Start)
            .Take(4)
            .ToList();

        var definitions = new List<ClaimsSourcePeriodDefinition>();
        for (var rank = 0; rank < candidates.Count; rank++)
        {
            var source = candidates[rank];
            var sourceMonthCount = InclusiveMonthCount(source.Start, source.End);
            var requested = applyComparisonOverride && data.RequestedReportMonthCount is 6 or 9 or 12
                ? data.RequestedReportMonthCount
                : sourceMonthCount;
            var selectedMonthCount = requested > 0 && requested <= sourceMonthCount
                ? requested
                : sourceMonthCount;
            var selectedEnd = selectedMonthCount < sourceMonthCount
                ? EndOfMonth(source.Start.AddMonths(selectedMonthCount - 1))
                : source.End;
            var role = PeriodRoleForRank(rank);
            definitions.Add(new ClaimsSourcePeriodDefinition(
                rank,
                candidates.Count - rank,
                role,
                PeriodIdForRank(rank, selectedMonthCount),
                BuildSourcePeriodKey(data, source.Start, source.End),
                FormatPeriodLabel(source.Start, selectedEnd),
                source.Start,
                source.End,
                selectedEnd,
                sourceMonthCount,
                selectedMonthCount,
                string.IsNullOrWhiteSpace(source.SourceFiles) ? "Source document" : source.SourceFiles));
        }

        return definitions;
    }

    private static List<object?[]> BuildCanonicalClaimsPeriodRows(ClaimsExperienceDataModelData data)
    {
        return BuildRankedSourcePeriodDefinitions(data, applyComparisonOverride: false)
            .OrderBy(period => period.Rank)
            .Select(period => new object?[]
            {
                period.Rank + 1,
                period.Role,
                period.PeriodKey,
                period.PeriodId,
                period.PeriodLabel,
                ToDateTime(period.SourceStart),
                ToDateTime(period.SourceEnd),
                period.SourceMonthCount,
                period.SourceFiles
            })
            .ToList();
    }

    private static List<object?[]> BuildCanonicalClaimsAdjudicationRows(ClaimsExperienceDataModelData data)
    {
        return SelectCanonicalServiceFactsForExport(data)
            .OrderBy(selection => selection.Period.Rank)
            .ThenBy(selection => selection.Fact.Benefit, StringComparer.OrdinalIgnoreCase)
            .ThenBy(selection => selection.Fact.Relationship, StringComparer.OrdinalIgnoreCase)
            .ThenBy(selection => selection.Fact.Category, StringComparer.OrdinalIgnoreCase)
            .ThenBy(selection => selection.Fact.Service, StringComparer.OrdinalIgnoreCase)
            .Select(selection => new object?[]
            {
                selection.Period.PeriodKey,
                selection.Period.Role,
                selection.Fact.Benefit,
                selection.Fact.Relationship,
                selection.Fact.Category,
                selection.Fact.Service,
                selection.Fact.AmountSubmitted,
                selection.Fact.AmountEligible,
                selection.Fact.AmountPaid,
                selection.Fact.ReportedPaidPercentage / 100m,
                selection.Fact.Currency,
                selection.Fact.SourceFile,
                selection.RelatedSourceFiles,
                selection.Fact.SourcePage,
                selection.Fact.SourceEvidence,
                selection.Fact.Confidence,
                selection.HasChangedRevision ? "NeedsReview" : selection.Fact.ValidationState
            })
            .ToList();
    }

    private static List<object?[]> BuildCanonicalClaimsControlRows(ClaimsExperienceDataModelData data)
    {
        var periods = BuildRankedSourcePeriodDefinitions(data, applyComparisonOverride: false);
        var byDate = periods.ToDictionary(period => PeriodDateKey(period.SourceStart, period.SourceEnd), StringComparer.OrdinalIgnoreCase);
        return data.ServiceControlTotals
            .Where(total => byDate.ContainsKey(PeriodDateKey(total.PeriodStart, total.PeriodEnd)))
            .GroupBy(total => string.Join("|",
                PeriodDateKey(total.PeriodStart, total.PeriodEnd),
                total.Benefit,
                total.TotalType,
                total.Relationship,
                total.Category,
                total.Service,
                total.AmountSubmitted,
                total.AmountEligible,
                total.AmountPaid,
                total.ReportedPaidPercentage), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(total => byDate[PeriodDateKey(total.PeriodStart, total.PeriodEnd)].Rank)
            .ThenBy(total => total.Benefit, StringComparer.OrdinalIgnoreCase)
            .ThenBy(total => total.TotalType, StringComparer.OrdinalIgnoreCase)
            .Select(total =>
            {
                var period = byDate[PeriodDateKey(total.PeriodStart, total.PeriodEnd)];
                return new object?[]
                {
                    period.PeriodKey,
                    period.Role,
                    total.TotalType,
                    total.Benefit,
                    total.Relationship,
                    total.Category,
                    total.Service,
                    total.AmountSubmitted,
                    total.AmountEligible,
                    total.AmountPaid,
                    total.ReportedPaidPercentage / 100m,
                    total.Currency,
                    total.SourceFile,
                    total.SourcePage,
                    total.SourceEvidence
                };
            })
            .ToList();
    }

    private static List<object?[]> BuildCanonicalClaimsRevisionRows(ClaimsExperienceDataModelData data)
    {
        var periods = BuildRankedSourcePeriodDefinitions(data, applyComparisonOverride: false);
        var byDate = periods.ToDictionary(period => PeriodDateKey(period.SourceStart, period.SourceEnd), StringComparer.OrdinalIgnoreCase);
        var rows = new List<object?[]>();
        var facts = data.ServiceAdjudicationFacts
            .Where(fact => byDate.ContainsKey(PeriodDateKey(fact.PeriodStart, fact.PeriodEnd)))
            .GroupBy(fact => CanonicalFactNaturalKey(data, fact), StringComparer.OrdinalIgnoreCase);

        foreach (var group in facts)
        {
            var valueGroups = group.GroupBy(CanonicalFactValueKey, StringComparer.OrdinalIgnoreCase).ToList();
            if (valueGroups.Count <= 1)
            {
                continue;
            }

            var chosen = LatestFact(data, group);
            var period = byDate[PeriodDateKey(chosen.PeriodStart, chosen.PeriodEnd)];
            foreach (var prior in valueGroups
                         .Select(values => LatestFact(data, values))
                         .Where(fact => CanonicalFactValueKey(fact) != CanonicalFactValueKey(chosen)))
            {
                rows.Add(
                [
                    period.PeriodKey,
                    period.Role,
                    group.Key,
                    $"{prior.SourceFile} p{prior.SourcePage}",
                    $"{chosen.SourceFile} p{chosen.SourcePage}",
                    prior.AmountSubmitted,
                    chosen.AmountSubmitted,
                    prior.AmountEligible,
                    chosen.AmountEligible,
                    prior.AmountPaid,
                    chosen.AmountPaid,
                    "NeedsReview"
                ]);
            }
        }

        return rows;
    }

    private static List<ClaimsFactExportSelection> SelectCanonicalServiceFactsForExport(ClaimsExperienceDataModelData data)
    {
        var periods = BuildRankedSourcePeriodDefinitions(data, applyComparisonOverride: false);
        var byDate = periods.ToDictionary(period => PeriodDateKey(period.SourceStart, period.SourceEnd), StringComparer.OrdinalIgnoreCase);
        var selections = new List<ClaimsFactExportSelection>();
        foreach (var group in data.ServiceAdjudicationFacts
                     .Where(fact => byDate.ContainsKey(PeriodDateKey(fact.PeriodStart, fact.PeriodEnd)))
                     .GroupBy(fact => CanonicalFactNaturalKey(data, fact), StringComparer.OrdinalIgnoreCase))
        {
            var chosen = LatestFact(data, group);
            var sourceFiles = string.Join("; ", group
                .Select(fact => $"{fact.SourceFile} p{fact.SourcePage}")
                .Distinct(StringComparer.OrdinalIgnoreCase));
            var changed = group.Select(CanonicalFactValueKey).Distinct(StringComparer.OrdinalIgnoreCase).Skip(1).Any();
            selections.Add(new ClaimsFactExportSelection(
                chosen,
                byDate[PeriodDateKey(chosen.PeriodStart, chosen.PeriodEnd)],
                sourceFiles,
                changed));
        }

        return selections;
    }

    private static ClaimsServiceAdjudicationFact LatestFact(
        ClaimsExperienceDataModelData data,
        IEnumerable<ClaimsServiceAdjudicationFact> facts)
    {
        return facts
            .OrderByDescending(fact => SourceModifiedUtc(data, fact.SourceFile))
            .ThenByDescending(fact => fact.SourceFile, StringComparer.OrdinalIgnoreCase)
            .ThenByDescending(fact => fact.SourcePage)
            .First();
    }

    private static DateTime SourceModifiedUtc(ClaimsExperienceDataModelData data, string sourceFile)
    {
        var path = data.ImportedFiles
            .FirstOrDefault(item => item.SourceFile.Equals(sourceFile, StringComparison.OrdinalIgnoreCase))
            ?.FullPath;
        try
        {
            return !string.IsNullOrWhiteSpace(path) && File.Exists(path)
                ? File.GetLastWriteTimeUtc(path)
                : DateTime.MinValue;
        }
        catch
        {
            return DateTime.MinValue;
        }
    }

    private static string CanonicalFactNaturalKey(ClaimsExperienceDataModelData data, ClaimsServiceAdjudicationFact fact)
    {
        return string.Join("|",
            data.ClientName.Trim(),
            data.Carrier.Trim(),
            data.PolicyNumber.Trim(),
            fact.Benefit.Trim(),
            PeriodDateKey(fact.PeriodStart, fact.PeriodEnd),
            NormalizeRelationshipKey(fact.Relationship),
            fact.Category.Trim(),
            fact.Service.Trim());
    }

    private static string CanonicalFactValueKey(ClaimsServiceAdjudicationFact fact)
    {
        return string.Join("|", fact.AmountSubmitted, fact.AmountEligible, fact.AmountPaid, fact.ReportedPaidPercentage, fact.Currency);
    }

    private static string NormalizeRelationshipKey(string value)
    {
        return new string((value ?? "").Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
    }

    private static string PeriodDateKey(DateOnly start, DateOnly end)
    {
        return $"{start:yyyy-MM-dd}|{end:yyyy-MM-dd}";
    }

    private static string BuildSourcePeriodKey(ClaimsExperienceDataModelData data, DateOnly start, DateOnly end)
    {
        return string.Join("|", data.Carrier.Trim(), data.PolicyNumber.Trim(), start.ToString("yyyy-MM-dd"), end.ToString("yyyy-MM-dd"));
    }

    private static string PeriodRoleForRank(int rank)
    {
        return rank switch
        {
            0 => "Current",
            1 => "Prior",
            2 => "Prior Prior",
            3 => "Prior Prior Prior",
            _ => $"Prior {rank}"
        };
    }

    private static string PeriodIdForRank(int rank, int monthCount)
    {
        var prefix = rank switch
        {
            0 => "CURR",
            1 => "PRIOR",
            2 => "PRIOR_PRIOR",
            3 => "PRIOR_PRIOR_PRIOR",
            _ => $"PRIOR_{rank}"
        };
        return $"{prefix}_{monthCount}";
    }

    private static string PeriodIdFormulaForRole(string role)
    {
        var prefix = NormalizeRelationshipKey(role) switch
        {
            "CURRENT" => "CURR",
            "PRIOR" => "PRIOR",
            "PRIORPRIOR" => "PRIOR_PRIOR",
            "PRIORPRIORPRIOR" => "PRIOR_PRIOR_PRIOR",
            _ => "PRIOR"
        };
        return $"{Quote(prefix + "_")}&{ReportMonthCountFormula()}";
    }

    private static string CanonicalOrMonthlyPaidClaimsFormula(int excelRow, string coverage)
    {
        var canonicalPeriodId = LookupFormula(
            $"B{excelRow}",
            "tblPPT_CanonicalClaimsPeriods[Period Role]",
            "tblPPT_CanonicalClaimsPeriods[Period ID]",
            Quote(""));
        var canonicalPaid = $"SUMIFS(tblPPT_CanonicalClaimsAdjudication[Amount Paid],tblPPT_CanonicalClaimsAdjudication[Period Role],B{excelRow},tblPPT_CanonicalClaimsAdjudication[Benefit],{Quote(coverage)})";
        var monthlyPaid = SumMonthlyFormula($"C{excelRow}", coverage, "Paid Claims");
        return $"IF(C{excelRow}=({canonicalPeriodId}),{canonicalPaid},{monthlyPaid})";
    }
}