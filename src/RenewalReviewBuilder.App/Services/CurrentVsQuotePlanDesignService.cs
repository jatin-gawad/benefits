using System.Globalization;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class CurrentVsQuotePlanDesignService
{
    private const string MissingValue = "Not found";
    private readonly PlanDesignCanonicalProfileService _profileService;

    public CurrentVsQuotePlanDesignService(string? databasePath = null)
        : this(new PlanDesignCanonicalProfileService(databasePath))
    {
    }

    public CurrentVsQuotePlanDesignService(PlanDesignCanonicalProfileService profileService)
    {
        _profileService = profileService ?? throw new ArgumentNullException(nameof(profileService));
    }

    public async Task<CurrentVsQuotePlanDesignComparison> BuildAsync(long clientId)
    {
        var profile = await _profileService.BuildAsync(clientId);
        return Build(profile);
    }

    public CurrentVsQuotePlanDesignComparison Build(PlanDesignCanonicalProfile profile)
    {
        ArgumentNullException.ThrowIfNull(profile);

        var result = new CurrentVsQuotePlanDesignComparison
        {
            ClientId = profile.ClientId,
            ClientName = profile.ClientName,
            ClientNumber = profile.ClientNumber,
            CurrentCarrier = profile.Carrier,
            ProfileVersion = profile.ProfileVersion,
            ReadinessStatus = profile.ReadinessStatus,
            BuiltAtUtc = profile.BuiltAtUtc,
            HasCurrentPlan = profile.HasCurrentPlan,
            HasQuotePlan = profile.HasQuotePlan,
            CanCompare = profile.CanCompareCurrentVsQuote
        };
        result.Issues.AddRange(profile.Issues);

        var quoteGroups = profile.QuoteFacts
            .GroupBy(BuildQuoteKey, StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => FirstNonBlank(group.First().Carrier, group.First().SourceFileName), StringComparer.OrdinalIgnoreCase)
            .ThenBy(group => group.First().ProposalId)
            .ToList();

        foreach (var group in quoteGroups)
        {
            var first = SelectBestFact(group);
            var carrier = FirstNonBlank(first.Carrier, first.SourceCarrier, "Saved Quote");
            var sourceFileName = FirstNonBlank(first.SourceFileName, first.SourceDocument);
            result.Quotes.Add(new CurrentVsQuotePlanDesignQuote
            {
                QuoteKey = group.Key,
                ProposalId = first.ProposalId,
                Carrier = carrier,
                SourceFileName = sourceFileName,
                DisplayName = carrier,
                Reference = first.ProposalId.HasValue
                    ? $"Proposal {first.ProposalId.Value.ToString(CultureInfo.InvariantCulture)}"
                    : Path.GetFileNameWithoutExtension(sourceFileName)
            });
        }

        var currentByKey = profile.CurrentPlanFacts
            .Where(fact => !string.IsNullOrWhiteSpace(fact.ComparisonKey))
            .GroupBy(fact => fact.ComparisonKey, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, SelectBestFact, StringComparer.OrdinalIgnoreCase);
        var quotesByKey = quoteGroups.ToDictionary(
            group => group.Key,
            group => group
                .Where(fact => !string.IsNullOrWhiteSpace(fact.ComparisonKey))
                .GroupBy(fact => fact.ComparisonKey, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(facts => facts.Key, SelectBestFact, StringComparer.OrdinalIgnoreCase),
            StringComparer.OrdinalIgnoreCase);
        var validationsByQuote = profile.ProposalValidationFacts
            .Where(fact => !string.IsNullOrWhiteSpace(fact.ComparisonKey))
            .GroupBy(BuildQuoteKey, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                group => group.Key,
                group => group
                    .GroupBy(fact => fact.ComparisonKey, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(facts => facts.Key, SelectBestFact, StringComparer.OrdinalIgnoreCase),
                StringComparer.OrdinalIgnoreCase);

        var comparisonKeys = currentByKey.Keys
            .Concat(quotesByKey.Values.SelectMany(rows => rows.Keys))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var comparisonKey in comparisonKeys)
        {
            currentByKey.TryGetValue(comparisonKey, out var currentFact);
            var representative = currentFact ?? result.Quotes
                .Select(quote => quotesByKey[quote.QuoteKey].GetValueOrDefault(comparisonKey))
                .FirstOrDefault(fact => fact is not null);
            if (representative is null)
            {
                continue;
            }

            var quoteValues = new List<CurrentVsQuotePlanDesignValue>();
            foreach (var quote in result.Quotes)
            {
                quotesByKey[quote.QuoteKey].TryGetValue(comparisonKey, out var quoteFact);
                PlanDesignCanonicalFact? validationFact = null;
                if (validationsByQuote.TryGetValue(quote.QuoteKey, out var quoteValidations) &&
                    quoteValidations.TryGetValue(comparisonKey, out var matchedValidation))
                {
                    validationFact = matchedValidation;
                }

                quoteValues.Add(BuildQuoteValue(quote, currentFact, quoteFact, validationFact));
            }

            var row = new CurrentVsQuotePlanDesignRow
            {
                ComparisonKey = comparisonKey,
                CanonicalBenefitCode = representative.CanonicalBenefitCode,
                BenefitCategory = FirstNonBlank(representative.BenefitCategory, representative.BenefitName, "Other"),
                BenefitName = FirstNonBlank(representative.BenefitName, representative.BenefitCategory, representative.FeatureName, "Plan Design"),
                ClassName = FirstNonBlank(representative.ClassName, "All Classes"),
                ProvisionName = FirstNonBlank(representative.FeatureName, representative.BenefitName, representative.BenefitCategory, "Provision"),
                CurrentValue = currentFact is null ? MissingValue : FormatValue(currentFact),
                CurrentAvailability = currentFact?.Availability ?? PlanDesignCanonicalAvailability.Missing,
                CurrentEvidence = currentFact?.Evidence,
                Status = ResolveRowStatus(currentFact, quoteValues)
            };
            row.QuoteValues.AddRange(quoteValues);
            result.Rows.Add(row);
        }

        result.Rows.Sort(CompareRows);
        return result;
    }

    private static CurrentVsQuotePlanDesignValue BuildQuoteValue(
        CurrentVsQuotePlanDesignQuote quote,
        PlanDesignCanonicalFact? currentFact,
        PlanDesignCanonicalFact? quoteFact,
        PlanDesignCanonicalFact? validationFact)
    {
        var status = ResolveValueStatus(currentFact, quoteFact, validationFact);
        return new CurrentVsQuotePlanDesignValue
        {
            QuoteKey = quote.QuoteKey,
            ProposalId = quote.ProposalId,
            Carrier = quote.Carrier,
            QuoteDisplayName = quote.DisplayName,
            Value = quoteFact is null ? MissingValue : FormatValue(quoteFact),
            Status = status,
            Availability = quoteFact?.Availability ?? PlanDesignCanonicalAvailability.Missing,
            ValidationStatus = validationFact?.ValidationStatus ?? "",
            Evidence = quoteFact?.Evidence
        };
    }

    private static string ResolveValueStatus(
        PlanDesignCanonicalFact? currentFact,
        PlanDesignCanonicalFact? quoteFact,
        PlanDesignCanonicalFact? validationFact)
    {
        if (validationFact is not null && PlanDesignCanonicalMapper.IsFailedValidation(validationFact))
        {
            return CurrentVsQuotePlanDesignStatuses.NeedsReview;
        }

        if (currentFact?.Availability is PlanDesignCanonicalAvailability.NeedsReview or PlanDesignCanonicalAvailability.Partial ||
            quoteFact?.Availability is PlanDesignCanonicalAvailability.NeedsReview or PlanDesignCanonicalAvailability.Partial)
        {
            return CurrentVsQuotePlanDesignStatuses.NeedsReview;
        }

        if (currentFact is null && quoteFact is null)
        {
            return CurrentVsQuotePlanDesignStatuses.Unavailable;
        }

        if (currentFact is null)
        {
            return CurrentVsQuotePlanDesignStatuses.QuoteOnly;
        }

        if (quoteFact is null)
        {
            return CurrentVsQuotePlanDesignStatuses.CurrentOnly;
        }

        var currentValue = NormalizeValue(FormatValue(currentFact));
        var quoteValue = NormalizeValue(FormatValue(quoteFact));
        if (string.IsNullOrWhiteSpace(currentValue) || string.IsNullOrWhiteSpace(quoteValue))
        {
            return CurrentVsQuotePlanDesignStatuses.NeedsReview;
        }

        return currentValue.Equals(quoteValue, StringComparison.OrdinalIgnoreCase)
            ? CurrentVsQuotePlanDesignStatuses.Same
            : CurrentVsQuotePlanDesignStatuses.Different;
    }

    private static string ResolveRowStatus(
        PlanDesignCanonicalFact? currentFact,
        IReadOnlyList<CurrentVsQuotePlanDesignValue> quoteValues)
    {
        if (quoteValues.Any(value => value.Status == CurrentVsQuotePlanDesignStatuses.NeedsReview))
        {
            return CurrentVsQuotePlanDesignStatuses.NeedsReview;
        }

        if (quoteValues.Any(value => value.Status == CurrentVsQuotePlanDesignStatuses.Different))
        {
            return CurrentVsQuotePlanDesignStatuses.Different;
        }

        if (quoteValues.Any(value => value.Status == CurrentVsQuotePlanDesignStatuses.QuoteOnly))
        {
            return CurrentVsQuotePlanDesignStatuses.QuoteOnly;
        }

        if (currentFact is not null &&
            (quoteValues.Count == 0 || quoteValues.Any(value => value.Status == CurrentVsQuotePlanDesignStatuses.CurrentOnly)))
        {
            return CurrentVsQuotePlanDesignStatuses.CurrentOnly;
        }

        return quoteValues.Count > 0 && quoteValues.All(value => value.Status == CurrentVsQuotePlanDesignStatuses.Same)
            ? CurrentVsQuotePlanDesignStatuses.Same
            : CurrentVsQuotePlanDesignStatuses.Unavailable;
    }

    private static string BuildQuoteKey(PlanDesignCanonicalFact fact)
    {
        if (fact.ProposalId.HasValue)
        {
            return $"proposal:{fact.ProposalId.Value.ToString(CultureInfo.InvariantCulture)}";
        }

        var source = FirstNonBlank(fact.SourceFileHash, fact.SourceFileName, fact.SourceDocument, "saved-quote");
        var carrier = FirstNonBlank(fact.Carrier, fact.SourceCarrier, "unknown-carrier");
        return $"source:{PlanDesignCanonicalMapper.NormalizeComparableText(carrier)}|{source.Trim().ToLowerInvariant()}";
    }

    private static PlanDesignCanonicalFact SelectBestFact(IEnumerable<PlanDesignCanonicalFact> facts)
    {
        return facts
            .OrderByDescending(fact => AvailabilityRank(fact.Availability))
            .ThenByDescending(fact => fact.ConfidenceScore ?? 0m)
            .ThenByDescending(fact => fact.SourceTable.Equals("plan_design_features", StringComparison.OrdinalIgnoreCase))
            .ThenByDescending(fact => fact.DateLoaded ?? DateTime.MinValue)
            .First();
    }

    private static int AvailabilityRank(string availability)
    {
        return availability switch
        {
            PlanDesignCanonicalAvailability.Available => 4,
            PlanDesignCanonicalAvailability.Partial => 3,
            PlanDesignCanonicalAvailability.NeedsReview => 2,
            _ => 1
        };
    }

    private static string FormatValue(PlanDesignCanonicalFact fact)
    {
        if (!string.IsNullOrWhiteSpace(fact.FeatureValue))
        {
            return fact.FeatureValue.Trim();
        }

        if (fact.ReimbursementPercent.HasValue)
        {
            var value = fact.ReimbursementPercent.Value;
            var percent = value is >= 0m and <= 1m ? value * 100m : value;
            return $"{percent.ToString("0.##", CultureInfo.InvariantCulture)}%";
        }

        var details = new List<string>();
        AddDetail(details, "Coinsurance", fact.Coinsurance);
        AddDetail(details, "Deductible", fact.Deductible);
        AddDetail(details, "Maximum", fact.Maximum);
        AddDetail(details, "Frequency", fact.Frequency);
        AddDetail(details, "Waiting period", fact.WaitingPeriod);
        AddDetail(details, "Termination", fact.TerminationAge);
        return details.Count == 0 ? MissingValue : string.Join(" | ", details);
    }

    private static void AddDetail(List<string> details, string label, string value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            details.Add($"{label}: {value.Trim()}");
        }
    }

    private static string NormalizeValue(string value)
    {
        return PlanDesignCanonicalMapper.NormalizeComparableText(value)
            .Replace("per cent", "%", StringComparison.Ordinal)
            .Replace("percent", "%", StringComparison.Ordinal)
            .Replace(" %", "%", StringComparison.Ordinal)
            .Replace("$ ", "$", StringComparison.Ordinal);
    }

    private static int CompareRows(CurrentVsQuotePlanDesignRow left, CurrentVsQuotePlanDesignRow right)
    {
        var category = string.Compare(left.BenefitCategory, right.BenefitCategory, StringComparison.OrdinalIgnoreCase);
        if (category != 0)
        {
            return category;
        }

        var benefit = string.Compare(left.BenefitName, right.BenefitName, StringComparison.OrdinalIgnoreCase);
        return benefit != 0
            ? benefit
            : string.Compare(left.ProvisionName, right.ProvisionName, StringComparison.OrdinalIgnoreCase);
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }
}
