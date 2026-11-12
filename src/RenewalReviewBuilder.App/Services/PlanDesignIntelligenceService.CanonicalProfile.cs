using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class PlanDesignIntelligenceService
{
    public async Task<PlanDesignCanonicalProfile> BuildCanonicalProfileAsync(long clientId)
    {
        var profile = new PlanDesignCanonicalProfile
        {
            ClientId = clientId,
            BuiltAtUtc = DateTime.UtcNow
        };

        if (!File.Exists(_databasePath))
        {
            AddProfileIssue(
                profile,
                "database_missing",
                PlanDesignCanonicalIssueSeverities.Error,
                "Benefits Intelligence database was not found.");
            return profile;
        }

        await using var connection = await OpenReadOnlyConnectionAsync();
        var existingTables = await GetExistingTablesAsync(connection);
        var client = await ReadClientAsync(connection, existingTables, clientId);
        profile.ClientName = client.ClientName;
        profile.ClientNumber = client.ClientNumber;
        profile.Carrier = client.Carrier;

        var clientKeys = BuildClientKeys(clientId, client);
        var proposalIds = await ReadProposalIdsAsync(connection, existingTables, clientKeys, client.ClientName);

        foreach (var tableName in SourceTableNames)
        {
            profile.Sources.Add(await BuildSourceUsageAsync(connection, existingTables, tableName, clientId, clientKeys, proposalIds));
        }

        var currentPlanRows = await ReadPlanDesignRowsAsync(connection, existingTables, clientId);
        currentPlanRows.AddRange(await ReadFeatureRowsAsync(connection, existingTables, clientId));
        var quoteRows = await ReadProposalRowsAsync(connection, existingTables, "proposal_plan_design", clientKeys, proposalIds);
        var validationRows = await ReadProposalRowsAsync(connection, existingTables, "proposal_booklet_validation", clientKeys, proposalIds);

        foreach (var fact in currentPlanRows.Concat(quoteRows).Concat(validationRows))
        {
            fact.ClientId = fact.ClientId == 0 ? clientId : fact.ClientId;
            fact.Carrier = FirstNonBlank(fact.Carrier, profile.Carrier);
            PlanDesignCanonicalMapper.CompleteFact(fact);
            profile.Facts.Add(fact);
        }

        AddCurrentPlanFallbackFromBookletValidation(profile, clientId);

        profile.FactorReferences.AddRange(await ReadPlanDesignFactorsAsync(connection, existingTables));
        profile.ContextBenefitLines.AddRange(await ReadContextBenefitLinesAsync(connection, existingTables, clientId));
        AddCanonicalProfileIssues(profile);
        return profile;
    }

    private static void AddCurrentPlanFallbackFromBookletValidation(PlanDesignCanonicalProfile profile, long clientId)
    {
        if (profile.HasCurrentPlan)
        {
            return;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var validationFact in profile.ProposalValidationFacts.Where(fact => HasText(fact.BookletValue)))
        {
            var key = FirstNonBlank(
                validationFact.ComparisonKey,
                $"{validationFact.BenefitName}|{validationFact.FeatureName}|{validationFact.BookletValue}");
            if (!seen.Add(key))
            {
                continue;
            }

            var currentFact = new PlanDesignCanonicalFact
            {
                Scope = PlanDesignCanonicalScopes.CurrentPlan,
                ClientId = validationFact.ClientId == 0 ? clientId : validationFact.ClientId,
                Carrier = FirstNonBlank(validationFact.SourceCarrier, validationFact.Carrier, profile.Carrier),
                SourceCarrier = FirstNonBlank(validationFact.SourceCarrier, validationFact.Carrier, profile.Carrier),
                CanonicalBenefitCode = validationFact.CanonicalBenefitCode,
                NormalizedBenefitKey = validationFact.NormalizedBenefitKey,
                BenefitCategory = validationFact.BenefitCategory,
                BenefitName = validationFact.BenefitName,
                ClassName = validationFact.ClassName,
                FeatureName = validationFact.FeatureName,
                FeatureValue = validationFact.BookletValue,
                BookletValue = validationFact.BookletValue,
                ValidationStatus = validationFact.ValidationStatus,
                Status = validationFact.Status,
                ComparisonKey = validationFact.ComparisonKey,
                Availability = validationFact.Availability,
                SourceTable = validationFact.SourceTable,
                SourceId = validationFact.SourceId,
                SourceDocument = validationFact.SourceDocument,
                SourceDocumentType = validationFact.SourceDocumentType,
                SourceType = validationFact.SourceType,
                SourceFileName = validationFact.SourceFileName,
                TrackedSourceFileName = validationFact.TrackedSourceFileName,
                SourceLineageStatus = validationFact.SourceLineageStatus,
                SourceLineageMessage = validationFact.SourceLineageMessage,
                SourceFileHash = validationFact.SourceFileHash,
                SourceSheet = validationFact.SourceSheet,
                SourcePage = validationFact.SourcePage,
                SourceText = validationFact.BookletValue,
                ConfidenceScore = validationFact.ConfidenceScore,
                ExtractionMethod = "Booklet validation current-plan fallback",
                RawSourceReference = validationFact.RawSourceReference,
                DateLoaded = validationFact.DateLoaded
            };
            PlanDesignCanonicalMapper.CompleteFact(currentFact);
            profile.Facts.Add(currentFact);
        }
    }

    private static void AddCanonicalProfileIssues(PlanDesignCanonicalProfile profile)
    {
        if (!profile.HasCurrentPlan)
        {
            AddProfileIssue(
                profile,
                "current_plan_missing",
                PlanDesignCanonicalIssueSeverities.Error,
                "No current plan design facts were found.");
        }

        if (!profile.HasQuotePlan)
        {
            AddProfileIssue(
                profile,
                "quote_plan_missing",
                PlanDesignCanonicalIssueSeverities.Warning,
                "No quote plan design facts were found.");
        }

        if (!profile.HasProposalValidation)
        {
            AddProfileIssue(
                profile,
                "proposal_validation_missing",
                PlanDesignCanonicalIssueSeverities.Warning,
                "No proposal-to-booklet validation facts were found.");
        }

        foreach (var fact in profile.Facts
                     .Where(fact => string.IsNullOrWhiteSpace(fact.CanonicalBenefitCode))
                     .GroupBy(fact => fact.NormalizedBenefitKey, StringComparer.OrdinalIgnoreCase)
                     .Select(group => group.First()))
        {
            AddProfileIssue(
                profile,
                "benefit_not_mapped",
                PlanDesignCanonicalIssueSeverities.Warning,
                $"Benefit '{FirstNonBlank(fact.BenefitName, fact.BenefitCategory, fact.FeatureName)}' is not mapped to a canonical benefit code.",
                fact);
        }

        foreach (var fact in profile.ProposalValidationFacts.Where(PlanDesignCanonicalMapper.IsFailedValidation))
        {
            AddProfileIssue(
                profile,
                "proposal_plan_design_mismatch",
                PlanDesignCanonicalIssueSeverities.Warning,
                $"Proposal value '{fact.FeatureValue}' does not match booklet value '{fact.BookletValue}'.",
                fact);
        }

        foreach (var fact in profile.Facts
                     .Where(fact => fact.SourceLineageStatus == CanonicalSourceLineageStatuses.NeedsReview)
                     .GroupBy(
                         fact => $"{fact.SourceDocumentId}|{fact.TrackedSourceFileName}|{fact.SourceFileName}",
                         StringComparer.OrdinalIgnoreCase)
                     .Select(group => group.First()))
        {
            AddProfileIssue(
                profile,
                CanonicalSourceLineageValueGuardrailIssueCodes.SourceLineageConflict,
                PlanDesignCanonicalIssueSeverities.Warning,
                fact.SourceLineageMessage,
                fact);
        }

        foreach (var fact in profile.CurrentPlanFacts
                     .Where(fact => fact.SourceLineageStatus is CanonicalSourceLineageStatuses.Unlinked or CanonicalSourceLineageStatuses.Missing)
                     .GroupBy(
                         fact => $"{fact.TrackedSourceFileName}|{fact.SourceFileName}",
                         StringComparer.OrdinalIgnoreCase)
                     .Select(group => group.First()))
        {
            AddProfileIssue(
                profile,
                CanonicalSourceLineageValueGuardrailIssueCodes.SourceLineageUnlinked,
                PlanDesignCanonicalIssueSeverities.Warning,
                fact.SourceLineageMessage,
                fact);
        }

        foreach (var fact in profile.Facts
                     .Where(fact => fact.ValueGuardrailStatus == CanonicalValueGuardrailStatuses.Rejected)
                     .GroupBy(fact => $"{fact.Scope}|{fact.ComparisonKey}", StringComparer.OrdinalIgnoreCase)
                     .Select(group => group.First()))
        {
            AddProfileIssue(
                profile,
                CanonicalSourceLineageValueGuardrailIssueCodes.PercentageOutOfRange,
                PlanDesignCanonicalIssueSeverities.Warning,
                fact.ValueGuardrailMessage,
                fact);
        }

        foreach (var fact in profile.Facts
                     .Where(fact => fact.ValueGuardrailStatus == CanonicalValueGuardrailStatuses.PreservedForReview)
                     .GroupBy(fact => $"{fact.Scope}|{fact.ComparisonKey}", StringComparer.OrdinalIgnoreCase)
                     .Select(group => group.First()))
        {
            AddProfileIssue(
                profile,
                CanonicalSourceLineageValueGuardrailIssueCodes.MultiplePercentagesPreserved,
                PlanDesignCanonicalIssueSeverities.Warning,
                fact.ValueGuardrailMessage,
                fact);
        }
    }

    private static void AddProfileIssue(
        PlanDesignCanonicalProfile profile,
        string code,
        string severity,
        string message,
        PlanDesignCanonicalFact? fact = null)
    {
        profile.Issues.Add(new PlanDesignCanonicalIssue
        {
            Code = code,
            Severity = severity,
            Message = message,
            Benefit = fact is null ? "" : FirstNonBlank(fact.BenefitName, fact.BenefitCategory),
            Provision = fact?.FeatureName ?? "",
            SourceTable = fact?.SourceTable ?? "",
            SourceId = fact?.SourceId ?? ""
        });
    }
}
