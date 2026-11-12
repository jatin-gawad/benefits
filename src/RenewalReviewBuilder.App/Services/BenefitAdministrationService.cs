using System.Text.Json;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Repositories;

namespace RenewalReviewBuilder.App.Services;

public sealed class BenefitAdministrationService
{
    private static readonly HashSet<string> ReviewStatuses =
    [
        BenefitAdministrationReviewStatuses.NeedsReview,
        BenefitAdministrationReviewStatuses.Confirmed,
        BenefitAdministrationReviewStatuses.Rejected
    ];

    private static readonly HashSet<string> ProviderRoles =
    [
        BenefitAdministrationProviderRoles.BenefitAdministrator,
        BenefitAdministrationProviderRoles.ClaimsAdministrator,
        BenefitAdministrationProviderRoles.DisabilityManagement,
        BenefitAdministrationProviderRoles.TravelAssistance,
        BenefitAdministrationProviderRoles.MedicalSecondOpinion,
        BenefitAdministrationProviderRoles.VirtualCare,
        BenefitAdministrationProviderRoles.OptionalBenefitAdministrator,
        BenefitAdministrationProviderRoles.Other
    ];

    private static readonly Dictionary<string, string> OrganizationAliases = BuildOrganizationAliases();
    private readonly string _databasePath;
    private readonly BenefitAdministrationRepository _repository;

    public BenefitAdministrationService(string? databasePath = null)
    {
        _databasePath = string.IsNullOrWhiteSpace(databasePath)
            ? AppPaths.BenefitsIntelligenceDatabasePath
            : databasePath;
        _repository = new BenefitAdministrationRepository(_databasePath);
    }

    public async Task<BenefitAdministrationImportResult> ImportAsync(BenefitAdministrationImportRequest request)
    {
        var initialization = new DatabaseService(_databasePath).Initialize();
        if (!initialization.Succeeded)
        {
            return new BenefitAdministrationImportResult { Message = initialization.Message };
        }

        await using var connection = new SqliteConnection($"Data Source={_databasePath}");
        await connection.OpenAsync();
        await using (var pragma = connection.CreateCommand())
        {
            pragma.CommandText = "PRAGMA foreign_keys = ON;";
            await pragma.ExecuteNonQueryAsync();
        }

        await using var transaction = connection.BeginTransaction();
        try
        {
            var result = await ImportInTransactionAsync(connection, transaction, request);
            if (!result.Succeeded)
            {
                await transaction.RollbackAsync();
                return result;
            }

            await transaction.CommitAsync();
            return result;
        }
        catch (Exception ex)
        {
            await transaction.RollbackAsync();
            AppLogger.Error(ex);
            return new BenefitAdministrationImportResult
            {
                Message = $"Benefit Administration import was rolled back: {ex.Message}"
            };
        }
    }

    internal async Task<BenefitAdministrationImportResult> ImportInTransactionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        BenefitAdministrationImportRequest request)
    {
        if (request.ClientId <= 0 || !await BenefitAdministrationRepository.ClientExistsAsync(connection, transaction, request.ClientId))
        {
            return new BenefitAdministrationImportResult { Message = "A confirmed client is required before Benefit Administration can be saved." };
        }

        var warnings = new List<string>();
        var inserted = 0;
        var providerAssignmentsInserted = 0;
        var duplicatesSkipped = 0;
        foreach (var source in request.Records)
        {
            if (!source.SourceDocumentId.HasValue ||
                !await BenefitAdministrationRepository.SourceDocumentBelongsToClientAsync(
                    connection,
                    transaction,
                    source.SourceDocumentId.Value,
                    request.ClientId))
            {
                warnings.Add($"{source.BenefitDisplayName}: source document was not confirmed for the selected client; row was not saved.");
                continue;
            }

            var benefitTypeCode = BenefitMatcher.ResolveCanonicalBenefitTypeCode(source.BenefitDisplayName);
            if (string.IsNullOrWhiteSpace(benefitTypeCode))
            {
                warnings.Add($"{source.BenefitDisplayName}: benefit type is unknown; row was not saved.");
                continue;
            }

            var benefitTypeId = await BenefitAdministrationRepository.ResolveBenefitTypeIdAsync(
                connection,
                transaction,
                benefitTypeCode);
            if (!benefitTypeId.HasValue)
            {
                throw new InvalidOperationException($"Canonical benefit type {benefitTypeCode} is missing from Benefit.");
            }

            var insurerId = await ResolveOrganizationIdAsync(
                connection,
                transaction,
                source.InsurerDisplayName,
                warnings,
                source.BenefitDisplayName);
            var providerOrganizations = new List<(long Id, string CanonicalName, string SourceName)>();
            foreach (var providerSourceName in source.ServiceProviderDisplayNames.Where(value => !string.IsNullOrWhiteSpace(value)))
            {
                var canonicalName = ResolveCanonicalOrganizationName(providerSourceName);
                if (string.IsNullOrWhiteSpace(canonicalName))
                {
                    warnings.Add($"{source.BenefitDisplayName}: unknown service provider '{providerSourceName}'.");
                    continue;
                }

                var providerId = await BenefitAdministrationRepository.ResolveOrganizationIdAsync(
                    connection,
                    transaction,
                    NormalizeOrganizationName(canonicalName));
                if (!providerId.HasValue)
                {
                    warnings.Add($"{source.BenefitDisplayName}: service provider '{providerSourceName}' could not be resolved.");
                    continue;
                }

                providerOrganizations.Add((providerId.Value, canonicalName, providerSourceName));
            }

            if (!insurerId.HasValue && providerOrganizations.Count == 0 && string.IsNullOrWhiteSpace(source.PolicyNumber))
            {
                warnings.Add($"{source.BenefitDisplayName}: insurer, service provider, and policy number are all missing; row was not saved.");
                continue;
            }

            AddCompletenessWarnings(source, warnings);
            if (HasUnresolvedOcrPolicyCharacters(source.PolicyNumber))
            {
                warnings.Add($"{source.BenefitDisplayName}: policy number contains unresolved OCR characters.");
            }

            var sourceLineage = BuildSourceLineageJson(
                source,
                insurerId.HasValue,
                providerOrganizations.Select(provider => provider.SourceName));

            var existing = await BenefitAdministrationRepository.FindActiveIdentityAsync(
                connection,
                transaction,
                request.ClientId,
                request.PlanDesignVersionId,
                benefitTypeId.Value,
                source.ProductName);
            long administrationId;
            if (existing is not null &&
                existing.InsurerOrganizationId == insurerId &&
                string.Equals(existing.PolicyNumber, source.PolicyNumber, StringComparison.OrdinalIgnoreCase))
            {
                administrationId = existing.ClientBenefitAdministrationId;
                await BenefitAdministrationRepository.RefreshExistingSourceAsync(
                    connection,
                    transaction,
                    administrationId,
                    source,
                    sourceLineage);
                duplicatesSkipped++;
            }
            else
            {
                if (existing is not null)
                {
                    if (string.Equals(existing.PolicyNumber, source.PolicyNumber, StringComparison.OrdinalIgnoreCase) &&
                        existing.InsurerOrganizationId != insurerId)
                    {
                        warnings.Add($"{source.BenefitDisplayName}: the same benefit and policy number now have a different insurer.");
                    }

                    var terminationDate = (request.EffectiveDate ?? DateTime.UtcNow.Date).AddDays(-1);
                    await BenefitAdministrationRepository.EndDateAdministrationAsync(
                        connection,
                        transaction,
                        existing.ClientBenefitAdministrationId,
                        terminationDate);
                }

                var record = new ClientBenefitAdministrationRecord
                {
                    ClientId = request.ClientId,
                    PlanDesignVersionId = request.PlanDesignVersionId,
                    BenefitTypeId = benefitTypeId.Value,
                    BenefitTypeCode = benefitTypeCode,
                    BenefitDisplayName = source.BenefitDisplayName.Trim(),
                    ProductName = source.ProductName.Trim(),
                    PolicyNumber = source.PolicyNumber.Trim(),
                    InsurerOrganizationId = insurerId,
                    IsOptional = source.IsOptional,
                    EffectiveDate = request.EffectiveDate,
                    SourceDocumentId = source.SourceDocumentId,
                    SourcePageNumber = source.SourcePageNumber,
                    SourceText = sourceLineage,
                    ExtractionConfidence = source.ExtractionConfidence,
                    ReviewStatusCode = BenefitAdministrationReviewStatuses.NeedsReview,
                    IsActive = true
                };
                administrationId = await BenefitAdministrationRepository.InsertAdministrationAsync(
                    connection,
                    transaction,
                    record,
                    source.InsurerDisplayName);
                inserted++;
            }

            var existingProviderKeys = await BenefitAdministrationRepository.GetActiveProviderKeysAsync(
                connection,
                transaction,
                administrationId);
            foreach (var provider in providerOrganizations)
            {
                var role = ResolveProviderRole(provider.CanonicalName, benefitTypeCode, source.IsOptional);
                if (existingProviderKeys.Any(key => key.ProviderId == provider.Id && !key.Role.Equals(role, StringComparison.OrdinalIgnoreCase)))
                {
                    warnings.Add($"{source.BenefitDisplayName}: {provider.CanonicalName} is assigned under conflicting provider roles.");
                }

                if (existingProviderKeys.Contains((provider.Id, role)))
                {
                    continue;
                }

                await BenefitAdministrationRepository.InsertProviderAssignmentAsync(
                    connection,
                    transaction,
                    new BenefitServiceProviderAssignmentRecord
                    {
                        ClientBenefitAdministrationId = administrationId,
                        ProviderOrganizationId = provider.Id,
                        ProviderRoleCode = role,
                        SourceDisplayName = provider.SourceName,
                        IsPrimary = true,
                        EffectiveDate = request.EffectiveDate,
                        SourceDocumentId = source.SourceDocumentId,
                        SourcePageNumber = source.SourcePageNumber,
                        ReviewStatusCode = BenefitAdministrationReviewStatuses.NeedsReview,
                        IsActive = true
                    });
                existingProviderKeys.Add((provider.Id, role));
                providerAssignmentsInserted++;
            }
        }

        return new BenefitAdministrationImportResult
        {
            Succeeded = true,
            RecordsInserted = inserted,
            ProviderAssignmentsInserted = providerAssignmentsInserted,
            DuplicatesSkipped = duplicatesSkipped,
            Warnings = warnings.Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
            Message = $"Saved {inserted} Benefit Administration row(s) and {providerAssignmentsInserted} provider assignment(s); {duplicatesSkipped} existing row(s) were refreshed."
        };
    }

    public async Task<BenefitAdministrationWorkspaceViewModel> GetWorkspaceAsync(
        long clientId,
        long? planDesignVersionId = null,
        DateTime? asOfDate = null)
    {
        var records = planDesignVersionId.HasValue
            ? await _repository.GetByPlanDesignVersionIdAsync(planDesignVersionId.Value)
            : await _repository.GetActiveAsOfAsync(clientId, asOfDate ?? DateTime.Today);
        var workspace = new BenefitAdministrationWorkspaceViewModel();
        foreach (var record in records.Where(record => record.ClientId == clientId && record.IsActive))
        {
            workspace.Rows.Add(new BenefitAdministrationWorkspaceRow
            {
                ClientBenefitAdministrationId = record.ClientBenefitAdministrationId,
                Benefit = record.BenefitDisplayName,
                Product = record.ProductName,
                InsuranceCompany = string.IsNullOrWhiteSpace(record.InsurerName) ? "Not assigned" : record.InsurerName,
                ServiceProvider = record.ServiceProviders.Count == 0
                    ? "Not assigned"
                    : string.Join(Environment.NewLine, record.ServiceProviders.Where(item => item.IsActive).Select(item => item.ProviderName)),
                ProviderRole = record.ServiceProviders.Count == 0
                    ? ""
                    : string.Join(Environment.NewLine, record.ServiceProviders.Where(item => item.IsActive).Select(item => FormatCode(item.ProviderRoleCode))),
                PolicyNumber = string.IsNullOrWhiteSpace(record.PolicyNumber) ? "Not saved" : record.PolicyNumber,
                EffectiveDate = record.EffectiveDate?.ToString("MMM d, yyyy") ?? "Not saved",
                Status = record.ReviewStatusCode,
                Source = BuildSourceLabel(record),
                SourcePath = record.SourceDocumentPath,
                SourceDocumentId = record.SourceDocumentId,
                SourcePageNumber = record.SourcePageNumber,
                IsOptional = record.IsOptional,
                ProviderAssignments = record.ServiceProviders.ToList()
            });
        }

        AddDistinctFilters(workspace.InsurerFilters, workspace.Rows.Select(row => row.InsuranceCompany));
        AddDistinctFilters(workspace.ProviderFilters, records.SelectMany(record => record.ServiceProviders).Select(item => item.ProviderName));
        AddDistinctFilters(workspace.StatusFilters, workspace.Rows.Select(row => row.Status));
        return workspace;
    }

    public async Task<IReadOnlyList<BenefitAdministrationClientExportRow>> GetClientFacingExportRowsAsync(
        long clientId,
        long? planDesignVersionId = null)
    {
        var workspace = await GetWorkspaceAsync(clientId, planDesignVersionId);
        return workspace.Rows.Select(row => new BenefitAdministrationClientExportRow
        {
            Benefit = row.Benefit,
            InsuranceCompany = row.InsuranceCompany,
            ServiceProvider = row.ServiceProvider.Replace(Environment.NewLine, "; "),
            PolicyNumber = row.PolicyNumber
        }).ToList();
    }

    public async Task<BenefitAdministrationImportResult> CarryForwardAsync(
        long clientId,
        long? sourcePlanDesignVersionId,
        long targetPlanDesignVersionId,
        DateTime? effectiveDate = null)
    {
        if (clientId <= 0 || targetPlanDesignVersionId <= 0)
        {
            return new BenefitAdministrationImportResult
            {
                Message = "A confirmed client and target plan-design version are required."
            };
        }

        var sourceRecords = sourcePlanDesignVersionId.HasValue
            ? await _repository.GetByPlanDesignVersionIdAsync(sourcePlanDesignVersionId.Value)
            : await _repository.GetActiveAsOfAsync(clientId, effectiveDate ?? DateTime.Today);
        sourceRecords = sourceRecords
            .Where(record => record.ClientId == clientId && record.IsActive)
            .ToList();

        var initialization = new DatabaseService(_databasePath).Initialize();
        if (!initialization.Succeeded)
        {
            return new BenefitAdministrationImportResult { Message = initialization.Message };
        }

        await using var connection = new SqliteConnection($"Data Source={_databasePath}");
        await connection.OpenAsync();
        await using (var pragma = connection.CreateCommand())
        {
            pragma.CommandText = "PRAGMA foreign_keys = ON;";
            await pragma.ExecuteNonQueryAsync();
        }

        await using var transaction = connection.BeginTransaction();
        try
        {
            var inserted = 0;
            var providerAssignmentsInserted = 0;
            var duplicatesSkipped = 0;
            foreach (var source in sourceRecords)
            {
                var duplicate = await BenefitAdministrationRepository.FindActiveIdentityAsync(
                    connection,
                    transaction,
                    clientId,
                    targetPlanDesignVersionId,
                    source.BenefitTypeId,
                    source.ProductName);
                if (duplicate is not null)
                {
                    duplicatesSkipped++;
                    continue;
                }

                var copiedRecord = new ClientBenefitAdministrationRecord
                {
                    ClientId = clientId,
                    PlanDesignVersionId = targetPlanDesignVersionId,
                    BenefitTypeId = source.BenefitTypeId,
                    BenefitTypeCode = source.BenefitTypeCode,
                    BenefitDisplayName = source.BenefitDisplayName,
                    ProductName = source.ProductName,
                    PolicyNumber = source.PolicyNumber,
                    InsurerOrganizationId = source.InsurerOrganizationId,
                    IsOptional = source.IsOptional,
                    EffectiveDate = effectiveDate,
                    SourceDocumentId = source.SourceDocumentId,
                    SourcePageNumber = source.SourcePageNumber,
                    SourceText = source.SourceText,
                    ExtractionConfidence = source.ExtractionConfidence,
                    ReviewStatusCode = BenefitAdministrationReviewStatuses.NeedsReview,
                    IsActive = true,
                    CopiedFromClientBenefitAdministrationId = source.ClientBenefitAdministrationId
                };
                var administrationId = await BenefitAdministrationRepository.InsertAdministrationAsync(
                    connection,
                    transaction,
                    copiedRecord,
                    source.InsurerName);
                inserted++;

                foreach (var provider in source.ServiceProviders.Where(item => item.IsActive))
                {
                    await BenefitAdministrationRepository.InsertProviderAssignmentAsync(
                        connection,
                        transaction,
                        new BenefitServiceProviderAssignmentRecord
                        {
                            ClientBenefitAdministrationId = administrationId,
                            ProviderOrganizationId = provider.ProviderOrganizationId,
                            ProviderRoleCode = provider.ProviderRoleCode,
                            SourceDisplayName = provider.SourceDisplayName,
                            IsPrimary = provider.IsPrimary,
                            EffectiveDate = effectiveDate,
                            SourceDocumentId = provider.SourceDocumentId,
                            SourcePageNumber = provider.SourcePageNumber,
                            ReviewStatusCode = BenefitAdministrationReviewStatuses.NeedsReview,
                            IsActive = true
                        });
                    providerAssignmentsInserted++;
                }
            }

            await transaction.CommitAsync();
            return new BenefitAdministrationImportResult
            {
                Succeeded = true,
                RecordsInserted = inserted,
                ProviderAssignmentsInserted = providerAssignmentsInserted,
                DuplicatesSkipped = duplicatesSkipped,
                Warnings = sourceRecords.Count == 0
                    ? ["No active Benefit Administration assignments were available to carry forward."]
                    : [],
                Message = $"Carried forward {inserted} Benefit Administration row(s) for review."
            };
        }
        catch (Exception ex)
        {
            await transaction.RollbackAsync();
            AppLogger.Error(ex);
            return new BenefitAdministrationImportResult
            {
                Message = $"Benefit Administration carry-forward was rolled back: {ex.Message}"
            };
        }
    }

    public async Task<bool> ConfirmAsync(long administrationId)
    {
        var record = await _repository.GetByIdAsync(administrationId)
                     ?? throw new InvalidOperationException("Benefit Administration record was not found.");
        if (string.IsNullOrWhiteSpace(record.BenefitTypeCode))
        {
            throw new InvalidOperationException("The benefit type must be resolved before confirmation.");
        }

        if (HasUnresolvedOcrPolicyCharacters(record.PolicyNumber))
        {
            throw new InvalidOperationException("The policy number contains unresolved OCR characters.");
        }

        if (HasUnresolvedOrganizationMatches(record.SourceText))
        {
            throw new InvalidOperationException("Every insurance company and service provider named by the source must be resolved before confirmation.");
        }

        return await _repository.UpdateReviewStatusAsync(administrationId, BenefitAdministrationReviewStatuses.Confirmed);
    }

    public Task<bool> UpdateReviewStatusAsync(long administrationId, string reviewStatusCode)
    {
        if (!ReviewStatuses.Contains(reviewStatusCode))
        {
            throw new ArgumentException("Review status is not supported.", nameof(reviewStatusCode));
        }

        return _repository.UpdateReviewStatusAsync(administrationId, reviewStatusCode);
    }

    public Task<bool> CorrectPolicyNumberAsync(long administrationId, string policyNumber)
    {
        return _repository.CorrectPolicyNumberAsync(administrationId, policyNumber);
    }

    public async Task<bool> CorrectInsurerAsync(long administrationId, long? insurerOrganizationId)
    {
        var organizations = await _repository.GetOrganizationsAsync();
        var insurer = insurerOrganizationId.HasValue
            ? organizations.FirstOrDefault(item => item.OrganizationId == insurerOrganizationId.Value)
            : null;
        if (insurerOrganizationId.HasValue && insurer is null)
        {
            throw new InvalidOperationException("The selected insurance company was not found.");
        }

        return await _repository.CorrectInsurerAsync(administrationId, insurerOrganizationId, insurer?.DisplayName ?? "");
    }

    public async Task<long> AddProviderAssignmentAsync(BenefitServiceProviderAssignmentRecord assignment)
    {
        if (!ProviderRoles.Contains(assignment.ProviderRoleCode))
        {
            throw new ArgumentException("Provider role is not supported.", nameof(assignment));
        }

        var organizations = await _repository.GetOrganizationsAsync();
        if (organizations.All(item => item.OrganizationId != assignment.ProviderOrganizationId))
        {
            throw new InvalidOperationException("The selected service provider was not found.");
        }

        assignment.ReviewStatusCode = BenefitAdministrationReviewStatuses.NeedsReview;
        assignment.IsActive = true;
        return await _repository.AddProviderAssignmentAsync(assignment);
    }

    public async Task<bool> UpdateProviderAssignmentAsync(BenefitServiceProviderAssignmentRecord assignment)
    {
        if (!ProviderRoles.Contains(assignment.ProviderRoleCode))
        {
            throw new ArgumentException("Provider role is not supported.", nameof(assignment));
        }

        var organizations = await _repository.GetOrganizationsAsync();
        if (organizations.All(item => item.OrganizationId != assignment.ProviderOrganizationId))
        {
            throw new InvalidOperationException("The selected service provider was not found.");
        }

        return await _repository.UpdateProviderAssignmentAsync(assignment);
    }

    public Task<bool> RemoveProviderAssignmentAsync(long assignmentId) =>
        _repository.DeactivateProviderAssignmentAsync(assignmentId);

    public Task<bool> EndDateProviderAssignmentAsync(long assignmentId, DateTime terminationDate) =>
        _repository.EndDateProviderAssignmentAsync(assignmentId, terminationDate);

    public Task<List<OrganizationRecord>> GetOrganizationsAsync() => _repository.GetOrganizationsAsync();

    private static async Task<long?> ResolveOrganizationIdAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string sourceName,
        ICollection<string> warnings,
        string benefitName)
    {
        if (string.IsNullOrWhiteSpace(sourceName))
        {
            return null;
        }

        var canonicalName = ResolveCanonicalOrganizationName(sourceName);
        if (string.IsNullOrWhiteSpace(canonicalName))
        {
            warnings.Add($"{benefitName}: unknown insurance company '{sourceName}'.");
            return null;
        }

        var id = await BenefitAdministrationRepository.ResolveOrganizationIdAsync(
            connection,
            transaction,
            NormalizeOrganizationName(canonicalName));
        if (!id.HasValue)
        {
            warnings.Add($"{benefitName}: insurance company '{sourceName}' could not be resolved.");
        }

        return id;
    }

    private static void AddCompletenessWarnings(
        BenefitAdministrationExtractionCandidate source,
        ICollection<string> warnings)
    {
        if (string.IsNullOrWhiteSpace(source.InsurerDisplayName))
        {
            warnings.Add($"{source.BenefitDisplayName}: missing insurer.");
        }

        if (source.ServiceProviderDisplayNames.All(string.IsNullOrWhiteSpace))
        {
            warnings.Add($"{source.BenefitDisplayName}: missing service provider.");
        }

        if (string.IsNullOrWhiteSpace(source.PolicyNumber))
        {
            warnings.Add($"{source.BenefitDisplayName}: missing policy number.");
        }
    }

    private static string BuildSourceLineageJson(
        BenefitAdministrationExtractionCandidate source,
        bool insurerResolved,
        IEnumerable<string> resolvedProviderSourceNames)
    {
        var resolvedProviders = resolvedProviderSourceNames
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        return JsonSerializer.Serialize(new
        {
            originalBenefitText = source.BenefitDisplayName,
            originalInsurerText = source.InsurerDisplayName,
            originalServiceProviderTexts = source.ServiceProviderDisplayNames,
            originalPolicyNumberText = source.PolicyNumber,
            extractedSourceText = source.SourceText,
            unresolvedInsurerText = !string.IsNullOrWhiteSpace(source.InsurerDisplayName) && !insurerResolved
                ? source.InsurerDisplayName
                : "",
            unresolvedServiceProviderTexts = source.ServiceProviderDisplayNames
                .Where(value => !string.IsNullOrWhiteSpace(value) && !resolvedProviders.Contains(value))
                .ToList()
        });
    }

    private static bool HasUnresolvedOrganizationMatches(string sourceLineage)
    {
        if (string.IsNullOrWhiteSpace(sourceLineage))
        {
            return false;
        }

        try
        {
            using var document = JsonDocument.Parse(sourceLineage);
            var root = document.RootElement;
            if (root.TryGetProperty("unresolvedInsurerText", out var insurer) &&
                !string.IsNullOrWhiteSpace(insurer.GetString()))
            {
                return true;
            }

            return root.TryGetProperty("unresolvedServiceProviderTexts", out var providers) &&
                   providers.ValueKind == JsonValueKind.Array &&
                   providers.EnumerateArray().Any(value => !string.IsNullOrWhiteSpace(value.GetString()));
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static string ResolveProviderRole(
        string canonicalProviderName,
        string benefitTypeCode,
        bool isOptional)
    {
        if (benefitTypeCode.Equals(CanonicalBenefitTypeCodes.MedicalSecondOpinion, StringComparison.OrdinalIgnoreCase))
        {
            return BenefitAdministrationProviderRoles.MedicalSecondOpinion;
        }

        if (benefitTypeCode.Equals(CanonicalBenefitTypeCodes.VirtualHealthcare, StringComparison.OrdinalIgnoreCase))
        {
            return BenefitAdministrationProviderRoles.VirtualCare;
        }

        if (canonicalProviderName.Equals("Disability Management Institute", StringComparison.OrdinalIgnoreCase))
        {
            return isOptional
                ? BenefitAdministrationProviderRoles.OptionalBenefitAdministrator
                : BenefitAdministrationProviderRoles.DisabilityManagement;
        }

        if (canonicalProviderName.Equals("GroupHEALTH", StringComparison.OrdinalIgnoreCase))
        {
            return BenefitAdministrationProviderRoles.BenefitAdministrator;
        }

        if (canonicalProviderName.Equals("Trident Global Assistance", StringComparison.OrdinalIgnoreCase))
        {
            return BenefitAdministrationProviderRoles.TravelAssistance;
        }

        return BenefitAdministrationProviderRoles.Other;
    }

    private static string ResolveCanonicalOrganizationName(string sourceName)
    {
        return OrganizationAliases.TryGetValue(NormalizeOrganizationName(sourceName), out var canonicalName)
            ? canonicalName
            : "";
    }

    private static Dictionary<string, string> BuildOrganizationAliases()
    {
        var aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        AddAliases(aliases, "SSQ, Life Insurance Company Inc.",
            "SSQ Life Insurance Company Inc.", "SSQ, Life Insurance Company Inc.", "SSQ Insurance");
        AddAliases(aliases, "Industrial Alliance",
            "Industrial Alliance", "iA Financial Group", "IA Financial Group");
        AddAliases(aliases, "Disability Management Institute",
            "Disability Management Institute", "DMI", "Disability Management Institute (DMI)");
        AddAliases(aliases, "GroupHEALTH",
            "GroupHEALTH", "Group Health", "GroupHEALTH Benefit Solutions");
        AddAliases(aliases, "Trident Global Assistance", "Trident Global Assistance");
        AddAliases(aliases, "WorldCare", "WorldCare", "World Care");
        AddAliases(aliases, "Advica Health", "Advica Health", "Advica");
        AddAliases(aliases, "Teladoc Health", "Teladoc Health", "Teladoc", "Teladoc Medical Experts");
        AddAliases(aliases, "Dialogue",
            "Dialogue", "Dialogue Health", "Dialogue Health Technologies Inc.", "Dialogue Medical Second Opinion");
        AddAliases(aliases, "Sparrow", "Sparrow");
        AddAliases(aliases, "TELUS Health Virtual Care", "TELUS Health Virtual Care");
        AddAliases(aliases, "The Co-operators", "The Co-operators", "Co-operators");
        return aliases;
    }

    private static void AddAliases(IDictionary<string, string> aliases, string canonicalName, params string[] values)
    {
        aliases[NormalizeOrganizationName(canonicalName)] = canonicalName;
        foreach (var value in values)
        {
            aliases[NormalizeOrganizationName(value)] = canonicalName;
        }
    }

    private static string NormalizeOrganizationName(string value)
    {
        return new string((value ?? "").Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
    }

    private static bool HasUnresolvedOcrPolicyCharacters(string value)
    {
        return !string.IsNullOrWhiteSpace(value) && value.Any(character => character is '?' or '\uFFFD');
    }

    private static string BuildSourceLabel(ClientBenefitAdministrationRecord record)
    {
        var source = string.IsNullOrWhiteSpace(record.SourceDocumentName) ? "Source document" : record.SourceDocumentName;
        return record.SourcePageNumber.HasValue ? $"{source}, page {record.SourcePageNumber}" : source;
    }

    private static string FormatCode(string value)
    {
        return string.Join(" ", value.Split('_', StringSplitOptions.RemoveEmptyEntries)
            .Select(word => char.ToUpperInvariant(word[0]) + word[1..].ToLowerInvariant()));
    }

    private static void AddDistinctFilters(ICollection<string> target, IEnumerable<string> values)
    {
        foreach (var value in values.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value))
        {
            target.Add(value);
        }
    }
}
