using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Repositories;

public interface IFlexibleBenefitRepository
{
    Task<FlexibleBenefitImportResult> SaveAsync(
        FlexibleBenefitImportRequest request,
        FlexibleBenefitExtractionResult extraction,
        FlexibleBenefitValidationResult validation);

    Task<FlexibleBenefitWorkspaceViewModel> GetWorkspaceAsync(long clientId);
    Task<DatabaseOperationResult> MarkReviewedAsync(long planId, string reviewedBy);
    Task<DatabaseOperationResult> CorrectSourceFieldAsync(
        long sourceFieldId,
        string newNormalizedValue,
        string correctedBy,
        string reason);
    Task<IReadOnlyList<FlexibleBenefitVersionDifferenceRow>> CompareVersionsAsync(long currentPlanId, long priorPlanId);
}

public sealed class FlexibleBenefitRepository : IFlexibleBenefitRepository
{
    private readonly DatabaseService _database;

    public FlexibleBenefitRepository(string? databasePath = null)
    {
        _database = new DatabaseService(databasePath);
    }

    public Task<FlexibleBenefitImportResult> SaveAsync(
        FlexibleBenefitImportRequest request,
        FlexibleBenefitExtractionResult extraction,
        FlexibleBenefitValidationResult validation)
    {
        return _database.SaveFlexibleBenefitExtractionAsync(request, extraction, validation);
    }

    public Task<FlexibleBenefitWorkspaceViewModel> GetWorkspaceAsync(long clientId)
    {
        return _database.GetFlexibleBenefitWorkspaceAsync(clientId);
    }

    public Task<DatabaseOperationResult> MarkReviewedAsync(long planId, string reviewedBy)
    {
        return _database.MarkFlexibleBenefitPlanReviewedAsync(planId, reviewedBy);
    }

    public Task<DatabaseOperationResult> CorrectSourceFieldAsync(
        long sourceFieldId,
        string newNormalizedValue,
        string correctedBy,
        string reason)
    {
        return _database.CorrectFlexibleBenefitSourceFieldAsync(
            sourceFieldId,
            newNormalizedValue,
            correctedBy,
            reason);
    }

    public Task<IReadOnlyList<FlexibleBenefitVersionDifferenceRow>> CompareVersionsAsync(long currentPlanId, long priorPlanId)
    {
        return _database.CompareFlexibleBenefitVersionsAsync(currentPlanId, priorPlanId);
    }
}
