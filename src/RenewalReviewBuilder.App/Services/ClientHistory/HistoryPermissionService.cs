using RenewalReviewBuilder.App.Models.History;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

/// <summary>Conservative desktop policy until user/role authorization is introduced.</summary>
public sealed class HistoryPermissionService : IHistoryPermissionService
{
    public bool CanReview(ClientHistoryEventReadModel historyEvent) =>
        historyEvent.ValidationStatus == HistoryValidationStatus.NeedsReview;
}