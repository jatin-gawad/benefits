namespace RenewalReviewBuilder.App.Services;

public sealed record RenewalWorkflowIdentity(
    long ClientId,
    long? SourceDocumentId,
    string ClientName,
    string PolicyNumber,
    string Carrier,
    DateOnly? RenewalDate);

public static class RenewalWorkflowContextGuard
{
    public static bool Matches(RenewalWorkflowIdentity selected, RenewalWorkflowIdentity current)
    {
        var hasStrongMatch = false;

        if (selected.ClientId > 0 && current.ClientId > 0)
        {
            if (selected.ClientId != current.ClientId)
            {
                return false;
            }

            hasStrongMatch = true;
        }

        if (selected.SourceDocumentId.HasValue && current.SourceDocumentId.HasValue)
        {
            if (selected.SourceDocumentId != current.SourceDocumentId)
            {
                return false;
            }

            hasStrongMatch = true;
        }

        if (!string.IsNullOrWhiteSpace(selected.PolicyNumber) &&
            !string.IsNullOrWhiteSpace(current.PolicyNumber))
        {
            if (!RevisedRateScheduleImporter.PolicyKey(selected.PolicyNumber)
                    .Equals(RevisedRateScheduleImporter.PolicyKey(current.PolicyNumber), StringComparison.Ordinal))
            {
                return false;
            }

            hasStrongMatch = true;
        }

        if (!string.IsNullOrWhiteSpace(selected.ClientName) &&
            !string.IsNullOrWhiteSpace(current.ClientName))
        {
            if (!RevisedRateScheduleImporter.ClientMatches(selected.ClientName, current.ClientName))
            {
                return false;
            }

            hasStrongMatch = true;
        }

        if (!string.IsNullOrWhiteSpace(selected.Carrier) &&
            !string.IsNullOrWhiteSpace(current.Carrier) &&
            !RevisedRateScheduleImporter.CarrierMatches(selected.Carrier, current.Carrier))
        {
            return false;
        }

        if (selected.RenewalDate.HasValue &&
            current.RenewalDate.HasValue &&
            selected.RenewalDate != current.RenewalDate)
        {
            return false;
        }

        return hasStrongMatch;
    }
}
