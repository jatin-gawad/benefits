namespace RenewalReviewBuilder.App.Services;

public sealed record ClientRenewalSelectionIdentity(
    long? DatabaseClientId,
    string ClientNumber,
    string ClientName,
    string Carrier,
    DateOnly? RenewalDate);

public sealed record ClientRenewalSelectionCandidate(
    int Index,
    RenewalWorkflowIdentity Identity,
    DateTime DateLoaded);

public static class ClientRenewalSelectionResolver
{
    public static ClientRenewalSelectionCandidate? Select(
        ClientRenewalSelectionIdentity client,
        IEnumerable<ClientRenewalSelectionCandidate> candidates)
    {
        var available = candidates.ToList();
        if (available.Count == 0)
        {
            return null;
        }

        if (client.DatabaseClientId is > 0)
        {
            return SelectBest(
                client,
                available.Where(candidate => candidate.Identity.ClientId == client.DatabaseClientId.Value));
        }

        var policyKey = RevisedRateScheduleImporter.PolicyKey(client.ClientNumber);
        if (!string.IsNullOrWhiteSpace(policyKey))
        {
            var policyMatches = available.Where(candidate =>
                !string.IsNullOrWhiteSpace(candidate.Identity.PolicyNumber) &&
                RevisedRateScheduleImporter.PolicyKey(candidate.Identity.PolicyNumber)
                    .Equals(policyKey, StringComparison.Ordinal));
            var policyMatch = SelectBest(client, policyMatches);
            if (policyMatch is not null)
            {
                return policyMatch;
            }
        }

        var clientKey = NormalizeClientName(client.ClientName);
        if (string.IsNullOrWhiteSpace(clientKey))
        {
            return null;
        }

        return SelectBest(
            client,
            available.Where(candidate =>
                NormalizeClientName(candidate.Identity.ClientName)
                    .Equals(clientKey, StringComparison.Ordinal)));
    }

    public static bool Matches(
        ClientRenewalSelectionIdentity client,
        ClientRenewalSelectionCandidate candidate)
    {
        return Select(client, [candidate]) is not null;
    }

    private static ClientRenewalSelectionCandidate? SelectBest(
        ClientRenewalSelectionIdentity client,
        IEnumerable<ClientRenewalSelectionCandidate> candidates)
    {
        return candidates
            .OrderByDescending(candidate =>
                client.RenewalDate.HasValue && candidate.Identity.RenewalDate == client.RenewalDate)
            .ThenByDescending(candidate =>
                RevisedRateScheduleImporter.CarrierMatches(client.Carrier, candidate.Identity.Carrier))
            .ThenByDescending(candidate => candidate.Identity.RenewalDate?.DayNumber ?? int.MinValue)
            .ThenByDescending(candidate => candidate.DateLoaded)
            .FirstOrDefault();
    }

    private static string NormalizeClientName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        var suffixes = new HashSet<string>(StringComparer.Ordinal)
        {
            "THE",
            "LTD",
            "LIMITED",
            "INC",
            "INCORPORATED",
            "CORP",
            "CORPORATION",
            "CO",
            "COMPANY"
        };
        var words = new string(value
                .ToUpperInvariant()
                .Select(character => char.IsLetterOrDigit(character) ? character : ' ')
                .ToArray())
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(word => !suffixes.Contains(word));

        return string.Concat(words);
    }
}
