using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class PlanDesignCanonicalProfileService
{
    private readonly PlanDesignIntelligenceService _readOnlySource;

    public PlanDesignCanonicalProfileService(string? databasePath = null)
    {
        _readOnlySource = new PlanDesignIntelligenceService(databasePath);
    }

    public string DatabasePath => _readOnlySource.DatabasePath;

    public Task<PlanDesignCanonicalProfile> BuildAsync(long clientId)
    {
        return _readOnlySource.BuildCanonicalProfileAsync(clientId);
    }
}
