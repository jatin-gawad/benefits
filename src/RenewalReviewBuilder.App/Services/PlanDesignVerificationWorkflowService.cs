using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Uses the existing profile, validation persistence, and workspace data; it does not own a second model or store.</summary>
public sealed class PlanDesignVerificationWorkflowService
{
    private readonly PlanDesignCanonicalProfileService _profiles;
    private readonly DatabaseService _database;
    private readonly PlanDesignVerificationService _comparison = new();

    public PlanDesignVerificationWorkflowService(string? databasePath = null)
    {
        _profiles = new PlanDesignCanonicalProfileService(databasePath);
        _database = new DatabaseService(databasePath);
    }

    public async Task<IReadOnlyList<PlanDesignVerificationResult>> ValidateAsync(long clientId)
    {
        var results = _comparison.CompareProfile(await _profiles.BuildAsync(clientId));
        await _database.SavePlanDesignVerificationAsync(clientId, results);
        return results;
    }
}
