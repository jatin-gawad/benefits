namespace RenewalReviewBuilder.App.Models;

/// <summary>
/// A verified policy attached to a client group. Clients remain the canonical
/// group record; policies identify the insured and ASO arrangements that can
/// contribute source data to that group.
/// </summary>
public sealed class ClientPolicyRecord
{
    public long ClientPolicyId { get; init; }
    public long ClientId { get; init; }
    public string PolicyNumber { get; init; } = "";
    public string PolicyRole { get; init; } = "";
    public string Carrier { get; init; } = "";
    public string VerificationStatus { get; init; } = "";
    public string SourceFileName { get; init; } = "";
    public string SourceDescription { get; init; } = "";
    public string ConfirmedBy { get; init; } = "";
    public DateTime ConfirmedAt { get; init; }
    public DateTime CreatedAt { get; init; }
    public DateTime UpdatedAt { get; init; }
}

public static class ClientPolicyRoles
{
    public const string Primary = "Primary";
    public const string Aso = "ASO";
}

public static class ClientPolicyVerificationStatuses
{
    public const string Verified = "Verified";
}
