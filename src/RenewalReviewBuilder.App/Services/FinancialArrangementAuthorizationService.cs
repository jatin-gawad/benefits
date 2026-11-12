namespace RenewalReviewBuilder.App.Services;

public static class FinancialArrangementPermissions
{
    public const string FinancialArrangementValidate = "FinancialArrangement.Validate";
    public const string FinancialArrangementConfirm = "FinancialArrangement.Confirm";
    public const string CommissionConfirm = "Commission.Confirm";
    public const string PoolingConfirm = "Pooling.Confirm";
    public const string RenewalFinancialComparisonRefresh = "RenewalFinancialComparison.Refresh";
    public const string RenewalFinancialComparisonFinalize = "RenewalFinancialComparison.Finalize";
}

public sealed class FinancialArrangementAuthorizationContext
{
    public string User { get; init; } = "";
    public ISet<string> GrantedPermissions { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    public static FinancialArrangementAuthorizationContext Allow(string user, params string[] permissions)
    {
        var context = new FinancialArrangementAuthorizationContext { User = user };
        foreach (var permission in permissions) context.GrantedPermissions.Add(permission);
        return context;
    }
}

public interface IFinancialArrangementAuthorizationService
{
    void Demand(FinancialArrangementAuthorizationContext context, string permission);
}

public sealed class FinancialArrangementAuthorizationService : IFinancialArrangementAuthorizationService
{
    public void Demand(FinancialArrangementAuthorizationContext context, string permission)
    {
        ArgumentNullException.ThrowIfNull(context);
        if (string.IsNullOrWhiteSpace(context.User) || !context.GrantedPermissions.Contains(permission)) throw new UnauthorizedAccessException($"{context.User ?? "Unknown user"} does not have {permission} permission.");
    }
}
