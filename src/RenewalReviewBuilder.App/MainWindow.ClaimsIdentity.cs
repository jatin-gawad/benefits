using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App;

public partial class MainWindow
{
    private static void ValidateClaimsSourceIdentityForWorkspace(
        ClientWorkspaceViewModel workspace,
        ClaimsExperienceDataModelResult export)
    {
        if (!string.IsNullOrWhiteSpace(export.ClientName) &&
            !string.IsNullOrWhiteSpace(workspace.ClientName) &&
            !BookletClientIdentityGuardrail.ClientNamesLikelyMatch(workspace.ClientName, export.ClientName))
        {
            throw new InvalidDataException(
                $"The claims reports are for '{export.ClientName}', not the selected client '{workspace.ClientName}'. No claims were saved.");
        }

        var sourcePolicy = NormalizeClaimsWorkspaceIdentifier(export.ClientNumber);
        var workspacePolicy = NormalizeClaimsWorkspaceIdentifier(workspace.ClientId);
        if (sourcePolicy.Length > 0 &&
            workspacePolicy.Length > 0 &&
            !sourcePolicy.Equals(workspacePolicy, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidDataException(
                $"The claims report contract/policy '{export.ClientNumber}' does not match selected client policy '{workspace.ClientId}'. No claims were saved.");
        }
    }

    private static string NormalizeClaimsWorkspaceIdentifier(string value)
    {
        return new string((value ?? "")
            .Where(char.IsLetterOrDigit)
            .Select(char.ToUpperInvariant)
            .ToArray());
    }
}
