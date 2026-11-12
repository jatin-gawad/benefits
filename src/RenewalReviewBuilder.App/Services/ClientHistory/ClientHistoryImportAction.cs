using System.Threading;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

/// <summary>
/// Coordinates the existing modal Client History importer with a canonical projection reload.
/// The caller supplies the persisted SQLite client key unchanged.
/// </summary>
public sealed class ClientHistoryImportAction
{
    public async Task RunAsync(
        long databaseClientId,
        Action<long> showImportDialog,
        Func<CancellationToken, Task> reloadProjectionAsync,
        CancellationToken cancellationToken = default)
    {
        if (databaseClientId <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(databaseClientId),
                "A persisted SQLite client ID is required to import Client History.");
        }

        ArgumentNullException.ThrowIfNull(showImportDialog);
        ArgumentNullException.ThrowIfNull(reloadProjectionAsync);

        cancellationToken.ThrowIfCancellationRequested();
        showImportDialog(databaseClientId);
        cancellationToken.ThrowIfCancellationRequested();
        await reloadProjectionAsync(cancellationToken);
    }
}
