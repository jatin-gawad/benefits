using System.Threading;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class BenefitsMasterDatabaseExportService
{
    private readonly DatabaseService _databaseService;
    private readonly PlanDesignWorkspaceDataService _workspaceDataService;

    public BenefitsMasterDatabaseExportService(string? databasePath = null)
        : this(new DatabaseService(databasePath))
    {
    }

    internal BenefitsMasterDatabaseExportService(DatabaseService databaseService)
    {
        _databaseService = databaseService;
        _workspaceDataService = new PlanDesignWorkspaceDataService(databaseService.DatabasePath);
    }

    public async Task<BenefitsMasterDatabaseWorkbookResult> RebuildAsync(
        string outputFolder,
        IReadOnlyCollection<long>? clientIds = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(outputFolder))
        {
            throw new ArgumentException("Choose an output folder before rebuilding the Benefits Master Database.", nameof(outputFolder));
        }

        var databaseStatus = _databaseService.Initialize();
        if (!databaseStatus.Succeeded)
        {
            throw new InvalidOperationException(databaseStatus.Message);
        }

        var selectedClientIds = clientIds is { Count: > 0 }
            ? clientIds.Where(clientId => clientId > 0).ToHashSet()
            : null;
        var clients = (await _databaseService.GetClientsAsync())
            .Where(client => selectedClientIds is null || selectedClientIds.Contains(client.ClientId))
            .OrderBy(client => client.ClientName, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var availableSnapshots = new List<(BenefitsMasterDatabaseClientSnapshot Snapshot, DateTime LatestLoaded)>();

        foreach (var client in clients)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var planDesignRows = await _databaseService.GetPlanDesignAsync(client.ClientId);
            if (planDesignRows.Count == 0)
            {
                continue;
            }

            foreach (var row in planDesignRows)
            {
                row.ClientName = client.ClientName;
                row.ClientNumber = client.ClientNumber;
                row.Carrier = client.Carrier;
            }

            var workspace = await _workspaceDataService.LoadAsync(
                client.ClientId,
                client.ClientName,
                client.Carrier);
            availableSnapshots.Add((
                new BenefitsMasterDatabaseClientSnapshot
                {
                    Client = client,
                    PlanDesignRows = planDesignRows,
                    BenefitSummaries = workspace.BenefitSummaries
                        .Where(summary => summary.Sections.Count > 0)
                        .ToList()
                },
                planDesignRows.Max(row => row.DateLoaded)));
        }

        if (availableSnapshots.Count == 0)
        {
            throw new InvalidOperationException(
                selectedClientIds is { Count: > 0 }
                    ? "The selected client does not have saved Plan Design rows in SQLite."
                    : "No saved Plan Design rows are available in SQLite.");
        }

        // A disposable Benefits Master workbook represents one extraction target.
        // SQLite retains every client; an explicit filter can still rebuild a
        // particular client, while the manual default selects the latest save.
        var snapshots = selectedClientIds is null
            ? new[]
            {
                availableSnapshots
                    .OrderByDescending(item => item.LatestLoaded)
                    .ThenBy(item => item.Snapshot.Client.ClientName, StringComparer.OrdinalIgnoreCase)
                    .First()
                    .Snapshot
            }
            : availableSnapshots.Select(item => item.Snapshot).ToArray();

        cancellationToken.ThrowIfCancellationRequested();
        return await Task.Run(
            () => BookletPlanDesignWorkbookWriter.WriteDatabaseSnapshot(
                new BenefitsMasterDatabaseWorkbookRequest
                {
                    OutputFolder = outputFolder,
                    GeneratedAt = DateTime.Now,
                    Clients = snapshots
                }),
            cancellationToken);
    }
}
