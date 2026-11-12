using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Microsoft.Data.Sqlite;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientHistoryImportPersistenceTests
{
    [Fact]
    public async Task Imports_And_Queries_Only_The_Selected_Database_Client()
    {
        var folder = Path.Combine(Path.GetTempPath(), $"history-import-client-selection-{Guid.NewGuid():N}"); Directory.CreateDirectory(folder);
        var db = Path.Combine(folder, "history.db"); var source = Path.Combine(folder, "selectedClientHistory.xlsx"); await File.WriteAllTextAsync(source, "fixture");
        try
        {
            var service = new DatabaseService(db);
            var firstClient = await service.RecordPolicyScheduleReviewAsync(new ClientRecord { ClientName = "First client" }, new DiscoveredDocument { FilePath = source, FileName = "selectedClientHistory.xlsx", DocumentType = DiscoveryDocumentTypes.GroupPolicySchedule }, "test", "setup");
            var selectedClient = await service.RecordPolicyScheduleReviewAsync(new ClientRecord { ClientName = "Selected client" }, new DiscoveredDocument { FilePath = source, FileName = "selectedClientHistory.xlsx", DocumentType = DiscoveryDocumentTypes.GroupPolicySchedule }, "test", "setup");
            var selectedClientId = Assert.IsType<long>(selectedClient.ClientId);
            var row = new ClientHistoryPreviewEvent { EffectiveDate = new DateOnly(2025, 1, 1), EventType = "Meetings", Title = "Selected client meeting", Description = "Selected client meeting", SourceFile = source, SourceSheet = "Group History", SourceRow = 61, OriginalText = "January 1 2025 | Selected client meeting", ExtractionMethod = "section-row-date", Confidence = .95m, PreviewStatus = ClientHistoryPreviewStatuses.Approved };

            await service.ImportApprovedClientHistoryEventsAsync(selectedClientId, [row]);

            Assert.Empty(await service.GetClientHistoryEventsAsync(Assert.IsType<long>(firstClient.ClientId)));
            var imported = Assert.Single(await service.GetClientHistoryEventsAsync(selectedClientId));
            Assert.Equal(selectedClientId, imported.ClientId);
            Assert.Equal("Selected client meeting", imported.Title);
        }
        finally { SqliteConnection.ClearAllPools(); Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task Imports_Approved_Row_Preserves_Evidence_And_Is_Idempotent()
    {
        var folder = Path.Combine(Path.GetTempPath(), $"history-import-{Guid.NewGuid():N}"); Directory.CreateDirectory(folder);
        var db = Path.Combine(folder, "history.db"); var source = Path.Combine(folder, "testHistory.xlsx"); await File.WriteAllTextAsync(source, "fixture");
        try
        {
            var service = new DatabaseService(db);
            var saved = await service.RecordPolicyScheduleReviewAsync(new ClientRecord { ClientName = "History Test Client" }, new DiscoveredDocument { FilePath = source, FileName = "testHistory.xlsx", DocumentType = DiscoveryDocumentTypes.GroupPolicySchedule }, "test", "setup");
            var row = new ClientHistoryPreviewEvent { EffectiveDate = new DateOnly(2025, 1, 1), EventType = "Meetings", Title = "Renewal meeting", Description = "Renewal meeting", SourceFile = source, SourceSheet = "Group History", SourceRow = 61, OriginalText = "January 1 2025 | Renewal meeting", ExtractionMethod = "section-row-date", Confidence = .95m, PreviewStatus = ClientHistoryPreviewStatuses.Approved };
            var first = await service.ImportApprovedClientHistoryEventsAsync(saved.ClientId!.Value, [row]);
            Assert.Equal(1, first.Imported);
            var imported = Assert.Single(await service.GetClientHistoryEventsAsync(saved.ClientId.Value));
            Assert.Equal("Group History", imported.SourceSheet); Assert.Equal(61, imported.SourceRow); Assert.Equal(row.OriginalText, imported.OriginalText);
            row.PreviewStatus = ClientHistoryPreviewStatuses.Approved;
            var second = await service.ImportApprovedClientHistoryEventsAsync(saved.ClientId.Value, [row]);
            Assert.Equal(1, second.SkippedDuplicates);
        }
        finally { SqliteConnection.ClearAllPools(); Directory.Delete(folder, true); }
    }
}
