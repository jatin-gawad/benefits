using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    public async Task<ClientHistoryImportResult> ImportApprovedClientHistoryEventsAsync(long clientId, IEnumerable<ClientHistoryPreviewEvent> previewRows, string importedBy = "Desktop user")
    {
        await InitializeCoreAsync();
        var result = new ClientHistoryImportResult();
        var rows = previewRows.ToList();
        result.Approved = rows.Count(row => row.PreviewStatus == ClientHistoryPreviewStatuses.Approved);
        result.Rejected = rows.Count(row => row.PreviewStatus == ClientHistoryPreviewStatuses.Rejected);
        using var connection = await OpenConnectionAsync();
        using var transaction = connection.BeginTransaction();
        try
        {
            var now = DateTime.UtcNow;
            var batchId = await InsertHistoryBatchAsync(connection, transaction, rows.FirstOrDefault()?.SourceFile ?? "", importedBy, now);
            foreach (var row in rows.Where(row => row.PreviewStatus == ClientHistoryPreviewStatuses.Approved))
            {
                var fingerprint = HistoryFingerprint(clientId, row);
                var duplicate = await FindHistoryDuplicateAsync(connection, transaction, clientId, row, fingerprint);
                if (duplicate)
                {
                    row.DuplicateStatus = "Skipped duplicate";
                    result.SkippedDuplicates++;
                    continue;
                }

                await InsertHistoryEventAsync(connection, transaction, clientId, row, batchId, fingerprint, importedBy, now);
                row.PreviewStatus = ClientHistoryPreviewStatuses.Imported;
                result.Imported++;
            }
            transaction.Commit();
            InvalidateExecutiveReportingCache();
        }
        catch (Exception ex)
        {
            transaction.Rollback();
            AppLogger.Error(ex);
            result.Failed = result.Approved;
            result.Warnings.Add($"No approved history events were committed: {ex.Message}");
        }
        return result;
    }

    public async Task<List<ClientHistoryEventRecord>> GetClientHistoryEventsAsync(long clientId)
    {
        await InitializeCoreAsync(); using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, "SELECT * FROM ClientHistoryEvent WHERE client_id=$client_id ORDER BY effective_date, history_event_id;");
        AddParameter(command, "$client_id", clientId); var items = new List<ClientHistoryEventRecord>(); using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync()) items.Add(new ClientHistoryEventRecord { HistoryEventId=reader.GetInt64(reader.GetOrdinal("history_event_id")), ClientId=clientId, EffectiveDate=DateOnly.TryParse(reader["effective_date"]?.ToString(), out var date) ? date : null, EventType=reader["event_type"].ToString() ?? "", BenefitCategory=reader["benefit_category"].ToString() ?? "", ChangeType=reader["change_type"].ToString() ?? "", Title=reader["title"].ToString() ?? "", Description=reader["description"].ToString() ?? "", PreviousValue=reader["previous_value"].ToString() ?? "", NewValue=reader["new_value"].ToString() ?? "", SourceDocumentId=reader["source_document_id"] is DBNull ? null : Convert.ToInt64(reader["source_document_id"], CultureInfo.InvariantCulture), SourceSheet=reader["source_sheet"].ToString() ?? "", SourceRow=Convert.ToInt32(reader["source_row"], CultureInfo.InvariantCulture), OriginalText=reader["original_text"].ToString() ?? "", ExtractionMethod=reader["extraction_method"].ToString() ?? "", Confidence=Convert.ToDecimal(reader["confidence"], CultureInfo.InvariantCulture), ReviewStatus=reader["review_status"].ToString() ?? "", ImportBatchId=reader["import_batch_id"] is DBNull ? null : Convert.ToInt64(reader["import_batch_id"], CultureInfo.InvariantCulture), SourceFingerprint=reader["source_fingerprint"].ToString() ?? "" });
        return items;
    }

    private static string HistoryFingerprint(long clientId, ClientHistoryPreviewEvent row)
    {
        var source = $"{clientId}|{row.SourceFile}|{row.SourceSheet}|{row.SourceRow}|{row.OriginalText}";
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(source)));
    }
    private static async Task<bool> FindHistoryDuplicateAsync(SqliteConnection c, SqliteTransaction t, long clientId, ClientHistoryPreviewEvent row, string fingerprint)
    {
        using var command=CreateCommand(c,t,"SELECT 1 FROM ClientHistoryEvent WHERE client_id=$client_id AND (source_fingerprint=$fingerprint OR (effective_date=$date AND event_type=$event_type AND LOWER(TRIM(description))=LOWER(TRIM($description))) OR (source_sheet=$sheet AND source_row=$row)) LIMIT 1;");
        AddParameter(command,"$client_id",clientId); AddParameter(command,"$fingerprint",fingerprint); AddParameter(command,"$date",row.EffectiveDate?.ToString("yyyy-MM-dd") ?? ""); AddParameter(command,"$event_type",row.EventType); AddParameter(command,"$description",row.Description); AddParameter(command,"$sheet",row.SourceSheet); AddParameter(command,"$row",row.SourceRow); return await command.ExecuteScalarAsync() is not null;
    }
    private static async Task<long> InsertHistoryBatchAsync(SqliteConnection c, SqliteTransaction t, string sourceFile, string user, DateTime now)
    {
        using var command=CreateCommand(c,t,"INSERT INTO ImportBatch(source_type,source_file_name,document_type,status,imported_at,imported_by,extraction_method,raw_source_reference) VALUES('Excel', $file, 'Client History', 'Imported', $at, $by, 'client-history-preview', $file); SELECT last_insert_rowid();"); AddParameter(command,"$file",sourceFile); AddParameter(command,"$at",now.ToString("O")); AddParameter(command,"$by",user); return Convert.ToInt64(await command.ExecuteScalarAsync(),CultureInfo.InvariantCulture);
    }
    private static async Task InsertHistoryEventAsync(SqliteConnection c, SqliteTransaction t, long clientId, ClientHistoryPreviewEvent row, long batchId, string fingerprint, string user, DateTime now)
    {
        using var command=CreateCommand(c,t,"INSERT INTO ClientHistoryEvent(client_id,effective_date,event_type,benefit_category,change_type,title,description,previous_value,new_value,source_sheet,source_row,original_text,extraction_method,confidence,review_status,created_at,created_by,updated_at,updated_by,import_batch_id,source_fingerprint) VALUES($client,$date,$type,$benefit,'',$title,$description,$previous,$new,$sheet,$row,$original,$method,$confidence,'Imported',$at,$by,$at,$by,$batch,$fingerprint);");
        AddParameter(command,"$client",clientId); AddParameter(command,"$date",row.EffectiveDate?.ToString("yyyy-MM-dd") ?? ""); AddParameter(command,"$type",row.EventType); AddParameter(command,"$benefit",row.BenefitCategory); AddParameter(command,"$title",row.Title); AddParameter(command,"$description",row.Description); AddParameter(command,"$previous",row.PreviousValue); AddParameter(command,"$new",row.NewValue); AddParameter(command,"$sheet",row.SourceSheet); AddParameter(command,"$row",row.SourceRow); AddParameter(command,"$original",row.OriginalText); AddParameter(command,"$method",row.ExtractionMethod); AddParameter(command,"$confidence",row.Confidence); AddParameter(command,"$at",now.ToString("O")); AddParameter(command,"$by",user); AddParameter(command,"$batch",batchId); AddParameter(command,"$fingerprint",fingerprint); await command.ExecuteNonQueryAsync();
    }
}
