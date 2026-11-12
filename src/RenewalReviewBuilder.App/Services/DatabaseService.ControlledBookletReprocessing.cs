using System.Text.Json;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    private static readonly JsonSerializerOptions ControlledBookletJsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private static async Task EnsureControlledBookletReprocessingSchemaAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, null, """
CREATE TABLE IF NOT EXISTS ControlledBookletReprocessingRun (
    run_id INTEGER PRIMARY KEY AUTOINCREMENT,
    workflow_version TEXT NOT NULL,
    source_document_id INTEGER NOT NULL,
    client_id INTEGER NOT NULL,
    client_name TEXT NOT NULL DEFAULT '',
    carrier TEXT NOT NULL DEFAULT '',
    source_file_name TEXT NOT NULL,
    source_file_path TEXT NOT NULL,
    source_file_hash TEXT NOT NULL,
    detected_client_name TEXT NOT NULL DEFAULT '',
    detected_carrier TEXT NOT NULL DEFAULT '',
    identity_confidence REAL,
    identity_source TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL,
    requested_by TEXT NOT NULL DEFAULT '',
    requested_at TEXT NOT NULL,
    staged_at TEXT,
    approved_by TEXT NOT NULL DEFAULT '',
    approved_at TEXT,
    committed_at TEXT,
    rejected_by TEXT NOT NULL DEFAULT '',
    rejected_at TEXT,
    output_folder TEXT NOT NULL DEFAULT '',
    master_workbook_path TEXT NOT NULL DEFAULT '',
    review_workbook_path TEXT NOT NULL DEFAULT '',
    backup_path TEXT NOT NULL DEFAULT '',
    before_row_count INTEGER NOT NULL DEFAULT 0,
    staged_row_count INTEGER NOT NULL DEFAULT 0,
    post_commit_row_count INTEGER NOT NULL DEFAULT 0,
    added_count INTEGER NOT NULL DEFAULT 0,
    changed_count INTEGER NOT NULL DEFAULT 0,
    removed_count INTEGER NOT NULL DEFAULT 0,
    blocking_issue_count INTEGER NOT NULL DEFAULT 0,
    warning_count INTEGER NOT NULL DEFAULT 0,
    before_snapshot_json TEXT NOT NULL DEFAULT '[]',
    staged_snapshot_json TEXT NOT NULL DEFAULT '[]',
    benefit_administration_json TEXT NOT NULL DEFAULT '[]',
    notes TEXT NOT NULL DEFAULT '',
    error_message TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id),
    FOREIGN KEY (client_id) REFERENCES Clients(client_id)
);

CREATE INDEX IF NOT EXISTS IX_ControlledBookletReprocessingRun_Document
ON ControlledBookletReprocessingRun(source_document_id, requested_at DESC);

CREATE INDEX IF NOT EXISTS IX_ControlledBookletReprocessingRun_Status
ON ControlledBookletReprocessingRun(status, requested_at DESC);

CREATE UNIQUE INDEX IF NOT EXISTS UX_ControlledBookletReprocessingRun_ActiveDocument
ON ControlledBookletReprocessingRun(source_document_id)
WHERE status IN ('Extracting', 'Ready for Approval', 'Committing');

CREATE TABLE IF NOT EXISTS ControlledBookletReprocessingIssue (
    issue_id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id INTEGER NOT NULL,
    issue_code TEXT NOT NULL,
    severity TEXT NOT NULL,
    is_blocking INTEGER NOT NULL DEFAULT 0 CHECK (is_blocking IN (0, 1)),
    benefit_category TEXT NOT NULL DEFAULT '',
    benefit_name TEXT NOT NULL DEFAULT '',
    field_name TEXT NOT NULL DEFAULT '',
    message TEXT NOT NULL,
    raw_source_reference TEXT NOT NULL DEFAULT '',
    source_page INTEGER,
    FOREIGN KEY (run_id) REFERENCES ControlledBookletReprocessingRun(run_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS IX_ControlledBookletReprocessingIssue_Run
ON ControlledBookletReprocessingIssue(run_id, is_blocking, issue_code);

CREATE TABLE IF NOT EXISTS ControlledBookletReprocessingChange (
    change_id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id INTEGER NOT NULL,
    change_type TEXT NOT NULL,
    comparison_key TEXT NOT NULL,
    class_name TEXT NOT NULL DEFAULT '',
    benefit_category TEXT NOT NULL DEFAULT '',
    benefit_name TEXT NOT NULL DEFAULT '',
    before_value TEXT NOT NULL DEFAULT '',
    after_value TEXT NOT NULL DEFAULT '',
    source_page INTEGER,
    FOREIGN KEY (run_id) REFERENCES ControlledBookletReprocessingRun(run_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS IX_ControlledBookletReprocessingChange_Run
ON ControlledBookletReprocessingChange(run_id, change_type, comparison_key);

CREATE TABLE IF NOT EXISTS ControlledBookletReprocessingEdit (
    edit_id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id INTEGER NOT NULL,
    comparison_key TEXT NOT NULL,
    field_name TEXT NOT NULL,
    previous_value TEXT NOT NULL DEFAULT '',
    new_value TEXT NOT NULL DEFAULT '',
    reason TEXT NOT NULL,
    edited_by TEXT NOT NULL,
    edited_at TEXT NOT NULL,
    FOREIGN KEY (run_id) REFERENCES ControlledBookletReprocessingRun(run_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS IX_ControlledBookletReprocessingEdit_Run
ON ControlledBookletReprocessingEdit(run_id, edited_at, edit_id);
""");
    }

    internal async Task<IReadOnlyList<ControlledBookletDocumentOption>> GetControlledBookletDocumentsAsync()
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, ControlledBookletDocumentSelectSql() + """

WHERE d.document_type = $document_type
  AND d.processing_status NOT IN ('identity_quarantined', 'superseded_duplicate')
ORDER BY c.client_name COLLATE NOCASE,
         carrier COLLATE NOCASE,
         d.file_name COLLATE NOCASE,
         d.document_id;
""");
        AddParameter(command, "$document_type", DocumentTypes.BenefitBooklet);
        return await ReadControlledBookletDocumentsAsync(command);
    }

    internal async Task<ControlledBookletDocumentOption?> GetControlledBookletDocumentAsync(long documentId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, ControlledBookletDocumentSelectSql() + """

WHERE d.document_id = $document_id
LIMIT 1;
""");
        AddParameter(command, "$document_id", documentId);
        return (await ReadControlledBookletDocumentsAsync(command)).SingleOrDefault();
    }

    internal async Task<IReadOnlyList<ControlledBookletDocumentOption>> GetControlledBookletHashAssignmentsAsync(string fileHash)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, ControlledBookletDocumentSelectSql() + """

WHERE d.file_hash = $file_hash
  AND d.processing_status NOT IN ('identity_quarantined', 'superseded_duplicate')
ORDER BY c.client_name, d.document_id;
""");
        AddParameter(command, "$file_hash", Clean(fileHash));
        return await ReadControlledBookletDocumentsAsync(command);
    }

    internal async Task<List<PlanDesignRecord>> GetPlanDesignForDocumentAsync(long documentId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM PlanDesign WHERE source_document_id = $document_id ORDER BY class_name, benefit_category, benefit_name, plan_design_id;");
        AddParameter(command, "$document_id", documentId);

        var rows = new List<PlanDesignRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var row = ReadPlanDesign(reader);
            row.ReimbursementText = row.ReimbursementPercent.HasValue
                ? $"{row.ReimbursementPercent.Value:0.####}%"
                : "";
            rows.Add(row);
        }

        return rows;
    }

    internal async Task<HashSet<string>> GetPlanDesignValidationIssueCodesForDocumentAsync(long documentId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, """
SELECT DISTINCT issue_code
FROM ValidationIssue
WHERE source_document_id = $document_id
  AND TRIM(issue_code) <> '';
""");
        AddParameter(command, "$document_id", documentId);
        var codes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            codes.Add(reader.GetString(0));
        }

        return codes;
    }

    internal async Task<long> BeginControlledBookletReprocessingRunAsync(
        ControlledBookletDocumentOption document,
        string requestedBy)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var transaction = connection.BeginTransaction();

        using (var active = CreateCommand(connection, transaction, """
SELECT run_id
FROM ControlledBookletReprocessingRun
WHERE source_document_id = $source_document_id
  AND status IN ('Extracting', 'Blocked', 'Ready for Approval', 'Committing')
ORDER BY run_id DESC
LIMIT 1;
"""))
        {
            AddParameter(active, "$source_document_id", document.DocumentId);
            var activeRunId = ToNullableLong(await active.ExecuteScalarAsync());
            if (activeRunId.HasValue)
            {
                throw new InvalidOperationException(
                    $"Document {document.DocumentId} already has active controlled reprocessing run {activeRunId.Value}.");
            }
        }

        using var insert = CreateCommand(connection, transaction, """
INSERT INTO ControlledBookletReprocessingRun (
    workflow_version, source_document_id, client_id, client_name, carrier,
    source_file_name, source_file_path, source_file_hash, status,
    requested_by, requested_at
)
VALUES (
    $workflow_version, $source_document_id, $client_id, $client_name, $carrier,
    $source_file_name, $source_file_path, $source_file_hash, $status,
    $requested_by, $requested_at
);
SELECT last_insert_rowid();
""");
        AddParameter(insert, "$workflow_version", ControlledBookletReprocessingContract.Version);
        AddParameter(insert, "$source_document_id", document.DocumentId);
        AddParameter(insert, "$client_id", document.ClientId);
        AddParameter(insert, "$client_name", Clean(document.ClientName));
        AddParameter(insert, "$carrier", Clean(document.Carrier));
        AddParameter(insert, "$source_file_name", Clean(document.FileName));
        AddParameter(insert, "$source_file_path", Clean(document.FilePath));
        AddParameter(insert, "$source_file_hash", Clean(document.FileHash));
        AddParameter(insert, "$status", ControlledBookletReprocessingStatuses.Extracting);
        AddParameter(insert, "$requested_by", Clean(requestedBy));
        AddParameter(insert, "$requested_at", ToDateTimeText(DateTime.Now));
        var runId = ToLong(await insert.ExecuteScalarAsync());
        transaction.Commit();
        return runId;
    }

    internal async Task SaveControlledBookletReprocessingStageAsync(
        ControlledBookletReprocessingRun run,
        IReadOnlyCollection<ControlledBookletReprocessingEdit>? newEdits = null)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var transaction = connection.BeginTransaction();

        using (var update = CreateCommand(connection, transaction, """
UPDATE ControlledBookletReprocessingRun
SET detected_client_name = $detected_client_name,
    detected_carrier = $detected_carrier,
    identity_confidence = $identity_confidence,
    identity_source = $identity_source,
    status = $status,
    staged_at = $staged_at,
    output_folder = $output_folder,
    master_workbook_path = $master_workbook_path,
    review_workbook_path = $review_workbook_path,
    before_row_count = $before_row_count,
    staged_row_count = $staged_row_count,
    added_count = $added_count,
    changed_count = $changed_count,
    removed_count = $removed_count,
    blocking_issue_count = $blocking_issue_count,
    warning_count = $warning_count,
    before_snapshot_json = $before_snapshot_json,
    staged_snapshot_json = $staged_snapshot_json,
    benefit_administration_json = $benefit_administration_json,
    notes = $notes,
    error_message = $error_message
WHERE run_id = $run_id;
"""))
        {
            AddParameter(update, "$detected_client_name", Clean(run.DetectedClientName));
            AddParameter(update, "$detected_carrier", Clean(run.DetectedCarrier));
            AddParameter(update, "$identity_confidence", run.IdentityConfidence);
            AddParameter(update, "$identity_source", Clean(run.IdentitySource));
            AddParameter(update, "$status", Clean(run.Status));
            AddParameter(update, "$staged_at", ToNullableDateTimeText(run.StagedAt));
            AddParameter(update, "$output_folder", Clean(run.OutputFolder));
            AddParameter(update, "$master_workbook_path", Clean(run.MasterWorkbookPath));
            AddParameter(update, "$review_workbook_path", Clean(run.ReviewWorkbookPath));
            AddParameter(update, "$before_row_count", run.BeforeRowCount);
            AddParameter(update, "$staged_row_count", run.StagedRowCount);
            AddParameter(update, "$added_count", run.AddedCount);
            AddParameter(update, "$changed_count", run.ChangedCount);
            AddParameter(update, "$removed_count", run.RemovedCount);
            AddParameter(update, "$blocking_issue_count", run.BlockingIssueCount);
            AddParameter(update, "$warning_count", run.WarningCount);
            AddParameter(update, "$before_snapshot_json", JsonSerializer.Serialize(run.BeforeRows, ControlledBookletJsonOptions));
            AddParameter(update, "$staged_snapshot_json", JsonSerializer.Serialize(run.StagedRows, ControlledBookletJsonOptions));
            AddParameter(update, "$benefit_administration_json", JsonSerializer.Serialize(run.BenefitAdministrationRecords, ControlledBookletJsonOptions));
            AddParameter(update, "$notes", Clean(run.Notes));
            AddParameter(update, "$error_message", Clean(run.ErrorMessage));
            AddParameter(update, "$run_id", run.RunId);
            if (await update.ExecuteNonQueryAsync() != 1)
            {
                throw new InvalidOperationException($"Controlled booklet reprocessing run {run.RunId} was not found.");
            }
        }

        await ExecuteNonQueryAsync(
            connection,
            transaction,
            "DELETE FROM ControlledBookletReprocessingIssue WHERE run_id = $run_id;",
            command => AddParameter(command, "$run_id", run.RunId));
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            "DELETE FROM ControlledBookletReprocessingChange WHERE run_id = $run_id;",
            command => AddParameter(command, "$run_id", run.RunId));

        foreach (var issue in run.Issues)
        {
            using var insertIssue = CreateCommand(connection, transaction, """
INSERT INTO ControlledBookletReprocessingIssue (
    run_id, issue_code, severity, is_blocking, benefit_category, benefit_name,
    field_name, message, raw_source_reference, source_page
)
VALUES (
    $run_id, $issue_code, $severity, $is_blocking, $benefit_category, $benefit_name,
    $field_name, $message, $raw_source_reference, $source_page
);
""");
            AddParameter(insertIssue, "$run_id", run.RunId);
            AddParameter(insertIssue, "$issue_code", Clean(issue.IssueCode));
            AddParameter(insertIssue, "$severity", Clean(issue.Severity));
            AddParameter(insertIssue, "$is_blocking", issue.IsBlocking ? 1 : 0);
            AddParameter(insertIssue, "$benefit_category", Clean(issue.BenefitCategory));
            AddParameter(insertIssue, "$benefit_name", Clean(issue.BenefitName));
            AddParameter(insertIssue, "$field_name", Clean(issue.FieldName));
            AddParameter(insertIssue, "$message", Clean(issue.Message));
            AddParameter(insertIssue, "$raw_source_reference", Clean(issue.RawSourceReference));
            AddParameter(insertIssue, "$source_page", issue.SourcePage);
            await insertIssue.ExecuteNonQueryAsync();
        }

        foreach (var change in run.Changes)
        {
            using var insertChange = CreateCommand(connection, transaction, """
INSERT INTO ControlledBookletReprocessingChange (
    run_id, change_type, comparison_key, class_name, benefit_category,
    benefit_name, before_value, after_value, source_page
)
VALUES (
    $run_id, $change_type, $comparison_key, $class_name, $benefit_category,
    $benefit_name, $before_value, $after_value, $source_page
);
""");
            AddParameter(insertChange, "$run_id", run.RunId);
            AddParameter(insertChange, "$change_type", Clean(change.ChangeType));
            AddParameter(insertChange, "$comparison_key", Clean(change.ComparisonKey));
            AddParameter(insertChange, "$class_name", Clean(change.ClassName));
            AddParameter(insertChange, "$benefit_category", Clean(change.BenefitCategory));
            AddParameter(insertChange, "$benefit_name", Clean(change.BenefitName));
            AddParameter(insertChange, "$before_value", Clean(change.BeforeValue));
            AddParameter(insertChange, "$after_value", Clean(change.AfterValue));
            AddParameter(insertChange, "$source_page", change.SourcePage);
            await insertChange.ExecuteNonQueryAsync();
        }

        foreach (var edit in newEdits ?? [])
        {
            using var insertEdit = CreateCommand(connection, transaction, """
INSERT INTO ControlledBookletReprocessingEdit (
    run_id, comparison_key, field_name, previous_value, new_value,
    reason, edited_by, edited_at
)
VALUES (
    $run_id, $comparison_key, $field_name, $previous_value, $new_value,
    $reason, $edited_by, $edited_at
);
""");
            AddParameter(insertEdit, "$run_id", run.RunId);
            AddParameter(insertEdit, "$comparison_key", Clean(edit.ComparisonKey));
            AddParameter(insertEdit, "$field_name", Clean(edit.FieldName));
            AddParameter(insertEdit, "$previous_value", Clean(edit.PreviousValue));
            AddParameter(insertEdit, "$new_value", Clean(edit.NewValue));
            AddParameter(insertEdit, "$reason", Clean(edit.Reason));
            AddParameter(insertEdit, "$edited_by", Clean(edit.EditedBy));
            AddParameter(insertEdit, "$edited_at", ToDateTimeText(edit.EditedAt));
            await insertEdit.ExecuteNonQueryAsync();
        }

        transaction.Commit();
    }

    internal async Task<ControlledBookletReprocessingRun?> GetControlledBookletReprocessingRunAsync(long runId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        ControlledBookletReprocessingRun? run;
        using (var command = CreateCommand(connection, null, "SELECT * FROM ControlledBookletReprocessingRun WHERE run_id = $run_id LIMIT 1;"))
        {
            AddParameter(command, "$run_id", runId);
            using var reader = await command.ExecuteReaderAsync();
            run = await reader.ReadAsync() ? ReadControlledBookletReprocessingRun(reader) : null;
        }

        if (run is null)
        {
            return null;
        }

        using (var command = CreateCommand(connection, null, "SELECT * FROM ControlledBookletReprocessingIssue WHERE run_id = $run_id ORDER BY is_blocking DESC, issue_id;"))
        {
            AddParameter(command, "$run_id", runId);
            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                run.Issues.Add(new ControlledBookletReprocessingIssue
                {
                    IssueId = ReadLong(reader, "issue_id"),
                    RunId = runId,
                    IssueCode = ReadString(reader, "issue_code"),
                    Severity = ReadString(reader, "severity"),
                    IsBlocking = (ReadInt(reader, "is_blocking") ?? 0) == 1,
                    BenefitCategory = ReadString(reader, "benefit_category"),
                    BenefitName = ReadString(reader, "benefit_name"),
                    FieldName = ReadString(reader, "field_name"),
                    Message = ReadString(reader, "message"),
                    RawSourceReference = ReadString(reader, "raw_source_reference"),
                    SourcePage = ReadInt(reader, "source_page")
                });
            }
        }

        using (var command = CreateCommand(connection, null, "SELECT * FROM ControlledBookletReprocessingChange WHERE run_id = $run_id ORDER BY change_type, comparison_key, change_id;"))
        {
            AddParameter(command, "$run_id", runId);
            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                run.Changes.Add(new ControlledBookletReprocessingChange
                {
                    ChangeId = ReadLong(reader, "change_id"),
                    RunId = runId,
                    ChangeType = ReadString(reader, "change_type"),
                    ComparisonKey = ReadString(reader, "comparison_key"),
                    ClassName = ReadString(reader, "class_name"),
                    BenefitCategory = ReadString(reader, "benefit_category"),
                    BenefitName = ReadString(reader, "benefit_name"),
                    BeforeValue = ReadString(reader, "before_value"),
                    AfterValue = ReadString(reader, "after_value"),
                    SourcePage = ReadInt(reader, "source_page")
                });
            }
        }

        using (var command = CreateCommand(connection, null, "SELECT * FROM ControlledBookletReprocessingEdit WHERE run_id = $run_id ORDER BY edited_at, edit_id;"))
        {
            AddParameter(command, "$run_id", runId);
            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                run.Edits.Add(new ControlledBookletReprocessingEdit
                {
                    EditId = ReadLong(reader, "edit_id"),
                    RunId = runId,
                    ComparisonKey = ReadString(reader, "comparison_key"),
                    FieldName = ReadString(reader, "field_name"),
                    PreviousValue = ReadString(reader, "previous_value"),
                    NewValue = ReadString(reader, "new_value"),
                    Reason = ReadString(reader, "reason"),
                    EditedBy = ReadString(reader, "edited_by"),
                    EditedAt = ReadDateTime(reader, "edited_at") ?? DateTime.MinValue
                });
            }
        }

        return run;
    }

    internal async Task<ControlledBookletReprocessingRun?> GetActiveControlledBookletReprocessingRunForDocumentAsync(
        long documentId)
    {
        await InitializeCoreAsync();
        long? activeRunId;
        using (var connection = await OpenConnectionAsync())
        using (var command = CreateCommand(connection, null, """
SELECT run_id
FROM ControlledBookletReprocessingRun
WHERE source_document_id = $source_document_id
  AND status IN ('Extracting', 'Blocked', 'Ready for Approval', 'Committing')
ORDER BY run_id DESC
LIMIT 1;
"""))
        {
            AddParameter(command, "$source_document_id", documentId);
            activeRunId = ToNullableLong(await command.ExecuteScalarAsync());
        }

        return activeRunId.HasValue
            ? await GetControlledBookletReprocessingRunAsync(activeRunId.Value)
            : null;
    }

    internal async Task BeginControlledBookletCommitAsync(long runId, string approvedBy, string backupPath)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var transaction = connection.BeginTransaction();
        using var command = CreateCommand(connection, transaction, """
UPDATE ControlledBookletReprocessingRun
SET status = $committing_status,
    approved_by = $approved_by,
    approved_at = $approved_at,
    backup_path = $backup_path,
    error_message = ''
WHERE run_id = $run_id
  AND status = $ready_status
  AND blocking_issue_count = 0
  AND staged_row_count > 0;
""");
        AddParameter(command, "$committing_status", ControlledBookletReprocessingStatuses.Committing);
        AddParameter(command, "$approved_by", Clean(approvedBy));
        AddParameter(command, "$approved_at", ToDateTimeText(DateTime.Now));
        AddParameter(command, "$backup_path", Clean(backupPath));
        AddParameter(command, "$run_id", runId);
        AddParameter(command, "$ready_status", ControlledBookletReprocessingStatuses.ReadyForApproval);
        if (await command.ExecuteNonQueryAsync() != 1)
        {
            throw new InvalidOperationException("This staged booklet is no longer eligible for approval.");
        }

        transaction.Commit();
    }

    internal async Task CompleteControlledBookletCommitAsync(
        long runId,
        int postCommitRowCount,
        string notes)
    {
        await UpdateControlledBookletRunStatusAsync(
            runId,
            ControlledBookletReprocessingStatuses.Committed,
            notes,
            "",
            postCommitRowCount,
            committedAt: DateTime.Now);
    }

    internal async Task FailControlledBookletRunAsync(long runId, string errorMessage)
    {
        await UpdateControlledBookletRunStatusAsync(
            runId,
            ControlledBookletReprocessingStatuses.Failed,
            "",
            errorMessage,
            null);
    }

    internal async Task RejectControlledBookletRunAsync(long runId, string rejectedBy, string reason)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, """
UPDATE ControlledBookletReprocessingRun
SET status = $status,
    rejected_by = $rejected_by,
    rejected_at = $rejected_at,
    notes = $notes
WHERE run_id = $run_id
  AND status IN ('Extracting', 'Blocked', 'Ready for Approval');
""");
        AddParameter(command, "$status", ControlledBookletReprocessingStatuses.Rejected);
        AddParameter(command, "$rejected_by", Clean(rejectedBy));
        AddParameter(command, "$rejected_at", ToDateTimeText(DateTime.Now));
        AddParameter(command, "$notes", Clean(reason));
        AddParameter(command, "$run_id", runId);
        if (await command.ExecuteNonQueryAsync() != 1)
        {
            throw new InvalidOperationException("Only extracting, blocked, or ready runs can be rejected.");
        }
    }

    internal async Task<string> CreateControlledBookletBackupAsync(long runId, long documentId)
    {
        await InitializeCoreAsync();
        var databaseFolder = Path.GetDirectoryName(DatabasePath)
            ?? throw new InvalidOperationException("The Benefits Intelligence database folder could not be resolved.");
        var applicationFolder = Path.GetFileName(databaseFolder).Equals("data", StringComparison.OrdinalIgnoreCase)
            ? Directory.GetParent(databaseFolder)?.FullName ?? databaseFolder
            : databaseFolder;
        var backupFolder = Path.Combine(applicationFolder, "backups", "booklet-reprocessing");
        Directory.CreateDirectory(backupFolder);
        var backupPath = Path.Combine(
            backupFolder,
            $"Benefits_Intelligence_before_booklet_{documentId}_run_{runId}_{DateTime.Now:yyyyMMdd_HHmmssfff}.db");

        using var source = await OpenConnectionAsync();
        using var destination = new SqliteConnection($"Data Source={backupPath}");
        await destination.OpenAsync();
        source.BackupDatabase(destination);
        using var check = destination.CreateCommand();
        check.CommandText = "PRAGMA quick_check;";
        var quickCheck = Convert.ToString(await check.ExecuteScalarAsync()) ?? "";
        if (!quickCheck.Equals("ok", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"The pre-commit database backup failed its integrity check: {quickCheck}");
        }

        return backupPath;
    }

    internal async Task<(int RowCount, int OutOfRangeCount)> GetPlanDesignPostCommitAuditAsync(long documentId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, """
SELECT
    COUNT(*) AS row_count,
    SUM(CASE WHEN reimbursement_percent < 0 OR reimbursement_percent > 100 THEN 1 ELSE 0 END) AS out_of_range_count
FROM PlanDesign
WHERE source_document_id = $document_id;
""");
        AddParameter(command, "$document_id", documentId);
        using var reader = await command.ExecuteReaderAsync();
        if (!await reader.ReadAsync())
        {
            return (0, 0);
        }

        return (ReadInt(reader, "row_count") ?? 0, ReadInt(reader, "out_of_range_count") ?? 0);
    }

    private async Task UpdateControlledBookletRunStatusAsync(
        long runId,
        string status,
        string notes,
        string errorMessage,
        int? postCommitRowCount,
        DateTime? committedAt = null)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, """
UPDATE ControlledBookletReprocessingRun
SET status = $status,
    notes = CASE WHEN TRIM($notes) = '' THEN notes ELSE $notes END,
    error_message = $error_message,
    post_commit_row_count = COALESCE($post_commit_row_count, post_commit_row_count),
    committed_at = COALESCE($committed_at, committed_at)
WHERE run_id = $run_id;
""");
        AddParameter(command, "$status", Clean(status));
        AddParameter(command, "$notes", Clean(notes));
        AddParameter(command, "$error_message", Clean(errorMessage));
        AddParameter(command, "$post_commit_row_count", postCommitRowCount);
        AddParameter(command, "$committed_at", ToNullableDateTimeText(committedAt));
        AddParameter(command, "$run_id", runId);
        await command.ExecuteNonQueryAsync();
    }

    private static string ControlledBookletDocumentSelectSql() => """
SELECT
    d.document_id,
    d.client_id,
    c.client_name,
    CASE
        WHEN NULLIF(TRIM(d.carrier), '') IS NOT NULL THEN d.carrier
        WHEN c.client_name = 'Unknown Client' COLLATE NOCASE THEN ''
        ELSE COALESCE(c.carrier, '')
    END AS carrier,
    d.document_type,
    d.file_name,
    d.file_path,
    d.file_hash,
    d.processing_status,
    (SELECT COUNT(*) FROM PlanDesign p WHERE p.source_document_id = d.document_id) AS current_row_count,
    (SELECT COUNT(*)
       FROM PlanDesign p
      WHERE p.source_document_id = d.document_id
        AND (p.reimbursement_percent < 0 OR p.reimbursement_percent > 100)) AS out_of_range_percentage_count,
    (SELECT COUNT(DISTINCT other.client_id)
       FROM Documents other
      WHERE other.file_hash = d.file_hash
        AND other.client_id <> d.client_id
        AND other.processing_status NOT IN ('identity_quarantined', 'superseded_duplicate')) AS cross_client_hash_assignment_count
FROM Documents d
JOIN Clients c ON c.client_id = d.client_id
""";

    private static async Task<List<ControlledBookletDocumentOption>> ReadControlledBookletDocumentsAsync(SqliteCommand command)
    {
        var documents = new List<ControlledBookletDocumentOption>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            documents.Add(new ControlledBookletDocumentOption
            {
                DocumentId = ReadLong(reader, "document_id"),
                ClientId = ReadLong(reader, "client_id"),
                ClientName = ClientNameResolver.ToDisplayName(ReadString(reader, "client_name")),
                Carrier = ReadString(reader, "carrier"),
                DocumentType = ReadString(reader, "document_type"),
                FileName = ReadString(reader, "file_name"),
                FilePath = ReadString(reader, "file_path"),
                FileHash = ReadString(reader, "file_hash"),
                ProcessingStatus = ReadString(reader, "processing_status"),
                CurrentRowCount = ReadInt(reader, "current_row_count") ?? 0,
                OutOfRangePercentageCount = ReadInt(reader, "out_of_range_percentage_count") ?? 0,
                CrossClientHashAssignmentCount = ReadInt(reader, "cross_client_hash_assignment_count") ?? 0
            });
        }

        return documents;
    }

    private static ControlledBookletReprocessingRun ReadControlledBookletReprocessingRun(SqliteDataReader reader)
    {
        var run = new ControlledBookletReprocessingRun
        {
            RunId = ReadLong(reader, "run_id"),
            WorkflowVersion = ReadString(reader, "workflow_version"),
            SourceDocumentId = ReadLong(reader, "source_document_id"),
            ClientId = ReadLong(reader, "client_id"),
            ClientName = ReadString(reader, "client_name"),
            Carrier = ReadString(reader, "carrier"),
            SourceFileName = ReadString(reader, "source_file_name"),
            SourceFilePath = ReadString(reader, "source_file_path"),
            SourceFileHash = ReadString(reader, "source_file_hash"),
            DetectedClientName = ReadString(reader, "detected_client_name"),
            DetectedCarrier = ReadString(reader, "detected_carrier"),
            IdentityConfidence = ReadDecimal(reader, "identity_confidence"),
            IdentitySource = ReadString(reader, "identity_source"),
            Status = ReadString(reader, "status"),
            RequestedBy = ReadString(reader, "requested_by"),
            RequestedAt = ReadDateTime(reader, "requested_at") ?? DateTime.MinValue,
            StagedAt = ReadDateTime(reader, "staged_at"),
            ApprovedBy = ReadString(reader, "approved_by"),
            ApprovedAt = ReadDateTime(reader, "approved_at"),
            CommittedAt = ReadDateTime(reader, "committed_at"),
            RejectedBy = ReadString(reader, "rejected_by"),
            RejectedAt = ReadDateTime(reader, "rejected_at"),
            OutputFolder = ReadString(reader, "output_folder"),
            MasterWorkbookPath = ReadString(reader, "master_workbook_path"),
            ReviewWorkbookPath = ReadString(reader, "review_workbook_path"),
            BackupPath = ReadString(reader, "backup_path"),
            BeforeRowCount = ReadInt(reader, "before_row_count") ?? 0,
            StagedRowCount = ReadInt(reader, "staged_row_count") ?? 0,
            PostCommitRowCount = ReadInt(reader, "post_commit_row_count") ?? 0,
            AddedCount = ReadInt(reader, "added_count") ?? 0,
            ChangedCount = ReadInt(reader, "changed_count") ?? 0,
            RemovedCount = ReadInt(reader, "removed_count") ?? 0,
            BlockingIssueCount = ReadInt(reader, "blocking_issue_count") ?? 0,
            WarningCount = ReadInt(reader, "warning_count") ?? 0,
            Notes = ReadString(reader, "notes"),
            ErrorMessage = ReadString(reader, "error_message")
        };

        AddDeserializedRows(run.BeforeRows, ReadString(reader, "before_snapshot_json"));
        AddDeserializedRows(run.StagedRows, ReadString(reader, "staged_snapshot_json"));
        run.BenefitAdministrationRecords.AddRange(
            JsonSerializer.Deserialize<List<BenefitAdministrationExtractionCandidate>>(
                ReadString(reader, "benefit_administration_json"),
                ControlledBookletJsonOptions) ?? []);
        return run;
    }

    private static void AddDeserializedRows(ICollection<PlanDesignRecord> target, string json)
    {
        foreach (var row in JsonSerializer.Deserialize<List<PlanDesignRecord>>(json, ControlledBookletJsonOptions) ?? [])
        {
            target.Add(row);
        }
    }

    private static string? ToNullableDateTimeText(DateTime? value) =>
        value.HasValue ? ToDateTimeText(value.Value) : null;
}
