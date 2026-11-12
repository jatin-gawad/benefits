using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Application service for canonical financial terms. Renewal comparison only reads its versions.</summary>
public sealed class FinancialArrangementService
{
    private readonly string _databasePath;
    private readonly IFinancialArrangementAuthorizationService _authorization;

    public FinancialArrangementService(string databasePath, IFinancialArrangementAuthorizationService? authorization = null)
    {
        _databasePath = databasePath;
        _authorization = authorization ?? new FinancialArrangementAuthorizationService();
    }

    /// <summary>
    /// Extracts reviewable financial terms behind the application-service boundary.
    /// Persistence still requires the existing explicit confirmation workflow.
    /// </summary>
    public FinancialArrangementExtraction ExtractSourceTerms(string path, DateOnly effectiveDate)
    {
        return new FinancialArrangementPdfExtractor().Extract(path, effectiveDate);
    }

    public async Task InitializeAsync()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_databasePath)!);
        await using var connection = await OpenAsync();
        await FinancialArrangementDatabaseMigration.ApplyAsync(connection);
    }

    public async Task<FinancialArrangement> SaveFinancialArrangementAsync(FinancialArrangement arrangement, bool supersedeExisting = true)
    {
        await InitializeAsync();
        await using var connection = await OpenAsync();
        using var transaction = connection.BeginTransaction();
        var now = DateTime.UtcNow.ToString("O");
        if (supersedeExisting)
        {
            await ExecuteAsync(connection, transaction, "UPDATE financial_arrangements SET superseded_at = $now WHERE client_id = $client AND carrier = $carrier AND effective_date = $effective AND superseded_at IS NULL;", command =>
            {
                Add(command, "$now", now); Add(command, "$client", arrangement.ClientId); Add(command, "$carrier", arrangement.Carrier); Add(command, "$effective", arrangement.EffectiveDate.ToString("yyyy-MM-dd"));
            });
        }
        arrangement.FinancialArrangementId = await InsertAsync(connection, transaction, @"INSERT INTO financial_arrangements(client_id,carrier,effective_date,expiry_date,status,source,source_document_id,target_loss_ratio,tlr_basis,target_loss_ratio_source_text,target_loss_ratio_source_page,target_loss_ratio_confidence,created_at) VALUES($client,$carrier,$effective,$expiry,$status,$source,$document,$tlr,$basis,$tlrText,$tlrPage,$tlrConfidence,$created);", command =>
        {
            Add(command, "$client", arrangement.ClientId); Add(command, "$carrier", arrangement.Carrier); Add(command, "$effective", arrangement.EffectiveDate.ToString("yyyy-MM-dd")); Add(command, "$expiry", arrangement.ExpiryDate?.ToString("yyyy-MM-dd")); Add(command, "$status", arrangement.Status); Add(command, "$source", arrangement.Source); Add(command, "$document", arrangement.SourceDocumentId); Add(command, "$tlr", arrangement.TargetLossRatio); Add(command, "$basis", arrangement.TlrBasis); Add(command, "$tlrText", arrangement.TargetLossRatioSourceText); Add(command, "$tlrPage", arrangement.TargetLossRatioSourcePage); Add(command, "$tlrConfidence", arrangement.TargetLossRatioExtractionConfidence); Add(command, "$created", now);
        });
        if (arrangement.SourceDocumentId.HasValue || arrangement.TargetLossRatioSourcePage.HasValue || !string.IsNullOrWhiteSpace(arrangement.TargetLossRatioSourceText))
        {
            await InsertAsync(connection, transaction, "INSERT INTO financial_arrangement_evidence(entity_type,entity_id,source_document_id,source_page,source_text,confidence,created_at) VALUES($type,$entity,$document,$page,$text,$confidence,$created);", command =>
            {
                Add(command, "$type", "FinancialArrangement"); Add(command, "$entity", arrangement.FinancialArrangementId); Add(command, "$document", arrangement.SourceDocumentId); Add(command, "$page", arrangement.TargetLossRatioSourcePage); Add(command, "$text", arrangement.TargetLossRatioSourceText); Add(command, "$confidence", arrangement.TargetLossRatioExtractionConfidence); Add(command, "$created", now);
            });
        }
        foreach (var charge in arrangement.Charges)
        {
            charge.FinancialArrangementChargeId = await InsertAsync(connection, transaction, @"INSERT INTO financial_arrangement_charges(financial_arrangement_id,charge_type,benefit,rate,amount,basis,formula,source_text,source_page,extraction_confidence) VALUES($arrangement,$type,$benefit,$rate,$amount,$basis,$formula,$text,$page,$confidence);", command =>
            {
                Add(command, "$arrangement", arrangement.FinancialArrangementId); Add(command, "$type", charge.ChargeType.ToString()); Add(command, "$benefit", charge.Benefit); Add(command, "$rate", charge.Rate); Add(command, "$amount", charge.Amount); Add(command, "$basis", charge.Basis); Add(command, "$formula", charge.Formula); Add(command, "$text", charge.SourceText); Add(command, "$page", charge.SourcePage); Add(command, "$confidence", charge.ExtractionConfidence);
            });
            if (arrangement.SourceDocumentId.HasValue || charge.SourcePage.HasValue || !string.IsNullOrWhiteSpace(charge.SourceText))
            {
                await InsertAsync(connection, transaction, "INSERT INTO financial_arrangement_evidence(entity_type,entity_id,source_document_id,source_page,source_text,confidence,created_at) VALUES($type,$entity,$document,$page,$text,$confidence,$created);", command =>
                {
                    Add(command, "$type", "FinancialArrangementCharge"); Add(command, "$entity", charge.FinancialArrangementChargeId); Add(command, "$document", arrangement.SourceDocumentId); Add(command, "$page", charge.SourcePage); Add(command, "$text", charge.SourceText); Add(command, "$confidence", charge.ExtractionConfidence); Add(command, "$created", now);
                });
            }
        }
        await transaction.CommitAsync();
        return arrangement;
    }

    public async Task<PoolingArrangement> SavePoolingArrangementAsync(PoolingArrangement arrangement)
    {
        await InitializeAsync(); await using var connection = await OpenAsync(); var now = DateTime.UtcNow.ToString("O");
        await ExecuteAsync(connection, null, "UPDATE pooling_arrangements SET superseded_at=$now WHERE client_id=$client AND carrier=$carrier AND benefit=$benefit AND effective_date=$effective AND superseded_at IS NULL;", command => { Add(command, "$now", now); Add(command, "$client", arrangement.ClientId); Add(command, "$carrier", arrangement.Carrier); Add(command, "$benefit", arrangement.Benefit); Add(command, "$effective", arrangement.EffectiveDate.ToString("yyyy-MM-dd")); });
        // The established desktop table retains its mandatory legacy date_loaded column.
        // Write both timestamps while the additive migration supplies created_at.
        arrangement.PoolingArrangementId = await InsertAsync(connection, null, "INSERT INTO pooling_arrangements(client_id,carrier,benefit,rate,threshold,basis,pooling_type,effective_date,expiry_date,source_document_id,source_text,source_page,extraction_confidence,created_at,date_loaded) VALUES($client,$carrier,$benefit,$rate,$threshold,$basis,$type,$effective,$expiry,$document,$text,$page,$confidence,$created,$loaded);", command => { Add(command, "$client", arrangement.ClientId); Add(command, "$carrier", arrangement.Carrier); Add(command, "$benefit", arrangement.Benefit); Add(command, "$rate", arrangement.Rate); Add(command, "$threshold", arrangement.Threshold); Add(command, "$basis", arrangement.Basis); Add(command, "$type", arrangement.PoolingType); Add(command, "$effective", arrangement.EffectiveDate.ToString("yyyy-MM-dd")); Add(command, "$expiry", arrangement.ExpiryDate?.ToString("yyyy-MM-dd")); Add(command, "$document", arrangement.SourceDocumentId); Add(command, "$text", arrangement.SourceText); Add(command, "$page", arrangement.SourcePage); Add(command, "$confidence", arrangement.ExtractionConfidence); Add(command, "$created", now); Add(command, "$loaded", now); });
        if (arrangement.SourceDocumentId.HasValue || arrangement.SourcePage.HasValue || !string.IsNullOrWhiteSpace(arrangement.SourceText))
        {
            await InsertAsync(connection, null, "INSERT INTO financial_arrangement_evidence(entity_type,entity_id,source_document_id,source_page,source_text,confidence,created_at) VALUES($type,$entity,$document,$page,$text,$confidence,$created);", command => { Add(command, "$type", "PoolingArrangement"); Add(command, "$entity", arrangement.PoolingArrangementId); Add(command, "$document", arrangement.SourceDocumentId); Add(command, "$page", arrangement.SourcePage); Add(command, "$text", arrangement.SourceText); Add(command, "$confidence", arrangement.ExtractionConfidence); Add(command, "$created", now); });
        }
        return arrangement;
    }

    public async Task<CommissionArrangement> CreateCommissionFromAorOrSoldProposalAsync(CommissionArrangement arrangement)
    {
        await InitializeAsync(); await using var connection = await OpenAsync(); var now = DateTime.UtcNow.ToString("O");
        await ExecuteAsync(connection, null, "UPDATE commission_arrangements SET superseded_at=$now WHERE client_id=$client AND carrier=$carrier AND effective_date=$effective AND superseded_at IS NULL;", command => { Add(command, "$now", now); Add(command, "$client", arrangement.ClientId); Add(command, "$carrier", arrangement.Carrier); Add(command, "$effective", arrangement.EffectiveDate.ToString("yyyy-MM-dd")); });
        arrangement.CommissionArrangementId = await InsertAsync(connection, null, "INSERT INTO commission_arrangements(client_id,carrier,rate,amount,basis,source,confirmation_status,effective_date,expiry_date,sold_proposal_id,source_document_id,created_at) VALUES($client,$carrier,$rate,$amount,$basis,$source,$status,$effective,$expiry,$proposal,$document,$created);", command => { Add(command, "$client", arrangement.ClientId); Add(command, "$carrier", arrangement.Carrier); Add(command, "$rate", arrangement.Rate); Add(command, "$amount", arrangement.Amount); Add(command, "$basis", arrangement.Basis); Add(command, "$source", arrangement.Source); Add(command, "$status", arrangement.ConfirmationStatus); Add(command, "$effective", arrangement.EffectiveDate.ToString("yyyy-MM-dd")); Add(command, "$expiry", arrangement.ExpiryDate?.ToString("yyyy-MM-dd")); Add(command, "$proposal", arrangement.SoldProposalId); Add(command, "$document", arrangement.SourceDocumentId); Add(command, "$created", now); });
        return arrangement;
    }

    public async Task<FinancialArrangement> ConfirmFinancialArrangementAsync(FinancialArrangement arrangement, FinancialArrangementAuthorizationContext authorization, string reason)
    {
        _authorization.Demand(authorization, FinancialArrangementPermissions.FinancialArrangementConfirm);
        var prior = arrangement.FinancialArrangementId;
        arrangement.Status = "Confirmed";
        var saved = await SaveFinancialArrangementAsync(arrangement);
        await WriteAuditAsync("FinancialArrangement", saved.FinancialArrangementId, saved.ClientId, null, "FinancialArrangementConfirmed", prior.ToString(), saved.FinancialArrangementId.ToString(), saved.EffectiveDate, authorization.User, reason, saved.Source);
        return saved;
    }

    public async Task<PoolingArrangement> ConfirmPoolingArrangementAsync(PoolingArrangement arrangement, FinancialArrangementAuthorizationContext authorization, string reason)
    {
        _authorization.Demand(authorization, FinancialArrangementPermissions.PoolingConfirm);
        var prior = arrangement.PoolingArrangementId;
        var saved = await SavePoolingArrangementAsync(arrangement);
        await WriteAuditAsync("PoolingArrangement", saved.PoolingArrangementId, saved.ClientId, null, "PoolingConfirmed", prior.ToString(), saved.PoolingArrangementId.ToString(), saved.EffectiveDate, authorization.User, reason, "FinancialArrangement");
        return saved;
    }

    public async Task<CommissionArrangement> ConfirmCommissionArrangementAsync(CommissionArrangement arrangement, FinancialArrangementAuthorizationContext authorization, string reason)
    {
        _authorization.Demand(authorization, FinancialArrangementPermissions.CommissionConfirm);
        var prior = arrangement.CommissionArrangementId;
        arrangement.ConfirmationStatus = "Confirmed";
        var saved = await CreateCommissionFromAorOrSoldProposalAsync(arrangement);
        await WriteAuditAsync("CommissionArrangement", saved.CommissionArrangementId, saved.ClientId, null, "CommissionConfirmed", prior.ToString(), saved.CommissionArrangementId.ToString(), saved.EffectiveDate, authorization.User, reason, saved.Source);
        return saved;
    }

    public async Task<IReadOnlyList<FinancialArrangementAuditEvent>> GetAuditEventsAsync(long clientId)
    {
        await InitializeAsync(); await using var connection = await OpenAsync(); await using var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM financial_arrangement_audit_events WHERE client_id=$client ORDER BY occurred_at_utc DESC, financial_arrangement_audit_event_id DESC;";
        Add(command, "$client", clientId);
        var events = new List<FinancialArrangementAuditEvent>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            events.Add(new FinancialArrangementAuditEvent { FinancialArrangementAuditEventId = reader.GetInt64(reader.GetOrdinal("financial_arrangement_audit_event_id")), EntityType = reader.GetString(reader.GetOrdinal("entity_type")), EntityId = reader.GetInt64(reader.GetOrdinal("entity_id")), ClientId = reader.GetInt64(reader.GetOrdinal("client_id")), RenewalId = reader.IsDBNull(reader.GetOrdinal("renewal_id")) ? null : reader.GetInt64(reader.GetOrdinal("renewal_id")), EventType = reader.GetString(reader.GetOrdinal("event_type")), PreviousValue = reader.GetString(reader.GetOrdinal("previous_value")), NewValue = reader.GetString(reader.GetOrdinal("new_value")), EffectiveDate = DateOnly.Parse(reader.GetString(reader.GetOrdinal("effective_date"))), User = reader.GetString(reader.GetOrdinal("user_name")), Reason = reader.GetString(reader.GetOrdinal("reason")), Source = reader.GetString(reader.GetOrdinal("source")), OccurredAtUtc = DateTime.Parse(reader.GetString(reader.GetOrdinal("occurred_at_utc"))).ToUniversalTime() });
        }
        return events;
    }

    public async Task<FinancialArrangementVersion> CreateVersionAsync(FinancialArrangementVersion version)
    {
        await InitializeAsync(); await using var connection = await OpenAsync();
        version.FinancialArrangementVersionId = await InsertAsync(connection, null, "INSERT OR IGNORE INTO financial_arrangement_versions(client_id,carrier,scenario,effective_date,financial_arrangement_id,commission_arrangement_id,pooling_arrangement_id,created_at) VALUES($client,$carrier,$scenario,$effective,$financial,$commission,$pooling,$created);", command => { Add(command, "$client", version.ClientId); Add(command, "$carrier", version.Carrier); Add(command, "$scenario", version.Scenario.ToString()); Add(command, "$effective", version.EffectiveDate.ToString("yyyy-MM-dd")); Add(command, "$financial", version.FinancialArrangementId); Add(command, "$commission", version.CommissionArrangementId); Add(command, "$pooling", version.PoolingArrangementId); Add(command, "$created", DateTime.UtcNow.ToString("O")); });
        if (version.FinancialArrangementVersionId == 0) version.FinancialArrangementVersionId = await ScalarLongAsync(connection, "SELECT financial_arrangement_version_id FROM financial_arrangement_versions WHERE client_id=$client AND carrier=$carrier AND scenario=$scenario AND effective_date=$effective AND financial_arrangement_id=$financial AND IFNULL(commission_arrangement_id,0)=IFNULL($commission,0) AND IFNULL(pooling_arrangement_id,0)=IFNULL($pooling,0);", command => { Add(command, "$client", version.ClientId); Add(command, "$carrier", version.Carrier); Add(command, "$scenario", version.Scenario.ToString()); Add(command, "$effective", version.EffectiveDate.ToString("yyyy-MM-dd")); Add(command, "$financial", version.FinancialArrangementId); Add(command, "$commission", version.CommissionArrangementId); Add(command, "$pooling", version.PoolingArrangementId); });
        return version;
    }

    public async Task<FinancialArrangementVersion?> GetLatestVersionAsync(long clientId, string carrier, FinancialArrangementScenario scenario)
    {
        await InitializeAsync(); await using var connection = await OpenAsync(); await using var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM financial_arrangement_versions WHERE client_id=$client AND carrier=$carrier AND scenario=$scenario ORDER BY effective_date DESC, financial_arrangement_version_id DESC LIMIT 1;";
        Add(command, "$client", clientId); Add(command, "$carrier", carrier); Add(command, "$scenario", scenario.ToString());
        await using var reader = await command.ExecuteReaderAsync();
        return await reader.ReadAsync() ? new FinancialArrangementVersion { FinancialArrangementVersionId = reader.GetInt64(reader.GetOrdinal("financial_arrangement_version_id")), ClientId = reader.GetInt64(reader.GetOrdinal("client_id")), Carrier = reader.GetString(reader.GetOrdinal("carrier")), Scenario = Enum.Parse<FinancialArrangementScenario>(reader.GetString(reader.GetOrdinal("scenario"))), EffectiveDate = DateOnly.Parse(reader.GetString(reader.GetOrdinal("effective_date"))), FinancialArrangementId = reader.GetInt64(reader.GetOrdinal("financial_arrangement_id")), CommissionArrangementId = reader.IsDBNull(reader.GetOrdinal("commission_arrangement_id")) ? null : reader.GetInt64(reader.GetOrdinal("commission_arrangement_id")), PoolingArrangementId = reader.IsDBNull(reader.GetOrdinal("pooling_arrangement_id")) ? null : reader.GetInt64(reader.GetOrdinal("pooling_arrangement_id")) } : null;
    }

    public async Task<IReadOnlyList<FinancialArrangementHistoryVersion>> GetHistoryVersionsForClientAsync(long clientId)
    {
        await InitializeAsync();
        await using var connection = await OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT version.financial_arrangement_version_id,
                   version.client_id,
                   version.carrier,
                   version.scenario,
                   version.effective_date,
                   version.financial_arrangement_id,
                   version.created_at AS version_created_at,
                   arrangement.status AS arrangement_status,
                   arrangement.source,
                   arrangement.source_document_id,
                   arrangement.target_loss_ratio,
                   arrangement.tlr_basis,
                   arrangement.target_loss_ratio_source_text,
                   arrangement.target_loss_ratio_source_page,
                   arrangement.target_loss_ratio_confidence,
                   version.commission_arrangement_id,
                   commission.confirmation_status AS commission_confirmation_status,
                   commission.rate AS commission_rate,
                   commission.amount AS commission_amount,
                   commission.basis AS commission_basis,
                   commission.source AS commission_source,
                   commission.source_document_id AS commission_source_document_id,
                   version.pooling_arrangement_id,
                   pooling.benefit AS pooling_benefit,
                   pooling.rate AS pooling_rate,
                   pooling.threshold AS pooling_threshold,
                   pooling.basis AS pooling_basis,
                   pooling.pooling_type,
                   pooling.source_document_id AS pooling_source_document_id,
                   pooling.source_text AS pooling_source_text,
                   pooling.source_page AS pooling_source_page,
                   pooling.extraction_confidence AS pooling_confidence
            FROM financial_arrangement_versions AS version
            JOIN financial_arrangements AS arrangement
              ON arrangement.financial_arrangement_id = version.financial_arrangement_id
            LEFT JOIN commission_arrangements AS commission
              ON commission.commission_arrangement_id = version.commission_arrangement_id
            LEFT JOIN pooling_arrangements AS pooling
              ON pooling.pooling_arrangement_id = version.pooling_arrangement_id
            WHERE version.client_id = $client
            ORDER BY version.effective_date, version.financial_arrangement_version_id;
            """;
        Add(command, "$client", clientId);

        var versions = new List<FinancialArrangementHistoryVersion>();
        await using (var reader = await command.ExecuteReaderAsync())
        {
            while (await reader.ReadAsync())
            {
                versions.Add(new FinancialArrangementHistoryVersion
                {
                    FinancialArrangementVersionId = reader.GetInt64(reader.GetOrdinal("financial_arrangement_version_id")),
                    ClientId = reader.GetInt64(reader.GetOrdinal("client_id")),
                    Carrier = ReadString(reader, "carrier"),
                    Scenario = Enum.Parse<FinancialArrangementScenario>(ReadString(reader, "scenario")),
                    EffectiveDate = DateOnly.Parse(ReadString(reader, "effective_date")),
                    FinancialArrangementId = reader.GetInt64(reader.GetOrdinal("financial_arrangement_id")),
                    CreatedAtUtc = DateTime.Parse(ReadString(reader, "version_created_at")).ToUniversalTime(),
                    ArrangementStatus = ReadString(reader, "arrangement_status"),
                    Source = ReadString(reader, "source"),
                    SourceDocumentId = ReadNullableLong(reader, "source_document_id"),
                    TargetLossRatio = ReadNullableDecimal(reader, "target_loss_ratio"),
                    TargetLossRatioBasis = ReadString(reader, "tlr_basis"),
                    TargetLossRatioSourceText = ReadString(reader, "target_loss_ratio_source_text"),
                    TargetLossRatioSourcePage = ReadNullableInt(reader, "target_loss_ratio_source_page"),
                    TargetLossRatioConfidence = ReadNullableDecimal(reader, "target_loss_ratio_confidence"),
                    CommissionArrangementId = ReadNullableLong(reader, "commission_arrangement_id"),
                    CommissionConfirmationStatus = ReadString(reader, "commission_confirmation_status"),
                    CommissionRate = ReadNullableDecimal(reader, "commission_rate"),
                    CommissionAmount = ReadNullableDecimal(reader, "commission_amount"),
                    CommissionBasis = ReadString(reader, "commission_basis"),
                    CommissionSource = ReadString(reader, "commission_source"),
                    CommissionSourceDocumentId = ReadNullableLong(reader, "commission_source_document_id"),
                    PoolingArrangementId = ReadNullableLong(reader, "pooling_arrangement_id"),
                    PoolingBenefit = ReadString(reader, "pooling_benefit"),
                    PoolingRate = ReadNullableDecimal(reader, "pooling_rate"),
                    PoolingThreshold = ReadNullableDecimal(reader, "pooling_threshold"),
                    PoolingBasis = ReadString(reader, "pooling_basis"),
                    PoolingType = ReadString(reader, "pooling_type"),
                    PoolingSourceDocumentId = ReadNullableLong(reader, "pooling_source_document_id"),
                    PoolingSourceText = ReadString(reader, "pooling_source_text"),
                    PoolingSourcePage = ReadNullableInt(reader, "pooling_source_page"),
                    PoolingConfidence = ReadNullableDecimal(reader, "pooling_confidence")
                });
            }
        }

        var versionsById = versions.ToDictionary(version => version.FinancialArrangementVersionId);
        await using var chargesCommand = connection.CreateCommand();
        chargesCommand.CommandText = """
            SELECT version.financial_arrangement_version_id,
                   charge.financial_arrangement_charge_id,
                   charge.charge_type,
                   charge.benefit,
                   charge.rate,
                   charge.amount,
                   charge.basis,
                   charge.formula,
                   charge.source_text,
                   charge.source_page,
                   charge.extraction_confidence
            FROM financial_arrangement_versions AS version
            JOIN financial_arrangement_charges AS charge
              ON charge.financial_arrangement_id = version.financial_arrangement_id
            WHERE version.client_id = $client
            ORDER BY version.financial_arrangement_version_id,
                     charge.financial_arrangement_charge_id;
            """;
        Add(chargesCommand, "$client", clientId);
        await using var chargesReader = await chargesCommand.ExecuteReaderAsync();
        while (await chargesReader.ReadAsync())
        {
            var versionId = chargesReader.GetInt64(chargesReader.GetOrdinal("financial_arrangement_version_id"));
            if (!versionsById.TryGetValue(versionId, out var version) ||
                !Enum.TryParse<FinancialChargeType>(ReadString(chargesReader, "charge_type"), out var chargeType))
            {
                continue;
            }

            version.Charges.Add(new FinancialArrangementCharge
            {
                FinancialArrangementChargeId = chargesReader.GetInt64(chargesReader.GetOrdinal("financial_arrangement_charge_id")),
                ChargeType = chargeType,
                Benefit = ReadString(chargesReader, "benefit"),
                Rate = ReadNullableDecimal(chargesReader, "rate"),
                Amount = ReadNullableDecimal(chargesReader, "amount"),
                Basis = ReadString(chargesReader, "basis"),
                Formula = ReadString(chargesReader, "formula"),
                SourceText = ReadString(chargesReader, "source_text"),
                SourcePage = ReadNullableInt(chargesReader, "source_page"),
                ExtractionConfidence = ReadNullableDecimal(chargesReader, "extraction_confidence")
            });
        }

        return versions;
    }

    /// <summary>
    /// Stores extracted evidence that is not complete enough to alter the verified client view.
    /// It intentionally does not supersede an existing confirmed arrangement or create a version.
    /// </summary>
    public async Task<FinancialArrangement> SaveFinancialArrangementForReviewAsync(FinancialArrangement arrangement, string reason)
    {
        arrangement.Status = "NeedsReview";
        var saved = await SaveFinancialArrangementAsync(arrangement, supersedeExisting: false);
        await WriteAuditAsync("FinancialArrangement", saved.FinancialArrangementId, saved.ClientId, null, "FinancialArrangementNeedsReview", "", reason, saved.EffectiveDate, "System", reason, saved.Source);
        return saved;
    }

    /// <summary>
    /// Returns only the latest confirmed client-specific financial version. It never falls back to
    /// renewal-queue estimates, carrier defaults, or ASO activity performance data.
    /// </summary>
    public async Task<CurrentClientFinancialTerms?> GetCurrentClientTermsAsync(long clientId, string carrier)
    {
        if (clientId <= 0 || string.IsNullOrWhiteSpace(carrier)) return null;

        await InitializeAsync();
        await using var connection = await OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT
    version.financial_arrangement_version_id,
    version.scenario,
    arrangement.client_id,
    arrangement.carrier,
    arrangement.effective_date,
    arrangement.source,
    arrangement.source_document_id,
    arrangement.target_loss_ratio,
    arrangement.tlr_basis,
    arrangement.target_loss_ratio_source_text,
    arrangement.target_loss_ratio_source_page,
    arrangement.target_loss_ratio_confidence,
    pooling.rate AS pooling_rate,
    pooling.threshold AS pooling_threshold,
    pooling.basis AS pooling_basis,
    pooling.source_text AS pooling_source_text,
    pooling.source_page AS pooling_source_page,
    pooling.extraction_confidence AS pooling_extraction_confidence
FROM financial_arrangement_versions AS version
INNER JOIN financial_arrangements AS arrangement
    ON arrangement.financial_arrangement_id = version.financial_arrangement_id
LEFT JOIN pooling_arrangements AS pooling
    ON pooling.pooling_arrangement_id = version.pooling_arrangement_id
WHERE version.client_id = $client
  AND lower(version.carrier) = lower($carrier)
  AND arrangement.status = 'Confirmed'
  AND arrangement.superseded_at IS NULL
  AND version.scenario IN ('Current', 'FinalAccepted')
ORDER BY
    arrangement.effective_date DESC,
    CASE version.scenario WHEN 'Current' THEN 0 WHEN 'FinalAccepted' THEN 1 ELSE 2 END,
    version.financial_arrangement_version_id DESC
LIMIT 1;";
        Add(command, "$client", clientId);
        Add(command, "$carrier", carrier.Trim());

        await using var reader = await command.ExecuteReaderAsync();
        if (!await reader.ReadAsync()) return null;

        return new CurrentClientFinancialTerms
        {
            FinancialArrangementVersionId = reader.GetInt64(reader.GetOrdinal("financial_arrangement_version_id")),
            Scenario = Enum.Parse<FinancialArrangementScenario>(reader.GetString(reader.GetOrdinal("scenario"))),
            ClientId = reader.GetInt64(reader.GetOrdinal("client_id")),
            Carrier = ReadString(reader, "carrier"),
            EffectiveDate = DateOnly.Parse(ReadString(reader, "effective_date")),
            Source = ReadString(reader, "source"),
            SourceDocumentId = ReadNullableLong(reader, "source_document_id"),
            TargetLossRatio = ReadNullableDecimal(reader, "target_loss_ratio"),
            TlrBasis = ReadString(reader, "tlr_basis"),
            TargetLossRatioSourceText = ReadString(reader, "target_loss_ratio_source_text"),
            TargetLossRatioSourcePage = ReadNullableInt(reader, "target_loss_ratio_source_page"),
            TargetLossRatioExtractionConfidence = ReadNullableDecimal(reader, "target_loss_ratio_confidence"),
            PoolingRate = ReadNullableDecimal(reader, "pooling_rate"),
            PoolingThreshold = ReadNullableDecimal(reader, "pooling_threshold"),
            PoolingBasis = ReadString(reader, "pooling_basis"),
            PoolingSourceText = ReadString(reader, "pooling_source_text"),
            PoolingSourcePage = ReadNullableInt(reader, "pooling_source_page"),
            PoolingExtractionConfidence = ReadNullableDecimal(reader, "pooling_extraction_confidence")
        };
    }

    public async Task<RenewalFinancialArrangementSnapshot> CreateSnapshotAsync(RenewalFinancialArrangementSnapshot snapshot)
    {
        await InitializeAsync(); await using var connection = await OpenAsync();
        snapshot.SourceVersion = $"{snapshot.CurrentArrangementVersionId}:{snapshot.ProposedArrangementVersionId}";
        snapshot.RenewalFinancialArrangementSnapshotId = await InsertAsync(connection, null, "INSERT OR IGNORE INTO renewal_financial_arrangement_snapshots(renewal_id,client_id,carrier,scenario,snapshot_type,snapshot_status,as_of_utc,current_arrangement_version_id,proposed_arrangement_version_id,source_version,is_stale,created_at) VALUES($renewal,$client,$carrier,$scenario,$type,$status,$asof,$current,$proposed,$source,0,$created);", command => { Add(command, "$renewal", snapshot.RenewalId); Add(command, "$client", snapshot.ClientId); Add(command, "$carrier", snapshot.Carrier); Add(command, "$scenario", snapshot.Scenario.ToString()); Add(command, "$type", snapshot.SnapshotType); Add(command, "$status", snapshot.SnapshotStatus); Add(command, "$asof", snapshot.AsOfUtc.ToString("O")); Add(command, "$current", snapshot.CurrentArrangementVersionId); Add(command, "$proposed", snapshot.ProposedArrangementVersionId); Add(command, "$source", snapshot.SourceVersion); Add(command, "$created", DateTime.UtcNow.ToString("O")); });
        if (snapshot.RenewalFinancialArrangementSnapshotId == 0) snapshot.RenewalFinancialArrangementSnapshotId = await ScalarLongAsync(connection, "SELECT renewal_financial_arrangement_snapshot_id FROM renewal_financial_arrangement_snapshots WHERE renewal_id=$renewal AND snapshot_type=$type AND source_version=$source;", command => { Add(command, "$renewal", snapshot.RenewalId); Add(command, "$type", snapshot.SnapshotType); Add(command, "$source", snapshot.SourceVersion); });
        return snapshot;
    }

    public async Task<int> MarkOpenSnapshotsStaleAsync(long clientId, string carrier, string reason)
    {
        await InitializeAsync(); await using var connection = await OpenAsync();
        var changed = await ExecuteCountAsync(connection, "UPDATE renewal_financial_arrangement_snapshots SET is_stale=1, snapshot_status='NeedsReview' WHERE client_id=$client AND carrier=$carrier AND snapshot_status NOT IN ('Finalized','Superseded') AND is_stale=0;", command => { Add(command, "$client", clientId); Add(command, "$carrier", carrier); });
        if (changed > 0) await WriteAuditAsync("RenewalFinancialArrangementSnapshot", 0, clientId, null, "RenewalSnapshotMarkedStale", "", reason, DateOnly.FromDateTime(DateTime.UtcNow), "System", reason, "Canonical financial term changed");
        return changed;
    }

    public async Task<RenewalFinancialArrangementSnapshot> RefreshSnapshotAsync(long snapshotId, long currentVersionId, long proposedVersionId, FinancialArrangementAuthorizationContext authorization)
    {
        _authorization.Demand(authorization, FinancialArrangementPermissions.RenewalFinancialComparisonRefresh);
        var existing = await GetSnapshotAsync(snapshotId) ?? throw new InvalidOperationException("The renewal financial snapshot was not found.");
        if (existing.SnapshotStatus.Equals("Finalized", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("A finalized renewal financial snapshot cannot be refreshed. Create a new renewal scenario instead.");
        if (existing.SnapshotType.Equals("Refreshed", StringComparison.OrdinalIgnoreCase) && existing.SourceVersion == $"{currentVersionId}:{proposedVersionId}" && !existing.IsStale) return existing;
        await InitializeAsync(); await using var connection = await OpenAsync();
        await ExecuteAsync(connection, null, "UPDATE renewal_financial_arrangement_snapshots SET snapshot_status='Superseded' WHERE renewal_financial_arrangement_snapshot_id=$snapshot;", command => Add(command, "$snapshot", snapshotId));
        var refreshed = await CreateSnapshotAsync(new RenewalFinancialArrangementSnapshot { RenewalId = existing.RenewalId, ClientId = existing.ClientId, Carrier = existing.Carrier, Scenario = existing.Scenario, SnapshotType = "Refreshed", SnapshotStatus = "Draft", AsOfUtc = DateTime.UtcNow, CurrentArrangementVersionId = currentVersionId, ProposedArrangementVersionId = proposedVersionId });
        await WriteAuditAsync("RenewalFinancialArrangementSnapshot", refreshed.RenewalFinancialArrangementSnapshotId, refreshed.ClientId, refreshed.RenewalId, "RenewalSnapshotRefreshed", snapshotId.ToString(), refreshed.RenewalFinancialArrangementSnapshotId.ToString(), DateOnly.FromDateTime(DateTime.UtcNow), authorization.User, "Canonical financial terms refreshed", "Renewal comparison");
        return refreshed;
    }

    public async Task FinalizeSnapshotAsync(long snapshotId, FinancialArrangementAuthorizationContext authorization)
    {
        _authorization.Demand(authorization, FinancialArrangementPermissions.RenewalFinancialComparisonFinalize);
        await InitializeAsync(); await using var connection = await OpenAsync();
        var now = DateTime.UtcNow;
        var changed = await ExecuteCountAsync(connection, "UPDATE renewal_financial_arrangement_snapshots SET snapshot_status='Finalized', is_stale=0, finalized_at=$finalized, finalized_by=$user WHERE renewal_financial_arrangement_snapshot_id=$snapshot AND snapshot_status NOT IN ('Finalized','Superseded');", command => { Add(command, "$finalized", now.ToString("O")); Add(command, "$user", authorization.User); Add(command, "$snapshot", snapshotId); });
        if (changed != 1) throw new InvalidOperationException("Only an open renewal financial snapshot can be finalized.");
        var snapshot = await GetSnapshotAsync(snapshotId) ?? throw new InvalidOperationException("The renewal financial snapshot was not found.");
        await WriteAuditAsync("RenewalFinancialArrangementSnapshot", snapshotId, snapshot.ClientId, snapshot.RenewalId, "RenewalSnapshotFinalized", "", "Finalized", DateOnly.FromDateTime(now), authorization.User, "Renewal terms accepted", "Renewal comparison");
    }

    public async Task<IReadOnlyList<RenewalFinancialArrangementSnapshot>> GetSnapshotsAsync(long renewalId)
    {
        await InitializeAsync(); await using var connection = await OpenAsync(); await using var command = connection.CreateCommand(); command.CommandText = "SELECT * FROM renewal_financial_arrangement_snapshots WHERE renewal_id=$renewal ORDER BY created_at, renewal_financial_arrangement_snapshot_id;"; Add(command, "$renewal", renewalId);
        var snapshots = new List<RenewalFinancialArrangementSnapshot>(); await using var reader = await command.ExecuteReaderAsync(); while (await reader.ReadAsync()) snapshots.Add(ReadSnapshot(reader));
        return snapshots;
    }

    public async Task<IReadOnlyList<RenewalFinancialArrangementSnapshot>> GetSnapshotsForClientAsync(long clientId)
    {
        await InitializeAsync(); await using var connection = await OpenAsync(); await using var command = connection.CreateCommand(); command.CommandText = "SELECT * FROM renewal_financial_arrangement_snapshots WHERE client_id=$client ORDER BY created_at DESC, renewal_financial_arrangement_snapshot_id DESC;"; Add(command, "$client", clientId);
        var snapshots = new List<RenewalFinancialArrangementSnapshot>(); await using var reader = await command.ExecuteReaderAsync(); while (await reader.ReadAsync()) snapshots.Add(ReadSnapshot(reader));
        return snapshots;
    }

    private async Task<RenewalFinancialArrangementSnapshot?> GetSnapshotAsync(long snapshotId)
    {
        await InitializeAsync(); await using var connection = await OpenAsync(); await using var command = connection.CreateCommand(); command.CommandText = "SELECT * FROM renewal_financial_arrangement_snapshots WHERE renewal_financial_arrangement_snapshot_id=$snapshot;"; Add(command, "$snapshot", snapshotId); await using var reader = await command.ExecuteReaderAsync(); return await reader.ReadAsync() ? ReadSnapshot(reader) : null;
    }

    private async Task<SqliteConnection> OpenAsync() { var connection = new SqliteConnection($"Data Source={_databasePath}"); await connection.OpenAsync(); return connection; }
    private static string ReadString(SqliteDataReader reader, string column)
    {
        var ordinal = reader.GetOrdinal(column);
        return reader.IsDBNull(ordinal) ? "" : reader.GetString(ordinal);
    }

    private static decimal? ReadNullableDecimal(SqliteDataReader reader, string column)
    {
        var ordinal = reader.GetOrdinal(column);
        return reader.IsDBNull(ordinal) ? null : Convert.ToDecimal(reader.GetValue(ordinal));
    }

    private static long? ReadNullableLong(SqliteDataReader reader, string column)
    {
        var ordinal = reader.GetOrdinal(column);
        return reader.IsDBNull(ordinal) ? null : reader.GetInt64(ordinal);
    }

    private static int? ReadNullableInt(SqliteDataReader reader, string column)
    {
        var ordinal = reader.GetOrdinal(column);
        return reader.IsDBNull(ordinal) ? null : reader.GetInt32(ordinal);
    }

    private static RenewalFinancialArrangementSnapshot ReadSnapshot(SqliteDataReader reader) => new() { RenewalFinancialArrangementSnapshotId = reader.GetInt64(reader.GetOrdinal("renewal_financial_arrangement_snapshot_id")), RenewalId = reader.GetInt64(reader.GetOrdinal("renewal_id")), ClientId = reader.GetInt64(reader.GetOrdinal("client_id")), Carrier = reader.GetString(reader.GetOrdinal("carrier")), Scenario = Enum.Parse<FinancialArrangementScenario>(reader.GetString(reader.GetOrdinal("scenario"))), SnapshotType = reader.GetString(reader.GetOrdinal("snapshot_type")), SnapshotStatus = reader.GetString(reader.GetOrdinal("snapshot_status")), AsOfUtc = DateTime.Parse(reader.GetString(reader.GetOrdinal("as_of_utc"))).ToUniversalTime(), CurrentArrangementVersionId = reader.GetInt64(reader.GetOrdinal("current_arrangement_version_id")), ProposedArrangementVersionId = reader.GetInt64(reader.GetOrdinal("proposed_arrangement_version_id")), SourceVersion = reader.GetString(reader.GetOrdinal("source_version")), IsStale = reader.GetInt64(reader.GetOrdinal("is_stale")) != 0, FinalizedBy = reader.IsDBNull(reader.GetOrdinal("finalized_by")) ? "" : reader.GetString(reader.GetOrdinal("finalized_by")), FinalizedAtUtc = reader.IsDBNull(reader.GetOrdinal("finalized_at")) ? null : DateTime.Parse(reader.GetString(reader.GetOrdinal("finalized_at"))).ToUniversalTime() };
    private async Task WriteAuditAsync(string entityType, long entityId, long clientId, long? renewalId, string eventType, string previousValue, string newValue, DateOnly effectiveDate, string user, string reason, string source)
    {
        await InitializeAsync(); await using var connection = await OpenAsync();
        await InsertAsync(connection, null, "INSERT INTO financial_arrangement_audit_events(entity_type,entity_id,client_id,renewal_id,event_type,previous_value,new_value,effective_date,user_name,reason,source,occurred_at_utc) VALUES($type,$entity,$client,$renewal,$event,$previous,$new,$effective,$user,$reason,$source,$occurred);", command => { Add(command, "$type", entityType); Add(command, "$entity", entityId); Add(command, "$client", clientId); Add(command, "$renewal", renewalId); Add(command, "$event", eventType); Add(command, "$previous", previousValue); Add(command, "$new", newValue); Add(command, "$effective", effectiveDate.ToString("yyyy-MM-dd")); Add(command, "$user", user); Add(command, "$reason", reason); Add(command, "$source", source); Add(command, "$occurred", DateTime.UtcNow.ToString("O")); });
    }
    private static async Task ExecuteAsync(SqliteConnection c, SqliteTransaction? t, string sql, Action<SqliteCommand> configure) { await using var cmd = c.CreateCommand(); cmd.Transaction = t; cmd.CommandText = sql; configure(cmd); await cmd.ExecuteNonQueryAsync(); }
    private static async Task<int> ExecuteCountAsync(SqliteConnection c, string sql, Action<SqliteCommand> configure) { await using var cmd = c.CreateCommand(); cmd.CommandText = sql; configure(cmd); return await cmd.ExecuteNonQueryAsync(); }
    private static async Task<long> InsertAsync(SqliteConnection c, SqliteTransaction? t, string sql, Action<SqliteCommand> configure) { await using var cmd = c.CreateCommand(); cmd.Transaction = t; cmd.CommandText = sql + " SELECT last_insert_rowid();"; configure(cmd); return Convert.ToInt64(await cmd.ExecuteScalarAsync()); }
    private static async Task<long> ScalarLongAsync(SqliteConnection c, string sql, Action<SqliteCommand> configure) { await using var cmd = c.CreateCommand(); cmd.CommandText = sql; configure(cmd); return Convert.ToInt64(await cmd.ExecuteScalarAsync()); }
    private static void Add(SqliteCommand command, string name, object? value) => command.Parameters.AddWithValue(name, value ?? DBNull.Value);
}
