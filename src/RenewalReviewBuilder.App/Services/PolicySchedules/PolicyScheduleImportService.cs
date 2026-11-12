using RenewalReviewBuilder.App.Connectors.Pdf;
using RenewalReviewBuilder.App.Connectors;
using RenewalReviewBuilder.App.Models;
using System.Threading;

namespace RenewalReviewBuilder.App.Services.PolicySchedules;

/// <summary>Thin coordinator so batch processing stays free of extraction, mapping, and persistence details.</summary>
public sealed class PolicyScheduleImportService
{
    private readonly DatabaseService _database;
    private readonly ManulifePolicyScheduleMapper _mapper;

    public PolicyScheduleImportService(DatabaseService database, ManulifePolicyScheduleMapper? mapper = null)
    {
        _database = database;
        _mapper = mapper ?? new ManulifePolicyScheduleMapper();
    }

    public async Task<DatabaseSaveResult> ImportAsync(ClientRecord client, IPolicyScheduleConnector connector, string sourcePath, CancellationToken cancellationToken = default)
    {
        var extraction = await connector.ExtractPolicyScheduleAsync(new BenefitsConnectorRequest { InputPath = sourcePath }, cancellationToken);
        var mapped = _mapper.Map(extraction, client);
        return await _database.SavePlanDesignAsync(
            client,
            [sourcePath],
            mapped.Rows,
            documentType: DocumentTypes.GroupPolicySchedule,
            extractionMethod: extraction.ExtractorVersion,
            planDesignSource: PlanDesignSources.PolicySchedule,
            validationIssues: mapped.ValidationIssues);
    }
}
