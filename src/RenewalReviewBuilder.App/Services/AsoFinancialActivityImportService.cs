using System.IO;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class AsoFinancialActivityImportRequest
{
    public string InputPath { get; init; } = "";
    public long? SelectedClientId { get; init; }
    public string SelectedCarrier { get; init; } = "";
    public string ImportedBy { get; init; } = "";
}

public sealed class AsoFinancialActivityImportResult
{
    public AsoFinancialActivityStatement? Statement { get; init; }
    public AsoFinancialActivitySaveResult? Persistence { get; init; }
    public bool Succeeded => Persistence?.Succeeded == true;
    public string Message { get; init; } = "";
}

public sealed class AsoFinancialActivityImportService
{
    private readonly DatabaseService _databaseService;
    private readonly AsoFinancialActivityWorkbookParser _parser;

    public AsoFinancialActivityImportService(
        DatabaseService databaseService,
        AsoFinancialActivityWorkbookParser? parser = null)
    {
        _databaseService = databaseService;
        _parser = parser ?? new AsoFinancialActivityWorkbookParser();
    }

    public async Task<AsoFinancialActivityImportResult> ImportAsync(
        AsoFinancialActivityImportRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.InputPath) ||
            !File.Exists(request.InputPath) ||
            !IsSupportedWorkbook(request.InputPath))
        {
            return new AsoFinancialActivityImportResult
            {
                Message = "Select a readable ASO Activity .xlsx or .xlsm workbook before importing."
            };
        }

        try
        {
            var statement = await Task.Run(() => _parser.Parse(request.InputPath));
            if (string.IsNullOrWhiteSpace(statement.Carrier))
            {
                statement.Carrier = request.SelectedCarrier;
            }

            var persistence = await _databaseService.SaveAsoFinancialActivityAsync(
                statement,
                string.IsNullOrWhiteSpace(request.ImportedBy)
                    ? Environment.UserName
                    : request.ImportedBy,
                request.SelectedClientId);

            return new AsoFinancialActivityImportResult
            {
                Statement = statement,
                Persistence = persistence,
                Message = persistence.Message
            };
        }
        catch (Exception exception)
        {
            AppLogger.Error(exception);
            return new AsoFinancialActivityImportResult
            {
                Message = $"The ASO Activity workbook could not be imported. {exception.Message}"
            };
        }
    }

    private static bool IsSupportedWorkbook(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
    }
}
