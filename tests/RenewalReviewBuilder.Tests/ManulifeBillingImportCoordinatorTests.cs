using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ManulifeBillingImportCoordinatorTests
{
    [Fact]
    public async Task Readable_Valid_Billing_Completes_Only_After_Persistence()
    {
        var input = CreateInputFile();
        var extraction = ManulifeBillingTestData.CreateValidExtraction(input);
        var persistence = new FakePersistence
        {
            Result = new ManulifeBillingPersistenceResult
            {
                Succeeded = true,
                Outcome = ManulifeBillingSaveOutcomes.Saved,
                Message = "Saved",
                DocumentId = 10,
                ImportBatchId = 11,
                StatementVersionId = 12,
                SavedRowCount = 9
            }
        };
        var csv = new FakeCsvExporter();
        var coordinator = new ManulifeBillingImportCoordinator(persistence, new FakeParser(extraction), csv);

        var result = await coordinator.ImportAsync(new ManulifeBillingImportRequest
        {
            InputPath = input,
            OutputFolder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"))
        });

        Assert.True(result.Succeeded);
        Assert.Equal(BillingImportStatuses.Completed, result.FinalStatus);
        Assert.True(persistence.WasCalled);
        Assert.True(csv.WasCalled);
        Assert.Equal(BillingExportStatuses.Completed, result.ExportStatus);
    }

    [Fact]
    public async Task Unreadable_Pdf_Cannot_Complete_Or_Save()
    {
        var input = CreateInputFile();
        var persistence = new FakePersistence();
        var csv = new FakeCsvExporter();
        var coordinator = new ManulifeBillingImportCoordinator(
            persistence,
            new ThrowingParser(new InvalidDataException("Unreadable PDF")),
            csv);

        var result = await coordinator.ImportAsync(new ManulifeBillingImportRequest { InputPath = input });

        Assert.False(result.Succeeded);
        Assert.Equal(BillingImportStatuses.Failed, result.FinalStatus);
        Assert.False(persistence.WasCalled);
        Assert.False(csv.WasCalled);
    }

    [Fact]
    public void Zero_Detail_Rows_Are_Blocking()
    {
        var extraction = ManulifeBillingTestData.CreateValidExtraction(CreateInputFile());
        extraction.Details.Clear();

        var validation = new ManulifeBillingCompletionPolicy().Evaluate(extraction);

        Assert.False(validation.CanPersist);
        Assert.Equal(BillingImportStatuses.Failed, validation.RecommendedStatus);
        Assert.Contains(validation.Issues, issue => issue.Code == "NoBillingDetails" && issue.IsBlocking);
    }

    [Fact]
    public void Client_Carrier_And_Billing_Period_Must_Be_Confirmed()
    {
        var source = CreateInputFile();
        var missingClient = ManulifeBillingTestData.CreateValidExtraction(source);
        missingClient.ClientIdentityConfirmed = false;
        var missingCarrier = ManulifeBillingTestData.CreateValidExtraction(source);
        missingCarrier.CarrierIdentityConfirmed = false;
        var missingPeriod = ManulifeBillingTestData.CreateValidExtraction(source);
        missingPeriod.PeriodStart = null;

        var policy = new ManulifeBillingCompletionPolicy();
        Assert.False(policy.Evaluate(missingClient).CanPersist);
        Assert.False(policy.Evaluate(missingCarrier).CanPersist);
        Assert.False(policy.Evaluate(missingPeriod).CanPersist);
    }

    [Fact]
    public void Missing_Statement_Totals_Cannot_Complete()
    {
        var extraction = ManulifeBillingTestData.CreateValidExtraction(CreateInputFile());
        extraction.PdfGrandTotalWithTax = null;

        var validation = new ManulifeBillingCompletionPolicy().Evaluate(extraction);

        Assert.True(validation.CanPersist);
        Assert.False(validation.CanComplete);
        Assert.Equal(BillingImportStatuses.NeedsReview, validation.RecommendedStatus);
    }

    [Fact]
    public void Required_NotAvailable_Reconciliation_Does_Not_Count_As_Passed()
    {
        var extraction = ManulifeBillingTestData.CreateValidExtraction(CreateInputFile());
        extraction.ReconciliationResults.Single(row => row.ReconciliationLevel == "Grand total").Status =
            BillingReconciliationStatuses.NotAvailable;

        var validation = new ManulifeBillingCompletionPolicy().Evaluate(extraction);

        Assert.False(validation.CanComplete);
        Assert.Equal(BillingImportStatuses.NeedsReview, validation.RecommendedStatus);
    }

    [Fact]
    public void Lives_Reconciliation_Is_Independent_And_Material_Mismatch_Requires_Review()
    {
        var exact = ManulifeBillingLivesReconciliation.Evaluate(100, 100);
        var warning = ManulifeBillingLivesReconciliation.Evaluate(100, 101);
        var failed = ManulifeBillingLivesReconciliation.Evaluate(100, 103);
        var unavailable = ManulifeBillingLivesReconciliation.Evaluate(null, 100);

        Assert.Equal(BillingReconciliationStatuses.Passed, exact.Status);
        Assert.Equal(BillingReconciliationStatuses.Warning, warning.Status);
        Assert.Equal(BillingReconciliationStatuses.Failed, failed.Status);
        Assert.Equal(3, failed.Difference);
        Assert.Equal(BillingReconciliationStatuses.NotAvailable, unavailable.Status);

        var extraction = ManulifeBillingTestData.CreateValidExtraction(CreateInputFile());
        var grandTotal = extraction.ReconciliationResults.Single(row => row.ReconciliationLevel == "Grand total");
        grandTotal.PdfLives = 100;
        grandTotal.CalculatedLives = 103;
        grandTotal.LivesDifference = failed.Difference;
        grandTotal.LivesStatus = failed.Status;
        grandTotal.LivesWarningMessage = failed.Message;

        var validation = new ManulifeBillingCompletionPolicy().Evaluate(extraction);
        Assert.False(validation.CanComplete);
        Assert.Equal(BillingImportStatuses.NeedsReview, validation.RecommendedStatus);
    }

    [Fact]
    public async Task Failed_Database_Save_Prevents_Csv_Export_And_Success_Message()
    {
        var input = CreateInputFile();
        var persistence = new FakePersistence
        {
            Result = new ManulifeBillingPersistenceResult
            {
                Outcome = ManulifeBillingSaveOutcomes.Failed,
                Message = "Database transaction failed"
            }
        };
        var csv = new FakeCsvExporter();
        var coordinator = new ManulifeBillingImportCoordinator(
            persistence,
            new FakeParser(ManulifeBillingTestData.CreateValidExtraction(input)),
            csv);

        var result = await coordinator.ImportAsync(new ManulifeBillingImportRequest
        {
            InputPath = input,
            OutputFolder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"))
        });

        Assert.False(result.Succeeded);
        Assert.Equal(BillingImportStatuses.Failed, result.FinalStatus);
        Assert.Equal(BillingExportStatuses.Skipped, result.ExportStatus);
        Assert.False(csv.WasCalled);
        Assert.DoesNotContain("completed", result.UserMessage, StringComparison.OrdinalIgnoreCase);
    }

    private static string CreateInputFile()
    {
        var folder = Path.Combine(Path.GetTempPath(), "ManulifeBillingCoordinatorTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "billing.pdf");
        File.WriteAllText(path, "test input");
        return path;
    }

    private sealed class FakeParser(BillingExtractionResult result) : IManulifeBillingStatementParser
    {
        public BillingExtractionResult Extract(string pdfPath) => result;
    }

    private sealed class ThrowingParser(Exception exception) : IManulifeBillingStatementParser
    {
        public BillingExtractionResult Extract(string pdfPath) => throw exception;
    }

    private sealed class FakePersistence : IManulifeBillingPersistence
    {
        public bool WasCalled { get; private set; }
        public ManulifeBillingPersistenceResult Result { get; init; } = new()
        {
            Succeeded = true,
            Outcome = ManulifeBillingSaveOutcomes.Saved,
            Message = "Saved"
        };

        public Task<ManulifeBillingPersistenceResult> SaveBillingExtractionAsync(
            BillingExtractionResult result,
            ManulifeBillingSaveMode saveMode = ManulifeBillingSaveMode.Cancel,
            long? selectedClientId = null,
            string performedBy = "")
        {
            WasCalled = true;
            return Task.FromResult(Result);
        }
    }

    private sealed class FakeCsvExporter : IManulifeBillingCsvExporter
    {
        public bool WasCalled { get; private set; }

        public BillingCsvExportResult Write(BillingExtractionResult result, string outputFolder)
        {
            WasCalled = true;
            return new BillingCsvExportResult { OutputFolder = outputFolder };
        }
    }
}
