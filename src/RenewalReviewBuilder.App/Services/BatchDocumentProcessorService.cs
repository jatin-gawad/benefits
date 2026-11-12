using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Connectors;
using RenewalReviewBuilder.App.Connectors.Pdf;
using RenewalReviewBuilder.App.Services.PolicySchedules;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class BatchDocumentProcessorService
{
    private readonly DocumentDiscoveryService _documentDiscoveryService;
    private readonly ReviewPackBuilder _reviewPackBuilder;
    private readonly ClaimsExperienceDataModelBuilder _claimsDataModelBuilder;
    private readonly ClaimsExperiencePowerPointExporter _claimsPowerPointExporter;
    private readonly GroupHealthBillingStatementParser _groupHealthBillingStatementParser;
    private readonly ManulifeBillingImportCoordinator _manulifeBillingImportCoordinator;
    private readonly ManulifeExcelBillingImportService _manulifeExcelBillingImportService;
    private readonly BookletExtractorRunner _bookletExtractorRunner;
    private readonly ProposalEngineRunner _proposalEngineRunner;
    private readonly DatabaseService _databaseService;
    private readonly BenefitsMasterDatabaseExportService _benefitsMasterDatabaseExportService;
    private readonly PdfTextExtractor _pdfTextExtractor;
    private readonly AsoFinancialActivityWorkbookParser _asoFinancialActivityWorkbookParser;
    private readonly ConnectorRegistry _connectorRegistry;
    private readonly PolicyScheduleImportService _policyScheduleImportService;
    private readonly Func<IReadOnlyList<string>, CancellationToken, Task<IReadOnlyList<BookletClientIdentity>>> _bookletClientIdentityProbe;

    public BatchDocumentProcessorService(
        DocumentDiscoveryService? documentDiscoveryService = null,
        ReviewPackBuilder? reviewPackBuilder = null,
        ClaimsExperienceDataModelBuilder? claimsDataModelBuilder = null,
        ClaimsExperiencePowerPointExporter? claimsPowerPointExporter = null,
        GroupHealthBillingStatementParser? groupHealthBillingStatementParser = null,
        ManulifeBillingImportCoordinator? manulifeBillingImportCoordinator = null,
        ManulifeExcelBillingImportService? manulifeExcelBillingImportService = null,
        BookletExtractorRunner? bookletExtractorRunner = null,
        ProposalEngineRunner? proposalEngineRunner = null,
        DatabaseService? databaseService = null,
        PdfTextExtractor? pdfTextExtractor = null,
        Func<IReadOnlyList<string>, CancellationToken, Task<IReadOnlyList<BookletClientIdentity>>>? bookletClientIdentityProbe = null,
        AsoFinancialActivityWorkbookParser? asoFinancialActivityWorkbookParser = null,
        ConnectorRegistry? connectorRegistry = null,
        BenefitsMasterDatabaseExportService? benefitsMasterDatabaseExportService = null)
    {
        _documentDiscoveryService = documentDiscoveryService ?? new DocumentDiscoveryService();
        _reviewPackBuilder = reviewPackBuilder ?? new ReviewPackBuilder();
        _claimsDataModelBuilder = claimsDataModelBuilder ?? new ClaimsExperienceDataModelBuilder();
        _claimsPowerPointExporter = claimsPowerPointExporter ?? new ClaimsExperiencePowerPointExporter();
        _groupHealthBillingStatementParser = groupHealthBillingStatementParser ?? new GroupHealthBillingStatementParser();
        _bookletExtractorRunner = bookletExtractorRunner ?? new BookletExtractorRunner();
        _proposalEngineRunner = proposalEngineRunner ?? new ProposalEngineRunner();
        _databaseService = databaseService ?? new DatabaseService();
        _benefitsMasterDatabaseExportService = benefitsMasterDatabaseExportService ??
            new BenefitsMasterDatabaseExportService(_databaseService.DatabasePath);
        _manulifeBillingImportCoordinator = manulifeBillingImportCoordinator ?? new ManulifeBillingImportCoordinator(_databaseService);
        _manulifeExcelBillingImportService = manulifeExcelBillingImportService ?? new ManulifeExcelBillingImportService(_databaseService);
        _pdfTextExtractor = pdfTextExtractor ?? new PdfTextExtractor();
        _asoFinancialActivityWorkbookParser = asoFinancialActivityWorkbookParser ?? new AsoFinancialActivityWorkbookParser();
        _connectorRegistry = connectorRegistry ?? new ConnectorRegistry();
        _policyScheduleImportService = new PolicyScheduleImportService(_databaseService);
        _bookletClientIdentityProbe = bookletClientIdentityProbe ??
            ((paths, cancellationToken) => _bookletExtractorRunner.DiscoverClientIdentitiesAsync(paths, cancellationToken));
    }

    public Task<BatchDiscoveryResult> DiscoverAsync(BatchProcessingRequest request, IProgress<BatchProcessingProgress>? progress = null)
    {
        return Task.Run(() => DiscoverCoreAsync(request, progress));
    }

    public BatchDiscoveryResult Discover(BatchProcessingRequest request, IProgress<BatchProcessingProgress>? progress = null)
    {
        return DiscoverCoreAsync(request, progress).GetAwaiter().GetResult();
    }

    private async Task<BatchDiscoveryResult> DiscoverCoreAsync(
        BatchProcessingRequest request,
        IProgress<BatchProcessingProgress>? progress)
    {
        var outputFolder = ResolveOutputFolder(request);
        var result = new BatchDiscoveryResult
        {
            ParentFolder = request.ParentFolder,
            OutputFolder = outputFolder
        };

        if (!Directory.Exists(request.ParentFolder))
        {
            result.Warnings.Add($"Parent folder was not found: {request.ParentFolder}");
            return result;
        }

        var selectedFolderClientName = await ResolveSelectedFolderClientNameAsync(request.ParentFolder);
        var clientFolders = string.IsNullOrWhiteSpace(selectedFolderClientName)
            ? ResolveClientFolders(request.ParentFolder)
            : [request.ParentFolder];
        var total = clientFolders.Count;
        for (var index = 0; index < total; index++)
        {
            var clientFolder = clientFolders[index];
            progress?.Report(new BatchProcessingProgress
            {
                CurrentClient = Path.GetFileName(clientFolder),
                CurrentWorkflow = "Document Discovery",
                Completed = index,
                Total = total,
                Message = $"Discovering documents for {Path.GetFileName(clientFolder)}..."
            });

            var discovery = _documentDiscoveryService.DiscoverClientFolder(clientFolder);
            if (IsFlatBookletTestFolder(request.ParentFolder, clientFolder, discovery))
            {
                progress?.Report(new BatchProcessingProgress
                {
                    CurrentClient = Path.GetFileName(clientFolder),
                    CurrentWorkflow = "Booklet Client Grouping",
                    Completed = index,
                    Total = total,
                    Message = "Identifying booklet client groups in the selected flat folder..."
                });
                result.Clients.AddRange(await BuildFlatBookletClientDiscoveriesAsync(discovery, outputFolder));
            }
            else
            {
                result.Clients.AddRange(BuildClientDiscoveries(discovery, outputFolder, selectedFolderClientName));
            }
        }

        progress?.Report(new BatchProcessingProgress
        {
            CurrentWorkflow = "Document Discovery",
            Completed = total,
            Total = total,
            Message = $"Document discovery finished for {total} selected folder(s) and {result.Clients.Count} client group(s)."
        });

        return result;
    }

    public async Task<BatchProcessingResult> ProcessBatchAsync(
        BatchDiscoveryResult discovery,
        BatchProcessingRequest request,
        IProgress<BatchProcessingProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var outputFolder = string.IsNullOrWhiteSpace(discovery.OutputFolder)
            ? ResolveOutputFolder(request)
            : discovery.OutputFolder;
        Directory.CreateDirectory(outputFolder);

        var result = new BatchProcessingResult
        {
            ParentFolder = discovery.ParentFolder,
            OutputFolder = outputFolder
        };

        var batchWatch = Stopwatch.StartNew();
        var total = discovery.Clients.Count;
        for (var index = 0; index < total; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var client = discovery.Clients[index];
            var completed = index;
            progress?.Report(BuildProgress(client, "Starting client", completed, total, result, batchWatch));
            var clientResult = await ProcessClientAsync(client, request, progress, completed, total, result, batchWatch, cancellationToken);
            result.Clients.Add(clientResult);
            progress?.Report(BuildProgress(client, "Client complete", index + 1, total, result, batchWatch));
        }

        result.LogPath = WriteBatchLog(result);
        return result;
    }

    private async Task<BatchClientProcessingResult> ProcessClientAsync(
        BatchClientDiscoveryResult client,
        BatchProcessingRequest request,
        IProgress<BatchProcessingProgress>? progress,
        int completed,
        int total,
        BatchProcessingResult batchResult,
        Stopwatch batchWatch,
        CancellationToken cancellationToken)
    {
        var watch = Stopwatch.StartNew();
        var state = new BatchClientProcessingState();
        Directory.CreateDirectory(client.OutputFolder);

        var result = new BatchClientProcessingResult
        {
            ClientName = client.ClientName,
            Carrier = client.Carrier,
            DiscoveryStatus = client.Status,
            OutputFolder = client.OutputFolder,
            DocumentTypesFound = client.DocumentTypesText,
            DocumentCarriersFound = client.DocumentCarriersText,
            PlanDesignSource = client.PlanDesignSource,
            PlanDesignStatus = client.PlanDesignStatus
        };
        result.Warnings.AddRange(client.Warnings);

        if (client.WorkflowAvailability.AvailableWorkflows.Count == 0)
        {
            result.Status = BatchProcessingStatuses.Skipped;
            result.Warnings.Add("No runnable workflow was available from document discovery.");
            result.ProcessingTime = watch.Elapsed;
            return result;
        }

        if (client.WorkflowAvailability.RenewalIntelligence)
        {
            await RunWorkflowAsync(
                client,
                result,
                BatchWorkflowNames.RenewalIntelligence,
                client.DocumentDiscovery.RenewalDocument?.FilePath ?? "",
                progress,
                completed,
                total,
                batchResult,
                batchWatch,
                () => RunRenewalWorkflowAsync(client, request, state));
        }

        if (client.WorkflowAvailability.ClaimsExperience)
        {
            ClaimsExperienceDataModelResult? claimsModel = null;
            await RunWorkflowAsync(
                client,
                result,
                BatchWorkflowNames.ClaimsExperience,
                FirstFile(ClaimsFiles(client)),
                progress,
                completed,
                total,
                batchResult,
                batchWatch,
                async () =>
                {
                    claimsModel = await CreateClaimsModelAsync(client, request);
                    return BuildClaimsWorkflowResult(claimsModel);
                });

            if (claimsModel is not null && File.Exists(request.ClaimsPowerPointTemplatePath))
            {
                await RunWorkflowAsync(
                    client,
                    result,
                    BatchWorkflowNames.ClaimsPowerPoint,
                    claimsModel.WorkbookPath,
                    progress,
                    completed,
                    total,
                    batchResult,
                    batchWatch,
                    () => RunClaimsPowerPointWorkflowAsync(claimsModel, request));
            }
            else if (claimsModel is not null && !string.IsNullOrWhiteSpace(request.ClaimsPowerPointTemplatePath))
            {
                result.Warnings.Add($"Claims PowerPoint template was not found: {request.ClaimsPowerPointTemplatePath}");
            }
        }

        foreach (var billingDocument in BillingFiles(client))
        {
            await RunWorkflowAsync(
                client,
                result,
                BatchWorkflowNames.Billing,
                billingDocument.FilePath,
                progress,
                completed,
                total,
                batchResult,
                batchWatch,
                () => RunBillingWorkflowAsync(client, billingDocument));
        }

        foreach (var financialDocument in AsoFinancialActivityFiles(client))
        {
            await RunWorkflowAsync(
                client,
                result,
                BatchWorkflowNames.FinancialIntelligence,
                financialDocument.FilePath,
                progress,
                completed,
                total,
                batchResult,
                batchWatch,
                () => RunAsoFinancialActivityWorkflowAsync(client, financialDocument));
        }

        if (client.WorkflowAvailability.BookletExtraction)
        {
            await RunWorkflowAsync(
                client,
                result,
                BatchWorkflowNames.BookletExtraction,
                FirstFile(BookletFiles(client)),
                progress,
                completed,
                total,
                batchResult,
                batchWatch,
                () => RunBookletWorkflowAsync(client, cancellationToken));
        }

        if (client.DocumentDiscovery.PlanDesign.Source == PlanDesignSources.PolicySchedule)
        {
            await RunWorkflowAsync(
                client,
                result,
                BatchWorkflowNames.PlanDesign,
                client.DocumentDiscovery.PlanDesign.FilePath,
                progress,
                completed,
                total,
                batchResult,
                batchWatch,
                () => RunPolicyScheduleWorkflowAsync(client));
        }
        else if (client.WorkflowAvailability.PlanDesign && !client.WorkflowAvailability.BookletExtraction)
        {
            await RunWorkflowAsync(
                client,
                result,
                BatchWorkflowNames.PlanDesign,
                client.DocumentDiscovery.PlanDesign.FilePath,
                progress,
                completed,
                total,
                batchResult,
                batchWatch,
                () => RunGroupInformationReportWorkflowAsync(client, state));
        }

        if (client.WorkflowAvailability.ProposalComparison)
        {
            await RunWorkflowAsync(
                client,
                result,
                BatchWorkflowNames.ProposalComparison,
                FirstFile(ProposalFiles(client)),
                progress,
                completed,
                total,
                batchResult,
                batchWatch,
                () => RunProposalWorkflowAsync(client, request));
        }

        watch.Stop();
        await ValidateClientOutputsAsync(client, result, state);
        result.ProcessingTime = watch.Elapsed;
        result.Status = DetermineClientStatus(result);
        return result;
    }

    private async Task RunWorkflowAsync(
        BatchClientDiscoveryResult client,
        BatchClientProcessingResult clientResult,
        string workflowName,
        string sourceFile,
        IProgress<BatchProcessingProgress>? progress,
        int completed,
        int total,
        BatchProcessingResult batchResult,
        Stopwatch batchWatch,
        Func<Task<BatchWorkflowResult>> workflow)
    {
        progress?.Report(BuildProgress(client, workflowName, completed, total, batchResult, batchWatch));
        var watch = Stopwatch.StartNew();
        try
        {
            var workflowResult = await workflow();
            workflowResult.WorkflowName = workflowName;
            workflowResult.SourceFile = sourceFile;
            workflowResult.ProcessingTime = watch.Elapsed;
            clientResult.Workflows.Add(workflowResult);
            clientResult.Warnings.AddRange(workflowResult.Warnings);
            if (workflowResult.Status == BatchProcessingStatuses.Failed)
            {
                clientResult.Errors.Add($"{workflowName}: {workflowResult.Message}");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            clientResult.Errors.Add($"{workflowName}: {ex.Message}");
            clientResult.Workflows.Add(new BatchWorkflowResult
            {
                WorkflowName = workflowName,
                SourceFile = sourceFile,
                Status = BatchProcessingStatuses.Failed,
                Message = ex.Message,
                Exception = ex.ToString(),
                StackTrace = ex.StackTrace ?? "",
                ProcessingTime = watch.Elapsed
            });
        }
    }

    private async Task<BatchWorkflowResult> RunRenewalWorkflowAsync(
        BatchClientDiscoveryResult client,
        BatchProcessingRequest request,
        BatchClientProcessingState state)
    {
        var renewalDocument = client.DocumentDiscovery.RenewalDocument;
        if (renewalDocument is null)
        {
            return Skipped(BatchWorkflowNames.RenewalIntelligence, "No renewal PDF was found.");
        }

        var workbookFiles = WorkbookFiles(client);
        var primaryWorkbook = SelectPrimaryRenewalWorkbook(workbookFiles);
        var renewalCase = new RenewalCase
        {
            PdfPath = renewalDocument.FilePath,
            AdditionalPdfPaths = SupportingPdfFiles(client, renewalDocument.FilePath),
            WorkbookPath = primaryWorkbook,
            AdditionalWorkbookPaths = workbookFiles
                .Where(path => !path.Equals(primaryWorkbook, StringComparison.OrdinalIgnoreCase))
                .ToList(),
            PowerPointTemplatePath = request.RenewalPowerPointTemplatePath,
            OutputFolder = client.OutputFolder
        };

        var analysis = await _reviewPackBuilder.AnalyzeAsync(renewalCase);
        analysis.RenewalCase.OutputFolder = client.OutputFolder;
        var resolvedClientName = FirstNonBlank(
            ClientNameResolver.Resolve(
                analysis.RenewalCase.ClientName,
                FirstNonBlank(analysis.Pdf.ClientName, client.ClientName),
                renewalDocument.FilePath),
            client.ClientName);
        analysis.RenewalCase.ClientName = resolvedClientName;
        if (ClientNameResolver.IsInvalidClientName(analysis.Pdf.ClientName))
        {
            analysis.Pdf.ClientName = resolvedClientName;
        }

        state.RenewalAnalysis = analysis;

        var databaseSave = await _databaseService.SaveRenewalAnalysisAsync(analysis);
        state.RenewalDatabaseSave = databaseSave;
        var export = await _reviewPackBuilder.ExportAsync(analysis);

        var workflow = new BatchWorkflowResult
        {
            SavedToDatabase = databaseSave.Succeeded && databaseSave.ClientId.HasValue && databaseSave.DocumentId.HasValue,
            DatabaseRowsSaved = databaseSave.InsertedRows,
            DatabaseSaveStatus = databaseSave.Message,
            Message = databaseSave.Message
        };
        AddExistingOutput(workflow, export.WorkbookPath);
        AddExistingOutput(workflow, export.PowerPointPath);
        ApplyClientCodeToWorkflowOutputs(workflow, client.ClientName);
        workflow.WorkbookGenerated = workflow.OutputPaths.Any(IsExcelWorkbookPath);
        workflow.PowerPointGenerated = workflow.OutputPaths.Any(IsPowerPointPath);
        if (!databaseSave.Succeeded)
        {
            workflow.Warnings.Add(databaseSave.Message);
        }
        else if (databaseSave.InsertedRows <= 0)
        {
            workflow.Warnings.Add("Renewal output generated but renewal database save did not report inserted renewal/rate/version rows.");
        }

        if (!workflow.WorkbookGenerated)
        {
            workflow.Warnings.Add("Renewal Review workbook was not generated.");
        }

        if (!workflow.PowerPointGenerated)
        {
            workflow.Warnings.Add("Renewal PowerPoint was not generated.");
        }

        workflow.Status = workflow.Warnings.Count == 0
            ? BatchProcessingStatuses.Completed
            : BatchProcessingStatuses.CompletedWithWarnings;

        return workflow;
    }

    private async Task<ClaimsExperienceDataModelResult> CreateClaimsModelAsync(
        BatchClientDiscoveryResult client,
        BatchProcessingRequest request)
    {
        var sourceFiles = ClaimsFiles(client);
        var claimsResult = await _claimsDataModelBuilder.ExportAsync(new ClaimsExperienceDataModelRequest
        {
            SourceFiles = sourceFiles,
            OutputFolder = client.OutputFolder,
            ReportMonthCount = request.ReportMonthCount,
            BenefitBookletPaths = BookletFiles(client)
        });

        await _databaseService.SaveClaimsExperienceAsync(claimsResult, sourceFiles, client.ClientFolder);
        return claimsResult;
    }

    private async Task<BatchWorkflowResult> RunClaimsPowerPointWorkflowAsync(
        ClaimsExperienceDataModelResult claimsModel,
        BatchProcessingRequest request)
    {
        var ppt = await _claimsPowerPointExporter.ExportAsync(new ClaimsExperiencePowerPointRequest
        {
            WorkbookPath = claimsModel.WorkbookPath,
            TemplatePath = request.ClaimsPowerPointTemplatePath,
            OutputFolder = claimsModel.OutputFolder
        });

        var workflow = new BatchWorkflowResult
        {
            Status = ppt.Warnings.Count > 0 ? BatchProcessingStatuses.CompletedWithWarnings : BatchProcessingStatuses.Completed,
            Message = $"Claims PowerPoint updated {ppt.SlideCountUpdated} slide(s)."
        };
        AddExistingOutput(workflow, ppt.PowerPointPath);
        workflow.Warnings.AddRange(ppt.Warnings);
        return workflow;
    }

    private static BatchWorkflowResult BuildClaimsWorkflowResult(ClaimsExperienceDataModelResult claimsModel)
    {
        var workflow = new BatchWorkflowResult
        {
            Status = claimsModel.Warnings.Count > 0 ? BatchProcessingStatuses.CompletedWithWarnings : BatchProcessingStatuses.Completed,
            Message = $"Claims model created with {claimsModel.TableCount} table(s)."
        };
        AddExistingOutput(workflow, claimsModel.WorkbookPath);
        workflow.Warnings.AddRange(claimsModel.Warnings);
        return workflow;
    }

    private async Task<BatchWorkflowResult> RunBillingWorkflowAsync(
        BatchClientDiscoveryResult client,
        DiscoveredDocument billingDocument)
    {
        if (!billingDocument.FileExtension.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return await RunManulifeExcelBillingPreviewWorkflowAsync(client, billingDocument);
        }

        return IsGroupHealthBillingDocument(billingDocument)
            ? await RunGroupHealthBillingWorkflowAsync(billingDocument)
            : await RunManulifeBillingWorkflowAsync(client, billingDocument);
    }

    private async Task<BatchWorkflowResult> RunAsoFinancialActivityWorkflowAsync(
        BatchClientDiscoveryResult client,
        DiscoveredDocument financialDocument)
    {
        var statement = await Task.Run(() => _asoFinancialActivityWorkbookParser.Parse(financialDocument.FilePath));
        if (string.IsNullOrWhiteSpace(statement.ClientName))
        {
            statement.ClientName = client.ClientName;
        }

        if (string.IsNullOrWhiteSpace(statement.Carrier) &&
            !string.IsNullOrWhiteSpace(client.Carrier) &&
            !client.Carrier.Equals("Unknown", StringComparison.OrdinalIgnoreCase))
        {
            statement.Carrier = client.Carrier;
        }

        var databaseSave = await _databaseService.SaveAsoFinancialActivityAsync(
            statement,
            Environment.UserName);
        var workflow = new BatchWorkflowResult
        {
            SavedToDatabase = databaseSave.Succeeded,
            DatabaseRowsSaved = databaseSave.InsertedRows,
            DatabaseSaveStatus = databaseSave.Message,
            Message = databaseSave.Message,
            Status = databaseSave.Succeeded
                ? statement.ImportStatus.Equals(AsoFinancialImportStatuses.Complete, StringComparison.OrdinalIgnoreCase)
                    ? BatchProcessingStatuses.Completed
                    : BatchProcessingStatuses.CompletedWithWarnings
                : BatchProcessingStatuses.Failed
        };

        if (statement.ValidationIssues.Count > 0)
        {
            workflow.Warnings.AddRange(statement.ValidationIssues);
        }

        if (!databaseSave.Succeeded)
        {
            workflow.Warnings.Add(databaseSave.Message);
        }

        return workflow;
    }

    private async Task<BatchWorkflowResult> RunManulifeExcelBillingPreviewWorkflowAsync(
        BatchClientDiscoveryResult client,
        DiscoveredDocument billingDocument)
    {
        var initialClientName = ResolveManulifeBillingClientName(client, billingDocument);
        var preview = await _manulifeExcelBillingImportService.CreatePreviewAsync(new ManulifeExcelBillingImportRequest
        {
            InputPath = billingDocument.FilePath,
            SelectedClientName = initialClientName,
            ImportedBy = Environment.UserName
        });

        var clientName = FirstNonBlank(preview.DetectedClientName, initialClientName);
        if (!string.Equals(clientName, initialClientName, StringComparison.OrdinalIgnoreCase))
        {
            preview = await _manulifeExcelBillingImportService.CreatePreviewAsync(new ManulifeExcelBillingImportRequest
            {
                InputPath = billingDocument.FilePath,
                SelectedClientName = clientName,
                ImportedBy = Environment.UserName
            });
        }

        if (!preview.CanConfirm)
        {
            var previewWorkflow = new BatchWorkflowResult
            {
                Status = BatchProcessingStatuses.CompletedWithWarnings,
                Message = $"Manulife Excel billing preview read {preview.EmployeeRecordCount} employee row(s) and {preview.BenefitRecordCount} benefit row(s).",
                DatabaseSaveStatus = "Preview only - a client name could not be confirmed from the selected folder or file name."
            };
            previewWorkflow.Warnings.Add($"Manulife Excel billing was read but not saved. Errors: {preview.ErrorCount}. Warnings: {preview.WarningCount}.");
            return previewWorkflow;
        }

        var databaseSave = await _manulifeExcelBillingImportService.ConfirmAsync(preview, new ManulifeExcelBillingImportRequest
        {
            InputPath = billingDocument.FilePath,
            SelectedClientName = clientName,
            BillingPeriod = preview.BillingPeriod,
            StatementDate = preview.StatementDate,
            ImportedBy = Environment.UserName
        });
        var workflow = new BatchWorkflowResult
        {
            SavedToDatabase = databaseSave.Succeeded,
            DatabaseRowsSaved = databaseSave.SavedEmployeeRecordCount + databaseSave.SavedBenefitRecordCount,
            DatabaseSaveStatus = databaseSave.Message,
            Message = databaseSave.Message,
            Status = databaseSave.Succeeded && preview.WarningCount == 0
                ? BatchProcessingStatuses.Completed
                : BatchProcessingStatuses.CompletedWithWarnings
        };
        if (!databaseSave.Succeeded)
        {
            workflow.Status = BatchProcessingStatuses.Failed;
            workflow.Warnings.Add(databaseSave.Message);
        }
        else if (preview.WarningCount > 0)
        {
            workflow.Warnings.Add($"Manulife Excel billing was saved with {preview.WarningCount} warning(s).");
        }

        return workflow;
    }

    private async Task<BatchWorkflowResult> RunGroupHealthBillingWorkflowAsync(DiscoveredDocument billingDocument)
    {
        var extraction = await Task.Run(() => _groupHealthBillingStatementParser.Extract(billingDocument.FilePath));
        var databaseSave = await _databaseService.SaveGroupHealthBillingStatementAsync(
            extraction,
            GroupHealthBillingSaveMode.NewVersion);

        var workflow = new BatchWorkflowResult
        {
            SavedToDatabase = databaseSave.Succeeded,
            DatabaseRowsSaved = databaseSave.InsertedRows,
            DatabaseSaveStatus = databaseSave.Message,
            Message = databaseSave.Message
        };

        if (databaseSave.Succeeded)
        {
            workflow.Status = extraction.ImportStatus.Equals(GroupHealthBillingImportStatuses.Passed, StringComparison.OrdinalIgnoreCase)
                ? BatchProcessingStatuses.Completed
                : BatchProcessingStatuses.CompletedWithWarnings;
            if (workflow.Status == BatchProcessingStatuses.CompletedWithWarnings)
            {
                workflow.Warnings.Add($"Billing statement was saved with status: {extraction.ImportStatus}.");
            }

            return workflow;
        }

        if (databaseSave.Message.Contains("already been imported", StringComparison.OrdinalIgnoreCase))
        {
            workflow.Status = BatchProcessingStatuses.CompletedWithWarnings;
            workflow.Warnings.Add("This billing statement was already imported, so the batch did not create a duplicate.");
            return workflow;
        }

        workflow.Status = BatchProcessingStatuses.Failed;
        workflow.Warnings.Add(databaseSave.Message);
        return workflow;
    }

    private async Task<BatchWorkflowResult> RunManulifeBillingWorkflowAsync(
        BatchClientDiscoveryResult client,
        DiscoveredDocument billingDocument)
    {
        var import = await _manulifeBillingImportCoordinator.ImportAsync(new ManulifeBillingImportRequest
        {
            InputPath = billingDocument.FilePath,
            OutputFolder = client.OutputFolder,
            ExportCsv = true,
            SaveMode = ManulifeBillingSaveMode.ImportAsNewVersion,
            PerformedBy = Environment.UserName
        });

        var workflow = new BatchWorkflowResult
        {
            SavedToDatabase = import.Succeeded,
            DatabaseRowsSaved = import.SavedRowCount,
            DatabaseSaveStatus = import.Persistence?.Message ?? import.UserMessage,
            Message = import.UserMessage
        };
        if (import.CsvExport is not null)
        {
            AddExistingOutput(workflow, import.CsvExport.ClassesDetectedPath);
            AddExistingOutput(workflow, import.CsvExport.ReconciliationResultsPath);
            AddExistingOutput(workflow, import.CsvExport.ValidationErrorsPath);
        }

        workflow.Warnings.AddRange(import.TechnicalErrors);
        if (import.Succeeded)
        {
            workflow.Status = import.WarningCount > 0 ||
                              import.FinalStatus.Equals(BillingImportStatuses.NeedsReview, StringComparison.OrdinalIgnoreCase)
                ? BatchProcessingStatuses.CompletedWithWarnings
                : BatchProcessingStatuses.Completed;
            return workflow;
        }

        workflow.Status = import.Persistence?.Outcome is ManulifeBillingSaveOutcomes.Duplicate or ManulifeBillingSaveOutcomes.ExistingVersion
            ? BatchProcessingStatuses.CompletedWithWarnings
            : BatchProcessingStatuses.Failed;
        return workflow;
    }

    private async Task<BatchWorkflowResult> RunBookletWorkflowAsync(
        BatchClientDiscoveryResult client,
        CancellationToken cancellationToken)
    {
        var inputPaths = BookletFiles(client);
        var bookletResult = await _bookletExtractorRunner.RunAsync(
            new BookletExtractionRequest
            {
                InputPaths = inputPaths,
                OutputFolder = client.OutputFolder
            },
            null,
            cancellationToken);

        var databaseSave = await _databaseService.SavePlanDesignFromBookletExtractionAsync(
            bookletResult,
            inputPaths,
            targetClient: new ClientRecord
            {
                ClientName = client.ClientName,
                Carrier = client.Carrier
            });
        BenefitsMasterDatabaseWorkbookResult? masterRebuild = null;
        string masterRebuildError = "";
        if (databaseSave.Succeeded && databaseSave.ClientId.HasValue)
        {
            try
            {
                masterRebuild = await _benefitsMasterDatabaseExportService.RebuildAsync(
                    client.OutputFolder,
                    [databaseSave.ClientId.Value],
                    cancellationToken);
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex);
                masterRebuildError = ex.Message;
            }
        }

        var workflow = new BatchWorkflowResult
        {
            Status = databaseSave.Succeeded && string.IsNullOrWhiteSpace(masterRebuildError)
                ? BatchProcessingStatuses.Completed
                : BatchProcessingStatuses.CompletedWithWarnings,
            Message = databaseSave.Message +
                      (masterRebuild is null
                          ? string.IsNullOrWhiteSpace(masterRebuildError)
                              ? ""
                              : $" Workbook rebuild failed: {masterRebuildError}"
                          : $" Rebuilt {masterRebuild.SummarySheetCount} Benefit Summary sheet(s) from SQLite."),
            SavedToDatabase = databaseSave.Succeeded && databaseSave.InsertedRows > 0,
            DatabaseRowsSaved = databaseSave.InsertedRows,
            PlanDesignRowsExtracted = databaseSave.InsertedRows,
            PlanDesignRowsSaved = databaseSave.InsertedRows,
            PlanDesignTargetTable = "PlanDesign",
            ScheduleOfBenefitsUpdated = masterRebuild is not null && File.Exists(masterRebuild.MasterWorkbookPath),
            DatabaseSaveStatus = databaseSave.Message
        };
        AddExistingOutput(workflow, bookletResult.MasterWorkbookPath);
        AddExistingOutput(workflow, bookletResult.ReviewWorkbookPath);
        if (!databaseSave.Succeeded)
        {
            workflow.Warnings.Add(databaseSave.Message);
        }
        else if (databaseSave.InsertedRows <= 0)
        {
            workflow.Warnings.Add("Booklet extraction completed, but no PlanDesign rows were saved.");
            workflow.Status = BatchProcessingStatuses.CompletedWithWarnings;
        }

        if (!string.IsNullOrWhiteSpace(masterRebuildError))
        {
            workflow.Warnings.Add(
                $"Plan Design was saved to SQLite, but the Benefits Master Database could not be rebuilt: {masterRebuildError}");
        }

        return workflow;
    }

    private async Task<BatchWorkflowResult> RunGroupInformationReportWorkflowAsync(
        BatchClientDiscoveryResult client,
        BatchClientProcessingState state)
    {
        var source = client.DocumentDiscovery.PlanDesign;
        if (string.IsNullOrWhiteSpace(source.FilePath) || !File.Exists(source.FilePath))
        {
            return Skipped(BatchWorkflowNames.PlanDesign, "No Group Information Report or embedded plan design source was available.");
        }

        var pages = _pdfTextExtractor.Extract(source.FilePath);
        var selectedPages = SelectPlanDesignPages(source, pages);
        var features = GenericRenewalDocumentAnalyzer.ExtractPlanDesignFeatures(selectedPages);
        foreach (var feature in features)
        {
            feature.SourceFileName = source.FileName;
            feature.PlanDesignSource = source.Source;
        }

        state.PlanDesignFeatures.AddRange(features);
        state.PlanDesignSourcePages.AddRange(selectedPages.Select(page => page.PageNumber));

        var workflow = new BatchWorkflowResult
        {
            PlanDesignRowsExtracted = features.Count,
            PlanDesignTargetTable = "plan_design_features"
        };

        if (state.RenewalAnalysis is null || state.RenewalDatabaseSave is null)
        {
            workflow.Status = BatchProcessingStatuses.CompletedWithWarnings;
            workflow.Message = "Plan design was detected, but it was not saved because the renewal workflow did not complete first.";
            workflow.Warnings.Add(workflow.Message);
            return workflow;
        }

        if (!state.RenewalDatabaseSave.Succeeded ||
            !state.RenewalDatabaseSave.ClientId.HasValue ||
            !state.RenewalDatabaseSave.DocumentId.HasValue)
        {
            workflow.Status = BatchProcessingStatuses.CompletedWithWarnings;
            workflow.Message = "Plan design was detected, but database save was skipped because the renewal database record was not confirmed.";
            workflow.Warnings.Add(workflow.Message);
            workflow.DatabaseSaveStatus = state.RenewalDatabaseSave.Message;
            return workflow;
        }

        var carrierDetection = GenericRenewalDocumentAnalyzer.DetectCarrier(ParsingHelpers.Clean(string.Join(" ", selectedPages.Select(page => page.Text))));
        var sections = GenericRenewalDocumentAnalyzer.DetectSections(selectedPages);
        var auditEntries = GenericRenewalDocumentAnalyzer.BuildAuditLog(
            source.FilePath,
            state.RenewalAnalysis,
            carrierDetection,
            sections,
            features);

        var databaseSave = await _databaseService.SaveGenericRenewalArtifactsAsync(
            state.RenewalAnalysis,
            state.RenewalDatabaseSave.ClientId,
            state.RenewalDatabaseSave.DocumentId,
            auditEntries,
            features);
        state.GenericArtifactSave = databaseSave;

        workflow.DatabaseSaveStatus = databaseSave.Message;
        workflow.DatabaseRowsSaved = databaseSave.InsertedRows;
        workflow.SavedToDatabase = databaseSave.Succeeded;

        if (databaseSave.Succeeded)
        {
            workflow.PlanDesignRowsSaved = await CountRowsAsync("plan_design_features", state.RenewalDatabaseSave.ClientId, state.RenewalDatabaseSave.DocumentId);
        }

        if (features.Count > 0)
        {
            var planDesignWorkbooks = _bookletExtractorRunner.WritePlanDesignWorkbooksFromFeatures(new BookletPlanDesignWorkbookRequest
            {
                OutputFolder = client.OutputFolder,
                ClientName = FirstNonBlank(state.RenewalAnalysis.RenewalCase.ClientName, client.ClientName),
                Carrier = FirstNonBlank(state.RenewalAnalysis.Pdf.Carrier, client.Carrier),
                PolicyNumber = state.RenewalAnalysis.RenewalCase.PolicyNumber,
                SourceFilePath = source.FilePath,
                SourceFileName = source.FileName,
                PlanDesignSource = source.Source,
                EffectiveDate = state.RenewalAnalysis.RenewalCase.RenewalDate?.ToString("MMM d, yyyy", CultureInfo.InvariantCulture) ?? "",
                Features = features
            });
            AddExistingOutput(workflow, planDesignWorkbooks.MasterWorkbookPath);
            AddExistingOutput(workflow, planDesignWorkbooks.ReviewWorkbookPath);
            ApplyClientCodeToWorkflowOutputs(workflow, client.ClientName);
            workflow.ScheduleOfBenefitsUpdated = workflow.OutputPaths.Any(path =>
                File.Exists(path) &&
                Path.GetFileName(path).StartsWith("Booklet_Extraction_Review", StringComparison.OrdinalIgnoreCase));
        }

        workflow.Status = databaseSave.Succeeded && workflow.PlanDesignRowsSaved > 0
            ? BatchProcessingStatuses.Completed
            : BatchProcessingStatuses.CompletedWithWarnings;
        workflow.Message = workflow.PlanDesignRowsSaved > 0
            ? $"Plan design features saved from {source.Source}: {workflow.PlanDesignRowsSaved} row(s) to plan_design_features."
            : $"Plan design source found ({source.Source}), but no feature rows were saved.";

        if (features.Count == 0)
        {
            workflow.Warnings.Add($"Plan design source found on page(s) {FormatPageList(source.PageNumbers)}, but the existing analyzer did not return feature rows.");
        }

        if (!databaseSave.Succeeded)
        {
            workflow.Warnings.Add(databaseSave.Message);
        }

        if (source.Source == PlanDesignSources.EmbeddedRenewal && workflow.PlanDesignRowsSaved == 0)
        {
            workflow.Warnings.Add($"Embedded GIR detected but no plan design rows were saved. Source page(s): {FormatPageList(source.PageNumbers)}.");
        }

        return workflow;
    }

    private async Task<BatchWorkflowResult> RunPolicyScheduleWorkflowAsync(BatchClientDiscoveryResult client)
    {
        var source = client.DocumentDiscovery.PlanDesign;
        var document = client.DocumentDiscovery.Documents.FirstOrDefault(item =>
            item.IsGroupPolicySchedule &&
            item.FilePath.Equals(source.FilePath, StringComparison.OrdinalIgnoreCase));
        if (document is null || !File.Exists(source.FilePath))
        {
            return Skipped(BatchWorkflowNames.PlanDesign, "The selected group policy schedule could not be found.");
        }

        var carrier = FirstNonBlank(document.Carrier, client.Carrier, "Unknown");
        var processorRegistered = _connectorRegistry.TryResolve(
            SourceTypes.Pdf,
            carrier,
            DocumentTypes.GroupPolicySchedule,
            out var connector);
        if (processorRegistered && connector is IPolicyScheduleConnector policyConnector && connector.IsEnabled)
        {
            var import = await _policyScheduleImportService.ImportAsync(
                new ClientRecord { ClientName = client.ClientName, Carrier = carrier },
                policyConnector,
                source.FilePath);
            var completed = import.Succeeded && import.InsertedRows > 0;
            var importedWorkflow = new BatchWorkflowResult
            {
                Status = completed ? BatchProcessingStatuses.Completed : BatchProcessingStatuses.CompletedWithWarnings,
                Message = import.Message,
                SavedToDatabase = import.Succeeded,
                DatabaseRowsSaved = import.InsertedRows,
                PlanDesignRowsExtracted = import.InsertedRows,
                PlanDesignRowsSaved = import.InsertedRows,
                PlanDesignTargetTable = "PlanDesign / plan_design_features",
                DatabaseSaveStatus = import.Message
            };
            if (!completed) importedWorkflow.Warnings.Add("Policy schedule was preserved but did not produce complete Plan Design rows.");
            return importedWorkflow;
        }

        var issueCode = !processorRegistered
            ? ValidationIssueCodes.PolicyScheduleProcessorNotRegistered
            : ValidationIssueCodes.PolicyScheduleExtractionNotImplemented;
        var message = !processorRegistered
            ? $"No policy-schedule processor is registered for {carrier}. The source was preserved for later reprocessing."
            : $"{connector.ConnectorName} is registered for {carrier}, but policy-schedule extraction is not implemented in Phase 2. The source was preserved for later reprocessing.";

        var persistence = await _databaseService.RecordPolicyScheduleReviewAsync(
            new ClientRecord { ClientName = client.ClientName, Carrier = carrier },
            document,
            issueCode,
            message);

        var workflow = new BatchWorkflowResult
        {
            Status = persistence.Succeeded ? BatchProcessingStatuses.CompletedWithWarnings : BatchProcessingStatuses.Failed,
            Message = persistence.Succeeded ? message : persistence.Message,
            SavedToDatabase = persistence.Succeeded,
            DatabaseRowsSaved = persistence.InsertedRows,
            PlanDesignTargetTable = "ImportBatch / ValidationIssue",
            DatabaseSaveStatus = persistence.Message
        };
        workflow.Warnings.Add(message);
        if (!persistence.Succeeded)
        {
            workflow.Warnings.Add(persistence.Message);
        }

        return workflow;
    }

    private async Task<BatchWorkflowResult> RunProposalWorkflowAsync(
        BatchClientDiscoveryResult client,
        BatchProcessingRequest request)
    {
        var proposalFiles = ProposalFiles(client);
        var result = await _proposalEngineRunner.RunComparisonAsync(
            proposalFiles,
            client.ClientName,
            BookletFiles(client).FirstOrDefault(),
            client.OutputFolder,
            string.IsNullOrWhiteSpace(request.DatabasePath) ? _databaseService.DatabasePath : request.DatabasePath,
            clientName: client.ClientName);

        return new BatchWorkflowResult
        {
            Status = result.Succeeded ? BatchProcessingStatuses.Completed : BatchProcessingStatuses.Failed,
            Message = result.Message,
            OutputPaths = { result.OutputFolder },
            SavedToDatabase = result.Succeeded,
            PlanDesignTargetTable = "proposal_plan_design",
            DatabaseSaveStatus = result.Succeeded
                ? $"Proposal engine completed. SQLite target: {result.DatabasePath}"
                : result.Error
        };
    }

    private static IEnumerable<BatchClientDiscoveryResult> BuildClientDiscoveries(
        RenewalDocumentDiscoveryResult discovery,
        string outputRoot,
        string? clientNameOverride = null)
    {
        if (!string.IsNullOrWhiteSpace(clientNameOverride) || !IsBillingOnlyFolder(discovery))
        {
            yield return BuildClientDiscovery(discovery, outputRoot, clientNameOverride);
            yield break;
        }

        foreach (var group in discovery.Documents
                     .GroupBy(ResolveBillingDiscoveryClientName, StringComparer.OrdinalIgnoreCase)
                     .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            var groupedDiscovery = new RenewalDocumentDiscoveryResult
            {
                ClientFolder = discovery.ClientFolder
            };
            groupedDiscovery.Documents.AddRange(group);
            yield return BuildClientDiscovery(groupedDiscovery, outputRoot, group.Key);
        }
    }

    private async Task<IReadOnlyList<BatchClientDiscoveryResult>> BuildFlatBookletClientDiscoveriesAsync(
        RenewalDocumentDiscoveryResult discovery,
        string outputRoot)
    {
        var bookletDocuments = discovery.Documents
            .Where(document => document.IsBenefitBooklet)
            .ToList();
        IReadOnlyList<BookletClientIdentity> identities;
        try
        {
            identities = await _bookletClientIdentityProbe(
                bookletDocuments.Select(document => document.FilePath).ToList(),
                CancellationToken.None);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return [BuildUnassignedBookletReview(
                discovery,
                bookletDocuments,
                outputRoot,
                "Booklet client grouping could not read the client identities. Review the assigned files before processing.")];
        }

        var identityByPath = identities
            .Where(identity => !string.IsNullOrWhiteSpace(identity.SourcePath))
            .GroupBy(identity => Path.GetFullPath(identity.SourcePath), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        var identifiedDocuments = bookletDocuments
            .Select(document => new FlatBookletDocument(document, identityByPath.GetValueOrDefault(Path.GetFullPath(document.FilePath))))
            .ToList();
        var confirmedGroups = BuildConfirmedBookletClientGroups(identifiedDocuments);
        var unconfirmedDocuments = identifiedDocuments
            .Where(item => !IsConfirmedBookletClientIdentity(item.Identity))
            .Select(item => item.Document)
            .ToList();

        if (confirmedGroups.Count == 0)
        {
            return [BuildUnassignedBookletReview(
                discovery,
                bookletDocuments,
                outputRoot,
                "No reliable client identity was found in the flat booklet folder. Nothing will be processed until the booklets are reviewed.")];
        }

        var results = new List<BatchClientDiscoveryResult>();
        foreach (var group in confirmedGroups)
        {
            var documents = group.Select(item => item.Document).ToList();
            if (confirmedGroups.Count == 1)
            {
                documents.AddRange(unconfirmedDocuments);
            }

            var clientName = ClientNameResolver.ToDisplayName(SelectCanonicalBookletIdentity(group).ClientName);
            var groupedDiscovery = BuildDiscoverySubset(discovery, documents);
            var client = BuildClientDiscovery(groupedDiscovery, outputRoot, clientName);
            if (confirmedGroups.Count == 1 && unconfirmedDocuments.Count > 0)
            {
                client.Warnings.Add(
                    $"{unconfirmedDocuments.Count} booklet file(s) did not name a client and were assigned to the only confirmed client group.");
            }

            results.Add(client);
        }

        if (confirmedGroups.Count > 1 && unconfirmedDocuments.Count > 0)
        {
            results.Add(BuildUnassignedBookletReview(
                discovery,
                unconfirmedDocuments,
                outputRoot,
                $"{unconfirmedDocuments.Count} booklet file(s) could not be assigned safely because more than one client was identified."));
        }

        return results;
    }

    private static List<List<FlatBookletDocument>> BuildConfirmedBookletClientGroups(
        IEnumerable<FlatBookletDocument> documents)
    {
        var groups = new List<List<FlatBookletDocument>>();
        foreach (var document in documents.Where(item => IsConfirmedBookletClientIdentity(item.Identity)))
        {
            var group = groups.FirstOrDefault(existing =>
                ClientIdentityKeysLikelyMatch(
                    NormalizeClientGroupKey(SelectCanonicalBookletIdentity(existing).ClientName),
                    NormalizeClientGroupKey(document.Identity!.ClientName)));
            if (group is null)
            {
                groups.Add([document]);
            }
            else
            {
                group.Add(document);
            }
        }

        return groups
            .OrderBy(group => NormalizeClientGroupKey(SelectCanonicalBookletIdentity(group).ClientName), StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static BookletClientIdentity SelectCanonicalBookletIdentity(IEnumerable<FlatBookletDocument> group)
    {
        return group
            .Where(item => item.Identity is not null)
            .Select(item => item.Identity!)
            .OrderByDescending(identity => NormalizeClientGroupKey(identity.ClientName).Length)
            .ThenBy(identity => identity.ClientName, StringComparer.OrdinalIgnoreCase)
            .First();
    }

    private static bool ClientIdentityKeysLikelyMatch(string left, string right)
    {
        return BookletClientIdentityGuardrail.GroupKeysLikelyMatch(left, right);
    }

    private static bool IsFlatBookletTestFolder(
        string parentFolder,
        string clientFolder,
        RenewalDocumentDiscoveryResult discovery)
    {
        if (!IsSameFolderPath(parentFolder, clientFolder) || discovery.Documents.Count == 0)
        {
            return false;
        }

        return discovery.Documents.All(document => document.IsBenefitBooklet);
    }

    private static bool IsSameFolderPath(string left, string right)
    {
        return string.Equals(
            Path.GetFullPath(left).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            Path.GetFullPath(right).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsConfirmedBookletClientIdentity(BookletClientIdentity? identity)
    {
        return BookletClientIdentityGuardrail.IsConfirmed(identity);
    }

    private static string NormalizeClientGroupKey(string value)
    {
        return BookletClientIdentityGuardrail.NormalizeGroupKey(value);
    }

    private static BatchClientDiscoveryResult BuildUnassignedBookletReview(
        RenewalDocumentDiscoveryResult discovery,
        IEnumerable<DiscoveredDocument> documents,
        string outputRoot,
        string warning)
    {
        var review = BuildClientDiscovery(
            BuildDiscoverySubset(discovery, documents),
            outputRoot,
            "Unassigned Booklets - Review");
        review.WorkflowAvailability = new BatchWorkflowAvailability();
        review.Status = BatchProcessingStatuses.NeedsReview;
        review.Warnings.Add(warning);
        return review;
    }

    private static RenewalDocumentDiscoveryResult BuildDiscoverySubset(
        RenewalDocumentDiscoveryResult source,
        IEnumerable<DiscoveredDocument> documents)
    {
        var subset = new RenewalDocumentDiscoveryResult
        {
            ClientFolder = source.ClientFolder
        };
        subset.Documents.AddRange(documents);
        subset.RenewalDocument = subset.Documents.FirstOrDefault(document => document.IsRenewal);
        subset.PlanDesign = DocumentDiscoveryService.SelectPlanDesignSource(subset.Documents, subset.RenewalDocument);
        return subset;
    }

    private sealed record FlatBookletDocument(DiscoveredDocument Document, BookletClientIdentity? Identity);

    private static bool IsBillingOnlyFolder(RenewalDocumentDiscoveryResult discovery)
    {
        return discovery.RenewalDocument is null &&
               discovery.Documents.Count > 0 &&
               discovery.Documents.All(document => document.IsBillingStatement);
    }

    private static string ResolveBillingDiscoveryClientName(DiscoveredDocument document)
    {
        if (ManulifeExcelBillingImportService.IsSupportedFile(document.FilePath))
        {
            var inferredClient = InferManulifeBillingClientName(document.FileName);
            return string.IsNullOrWhiteSpace(inferredClient)
                ? "Unconfirmed Manulife Billing"
                : inferredClient;
        }

        return IsGroupHealthBillingDocument(document)
            ? "GroupHEALTH Billing"
            : "Billing";
    }

    private static BatchClientDiscoveryResult BuildClientDiscovery(
        RenewalDocumentDiscoveryResult discovery,
        string outputRoot,
        string? clientNameOverride = null)
    {
        var financialClientName = discovery.Documents
            .Where(document => document.IsAsoFinancialActivity)
            .Select(document => document.ContractHolder)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
        var clientName = FirstNonBlank(
            clientNameOverride ?? "",
            financialClientName,
            Path.GetFileName(discovery.ClientFolder.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)));
        if (string.IsNullOrWhiteSpace(clientName))
        {
            clientName = "Client";
        }

        var carrier = discovery.Documents
            .Select(document => document.Carrier)
            .FirstOrDefault(carrier => !string.IsNullOrWhiteSpace(carrier) && !carrier.Equals("Unknown", StringComparison.OrdinalIgnoreCase)) ?? "Unknown";

        var result = new BatchClientDiscoveryResult
        {
            ClientName = clientName,
            ClientFolder = discovery.ClientFolder,
            Carrier = carrier,
            DocumentDiscovery = discovery,
            OutputFolder = Path.Combine(outputRoot, SafePathPart(clientName)),
            WorkflowAvailability = BuildWorkflowAvailability(discovery)
        };

        result.DocumentTypesFound.AddRange(discovery.Documents
            .Select(document => document.DocumentType)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value));
        result.DocumentCarriersFound.AddRange(discovery.Documents
            .Select(document => document.Carrier)
            .Where(value => !string.IsNullOrWhiteSpace(value) && !value.Equals("Unknown", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value));
        AddEmbeddedPlanDesignDocumentType(result, discovery);

        if (discovery.RenewalDocument is not null && !result.WorkflowAvailability.PlanDesign)
        {
            result.Warnings.Add("Plan Design Status = Not Available.");
        }

        if (result.WorkflowAvailability.AvailableWorkflows.Count == 0)
        {
            result.Status = BatchProcessingStatuses.NeedsReview;
            result.Warnings.Add("No workflow is available for this folder.");
        }
        else
        {
            result.Status = BatchProcessingStatuses.ReadyToProcess;
        }

        return result;
    }

    private static void AddEmbeddedPlanDesignDocumentType(
        BatchClientDiscoveryResult result,
        RenewalDocumentDiscoveryResult discovery)
    {
        if (discovery.PlanDesign.Source != PlanDesignSources.EmbeddedRenewal)
        {
            return;
        }

        var section = discovery.PlanDesign.SectionType.Contains("Group Information Report", StringComparison.OrdinalIgnoreCase)
            ? "Embedded GIR"
            : FirstNonBlank(discovery.PlanDesign.SectionType, "Embedded Renewal");

        if (!result.DocumentTypesFound.Contains(section, StringComparer.OrdinalIgnoreCase))
        {
            result.DocumentTypesFound.Add(section);
        }
    }

    private static BatchWorkflowAvailability BuildWorkflowAvailability(RenewalDocumentDiscoveryResult discovery)
    {
        return new BatchWorkflowAvailability
        {
            RenewalIntelligence = discovery.RenewalDocument is not null,
            ClaimsExperience = discovery.Documents.Any(document => document.IsClaimsExperience),
            Billing = discovery.Documents.Any(document => document.IsBillingStatement),
            FinancialIntelligence = discovery.Documents.Any(document => document.IsAsoFinancialActivity),
            PlanDesign = discovery.PlanDesign.Source is PlanDesignSources.SeparateGir or PlanDesignSources.EmbeddedRenewal or PlanDesignSources.PolicySchedule,
            BookletExtraction = discovery.PlanDesign.Source == PlanDesignSources.Booklet,
            ProposalComparison = discovery.Documents.Any(document => document.IsProposal)
        };
    }

    private static string WriteBatchLog(BatchProcessingResult result)
    {
        var logPath = Path.Combine(result.OutputFolder, "BatchProcessingLog.xlsx");
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Batch Log");
        var headers = new[]
        {
            "Client",
            "Carrier",
            "Document Carriers Found",
            "Document Types Found",
            "Discovery Status",
            "Renewal Processed",
            "Renewal Saved To Database",
            "PowerPoint Generated",
            "Billing Processed",
            "Billing Saved To Database",
            "Financial Intelligence Processed",
            "Financial Intelligence Saved To Database",
            "Plan Design Source",
            "Plan Design Status",
            "Plan Design Extracted",
            "Plan Design Rows Saved",
            "Plan Design Target Table",
            "Schedule_of_Benefits Updated",
            "Database Save Status",
            "Workflows Executed",
            "Status",
            "Warnings",
            "Errors",
            "Processing Time",
            "Output Folder"
        };

        for (var index = 0; index < headers.Length; index++)
        {
            sheet.Cell(1, index + 1).Value = headers[index];
        }

        for (var index = 0; index < result.Clients.Count; index++)
        {
            var row = result.Clients[index];
            var rowNumber = index + 2;
            sheet.Cell(rowNumber, 1).Value = row.ClientName;
            sheet.Cell(rowNumber, 2).Value = row.Carrier;
            sheet.Cell(rowNumber, 3).Value = row.DocumentCarriersFound;
            sheet.Cell(rowNumber, 4).Value = row.DocumentTypesFound;
            sheet.Cell(rowNumber, 5).Value = row.DiscoveryStatus;
            sheet.Cell(rowNumber, 6).Value = YesNo(row.RenewalProcessed);
            sheet.Cell(rowNumber, 7).Value = YesNo(row.RenewalSavedToDatabase);
            sheet.Cell(rowNumber, 8).Value = YesNo(row.PowerPointGenerated);
            sheet.Cell(rowNumber, 9).Value = YesNo(row.BillingProcessed);
            sheet.Cell(rowNumber, 10).Value = YesNo(row.BillingSavedToDatabase);
            sheet.Cell(rowNumber, 11).Value = YesNo(row.FinancialIntelligenceProcessed);
            sheet.Cell(rowNumber, 12).Value = YesNo(row.FinancialIntelligenceSavedToDatabase);
            sheet.Cell(rowNumber, 13).Value = row.PlanDesignSource;
            sheet.Cell(rowNumber, 14).Value = row.PlanDesignStatus;
            sheet.Cell(rowNumber, 15).Value = YesNo(row.PlanDesignExtracted);
            sheet.Cell(rowNumber, 16).Value = row.PlanDesignRowsSaved;
            sheet.Cell(rowNumber, 17).Value = row.PlanDesignTargetTable;
            sheet.Cell(rowNumber, 18).Value = YesNo(row.ScheduleOfBenefitsUpdated);
            sheet.Cell(rowNumber, 19).Value = row.DatabaseSaveStatus;
            sheet.Cell(rowNumber, 20).Value = row.WorkflowsExecuted;
            sheet.Cell(rowNumber, 21).Value = row.Status;
            sheet.Cell(rowNumber, 22).Value = row.WarningsText;
            sheet.Cell(rowNumber, 23).Value = row.ErrorsText;
            sheet.Cell(rowNumber, 24).Value = FormatDuration(row.ProcessingTime);
            sheet.Cell(rowNumber, 25).Value = row.OutputFolder;
        }

        var header = sheet.Range(1, 1, 1, headers.Length);
        header.Style.Fill.BackgroundColor = XLColor.FromHtml("#0E7C7B");
        header.Style.Font.FontColor = XLColor.White;
        header.Style.Font.Bold = true;
        sheet.Columns().AdjustToContents();

        WriteWorkflowDetailSheet(workbook, result.Clients);
        WriteSummarySheet(workbook, result);
        workbook.SaveAs(logPath);
        return logPath;
    }

    private static void WriteWorkflowDetailSheet(
        XLWorkbook workbook,
        IReadOnlyList<BatchClientProcessingResult> clients)
    {
        var sheet = workbook.AddWorksheet("Workflow Details");
        var headers = new[]
        {
            "Client",
            "Workflow",
            "File",
            "Status",
            "Message",
            "Saved To Database",
            "Database Rows Saved",
            "Plan Design Rows Extracted",
            "Plan Design Rows Saved",
            "Plan Design Target Table",
            "Schedule_of_Benefits Updated",
            "Database Save Status",
            "Exception",
            "Stack Trace",
            "Processing Time",
            "Outputs"
        };

        for (var index = 0; index < headers.Length; index++)
        {
            sheet.Cell(1, index + 1).Value = headers[index];
        }

        var rowNumber = 2;
        foreach (var client in clients)
        {
            foreach (var workflow in client.Workflows)
            {
                sheet.Cell(rowNumber, 1).Value = client.ClientName;
                sheet.Cell(rowNumber, 2).Value = workflow.WorkflowName;
                sheet.Cell(rowNumber, 3).Value = workflow.SourceFile;
                sheet.Cell(rowNumber, 4).Value = workflow.Status;
                sheet.Cell(rowNumber, 5).Value = workflow.Message;
                sheet.Cell(rowNumber, 6).Value = YesNo(workflow.SavedToDatabase);
                sheet.Cell(rowNumber, 7).Value = workflow.DatabaseRowsSaved;
                sheet.Cell(rowNumber, 8).Value = workflow.PlanDesignRowsExtracted;
                sheet.Cell(rowNumber, 9).Value = workflow.PlanDesignRowsSaved;
                sheet.Cell(rowNumber, 10).Value = workflow.PlanDesignTargetTable;
                sheet.Cell(rowNumber, 11).Value = YesNo(workflow.ScheduleOfBenefitsUpdated);
                sheet.Cell(rowNumber, 12).Value = workflow.DatabaseSaveStatus;
                sheet.Cell(rowNumber, 13).Value = workflow.Exception;
                sheet.Cell(rowNumber, 14).Value = workflow.StackTrace;
                sheet.Cell(rowNumber, 15).Value = FormatDuration(workflow.ProcessingTime);
                sheet.Cell(rowNumber, 16).Value = string.Join(" | ", workflow.OutputPaths);
                rowNumber++;
            }
        }

        var header = sheet.Range(1, 1, 1, headers.Length);
        header.Style.Fill.BackgroundColor = XLColor.FromHtml("#0E7C7B");
        header.Style.Font.FontColor = XLColor.White;
        header.Style.Font.Bold = true;
        sheet.Columns().AdjustToContents();
    }

    private static void WriteSummarySheet(XLWorkbook workbook, BatchProcessingResult result)
    {
        var sheet = workbook.AddWorksheet("Summary");
        var rows = new (string Metric, object Value)[]
        {
            ("Clients discovered", result.Clients.Count),
            ("Renewals processed", result.Clients.Count(client => client.RenewalProcessed)),
            ("Renewals saved to database", result.Clients.Count(client => client.RenewalSavedToDatabase)),
            ("PowerPoints generated", result.Clients.Count(client => client.PowerPointGenerated)),
            ("Plan design sources found", result.Clients.Count(client => !client.PlanDesignSource.Equals(PlanDesignSources.None, StringComparison.OrdinalIgnoreCase))),
            ("Plan design rows saved", result.Clients.Sum(client => client.PlanDesignRowsSaved)),
            ("Clients with warnings", result.Clients.Count(client => client.Status == BatchProcessingStatuses.CompletedWithWarnings)),
            ("Clients failed", result.Clients.Count(client => client.Status == BatchProcessingStatuses.Failed))
        };

        sheet.Cell(1, 1).Value = "Metric";
        sheet.Cell(1, 2).Value = "Value";
        for (var index = 0; index < rows.Length; index++)
        {
            sheet.Cell(index + 2, 1).Value = rows[index].Metric;
            sheet.Cell(index + 2, 2).Value = rows[index].Value.ToString();
        }

        var header = sheet.Range(1, 1, 1, 2);
        header.Style.Fill.BackgroundColor = XLColor.FromHtml("#0E7C7B");
        header.Style.Font.FontColor = XLColor.White;
        header.Style.Font.Bold = true;
        sheet.Columns().AdjustToContents();
    }

    private async Task ValidateClientOutputsAsync(
        BatchClientDiscoveryResult client,
        BatchClientProcessingResult result,
        BatchClientProcessingState state)
    {
        var renewalWorkflow = result.Workflows.FirstOrDefault(workflow =>
            workflow.WorkflowName.Equals(BatchWorkflowNames.RenewalIntelligence, StringComparison.OrdinalIgnoreCase));
        if (renewalWorkflow is not null && renewalWorkflow.Status != BatchProcessingStatuses.Skipped)
        {
            result.RenewalProcessed = renewalWorkflow.Status != BatchProcessingStatuses.Failed;
            result.RenewalWorkbookGenerated = renewalWorkflow.WorkbookGenerated ||
                                              renewalWorkflow.OutputPaths.Any(IsExcelWorkbookPath);
            result.PowerPointGenerated = renewalWorkflow.PowerPointGenerated ||
                                         renewalWorkflow.OutputPaths.Any(IsPowerPointPath);

            var renewalRows = 0;
            if (state.RenewalDatabaseSave?.ClientId is not null && state.RenewalDatabaseSave.DocumentId is not null)
            {
                renewalRows = await CountRowsAsync("Renewals", state.RenewalDatabaseSave.ClientId, state.RenewalDatabaseSave.DocumentId);
            }

            result.RenewalSavedToDatabase = state.RenewalDatabaseSave?.Succeeded == true && renewalRows > 0;
            if (!result.RenewalWorkbookGenerated)
            {
                result.Warnings.Add("Renewal workflow ran, but Renewal Review workbook was not confirmed.");
            }

            if (!result.PowerPointGenerated)
            {
                result.Warnings.Add("Renewal workflow ran, but Renewal PowerPoint was not confirmed.");
            }

            if (result.RenewalProcessed && !result.RenewalSavedToDatabase)
            {
                result.Warnings.Add("Renewal output generated but renewal database save not confirmed.");
            }
        }

        var planWorkflow = result.Workflows.FirstOrDefault(workflow =>
            workflow.WorkflowName.Equals(BatchWorkflowNames.PlanDesign, StringComparison.OrdinalIgnoreCase));
        var bookletWorkflow = result.Workflows.FirstOrDefault(workflow =>
            workflow.WorkflowName.Equals(BatchWorkflowNames.BookletExtraction, StringComparison.OrdinalIgnoreCase));
        var proposalWorkflow = result.Workflows.FirstOrDefault(workflow =>
            workflow.WorkflowName.Equals(BatchWorkflowNames.ProposalComparison, StringComparison.OrdinalIgnoreCase));
        var billingWorkflows = result.Workflows
            .Where(workflow => workflow.WorkflowName.Equals(BatchWorkflowNames.Billing, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var financialWorkflows = result.Workflows
            .Where(workflow => workflow.WorkflowName.Equals(BatchWorkflowNames.FinancialIntelligence, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (billingWorkflows.Count > 0)
        {
            result.BillingProcessed = billingWorkflows.Any(workflow => workflow.Status != BatchProcessingStatuses.Failed);
            result.BillingSavedToDatabase = billingWorkflows.Any(workflow => workflow.SavedToDatabase);
            if (result.BillingProcessed && !result.BillingSavedToDatabase)
            {
                result.Warnings.Add("Billing workflow ran, but no billing statement save was confirmed.");
            }
        }

        if (financialWorkflows.Count > 0)
        {
            result.FinancialIntelligenceProcessed = financialWorkflows.Any(
                workflow => workflow.Status != BatchProcessingStatuses.Failed);
            result.FinancialIntelligenceSavedToDatabase = financialWorkflows.Any(
                workflow => workflow.SavedToDatabase);
            if (result.FinancialIntelligenceProcessed && !result.FinancialIntelligenceSavedToDatabase)
            {
                result.Warnings.Add(
                    "Financial Intelligence workflow ran, but no ASO financial statement save was confirmed.");
            }
        }

        if (planWorkflow is not null && planWorkflow.Status != BatchProcessingStatuses.Skipped)
        {
            result.PlanDesignExtracted = planWorkflow.PlanDesignRowsExtracted > 0;
            result.PlanDesignRowsSaved = planWorkflow.PlanDesignRowsSaved;
            result.PlanDesignTargetTable = FirstNonBlank(planWorkflow.PlanDesignTargetTable, "plan_design_features");
            result.ScheduleOfBenefitsUpdated = planWorkflow.ScheduleOfBenefitsUpdated;
            if (result.PlanDesignRowsSaved == 0)
            {
                var pages = FormatPageList(client.DocumentDiscovery.PlanDesign.PageNumbers);
                result.Warnings.Add($"{client.DocumentDiscovery.PlanDesign.Source} detected but no plan design rows were saved. Source page(s): {pages}.");
            }
        }
        else if (bookletWorkflow is not null && bookletWorkflow.Status != BatchProcessingStatuses.Skipped)
        {
            result.PlanDesignExtracted = bookletWorkflow.PlanDesignRowsExtracted > 0;
            result.PlanDesignRowsSaved = bookletWorkflow.PlanDesignRowsSaved;
            result.PlanDesignTargetTable = FirstNonBlank(bookletWorkflow.PlanDesignTargetTable, "PlanDesign");
            result.ScheduleOfBenefitsUpdated = bookletWorkflow.ScheduleOfBenefitsUpdated;
            if (result.PlanDesignRowsSaved == 0)
            {
                result.Warnings.Add("Booklet workflow ran, but no PlanDesign rows were confirmed.");
            }
        }
        else if (proposalWorkflow is not null &&
                 proposalWorkflow.Status != BatchProcessingStatuses.Skipped &&
                 result.PlanDesignSource.Equals(PlanDesignSources.Proposal, StringComparison.OrdinalIgnoreCase))
        {
            result.PlanDesignExtracted = proposalWorkflow.Status != BatchProcessingStatuses.Failed;
            result.PlanDesignTargetTable = FirstNonBlank(proposalWorkflow.PlanDesignTargetTable, "proposal_plan_design");
        }

        result.DatabaseSaveStatus = string.Join(" | ", result.Workflows
            .Select(workflow => workflow.DatabaseSaveStatus)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase));
    }

    private async Task<int> CountRowsAsync(string tableName, long? clientId, long? sourceDocumentId)
    {
        if (string.IsNullOrWhiteSpace(_databaseService.DatabasePath) || !File.Exists(_databaseService.DatabasePath))
        {
            return 0;
        }

        var safeTableName = tableName switch
        {
            "Renewals" => "Renewals",
            "Rates" => "Rates",
            "PlanDesign" => "PlanDesign",
            "plan_design_features" => "plan_design_features",
            _ => throw new ArgumentOutOfRangeException(nameof(tableName), tableName, "Unsupported validation table.")
        };

        await using var connection = new SqliteConnection($"Data Source={_databaseService.DatabasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        var conditions = new List<string>();
        if (clientId.HasValue)
        {
            conditions.Add("client_id = $client_id");
            command.Parameters.AddWithValue("$client_id", clientId.Value);
        }

        if (sourceDocumentId.HasValue)
        {
            conditions.Add("source_document_id = $source_document_id");
            command.Parameters.AddWithValue("$source_document_id", sourceDocumentId.Value);
        }

        command.CommandText = conditions.Count == 0
            ? $"SELECT COUNT(*) FROM {safeTableName};"
            : $"SELECT COUNT(*) FROM {safeTableName} WHERE {string.Join(" AND ", conditions)};";
        var value = await command.ExecuteScalarAsync();
        return Convert.ToInt32(value, CultureInfo.InvariantCulture);
    }

    private static List<PdfPageText> SelectPlanDesignPages(
        PlanDesignSourceDiscovery source,
        IReadOnlyList<PdfPageText> pages)
    {
        if (source.Source == PlanDesignSources.SeparateGir || source.PageNumbers.Count == 0)
        {
            return pages.ToList();
        }

        var firstPage = source.PageNumbers.Min();
        var lastDetectedPage = source.PageNumbers.Max();
        var maxPage = pages.Count == 0 ? lastDetectedPage : pages.Max(page => page.PageNumber);
        var lastPage = Math.Min(maxPage, Math.Max(lastDetectedPage, firstPage + 8));
        return pages
            .Where(page => page.PageNumber >= firstPage && page.PageNumber <= lastPage)
            .OrderBy(page => page.PageNumber)
            .ToList();
    }

    private static string WriteScheduleOfBenefitsAdapter(
        BatchClientDiscoveryResult client,
        PlanDesignSourceDiscovery source,
        IReadOnlyList<PlanDesignFeatureExtraction> features)
    {
        var path = Path.Combine(client.OutputFolder, "Schedule_of_Benefits.xlsx");
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Schedule_of_Benefits");
        var headers = new[]
        {
            "Slide Page",
            "Page Total",
            "Section Order",
            "Client Name",
            "Carrier",
            "Class Name",
            "Section Name",
            "Detail Text",
            "Source File",
            "Notes"
        };

        sheet.Cell(1, 1).Value = "Schedule of Benefits";
        sheet.Range(1, 1, 1, headers.Length).Merge();
        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontSize = 16;
        sheet.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#004751");

        for (var index = 0; index < headers.Length; index++)
        {
            sheet.Cell(3, index + 1).Value = headers[index];
        }

        var groupedRows = features
            .GroupBy(feature => new
            {
                FeatureClass = FirstNonBlank(feature.ClassName, feature.ClassNumber, "Class 1"),
                feature.Benefit
            })
            .OrderBy(group => group.Key.FeatureClass, StringComparer.OrdinalIgnoreCase)
            .ThenBy(group => group.Key.Benefit, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var pageTotal = Math.Max(1, (int)Math.Ceiling(groupedRows.Count / 8m));
        var rowNumber = 4;
        var sectionOrder = 1;
        foreach (var group in groupedRows)
        {
            var slidePage = Math.Max(1, (int)Math.Ceiling(sectionOrder / 8m));
            var details = string.Join(Environment.NewLine, group
                .Select(feature => $"{feature.FeatureName}: {feature.FeatureValue}")
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase));
            sheet.Cell(rowNumber, 1).Value = slidePage;
            sheet.Cell(rowNumber, 2).Value = pageTotal;
            sheet.Cell(rowNumber, 3).Value = sectionOrder;
            sheet.Cell(rowNumber, 4).Value = client.ClientName;
            sheet.Cell(rowNumber, 5).Value = client.Carrier;
            sheet.Cell(rowNumber, 6).Value = group.Key.FeatureClass;
            sheet.Cell(rowNumber, 7).Value = FirstNonBlank(group.Key.Benefit, source.SectionType, "Schedule of Benefits");
            sheet.Cell(rowNumber, 8).Value = details;
            sheet.Cell(rowNumber, 9).Value = source.FileName;
            sheet.Cell(rowNumber, 10).Value = $"{source.Source}; source page(s): {FormatPageList(group.Select(feature => feature.SourcePage))}";
            rowNumber++;
            sectionOrder++;
        }

        var range = sheet.Range(3, 1, Math.Max(3, rowNumber - 1), headers.Length);
        var table = range.CreateTable("tblPPT_ScheduleOfBenefits");
        table.Theme = XLTableTheme.TableStyleMedium4;
        sheet.Columns().AdjustToContents();
        sheet.Column(8).Width = 60;
        sheet.Column(8).Style.Alignment.WrapText = true;

        var raw = workbook.AddWorksheet("plan_design_features");
        var rawHeaders = new[]
        {
            "Benefit",
            "Category",
            "Feature Name",
            "Feature Value",
            "Coinsurance",
            "Deductible",
            "Maximum",
            "Frequency",
            "Termination Age",
            "Class Number",
            "Class Name",
            "Source Page",
            "Source Text",
            "Source File",
            "Plan Design Source",
            "Confidence"
        };
        for (var index = 0; index < rawHeaders.Length; index++)
        {
            raw.Cell(1, index + 1).Value = rawHeaders[index];
        }

        for (var index = 0; index < features.Count; index++)
        {
            var feature = features[index];
            var row = index + 2;
            raw.Cell(row, 1).Value = feature.Benefit;
            raw.Cell(row, 2).Value = feature.Category;
            raw.Cell(row, 3).Value = feature.FeatureName;
            raw.Cell(row, 4).Value = feature.FeatureValue;
            raw.Cell(row, 5).Value = feature.Coinsurance;
            raw.Cell(row, 6).Value = feature.Deductible;
            raw.Cell(row, 7).Value = feature.Maximum;
            raw.Cell(row, 8).Value = feature.Frequency;
            raw.Cell(row, 9).Value = feature.TerminationAge;
            raw.Cell(row, 10).Value = feature.ClassNumber;
            raw.Cell(row, 11).Value = feature.ClassName;
            raw.Cell(row, 12).Value = feature.SourcePage;
            raw.Cell(row, 13).Value = feature.SourceText;
            raw.Cell(row, 14).Value = feature.SourceFileName;
            raw.Cell(row, 15).Value = feature.PlanDesignSource;
            raw.Cell(row, 16).Value = feature.ConfidenceScore;
        }

        raw.Columns().AdjustToContents();
        workbook.SaveAs(path);
        return path;
    }

    private static BatchProcessingProgress BuildProgress(
        BatchClientDiscoveryResult client,
        string workflow,
        int completed,
        int total,
        BatchProcessingResult result,
        Stopwatch batchWatch)
    {
        return new BatchProcessingProgress
        {
            CurrentClient = client.ClientName,
            CurrentWorkflow = workflow,
            Completed = completed,
            Total = total,
            WarningCount = result.Clients.Count(row => row.Status == BatchProcessingStatuses.CompletedWithWarnings),
            FailedCount = result.Clients.Count(row => row.Status == BatchProcessingStatuses.Failed),
            EstimatedTimeRemaining = EstimateRemaining(batchWatch.Elapsed, completed, total),
            Message = $"{client.ClientName}: {workflow}"
        };
    }

    private static string DetermineClientStatus(BatchClientProcessingResult result)
    {
        if (result.Workflows.Count == 0)
        {
            return BatchProcessingStatuses.Skipped;
        }

        if (result.Workflows.All(workflow => workflow.Status == BatchProcessingStatuses.Failed))
        {
            return BatchProcessingStatuses.Failed;
        }

        return result.Errors.Count > 0 ||
               result.Warnings.Count > 0 ||
               result.Workflows.Any(workflow => workflow.Status == BatchProcessingStatuses.Failed) ||
               result.Workflows.Any(workflow => workflow.Status == BatchProcessingStatuses.CompletedWithWarnings)
            ? BatchProcessingStatuses.CompletedWithWarnings
            : BatchProcessingStatuses.Completed;
    }

    private static TimeSpan EstimateRemaining(TimeSpan elapsed, int completed, int total)
    {
        if (completed <= 0 || total <= completed)
        {
            return TimeSpan.Zero;
        }

        var averageTicks = elapsed.Ticks / completed;
        return TimeSpan.FromTicks(averageTicks * (total - completed));
    }

    private static List<string> SupportingPdfFiles(BatchClientDiscoveryResult client, string renewalPath)
    {
        return client.DocumentDiscovery.Documents
            .Where(document => document.FileExtension.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            .Select(document => document.FilePath)
            .Where(path => !PathsEqual(path, renewalPath))
            .Where(path => !IsGeneratedOutputPath(client.ClientFolder, path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> ClaimsFiles(BatchClientDiscoveryResult client)
    {
        return client.DocumentDiscovery.Documents
            .Where(document => document.IsClaimsExperience)
            .Select(document => document.FilePath)
            .Where(File.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<DiscoveredDocument> BillingFiles(BatchClientDiscoveryResult client)
    {
        return client.DocumentDiscovery.Documents
            .Where(document => document.IsBillingStatement)
            .Where(document => File.Exists(document.FilePath))
            .OrderBy(document => document.FilePath, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<DiscoveredDocument> AsoFinancialActivityFiles(BatchClientDiscoveryResult client)
    {
        return client.DocumentDiscovery.Documents
            .Where(document => document.IsAsoFinancialActivity)
            .Where(document => File.Exists(document.FilePath))
            .OrderBy(document => document.FilePath, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsGroupHealthBillingDocument(DiscoveredDocument document)
    {
        return document.Carrier.Contains("grouphealth", StringComparison.OrdinalIgnoreCase) ||
               document.FileName.Contains("grouphealth", StringComparison.OrdinalIgnoreCase);
    }

    private static string ResolveManulifeBillingClientName(BatchClientDiscoveryResult client, DiscoveredDocument document)
    {
        if (!IsGenericBillingClientName(client.ClientName))
        {
            return client.ClientName;
        }

        return InferManulifeBillingClientName(document.FileName);
    }

    private static bool IsGenericBillingClientName(string value)
    {
        return string.IsNullOrWhiteSpace(value) ||
               value.Equals("Billing", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Billings", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Batch", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Batch Processing", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Batch Intake", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Manulife Billing", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Unconfirmed Manulife Billing", StringComparison.OrdinalIgnoreCase);
    }

    private static string InferManulifeBillingClientName(string fileName)
    {
        var sourceName = Path.GetFileNameWithoutExtension(fileName).Trim();
        if (string.IsNullOrWhiteSpace(sourceName))
        {
            return "";
        }

        var match = Regex.Match(
            sourceName,
            @"^(?<client>.+?)\s*[-_]\s*(?:div(?:ision)?|class|plan)\s*[A-Za-z0-9]+\s*[-_]\s*(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\s*20\d{2}\b",
            RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return "";
        }

        return match.Groups["client"].Value.Trim().Trim('-', '_').Trim();
    }

    private static List<string> BookletFiles(BatchClientDiscoveryResult client)
    {
        return client.DocumentDiscovery.Documents
            .Where(document => document.IsBenefitBooklet)
            .Select(document => document.FilePath)
            .Where(File.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> ProposalFiles(BatchClientDiscoveryResult client)
    {
        return client.DocumentDiscovery.Documents
            .Where(document => document.IsProposal)
            .Select(document => document.FilePath)
            .Where(File.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> WorkbookFiles(BatchClientDiscoveryResult client)
    {
        return client.DocumentDiscovery.Documents
            .Where(document => !document.IsAsoFinancialActivity)
            .Where(document => document.FileExtension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                               document.FileExtension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                               document.FileExtension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            .Select(document => document.FilePath)
            .Where(File.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    internal static string SelectPrimaryRenewalWorkbook(IEnumerable<string> workbookPaths)
    {
        return workbookPaths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Where(path =>
            {
                var fileName = Path.GetFileNameWithoutExtension(path);
                return fileName.Contains("Client_Copy", StringComparison.OrdinalIgnoreCase) ||
                       fileName.Contains("Client Copy", StringComparison.OrdinalIgnoreCase) ||
                       fileName.Contains("Client-Copy", StringComparison.OrdinalIgnoreCase);
            })
            // Manulife commonly leaves the current Client_Copy undated and retains dated archive copies beside it.
            .OrderByDescending(path => LatestYearInFileName(Path.GetFileNameWithoutExtension(path)) == 0)
            .ThenByDescending(path => LatestYearInFileName(Path.GetFileNameWithoutExtension(path)))
            .ThenByDescending(path => File.Exists(path) ? File.GetLastWriteTimeUtc(path) : DateTime.MinValue)
            .FirstOrDefault() ?? "";
    }

    private static int LatestYearInFileName(string fileName)
    {
        return Regex.Matches(fileName, @"(?<!\d)20\d{2}(?!\d)")
            .Select(match => int.Parse(match.Value, CultureInfo.InvariantCulture))
            .DefaultIfEmpty(0)
            .Max();
    }

    private async Task<string> ResolveSelectedFolderClientNameAsync(string parentFolder)
    {
        var childFolders = Directory
            .EnumerateDirectories(parentFolder, "*", SearchOption.TopDirectoryOnly)
            .Where(path => !IsGeneratedOutputFolder(path))
            .ToList();
        if (childFolders.Count == 0)
        {
            return "";
        }

        var selectedFolderName = ClientNameResolver.ToDisplayName(
            Path.GetFileName(parentFolder.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)));
        if (IsGenericBatchContainerName(selectedFolderName))
        {
            return "";
        }

        var bookletCandidates = Directory
            .EnumerateFiles(parentFolder, "*.pdf", SearchOption.AllDirectories)
            .Where(path => !IsGeneratedOutputPath(parentFolder, path))
            .Where(path => Regex.IsMatch(
                Path.GetFileNameWithoutExtension(path),
                @"\b(booklet|handbook|benefits?\s*(?:book|summary))\b",
                RegexOptions.IgnoreCase))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .Take(4)
            .ToList();
        if (bookletCandidates.Count == 0)
        {
            return "";
        }

        IReadOnlyList<BookletClientIdentity> identities;
        try
        {
            identities = await _bookletClientIdentityProbe(bookletCandidates, CancellationToken.None);
        }
        catch (Exception exception)
        {
            AppLogger.Error(exception);
            return "";
        }

        var matches = identities
            .Where(IsConfirmedBookletClientIdentity)
            .Where(identity => SelectedFolderMatchesBookletIdentity(selectedFolderName, identity.ClientName))
            .GroupBy(identity => NormalizeClientGroupKey(identity.ClientName), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.OrderByDescending(identity => identity.ConfidenceScore).First())
            .ToList();

        return matches.Count == 1
            ? ClientNameResolver.ToDisplayName(matches[0].ClientName)
            : "";
    }

    private static bool SelectedFolderMatchesBookletIdentity(string selectedFolderName, string detectedClientName)
    {
        if (BookletClientIdentityGuardrail.ClientNamesLikelyMatch(selectedFolderName, detectedClientName))
        {
            return true;
        }

        var selectedKey = NormalizeClientGroupKey(selectedFolderName);
        var detectedKey = NormalizeClientGroupKey(detectedClientName);
        return selectedKey.Length >= 6 && detectedKey.Contains(selectedKey, StringComparison.Ordinal);
    }

    private static bool IsGenericBatchContainerName(string value)
    {
        var normalized = NormalizeClientGroupKey(value);
        return normalized is "BATCH" or "BATCHPROCESSING" or "BOOKLETTEST" or "CLIENTS" or
               "CLIENTFOLDERS" or "DOCUMENTS" or "DOCUMENTINTAKE" or "INPUT" or "INPUTPDFS" or
               "UPLOAD" or "UPLOADS" or "TESTDATA";
    }

    private static List<string> ResolveClientFolders(string parentFolder)
    {
        var childFolders = Directory
            .EnumerateDirectories(parentFolder, "*", SearchOption.TopDirectoryOnly)
            .Where(path => !IsGeneratedOutputFolder(path))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (childFolders.Count > 0)
        {
            return childFolders;
        }

        return [parentFolder];
    }

    private static string ResolveOutputFolder(BatchProcessingRequest request)
    {
        if (!string.IsNullOrWhiteSpace(request.OutputFolder))
        {
            return request.OutputFolder;
        }

        return string.IsNullOrWhiteSpace(request.ParentFolder)
            ? Path.Combine(AppContext.BaseDirectory, "Batch Output")
            : Path.Combine(request.ParentFolder, "Batch Output");
    }

    private static bool IsGeneratedOutputFolder(string folderPath)
    {
        var folderName = Path.GetFileName(folderPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        return IsGeneratedOutputFolderName(folderName);
    }

    private static bool IsGeneratedOutputPath(string clientFolder, string filePath)
    {
        if (string.IsNullOrWhiteSpace(clientFolder) || string.IsNullOrWhiteSpace(filePath))
        {
            return false;
        }

        var relativePath = Path.GetRelativePath(clientFolder, filePath);
        var parts = relativePath.Split(
            [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
            StringSplitOptions.RemoveEmptyEntries);

        return parts.Any(IsGeneratedOutputFolderName);
    }

    private static bool IsGeneratedOutputFolderName(string folderName)
    {
        return folderName.Equals("Batch Output", StringComparison.OrdinalIgnoreCase) ||
               folderName.Equals("Supporting PDFs", StringComparison.OrdinalIgnoreCase) ||
               folderName.Equals("Supporting Excel Files", StringComparison.OrdinalIgnoreCase);
    }

    private static BatchWorkflowResult Skipped(string workflow, string message)
    {
        return new BatchWorkflowResult
        {
            WorkflowName = workflow,
            Status = BatchProcessingStatuses.Skipped,
            Message = message
        };
    }

    private static void AddExistingOutput(BatchWorkflowResult workflow, string? path)
    {
        if (!string.IsNullOrWhiteSpace(path) && (File.Exists(path) || Directory.Exists(path)))
        {
            workflow.OutputPaths.Add(path);
        }
    }

    private static void ApplyClientCodeToWorkflowOutputs(BatchWorkflowResult workflow, string clientName)
    {
        var clientCode = BuildClientFileCode(clientName);
        if (string.IsNullOrWhiteSpace(clientCode))
        {
            return;
        }

        for (var index = 0; index < workflow.OutputPaths.Count; index++)
        {
            var outputPath = workflow.OutputPaths[index];
            if (!File.Exists(outputPath))
            {
                continue;
            }

            try
            {
                workflow.OutputPaths[index] = RenameFileWithClientCode(outputPath, clientCode);
            }
            catch (Exception ex)
            {
                workflow.Warnings.Add($"Could not add client code to {Path.GetFileName(outputPath)}: {ex.Message}");
            }
        }
    }

    private static string RenameFileWithClientCode(string path, string clientCode)
    {
        var folder = Path.GetDirectoryName(path);
        if (string.IsNullOrWhiteSpace(folder))
        {
            return path;
        }

        var extension = Path.GetExtension(path);
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);
        var suffix = $"_{clientCode}";
        if (fileNameWithoutExtension.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
        {
            return path;
        }

        var destination = UniqueClientCodePath(folder, $"{fileNameWithoutExtension}{suffix}", extension);
        File.Move(path, destination);
        return destination;
    }

    private static string UniqueClientCodePath(string folder, string fileNameWithoutExtension, string extension)
    {
        var destination = Path.Combine(folder, $"{fileNameWithoutExtension}{extension}");
        if (!File.Exists(destination))
        {
            return destination;
        }

        for (var index = 2; index < 1000; index++)
        {
            var candidate = Path.Combine(folder, $"{fileNameWithoutExtension}_{index}{extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }

        return Path.Combine(folder, $"{fileNameWithoutExtension}_{DateTime.Now:yyyyMMddHHmmss}{extension}");
    }

    private static string BuildClientFileCode(string clientName)
    {
        var firstWord = FirstNonBlank(clientName.Split(' ', StringSplitOptions.RemoveEmptyEntries));
        var letters = new string(firstWord
            .Where(char.IsLetterOrDigit)
            .Take(3)
            .ToArray());

        if (letters.Length < 3)
        {
            letters += new string(clientName
                .Where(char.IsLetterOrDigit)
                .Skip(letters.Length)
                .Take(3 - letters.Length)
                .ToArray());
        }

        return string.IsNullOrWhiteSpace(letters)
            ? "CLI"
            : letters.ToUpperInvariant();
    }

    private static string FirstFile(IEnumerable<string> paths)
    {
        return paths.FirstOrDefault() ?? "";
    }

    private static bool IsExcelWorkbookPath(string path)
    {
        return File.Exists(path) &&
               (Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                Path.GetExtension(path).Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsPowerPointPath(string path)
    {
        return File.Exists(path) &&
               Path.GetExtension(path).Equals(".pptx", StringComparison.OrdinalIgnoreCase);
    }

    private static string YesNo(bool value)
    {
        return value ? "Yes" : "No";
    }

    private static string FormatPageList(IEnumerable<int?> pageNumbers)
    {
        return FormatPageList(pageNumbers
            .Where(page => page.HasValue)
            .Select(page => page!.Value));
    }

    private static string FormatPageList(IEnumerable<int> pageNumbers)
    {
        var pages = pageNumbers
            .Distinct()
            .OrderBy(page => page)
            .Select(page => page.ToString(CultureInfo.InvariantCulture))
            .ToList();
        return pages.Count == 0 ? "not captured" : string.Join(", ", pages);
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private static string SafePathPart(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(value.Select(ch => invalid.Contains(ch) ? '-' : ch).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? "Client" : cleaned;
    }

    private static bool PathsEqual(string left, string right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
        {
            return false;
        }

        return string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
    }

    private static string FormatDuration(TimeSpan duration)
    {
        return duration.TotalSeconds < 1
            ? "<1 sec"
            : duration.TotalMinutes < 1
                ? $"{duration.TotalSeconds:0} sec"
                : $"{duration.TotalMinutes:0.0} min";
    }

    private sealed class BatchClientProcessingState
    {
        public AnalysisResult? RenewalAnalysis { get; set; }
        public DatabaseSaveResult? RenewalDatabaseSave { get; set; }
        public DatabaseSaveResult? GenericArtifactSave { get; set; }
        public List<PlanDesignFeatureExtraction> PlanDesignFeatures { get; } = [];
        public List<int> PlanDesignSourcePages { get; } = [];
    }
}
