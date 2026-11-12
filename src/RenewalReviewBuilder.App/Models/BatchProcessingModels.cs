namespace RenewalReviewBuilder.App.Models;

public static class BatchProcessingStatuses
{
    public const string Discovered = "Discovered";
    public const string ReadyToProcess = "Ready to Process";
    public const string Ready = "Ready";
    public const string Processing = "Processing";
    public const string Completed = "Completed";
    public const string CompletedWithWarnings = "Completed With Warnings";
    public const string NeedsReview = "Needs Review";
    public const string Failed = "Failed";
    public const string Skipped = "Skipped";
}

public static class BatchWorkflowNames
{
    public const string RenewalIntelligence = "Renewal Intelligence";
    public const string ClaimsExperience = "Claims Experience";
    public const string ClaimsPowerPoint = "Claims Experience PowerPoint";
    public const string Billing = "Billing Import";
    public const string FinancialIntelligence = "Financial Intelligence";
    public const string PlanDesign = "Plan Design";
    public const string BookletExtraction = "Booklet Extraction";
    public const string ProposalComparison = "Proposal Comparison";
}

public sealed class BatchProcessingRequest
{
    public string ParentFolder { get; set; } = "";
    public string OutputFolder { get; set; } = "";
    public string RenewalPowerPointTemplatePath { get; set; } = "";
    public string ClaimsPowerPointTemplatePath { get; set; } = "";
    public int ReportMonthCount { get; set; }
    public string DatabasePath { get; set; } = "";
}

public sealed class BatchDiscoveryResult
{
    public string ParentFolder { get; set; } = "";
    public string OutputFolder { get; set; } = "";
    public List<BatchClientDiscoveryResult> Clients { get; } = [];
    public List<string> Warnings { get; } = [];
    public string Status => Clients.Count == 0
        ? BatchProcessingStatuses.NeedsReview
        : Clients.Any(client => client.Status == BatchProcessingStatuses.ReadyToProcess)
            ? BatchProcessingStatuses.ReadyToProcess
            : BatchProcessingStatuses.NeedsReview;
}

public sealed class BatchClientDiscoveryResult
{
    public string ClientName { get; set; } = "";
    public string ClientFolder { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string Status { get; set; } = BatchProcessingStatuses.NeedsReview;
    public string OutputFolder { get; set; } = "";
    public RenewalDocumentDiscoveryResult DocumentDiscovery { get; set; } = new();
    public BatchWorkflowAvailability WorkflowAvailability { get; set; } = new();
    public List<string> DocumentTypesFound { get; } = [];
    public List<string> DocumentCarriersFound { get; } = [];
    public List<string> Warnings { get; } = [];

    public string PlanDesignSource => DocumentDiscovery.PlanDesign.Source;
    public string PlanDesignStatus => DocumentDiscovery.PlanDesign.Status;
    public string DocumentTypesText => DocumentTypesFound.Count == 0 ? "None" : string.Join(", ", DocumentTypesFound);
    public string DocumentCarriersText => DocumentCarriersFound.Count == 0 ? "Not confirmed" : string.Join(", ", DocumentCarriersFound);
    public int BookletFileCount => DocumentDiscovery.Documents.Count(document => document.IsBenefitBooklet);
    public string BookletFilesText => string.Join(", ", DocumentDiscovery.Documents
        .Where(document => document.IsBenefitBooklet)
        .Select(document => document.FileName)
        .OrderBy(fileName => fileName, StringComparer.OrdinalIgnoreCase));
    public string WorkflowsText => WorkflowAvailability.AvailableWorkflows.Count == 0 ? "None" : string.Join(", ", WorkflowAvailability.AvailableWorkflows);
    public string WarningsText => Warnings.Count == 0 ? "" : string.Join(" | ", Warnings);
}

public sealed class BookletClientIdentity
{
    public string SourcePath { get; init; } = "";
    public string SourceFile { get; init; } = "";
    public string ClientName { get; init; } = "";
    public string DocumentClientName { get; init; } = "";
    public string FileNameClientName { get; init; } = "";
    public string Carrier { get; init; } = "";
    public string GroupNumber { get; init; } = "";
    public string PolicyNumber { get; init; } = "";
    public decimal ConfidenceScore { get; init; }
    public string IdentitySource { get; init; } = "";
    public string IdentityAgreement { get; init; } = "";
    public string ErrorMessage { get; init; } = "";
}

public sealed class BatchWorkflowAvailability
{
    public bool RenewalIntelligence { get; set; }
    public bool ClaimsExperience { get; set; }
    public bool Billing { get; set; }
    public bool FinancialIntelligence { get; set; }
    public bool PlanDesign { get; set; }
    public bool BookletExtraction { get; set; }
    public bool ProposalComparison { get; set; }

    public List<string> AvailableWorkflows
    {
        get
        {
            var workflows = new List<string>();
            if (RenewalIntelligence)
            {
                workflows.Add(BatchWorkflowNames.RenewalIntelligence);
            }

            if (ClaimsExperience)
            {
                workflows.Add(BatchWorkflowNames.ClaimsExperience);
            }

            if (Billing)
            {
                workflows.Add(BatchWorkflowNames.Billing);
            }

            if (FinancialIntelligence)
            {
                workflows.Add(BatchWorkflowNames.FinancialIntelligence);
            }

            if (PlanDesign)
            {
                workflows.Add(BatchWorkflowNames.PlanDesign);
            }

            if (BookletExtraction)
            {
                workflows.Add(BatchWorkflowNames.BookletExtraction);
            }

            if (ProposalComparison)
            {
                workflows.Add(BatchWorkflowNames.ProposalComparison);
            }

            return workflows;
        }
    }
}

public sealed class BatchProcessingResult
{
    public string ParentFolder { get; set; } = "";
    public string OutputFolder { get; set; } = "";
    public string LogPath { get; set; } = "";
    public List<BatchClientProcessingResult> Clients { get; } = [];
    public string Status => Clients.Count == 0
        ? BatchProcessingStatuses.NeedsReview
        : Clients.Any(client => client.Status == BatchProcessingStatuses.Failed)
        ? BatchProcessingStatuses.CompletedWithWarnings
        : Clients.Any(client => client.Status == BatchProcessingStatuses.CompletedWithWarnings)
            ? BatchProcessingStatuses.CompletedWithWarnings
            : BatchProcessingStatuses.Completed;
}

public sealed class BatchClientProcessingResult
{
    public string ClientName { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string Status { get; set; } = BatchProcessingStatuses.Skipped;
    public string DiscoveryStatus { get; set; } = BatchProcessingStatuses.Discovered;
    public string OutputFolder { get; set; } = "";
    public TimeSpan ProcessingTime { get; set; }
    public List<string> Warnings { get; } = [];
    public List<string> Errors { get; } = [];
    public List<BatchWorkflowResult> Workflows { get; } = [];

    public string DocumentTypesFound { get; set; } = "";
    public string DocumentCarriersFound { get; set; } = "";
    public string PlanDesignSource { get; set; } = PlanDesignSources.None;
    public string PlanDesignStatus { get; set; } = PlanDesignStatuses.NotAvailable;
    public bool RenewalProcessed { get; set; }
    public bool RenewalWorkbookGenerated { get; set; }
    public bool RenewalSavedToDatabase { get; set; }
    public bool PowerPointGenerated { get; set; }
    public bool BillingProcessed { get; set; }
    public bool BillingSavedToDatabase { get; set; }
    public bool FinancialIntelligenceProcessed { get; set; }
    public bool FinancialIntelligenceSavedToDatabase { get; set; }
    public bool PlanDesignExtracted { get; set; }
    public int PlanDesignRowsSaved { get; set; }
    public string PlanDesignTargetTable { get; set; } = "";
    public bool ScheduleOfBenefitsUpdated { get; set; }
    public string DatabaseSaveStatus { get; set; } = "";
    public string WorkflowsExecuted => Workflows.Count == 0 ? "" : string.Join(", ", Workflows.Select(workflow => workflow.WorkflowName));
    public string WarningsText => Warnings.Count == 0 ? "" : string.Join(" | ", Warnings);
    public string ErrorsText => Errors.Count == 0 ? "" : string.Join(" | ", Errors);
}

public sealed class BatchWorkflowResult
{
    public string WorkflowName { get; set; } = "";
    public string Status { get; set; } = BatchProcessingStatuses.Skipped;
    public string SourceFile { get; set; } = "";
    public string Message { get; set; } = "";
    public string Exception { get; set; } = "";
    public string StackTrace { get; set; } = "";
    public TimeSpan ProcessingTime { get; set; }
    public List<string> OutputPaths { get; } = [];
    public List<string> Warnings { get; } = [];
    public bool WorkbookGenerated { get; set; }
    public bool PowerPointGenerated { get; set; }
    public bool SavedToDatabase { get; set; }
    public int DatabaseRowsSaved { get; set; }
    public int PlanDesignRowsExtracted { get; set; }
    public int PlanDesignRowsSaved { get; set; }
    public string PlanDesignTargetTable { get; set; } = "";
    public bool ScheduleOfBenefitsUpdated { get; set; }
    public string DatabaseSaveStatus { get; set; } = "";
}

public sealed class BatchProcessingProgress
{
    public string CurrentClient { get; set; } = "";
    public string CurrentWorkflow { get; set; } = "";
    public int Completed { get; set; }
    public int Total { get; set; }
    public int WarningCount { get; set; }
    public int FailedCount { get; set; }
    public TimeSpan EstimatedTimeRemaining { get; set; }
    public string Message { get; set; } = "";
}
