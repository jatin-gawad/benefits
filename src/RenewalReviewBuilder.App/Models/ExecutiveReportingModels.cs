namespace RenewalReviewBuilder.App.Models;

public sealed class ExecutiveReportingSnapshot
{
    public DateTime LoadedAt { get; set; } = DateTime.Now;
    public List<ClientRecord> Clients { get; } = [];
    public List<RenewalQueueRecord> Renewals { get; } = [];
    public List<RenewalVersionRecord> RenewalVersions { get; } = [];
    public List<RenewalAssumptionRecord> RenewalAssumptions { get; } = [];
    public List<ClaimsExperienceRecord> Claims { get; } = [];
    public List<MonthlyClaimsRecord> MonthlyClaims { get; } = [];
    public List<PlanDesignRecord> PlanDesign { get; } = [];
    public List<ExecutiveProposalRecord> Proposals { get; } = [];
    public List<ExecutiveQuoteScenarioRecord> QuoteScenarios { get; } = [];
    public List<ExecutiveDataCategoryRecord> DataCategories { get; } = [];
}

public sealed class ExecutiveProposalRecord
{
    public long ProposalId { get; set; }
    public long ClientId { get; set; }
    public string ClientName { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string Status { get; set; } = "";
    public decimal? CurrentPremium { get; set; }
    public decimal? ProposedPremium { get; set; }
    public decimal? Savings { get; set; }
    public decimal? EmployerCost { get; set; }
    public decimal? EmployeeCost { get; set; }
    public DateTime DateLoaded { get; set; }
}

public sealed class ExecutiveQuoteScenarioRecord
{
    public long ScenarioId { get; set; }
    public long ClientId { get; set; }
    public string ScenarioName { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string Status { get; set; } = "";
    public decimal? CurrentPremium { get; set; }
    public decimal? EstimatedPremium { get; set; }
    public DateTime CreatedAt { get; set; }
}

public sealed class ExecutiveDataCategoryRecord
{
    public string Category { get; set; } = "";
    public int Records { get; set; }
    public string Status { get; set; } = "Detected";
}

public sealed class ExecutiveHealthWeights
{
    public decimal LossRatio { get; init; } = 0.25m;
    public decimal RenewalIncrease { get; init; } = 0.20m;
    public decimal ClaimsVolatility { get; init; } = 0.15m;
    public decimal PoolingLargeClaimsRisk { get; init; } = 0.15m;
    public decimal PlanCompetitiveness { get; init; } = 0.10m;
    public decimal ContributionCompetitiveness { get; init; } = 0.05m;
    public decimal ProposalMarketActivity { get; init; } = 0.05m;
    public decimal ConsultantRiskFlag { get; init; } = 0.05m;
}
