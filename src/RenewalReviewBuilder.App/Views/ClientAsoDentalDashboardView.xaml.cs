using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Views;

public partial class ClientAsoDentalDashboardView : UserControl
{
    public ClientAsoDentalDashboardView()
    {
        InitializeComponent();
        DataContext = ClientAsoDentalDashboardModel.Empty();
    }

    public void LoadDashboard(ClientWorkspaceViewModel workspace, IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        DataContext = ClientAsoDentalDashboardModel.FromMonthlyClaims(workspace, monthlyClaims);
    }

    public void LoadDashboard(
        ClientWorkspaceViewModel workspace,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        IReadOnlyCollection<AsoFinancialActivityStatement> asoStatements)
    {
        DataContext = ClientAsoDentalDashboardModel.FromSources(workspace, monthlyClaims, asoStatements);
    }

    public void LoadEmpty(ClientWorkspaceViewModel workspace)
    {
        DataContext = ClientAsoDentalDashboardModel.Empty(workspace);
    }
}

public sealed class ClientAsoDentalDashboardModel
{
    private const decimal AdministrationFeeRate = 0.15m;
    private const decimal GstRate = 0.05m;

    public string HeaderSubtitle { get; init; } = "Select a client to load its ASO dental cost schedule.";
    public string EstimateNote { get; init; } = "Estimated from saved paid dental claims using a 15.0% administration fee and 5.0% GST on that fee. Confirm the billed amount against the monthly activity report or invoice.";
    public string PaidClaimsText { get; init; } = "$0";
    public string AdministrationFeeText { get; init; } = "$0";
    public string GstText { get; init; } = "$0";
    public string TotalCostText { get; init; } = "$0";
    public string FeeBasisText { get; init; } = "15.0% consultant estimate";
    public string GstBasisText { get; init; } = "5.0% of estimated administration fee";
    public string TotalCostTitle { get; init; } = "Estimated ASO cost";
    public string TotalBasisText { get; init; } = "Claims + estimated fee + GST";
    public string PeriodLabel { get; init; } = "No saved service months";
    public string RowCountText { get; init; } = "0 months";
    public IReadOnlyList<ClientAsoDentalMonthlyRow> MonthlyRows { get; init; } = [];
    public bool HasData => MonthlyRows.Count > 0;
    public Visibility DataVisibility => HasData ? Visibility.Visible : Visibility.Collapsed;
    public Visibility EmptyVisibility => HasData ? Visibility.Collapsed : Visibility.Visible;

    public static ClientAsoDentalDashboardModel Empty(ClientWorkspaceViewModel? workspace = null)
    {
        return new ClientAsoDentalDashboardModel
        {
            HeaderSubtitle = workspace is null
                ? "Select a client to load its ASO dental cost schedule."
                : $"{workspace.ClientName} | ASO dental cost schedule"
        };
    }

    public static ClientAsoDentalDashboardModel FromMonthlyClaims(
        ClientWorkspaceViewModel workspace,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        var rows = ClientClaimsDashboardModel.SelectAuthoritativeMonthlyRows(monthlyClaims)
            .Where(row => row.Month.HasValue &&
                          row.PaidClaims.HasValue &&
                          row.PaidClaims.Value >= 0m &&
                          row.BenefitLine.Contains("dental", StringComparison.OrdinalIgnoreCase))
            .OrderBy(row => row.Month)
            .Select(row => ClientAsoDentalMonthlyRow.FromRecord(row, AdministrationFeeRate, GstRate))
            .ToList();

        return FromRows(workspace, rows, usesActualStatements: false);
    }

    public static ClientAsoDentalDashboardModel FromSources(
        ClientWorkspaceViewModel workspace,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        IReadOnlyCollection<AsoFinancialActivityStatement> asoStatements)
    {
        var actualRows = SelectActualInvoiceRows(asoStatements);
        return actualRows.Count > 0
            ? FromRows(workspace, actualRows, usesActualStatements: true)
            : FromMonthlyClaims(workspace, monthlyClaims);
    }

    private static List<ClientAsoDentalMonthlyRow> SelectActualInvoiceRows(
        IReadOnlyCollection<AsoFinancialActivityStatement> statements)
    {
        var completeStatements = statements
            .Where(statement => statement.IsCurrent &&
                                statement.ImportStatus.Equals(AsoFinancialImportStatuses.Complete, StringComparison.OrdinalIgnoreCase) &&
                                statement.Benefits.Contains("dental", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var invoiceStatements = completeStatements
            .Where(statement => statement.SourceSheet.Contains("ASO Invoice", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var authoritativeStatements = invoiceStatements.Count > 0 ? invoiceStatements : completeStatements;

        return authoritativeStatements
            .SelectMany(statement => statement.Periods
                .Where(period => period.ActivityMonth.HasValue && period.DentalClaimsPaid.HasValue)
                .Select(period => new { Statement = statement, Period = period }))
            .GroupBy(item => item.Period.ActivityMonth!.Value)
            .Select(group => group
                .OrderByDescending(item => item.Statement.StatementDate)
                .ThenByDescending(item => item.Statement.ImportedAt)
                .ThenByDescending(item => item.Statement.StatementId)
                .First())
            .OrderBy(item => item.Period.ActivityMonth)
            .Select(item => ClientAsoDentalMonthlyRow.FromStatement(item.Statement, item.Period))
            .ToList();
    }

    private static ClientAsoDentalDashboardModel FromRows(
        ClientWorkspaceViewModel workspace,
        List<ClientAsoDentalMonthlyRow> rows,
        bool usesActualStatements)
    {

        if (rows.Count == 0)
        {
            return Empty(workspace);
        }

        var paidClaims = rows.Sum(row => row.PaidClaims);
        var administrationFee = rows.Sum(row => row.AdministrationFee);
        var gst = rows.Sum(row => row.Gst);
        return new ClientAsoDentalDashboardModel
        {
            HeaderSubtitle = $"{workspace.ClientName} | Carrier: {DisplayOrDefault(workspace.PrimaryCarrier)} | Dental ASO renewal support",
            PaidClaimsText = FormatMoney(paidClaims),
            AdministrationFeeText = FormatMoney(administrationFee),
            GstText = FormatMoney(gst),
            TotalCostText = FormatMoney(rows.Sum(row => row.TotalCost)),
            EstimateNote = usesActualStatements
                ? "Actual carrier-observed Sun Life ASO invoice values. Billed fees include claims handling, general administration, profit, and consultant fees; total charges also include GST and statement interest or credits."
                : "Estimated from saved paid dental claims using a 15.0% administration fee and 5.0% GST on that fee. Confirm the billed amount against the monthly activity report or invoice.",
            FeeBasisText = usesActualStatements ? "Actual billed statement fees" : "15.0% consultant estimate",
            GstBasisText = usesActualStatements ? "Actual billed GST" : "5.0% of estimated administration fee",
            TotalCostTitle = usesActualStatements ? "Actual total charges" : "Estimated ASO cost",
            TotalBasisText = usesActualStatements ? "Claims + fees + tax + interest / credit" : "Claims + estimated fee + GST",
            PeriodLabel = FormatPeriod(rows.First().Month, rows.Last().Month),
            RowCountText = $"{rows.Count} {(rows.Count == 1 ? "month" : "months")}",
            MonthlyRows = rows
        };
    }

    private static string DisplayOrDefault(string? value) => string.IsNullOrWhiteSpace(value) ? "Not set" : value;
    private static string FormatMoney(decimal value) => value.ToString("C2", CultureInfo.GetCultureInfo("en-CA"));
    private static string FormatPeriod(DateOnly start, DateOnly end) => start == end
        ? start.ToString("MMMM yyyy", CultureInfo.InvariantCulture)
        : $"{start:MMM yyyy} - {end:MMM yyyy}";
}

public sealed record ClientAsoDentalMonthlyRow(
    DateOnly Month,
    decimal PaidClaims,
    decimal AdministrationFee,
    decimal Gst,
    decimal InterestCredit,
    decimal TotalCost,
    string SourceFileName)
{
    public string MonthText => Month.ToString("MMM yyyy", CultureInfo.InvariantCulture);
    public string PaidClaimsText => FormatMoney(PaidClaims);
    public string AdministrationFeeText => FormatMoney(AdministrationFee);
    public string GstText => FormatMoney(Gst);
    public string InterestCreditText => FormatMoney(InterestCredit);
    public string TotalCostText => FormatMoney(TotalCost);

    public static ClientAsoDentalMonthlyRow FromRecord(MonthlyClaimsRecord record, decimal administrationFeeRate, decimal gstRate)
    {
        var paidClaims = record.PaidClaims ?? 0m;
        var administrationFee = Math.Round(paidClaims * administrationFeeRate, 2, MidpointRounding.AwayFromZero);
        var gst = Math.Round(administrationFee * gstRate, 2, MidpointRounding.AwayFromZero);
        return new ClientAsoDentalMonthlyRow(
            record.Month!.Value,
            paidClaims,
            administrationFee,
            gst,
            0m,
            paidClaims + administrationFee + gst,
            "Saved monthly claims (estimated)");
    }

    public static ClientAsoDentalMonthlyRow FromStatement(
        AsoFinancialActivityStatement statement,
        AsoFinancialActivityPeriod period)
    {
        var paidClaims = period.DentalClaimsPaid ?? 0m;
        var billedFees = (period.TotalServiceFees ?? 0m) + (period.ConsultingFees ?? 0m);
        var gst = period.Gst ?? period.TotalTaxes ?? 0m;
        var interestCredit = period.InterestCreditCharge ?? 0m;
        var totalCost = period.TotalCharges ?? paidClaims + billedFees + gst + interestCredit;
        return new ClientAsoDentalMonthlyRow(
            period.ActivityMonth!.Value,
            paidClaims,
            billedFees,
            gst,
            interestCredit,
            totalCost,
            statement.SourceFileName);
    }

    private static string FormatMoney(decimal value) => value.ToString("C2", CultureInfo.GetCultureInfo("en-CA"));
}
