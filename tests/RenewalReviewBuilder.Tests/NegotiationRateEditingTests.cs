using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Markup;
using System.Xml.Linq;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class NegotiationRateEditingTests
{
    [Theory]
    [InlineData(1)]
    [InlineData(12)]
    public void EditingBaselines_RecalculatesMatchingPremium_InExistingUnits(int scale)
    {
        var row = HealthRow(scale);
        row.EditableCurrentRate = 50m;
        row.EditableProposedRate = 60m;

        Assert.Equal(100m * scale, row.CurrentPremium);
        Assert.Equal(120m * scale, row.ProposedPremium);
        Assert.Equal(110m, row.NegotiatedRate);
        Assert.Equal(220m * scale, row.NegotiatedPremium);
        Assert.Contains("Current rate corrected from 100 to 50", row.Notes);
        Assert.Contains("Proposed rate corrected from 120 to 60", row.ToSaveRecord().Notes);
    }

    [Fact]
    public void EditingCriticalIllness_UsesPrintedTenThousandDivisor()
    {
        var row = new NegotiationRateRow
        {
            BenefitLine = "Critical Illness", RateBasis = "Per $10,000", Volume = 1_620_000m,
            CurrentRate = 6.74m, CurrentPremium = 1_091.88m,
            ProposedRate = 7.84m, ProposedPremium = 1_270.08m
        };
        row.EditableCurrentRate = 6m;
        row.EditableProposedRate = 7m;
        Assert.Equal(972m, row.CurrentPremium);
        Assert.Equal(1_134m, row.ProposedPremium);
    }

    [Fact]
    public void MissingExposure_PreservesSourcePremium_AndRecordsReason()
    {
        var row = new NegotiationRateRow
        {
            BenefitLine = "Basic Life", RateBasis = "volume",
            CurrentRate = 1m, CurrentPremium = 25m
        };
        row.EditableCurrentRate = 2m;
        Assert.Equal(25m, row.CurrentPremium);
        Assert.Contains("Premium retained", row.Notes);
    }

    [Fact]
    public void GrossCorrections_ReachExport_WithoutChangingNetOrNegotiatedRates()
    {
        var row = HealthRow();
        row.RateValueBasis = RateValueBasisNames.Gross;
        row.ClientPresentationRateBasis = RateValueBasisNames.Gross;
        row.CurrentNetRate = 90m;
        row.CurrentGrossRate = 100m;
        row.ProposedNetRate = 108m;
        row.ProposedGrossRate = 120m;
        row.VersionNetRate = 99m;
        row.VersionGrossRate = 110m;
        row.EditableCurrentRate = 105m;
        row.EditableProposedRate = 125m;
        var request = SaveRequest(RenewalVersionNames.Negotiated, row);
        var analysis = new AnalysisResult();
        var version = new NegotiationWorkbenchService().ApplySavedVersion(analysis, request, request.VersionName);
        var line = Assert.Single(version.RateLines);
        Assert.Equal(105m, ClientPresentationRateSelector.CurrentRate(line));
        Assert.Equal(125m, ClientPresentationRateSelector.ProposedRate(line));
        Assert.Equal(110m, ClientPresentationRateSelector.VersionRate(line));
        Assert.Equal(90m, line.CurrentNetRate);
        Assert.Equal(108m, line.ProposedNetRate);
    }

    [Fact]
    public async Task SaveReloadAndExport_PreserveCorrectedBaselines_OverOriginalCarrierValues()
    {
        var folder = Path.Combine(Path.GetTempPath(), "RenewalReviewBuilderTests", $"RateEditing-{Guid.NewGuid():N}");
        Directory.CreateDirectory(folder);
        var db = new DatabaseService(Path.Combine(folder, "test.db"));
        try
        {
            var original = SaveRequest(RenewalVersionNames.CarrierProposed, HealthRow());
            var originalSave = await db.SaveNegotiationWorkbenchAsync(original);
            Assert.True(originalSave.Succeeded, originalSave.Message);
            var edited = HealthRow();
            edited.EditableCurrentRate = 50m;
            edited.EditableProposedRate = 60m;
            var request = SaveRequest(RenewalVersionNames.Negotiated, edited);
            var saved = await db.SaveNegotiationWorkbenchAsync(request);
            Assert.True(saved.Succeeded, saved.Message);

            var history = new ClientHistory();
            history.RenewalVersions.AddRange(await db.GetRenewalVersionsAsync(saved.ClientId!.Value));
            var reopened = Assert.Single(MainWindow.BuildNegotiationRowsFromHistory(new RenewalQueueItem
            {
                ClientId = saved.ClientId.Value, Client = request.ClientName,
                RenewalDate = request.EffectiveDate, CurrentPremium = 200m, RenewalPremium = 240m
            }, history));
            Assert.Equal(50m, reopened.CurrentRate);
            Assert.Equal(60m, reopened.ProposedRate);
            Assert.Equal(100m, reopened.CurrentPremium);
            Assert.Equal(120m, reopened.ProposedPremium);
            Assert.Equal(110m, reopened.NegotiatedRate);
            Assert.Equal(220m, reopened.NegotiatedPremium);
            var originalRate = Assert.Single(history.RenewalVersions.Single(v => v.VersionName == RenewalVersionNames.CarrierProposed).Rates);
            Assert.Equal(100m, originalRate.CurrentRate);
            Assert.Equal(120m, originalRate.ProposedRate);

            var analysis = new AnalysisResult();
            var workbench = new NegotiationWorkbenchService();
            var reopenedRequest = SaveRequest(RenewalVersionNames.Negotiated, reopened);
            workbench.ApplySavedVersion(analysis, reopenedRequest, reopenedRequest.VersionName);
            using var workbook = new XLWorkbook(new ExcelReportWriter().Write(analysis, folder));
            var rates = workbook.Worksheet("Rate Reconciliation");
            var health = Assert.Single(rates.RowsUsed(), r => r.Cell(1).GetString() == "Extended Health Care");
            Assert.Equal(50m, health.Cell(5).GetValue<decimal>());
            var ppt = PowerPointRateTableBuilder.BuildPlaceholderTables(analysis)["<<RateTable_LineRates>>"];
            var pptHealth = Assert.Single(ppt.Rows, r => r.Length > 2 && r[1].StartsWith("Extended Health Care", StringComparison.Ordinal));
            Assert.Equal(50m, decimal.Parse(pptHealth[5], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(60m, decimal.Parse(pptHealth[7], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(110m, decimal.Parse(pptHealth[9], System.Globalization.CultureInfo.InvariantCulture));
        }
        finally
        {
            SqliteConnection.ClearAllPools();
        }
    }

    [Fact]
    public void ActualWpfRateFields_AcceptCorrections_AndRejectInvalidInput()
    {
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                var folder = new DirectoryInfo(AppContext.BaseDirectory);
                while (folder is not null && !File.Exists(Path.Combine(folder.FullName, "RenewalReviewBuilder.sln")))
                    folder = folder.Parent;
                var document = XDocument.Load(Path.Combine(folder!.FullName, "src", "RenewalReviewBuilder.App", "MainWindow.xaml"));
                XNamespace ns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";
                XNamespace x = "http://schemas.microsoft.com/winfx/2006/xaml";
                var row = HealthRow();
                foreach (var name in new[] { "NegotiationCurrentRateColumn", "NegotiationProposedRateColumn" })
                {
                    var definition = new XElement(document.Descendants(ns + "DataGridTemplateColumn").Single(c => (string?)c.Attribute(x + "Name") == name));
                    definition.SetAttributeValue(XNamespace.Xmlns + "x", x.NamespaceName);
                    definition.Descendants(ns + "TextBox").Attributes("LostFocus").Remove();
                    var column = (DataGridTemplateColumn)XamlReader.Parse(definition.ToString());
                    var input = (TextBox)column.CellTemplate.LoadContent();
                    input.DataContext = row;
                    System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(
                        () => { }, System.Windows.Threading.DispatcherPriority.ContextIdle);
                    var binding = input.GetBindingExpression(TextBox.TextProperty)!;
                    binding.UpdateTarget();
                    Assert.Same(row, binding.DataItem);
                    Assert.Equal(BindingStatus.Active, binding.Status);
                    Assert.Equal(BindingMode.TwoWay, binding.ParentBinding.Mode);
                    input.SetCurrentValue(TextBox.TextProperty, "55.25");
                    binding.UpdateSource();
                    Assert.False(Validation.GetHasError(input));
                    Assert.Equal(55.25m, name.Contains("Current", StringComparison.Ordinal) ? row.CurrentRate : row.ProposedRate);
                    Assert.Equal(110.50m, name.Contains("Current", StringComparison.Ordinal) ? row.CurrentPremium : row.ProposedPremium);
                    foreach (var invalid in new[] { "-1", "wrong", "" })
                    {
                        input.SetCurrentValue(TextBox.TextProperty, invalid);
                        binding.UpdateSource();
                        Assert.True(Validation.GetHasError(input));
                        Assert.Equal(55.25m, name.Contains("Current", StringComparison.Ordinal) ? row.CurrentRate : row.ProposedRate);
                    }
                    input.SetCurrentValue(TextBox.TextProperty, "0");
                    binding.UpdateSource();
                    Assert.False(Validation.GetHasError(input));
                }
            }
            catch (Exception ex) { failure = ex; }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.True(thread.Join(TimeSpan.FromSeconds(30)), "Rate editing UI check timed out.");
        Assert.Null(failure);
    }

    private static NegotiationRateRow HealthRow(int scale = 1) => new()
    {
        BenefitLine = "Extended Health Care", AccountClass = "Class 1", ClassName = "Single", RateBasis = "single",
        Lives = 2, CurrentRate = 100m, CurrentPremium = 200m * scale,
        ProposedRate = 120m, ProposedPremium = 240m * scale,
        NegotiatedRate = 110m, NegotiatedPremium = 220m * scale
    };

    private static NegotiationWorkbenchSaveRequest SaveRequest(string version, NegotiationRateRow row)
    {
        var request = new NegotiationWorkbenchSaveRequest
        {
            ClientName = "Rate correction test", Carrier = "Test Carrier",
            EffectiveDate = new DateOnly(2026, 10, 1), RenewalYear = 2026,
            VersionName = version, Status = "Negotiated"
        };
        request.Rates.Add(row.ToSaveRecord());
        return request;
    }
}
