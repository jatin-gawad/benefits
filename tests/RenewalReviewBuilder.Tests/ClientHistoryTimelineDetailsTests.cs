using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientHistoryTimelineDetailsTests
{
    [Fact]
    public async Task Details_Action_Reopens_The_Shared_Drawer_And_Moves_Canonical_Selection()
    {
        var state = new HistorySelectionState();
        var events = new[]
        {
            Event("first", new DateOnly(2024, 4, 1), 41, 7, 101),
            Event("second", new DateOnly(2025, 5, 1), 42, 8, 202)
        };
        using var controller = new ClientHistoryWorkspaceController(
            state,
            new StaticProjection(events),
            (_, _) => Task.FromResult<IReadOnlyList<RenewalVersionRecord>>([]));
        await controller.LoadAsync(90);
        controller.Presentation.SelectMode(HistoryMode.Evolution);
        controller.SetFilters(
            "coverage",
            "Plan Design",
            new DateOnly(2024, 1, 1),
            new DateOnly(2025, 12, 31));

        Assert.True(controller.OpenHistoryEventDetails("first"));
        Assert.Equal("first", state.SelectedHistoryEventId);
        Assert.Equal("first", state.SelectedEvolutionStepId);
        Assert.Equal(41, state.SelectedRenewalId);
        Assert.Equal(7, state.SelectedCarrierId);
        Assert.Equal(2024, state.SelectedYear);
        Assert.Equal([101L], state.SelectedDocumentIds);
        Assert.True(state.IsDetailsDrawerOpen);

        controller.Presentation.CloseDrawer();
        Assert.False(state.IsDetailsDrawerOpen);
        Assert.Equal("first", state.SelectedHistoryEventId);

        Assert.True(controller.OpenHistoryEventDetails("first"));
        Assert.True(state.IsDetailsDrawerOpen);
        Assert.Equal("first", state.SelectedHistoryEventId);

        Assert.True(controller.OpenHistoryEventDetails("second"));
        Assert.Equal("second", state.SelectedHistoryEventId);
        Assert.Equal("second", state.SelectedEvolutionStepId);
        Assert.Equal(42, state.SelectedRenewalId);
        Assert.Equal(8, state.SelectedCarrierId);
        Assert.Equal(2025, state.SelectedYear);
        Assert.Equal([202L], state.SelectedDocumentIds);
        Assert.True(state.IsDetailsDrawerOpen);
        Assert.Equal(HistoryMode.Evolution, state.ActiveHistoryMode);
        Assert.Equal("coverage", controller.Presentation.SearchText);
        Assert.Equal("Plan Design", controller.Presentation.SelectedCategory);
        Assert.Equal(new DateOnly(2024, 1, 1), controller.Presentation.DateFrom);
        Assert.Equal(new DateOnly(2025, 12, 31), controller.Presentation.DateTo);
    }

    [Fact]
    public void Drawer_Content_Uses_Available_Canonical_Details_And_Original_Evidence_Text()
    {
        var historyEvent = new ClientHistoryEventReadModel
        {
            EventId = "detail",
            Summary = "Dental coverage was expanded.",
            BusinessRationale = "Improve employee retention.",
            Category = "Plan Design",
            RequestedBy = "Consultant",
            ApprovedBy = "Client",
            Notes = "Approved during the renewal meeting.",
            Evidence =
            [
                new HistoryEvidenceReference
                {
                    DocumentId = 88,
                    SheetName = "Group History",
                    CellRange = "B14",
                    ExtractedTextReference = "Original imported workbook wording",
                    ValidationStatus = HistoryValidationStatus.Valid
                }
            ]
        };

        var details = ClientHistoryDrawerContentFormatter.Details(historyEvent);
        var evidence = Assert.Single(ClientHistoryDrawerContentFormatter.Evidence(historyEvent));

        Assert.Equal(
        [
            "Summary: Dental coverage was expanded.",
            "Business rationale: Improve employee retention.",
            "Category: Plan Design",
            "Requested by: Consultant",
            "Approved by: Client",
            "Notes: Approved during the renewal meeting."
        ], details);
        Assert.Contains("Document 88", evidence, StringComparison.Ordinal);
        Assert.Contains("Group History B14", evidence, StringComparison.Ordinal);
        Assert.Contains("Source text: Original imported workbook wording", evidence, StringComparison.Ordinal);
        Assert.Contains("Valid", evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void Timeline_Details_Wiring_Uses_Row_Context_And_The_Shared_Details_Path()
    {
        var root = FindRepositoryRoot();
        var xamlPath = Path.Combine(root, "src", "RenewalReviewBuilder.App", "Views", "ClientHistoryWorkspaceView.xaml");
        var codePath = Path.Combine(root, "src", "RenewalReviewBuilder.App", "Views", "ClientHistoryWorkspaceView.xaml.cs");
        var document = XDocument.Load(xamlPath);
        XNamespace presentation = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";

        var button = Assert.Single(
            document.Descendants(presentation + "Button"),
            element => element.Attribute("Content")?.Value == "Details  >");
        var source = File.ReadAllText(codePath);

        Assert.Equal("TimelineDetails_Click", button.Attribute("Click")?.Value);
        Assert.Null(button.Attribute("Tag"));
        Assert.Contains("DataContext is not ClientHistoryTimelineRow row", source, StringComparison.Ordinal);
        Assert.Contains("OpenTimelineDetails(row, focusDrawer: true)", source, StringComparison.Ordinal);
        Assert.Contains("Workspace.OpenHistoryEventDetails(row.EventId)", source, StringComparison.Ordinal);
        Assert.Contains("SetDrawer();", source, StringComparison.Ordinal);
        Assert.Contains("RenderDrawer();", source, StringComparison.Ordinal);
    }

    [Fact]
    public void Kpi_Typography_Is_Doubled_Without_The_Legacy_Fixed_Height()
    {
        var root = FindRepositoryRoot();
        var xamlPath = Path.Combine(root, "src", "RenewalReviewBuilder.App", "Views", "ClientHistoryWorkspaceView.xaml");
        var document = XDocument.Load(xamlPath);
        XNamespace presentation = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";
        XNamespace xaml = "http://schemas.microsoft.com/winfx/2006/xaml";

        Assert.True(StyleFontSize(document, presentation, xaml, "HistoryKpiLabelText") >= 16);
        Assert.True(StyleFontSize(document, presentation, xaml, "HistoryKpiValueText") >= 30);
        Assert.True(StyleFontSize(document, presentation, xaml, "HistoryKpiSubtitleText") >= 14);

        var grid = Assert.Single(
            document.Descendants(presentation + "UniformGrid"),
            element => element.Attribute(xaml + "Name")?.Value == "HistoryKpiGrid");
        Assert.Equal("5", grid.Attribute("Columns")?.Value);
        Assert.Null(grid.Attribute("Height"));
        Assert.True(double.Parse(grid.Attribute("MinHeight")!.Value, CultureInfo.InvariantCulture) >= 120);

        foreach (var binding in new[]
                 {
                     "Workspace.Kpis.YearsAsClient",
                     "Workspace.Kpis.TotalValueCreated",
                     "Workspace.Kpis.AverageRenewalIncrease",
                     "Workspace.CurrentCarrierDisplay"
                 })
        {
            var value = Assert.Single(
                grid.Descendants(presentation + "TextBlock"),
                element => element.Attribute("Text")?.Value.Contains(binding, StringComparison.Ordinal) == true);
            Assert.Equal("{StaticResource HistoryKpiValueText}", value.Attribute("Style")?.Value);
            Assert.NotNull(value.Attribute("ToolTip"));
        }

        var carrier = Assert.Single(
            grid.Descendants(presentation + "TextBlock"),
            element => element.Attribute("Text")?.Value.Contains("Workspace.CurrentCarrierDisplay", StringComparison.Ordinal) == true);
        Assert.Equal("CharacterEllipsis", carrier.Attribute("TextTrimming")?.Value);
    }

    private static double StyleFontSize(
        XDocument document,
        XNamespace presentation,
        XNamespace xaml,
        string key)
    {
        var style = Assert.Single(
            document.Descendants(presentation + "Style"),
            element => element.Attribute(xaml + "Key")?.Value == key);
        var setter = Assert.Single(
            style.Elements(presentation + "Setter"),
            element => element.Attribute("Property")?.Value == "FontSize");
        return double.Parse(setter.Attribute("Value")!.Value, CultureInfo.InvariantCulture);
    }

    private static ClientHistoryEventReadModel Event(
        string eventId,
        DateOnly date,
        long renewalId,
        long carrierId,
        long documentId) => new()
        {
            EventId = eventId,
            StableSourceKey = eventId,
            ClientId = 90,
            SourceModule = "Manual",
            SourceEntityType = "History",
            SourceEntityId = eventId,
            EffectiveDate = date,
            EffectiveDateReference = HistoryDateReference.FromExact(date),
            EventType = "Benefit Plan Amendment",
            Category = "Plan Design",
            Title = $"{eventId} coverage change",
            Summary = "coverage update",
            RelatedRecords = [$"Renewal:{renewalId}", $"Carrier:{carrierId}"],
            Evidence = [new HistoryEvidenceReference { DocumentId = documentId }],
            ValidationStatus = HistoryValidationStatus.Valid
        };

    private static string FindRepositoryRoot()
    {
        var folder = new DirectoryInfo(AppContext.BaseDirectory);
        while (folder is not null && !File.Exists(Path.Combine(folder.FullName, "RenewalReviewBuilder.sln")))
        {
            folder = folder.Parent;
        }

        return folder?.FullName ?? throw new DirectoryNotFoundException("Repository root was not found.");
    }

    private sealed class StaticProjection(IReadOnlyList<ClientHistoryEventReadModel> events) : IClientHistoryProjectionBuilder
    {
        public Task<IReadOnlyList<ClientHistoryEventReadModel>> BuildAsync(
            long clientId,
            CancellationToken cancellationToken = default) => Task.FromResult(events);

        public Task<ClientHistoryAsOfQueryResult> BuildAsOfAsync(AsOfQueryContext context) =>
            Task.FromResult(new ClientHistoryAsOfQueryResult
            {
                Context = context,
                Events = events,
                ContributorSupport = []
            });
    }
}
