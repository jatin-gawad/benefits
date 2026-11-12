using System.Collections.Generic;
using System.Xml.Linq;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientHistoryAddHistoryWorkflowTests
{
    [Fact]
    public void Add_History_Button_Is_Wired_To_The_Workspace_Handler()
    {
        var root = FindRepositoryRoot();
        var xamlPath = Path.Combine(
            root,
            "src",
            "RenewalReviewBuilder.App",
            "Views",
            "ClientHistoryWorkspaceView.xaml");
        var document = XDocument.Load(xamlPath);
        XNamespace presentation = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";
        XNamespace xaml = "http://schemas.microsoft.com/winfx/2006/xaml";

        var button = Assert.Single(
            document.Descendants(presentation + "Button"),
            element => element.Attribute(xaml + "Name")?.Value == "AddHistoryButton");

        Assert.Equal("AddHistory_Click", button.Attribute("Click")?.Value);
    }

    [Fact]
    public void Historical_Renewal_Button_Is_Wired_To_The_Protected_Import_Handler()
    {
        var root = FindRepositoryRoot();
        var xamlPath = Path.Combine(root, "src", "RenewalReviewBuilder.App", "Views", "ClientHistoryWorkspaceView.xaml");
        var document = XDocument.Load(xamlPath);
        XNamespace presentation = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";
        XNamespace xaml = "http://schemas.microsoft.com/winfx/2006/xaml";

        var button = Assert.Single(
            document.Descendants(presentation + "Button"),
            element => element.Attribute(xaml + "Name")?.Value == "AddHistoricalRenewalButton");

        Assert.Equal("AddHistoricalRenewal_Click", button.Attribute("Click")?.Value);
    }

    [Fact]
    public async Task Import_Action_Passes_The_Persisted_Client_Id_And_Reloads_After_The_Dialog_Closes()
    {
        const long databaseClientId = 987654321;
        long? dialogClientId = null;
        var sequence = new List<string>();
        var action = new ClientHistoryImportAction();

        await action.RunAsync(
            databaseClientId,
            clientId =>
            {
                dialogClientId = clientId;
                sequence.Add("dialog");
            },
            _ =>
            {
                sequence.Add("reload");
                return Task.CompletedTask;
            });

        Assert.Equal(databaseClientId, dialogClientId);
        Assert.Equal(["dialog", "reload"], sequence);
    }

    [Fact]
    public void Workspace_Uses_The_Loaded_Database_Client_Id_Without_External_Id_Conversion()
    {
        var root = FindRepositoryRoot();
        var sourcePath = Path.Combine(
            root,
            "src",
            "RenewalReviewBuilder.App",
            "Views",
            "ClientHistoryWorkspaceView.xaml.cs");
        var source = File.ReadAllText(sourcePath);

        Assert.Contains("Workspace.DatabaseClientId.Value", source, StringComparison.Ordinal);
        Assert.Contains("new ClientHistoryImportPreviewWindow(databaseClientId)", source, StringComparison.Ordinal);
        Assert.Contains("Owner = owner", source, StringComparison.Ordinal);
        Assert.Contains("await Workspace.RefreshAsync(cancellationToken)", source, StringComparison.Ordinal);
        Assert.DoesNotContain("long.Parse", source, StringComparison.Ordinal);
        Assert.DoesNotContain("long.TryParse", source, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Import_Action_Rejects_A_Nonpersisted_Client_Before_Opening_The_Dialog()
    {
        var opened = false;
        var reloaded = false;
        var action = new ClientHistoryImportAction();

        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => action.RunAsync(
            0,
            _ => opened = true,
            _ =>
            {
                reloaded = true;
                return Task.CompletedTask;
            }));

        Assert.False(opened);
        Assert.False(reloaded);
    }

    private static string FindRepositoryRoot()
    {
        var folder = new DirectoryInfo(AppContext.BaseDirectory);
        while (folder is not null && !File.Exists(Path.Combine(folder.FullName, "RenewalReviewBuilder.sln")))
        {
            folder = folder.Parent;
        }

        return folder?.FullName ?? throw new DirectoryNotFoundException("Repository root was not found.");
    }
}
