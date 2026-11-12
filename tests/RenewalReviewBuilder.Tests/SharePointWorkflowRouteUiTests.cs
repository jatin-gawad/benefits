using System;
using System.IO;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class SharePointWorkflowRouteUiTests
{
    [Fact]
    public void WorkflowSetup_Uses_Configured_SharePoint_Documents_Route()
    {
        var root = FindRepositoryRoot();
        var xaml = File.ReadAllText(Path.Combine(root, "src", "RenewalReviewBuilder.App", "MainWindow.xaml"));
        var code = File.ReadAllText(Path.Combine(root, "src", "RenewalReviewBuilder.App", "MainWindow.xaml.cs"));

        Assert.Contains("x:Name=\"OpenWestlandSharePointButton\"", xaml, StringComparison.Ordinal);
        Assert.Contains("Content=\"Open Westland SharePoint\"", xaml, StringComparison.Ordinal);
        Assert.Contains("Click=\"OpenWestlandSharePoint_Click\"", xaml, StringComparison.Ordinal);
        Assert.Contains("WESTLAND_SHAREPOINT_DOCUMENTS_URL", code, StringComparison.Ordinal);
        Assert.Contains("string.IsNullOrWhiteSpace(WestlandSharePointDocumentsUrl)", code, StringComparison.Ordinal);
        Assert.Contains("UseShellExecute = true", code, StringComparison.Ordinal);
    }

    private static string FindRepositoryRoot()
    {
        var current = new DirectoryInfo(AppContext.BaseDirectory);
        while (current is not null)
        {
            if (Directory.Exists(Path.Combine(current.FullName, "src", "RenewalReviewBuilder.App")))
            {
                return current.FullName;
            }

            current = current.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate the Build Renewal repository root.");
    }
}
