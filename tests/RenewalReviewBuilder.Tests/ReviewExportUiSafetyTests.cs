using System;
using System.IO;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ReviewExportUiSafetyTests
{
    [Fact]
    public void MainWindow_Binds_Export_Mutated_Review_Collections_Through_Snapshots()
    {
        var root = FindRepositoryRoot();
        var sourcePath = Path.Combine(root, "src", "RenewalReviewBuilder.App", "MainWindow.xaml.cs");
        var source = File.ReadAllText(sourcePath);

        Assert.Contains(
            "FindingsGrid.ItemsSource = _analysis.Findings.ToList();",
            source,
            StringComparison.Ordinal);
        Assert.Contains(
            "SourceChecksGrid.ItemsSource = _analysis?.SourceChecks.ToList();",
            source,
            StringComparison.Ordinal);
        Assert.DoesNotContain(
            "SourceChecksGrid.ItemsSource = _analysis.SourceChecks;",
            source,
            StringComparison.Ordinal);

        var exportStart = source.IndexOf("private async void Export_Click(", StringComparison.Ordinal);
        var exportCall = source.IndexOf("await _builder.ExportAsync(_analysis", exportStart, StringComparison.Ordinal);
        var refreshAfterExport = source.IndexOf("RefreshSourceChecksGrid();", exportCall, StringComparison.Ordinal);

        Assert.True(exportStart >= 0, "Export_Click was not found.");
        Assert.True(exportCall > exportStart, "ExportAsync call was not found in Export_Click.");
        Assert.True(
            refreshAfterExport > exportCall,
            "The Source Checks snapshot must be refreshed after background export completes.");
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
