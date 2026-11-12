using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class AnalysisEngineTests
{
    [Fact]
    public void Recognized_Prior_Renewal_Does_Not_Report_Missing_Input()
    {
        var result = new AnalysisResult();
        result.PreviousRenewals.Add(new PreviousRenewalSnapshot
        {
            SourcePath = @"C:\Renewals\RenewalReport - 2025-08-06.pdf",
            SourceFileName = "RenewalReport - 2025-08-06.pdf",
            PolicyNumber = "123456",
            RenewalDate = new DateOnly(2025, 8, 1)
        });

        new AnalysisEngine().Analyze(result);

        Assert.DoesNotContain(result.Findings, finding =>
            finding.Title.Contains("Prior renewal", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Supplied_Unrecognized_Prior_Renewal_Is_Not_Reported_As_Not_Provided()
    {
        var result = new AnalysisResult
        {
            RenewalCase = new RenewalCase
            {
                AdditionalPdfPaths = [@"C:\Renewals\RenewalReport - 2025-08-06.pdf"]
            }
        };

        new AnalysisEngine().Analyze(result);

        Assert.Contains(result.Findings, finding =>
            finding.Title.Equals("Prior renewal supplied but not recognized", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(result.Findings, finding =>
            finding.Title.Equals("Prior renewal not provided", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task Golden_File_Produces_Client_Review_Findings()
    {
        if (!SampleData.HasSampleFiles)
        {
            return;
        }

        var builder = new ReviewPackBuilder();
        var result = await builder.AnalyzeAsync(new RenewalCase
        {
            PdfPath = SampleData.PdfPath,
            WorkbookPath = SampleData.WorkbookPath,
            OutputFolder = Path.Combine(Path.GetTempPath(), "RenewalReviewBuilderTests")
        });

        Assert.Contains(result.Findings, finding =>
            finding.Title.Contains("Total monthly premium", StringComparison.OrdinalIgnoreCase));

        Assert.Contains(result.Findings, finding =>
            finding.Title.Contains("paramedical", StringComparison.OrdinalIgnoreCase));

        Assert.Contains(result.Findings, finding =>
            finding.Title.Contains("diabetes", StringComparison.OrdinalIgnoreCase));
    }
}
