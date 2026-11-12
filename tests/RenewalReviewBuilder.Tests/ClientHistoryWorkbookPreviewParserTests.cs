using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientHistoryWorkbookPreviewParserTests
{
    [Fact]
    public void Parses_GroupHistory_Sections_ContinuationRows_And_RowEvidence()
    {
        var path = CreateFixture();
        try
        {
            var parser = new ClientHistoryWorkbookPreviewParser();
            Assert.True(parser.LooksLikeClientHistoryWorkbook(path));

            var preview = parser.Parse(path);

            Assert.Equal(3, preview.CountsByEventType["Benefit Plan Amendments"]);
            Assert.Equal(2, preview.CountsByEventType["Market Survey"]);
            Assert.Equal(1, preview.CountsByEventType["Change In Carrier"]);
            Assert.Equal(1, preview.CountsByEventType["Meetings"]);
            var continuation = Assert.Single(preview.Records, row => row.Description.Contains("continued dental", StringComparison.OrdinalIgnoreCase));
            Assert.Equal(new DateOnly(2025, 9, 1), continuation.EffectiveDate);
            Assert.Equal("Group History", continuation.SourceSheet);
            Assert.Equal(5, continuation.SourceRow);
            Assert.Equal(path, continuation.SourceFile);
            Assert.Contains("continued dental", continuation.OriginalText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(preview.Records, row => row.PreviewStatus == ClientHistoryPreviewStatuses.NeedsReview && row.Description.Contains("unresolved", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(preview.Records, row => row.PreviewStatus == ClientHistoryPreviewStatuses.Parsed && row.ExtractionMethod == "other-client-attribute");
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Keeps_A_Malformed_Row_In_Review_Without_Stopping_The_Workbook()
    {
        var path = CreateFixture();
        try
        {
            var preview = new ClientHistoryWorkbookPreviewParser().Parse(path);
            Assert.Contains(preview.Records, row => row.EventType == "Market Survey");
            Assert.Contains(preview.Records, row => row.PreviewStatus == ClientHistoryPreviewStatuses.NeedsReview);
        }
        finally { File.Delete(path); }
    }

    private static string CreateFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"testHistory-{Guid.NewGuid():N}.xlsx");
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Group History");
        sheet.Cell("A1").Value = "Group History";
        sheet.Cell("A3").Value = "Benefit Plan Amendments";
        sheet.Cell("A4").Value = new DateTime(2025, 9, 1);
        sheet.Cell("B4").Value = "Health maximum increased";
        sheet.Cell("B5").Value = "continued dental plan amendment";
        sheet.Cell("A6").Value = "1st February 2026";
        sheet.Cell("B6").Value = "Dental deductible removed";
        sheet.Cell("A8").Value = "Market Survey";
        sheet.Cell("B9").Value = "unresolved market note without a date";
        sheet.Cell("A10").Value = "Sept 2nd 2025";
        sheet.Cell("B10").Value = "Market survey completed";
        sheet.Cell("A11").Value = "Change In Carrier";
        sheet.Cell("A12").Value = "March 3rd 2025";
        sheet.Cell("B12").Value = "Changed carrier from Alpha to Beta";
        sheet.Cell("A14").Value = "Meetings";
        sheet.Cell("A15").Value = "4/5/2025";
        sheet.Cell("B15").Value = "Annual renewal meeting";
        sheet.Cell("A17").Value = "Other";
        sheet.Cell("B18").Value = "Unionized workforce";
        sheet.Cell("B19").Value = "unresolved note without a date";
        workbook.SaveAs(path);
        return path;
    }
}
