using RenewalReviewBuilder.App.Models;
using UglyToad.PdfPig;

namespace RenewalReviewBuilder.App.Services;

public sealed class PdfTextExtractor
{
    public List<PdfPageText> Extract(string pdfPath)
    {
        var pages = new List<PdfPageText>();

        using var document = PdfDocument.Open(pdfPath);
        foreach (var page in document.GetPages())
        {
            var lines = ExtractLines(page)
                .Select(ParsingHelpers.Clean)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .ToList();
            var textParts = lines.Count > 0 ? lines : new List<string> { page.Text };

            pages.Add(new PdfPageText
            {
                PageNumber = page.Number,
                Text = ParsingHelpers.Clean(string.Join(" ", textParts)),
                Lines = lines.Count > 0 ? lines : SplitFallback(page.Text)
            });
        }

        return pages;
    }

    private static List<string> ExtractLines(UglyToad.PdfPig.Content.Page page)
    {
        var words = page.GetWords()
            .OrderByDescending(word => word.BoundingBox.Bottom)
            .ThenBy(word => word.BoundingBox.Left)
            .ToList();

        if (words.Count == 0)
        {
            return [];
        }

        var rows = new List<List<UglyToad.PdfPig.Content.Word>>();
        foreach (var word in words)
        {
            var row = rows.FirstOrDefault(existing =>
                Math.Abs(existing[0].BoundingBox.Bottom - word.BoundingBox.Bottom) < 3.0);

            if (row is null)
            {
                rows.Add(new List<UglyToad.PdfPig.Content.Word> { word });
            }
            else
            {
                row.Add(word);
            }
        }

        return rows
            .Select(row => string.Join(" ", row.OrderBy(word => word.BoundingBox.Left).Select(word => word.Text)))
            .ToList();
    }

    private static List<string> SplitFallback(string text)
    {
        return text
            .Split('\n')
            .Select(ParsingHelpers.Clean)
            .Where(line => line.Length > 0)
            .ToList();
    }
}
