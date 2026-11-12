namespace RenewalReviewBuilder.Tests;

// Optional, synthetic fixtures only. No private fixture locations are distributed.
internal static class SampleData
{
    private static readonly string Root = Environment.GetEnvironmentVariable("RENEWAL_TEST_FIXTURES") ?? "";
    public static string PdfPath => string.IsNullOrWhiteSpace(Root) ? "" : Path.Combine(Root, "sample-renewal.pdf");
    public static string WorkbookPath => string.IsNullOrWhiteSpace(Root) ? "" : Path.Combine(Root, "sample-claims.xlsm");
    public static bool HasSampleFiles => File.Exists(PdfPath) && File.Exists(WorkbookPath);
}
