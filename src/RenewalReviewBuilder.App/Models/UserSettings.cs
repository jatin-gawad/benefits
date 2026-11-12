namespace RenewalReviewBuilder.App.Models;

public sealed class UserSettings
{
    public string PdfPath { get; set; } = "";
    public List<string> AdditionalPdfPaths { get; set; } = [];
    public string WorkbookPath { get; set; } = "";
    public List<string> AdditionalWorkbookPaths { get; set; } = [];
    public List<string> BenefitBookletPaths { get; set; } = [];
    public string CurrentYearFolder { get; set; } = "";
    public string PriorYearFolder { get; set; } = "";
    public string AsoBillingFolder { get; set; } = "";
    public HcsaCalculation HcsaCalculation { get; set; } = new();
    public AllstateCriticalIllnessEnrollment AllstateCriticalIllnessEnrollment { get; set; } = new();
    public int ReportMonthCount { get; set; } = 12;
    public bool AutoDetectReportMonthCount { get; set; } = true;
    public string TemplatePath { get; set; } = "";
    public string OutputFolder { get; set; } = "";
}
