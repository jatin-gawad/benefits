namespace RenewalReviewBuilder.App.Models;

public sealed class RenewalCase
{
    public string ClientName { get; set; } = "";
    public string PolicyNumber { get; set; } = "";
    public DateOnly? RenewalDate { get; set; }
    public string PdfPath { get; set; } = "";
    public List<string> AdditionalPdfPaths { get; set; } = [];
    public string WorkbookPath { get; set; } = "";
    public List<string> AdditionalWorkbookPaths { get; set; } = [];
    public string CurrentYearFolder { get; set; } = "";
    public string PriorYearFolder { get; set; } = "";
    public string AsoBillingFolder { get; set; } = "";
    public HcsaCalculation HcsaCalculation { get; set; } = new();
    public AllstateCriticalIllnessEnrollment AllstateCriticalIllnessEnrollment { get; set; } = new();
    public string PowerPointTemplatePath { get; set; } = "";
    public string OutputFolder { get; set; } = "";
}
