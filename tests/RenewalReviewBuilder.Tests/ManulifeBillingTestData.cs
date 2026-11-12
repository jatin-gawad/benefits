using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.Tests;

internal static class ManulifeBillingTestData
{
    public const decimal ExpectedBenefitTotal = 467.36m;

    public static string CreateSyntheticPdf(string folder, string fileName = "synthetic-manulife-billing.pdf")
    {
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, fileName);
        var lines = new[]
        {
            "Billing period:",
            "MAR 1 2026 to MAR 31 2026",
            "Payment due: MAR 1 2026",
            "Portable Billing Test Client",
            "Questions?",
            "Internet: www.manulife.ca/groupbenefits",
            "Write: Manulife Financial Group Benefits",
            "Your Premium Due $ 467.36",
            "Your Taxes Due $ 0.00",
            "Your Current Month Amount Due $ 467.36",
            "Billing Id 9001001",
            "Group Benefits Billing Statement: Summary by Billing Division",
            "Class 110 British Columbia",
            "Extended Health S 01 2 1000 450.00 17.36 467.36 0.00 467.36",
            "Total Monthly Summary 450.00 17.36 467.36 0.00 467.36",
            "Total Tax Summary 0.00"
        };

        var content = new StringBuilder("BT\n/F1 9 Tf\n50 760 Td\n");
        foreach (var line in lines)
        {
            content.Append('(').Append(EscapePdfText(line)).Append(") Tj\n0 -14 Td\n");
        }
        content.Append("ET\n");

        var objects = new[]
        {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            $"<< /Length {Encoding.ASCII.GetByteCount(content.ToString())} >>\nstream\n{content}endstream"
        };

        var pdf = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int> { 0 };
        for (var index = 0; index < objects.Length; index++)
        {
            offsets.Add(Encoding.ASCII.GetByteCount(pdf.ToString()));
            pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }

        var xrefOffset = Encoding.ASCII.GetByteCount(pdf.ToString());
        pdf.Append("xref\n0 6\n0000000000 65535 f \n");
        foreach (var offset in offsets.Skip(1))
        {
            pdf.Append(offset.ToString("0000000000", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        }

        pdf.Append("trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n")
            .Append(xrefOffset.ToString(CultureInfo.InvariantCulture))
            .Append("\n%%EOF\n");
        File.WriteAllText(path, pdf.ToString(), Encoding.ASCII);
        return path;
    }

    public static BillingExtractionResult CreateValidExtraction(
        string sourceFile,
        string clientName = "Portable Billing Test Client",
        string clientNumber = "9001001")
    {
        var result = new BillingExtractionResult
        {
            ClientName = clientName,
            ClientNumber = clientNumber,
            BillingId = clientNumber,
            Carrier = "Manulife",
            BillingPeriod = "MAR 1 2026 to MAR 31 2026",
            PeriodStart = new DateOnly(2026, 3, 1),
            PeriodEnd = new DateOnly(2026, 3, 31),
            SourceFile = sourceFile,
            PageCount = 1,
            ReadablePdfExtraction = true,
            ClientIdentityConfirmed = true,
            CarrierIdentityConfirmed = true,
            ReconciliationExecuted = true,
            PdfGrandPremiumBeforeTax = ExpectedBenefitTotal,
            PdfGrandTax = 0m,
            PdfGrandTotalWithTax = ExpectedBenefitTotal
        };

        result.Details.Add(new BillingOptionPremiumDetail
        {
            BillingPeriod = result.BillingPeriod,
            Province = "British Columbia",
            Benefit = "Extended Health",
            ClassCode = "01",
            OptionLabel = "Option 1",
            DependentCode = "S",
            Lives = 2,
            Volume = 1000m,
            MonthlyPremium = 450m,
            PremiumAdjustments = 17.36m,
            TotalPremiumAdjustments = ExpectedBenefitTotal,
            SourcePage = 1,
            SourceFile = sourceFile
        });
        result.ProvinceSummaries.Add(new BillingProvinceSummary
        {
            BillingPeriod = result.BillingPeriod,
            Province = "British Columbia",
            TotalLives = 2,
            TotalMonthlyPremium = ExpectedBenefitTotal,
            EstimatedAnnualPremium = ExpectedBenefitTotal * 12m,
            TotalTax = 0m,
            TotalPremiumWithTax = ExpectedBenefitTotal,
            SourceFile = sourceFile
        });
        result.OptionSummaries.Add(new BillingOptionSummary
        {
            BillingPeriod = result.BillingPeriod,
            Province = "British Columbia",
            Benefit = "Extended Health",
            ClassCode = "01",
            OptionLabel = "Option 1",
            Lives = 2,
            MonthlyPremium = ExpectedBenefitTotal,
            EstimatedAnnualPremium = ExpectedBenefitTotal * 12m,
            SourceFile = sourceFile
        });
        result.ClassMappings.Add(new BillingClassMapping
        {
            Carrier = "Manulife",
            ClassCode = "01",
            OptionLabel = "Option 1",
            FirstSeenBillingPeriod = result.BillingPeriod,
            LastSeenBillingPeriod = result.BillingPeriod,
            SourceFile = sourceFile
        });

        result.ReconciliationResults.Add(CreateReconciliation("Row-level section total", result, includeLives: true));
        result.ReconciliationResults.Add(CreateReconciliation("Province total", result, includeLives: true));
        result.ReconciliationResults.Add(CreateReconciliation("Tax validation", result, includeLives: false));
        result.ReconciliationResults.Add(CreateReconciliation("Grand total", result, includeLives: true));
        return result;
    }

    private static BillingReconciliationResult CreateReconciliation(
        string level,
        BillingExtractionResult result,
        bool includeLives)
    {
        return new BillingReconciliationResult
        {
            BillingPeriod = result.BillingPeriod,
            SourceFile = result.SourceFile,
            ReconciliationLevel = level,
            Province = level == "Grand total" ? "" : "British Columbia",
            PdfLives = includeLives ? 2 : null,
            CalculatedLives = includeLives ? 2 : null,
            LivesDifference = includeLives ? 0 : null,
            LivesStatus = includeLives ? BillingReconciliationStatuses.Passed : BillingReconciliationStatuses.NotAvailable,
            PdfPremium = level == "Tax validation" ? null : ExpectedBenefitTotal,
            CalculatedPremium = level == "Tax validation" ? null : ExpectedBenefitTotal,
            PremiumDifference = level == "Tax validation" ? null : 0m,
            PdfTax = 0m,
            CalculatedTax = 0m,
            TaxDifference = 0m,
            PdfTotalWithTax = ExpectedBenefitTotal,
            CalculatedTotalWithTax = ExpectedBenefitTotal,
            TotalDifference = 0m,
            Status = BillingReconciliationStatuses.Passed,
            SourcePage = 1
        };
    }

    private static string EscapePdfText(string value)
    {
        return value.Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("(", "\\(", StringComparison.Ordinal)
            .Replace(")", "\\)", StringComparison.Ordinal);
    }
}
