using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.Normalization;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class NormalizedDataValidatorTests
{
    [Fact]
    public void Validate_RenewalPayload_FlagsMissingCoreFields()
    {
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.Renewal,
            Source = new NormalizedSourceMetadata
            {
                SourceType = SourceTypes.Pdf,
                RawSourceReference = "renewal.pdf"
            },
            Client = new NormalizedClient(),
            Carrier = new NormalizedCarrier()
        };

        var issues = new NormalizedDataValidator().Validate(payload);
        var codes = issues.Select(issue => issue.IssueCode).ToList();

        Assert.Contains(ValidationIssueCodes.MissingClientName, codes);
        Assert.Contains(ValidationIssueCodes.MissingCarrier, codes);
        Assert.Contains(ValidationIssueCodes.MissingRenewalEffectiveDate, codes);
        Assert.Contains(ValidationIssueCodes.MissingPremium, codes);
        Assert.Contains(ValidationIssueCodes.MissingClaims, codes);
        Assert.Contains(ValidationIssueCodes.MissingRates, codes);
    }

    [Fact]
    public void Validate_BillingPayload_FlagsMissingProvinceAndOptionClass()
    {
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.Billing,
            Source = new NormalizedSourceMetadata
            {
                SourceType = SourceTypes.Pdf,
                SourceCarrier = "Manulife"
            },
            Client = new NormalizedClient { ClientName = "Test Client" },
            Carrier = new NormalizedCarrier { CarrierName = "Manulife" }
        };
        payload.BillingStatements.Add(new NormalizedBillingStatement
        {
            BenefitLine = "Health",
            MonthlyPremium = 100m
        });
        payload.ApplySourceToRecords();

        var issues = new NormalizedDataValidator().Validate(payload);
        var codes = issues.Select(issue => issue.IssueCode).ToList();

        Assert.Contains(ValidationIssueCodes.ProvinceMissing, codes);
        Assert.Contains(ValidationIssueCodes.OptionClassMissing, codes);
    }

    [Fact]
    public void Validate_PlanDesignPayload_FlagsRejectedAndTieredPercentages()
    {
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.BenefitBooklet,
            Source = new NormalizedSourceMetadata
            {
                SourceType = SourceTypes.Pdf,
                SourceFileName = "booklet.pdf",
                SourceCarrier = "Equitable"
            },
            Client = new NormalizedClient { ClientName = "Guardrail Test Client" },
            Carrier = new NormalizedCarrier { CarrierName = "Equitable" }
        };
        payload.PlanDesign.Add(new NormalizedPlanDesign
        {
            BenefitName = "Prescription Drugs",
            ReimbursementPercent = 8050m,
            ReimbursementText = "8050%"
        });
        payload.PlanDesign.Add(new NormalizedPlanDesign
        {
            BenefitName = "Specialty Drugs",
            ReimbursementText = "80% for the first tier; 100% thereafter"
        });
        payload.ApplySourceToRecords();

        var issues = new NormalizedDataValidator().Validate(payload);
        var codes = issues.Select(issue => issue.IssueCode).ToList();

        Assert.Contains(ValidationIssueCodes.PlanDesignPercentageOutOfRange, codes);
        Assert.Contains(ValidationIssueCodes.PlanDesignMultiplePercentagesPreserved, codes);
    }

    [Fact]
    public void Validate_PlanDesignPayload_FlagsMissingAndDuplicateProvisionIdentity()
    {
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.BenefitBooklet,
            Source = new NormalizedSourceMetadata
            {
                SourceType = SourceTypes.Pdf,
                SourceFileName = "booklet.pdf",
                SourceCarrier = "Equitable"
            },
            Client = new NormalizedClient { ClientName = "Provision Identity Test Client" },
            Carrier = new NormalizedCarrier { CarrierName = "Equitable" }
        };
        payload.PlanDesign.Add(new NormalizedPlanDesign
        {
            BenefitCategory = "Prescription Drugs",
            BenefitName = "Prescription Drugs",
            ClassName = "Class A",
            ProvisionCode = "DEDUCTIBLE",
            ProvisionLabel = "Deductible"
        });
        payload.PlanDesign.Add(new NormalizedPlanDesign
        {
            BenefitCategory = "Prescription Drugs",
            BenefitName = "Prescription Drugs",
            ClassName = "Class A",
            ProvisionCode = "DEDUCTIBLE",
            ProvisionLabel = "Deductible"
        });
        payload.PlanDesign.Add(new NormalizedPlanDesign
        {
            BenefitCategory = "Vision",
            BenefitName = "Vision",
            ClassName = "Class A"
        });
        payload.ApplySourceToRecords();

        var codes = new NormalizedDataValidator().Validate(payload)
            .Select(issue => issue.IssueCode)
            .ToList();

        Assert.Contains(ValidationIssueCodes.PlanDesignProvisionIdentityMissing, codes);
        Assert.Contains(ValidationIssueCodes.PlanDesignDuplicateProvisionIdentity, codes);
    }
}
