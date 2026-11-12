using System.Security.Cryptography;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.Services.FlexibleBenefits;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class FlexibleBenefitDatabaseTests
{
    [Fact]
    public async Task Save_CommitsNormalizedGraphLineageAndFlexibleOnlyWorkspace()
    {
        var test = await CreateTestContextAsync();
        try
        {
            var extraction = await BuildExtractionAsync(test.SourceFile, 500m);
            var validation = new FlexibleBenefitValidationService().Validate(extraction, test.ClientId);

            var save = await test.Database.SaveFlexibleBenefitExtractionAsync(
                Request(test, extraction.SourcePath),
                extraction,
                validation);

            Assert.True(save.Succeeded, save.Message);
            Assert.Equal(FlexibleBenefitImportOutcome.Completed, save.Outcome);
            await using var connection = await OpenAsync(test.DatabasePath);
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitPlan"));
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitClass"));
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitCategory"));
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitEligibilityRule"));
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitFundingConfiguration"));
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitFee"));
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitAgreement"));
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitContact"));
            Assert.Equal(extraction.SourceFields.Count, await CountAsync(connection, "FlexibleBenefitPlanSourceField"));
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitExtractionStaging"));
            Assert.Equal(0, await ForeignKeyViolationCountAsync(connection));

            var workspace = await test.Database.GetFlexibleBenefitWorkspaceAsync(test.ClientId);
            Assert.True(workspace.HasPlan);
            Assert.Equal("National HealthClaim", workspace.Administrator);
            Assert.Equal("$0 per year default allocation", workspace.HcsaSummary);
            Assert.Equal("$500 per year default allocation", workspace.LsaSummary);
            Assert.Single(workspace.Classes);
            Assert.Single(workspace.Categories);
            Assert.Equal(extraction.SourceFields.Count, workspace.SourceFields.Count);

            var planDesign = await new PlanDesignWorkspaceDataService(test.DatabasePath)
                .LoadAsync(test.ClientId, "Flexible Test Client", "National HealthClaim");
            Assert.False(planDesign.HasPlanDesign);
            Assert.True(planDesign.HasAnyPlanDesign);
            Assert.True(planDesign.FlexibleBenefits.HasPlan);
            Assert.Equal("$500 per year default allocation", planDesign.FlexibleBenefits.LsaSummary);
        }
        finally
        {
            test.Dispose();
        }
    }

    [Fact]
    public async Task SameHashSameClient_ReturnsDuplicateWithoutAdditionalRows()
    {
        var test = await CreateTestContextAsync();
        try
        {
            var extraction = await BuildExtractionAsync(test.SourceFile, 500m);
            var validation = new FlexibleBenefitValidationService().Validate(extraction, test.ClientId);
            var request = Request(test, extraction.SourcePath);

            var first = await test.Database.SaveFlexibleBenefitExtractionAsync(request, extraction, validation);
            var duplicate = await test.Database.SaveFlexibleBenefitExtractionAsync(request, extraction, validation);

            Assert.True(first.Succeeded, first.Message);
            Assert.Equal(FlexibleBenefitImportOutcome.Duplicate, duplicate.Outcome);
            await using var connection = await OpenAsync(test.DatabasePath);
            Assert.Equal(1, await CountAsync(connection, "Documents"));
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitPlan"));
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitClass"));
            Assert.Equal(extraction.SourceFields.Count, await CountAsync(connection, "FlexibleBenefitPlanSourceField"));
        }
        finally
        {
            test.Dispose();
        }
    }

    [Fact]
    public async Task DifferentFileSamePeriod_DefaultCancelThenNewVersionPreservesHistory()
    {
        var test = await CreateTestContextAsync();
        try
        {
            var secondFile = Path.Combine(test.Folder, "flex-plan-v2.pdf");
            await File.WriteAllTextAsync(secondFile, "version two flexible benefits");
            var firstExtraction = await BuildExtractionAsync(test.SourceFile, 500m);
            var secondExtraction = await BuildExtractionAsync(secondFile, 750m);
            var firstValidation = new FlexibleBenefitValidationService().Validate(firstExtraction, test.ClientId);
            var secondValidation = new FlexibleBenefitValidationService().Validate(secondExtraction, test.ClientId);

            var first = await test.Database.SaveFlexibleBenefitExtractionAsync(
                Request(test, test.SourceFile), firstExtraction, firstValidation);
            var cancelled = await test.Database.SaveFlexibleBenefitExtractionAsync(
                Request(test, secondFile), secondExtraction, secondValidation);
            var second = await test.Database.SaveFlexibleBenefitExtractionAsync(
                Request(test, secondFile, FlexibleBenefitImportMode.ImportAsNewVersion), secondExtraction, secondValidation);

            Assert.True(first.Succeeded, first.Message);
            Assert.Equal(FlexibleBenefitImportOutcome.ExistingVersion, cancelled.Outcome);
            Assert.True(second.Succeeded, second.Message);
            Assert.Equal(2, second.VersionNumber);
            await using var connection = await OpenAsync(test.DatabasePath);
            Assert.Equal(2, await CountAsync(connection, "FlexibleBenefitPlan"));
            Assert.Equal(2, await CountAsync(connection, "FlexibleBenefitClass"));
            Assert.Equal(2, await CountAsync(connection, "FlexibleBenefitDocumentVersion"));
            Assert.Equal(1, await ScalarIntAsync(connection, "SELECT COUNT(*) FROM FlexibleBenefitPlan WHERE version_number = 2 AND supersedes_plan_id IS NOT NULL;"));
            Assert.Equal(1, await ScalarIntAsync(connection, "SELECT COUNT(*) FROM FlexibleBenefitPlan WHERE is_current = 1 AND version_number = 2;"));
            Assert.Equal(1, await ScalarIntAsync(connection, "SELECT COUNT(*) FROM FlexibleBenefitPlan WHERE is_current = 0 AND version_number = 1;"));

            var workspace = await test.Database.GetFlexibleBenefitWorkspaceAsync(test.ClientId);
            Assert.Equal(2, workspace.VersionNumber);
            Assert.Equal("$750 per year default allocation", workspace.LsaSummary);
            Assert.Equal(2, workspace.Versions.Count);
            var differences = await test.Database.CompareFlexibleBenefitVersionsAsync(
                workspace.PlanId!.Value,
                workspace.Versions.Single(version => !version.IsCurrent).PlanId);
            Assert.Contains(differences, difference => difference.Area == "Class Limits" && difference.Field.Contains("Single", StringComparison.Ordinal));
        }
        finally
        {
            test.Dispose();
        }
    }

    [Fact]
    public async Task ReplaceExistingVersion_PreservesPriorRowsAndMarksPriorReplaced()
    {
        var test = await CreateTestContextAsync();
        try
        {
            var replacementFile = Path.Combine(test.Folder, "flex-plan-replacement.pdf");
            await File.WriteAllTextAsync(replacementFile, "replacement flexible benefits");
            var original = await BuildExtractionAsync(test.SourceFile, 500m);
            var replacement = await BuildExtractionAsync(replacementFile, 900m);
            var validator = new FlexibleBenefitValidationService();

            await test.Database.SaveFlexibleBenefitExtractionAsync(
                Request(test, test.SourceFile), original, validator.Validate(original, test.ClientId));
            var save = await test.Database.SaveFlexibleBenefitExtractionAsync(
                Request(test, replacementFile, FlexibleBenefitImportMode.ReplaceExistingVersion),
                replacement,
                validator.Validate(replacement, test.ClientId));

            Assert.True(save.Succeeded, save.Message);
            await using var connection = await OpenAsync(test.DatabasePath);
            Assert.Equal(2, await CountAsync(connection, "FlexibleBenefitPlan"));
            Assert.Equal(1, await ScalarIntAsync(connection, "SELECT COUNT(*) FROM FlexibleBenefitPlan WHERE version_number = 1 AND review_status = 'Replaced' AND is_current = 0;"));
            Assert.Equal(1, await ScalarIntAsync(connection, "SELECT COUNT(*) FROM FlexibleBenefitDocumentVersion WHERE version_number = 1 AND review_status = 'Replaced' AND is_current = 0;"));
        }
        finally
        {
            test.Dispose();
        }
    }

    [Fact]
    public async Task ChildInsertFailure_RollsBackDocumentPlanAndAllDetailRows()
    {
        var test = await CreateTestContextAsync();
        try
        {
            var extraction = await BuildExtractionAsync(test.SourceFile, 500m);
            extraction.Categories.Add(new FlexibleBenefitCategory
            {
                CategoryCode = extraction.Categories[0].CategoryCode,
                CategoryGroup = "General",
                CategoryName = "Duplicate category",
                Enabled = false,
                AccountType = FlexibleBenefitAccountType.Lsa,
                SourcePageNumber = 2,
                ExtractionConfidence = 0.9m
            });
            var intentionallyPermissiveValidation = new FlexibleBenefitValidationResult
            {
                Status = FlexibleBenefitExtractionStatuses.Completed
            };

            var save = await test.Database.SaveFlexibleBenefitExtractionAsync(
                Request(test, test.SourceFile), extraction, intentionallyPermissiveValidation);

            Assert.False(save.Succeeded);
            Assert.Equal(FlexibleBenefitImportOutcome.Failed, save.Outcome);
            await using var connection = await OpenAsync(test.DatabasePath);
            Assert.Equal(0, await CountAsync(connection, "Documents"));
            Assert.Equal(0, await CountAsync(connection, "FlexibleBenefitPlan"));
            Assert.Equal(0, await CountAsync(connection, "FlexibleBenefitClass"));
            Assert.Equal(0, await CountAsync(connection, "FlexibleBenefitCategory"));
            Assert.Equal(0, await CountAsync(connection, "FlexibleBenefitPlanSourceField"));
        }
        finally
        {
            test.Dispose();
        }
    }

    [Fact]
    public async Task CorrectionAndReview_UpdateNormalizedValueAndPreserveAudit()
    {
        var test = await CreateTestContextAsync();
        try
        {
            var extraction = await BuildExtractionAsync(test.SourceFile, 500m);
            var validation = new FlexibleBenefitValidationService().Validate(extraction, test.ClientId);
            var save = await test.Database.SaveFlexibleBenefitExtractionAsync(
                Request(test, test.SourceFile), extraction, validation);
            Assert.True(save.Succeeded, save.Message);

            long sourceFieldId;
            await using (var lookupConnection = await OpenAsync(test.DatabasePath))
            {
                sourceFieldId = await ScalarLongAsync(lookupConnection, @"
SELECT flexible_benefit_plan_source_field_id
FROM FlexibleBenefitPlanSourceField
WHERE entity_type = 'FlexibleBenefitPlan' AND field_name = 'RunoffDays'
LIMIT 1;");
            }

            var correction = await test.Database.CorrectFlexibleBenefitSourceFieldAsync(
                sourceFieldId,
                "90",
                "Test Analyst",
                "Confirmed in signed amendment");
            Assert.True(correction.Succeeded, correction.Message);

            await using var connection = await OpenAsync(test.DatabasePath);
            Assert.Equal(90, await ScalarIntAsync(connection, "SELECT runoff_days FROM FlexibleBenefitPlan LIMIT 1;"));
            Assert.Equal("90", await ScalarStringAsync(connection, "SELECT normalized_value FROM FlexibleBenefitPlanSourceField WHERE flexible_benefit_plan_source_field_id = " + sourceFieldId + ";"));
            Assert.Equal(1, await CountAsync(connection, "FlexibleBenefitCorrectionAudit"));
            Assert.Equal("60", await ScalarStringAsync(connection, "SELECT previous_normalized_value FROM FlexibleBenefitCorrectionAudit LIMIT 1;"));

            var reviewed = await test.Database.MarkFlexibleBenefitPlanReviewedAsync(save.FlexibleBenefitPlanId!.Value, "Test Analyst");
            Assert.True(reviewed.Succeeded, reviewed.Message);
            Assert.Equal("Reviewed", await ScalarStringAsync(connection, "SELECT review_status FROM FlexibleBenefitPlan LIMIT 1;"));
            Assert.Equal(0, await ScalarIntAsync(connection, "SELECT COUNT(*) FROM FlexibleBenefitPlanSourceField WHERE review_status <> 'Reviewed';"));
        }
        finally
        {
            test.Dispose();
        }
    }

    [Fact]
    public async Task SameHashDifferentClient_IsSeparateAndNeedsReview()
    {
        var test = await CreateTestContextAsync();
        try
        {
            var secondClientId = await InsertClientAsync(test.DatabasePath, "Second Flexible Client", "FLEX-002");
            var firstExtraction = await BuildExtractionAsync(test.SourceFile, 500m);
            var secondExtraction = await BuildExtractionAsync(test.SourceFile, 500m);
            var validator = new FlexibleBenefitValidationService();
            var first = await test.Database.SaveFlexibleBenefitExtractionAsync(
                Request(test, test.SourceFile), firstExtraction, validator.Validate(firstExtraction, test.ClientId));
            var secondRequest = new FlexibleBenefitImportRequest
            {
                ClientId = secondClientId,
                ClientName = "Second Flexible Client",
                FilePath = test.SourceFile,
                ImportedBy = "Test Analyst"
            };
            var second = await test.Database.SaveFlexibleBenefitExtractionAsync(
                secondRequest,
                secondExtraction,
                validator.Validate(secondExtraction, secondClientId));

            Assert.True(first.Succeeded, first.Message);
            Assert.True(second.Succeeded, second.Message);
            Assert.Equal(FlexibleBenefitImportOutcome.NeedsReview, second.Outcome);
            Assert.Contains(second.ValidationIssues, issue => issue.Code == "same_hash_different_client");
            await using var connection = await OpenAsync(test.DatabasePath);
            Assert.Equal(2, await CountAsync(connection, "Documents"));
            Assert.Equal(2, await CountAsync(connection, "FlexibleBenefitPlan"));
        }
        finally
        {
            test.Dispose();
        }
    }

    private static FlexibleBenefitImportRequest Request(
        FlexibleBenefitTestContext test,
        string filePath,
        FlexibleBenefitImportMode mode = FlexibleBenefitImportMode.Cancel)
    {
        return new FlexibleBenefitImportRequest
        {
            ClientId = test.ClientId,
            ClientName = "Flexible Test Client",
            FilePath = filePath,
            ImportMode = mode,
            ImportedBy = "Test Analyst"
        };
    }

    private static async Task<FlexibleBenefitTestContext> CreateTestContextAsync()
    {
        var folder = Path.Combine(Path.GetTempPath(), "FlexibleBenefitDatabaseTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        var sourceFile = Path.Combine(folder, "flex-plan-v1.pdf");
        await File.WriteAllTextAsync(sourceFile, "version one flexible benefits");
        var database = new DatabaseService(databasePath);
        Assert.True(database.Initialize().Succeeded);
        var clientId = await InsertClientAsync(databasePath, "Flexible Test Client", "FLEX-001");
        return new FlexibleBenefitTestContext(folder, databasePath, sourceFile, database, clientId);
    }

    private static async Task<long> InsertClientAsync(string databasePath, string clientName, string clientNumber)
    {
        await using var connection = await OpenAsync(databasePath);
        await using var command = connection.CreateCommand();
        command.CommandText = @"
INSERT INTO Clients (client_name, client_number, carrier, status, date_created, date_updated)
VALUES ($client_name, $client_number, 'National HealthClaim', 'Active', $now, $now);
SELECT last_insert_rowid();";
        command.Parameters.AddWithValue("$client_name", clientName);
        command.Parameters.AddWithValue("$client_number", clientNumber);
        command.Parameters.AddWithValue("$now", DateTime.UtcNow.ToString("O"));
        return Convert.ToInt64(await command.ExecuteScalarAsync());
    }

    private static async Task<FlexibleBenefitExtractionResult> BuildExtractionAsync(string sourceFile, decimal annualAmount)
    {
        await using var sourceStream = File.OpenRead(sourceFile);
        var sourceHash = Convert.ToHexString(await SHA256.HashDataAsync(sourceStream));
        var extraction = new FlexibleBenefitExtractionResult
        {
            SourcePath = sourceFile,
            SourceHash = sourceHash,
            Confidence = 0.95m,
            Classification = new FlexibleBenefitDocumentClassification
            {
                Carrier = "National HealthClaim",
                Administrator = "National HealthClaim",
                DocumentType = FlexibleBenefitDocumentTypes.FlexibleBenefitsPlanDocument,
                DocumentSubtype = FlexibleBenefitDocumentSubtype.CombinedEnrollmentAndAgreement,
                IsSigned = true,
                Confidence = 0.95m
            },
            Client = new FlexibleBenefitClientIdentity
            {
                LegalName = "Flexible Test Client Inc.",
                DbaName = "Flexible Test Client",
                Address = "100 Test Street",
                City = "Vancouver",
                Province = "BC",
                PostalCode = "V6B 1A1",
                Phone = "604-555-0100"
            },
            Plan = new FlexibleBenefitPlan
            {
                PlanName = "LifeHealth+",
                PlanType = FlexibleBenefitPlanType.CombinedHcsaAndLsa,
                HcsaEnabled = true,
                LsaEnabled = true,
                EffectiveDate = new DateOnly(2023, 4, 1),
                BenefitYearType = FlexibleBenefitYearType.CalendarYear,
                DefaultAllocationType = FlexibleBenefitAllocationType.AllLsa,
                DefaultHcsaPercentage = 0m,
                DefaultLsaPercentage = 100m,
                RunoffDays = 60,
                CarryForwardType = FlexibleBenefitCarryForwardType.Credit,
                CreditCarryForwardEnabled = true,
                ChildDependentCutoffAge = 21,
                StudentDependentCutoffAge = 25,
                FundingMethod = FlexibleBenefitFundingMethod.Pad,
                PadRequested = true,
                PayAsYouGo = true,
                SetupFee = 0m,
                AdministrationFeePercentage = 9m,
                TaxTreatment = "GST on administration fee"
            }
        };

        var planClass = new FlexibleBenefitClass
        {
            SourceKey = "class-all",
            ClassName = "All Eligible",
            AnnualSingleAmount = annualAmount,
            AnnualCoupleAmount = annualAmount,
            AnnualFamilyAmount = annualAmount,
            WaitingPeriodDays = 90,
            NewHireProrationEnabled = false,
            SourcePageNumber = 5,
            ExtractionConfidence = 0.95m
        };
        extraction.Classes.Add(planClass);
        var category = new FlexibleBenefitCategory
        {
            SourceKey = "category-pay-all",
            CategoryCode = "PAY_ALL_RECEIPTS_AS_SUBMITTED",
            CategoryGroup = "General",
            CategoryName = "Pay All Receipts As Submitted",
            Enabled = false,
            AccountType = FlexibleBenefitAccountType.Lsa,
            SourcePageNumber = 4,
            ExtractionConfidence = 0.95m
        };
        extraction.Categories.Add(category);
        var rule = new FlexibleBenefitEligibilityRule
        {
            SourceKey = "rule-waiting",
            ClassSourceKey = planClass.SourceKey,
            RuleType = "WaitingPeriod",
            Operator = "Equals",
            RuleValue = "90",
            RuleUnit = "Days",
            SourcePageNumber = 5,
            ExtractionConfidence = 0.95m
        };
        extraction.EligibilityRules.Add(rule);
        var funding = new FlexibleBenefitFundingConfiguration
        {
            SourceKey = "funding-1",
            FundingMethod = FlexibleBenefitFundingMethod.Pad,
            PadSelected = true,
            OnlineBankPaymentAllowed = true,
            ChequeAllowed = true,
            ClaimsPaidOnlyAfterFundsClear = true,
            InterestPaidOnFunds = false,
            SourcePageNumber = 7,
            ExtractionConfidence = 0.95m
        };
        extraction.FundingConfigurations.Add(funding);
        var fee = new FlexibleBenefitFee
        {
            SourceKey = "fee-setup",
            FeeType = "SetupFee",
            FeeCalculationType = "FlatAmount",
            FlatAmount = 0m,
            SourcePageNumber = 7,
            ExtractionConfidence = 0.95m
        };
        extraction.Fees.Add(fee);
        var agreement = new FlexibleBenefitAgreement
        {
            SourceKey = "agreement-1",
            AgreementType = "AdministrativeServicesIndemnityAgreement",
            AgreementDate = new DateOnly(2023, 2, 14),
            ExecutionDate = new DateTime(2023, 2, 15),
            GoverningProvince = "Alberta",
            SubscriberSignerName = "Test Subscriber",
            AdministratorSignerName = "Test Administrator",
            SubscriberSignaturePresent = true,
            AdministratorSignaturePresent = true,
            ElectronicSignature = true,
            SourcePageNumber = 13
        };
        extraction.Agreements.Add(agreement);
        var contact = new FlexibleBenefitContact
        {
            SourceKey = "contact-admin",
            ContactRole = "PlanAdministrator",
            ContactName = "Plan Admin",
            Email = "admin@example.com",
            IsPrimary = true,
            SourcePageNumber = 1,
            ExtractionConfidence = 0.95m
        };
        extraction.Contacts.Add(contact);

        AddSource(extraction, "FlexibleBenefitPlan", "plan", "RunoffDays", "60 days", "60", 5);
        AddSource(extraction, "FlexibleBenefitClass", planClass.SourceKey, "AnnualSingleAmount", annualAmount.ToString(), annualAmount.ToString(), 5);
        AddSource(extraction, "FlexibleBenefitCategory", category.SourceKey, "Enabled", "Unchecked", "false", 4);
        AddSource(extraction, "FlexibleBenefitEligibilityRule", rule.SourceKey, "RuleValue", "90 days", "90", 5);
        AddSource(extraction, "FlexibleBenefitFundingConfiguration", funding.SourceKey, "PadSelected", "Selected", "true", 7);
        AddSource(extraction, "FlexibleBenefitFee", fee.SourceKey, "FlatAmount", "$0", "0", 7);
        AddSource(extraction, "FlexibleBenefitAgreement", agreement.SourceKey, "SubscriberSignaturePresent", "Signature ink", "true", 13);
        AddSource(extraction, "FlexibleBenefitAgreement", agreement.SourceKey, "AdministratorSignaturePresent", "Signature ink", "true", 13);
        AddSource(extraction, "FlexibleBenefitContact", contact.SourceKey, "Email", "admin@example.com", "admin@example.com", 1);
        return extraction;
    }

    private static void AddSource(
        FlexibleBenefitExtractionResult extraction,
        string entityType,
        string entityKey,
        string fieldName,
        string rawValue,
        string normalizedValue,
        int page)
    {
        extraction.SourceFields.Add(new FlexibleBenefitPlanSourceField
        {
            EntityType = entityType,
            EntityKey = entityKey,
            FieldName = fieldName,
            RawValue = rawValue,
            NormalizedValue = normalizedValue,
            SourcePageNumber = page,
            BoundingBoxJson = "{\"x\":10,\"y\":10,\"width\":20,\"height\":10}",
            ExtractionMethod = "Test fixture",
            ExtractionConfidence = 0.95m,
            ReviewStatus = FlexibleBenefitReviewStatuses.Pending
        });
    }

    private static async Task<SqliteConnection> OpenAsync(string databasePath)
    {
        var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA foreign_keys = ON;";
        await command.ExecuteNonQueryAsync();
        return connection;
    }

    private static async Task<int> CountAsync(SqliteConnection connection, string table)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = $"SELECT COUNT(*) FROM {table};";
        return Convert.ToInt32(await command.ExecuteScalarAsync());
    }

    private static async Task<int> ForeignKeyViolationCountAsync(SqliteConnection connection)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA foreign_key_check;";
        await using var reader = await command.ExecuteReaderAsync();
        var count = 0;
        while (await reader.ReadAsync())
        {
            count++;
        }

        return count;
    }

    private static async Task<int> ScalarIntAsync(SqliteConnection connection, string sql)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        return Convert.ToInt32(await command.ExecuteScalarAsync());
    }

    private static async Task<long> ScalarLongAsync(SqliteConnection connection, string sql)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        return Convert.ToInt64(await command.ExecuteScalarAsync());
    }

    private static async Task<string> ScalarStringAsync(SqliteConnection connection, string sql)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        return Convert.ToString(await command.ExecuteScalarAsync()) ?? "";
    }

    private sealed record FlexibleBenefitTestContext(
        string Folder,
        string DatabasePath,
        string SourceFile,
        DatabaseService Database,
        long ClientId) : IDisposable
    {
        public void Dispose()
        {
            SqliteConnection.ClearAllPools();
            if (Directory.Exists(Folder))
            {
                Directory.Delete(Folder, true);
            }
        }
    }
}
