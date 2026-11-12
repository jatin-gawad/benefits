using System.Collections.Generic;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class CanonicalControlledBookletReprocessingWorkflowTests
{
    [Fact]
    public void NormalizeProvisionRows_SplitsLabeledRatesAndRemovesDerivedCostSharing()
    {
        var rows = PlanDesignCanonicalMapper.NormalizeProvisionRows(
        [
            new PlanDesignRecord
            {
                ClassName = "Class A",
                BenefitCategory = "Prescription Drugs",
                BenefitName = "Prescription Drugs",
                CoverageDescription = "Coinsurance: 80% for SDMP drugs; 90% at Costco Pharmacies; 80% at other pharmacies",
                ReimbursementText = "80%; 90%"
            },
            new PlanDesignRecord
            {
                ClassName = "Class A",
                BenefitCategory = "Prescription Drugs",
                BenefitName = "Prescription Drugs",
                CoverageDescription = "Cost Sharing: SDMP drugs: plan 80% / member 20%; Costco Pharmacies: plan 90% / member 10%; Other pharmacies: plan 80% / member 20%",
                ReimbursementText = "80%; 90%"
            }
        ]);

        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { 80m, 80m, 90m }, rows.Select(row => row.ReimbursementPercent!.Value).OrderBy(value => value));
        Assert.Contains(rows, row => row.ProvisionLabel.Equals("SDMP drugs Coinsurance", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(rows, row => row.ProvisionLabel.Equals("Costco Pharmacies Coinsurance", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(rows, row => row.ProvisionLabel.Equals("Other pharmacies Coinsurance", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(rows, row => row.CoverageDescription.StartsWith("Cost Sharing:", StringComparison.OrdinalIgnoreCase));
        Assert.All(rows, row => Assert.False(
            CanonicalSourceLineageValueGuardrails.HasMultiplePercentageValues(row.ReimbursementText)));
    }

    [Fact]
    public void NormalizeProvisionRows_RemovesRedundantDentalSummaryRowsButKeepsDistinctProvisions()
    {
        var rows = PlanDesignCanonicalMapper.NormalizeProvisionRows(
        [
            DentalRow("Dental Care", "Deductible: nil"),
            DentalRow("Dental Care", "Recall: once in any period of 6 months"),
            DentalRow("Dental Care", "Basic Services: 80%"),
            DentalRow("Dental Care", "Major Restorative Services: 50%"),
            DentalRow("Dental Care", "Orthodontics: 50%"),
            DentalRow("Dental Care", "Basic and Major Maximum: $2,500 per calendar year"),
            DentalRow("Dental Care", "Orthodontic Maximum: $2,500 lifetime"),
            DentalRow("Dental Care", "Orthodontic Eligibility: adults and dependent children"),
            DentalRow("Dental Basic Services", "Type A - Basic services at 80%"),
            DentalRow("Dental Major Restorative Services", "Type B - Major restorative services at 50%"),
            DentalRow("Dental Orthodontic Services", "Type C - Orthodontic services at 50%")
        ]);

        Assert.Equal(6, rows.Count);
        Assert.Contains(rows, row => row.ProvisionLabel == "Deductible");
        Assert.Contains(rows, row => row.ProvisionLabel == "Recall");
        Assert.Contains(rows, row => row.ProvisionLabel == "Orthodontic Eligibility");
        Assert.Contains(rows, row => row.ProvisionLabel == "Basic Services");
        Assert.Contains(rows, row => row.ProvisionLabel == "Major Restorative Services");
        Assert.Contains(rows, row => row.ProvisionLabel == "Orthodontic Services");
        Assert.DoesNotContain(rows, row => row.ProvisionLabel == "Basic and Major Maximum");
        Assert.DoesNotContain(rows, row => row.ProvisionLabel == "Orthodontic Maximum");
    }

    [Fact]
    public void NormalizeProvisionRows_RemovesSyntheticClassSummary()
    {
        var rows = PlanDesignCanonicalMapper.NormalizeProvisionRows(
        [
            new PlanDesignRecord
            {
                ClassName = "Class A",
                BenefitCategory = "Class Summary",
                BenefitName = "Class Benchmark Summary",
                ProvisionCode = "CLASS_BENCHMARK_SUMMARY",
                CoverageDescription = "Synthetic benchmark text"
            },
            new PlanDesignRecord
            {
                ClassName = "Class A",
                BenefitCategory = "AD&D",
                BenefitName = "AD&D",
                CoverageDescription = "AD&D Benefit Amount: Same as employee life"
            }
        ]);

        var row = Assert.Single(rows);
        Assert.Equal("AD&D", row.BenefitCategory);
        Assert.DoesNotContain(rows, item => item.ProvisionCode == "CLASS_BENCHMARK_SUMMARY");
    }

    [Fact]
    public void TopLevelBenefitKey_NormalizesAdAndDFromLegacyAndCanonicalShapes()
    {
        var legacy = new PlanDesignRecord
        {
            BenefitName = "AD&D"
        };
        var canonical = new PlanDesignRecord
        {
            BenefitCategory = "AD&D",
            BenefitName = "AD&D",
            ProvisionCode = "PLAN_SUMMARY"
        };

        Assert.Equal("AD_AND_D", CanonicalControlledBookletReprocessingService.TopLevelBenefitKey(legacy));
        Assert.Equal(
            CanonicalControlledBookletReprocessingService.TopLevelBenefitKey(legacy),
            CanonicalControlledBookletReprocessingService.TopLevelBenefitKey(canonical));
    }

    [Fact]
    public void NormalizeProvisionRows_PreservesSeparateLifeAndAdAndDIdentitiesUnderSharedHeading()
    {
        var rows = PlanDesignCanonicalMapper.NormalizeProvisionRows(
        [
            new PlanDesignRecord
            {
                ClassName = "All employees",
                BenefitCategory = "Life / AD&D",
                BenefitName = "Life",
                CoverageDescription = "Employee Life Amount: $50,000"
            },
            new PlanDesignRecord
            {
                ClassName = "All employees",
                BenefitCategory = "Life / AD&D",
                BenefitName = "AD&D",
                CoverageDescription = "AD&D Amount: Equal to Employee Life coverage"
            }
        ]);

        Assert.Equal(2, rows.Count);
        Assert.Contains(rows, row => PlanDesignCanonicalMapper.ResolveProvisionBenefitCode(row) == "LIFE");
        Assert.Contains(rows, row => PlanDesignCanonicalMapper.ResolveProvisionBenefitCode(row) == "AD_AND_D");
    }

    [Fact]
    public void ProvisionComparisonKey_RemainsStableWhenOnlyTheValueChanges()
    {
        var before = PlanDesignCanonicalMapper.NormalizeProvisionRows(
        [
            PrescriptionDrugRow("Drug Maximum: $10,000 per calendar year"),
            PrescriptionDrugRow("Deductible: nil")
        ]);
        var after = PlanDesignCanonicalMapper.NormalizeProvisionRows(
        [
            PrescriptionDrugRow("Drug Maximum: $15,000 per calendar year"),
            PrescriptionDrugRow("Deductible: nil")
        ]);
        var beforeMaximum = Assert.Single(before, row => row.ProvisionLabel == "Drug Maximum");
        var afterMaximum = Assert.Single(after, row => row.ProvisionLabel == "Drug Maximum");

        Assert.Equal(
            PlanDesignCanonicalMapper.BuildProvisionComparisonKey(beforeMaximum),
            PlanDesignCanonicalMapper.BuildProvisionComparisonKey(afterMaximum));
        Assert.NotEqual(
            CanonicalControlledBookletReprocessingService.BuildDisplayValue(beforeMaximum),
            CanonicalControlledBookletReprocessingService.BuildDisplayValue(afterMaximum));
    }

    [Fact]
    public void ProvisionComparisonKey_TreatsSingletonCoverageAsTheExistingPlanSummary()
    {
        var before = Assert.Single(PlanDesignCanonicalMapper.NormalizeProvisionRows(
        [
            new PlanDesignRecord
            {
                ClassName = "Class 02",
                BenefitCategory = "EAP",
                BenefitName = "EAP",
                CoverageDescription = "Sparrow"
            }
        ]));
        var after = Assert.Single(PlanDesignCanonicalMapper.NormalizeProvisionRows(
        [
            new PlanDesignRecord
            {
                ClassName = "Class 02",
                BenefitCategory = "EAP",
                BenefitName = "EAP",
                CoverageDescription = "Coverage: Sparrow"
            }
        ]));

        Assert.Equal("Plan Summary", before.ProvisionLabel);
        Assert.Equal("Plan Summary", after.ProvisionLabel);
        Assert.Equal(
            PlanDesignCanonicalMapper.BuildProvisionComparisonKey(before),
            PlanDesignCanonicalMapper.BuildProvisionComparisonKey(after));
    }

    [Fact]
    public async Task SavePlanDesign_NormalizesRowsAndPersistsCanonicalFeatureNames()
    {
        var folder = CreateTestFolder();
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        var bookletPath = Path.Combine(folder, "Canonical Provision Booklet.pdf");
        await File.WriteAllTextAsync(bookletPath, "canonical provision source");

        try
        {
            var database = new DatabaseService(databasePath);
            var save = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Canonical Provision Client", Carrier = "Equitable" },
                [bookletPath],
                [
                    new PlanDesignRecord
                    {
                        ClassName = "Class A",
                        BenefitCategory = "Prescription Drugs",
                        BenefitName = "Prescription Drugs",
                        CoverageDescription = "Coinsurance: 80% for SDMP drugs; 90% at Costco Pharmacies; 80% at other pharmacies",
                        ReimbursementText = "80%; 90%",
                        SourceFileName = Path.GetFileName(bookletPath)
                    },
                    new PlanDesignRecord
                    {
                        ClassName = "Class A",
                        BenefitCategory = "Prescription Drugs",
                        BenefitName = "Prescription Drugs",
                        CoverageDescription = "Cost Sharing: SDMP drugs: plan 80% / member 20%; Costco Pharmacies: plan 90% / member 10%; Other pharmacies: plan 80% / member 20%",
                        ReimbursementText = "80%; 90%",
                        SourceFileName = Path.GetFileName(bookletPath)
                    }
                ]);

            Assert.True(save.Succeeded, save.Message);
            Assert.Equal(3, save.InsertedRows);
            Assert.Equal(3, await CountAsync(databasePath, "PlanDesign"));
            Assert.Equal(
                new[]
                {
                    "Costco Pharmacies Coinsurance",
                    "Other pharmacies Coinsurance",
                    "SDMP drugs Coinsurance"
                },
                await ReadFeatureNamesAsync(databasePath));
        }
        finally
        {
            Cleanup(folder);
        }
    }

    [Fact]
    public void BookletReviewCounts_SeparateMediumRecommendationsFromRequiredReview()
    {
        var folder = CreateTestFolder();
        var reviewPath = Path.Combine(folder, "Booklet_Extraction_Review.xlsx");
        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add("Booklet_Extraction_Review");
                sheet.Cell(1, 1).Value = "Benefit Category";
                sheet.Cell(1, 2).Value = "Benefit Name";
                sheet.Cell(1, 3).Value = "ReviewFlag";
                sheet.Cell(1, 4).Value = "Confidence";
                sheet.Cell(2, 1).Value = "Class Summary";
                sheet.Cell(2, 2).Value = "Class Benchmark Summary";
                sheet.Cell(2, 3).Value = "Review";
                sheet.Cell(2, 4).Value = "Medium";
                sheet.Cell(3, 1).Value = "Prescription Drugs";
                sheet.Cell(3, 2).Value = "Prescription Drugs";
                sheet.Cell(3, 3).Value = "No";
                sheet.Cell(3, 4).Value = "High";
                sheet.Cell(4, 1).Value = "Vision";
                sheet.Cell(4, 2).Value = "Vision";
                sheet.Cell(4, 3).Value = "Review";
                sheet.Cell(4, 4).Value = "Low";
                workbook.SaveAs(reviewPath);
            }

            var counts = BookletExtractorRunner.CountReviewRows(reviewPath);

            Assert.Equal(1, counts.Required);
            Assert.Equal(1, counts.Recommended);
        }
        finally
        {
            Cleanup(folder);
        }
    }

    [Fact]
    public void ComparisonDisplay_UsesReimbursementAndDoesNotRepeatStructuredDetails()
    {
        const string maximum = "$500 per practitioner per calendar year";
        var row = new PlanDesignRecord
        {
            ReimbursementPercent = 80m,
            ReimbursementText = "80%",
            MaximumAmount = maximum,
            Frequency = "up to 180 days per trip",
            CoverageDescription = $"Coinsurance: 80%\r\nDeductible: $45 Single / $45 Family\r\nMaximum: {maximum}\r\nTrip Duration: up to 180 days per trip"
        };

        var display = CanonicalControlledBookletReprocessingService.BuildDisplayValue(row);

        Assert.Contains("Reimbursement: 80%", display);
        Assert.DoesNotContain("Rate:", display);
        Assert.DoesNotContain("Coinsurance: 80%", display);
        Assert.Contains("Deductible: $45 Single / $45 Family", display);
        Assert.Equal(1, display.Split("Maximum:", StringSplitOptions.None).Length - 1);
        Assert.Equal(1, display.Split("up to 180 days per trip", StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public async Task CleanStage_DoesNotChangeLiveRows_AndApprovedCommitCreatesBackupAndAudit()
    {
        var folder = CreateTestFolder();
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        var bookletPath = Path.Combine(folder, "Synthetic Community Services Booklet.pdf");
        var workbookPath = Path.Combine(folder, "staged-plan-design.xlsx");
        await File.WriteAllTextAsync(bookletPath, "stable source booklet");

        try
        {
            var database = new DatabaseService(databasePath);
            var seed = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Synthetic Community Services Ltd.", Carrier = "Industrial Alliance" },
                [bookletPath],
                [
                    PlanRow("Dental", "Basic Services", 100m, bookletPath),
                    PlanRow("Dental", "Major Restorative Services", 60m, bookletPath)
                ]);
            Assert.True(seed.Succeeded, seed.Message);
            Assert.True(seed.DocumentId.HasValue);
            WritePlanDesignWorkbook(
                workbookPath,
                bookletPath,
                ("Dental", "Basic Services", "80%", "$2,500 per calendar year"),
                ("Dental", "Major Restorative Services", "50%", "$2,500 per calendar year"));

            var workflow = CreateWorkflow(database, bookletPath, workbookPath, "Synthetic Community Services");
            var staged = await workflow.StageAsync(new ControlledBookletReprocessingRequest
            {
                SourceDocumentId = seed.DocumentId.Value,
                OutputRoot = Path.Combine(folder, "staging"),
                RequestedBy = "workflow-test"
            });

            Assert.Equal(ControlledBookletReprocessingStatuses.ReadyForApproval, staged.Status);
            Assert.True(staged.CanApprove);
            Assert.Equal(0, staged.BlockingIssueCount);
            Assert.Equal(2, staged.BeforeRowCount);
            Assert.Equal(2, staged.StagedRowCount);
            Assert.Equal(2, staged.ChangedCount);
            Assert.Equal(new[] { 60m, 100m }, await ReadRatesAsync(databasePath, seed.DocumentId.Value));

            var activeRun = await workflow.GetActiveRunForDocumentAsync(seed.DocumentId.Value);
            Assert.NotNull(activeRun);
            Assert.Equal(staged.RunId, activeRun.RunId);
            Assert.Equal(ControlledBookletReprocessingStatuses.ReadyForApproval, activeRun.Status);

            var committed = await workflow.ApproveAndCommitAsync(staged.RunId, "approved-user");

            Assert.Equal(ControlledBookletReprocessingStatuses.Committed, committed.Status);
            Assert.Equal(2, committed.PostCommitRowCount);
            Assert.True(File.Exists(committed.BackupPath));
            Assert.Equal(new[] { 50m, 80m }, await ReadRatesAsync(databasePath, seed.DocumentId.Value));
            Assert.Equal("ok", await QuickCheckAsync(committed.BackupPath));
            Assert.Equal(1, await CountAsync(databasePath, "ControlledBookletReprocessingRun"));
            Assert.Equal(2, await CountAsync(databasePath, "ControlledBookletReprocessingChange", "change_type = 'Changed'"));
            Assert.Null(await workflow.GetActiveRunForDocumentAsync(seed.DocumentId.Value));
        }
        finally
        {
            Cleanup(folder);
        }
    }

    [Fact]
    public async Task OutOfRangePercentage_BlocksApproval_AndLeavesLiveRowsUntouched()
    {
        var folder = CreateTestFolder();
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        var bookletPath = Path.Combine(folder, "Percentage Guardrail Booklet.pdf");
        var workbookPath = Path.Combine(folder, "unsafe-plan-design.xlsx");
        await File.WriteAllTextAsync(bookletPath, "percentage guardrail source");

        try
        {
            var database = new DatabaseService(databasePath);
            var seed = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Percentage Guardrail Society", Carrier = "Beneva" },
                [bookletPath],
                [PlanRow("Dental", "Basic Services", 80m, bookletPath)]);
            Assert.True(seed.Succeeded, seed.Message);
            WritePlanDesignWorkbook(
                workbookPath,
                bookletPath,
                ("Dental", "Basic Services", "8050%", "$2,500 per calendar year"));

            var workflow = CreateWorkflow(database, bookletPath, workbookPath, "Percentage Guardrail Society");
            var staged = await workflow.StageAsync(new ControlledBookletReprocessingRequest
            {
                SourceDocumentId = seed.DocumentId!.Value,
                OutputRoot = Path.Combine(folder, "staging")
            });

            Assert.Equal(ControlledBookletReprocessingStatuses.Blocked, staged.Status);
            Assert.False(staged.CanApprove);
            Assert.Contains(staged.Issues, issue =>
                issue.IssueCode == ControlledBookletReprocessingIssueCodes.PercentageOutOfRange && issue.IsBlocking);
            Assert.Equal(new[] { 80m }, await ReadRatesAsync(databasePath, seed.DocumentId.Value));
            await Assert.ThrowsAsync<InvalidOperationException>(() => workflow.ApproveAndCommitAsync(staged.RunId, "approver"));

            var loadedBlockedRun = await workflow.GetActiveRunForDocumentAsync(seed.DocumentId.Value);
            Assert.NotNull(loadedBlockedRun);
            Assert.Equal(staged.RunId, loadedBlockedRun.RunId);
            var repeatedStage = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                workflow.StageAsync(new ControlledBookletReprocessingRequest
                {
                    SourceDocumentId = seed.DocumentId.Value,
                    OutputRoot = Path.Combine(folder, "repeated-staging")
                }));
            Assert.Contains($"active controlled reprocessing run {staged.RunId}", repeatedStage.Message);

            await workflow.RejectAsync(staged.RunId, "reviewer", "Guardrail test completed.");
            Assert.Null(await workflow.GetActiveRunForDocumentAsync(seed.DocumentId.Value));
        }
        finally
        {
            Cleanup(folder);
        }
    }

    [Fact]
    public async Task ManualStagedCorrection_RevalidatesAndAuditsWithoutChangingLiveRowsUntilApproval()
    {
        var folder = CreateTestFolder();
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        var bookletPath = Path.Combine(folder, "Manual Correction Booklet.pdf");
        var workbookPath = Path.Combine(folder, "manual-correction-stage.xlsx");
        await File.WriteAllTextAsync(bookletPath, "manual correction source");

        try
        {
            var database = new DatabaseService(databasePath);
            var seed = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Manual Correction Society", Carrier = "Industrial Alliance" },
                [bookletPath],
                [PlanRow("Dental", "Basic Services", 100m, bookletPath)]);
            Assert.True(seed.Succeeded, seed.Message);
            WritePlanDesignWorkbook(
                workbookPath,
                bookletPath,
                ("Dental", "Basic Services", "8050%", "$2,500 per calendar year"));

            var workflow = CreateWorkflow(database, bookletPath, workbookPath, "Manual Correction Society");
            var staged = await workflow.StageAsync(new ControlledBookletReprocessingRequest
            {
                SourceDocumentId = seed.DocumentId!.Value,
                OutputRoot = Path.Combine(folder, "staging")
            });
            Assert.Equal(ControlledBookletReprocessingStatuses.Blocked, staged.Status);
            var stagedRow = Assert.Single(staged.StagedRows);
            var comparisonKey = PlanDesignCanonicalMapper.BuildProvisionComparisonKey(stagedRow);

            var missingReason = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                workflow.ApplyStagedEditAsync(staged.RunId, new ControlledBookletStagedEditRequest
                {
                    ComparisonKey = comparisonKey,
                    ReimbursementPercent = 80m,
                    MaximumAmount = stagedRow.MaximumAmount,
                    CoverageDescription = stagedRow.CoverageDescription
                }));
            Assert.Contains("reason is required", missingReason.Message, StringComparison.OrdinalIgnoreCase);

            var corrected = await workflow.ApplyStagedEditAsync(
                staged.RunId,
                new ControlledBookletStagedEditRequest
                {
                    ComparisonKey = comparisonKey,
                    ReimbursementPercent = 80m,
                    MaximumAmount = stagedRow.MaximumAmount,
                    Frequency = stagedRow.Frequency,
                    WaitingPeriod = stagedRow.WaitingPeriod,
                    TerminationAge = stagedRow.TerminationAge,
                    CoverageDescription = stagedRow.CoverageDescription,
                    Notes = stagedRow.Notes,
                    Reason = "Confirmed against the source booklet.",
                    EditedBy = "plan-reviewer"
                });

            Assert.Equal(ControlledBookletReprocessingStatuses.ReadyForApproval, corrected.Status);
            Assert.True(corrected.CanApprove);
            Assert.DoesNotContain(corrected.Issues, issue =>
                issue.IssueCode == ControlledBookletReprocessingIssueCodes.PercentageOutOfRange);
            Assert.Equal(80m, Assert.Single(corrected.StagedRows).ReimbursementPercent);
            Assert.Equal(new[] { 100m }, await ReadRatesAsync(databasePath, seed.DocumentId.Value));
            var audit = Assert.Single(corrected.Edits);
            Assert.Equal("reimbursement_percent", audit.FieldName);
            Assert.Equal("8050%", audit.PreviousValue);
            Assert.Equal("80%", audit.NewValue);
            Assert.Equal("Confirmed against the source booklet.", audit.Reason);
            Assert.Equal("plan-reviewer", audit.EditedBy);
            Assert.Equal(1, await CountAsync(databasePath, "ControlledBookletReprocessingEdit"));

            var committed = await workflow.ApproveAndCommitAsync(corrected.RunId, "approver");
            Assert.Equal(ControlledBookletReprocessingStatuses.Committed, committed.Status);
            Assert.Equal(new[] { 80m }, await ReadRatesAsync(databasePath, seed.DocumentId.Value));
            Assert.Single(committed.Edits);
        }
        finally
        {
            Cleanup(folder);
        }
    }

    [Fact]
    public async Task CrossClientDuplicateHash_BlocksBeforeExtraction()
    {
        var folder = CreateTestFolder();
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        var firstBooklet = Path.Combine(folder, "First Client Booklet.pdf");
        var secondBooklet = Path.Combine(folder, "Second Client Booklet.pdf");
        await File.WriteAllTextAsync(firstBooklet, "same booklet bytes");
        await File.WriteAllTextAsync(secondBooklet, "same booklet bytes");

        try
        {
            var database = new DatabaseService(databasePath);
            var first = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "First Client Society", Carrier = "Industrial Alliance" },
                [firstBooklet],
                [PlanRow("Dental", "Basic Services", 80m, firstBooklet)]);
            var second = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Second Client Society", Carrier = "Industrial Alliance" },
                [secondBooklet],
                [PlanRow("Dental", "Basic Services", 80m, secondBooklet)]);
            Assert.True(first.Succeeded, first.Message);
            Assert.True(second.Succeeded, second.Message);
            var extractionCalled = false;
            var workflow = new CanonicalControlledBookletReprocessingService(
                database,
                (request, progress, token) =>
                {
                    extractionCalled = true;
                    return Task.FromResult(new BookletExtractionResult());
                },
                (paths, token) => Task.FromResult<IReadOnlyList<BookletClientIdentity>>(new List<BookletClientIdentity>()));

            var staged = await workflow.StageAsync(new ControlledBookletReprocessingRequest
            {
                SourceDocumentId = first.DocumentId!.Value,
                OutputRoot = Path.Combine(folder, "staging")
            });

            Assert.False(extractionCalled);
            Assert.Equal(ControlledBookletReprocessingStatuses.Blocked, staged.Status);
            Assert.Contains(staged.Issues, issue =>
                issue.IssueCode == ControlledBookletReprocessingIssueCodes.CrossClientDuplicateHash && issue.IsBlocking);

            await workflow.RejectAsync(staged.RunId, "reviewer", "Duplicate ownership confirmed.");
            await SetDocumentProcessingStatusAsync(
                databasePath,
                second.DocumentId!.Value,
                DocumentIdentityReconciliationStatuses.IdentityQuarantined);

            var availableDocuments = await workflow.GetDocumentsAsync();
            var availableFirst = Assert.Single(availableDocuments, document => document.DocumentId == first.DocumentId);
            Assert.Equal(0, availableFirst.CrossClientHashAssignmentCount);
            Assert.DoesNotContain(availableDocuments, document => document.DocumentId == second.DocumentId);

            var restaged = await workflow.StageAsync(new ControlledBookletReprocessingRequest
            {
                SourceDocumentId = first.DocumentId.Value,
                OutputRoot = Path.Combine(folder, "restaged")
            });
            Assert.DoesNotContain(restaged.Issues, issue =>
                issue.IssueCode == ControlledBookletReprocessingIssueCodes.CrossClientDuplicateHash);
            Assert.Contains(restaged.Issues, issue =>
                issue.IssueCode == ControlledBookletReprocessingIssueCodes.SourceIdentityUnconfirmed);
        }
        finally
        {
            Cleanup(folder);
        }
    }

    [Fact]
    public async Task GetDocuments_ReturnsRegisteredBookletsAlphabetically()
    {
        var folder = CreateTestFolder();
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        var zuluBooklet = Path.Combine(folder, "Zulu Booklet.pdf");
        var alphaBooklet = Path.Combine(folder, "Alpha Booklet.pdf");
        await File.WriteAllTextAsync(zuluBooklet, "zulu booklet bytes");
        await File.WriteAllTextAsync(alphaBooklet, "alpha booklet bytes");

        try
        {
            var database = new DatabaseService(databasePath);
            var zulu = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Zulu Foundation", Carrier = "Manulife" },
                [zuluBooklet],
                [PlanRow("Dental", "Basic Services", 80m, zuluBooklet)]);
            var alpha = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Alpha Foundation", Carrier = "Industrial Alliance" },
                [alphaBooklet],
                [PlanRow("Dental", "Basic Services", 80m, alphaBooklet)]);
            Assert.True(zulu.Succeeded, zulu.Message);
            Assert.True(alpha.Succeeded, alpha.Message);
            await SetPlanDesignReimbursementAsync(databasePath, zulu.DocumentId!.Value, 8050m);

            var workflow = new CanonicalControlledBookletReprocessingService(database);
            var documents = await workflow.GetDocumentsAsync();

            Assert.Equal(
                new[] { "Alpha Foundation", "Zulu Foundation" },
                documents.Select(document => document.ClientName));
            Assert.Equal(
                1,
                documents.Single(document => document.DocumentId == zulu.DocumentId).OutOfRangePercentageCount);
        }
        finally
        {
            Cleanup(folder);
        }
    }

    [Fact]
    public async Task GetDocuments_CleansLegacyClientNameAndAvoidsUnknownCarrierFallback()
    {
        var folder = CreateTestFolder();
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        var legacyBooklet = Path.Combine(folder, "Legacy Booklet.pdf");
        var unknownBooklet = Path.Combine(folder, "Unknown Booklet.pdf");
        await File.WriteAllTextAsync(legacyBooklet, "legacy booklet bytes");
        await File.WriteAllTextAsync(unknownBooklet, "unknown booklet bytes");

        try
        {
            var database = new DatabaseService(databasePath);
            var legacy = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Legacy Foundation", Carrier = "Beneva" },
                [legacyBooklet],
                [PlanRow("Dental", "Basic Services", 80m, legacyBooklet)]);
            var unknown = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Unknown Client", Carrier = "Manulife" },
                [unknownBooklet],
                [PlanRow("Dental", "Basic Services", 80m, unknownBooklet)]);
            Assert.True(legacy.Succeeded, legacy.Message);
            Assert.True(unknown.Succeeded, unknown.Message);
            await SetClientNameAsync(
                databasePath,
                legacy.ClientId!.Value,
                "Legacy Foundation [[PAGE 2]] Introduction Policyholder Legacy Foundation");
            await SetDocumentCarrierAsync(databasePath, unknown.DocumentId!.Value, null);

            var documents = await new CanonicalControlledBookletReprocessingService(database).GetDocumentsAsync();
            var legacyDocument = Assert.Single(documents, document => document.DocumentId == legacy.DocumentId);
            var unknownDocument = Assert.Single(documents, document => document.DocumentId == unknown.DocumentId);

            Assert.Equal("Legacy Foundation", legacyDocument.ClientName);
            Assert.DoesNotContain("[[PAGE", legacyDocument.DisplayName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal("", unknownDocument.Carrier);
            Assert.Contains("Carrier not recorded", unknownDocument.DisplayName, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            Cleanup(folder);
        }
    }

    [Fact]
    public async Task FileChangedAfterStaging_FailsCommitWithoutReplacingRows()
    {
        var folder = CreateTestFolder();
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        var bookletPath = Path.Combine(folder, "Commit Identity Booklet.pdf");
        var workbookPath = Path.Combine(folder, "commit-identity-stage.xlsx");
        await File.WriteAllTextAsync(bookletPath, "original booklet bytes");

        try
        {
            var database = new DatabaseService(databasePath);
            var seed = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Commit Identity Foundation", Carrier = "Manulife" },
                [bookletPath],
                [PlanRow("Dental", "Basic Services", 100m, bookletPath)]);
            Assert.True(seed.Succeeded, seed.Message);
            WritePlanDesignWorkbook(
                workbookPath,
                bookletPath,
                ("Dental", "Basic Services", "80%", "$2,000 per calendar year"));
            var workflow = CreateWorkflow(database, bookletPath, workbookPath, "Commit Identity Foundation");
            var staged = await workflow.StageAsync(new ControlledBookletReprocessingRequest
            {
                SourceDocumentId = seed.DocumentId!.Value,
                OutputRoot = Path.Combine(folder, "staging")
            });
            Assert.True(staged.CanApprove);

            await File.AppendAllTextAsync(bookletPath, " changed after staging");
            var result = await workflow.ApproveAndCommitAsync(staged.RunId, "approver");

            Assert.Equal(ControlledBookletReprocessingStatuses.Failed, result.Status);
            Assert.Contains("changed after staging", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(new[] { 100m }, await ReadRatesAsync(databasePath, seed.DocumentId.Value));
            Assert.True(string.IsNullOrWhiteSpace(result.BackupPath));
        }
        finally
        {
            Cleanup(folder);
        }
    }

    private static CanonicalControlledBookletReprocessingService CreateWorkflow(
        DatabaseService database,
        string bookletPath,
        string workbookPath,
        string detectedClientName)
    {
        return new CanonicalControlledBookletReprocessingService(
            database,
            (request, progress, token) => Task.FromResult(new BookletExtractionResult
            {
                OutputFolder = request.OutputFolder,
                MasterWorkbookPath = workbookPath,
                ReviewWorkbookPath = Path.Combine(request.OutputFolder, "review.xlsx"),
                TotalPdfsFound = 1,
                SuccessfullyProcessed = 1,
                ErrorCount = 0,
                ReviewRequiredCount = 0
            }),
            (paths, token) => Task.FromResult<IReadOnlyList<BookletClientIdentity>>(
                new List<BookletClientIdentity>
                {
                    new BookletClientIdentity
                    {
                        SourcePath = bookletPath,
                        SourceFile = Path.GetFileName(bookletPath),
                        ClientName = detectedClientName,
                        Carrier = "Industrial Alliance",
                        ConfidenceScore = 0.95m,
                        IdentitySource = "Document"
                    }
                }));
    }

    private static PlanDesignRecord PlanRow(
        string category,
        string benefit,
        decimal reimbursement,
        string sourcePath)
    {
        return new PlanDesignRecord
        {
            ClassName = "All Classes",
            BenefitCategory = category,
            BenefitName = benefit,
            CoverageDescription = $"{benefit} at {reimbursement:0.####}%",
            ReimbursementPercent = reimbursement,
            ReimbursementText = $"{reimbursement:0.####}%",
            MaximumAmount = "$2,500 per calendar year",
            SourceFileName = Path.GetFileName(sourcePath),
            SourcePage = 10
        };
    }

    private static PlanDesignRecord DentalRow(string category, string description) => new()
    {
        ClassName = "Class A",
        BenefitCategory = category,
        BenefitName = category,
        CoverageDescription = description
    };

    private static PlanDesignRecord PrescriptionDrugRow(string description) => new()
    {
        ClassName = "Class A",
        BenefitCategory = "Prescription Drugs",
        BenefitName = "Prescription Drugs",
        CoverageDescription = description
    };

    private static void WritePlanDesignWorkbook(
        string workbookPath,
        string sourcePath,
        params (string Category, string Benefit, string Reimbursement, string Maximum)[] rows)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("PlanDesign");
        var headers = new[]
        {
            "Client Name", "Client Number", "Carrier", "Class Name", "Benefit Category",
            "Benefit Name", "Coverage Description", "Reimbursement Percent", "Maximum Amount",
            "Frequency", "Waiting Period", "Termination Age", "Notes", "Source File",
            "Source Section", "Source Page"
        };
        for (var column = 0; column < headers.Length; column++)
        {
            sheet.Cell(1, column + 1).Value = headers[column];
        }

        for (var index = 0; index < rows.Length; index++)
        {
            var rowNumber = index + 2;
            var row = rows[index];
            sheet.Cell(rowNumber, 3).Value = "Industrial Alliance";
            sheet.Cell(rowNumber, 4).Value = "All Classes";
            sheet.Cell(rowNumber, 5).Value = row.Category;
            sheet.Cell(rowNumber, 6).Value = row.Benefit;
            sheet.Cell(rowNumber, 7).Value = $"{row.Benefit} at {row.Reimbursement}";
            sheet.Cell(rowNumber, 8).Value = row.Reimbursement;
            sheet.Cell(rowNumber, 9).Value = row.Maximum;
            sheet.Cell(rowNumber, 14).Value = Path.GetFileName(sourcePath);
            sheet.Cell(rowNumber, 15).Value = row.Category;
            sheet.Cell(rowNumber, 16).Value = 10 + index;
        }

        workbook.SaveAs(workbookPath);
    }

    private static async Task<decimal[]> ReadRatesAsync(string databasePath, long documentId)
    {
        await using var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT reimbursement_percent FROM PlanDesign WHERE source_document_id = $document_id ORDER BY reimbursement_percent;";
        command.Parameters.AddWithValue("$document_id", documentId);
        var values = new List<decimal>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            values.Add(reader.GetDecimal(0));
        }

        return values.ToArray();
    }

    private static async Task<string[]> ReadFeatureNamesAsync(string databasePath)
    {
        await using var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT feature_name FROM plan_design_features ORDER BY feature_name COLLATE NOCASE;";
        var values = new List<string>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            values.Add(reader.GetString(0));
        }

        return values.ToArray();
    }

    private static async Task<int> CountAsync(string databasePath, string tableName, string? where = null)
    {
        await using var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = $"SELECT COUNT(*) FROM {tableName}{(string.IsNullOrWhiteSpace(where) ? "" : $" WHERE {where}")};";
        return Convert.ToInt32(await command.ExecuteScalarAsync());
    }

    private static async Task SetDocumentProcessingStatusAsync(
        string databasePath,
        long documentId,
        string status)
    {
        await using var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = "UPDATE Documents SET processing_status = $status WHERE document_id = $document_id;";
        command.Parameters.AddWithValue("$status", status);
        command.Parameters.AddWithValue("$document_id", documentId);
        Assert.Equal(1, await command.ExecuteNonQueryAsync());
    }

    private static async Task SetPlanDesignReimbursementAsync(
        string databasePath,
        long documentId,
        decimal reimbursementPercent)
    {
        await using var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = "UPDATE PlanDesign SET reimbursement_percent = $value WHERE source_document_id = $document_id;";
        command.Parameters.AddWithValue("$value", reimbursementPercent);
        command.Parameters.AddWithValue("$document_id", documentId);
        Assert.True(await command.ExecuteNonQueryAsync() > 0);
    }

    private static async Task SetClientNameAsync(string databasePath, long clientId, string clientName)
    {
        await using var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = "UPDATE Clients SET client_name = $client_name WHERE client_id = $client_id;";
        command.Parameters.AddWithValue("$client_name", clientName);
        command.Parameters.AddWithValue("$client_id", clientId);
        Assert.Equal(1, await command.ExecuteNonQueryAsync());
    }

    private static async Task SetDocumentCarrierAsync(string databasePath, long documentId, string? carrier)
    {
        await using var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = "UPDATE Documents SET carrier = $carrier WHERE document_id = $document_id;";
        command.Parameters.AddWithValue("$carrier", (object?)carrier ?? DBNull.Value);
        command.Parameters.AddWithValue("$document_id", documentId);
        Assert.Equal(1, await command.ExecuteNonQueryAsync());
    }

    private static async Task<string> QuickCheckAsync(string databasePath)
    {
        await using var connection = new SqliteConnection($"Data Source={databasePath};Mode=ReadOnly");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA quick_check;";
        return Convert.ToString(await command.ExecuteScalarAsync()) ?? "";
    }

    private static string CreateTestFolder()
    {
        var folder = Path.Combine(Path.GetTempPath(), "ControlledBookletReprocessingTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        return folder;
    }

    private static void Cleanup(string folder)
    {
        SqliteConnection.ClearAllPools();
        if (Directory.Exists(folder))
        {
            Directory.Delete(folder, recursive: true);
        }
    }
}
