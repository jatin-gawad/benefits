using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Media;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Repositories;

namespace RenewalReviewBuilder.App.Services;

public sealed class PlanDesignWorkspaceDataService
{
    private static readonly (string SearchText, string DisplayName)[] BenefitProviderNames =
    [
        ("Manulife Financial", "Manulife"),
        ("Pacific Blue Cross", "Pacific Blue Cross"),
        ("Equitable Life", "Equitable Life"),
        ("Empire Life", "Empire Life"),
        ("Canada Life", "Canada Life"),
        ("Green Shield", "Green Shield"),
        ("GreenShield", "Green Shield"),
        ("Sun Life", "Sun Life"),
        ("Blue Cross", "Blue Cross"),
        ("Allstate", "Allstate"),
        ("Desjardins", "Desjardins"),
        ("Manulife", "Manulife"),
        ("Equitable", "Equitable Life")
    ];

    private readonly string _databasePath;
    private readonly CurrentVsQuotePlanDesignService _currentVsQuotePlanDesignService;
    private readonly PlanDesignComparisonPresentationService _planDesignComparisonPresentationService = new();

    public PlanDesignWorkspaceDataService(string? databasePath = null)
    {
        _databasePath = string.IsNullOrWhiteSpace(databasePath)
            ? AppPaths.BenefitsIntelligenceDatabasePath
            : databasePath;
        _currentVsQuotePlanDesignService = new CurrentVsQuotePlanDesignService(_databasePath);
    }

    public async Task<PlanDesignWorkspaceViewModel> LoadAsync(
        long? clientId,
        string? clientName,
        string? carrier,
        PlanDesignWorkspaceClientContext? clientContext = null)
    {
        AppPaths.EnsureCreated();

        if (!File.Exists(_databasePath))
        {
            return CreateEmptyModel(null, FirstNonBlank(clientName, "Selected client"), carrier, "Benefits Intelligence database was not found.");
        }

        var databaseService = new DatabaseService(_databasePath);
        var databaseStatus = databaseService.Initialize();
        if (!databaseStatus.Succeeded)
        {
            return CreateEmptyModel(clientId, FirstNonBlank(clientName, "Selected client"), carrier, databaseStatus.Message);
        }

        await using var connection = new SqliteConnection($"Data Source={_databasePath}");
        await connection.OpenAsync();
        var clientOptions = await ReadClientOptionsAsync(connection);

        var client = await ResolveClientAsync(connection, clientId, clientName, carrier)
                     ?? new ClientWorkspaceClient(
                         clientId,
                         FirstNonBlank(clientName, "Selected client"),
                         "",
                         FirstNonBlank(carrier),
                         null,
                         "",
                         "",
                         "",
                         FirstNonBlank(clientContext?.PlanStatus));
        var currentVsQuote = client.ClientId.HasValue
            ? await _currentVsQuotePlanDesignService.BuildAsync(client.ClientId.Value)
            : new CurrentVsQuotePlanDesignComparison();
        var planComparison = _planDesignComparisonPresentationService.Build(currentVsQuote);


        var featureRows = await ReadFeatureRowsAsync(connection, client.ClientId);
        var legacyRows = await ReadLegacyPlanDesignRowsAsync(connection, client.ClientId);
        var flexibleBenefits = client.ClientId.HasValue
            ? await new FlexibleBenefitRepository(_databasePath).GetWorkspaceAsync(client.ClientId.Value)
            : new FlexibleBenefitWorkspaceViewModel();
        var benefitAdministration = client.ClientId.HasValue
            ? await new BenefitAdministrationService(_databasePath).GetWorkspaceAsync(client.ClientId.Value)
            : new BenefitAdministrationWorkspaceViewModel();
        var policyScheduleWarnings = client.ClientId.HasValue
            ? await ReadPolicyScheduleReviewWarningsAsync(connection, client.ClientId.Value)
            : [];
        var verificationRows = client.ClientId.HasValue
            ? await ReadPlanDesignVerificationRowsAsync(connection, client.ClientId.Value)
            : [];
        var clientPolicies = client.ClientId.HasValue
            ? await databaseService.GetClientPoliciesAsync(client.ClientId.Value)
            : [];
        if (featureRows.Count == 0)
        {
            featureRows = legacyRows;
        }
        else if (legacyRows.Count > 0)
        {
            var normalizedSources = featureRows
                .Select(row => FirstNonBlank(row.SourceDocumentPath, row.SourceDocument))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            featureRows.AddRange(legacyRows.Where(row =>
            {
                var source = FirstNonBlank(row.SourceDocumentPath, row.SourceDocument);
                return string.IsNullOrWhiteSpace(source) || !normalizedSources.Contains(source);
            }));
        }

        if (featureRows.Count == 0)
        {
            var emptyModel = CreateEmptyModel(
                client.ClientId,
                client.ClientName,
                FirstNonBlank(client.Carrier, carrier),
                flexibleBenefits.HasPlan || benefitAdministration.HasRecords
                    ? "Administrative or flexible-benefit plan information is loaded. No insured-benefit provisions have been extracted yet."
                    : "No plan design has been extracted for this client yet.",
                flexibleBenefits,
                benefitAdministration,
                planComparison);
            AddClientSelectorOptions(emptyModel, clientOptions);
            AddPolicyScheduleReviewWarnings(emptyModel, policyScheduleWarnings);
            AddVerificationRows(emptyModel, verificationRows);
            return emptyModel;
        }

        var classComparisonData = await ReadClassComparisonDataAsync(connection, client.ClientId);
        var model = BuildModel(client, featureRows, classComparisonData, clientContext, flexibleBenefits, benefitAdministration, planComparison, clientPolicies);
        AddClientSelectorOptions(model, clientOptions);
        AddPolicyScheduleReviewWarnings(model, policyScheduleWarnings);
        AddVerificationRows(model, verificationRows);
        return model;
    }

    private static async Task<List<PlanDesignVerificationWorkspaceRow>> ReadPlanDesignVerificationRowsAsync(SqliteConnection connection, long clientId)
    {
        if (!await TableExistsAsync(connection, "ValidationIssue") || !await TableExistsAsync(connection, "Documents")) return [];
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT v.field_name, v.issue_code, v.message, v.raw_source_reference, d.file_name
            FROM ValidationIssue v JOIN Documents d ON d.document_id = v.source_document_id
            WHERE d.client_id = $client_id AND v.issue_code LIKE 'plan_design_verification_%'
            ORDER BY v.created_at DESC;
            """;
        command.Parameters.AddWithValue("$client_id", clientId);
        var rows = new List<PlanDesignVerificationWorkspaceRow>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync()) rows.Add(new PlanDesignVerificationWorkspaceRow
        {
            Field = ReadString(reader, "field_name"), Status = ReadString(reader, "issue_code").Replace("plan_design_verification_", "", StringComparison.OrdinalIgnoreCase),
            SourceDocument = ReadString(reader, "file_name"), SourceExcerpt = ReadString(reader, "raw_source_reference"), EvidenceValue = ReadString(reader, "message")
        });
        return rows;
    }

    private static void AddVerificationRows(PlanDesignWorkspaceViewModel model, IEnumerable<PlanDesignVerificationWorkspaceRow> rows)
    {
        foreach (var row in rows) model.VerificationRows.Add(row);
    }

    private static async Task<List<PlanDesignFlag>> ReadPolicyScheduleReviewWarningsAsync(
        SqliteConnection connection,
        long clientId)
    {
        if (!await TableExistsAsync(connection, "ValidationIssue") || !await TableExistsAsync(connection, "Documents"))
        {
            return [];
        }

        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT DISTINCT v.message
            FROM ValidationIssue v
            JOIN Documents d ON d.document_id = v.source_document_id
            WHERE d.client_id = $client_id
              AND d.document_type = $document_type
              AND v.issue_code IN ($processor_missing, $extraction_pending, $classification_review)
            ORDER BY v.created_at DESC;
            """;
        command.Parameters.AddWithValue("$client_id", clientId);
        command.Parameters.AddWithValue("$document_type", DocumentTypes.GroupPolicySchedule);
        command.Parameters.AddWithValue("$processor_missing", ValidationIssueCodes.PolicyScheduleProcessorNotRegistered);
        command.Parameters.AddWithValue("$extraction_pending", ValidationIssueCodes.PolicyScheduleExtractionNotImplemented);
        command.Parameters.AddWithValue("$classification_review", ValidationIssueCodes.PolicyScheduleClassificationReviewRequired);
        var warnings = new List<PlanDesignFlag>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            warnings.Add(new PlanDesignFlag
            {
                Label = ReadString(reader, "message"),
                Value = "Review",
                IndicatorBrush = Brushes.DarkOrange
            });
        }

        return warnings;
    }

    private static void AddPolicyScheduleReviewWarnings(
        PlanDesignWorkspaceViewModel model,
        IEnumerable<PlanDesignFlag> warnings)
    {
        foreach (var warning in warnings)
        {
            model.SourceVerificationWarnings.Add(warning);
            model.MissingReviewItems.Add(warning);
        }
    }

    private static PlanDesignWorkspaceViewModel BuildModel(
        ClientWorkspaceClient client,
        IReadOnlyList<PlanFeatureRow> rows,
        ClassComparisonData classComparisonData,
        PlanDesignWorkspaceClientContext? clientContext,
        FlexibleBenefitWorkspaceViewModel flexibleBenefits,
        BenefitAdministrationWorkspaceViewModel benefitAdministration,
        PlanDesignComparisonWorkspaceViewModel planComparison,
        IReadOnlyList<ClientPolicyRecord> clientPolicies)
    {
        var carrier = FirstNonBlank(client.Carrier, rows.Select(row => row.Carrier).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)));
        var sourceDocument = FirstNonBlank(rows.Select(row => row.SourceDocument).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)), "SQLite plan design features");
        var sourceDocumentPath = FirstNonBlank(rows.Select(row => row.SourceDocumentPath).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)));
        var latestLoaded = rows.Select(row => row.DateLoaded).Where(value => value.HasValue).Max();
        var effectiveDate = FirstNonBlank(
            rows.Select(row => row.EffectiveDateText).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)),
            client.EffectiveDate?.ToString("MMM d, yyyy", CultureInfo.InvariantCulture),
            "Not saved");
        var classNames = rows
            .Select(GetClassName)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var concreteClassNames = classNames
            .Where(value => !IsGenericAllClasses(value))
            .ToList();

        if (concreteClassNames.Count > 0)
        {
            classNames = concreteClassNames;
        }

        if (classNames.Count == 0)
        {
            classNames.Add("All Classes");
        }

        var sourceVerificationWarnings = BuildSourceVerificationWarnings(rows);
        var missingItems = CalculateMissingItems(rows, sourceDocumentPath) + sourceVerificationWarnings.Count;

        var classComparison = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            ClientId = client.ClientId,
            ClientName = client.ClientName,
            EffectiveDateText = effectiveDate,
            TotalEligibleEmployees = clientContext?.TotalEligibleEmployees ?? classComparisonData.TotalEligibleEmployees,
            Consultant = FirstNonBlank(clientContext?.Consultant, client.Advisor, "Not available"),
            PlanStatus = FirstNonBlank(clientContext?.PlanStatus, client.Status, "Not available"),
            BenefitRows = rows.Select(ToClassBenefitSourceRow).ToList(),
            RateRows = classComparisonData.Rates,
            CostRows = classComparisonData.Costs,
            Notes = classComparisonData.Notes,
            History = classComparisonData.History,
            Preferences = classComparisonData.Preferences
        });
        classNames = classComparison.Classes.Select(item => item.ClassName).ToList();
        var class1Name = classNames.ElementAtOrDefault(0) ?? "Class 1";
        var class2Name = classNames.ElementAtOrDefault(1) ?? (classNames.Count == 1 ? "" : "Class 2");
        var class3Name = classNames.ElementAtOrDefault(2) ?? "";

        var model = new PlanDesignWorkspaceViewModel
        {
            ClientId = client.ClientId,
            ClientName = client.ClientName,
            PolicyNumber = client.ClientNumber,
            Carrier = carrier,
            SourceDocument = sourceDocument,
            SourceDocumentPath = sourceDocumentPath,
            Quote = "Saved plan design",
            PolicyEffectiveDate = effectiveDate,
            CalcEffectiveDate = effectiveDate,
            RatesVersion = "SQLite",
            ClassesFound = classNames.Count.ToString(CultureInfo.InvariantCulture),
            Class1Name = class1Name,
            Class2Name = class2Name,
            Class3Name = class3Name,
            Readiness = missingItems == 0 ? "Ready" : "Review",
            MissingItemsCount = missingItems,
            HasPlanDesign = true,
            ClassComparison = classComparison,
            PlanComparison = planComparison,
            FlexibleBenefits = flexibleBenefits,
            BenefitAdministration = benefitAdministration,
            EmptyStateMessage = "No plan design has been extracted for this client yet.",
            LastRefreshedText = $"Loaded {rows.Count:#,0} plan design feature row(s) from SQLite for {client.ClientName}."
        };

        foreach (var warning in sourceVerificationWarnings)
        {
            model.SourceVerificationWarnings.Add(warning);
        }

        AddBenefitSummaries(model, rows, latestLoaded, clientPolicies);
        AddSummaryCards(model, rows, effectiveDate, sourceDocument);
        AddPlanSummaryRows(model, rows, classNames);
        AddClassComparisonRows(model, rows, class1Name, class2Name, class3Name);
        AddEhcRows(model, rows, class1Name, class2Name, class3Name);
        AddParamedicalRows(model, rows, class1Name, class2Name, class3Name);
        AddRawExtractionRows(model, rows);
        AddRightRailRows(model, rows, sourceDocumentPath, latestLoaded);

        return model;
    }

    private static PlanDesignWorkspaceViewModel CreateEmptyModel(
        long? clientId,
        string clientName,
        string? carrier,
        string message,
        FlexibleBenefitWorkspaceViewModel? flexibleBenefits = null,
        BenefitAdministrationWorkspaceViewModel? benefitAdministration = null,
        PlanDesignComparisonWorkspaceViewModel? planComparison = null)
    {
        flexibleBenefits ??= new FlexibleBenefitWorkspaceViewModel();
        benefitAdministration ??= new BenefitAdministrationWorkspaceViewModel();
        planComparison ??= new PlanDesignComparisonWorkspaceViewModel();
        var flexibleAdministrator = flexibleBenefits.HasPlan ? flexibleBenefits.Administrator : "";
        var model = new PlanDesignWorkspaceViewModel
        {
            ClientId = clientId,
            ClientName = clientName,
            Carrier = FirstNonBlank(flexibleAdministrator, carrier),
            SourceDocument = flexibleBenefits.HasPlan
                ? FirstNonBlank(Path.GetFileName(flexibleBenefits.SourceDocumentPath), "Flexible-benefits source document")
                : "Not linked",
            SourceDocumentPath = flexibleBenefits.SourceDocumentPath,
            Quote = flexibleBenefits.HasPlan ? "Flexible-benefits plan" : "No extracted plan design",
            PolicyEffectiveDate = flexibleBenefits.HasPlan ? flexibleBenefits.EffectiveDate : "N/A",
            CalcEffectiveDate = flexibleBenefits.HasPlan ? flexibleBenefits.EffectiveDate : "N/A",
            RatesVersion = "N/A",
            ClassesFound = flexibleBenefits.HasPlan ? flexibleBenefits.Classes.Count.ToString(CultureInfo.InvariantCulture) : "0",
            Readiness = flexibleBenefits.HasPlan ? flexibleBenefits.ReviewStatus : "Needs extraction",
            MissingItemsCount = flexibleBenefits.HasPlan ? flexibleBenefits.ReviewItems.Count : 1,
            HasPlanDesign = false,
            PlanComparison = planComparison,
            FlexibleBenefits = flexibleBenefits,
            BenefitAdministration = benefitAdministration,
            EmptyStateMessage = message,
            LastRefreshedText = message
        };

        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Client", Value = clientName, Detail = "Selected workspace client" });
        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Carrier", Value = FirstNonBlank(flexibleAdministrator, carrier, "Not saved"), Detail = "From client profile or flexible-benefits source" });
        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Source Document", Value = model.SourceDocument, Detail = flexibleBenefits.HasPlan ? "Signed flexible-benefits source" : "Upload or link a booklet/GIR" });
        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Effective Date", Value = flexibleBenefits.HasPlan ? flexibleBenefits.EffectiveDate : "N/A", Detail = flexibleBenefits.HasPlan ? "Flexible-benefits plan" : "No extracted plan design" });
        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Classes Found", Value = model.ClassesFound, Detail = flexibleBenefits.HasPlan ? "Flexible-benefits classes" : "No saved class data" });
        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Readiness", Value = model.Readiness, Detail = flexibleBenefits.HasPlan ? "Flexible-benefits review state" : "Missing plan design" });
        AddBenefitSummaries(model, [], null);
        return model;
    }

    private static void AddBenefitSummaries(
        PlanDesignWorkspaceViewModel model,
        IReadOnlyList<PlanFeatureRow> rows,
        DateTime? latestLoaded,
        IReadOnlyList<ClientPolicyRecord>? clientPolicies = null)
    {
        var sourceRows = rows.Select(row => new BenefitSummarySourceRow
        {
            ClassName = GetClassName(row),
            Benefit = row.Benefit,
            Category = row.Category,
            ProvisionName = row.FeatureName,
            RawValue = row.FeatureValue,
            Coinsurance = row.Coinsurance,
            Deductible = row.Deductible,
            Maximum = row.Maximum,
            Frequency = row.Frequency,
            WaitingPeriod = row.WaitingPeriod,
            TerminationAge = row.TerminationAge,
            Carrier = row.Carrier,
            SourceText = row.SourceText,
            SourceDocument = row.SourceDocument,
            SourcePage = ParseSourcePage(row.SourcePage),
            PlanDesignFeatureId = row.PlanDesignFeatureId,
            Rules = row.Rules
        }).ToList();

        if (model.FlexibleBenefits.HasPlan)
        {
            if (IsConfirmedFlexibleBenefitSummary(model.FlexibleBenefits.HcsaSummary))
            {
                sourceRows.Add(new BenefitSummarySourceRow
                {
                    ClassName = "All Classes",
                    Benefit = "Health Spending Account",
                    Category = "Flexible Benefits",
                    ProvisionName = "Benefit Amount",
                    RawValue = model.FlexibleBenefits.HcsaSummary,
                    Carrier = model.FlexibleBenefits.Administrator,
                    SourceDocument = Path.GetFileName(model.FlexibleBenefits.SourceDocumentPath),
                    SourceText = $"Administrator: {model.FlexibleBenefits.Administrator}"
                });
            }

            if (IsConfirmedFlexibleBenefitSummary(model.FlexibleBenefits.LsaSummary))
            {
                sourceRows.Add(new BenefitSummarySourceRow
                {
                    ClassName = "All Classes",
                    Benefit = "Wellness Spending Account",
                    Category = "Flexible Benefits",
                    ProvisionName = "Benefit Amount",
                    RawValue = model.FlexibleBenefits.LsaSummary,
                    Carrier = model.FlexibleBenefits.Administrator,
                    SourceDocument = Path.GetFileName(model.FlexibleBenefits.SourceDocumentPath),
                    SourceText = $"Administrator: {model.FlexibleBenefits.Administrator}"
                });
            }
        }

        var carrierPolicies = new List<BenefitSummaryCarrierPolicy>();
        carrierPolicies.AddRange((clientPolicies ?? [])
            .Where(policy =>
                policy.VerificationStatus.Equals(ClientPolicyVerificationStatuses.Verified, StringComparison.OrdinalIgnoreCase) &&
                !string.IsNullOrWhiteSpace(policy.PolicyNumber))
            .Select(policy => new BenefitSummaryCarrierPolicy
            {
                Carrier = FirstNonBlank(policy.Carrier, model.Carrier),
                PolicyNumber = policy.PolicyNumber
            }));
        foreach (var assignment in model.BenefitAdministration.Rows)
        {
            if (!string.IsNullOrWhiteSpace(assignment.InsuranceCompany) &&
                !assignment.InsuranceCompany.Equals("Not assigned", StringComparison.OrdinalIgnoreCase))
            {
                carrierPolicies.Add(new BenefitSummaryCarrierPolicy
                {
                    Carrier = assignment.InsuranceCompany,
                    PolicyNumber = assignment.PolicyNumber,
                    Coverage = assignment.Benefit
                });
            }

            foreach (var provider in assignment.ServiceProvider
                         .Split([';', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (!provider.Equals("Not assigned", StringComparison.OrdinalIgnoreCase))
                {
                    carrierPolicies.Add(new BenefitSummaryCarrierPolicy
                    {
                        Carrier = provider,
                        Coverage = assignment.Benefit
                    });
                }
            }
        }

        var summaries = BenefitSummaryDisplayFormatter.Build(new BenefitSummaryBuildRequest
        {
            ClientName = model.ClientName,
            Carrier = model.Carrier,
            PolicyNumber = model.PolicyNumber,
            EffectiveDate = model.EffectiveDateText,
            UpdatedDate = latestLoaded,
            UpdatedDateSource = latestLoaded.HasValue ? "Latest saved plan-design feature" : "",
            Rows = sourceRows,
            CarrierPolicies = carrierPolicies
        });

        foreach (var summary in summaries.Where(summary => summary.Sections.Count > 0))
        {
            model.BenefitSummaries.Add(summary);
        }

        model.SelectedBenefitSummary = model.BenefitSummaries.FirstOrDefault();
    }

    private static bool IsConfirmedFlexibleBenefitSummary(string value)
    {
        return !string.IsNullOrWhiteSpace(value) &&
               !value.Equals("Not confirmed", StringComparison.OrdinalIgnoreCase);
    }

    private static int? ParseSourcePage(string value)
    {
        var match = Regex.Match(value ?? "", @"\d+");
        return match.Success && int.TryParse(match.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var page)
            ? page
            : null;
    }

    private static void AddClientSelectorOptions(PlanDesignWorkspaceViewModel model, IReadOnlyList<ClientWorkspaceClient> clients)
    {
        model.ClientSelectorOptions.Clear();
        model.ClientSelectorOptions.Add(new PlanDesignClientOption { ClientName = "Select a client" });

        foreach (var client in clients
                     .Where(client => client.ClientId.HasValue && !string.IsNullOrWhiteSpace(client.ClientName))
                     .GroupBy(client => client.ClientId!.Value)
                     .Select(group => group.First())
                     .OrderBy(client => client.ClientName, StringComparer.OrdinalIgnoreCase))
        {
            model.ClientSelectorOptions.Add(new PlanDesignClientOption
            {
                ClientId = client.ClientId,
                ClientName = client.ClientName,
                Carrier = client.Carrier
            });
        }

        if (model.ClientId.HasValue && model.ClientSelectorOptions.All(option => option.ClientId != model.ClientId))
        {
            model.ClientSelectorOptions.Add(new PlanDesignClientOption
            {
                ClientId = model.ClientId,
                ClientName = model.ClientName,
                Carrier = model.Carrier
            });
        }
    }

    private static void AddSummaryCards(PlanDesignWorkspaceViewModel model, IReadOnlyList<PlanFeatureRow> rows, string effectiveDate, string sourceDocument)
    {
        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Client", Value = model.ClientName, Detail = string.IsNullOrWhiteSpace(model.PolicyNumber) ? "Saved client profile" : $"Policy #: {model.PolicyNumber}" });
        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Carrier", Value = FirstNonBlank(model.Carrier, "Not saved"), Detail = $"Rates Version: {model.RatesVersion}" });
        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Source Document", Value = sourceDocument, Detail = "Saved extraction source" });
        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Effective Date", Value = effectiveDate, Detail = "Client/profile source" });
        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Classes Found", Value = model.ClassesFound, Detail = ClassDetail(model) });
        model.SummaryCards.Add(new PlanDesignSummaryCard { Label = "Readiness", Value = model.Readiness, Detail = $"Missing items: {model.MissingItemsCount}" });
    }

    private static string ClassDetail(PlanDesignWorkspaceViewModel model)
    {
        var classNames = new[] { model.Class1Name, model.Class2Name, model.Class3Name }
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToList();

        return classNames.Count == 0
            ? "All Classes"
            : string.Join(" / ", classNames);
    }

    private static void AddPlanSummaryRows(
        PlanDesignWorkspaceViewModel model,
        IReadOnlyList<PlanFeatureRow> rows,
        IReadOnlyList<string> classNames)
    {
        foreach (var group in rows
                     .GroupBy(row => NormalizeBenefitArea(row), StringComparer.OrdinalIgnoreCase)
                     .OrderBy(group => BenefitOrder(group.Key))
                     .ThenBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            var classValues = classNames
                .Select(className => new PlanDesignClassSummaryValue
                {
                    ClassName = className,
                    Value = SummarizeCompactComparisonRows(group, className, group.Key)
                })
                .ToList();
            var comparison = classValues.Count > 1
                ? CompareClassValues(classValues.Select(item => item.Value).ToArray())
                : (true, "Single class", Brushes.SeaGreen);
            model.SummaryRows.Add(new PlanDesignBenefitSummaryRow
            {
                BenefitArea = group.Key,
                Class1Value = classValues.ElementAtOrDefault(0)?.Value ?? "",
                Class2Value = classValues.ElementAtOrDefault(1)?.Value ?? "",
                Class3Value = classValues.ElementAtOrDefault(2)?.Value ?? "",
                ClassValues = classValues,
                KeyDifference = comparison.Text,
                DifferenceBrush = comparison.Brush
            });
        }
    }

    private static void AddClassComparisonRows(PlanDesignWorkspaceViewModel model, IReadOnlyList<PlanFeatureRow> rows, string class1Name, string class2Name, string class3Name)
    {
        foreach (var group in rows
                     .GroupBy(row => NormalizeBenefitArea(row), StringComparer.OrdinalIgnoreCase)
                     .OrderBy(group => BenefitOrder(group.Key))
                     .ThenBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            var hasSecondClass = !string.IsNullOrWhiteSpace(class2Name);
            var hasThirdClass = !string.IsNullOrWhiteSpace(class3Name);
            var class1Value = SummarizeClassRows(group, class1Name, group.Key);
            var class2Value = hasSecondClass ? SummarizeClassRows(group, class2Name, group.Key) : "";
            var class3Value = hasThirdClass ? SummarizeClassRows(group, class3Name, group.Key) : "";
            var comparison = hasSecondClass || hasThirdClass
                ? CompareClassValues(class1Value, class2Value, class3Value)
                : (true, "Single class", Brushes.SeaGreen);
            model.ClassComparisonRows.Add(new PlanDesignClassComparisonRow
            {
                BenefitArea = group.Key,
                Class1Value = class1Value,
                Class2Value = class2Value,
                Class3Value = class3Value,
                Difference = comparison.IsSame ? "Same" : "Material",
                PricingImpact = comparison.IsSame ? "-" : "Review pricing impact",
                DifferenceBrush = comparison.Brush
            });
        }
    }

    private static void AddEhcRows(PlanDesignWorkspaceViewModel model, IReadOnlyList<PlanFeatureRow> rows, string class1Name, string class2Name, string class3Name)
    {
        var ehcRows = rows
            .Where(row => ContainsAny(row.Benefit, row.Category, "ehc", "extended health", "health care", "healthcare"))
            .ToList();

        if (ehcRows.Count == 0)
        {
            model.EhcCoverageRows.Add(new PlanDesignCoverageRow
            {
                CoverageFeature = "Extended Healthcare",
                Class1Value = "No EHC extraction found",
                Class2Value = string.IsNullOrWhiteSpace(class2Name) ? "" : "No EHC extraction found",
                Class3Value = string.IsNullOrWhiteSpace(class3Name) ? "" : "No EHC extraction found",
                Difference = "Review",
                DifferenceBrush = Brushes.DarkOrange
            });
            return;
        }

        foreach (var group in ehcRows
                     .GroupBy(row => FirstNonBlank(row.FeatureName, row.Category, row.Benefit), StringComparer.OrdinalIgnoreCase)
                     .OrderBy(group => FeatureOrder(group.Key))
                     .ThenBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            var hasSecondClass = !string.IsNullOrWhiteSpace(class2Name);
            var hasThirdClass = !string.IsNullOrWhiteSpace(class3Name);
            var class1Value = SummarizeFeatureRows(group, class1Name);
            var class2Value = hasSecondClass ? SummarizeFeatureRows(group, class2Name) : "";
            var class3Value = hasThirdClass ? SummarizeFeatureRows(group, class3Name) : "";
            var comparison = hasSecondClass || hasThirdClass
                ? CompareClassValues(class1Value, class2Value, class3Value)
                : (true, "Single class", Brushes.SeaGreen);
            model.EhcCoverageRows.Add(new PlanDesignCoverageRow
            {
                CoverageFeature = group.Key,
                Class1Value = class1Value,
                Class2Value = class2Value,
                Class3Value = class3Value,
                Difference = comparison.Text,
                DifferenceBrush = comparison.Brush
            });
        }
    }

    private static void AddParamedicalRows(PlanDesignWorkspaceViewModel model, IReadOnlyList<PlanFeatureRow> rows, string class1Name, string class2Name, string class3Name)
    {
        var paramedicalRows = rows
            .Where(row => ContainsAny(
                $"{row.Benefit} {row.Category} {row.FeatureName}",
                $"{row.FeatureValue} {row.SourceText}",
                "paramedical",
                "practitioner"))
            .ToList();

        if (paramedicalRows.Count == 0)
        {
            model.ParamedicalRows.Add(new PlanDesignCoverageRow
            {
                CoverageFeature = "Paramedical Services",
                Class1Value = "No paramedical extraction found",
                Class2Value = string.IsNullOrWhiteSpace(class2Name) ? "" : "No paramedical extraction found",
                Class3Value = string.IsNullOrWhiteSpace(class3Name) ? "" : "No paramedical extraction found",
                Difference = "Review",
                DifferenceBrush = Brushes.DarkOrange
            });
            model.ParamedicalDetails.Add(new PlanDesignDetailValue
            {
                Label = "Covered Practitioners",
                Value = "No practitioner list was saved for this client."
            });
            return;
        }

        AddParamedicalComparisonRow(model, paramedicalRows, class1Name, class2Name, class3Name, "Coinsurance", row =>
            FirstNonBlank(row.Coinsurance, ExtractPercent(row.FeatureValue)));
        AddParamedicalComparisonRow(model, paramedicalRows, class1Name, class2Name, class3Name, "Maximums", row =>
            FirstNonBlank(row.Maximum, ExtractMaximumText(row.FeatureValue)));
        AddParamedicalComparisonRow(model, paramedicalRows, class1Name, class2Name, class3Name, "Practitioner Limit", row =>
            ExtractPractitionerLimit(row.FeatureValue));

        var practitionerList = paramedicalRows
            .Select(row => ExtractPractitionerList(FirstNonBlank(row.FeatureValue, row.SourceText)))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(value => value.Length)
            .FirstOrDefault();

        model.ParamedicalDetails.Add(new PlanDesignDetailValue
        {
            Label = "Covered Practitioners",
            Value = FirstNonBlank(practitionerList, "Practitioner list is not separately stated in the saved source.")
        });
    }

    private static void AddParamedicalComparisonRow(
        PlanDesignWorkspaceViewModel model,
        IReadOnlyList<PlanFeatureRow> rows,
        string class1Name,
        string class2Name,
        string class3Name,
        string label,
        Func<PlanFeatureRow, string> valueSelector)
    {
        var hasSecondClass = !string.IsNullOrWhiteSpace(class2Name);
        var hasThirdClass = !string.IsNullOrWhiteSpace(class3Name);
        var class1Value = SelectParamedicalValue(rows, class1Name, valueSelector);
        var class2Value = hasSecondClass ? SelectParamedicalValue(rows, class2Name, valueSelector) : "";
        var class3Value = hasThirdClass ? SelectParamedicalValue(rows, class3Name, valueSelector) : "";

        if (new[] { class1Value, class2Value, class3Value }
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .All(value => value.Equals("N/A", StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        var comparison = hasSecondClass || hasThirdClass
            ? CompareClassValues(class1Value, class2Value, class3Value)
            : (true, "Single class", Brushes.SeaGreen);
        model.ParamedicalRows.Add(new PlanDesignCoverageRow
        {
            CoverageFeature = label,
            Class1Value = class1Value,
            Class2Value = class2Value,
            Class3Value = class3Value,
            Difference = comparison.Text,
            DifferenceBrush = comparison.Brush
        });
    }

    private static string SelectParamedicalValue(
        IEnumerable<PlanFeatureRow> rows,
        string className,
        Func<PlanFeatureRow, string> valueSelector)
    {
        return FirstNonBlank(
            rows.Where(row => ClassMatches(row, className))
                .Select(valueSelector)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)),
            "N/A");
    }

    private static string ExtractPercent(string value)
    {
        var match = Regex.Match(value ?? "", @"\b(100|[1-9]?\d(?:\.\d+)?)\s*%");
        return match.Success ? $"{match.Groups[1].Value}%" : "";
    }

    private static string ExtractMaximumText(string value)
    {
        var text = FirstNonBlank(value);
        var practitionerIndex = text.IndexOf("Practitioners:", StringComparison.OrdinalIgnoreCase);
        if (practitionerIndex >= 0)
        {
            text = text[..practitionerIndex];
        }

        var amounts = Regex.Matches(text, @"\$[\d,]+(?:\.\d{2})?(?:\s+[^.;]+)?")
            .Select(match => match.Value.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        return string.Join("; ", amounts);
    }

    private static string ExtractPractitionerLimit(string value)
    {
        var text = FirstNonBlank(value);
        if (text.Contains("no practitioner limit", StringComparison.OrdinalIgnoreCase))
        {
            return "No practitioner limit";
        }

        var match = Regex.Match(text, @"\$[\d,]+(?:\.\d{2})?\s+(?:per|each)\s+practitioner", RegexOptions.IgnoreCase);
        return match.Success ? match.Value : "";
    }

    private static string ExtractPractitionerList(string value)
    {
        var text = FirstNonBlank(value);
        var markerIndex = text.IndexOf("Practitioners:", StringComparison.OrdinalIgnoreCase);
        return markerIndex < 0
            ? ""
            : text[(markerIndex + "Practitioners:".Length)..].Trim().TrimEnd('.');
    }

    private static void AddRawExtractionRows(PlanDesignWorkspaceViewModel model, IReadOnlyList<PlanFeatureRow> rows)
    {
        foreach (var row in rows
                     .OrderBy(row => BenefitOrder(NormalizeBenefitArea(row)))
                     .ThenBy(row => NormalizeBenefitArea(row), StringComparer.OrdinalIgnoreCase)
                     .ThenBy(row => row.Category, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(row => GetClassName(row), StringComparer.OrdinalIgnoreCase)
                     .ThenBy(row => row.SourcePage, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(row => row.FeatureName, StringComparer.OrdinalIgnoreCase))
        {
            model.RawExtractionRows.Add(new PlanDesignRawExtractionRow
            {
                Benefit = FirstNonBlank(row.Benefit, NormalizeBenefitArea(row)),
                Category = row.Category,
                FeatureName = row.FeatureName,
                FeatureValue = row.FeatureValue,
                Coinsurance = row.Coinsurance,
                Deductible = row.Deductible,
                Maximum = row.Maximum,
                Frequency = row.Frequency,
                TerminationAge = row.TerminationAge,
                ClassText = GetClassName(row),
                SourcePage = row.SourcePage,
                ConfidenceText = row.ConfidenceScore.HasValue
                    ? row.ConfidenceScore.Value.ToString("0.##", CultureInfo.InvariantCulture)
                    : "",
                SourceText = row.SourceText,
                LoadedText = row.DateLoaded.HasValue
                    ? row.DateLoaded.Value.ToString("MMM d, yyyy h:mm tt", CultureInfo.InvariantCulture)
                    : ""
            });
        }
    }

    private static void AddRightRailRows(PlanDesignWorkspaceViewModel model, IReadOnlyList<PlanFeatureRow> rows, string sourceDocumentPath, DateTime? latestLoaded)
    {
        var classCount = rows.Select(GetClassName).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase).Count();
        var coinsuranceCount = rows.Count(row => !string.IsNullOrWhiteSpace(row.Coinsurance) || ContainsAny(row.FeatureName, row.FeatureValue, "coinsurance", "%"));
        var maximumCount = rows.Count(row => !string.IsNullOrWhiteSpace(row.Maximum) || ContainsAny(row.FeatureName, row.FeatureValue, "maximum", "max"));
        var deductibleCount = rows.Count(row => !string.IsNullOrWhiteSpace(row.Deductible) || ContainsAny(row.FeatureName, row.FeatureValue, "deductible"));
        var lowConfidenceCount = rows.Count(row => row.ConfidenceScore.HasValue && row.ConfidenceScore.Value < 0.75m);

        model.PricingImpactFlags.Add(new PlanDesignFlag { Label = "Classes detected", Value = classCount.ToString(CultureInfo.InvariantCulture), IndicatorBrush = classCount > 1 ? Brushes.DarkOrange : Brushes.SeaGreen });
        model.PricingImpactFlags.Add(new PlanDesignFlag { Label = "Coinsurance items", Value = coinsuranceCount.ToString(CultureInfo.InvariantCulture), IndicatorBrush = coinsuranceCount > 0 ? Brushes.DarkOrange : Brushes.Gray });
        model.PricingImpactFlags.Add(new PlanDesignFlag { Label = "Maximum / limit items", Value = maximumCount.ToString(CultureInfo.InvariantCulture), IndicatorBrush = maximumCount > 0 ? Brushes.DarkOrange : Brushes.Gray });
        model.PricingImpactFlags.Add(new PlanDesignFlag { Label = "Deductible items", Value = deductibleCount.ToString(CultureInfo.InvariantCulture), IndicatorBrush = deductibleCount > 0 ? Brushes.SeaGreen : Brushes.Gray });

        if (string.IsNullOrWhiteSpace(sourceDocumentPath))
        {
            model.MissingReviewItems.Add(new PlanDesignFlag { Label = "Source document path is not linked", Value = "Review", IndicatorBrush = Brushes.DarkOrange });
        }

        if (lowConfidenceCount > 0)
        {
            model.MissingReviewItems.Add(new PlanDesignFlag { Label = "Low confidence extraction rows", Value = lowConfidenceCount.ToString(CultureInfo.InvariantCulture), IndicatorBrush = Brushes.DarkOrange });
        }

        if (model.EhcCoverageRows.Count == 0 || model.EhcCoverageRows.Any(row => row.Class1Value.Contains("No EHC extraction", StringComparison.OrdinalIgnoreCase)))
        {
            model.MissingReviewItems.Add(new PlanDesignFlag { Label = "EHC details need review", Value = "Review", IndicatorBrush = Brushes.DarkOrange });
        }

        foreach (var warning in model.SourceVerificationWarnings)
        {
            model.MissingReviewItems.Add(warning);
        }

        if (model.MissingReviewItems.Count == 0)
        {
            model.MissingReviewItems.Add(new PlanDesignFlag { Label = "No required review items detected", Value = "", IndicatorBrush = Brushes.SeaGreen });
        }

        foreach (var row in model.ClassComparisonRows.Where(row => !row.Difference.Equals("Same", StringComparison.OrdinalIgnoreCase)).Take(5))
        {
            model.KeyHighlights.Add(new PlanDesignFlag { Label = $"{row.BenefitArea} differs by class", Value = "Review", IndicatorBrush = Brushes.DarkOrange });
            model.PricingDrivers.Add(new PlanDesignFlag { Label = row.BenefitArea, Value = "Review", Detail = row.PricingImpact, IndicatorBrush = Brushes.DarkOrange });
        }

        if (model.KeyHighlights.Count == 0)
        {
            model.KeyHighlights.Add(new PlanDesignFlag { Label = "No material class differences detected", Value = "", IndicatorBrush = Brushes.SeaGreen });
            model.PricingDrivers.Add(new PlanDesignFlag { Label = "No obvious pricing drivers from class differences", Value = "", IndicatorBrush = Brushes.SeaGreen });
        }

        model.UnderwritingConsiderations.Add(new PlanDesignFlag { Label = "Review benefit differences against claims and demographics.", IndicatorBrush = Brushes.SeaGreen });
        model.UnderwritingConsiderations.Add(new PlanDesignFlag { Label = "Validate any rows marked for review before pricing reliance.", IndicatorBrush = Brushes.DarkOrange });
        model.ConsultantTalkingPoints.Add(new PlanDesignFlag { Label = "Use this client-specific extraction when discussing plan design.", IndicatorBrush = Brushes.SeaGreen });
        model.ConsultantTalkingPoints.Add(new PlanDesignFlag { Label = "Confirm source document wording for high-impact provisions.", IndicatorBrush = Brushes.DarkOrange });

        model.ReadinessChecklist.Add(new PlanDesignFlag { Label = "Client-specific plan design data loaded", Value = "Done", IndicatorBrush = Brushes.SeaGreen });
        model.ReadinessChecklist.Add(new PlanDesignFlag { Label = "Source document linked", Value = string.IsNullOrWhiteSpace(sourceDocumentPath) ? "Review" : "Done", IndicatorBrush = string.IsNullOrWhiteSpace(sourceDocumentPath) ? Brushes.DarkOrange : Brushes.SeaGreen });
        foreach (var warning in model.SourceVerificationWarnings)
        {
            model.ReadinessChecklist.Add(warning);
        }
        model.ReadinessChecklist.Add(new PlanDesignFlag { Label = latestLoaded.HasValue ? $"Last loaded {latestLoaded.Value:MMM d, yyyy}" : "Loaded date unavailable", Value = "", IndicatorBrush = Brushes.SeaGreen });

        model.EhcAdditionalDetails.Add(new PlanDesignDetailValue { Label = "Source", Value = model.SourceDocument });
        model.EhcAdditionalDetails.Add(new PlanDesignDetailValue { Label = "Carrier", Value = FirstNonBlank(model.Carrier, "N/A") });
        model.EhcAdditionalDetails.Add(new PlanDesignDetailValue { Label = "Classes", Value = model.ClassesFound });
        model.EhcAdditionalDetails.Add(new PlanDesignDetailValue { Label = "Feature rows", Value = rows.Count.ToString("#,0", CultureInfo.InvariantCulture) });
        model.EhcAdditionalDetails.Add(new PlanDesignDetailValue { Label = "Readiness", Value = model.Readiness });
    }

    public async Task<DatabaseOperationResult> SaveClassComparisonPreferencesAsync(
        ClassComparisonWorkspaceViewModel model,
        string updatedBy)
    {
        if (!model.ClientId.HasValue)
        {
            return DatabaseOperationResult.Failure("Select a saved client before saving the class comparison.", _databasePath);
        }

        try
        {
            var databaseStatus = new DatabaseService(_databasePath).Initialize();
            if (!databaseStatus.Succeeded)
            {
                return databaseStatus;
            }

            await using var connection = new SqliteConnection($"Data Source={_databasePath}");
            await connection.OpenAsync();
            var previous = await ReadClassComparisonPreferencesAsync(connection, model.ClientId);
            var state = model.ToPreferenceState(updatedBy);
            using var transaction = connection.BeginTransaction();
            await using (var command = connection.CreateCommand())
            {
                command.Transaction = transaction;
                command.CommandText = """
                    INSERT INTO class_comparison_preferences (
                        client_id,
                        baseline_class_name,
                        selected_class_names,
                        active_tab,
                        key_benefits_only,
                        highlight_differences,
                        collapsed_sections,
                        recommended_class_name,
                        updated_at,
                        updated_by
                    )
                    VALUES (
                        $client_id,
                        $baseline_class_name,
                        $selected_class_names,
                        $active_tab,
                        $key_benefits_only,
                        $highlight_differences,
                        $collapsed_sections,
                        $recommended_class_name,
                        $updated_at,
                        $updated_by
                    )
                    ON CONFLICT(client_id) DO UPDATE SET
                        baseline_class_name = excluded.baseline_class_name,
                        selected_class_names = excluded.selected_class_names,
                        active_tab = excluded.active_tab,
                        key_benefits_only = excluded.key_benefits_only,
                        highlight_differences = excluded.highlight_differences,
                        collapsed_sections = excluded.collapsed_sections,
                        recommended_class_name = excluded.recommended_class_name,
                        updated_at = excluded.updated_at,
                        updated_by = excluded.updated_by;
                    """;
                command.Parameters.AddWithValue("$client_id", model.ClientId.Value);
                command.Parameters.AddWithValue("$baseline_class_name", state.BaselineClassName);
                command.Parameters.AddWithValue("$selected_class_names", JsonSerializer.Serialize(state.SelectedClassNames));
                command.Parameters.AddWithValue("$active_tab", state.ActiveTab);
                command.Parameters.AddWithValue("$key_benefits_only", state.KeyBenefitsOnly ? 1 : 0);
                command.Parameters.AddWithValue("$highlight_differences", state.HighlightDifferences ? 1 : 0);
                command.Parameters.AddWithValue("$collapsed_sections", JsonSerializer.Serialize(state.CollapsedSections));
                command.Parameters.AddWithValue("$recommended_class_name", state.RecommendedClassName);
                command.Parameters.AddWithValue("$updated_at", DateTime.Now.ToString("O", CultureInfo.InvariantCulture));
                command.Parameters.AddWithValue("$updated_by", FirstNonBlank(updatedBy, "Desktop user"));
                await command.ExecuteNonQueryAsync();
            }

            await AddPreferenceAuditRowsAsync(connection, transaction, model.ClientId.Value, previous, state);
            transaction.Commit();
            return DatabaseOperationResult.Success("Class comparison selections were saved.", _databasePath);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseOperationResult.Failure($"Class comparison selections could not be saved: {ex.Message}", _databasePath);
        }
    }

    public async Task<ClassComparisonNoteRow> SaveClassComparisonNoteAsync(
        long clientId,
        string noteText,
        string className,
        string author)
    {
        var cleanedNote = FirstNonBlank(noteText);
        if (string.IsNullOrWhiteSpace(cleanedNote))
        {
            throw new InvalidOperationException("Enter a note before saving.");
        }

        var databaseStatus = new DatabaseService(_databasePath).Initialize();
        if (!databaseStatus.Succeeded)
        {
            throw new InvalidOperationException(databaseStatus.Message);
        }

        var createdAt = DateTime.Now;
        var createdBy = FirstNonBlank(author, "Desktop user");
        await using var connection = new SqliteConnection($"Data Source={_databasePath}");
        await connection.OpenAsync();
        using var transaction = connection.BeginTransaction();
        long noteId;
        await using (var command = connection.CreateCommand())
        {
            command.Transaction = transaction;
            command.CommandText = """
                INSERT INTO class_comparison_notes (
                    client_id,
                    class_name,
                    note_text,
                    created_at,
                    created_by,
                    updated_at,
                    updated_by
                )
                VALUES (
                    $client_id,
                    $class_name,
                    $note_text,
                    $created_at,
                    $created_by,
                    $created_at,
                    $created_by
                );
                SELECT last_insert_rowid();
                """;
            command.Parameters.AddWithValue("$client_id", clientId);
            command.Parameters.AddWithValue("$class_name", FirstNonBlank(className));
            command.Parameters.AddWithValue("$note_text", cleanedNote);
            command.Parameters.AddWithValue("$created_at", createdAt.ToString("O", CultureInfo.InvariantCulture));
            command.Parameters.AddWithValue("$created_by", createdBy);
            noteId = Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
        }

        await InsertClassComparisonAuditAsync(
            connection,
            transaction,
            clientId,
            className,
            "Analyst Note",
            "",
            cleanedNote,
            "Analyst note added in Class Comparison.",
            "",
            createdBy,
            createdAt);
        transaction.Commit();

        return new ClassComparisonNoteRow
        {
            NoteId = noteId,
            ClientId = clientId,
            NoteText = cleanedNote,
            ClassName = FirstNonBlank(className),
            Author = createdBy,
            CreatedAt = createdAt
        };
    }

    private static ClassBenefitSourceRow ToClassBenefitSourceRow(PlanFeatureRow row)
    {
        var tieredReimbursement = GetTieredReimbursementDisplay(row, compact: true);
        return new ClassBenefitSourceRow(
            GetClassName(row),
            row.Benefit,
            row.Category,
            row.FeatureName,
            FirstNonBlank(tieredReimbursement, row.FeatureValue),
            string.IsNullOrWhiteSpace(tieredReimbursement) ? row.Coinsurance : "",
            row.Deductible,
            row.Maximum,
            row.Frequency,
            row.WaitingPeriod,
            row.TerminationAge,
            row.SourceSection,
            row.SourceDocument,
            row.SourcePage,
            row.SourceText,
            row.SourceText,
            row.ConfidenceScore);
    }

    private async Task<ClassComparisonData> ReadClassComparisonDataAsync(SqliteConnection connection, long? clientId)
    {
        if (!clientId.HasValue)
        {
            return ClassComparisonData.Empty;
        }

        var rates = await ReadClassRatesAsync(connection, clientId.Value);
        var costs = await ReadClassCostsAsync(connection, clientId.Value);
        var totalEligibleEmployees = await ReadTotalEligibleEmployeesAsync(connection, clientId.Value, costs);
        var preferences = await ReadClassComparisonPreferencesAsync(connection, clientId.Value);
        var notes = await ReadClassComparisonNotesAsync(connection, clientId.Value);
        var history = await ReadClassComparisonHistoryAsync(connection, clientId.Value);
        return new ClassComparisonData(rates, costs, totalEligibleEmployees, preferences, notes, history);
    }

    private static async Task<List<ClassRateSourceRow>> ReadClassRatesAsync(SqliteConnection connection, long clientId)
    {
        if (!await TableExistsAsync(connection, "Rates"))
        {
            return [];
        }

        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                r.benefit_line,
                r.class AS class_name,
                r.current_rate,
                r.proposed_rate,
                r.negotiated_rate,
                r.rate_basis,
                r.effective_date,
                COALESCE(NULLIF(d.file_name, ''), NULLIF(r.source_file_name, ''), 'Saved rate data') AS source_document
            FROM Rates r
            LEFT JOIN Documents d ON d.document_id = r.source_document_id
            WHERE r.client_id = $client_id
            ORDER BY r.effective_date DESC, r.benefit_line, r.class;
            """;
        command.Parameters.AddWithValue("$client_id", clientId);

        var rows = new List<ClassRateSourceRow>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rows.Add(new ClassRateSourceRow(
                ReadString(reader, "benefit_line"),
                ReadString(reader, "class_name"),
                ReadDecimal(reader, "current_rate"),
                ReadDecimal(reader, "proposed_rate"),
                ReadDecimal(reader, "negotiated_rate"),
                ReadString(reader, "rate_basis"),
                ReadString(reader, "effective_date"),
                ReadString(reader, "source_document")));
        }

        return rows;
    }

    private static async Task<List<ClassCostSourceRow>> ReadClassCostsAsync(SqliteConnection connection, long clientId)
    {
        var rows = new List<ClassCostSourceRow>();
        if (await TableExistsAsync(connection, "BillingStatements") && await TableExistsAsync(connection, "BillingStatementRows"))
        {
            await using var command = connection.CreateCommand();
            command.CommandText = """
                WITH latest_statement AS (
                    SELECT billing_statement_id, source_file_name
                    FROM BillingStatements
                    WHERE client_id = $client_id
                      AND is_deleted = 0
                      AND is_replaced = 0
                    ORDER BY statement_year DESC, statement_month DESC, import_version DESC, billing_statement_id DESC
                    LIMIT 1
                )
                SELECT
                    COALESCE(NULLIF(TRIM(r.class_code), ''), NULLIF(TRIM(r.employment_class), ''), '') AS class_name,
                    COUNT(DISTINCT COALESCE(NULLIF(TRIM(r.employee_id), ''), NULLIF(TRIM(r.certificate_id), ''), NULLIF(TRIM(r.employee_name), ''))) AS employee_count,
                    SUM(r.employer_premium) AS employer_premium,
                    SUM(r.employee_premium) AS employee_premium,
                    SUM(COALESCE(r.total_premium, r.row_total)) AS total_premium,
                    MAX(s.source_file_name) AS source_document
                FROM BillingStatementRows r
                JOIN latest_statement s ON s.billing_statement_id = r.billing_statement_id
                WHERE r.client_id = $client_id
                GROUP BY COALESCE(NULLIF(TRIM(r.class_code), ''), NULLIF(TRIM(r.employment_class), ''), '')
                HAVING TRIM(class_name) <> '';
                """;
            command.Parameters.AddWithValue("$client_id", clientId);
            await using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                rows.Add(new ClassCostSourceRow(
                    ReadString(reader, "class_name"),
                    ReadNullableInt(reader, "employee_count"),
                    ReadDecimal(reader, "employer_premium"),
                    ReadDecimal(reader, "employee_premium"),
                    ReadDecimal(reader, "total_premium"),
                    ReadString(reader, "source_document")));
            }
        }

        if (rows.Count == 0 && await TableExistsAsync(connection, "BillingStatementTotals"))
        {
            await using var command = connection.CreateCommand();
            command.CommandText = """
                SELECT
                    class_code AS class_name,
                    MAX(lives) AS employee_count,
                    SUM(COALESCE(total_premium, premium)) AS total_premium
                FROM BillingStatementTotals
                WHERE client_id = $client_id
                  AND TRIM(COALESCE(class_code, '')) <> ''
                GROUP BY class_code;
                """;
            command.Parameters.AddWithValue("$client_id", clientId);
            await using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                rows.Add(new ClassCostSourceRow(
                    ReadString(reader, "class_name"),
                    ReadNullableInt(reader, "employee_count"),
                    null,
                    null,
                    ReadDecimal(reader, "total_premium"),
                    "Billing statement totals"));
            }
        }

        if (rows.Count == 0 && await TableExistsAsync(connection, "Premium"))
        {
            await using var command = connection.CreateCommand();
            command.CommandText = """
                SELECT
                    class_name,
                    SUM(CASE WHEN LOWER(COALESCE(period_label, '')) LIKE '%annual%' THEN amount / 12.0 ELSE amount END) AS monthly_total,
                    MAX(COALESCE(source_file_name, '')) AS source_document
                FROM Premium
                WHERE client_id = $client_id
                  AND TRIM(COALESCE(class_name, '')) <> ''
                GROUP BY class_name;
                """;
            command.Parameters.AddWithValue("$client_id", clientId);
            await using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                rows.Add(new ClassCostSourceRow(
                    ReadString(reader, "class_name"),
                    null,
                    null,
                    null,
                    ReadDecimal(reader, "monthly_total"),
                    ReadString(reader, "source_document")));
            }
        }

        return rows;
    }

    private static async Task<int?> ReadTotalEligibleEmployeesAsync(
        SqliteConnection connection,
        long clientId,
        IReadOnlyList<ClassCostSourceRow> costs)
    {
        var classEmployeeCount = costs.Where(item => item.EmployeeCount.HasValue).Sum(item => item.EmployeeCount!.Value);
        if (classEmployeeCount > 0)
        {
            return classEmployeeCount;
        }

        if (await TableExistsAsync(connection, "BillingStatements"))
        {
            await using var billingCommand = connection.CreateCommand();
            billingCommand.CommandText = """
                SELECT employee_count
                FROM BillingStatements
                WHERE client_id = $client_id
                  AND is_deleted = 0
                  AND is_replaced = 0
                ORDER BY statement_year DESC, statement_month DESC, import_version DESC
                LIMIT 1;
                """;
            billingCommand.Parameters.AddWithValue("$client_id", clientId);
            var billingCount = await billingCommand.ExecuteScalarAsync();
            if (billingCount is not null && billingCount != DBNull.Value && Convert.ToInt32(billingCount, CultureInfo.InvariantCulture) > 0)
            {
                return Convert.ToInt32(billingCount, CultureInfo.InvariantCulture);
            }
        }

        if (await TableExistsAsync(connection, "Renewals"))
        {
            await using var renewalCommand = connection.CreateCommand();
            renewalCommand.CommandText = "SELECT MAX(lives) FROM Renewals WHERE client_id = $client_id;";
            renewalCommand.Parameters.AddWithValue("$client_id", clientId);
            var renewalCount = await renewalCommand.ExecuteScalarAsync();
            if (renewalCount is not null && renewalCount != DBNull.Value && Convert.ToInt32(renewalCount, CultureInfo.InvariantCulture) > 0)
            {
                return Convert.ToInt32(renewalCount, CultureInfo.InvariantCulture);
            }
        }

        return null;
    }

    private static async Task<ClassComparisonPreferenceState> ReadClassComparisonPreferencesAsync(
        SqliteConnection connection,
        long? clientId)
    {
        if (!clientId.HasValue || !await TableExistsAsync(connection, "class_comparison_preferences"))
        {
            return new ClassComparisonPreferenceState { ClientId = clientId };
        }

        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                baseline_class_name,
                selected_class_names,
                active_tab,
                key_benefits_only,
                highlight_differences,
                collapsed_sections,
                recommended_class_name,
                updated_by
            FROM class_comparison_preferences
            WHERE client_id = $client_id
            LIMIT 1;
            """;
        command.Parameters.AddWithValue("$client_id", clientId.Value);
        await using var reader = await command.ExecuteReaderAsync();
        if (!await reader.ReadAsync())
        {
            return new ClassComparisonPreferenceState { ClientId = clientId };
        }

        return new ClassComparisonPreferenceState
        {
            ClientId = clientId,
            BaselineClassName = ReadString(reader, "baseline_class_name"),
            SelectedClassNames = DeserializeStringList(ReadString(reader, "selected_class_names")),
            ActiveTab = FirstNonBlank(ReadString(reader, "active_tab"), "Class Comparison"),
            KeyBenefitsOnly = ReadLong(reader, "key_benefits_only") == 1,
            HighlightDifferences = ReadLong(reader, "highlight_differences") != 0,
            CollapsedSections = DeserializeStringList(ReadString(reader, "collapsed_sections")),
            RecommendedClassName = ReadString(reader, "recommended_class_name"),
            UpdatedBy = ReadString(reader, "updated_by")
        };
    }

    private static async Task<List<ClassComparisonNoteRow>> ReadClassComparisonNotesAsync(
        SqliteConnection connection,
        long clientId)
    {
        if (!await TableExistsAsync(connection, "class_comparison_notes"))
        {
            return [];
        }

        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT note_id, client_id, class_name, note_text, created_at, created_by
            FROM class_comparison_notes
            WHERE client_id = $client_id
            ORDER BY created_at DESC
            LIMIT 100;
            """;
        command.Parameters.AddWithValue("$client_id", clientId);
        var rows = new List<ClassComparisonNoteRow>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rows.Add(new ClassComparisonNoteRow
            {
                NoteId = ReadLong(reader, "note_id"),
                ClientId = ReadLong(reader, "client_id"),
                ClassName = ReadString(reader, "class_name"),
                NoteText = ReadString(reader, "note_text"),
                Author = ReadString(reader, "created_by"),
                CreatedAt = ReadDateTime(reader, "created_at") ?? DateTime.MinValue
            });
        }

        return rows;
    }

    private static async Task<List<ClassComparisonHistoryRow>> ReadClassComparisonHistoryAsync(
        SqliteConnection connection,
        long clientId)
    {
        var hasComparisonAudit = await TableExistsAsync(connection, "class_comparison_audit");
        var hasExtractionAudit = await TableExistsAsync(connection, "extraction_audit_log") && await TableExistsAsync(connection, "Documents");
        if (!hasComparisonAudit && !hasExtractionAudit)
        {
            return [];
        }

        var statements = new List<string>();
        if (hasComparisonAudit)
        {
            statements.Add("""
                SELECT
                    audit_id AS history_id,
                    created_at,
                    created_by AS author,
                    class_name,
                    field_name,
                    previous_value,
                    new_value,
                    reason,
                    source_document
                FROM class_comparison_audit
                WHERE client_id = $client_id
                """);
        }

        if (hasExtractionAudit)
        {
            statements.Add("""
                SELECT
                    -a.audit_id AS history_id,
                    a.created_at,
                    'Extraction' AS author,
                    '' AS class_name,
                    a.field_name,
                    '' AS previous_value,
                    a.extracted_value AS new_value,
                    COALESCE(NULLIF(a.validation_message, ''), NULLIF(a.section_type, ''), 'Source extraction') AS reason,
                    COALESCE(NULLIF(a.source_file, ''), NULLIF(d.file_name, ''), '') AS source_document
                FROM extraction_audit_log a
                JOIN Documents d ON d.document_id = a.source_document_id
                WHERE d.client_id = $client_id
                """);
        }

        await using var command = connection.CreateCommand();
        command.CommandText = $"""
            SELECT * FROM (
                {string.Join(" UNION ALL ", statements)}
            ) history
            ORDER BY created_at DESC
            LIMIT 100;
            """;
        command.Parameters.AddWithValue("$client_id", clientId);
        var rows = new List<ClassComparisonHistoryRow>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var createdAt = ReadDateTime(reader, "created_at");
            rows.Add(new ClassComparisonHistoryRow
            {
                HistoryId = ReadLong(reader, "history_id"),
                DateText = createdAt.HasValue
                    ? createdAt.Value.ToString("MMM d, yyyy h:mm tt", CultureInfo.InvariantCulture)
                    : ReadString(reader, "created_at"),
                Author = ReadString(reader, "author"),
                ClassName = ReadString(reader, "class_name"),
                FieldName = ReadString(reader, "field_name"),
                PreviousValue = FirstNonBlank(ReadString(reader, "previous_value"), "-") ,
                NewValue = FirstNonBlank(ReadString(reader, "new_value"), "-"),
                Reason = ReadString(reader, "reason"),
                SourceDocument = ReadString(reader, "source_document")
            });
        }

        return rows;
    }

    private static async Task AddPreferenceAuditRowsAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        ClassComparisonPreferenceState previous,
        ClassComparisonPreferenceState current)
    {
        var changes = new List<(string Field, string Previous, string Current)>
        {
            ("Baseline Class", previous.BaselineClassName, current.BaselineClassName),
            ("Selected Classes", string.Join(" | ", previous.SelectedClassNames), string.Join(" | ", current.SelectedClassNames)),
            ("Active Tab", previous.ActiveTab, current.ActiveTab),
            ("Benefit View", previous.KeyBenefitsOnly ? "Key Benefits Only" : "All Benefits", current.KeyBenefitsOnly ? "Key Benefits Only" : "All Benefits"),
            ("Highlight Differences", previous.HighlightDifferences ? "On" : "Off", current.HighlightDifferences ? "On" : "Off"),
            ("Collapsed Sections", string.Join(" | ", previous.CollapsedSections), string.Join(" | ", current.CollapsedSections)),
            ("Recommended Class", previous.RecommendedClassName, current.RecommendedClassName)
        };

        foreach (var change in changes.Where(change => !change.Previous.Equals(change.Current, StringComparison.OrdinalIgnoreCase)))
        {
            await InsertClassComparisonAuditAsync(
                connection,
                transaction,
                clientId,
                change.Field is "Baseline Class" or "Recommended Class" ? change.Current : "",
                change.Field,
                change.Previous,
                change.Current,
                "Class comparison preference changed in Client Workspace.",
                "",
                FirstNonBlank(current.UpdatedBy, "Desktop user"),
                DateTime.Now);
        }
    }

    private static async Task InsertClassComparisonAuditAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string className,
        string fieldName,
        string previousValue,
        string newValue,
        string reason,
        string sourceDocument,
        string createdBy,
        DateTime createdAt)
    {
        await using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            INSERT INTO class_comparison_audit (
                client_id,
                class_name,
                field_name,
                previous_value,
                new_value,
                reason,
                source_document,
                created_at,
                created_by
            )
            VALUES (
                $client_id,
                $class_name,
                $field_name,
                $previous_value,
                $new_value,
                $reason,
                $source_document,
                $created_at,
                $created_by
            );
            """;
        command.Parameters.AddWithValue("$client_id", clientId);
        command.Parameters.AddWithValue("$class_name", FirstNonBlank(className));
        command.Parameters.AddWithValue("$field_name", FirstNonBlank(fieldName));
        command.Parameters.AddWithValue("$previous_value", FirstNonBlank(previousValue));
        command.Parameters.AddWithValue("$new_value", FirstNonBlank(newValue));
        command.Parameters.AddWithValue("$reason", FirstNonBlank(reason));
        command.Parameters.AddWithValue("$source_document", FirstNonBlank(sourceDocument));
        command.Parameters.AddWithValue("$created_at", createdAt.ToString("O", CultureInfo.InvariantCulture));
        command.Parameters.AddWithValue("$created_by", FirstNonBlank(createdBy, "Desktop user"));
        await command.ExecuteNonQueryAsync();
    }

    private static IReadOnlyList<string> DeserializeStringList(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return [];
        }

        try
        {
            return JsonSerializer.Deserialize<List<string>>(json)?
                       .Where(value => !string.IsNullOrWhiteSpace(value))
                       .Select(value => value.Trim())
                       .Distinct(StringComparer.OrdinalIgnoreCase)
                       .ToList()
                   ?? [];
        }
        catch (JsonException)
        {
            return json.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        }
    }

    private async Task<ClientWorkspaceClient?> ResolveClientAsync(SqliteConnection connection, long? clientId, string? clientName, string? carrier)
    {
        if (!await TableExistsAsync(connection, "Clients"))
        {
            return null;
        }

        if (clientId.HasValue)
        {
            var byId = await ReadClientAsync(connection, "WHERE client_id = $client_id", command => command.Parameters.AddWithValue("$client_id", clientId.Value));
            if (byId is not null)
            {
                return byId;
            }
        }

        if (!string.IsNullOrWhiteSpace(clientName))
        {
            var byName = await ReadClientAsync(connection, "WHERE client_name = $client_name COLLATE NOCASE", command => command.Parameters.AddWithValue("$client_name", clientName.Trim()));
            if (byName is not null)
            {
                return byName;
            }

            var byLike = await ReadClientAsync(connection, "WHERE client_name LIKE $client_like COLLATE NOCASE ORDER BY LENGTH(client_name) LIMIT 1", command => command.Parameters.AddWithValue("$client_like", $"%{clientName.Trim()}%"));
            if (byLike is not null)
            {
                return byLike;
            }
        }

        return null;
    }

    private static async Task<List<ClientWorkspaceClient>> ReadClientOptionsAsync(SqliteConnection connection)
    {
        if (!await TableExistsAsync(connection, "Clients"))
        {
            return [];
        }

        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT client_id, client_name, client_number, carrier, effective_date, industry, province, advisor, status
            FROM Clients
            WHERE TRIM(COALESCE(client_name, '')) <> ''
            ORDER BY client_name;
            """;

        var clients = new List<ClientWorkspaceClient>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            clients.Add(new ClientWorkspaceClient(
                ReadLong(reader, "client_id"),
                ReadString(reader, "client_name"),
                ReadString(reader, "client_number"),
                ReadString(reader, "carrier"),
                ReadDateOnly(reader, "effective_date"),
                ReadString(reader, "industry"),
                ReadString(reader, "province"),
                ReadString(reader, "advisor"),
                ReadString(reader, "status")));
        }

        return clients;
    }

    private static async Task<ClientWorkspaceClient?> ReadClientAsync(SqliteConnection connection, string whereClause, Action<SqliteCommand> addParameters)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = $"""
            SELECT client_id, client_name, client_number, carrier, effective_date, industry, province, advisor, status
            FROM Clients
            {whereClause}
            LIMIT 1;
            """;
        addParameters(command);
        await using var reader = await command.ExecuteReaderAsync();
        if (!await reader.ReadAsync())
        {
            return null;
        }

        return new ClientWorkspaceClient(
            ReadLong(reader, "client_id"),
            ReadString(reader, "client_name"),
            ReadString(reader, "client_number"),
            ReadString(reader, "carrier"),
            ReadDateOnly(reader, "effective_date"),
            ReadString(reader, "industry"),
            ReadString(reader, "province"),
            ReadString(reader, "advisor"),
            ReadString(reader, "status"));
    }

    private async Task<List<PlanFeatureRow>> ReadFeatureRowsAsync(SqliteConnection connection, long? clientId)
    {
        if (!clientId.HasValue || !await TableExistsAsync(connection, "plan_design_features"))
        {
            return [];
        }

        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                f.plan_design_feature_id,
                f.benefit,
                f.category,
                f.feature_name,
                f.feature_value,
                f.coinsurance,
                f.deductible,
                f.maximum,
                f.frequency,
                '' AS waiting_period,
                f.termination_age,
                '' AS source_section,
                f.class_number,
                f.class_name,
                f.source_page,
                f.source_text,
                f.confidence_score,
                f.date_loaded,
                COALESCE(NULLIF(d.file_name, ''), 'Saved plan design features') AS source_document,
                COALESCE(NULLIF(d.file_path, ''), '') AS source_document_path,
                COALESCE(NULLIF(d.document_type, ''), '') AS source_document_type,
                COALESCE(NULLIF(d.effective_date, ''), NULLIF(c.effective_date, ''), '') AS effective_date,
                COALESCE(NULLIF(d.carrier, ''), NULLIF(c.carrier, ''), '') AS carrier,
                r.PlanBenefitRuleId AS rule_id,
                r.SourceDocumentId AS rule_source_document_id,
                r.RuleType AS rule_type,
                r.Tier1ReimbursementPercent AS rule_tier1_percent,
                r.Tier1ThresholdAmount AS rule_tier1_threshold,
                r.Tier2ReimbursementPercent AS rule_tier2_percent,
                r.ThresholdBasis AS rule_threshold_basis,
                r.ResetPeriod AS rule_reset_period,
                r.CurrencyCode AS rule_currency_code,
                r.SourceText AS rule_source_text,
                r.SourceLocation AS rule_source_location,
                r.NormalizationStatus AS rule_normalization_status,
                r.ConfidenceScore AS rule_confidence_score,
                r.DisplayOrder AS rule_display_order,
                r.IsActive AS rule_is_active,
                r.EffectiveDate AS rule_effective_date,
                r.TerminationDate AS rule_termination_date,
                r.CreatedAtUtc AS rule_created_at,
                r.UpdatedAtUtc AS rule_updated_at
            FROM plan_design_features f
            LEFT JOIN PlanBenefitRule r
                ON r.PlanDesignFeatureId = f.plan_design_feature_id
               AND r.RuleType = 'TieredReimbursement'
               AND r.IsActive = 1
            LEFT JOIN Documents d ON d.document_id = f.source_document_id
            LEFT JOIN Clients c ON c.client_id = f.client_id
            WHERE f.client_id = $client_id
            ORDER BY f.class_number, f.benefit, f.category, f.feature_name;
            """;
        command.Parameters.AddWithValue("$client_id", clientId.Value);

        var rows = new List<PlanFeatureRow>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rows.Add(new PlanFeatureRow(
                ReadString(reader, "benefit"),
                ReadString(reader, "category"),
                ReadString(reader, "feature_name"),
                ReadString(reader, "feature_value"),
                ReadString(reader, "coinsurance"),
                ReadString(reader, "deductible"),
                ReadString(reader, "maximum"),
                ReadString(reader, "frequency"),
                ReadString(reader, "waiting_period"),
                ReadString(reader, "termination_age"),
                ReadString(reader, "source_section"),
                ReadString(reader, "class_number"),
                ReadString(reader, "class_name"),
                ReadString(reader, "source_page"),
                ReadString(reader, "source_document"),
                ReadString(reader, "source_document_path"),
                ReadString(reader, "source_document_type"),
                ReadString(reader, "effective_date"),
                ReadString(reader, "carrier"),
                ReadString(reader, "source_text"),
                ReadDecimal(reader, "confidence_score"),
                ReadDateTime(reader, "date_loaded"),
                ReadLong(reader, "plan_design_feature_id"),
                ReadPlanBenefitRules(reader, clientId.Value, ReadLong(reader, "plan_design_feature_id"))));
        }

        return rows;
    }

    private async Task<List<PlanFeatureRow>> ReadLegacyPlanDesignRowsAsync(SqliteConnection connection, long? clientId)
    {
        if (!clientId.HasValue || !await TableExistsAsync(connection, "PlanDesign"))
        {
            return [];
        }

        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                p.benefit_category,
                p.benefit_name,
                p.class_name,
                p.coverage_description,
                p.reimbursement_percent,
                p.maximum_amount,
                p.frequency,
                p.waiting_period,
                p.termination_age,
                p.source_section,
                p.notes,
                p.source_page,
                p.date_loaded,
                COALESCE(NULLIF(d.file_name, ''), NULLIF(p.booklet_source_file_name, ''), 'PlanDesign') AS source_document,
                COALESCE(NULLIF(d.file_path, ''), '') AS source_document_path,
                COALESCE(NULLIF(d.document_type, ''), '') AS source_document_type,
                COALESCE(NULLIF(d.effective_date, ''), NULLIF(c.effective_date, ''), '') AS effective_date,
                COALESCE(NULLIF(d.carrier, ''), NULLIF(c.carrier, ''), '') AS carrier
            FROM PlanDesign p
            LEFT JOIN Documents d ON d.document_id = p.source_document_id
            LEFT JOIN Clients c ON c.client_id = p.client_id
            WHERE p.client_id = $client_id
            ORDER BY p.class_name, p.benefit_category, p.benefit_name;
            """;
        command.Parameters.AddWithValue("$client_id", clientId.Value);

        var rows = new List<PlanFeatureRow>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var reimbursement = ReadDecimal(reader, "reimbursement_percent");
            var benefitName = ReadString(reader, "benefit_name");
            var coverageDescription = ReadString(reader, "coverage_description");
            var (featureName, featureValue) = ParseLegacyPlanDesignFeature(benefitName, coverageDescription);
            rows.Add(new PlanFeatureRow(
                ReadString(reader, "benefit_category"),
                benefitName,
                featureName,
                featureValue,
                reimbursement.HasValue ? $"{reimbursement.Value:0.#}%" : "",
                "",
                ReadString(reader, "maximum_amount"),
                ReadString(reader, "frequency"),
                ReadString(reader, "waiting_period"),
                ReadString(reader, "termination_age"),
                ReadString(reader, "source_section"),
                "",
                FirstNonBlank(ReadString(reader, "class_name"), "All Classes"),
                ReadString(reader, "source_page"),
                ReadString(reader, "source_document"),
                ReadString(reader, "source_document_path"),
                ReadString(reader, "source_document_type"),
                ReadString(reader, "effective_date"),
                ReadString(reader, "carrier"),
                ReadString(reader, "notes"),
                null,
                ReadDateTime(reader, "date_loaded"),
                null,
                []));
        }

        return rows;
    }

    private static (string FeatureName, string FeatureValue) ParseLegacyPlanDesignFeature(
        string benefitName,
        string coverageDescription)
    {
        var labelledValue = benefitName.Contains("Critical Illness", StringComparison.OrdinalIgnoreCase)
            ? Regex.Match(coverageDescription, @"^(?<label>[A-Za-z][A-Za-z0-9 &/()'\-]{1,80}):\s*(?<value>.+)$")
            : Match.Empty;
        if (labelledValue.Success)
        {
            return (
                labelledValue.Groups["label"].Value.Trim(),
                labelledValue.Groups["value"].Value.Trim());
        }

        return (benefitName, coverageDescription);
    }

    private static async Task<bool> TableExistsAsync(SqliteConnection connection, string tableName)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type = 'table' AND name = $name;";
        command.Parameters.AddWithValue("$name", tableName);
        var result = await command.ExecuteScalarAsync();
        return Convert.ToInt32(result, CultureInfo.InvariantCulture) > 0;
    }

    private static string SummarizeClassRows(IEnumerable<PlanFeatureRow> rows, string className, string benefitArea)
    {
        var selected = rows
            .Where(row => ClassMatches(row, className))
            .ToList();

        if (selected.Count == 0 && string.IsNullOrWhiteSpace(className))
        {
            selected = rows.ToList();
        }

        if (benefitArea.Equals("Paramedical Services", StringComparison.OrdinalIgnoreCase))
        {
            return SummarizeParamedicalClass(selected);
        }

        if (benefitArea.Equals("Critical Illness", StringComparison.OrdinalIgnoreCase))
        {
            return SummarizeCriticalIllnessClass(selected);
        }

        return SummarizeRows(selected);
    }

    private static string SummarizeParamedicalClass(IReadOnlyList<PlanFeatureRow> rows)
    {
        if (rows.Count == 0)
        {
            return "N/A";
        }

        var row = rows[0];
        var tieredReimbursement = GetTieredReimbursementDisplay(row, compact: true);
        var values = new[]
        {
            FirstNonBlank(tieredReimbursement, row.Coinsurance, ExtractPercent(row.FeatureValue)),
            FirstNonBlank(row.Maximum, ExtractMaximumText(row.FeatureValue)),
            ExtractPractitionerLimit(row.FeatureValue)
        }
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return values.Count == 0 ? DisplayValue(row) : string.Join(Environment.NewLine, values);
    }

    private static string SummarizeCriticalIllnessClass(IReadOnlyList<PlanFeatureRow> rows)
    {
        if (rows.Count == 0)
        {
            return "N/A";
        }

        var lines = new List<string>();
        var provider = FirstNonBlank(rows.Select(FindBenefitProvider));
        if (!string.IsNullOrWhiteSpace(provider))
        {
            lines.Add($"Provider: {provider}");
        }

        var employeeAmount = FindCriticalIllnessAmount(rows, "employee", "member");
        var spouseAmount = FindCriticalIllnessAmount(rows, "spouse");
        var childAmount = FindCriticalIllnessAmount(rows, "child");
        var dependentAmount = FindCriticalIllnessAmount(rows, "dependent");

        if (!string.IsNullOrWhiteSpace(employeeAmount) &&
            employeeAmount.Equals(spouseAmount, StringComparison.OrdinalIgnoreCase) &&
            employeeAmount.Equals(childAmount, StringComparison.OrdinalIgnoreCase))
        {
            lines.Add($"Employee, spouse & child: {employeeAmount}");
        }
        else
        {
            if (!string.IsNullOrWhiteSpace(employeeAmount))
            {
                lines.Add($"Employee: {employeeAmount}");
            }

            if (!string.IsNullOrWhiteSpace(spouseAmount) &&
                spouseAmount.Equals(childAmount, StringComparison.OrdinalIgnoreCase))
            {
                lines.Add($"Spouse & child: {spouseAmount}");
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(spouseAmount))
                {
                    lines.Add($"Spouse: {spouseAmount}");
                }

                if (!string.IsNullOrWhiteSpace(childAmount))
                {
                    lines.Add($"Child: {childAmount}");
                }
            }

            if (string.IsNullOrWhiteSpace(spouseAmount) &&
                string.IsNullOrWhiteSpace(childAmount) &&
                !string.IsNullOrWhiteSpace(dependentAmount))
            {
                lines.Add($"Dependent: {dependentAmount}");
            }
        }

        if (string.IsNullOrWhiteSpace(employeeAmount) &&
            string.IsNullOrWhiteSpace(spouseAmount) &&
            string.IsNullOrWhiteSpace(childAmount) &&
            string.IsNullOrWhiteSpace(dependentAmount))
        {
            var sourceLabelledAmount = rows
                .Where(row => ContainsAny(row.FeatureValue, row.SourceText, "employee", "member", "spouse", "child", "dependent"))
                .Select(row => CleanComparisonText(DisplayValue(row), "Critical Illness", row.FeatureName))
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value) && value.Contains('$'));
            if (!string.IsNullOrWhiteSpace(sourceLabelledAmount))
            {
                lines.Add(sourceLabelledAmount);
            }
        }

        var waitingRow = rows.FirstOrDefault(row =>
            ContainsAny(row.FeatureName, row.WaitingPeriod, "waiting period", "eligibility waiting"));
        if (waitingRow is not null)
        {
            var waitingPeriod = CleanComparisonText(
                FirstNonBlank(waitingRow.WaitingPeriod, waitingRow.FeatureValue),
                "Critical Illness",
                waitingRow.FeatureName);
            if (!string.IsNullOrWhiteSpace(waitingPeriod))
            {
                lines.Add($"Waiting period: {waitingPeriod}");
            }
        }

        return lines.Count == 0
            ? SummarizeRows(rows)
            : string.Join(Environment.NewLine, lines.Distinct(StringComparer.OrdinalIgnoreCase));
    }

    private static string FindCriticalIllnessAmount(
        IReadOnlyList<PlanFeatureRow> rows,
        params string[] insuredPersonTerms)
    {
        foreach (var row in rows.Where(row => ContainsAny(row.FeatureName, "", insuredPersonTerms)))
        {
            var value = FirstNonBlank(row.Maximum, row.FeatureValue);
            if (string.IsNullOrWhiteSpace(value) ||
                (string.IsNullOrWhiteSpace(row.Maximum) &&
                 !ContainsAny(row.FeatureName, "", "amount", "maximum", "benefit")))
            {
                continue;
            }

            return CleanComparisonText(value, "Critical Illness", row.FeatureName);
        }

        return "";
    }

    private static string SummarizeCompactComparisonRows(IEnumerable<PlanFeatureRow> rows, string className, string benefitArea)
    {
        var selected = rows
            .Where(row => ClassMatches(row, className))
            .ToList();

        if (selected.Count == 0 && string.IsNullOrWhiteSpace(className))
        {
            selected = rows.ToList();
        }

        if (selected.Count == 0)
        {
            return "N/A";
        }

        if (benefitArea.Equals("Critical Illness", StringComparison.OrdinalIgnoreCase))
        {
            return SummarizeCriticalIllnessClass(selected);
        }

        var values = selected
            .OrderBy(row => FeatureOrder(row.FeatureName))
            .Select(row => CompactComparisonValue(row, benefitArea))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(3)
            .ToList();

        return values.Count == 0 ? "Saved" : string.Join("  |  ", values);
    }

    private static string SummarizeFeatureRows(IEnumerable<PlanFeatureRow> rows, string className)
    {
        var selected = rows
            .Where(row => ClassMatches(row, className))
            .ToList();

        return FirstNonBlank(selected.Select(DisplayValue).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)), "N/A");
    }

    private static bool ClassMatches(PlanFeatureRow row, string className)
    {
        if (string.IsNullOrWhiteSpace(className))
        {
            return true;
        }

        var rowClassName = GetClassName(row);
        return rowClassName.Equals(className, StringComparison.OrdinalIgnoreCase)
               || IsGenericAllClasses(rowClassName);
    }

    private static bool IsGenericAllClasses(string className)
    {
        var normalized = (className ?? "").Trim();
        return normalized.Length == 0
               || normalized.Equals("All", StringComparison.OrdinalIgnoreCase)
               || normalized.Equals("All Classes", StringComparison.OrdinalIgnoreCase);
    }

    private static string SummarizeRows(IReadOnlyList<PlanFeatureRow> rows)
    {
        if (rows.Count == 0)
        {
            return "N/A";
        }

        var values = rows
            .OrderBy(row => FeatureOrder(row.FeatureName))
            .Select(row =>
            {
                var label = FirstNonBlank(row.FeatureName, row.Category, row.Benefit);
                var value = DisplayValue(row);
                return string.IsNullOrWhiteSpace(label) || label.Equals(value, StringComparison.OrdinalIgnoreCase)
                    ? value
                    : $"{label}: {value}";
            })
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(3)
            .ToList();

        return values.Count == 0 ? "Saved" : string.Join(Environment.NewLine, values);
    }

    private static string DisplayValue(PlanFeatureRow row)
    {
        return FirstNonBlank(
            GetTieredReimbursementDisplay(row, compact: false),
            row.FeatureValue,
            row.Coinsurance,
            row.Maximum,
            row.Deductible,
            row.Frequency,
            row.TerminationAge,
            row.SourceText);
    }

    private static string CompactComparisonValue(PlanFeatureRow row, string benefitArea)
    {
        var tieredReimbursement = GetTieredReimbursementDisplay(row, compact: true);
        if (!string.IsNullOrWhiteSpace(tieredReimbursement))
        {
            return tieredReimbursement;
        }

        if (!string.IsNullOrWhiteSpace(row.Coinsurance))
        {
            return LimitComparisonText($"{NormalizePercent(row.Coinsurance)} coinsurance", 42);
        }

        if (!string.IsNullOrWhiteSpace(row.Maximum))
        {
            return LimitComparisonText($"{CleanComparisonText(row.Maximum, benefitArea, row.FeatureName)} max", 42);
        }

        if (!string.IsNullOrWhiteSpace(row.Deductible))
        {
            var deductible = CleanComparisonText(row.Deductible, benefitArea, row.FeatureName);
            return IsZeroLikeValue(deductible) ? "No deductible" : LimitComparisonText($"{deductible} deductible", 42);
        }

        if (!string.IsNullOrWhiteSpace(row.Frequency))
        {
            return LimitComparisonText(CleanComparisonText(row.Frequency, benefitArea, row.FeatureName), 42);
        }

        if (!string.IsNullOrWhiteSpace(row.TerminationAge))
        {
            var terminationAge = CleanComparisonText(row.TerminationAge, benefitArea, row.FeatureName)
                .Replace("Age ", "", StringComparison.OrdinalIgnoreCase)
                .Replace("age ", "", StringComparison.OrdinalIgnoreCase);
            return LimitComparisonText($"Term age {terminationAge}", 42);
        }

        return LimitComparisonText(CleanComparisonText(DisplayValue(row), benefitArea, row.FeatureName), 48);
    }

    private static string GetTieredReimbursementDisplay(PlanFeatureRow row, bool compact)
    {
        var tieredRules = row.Rules
            .Where(candidate => candidate.RuleType == BenefitRuleType.TieredReimbursement)
            .OrderBy(candidate => candidate.DisplayOrder)
            .ToList();
        var rule = tieredRules
            .Where(candidate => candidate.IsDisplayable)
            .FirstOrDefault();
        if (rule is null)
        {
            return tieredRules.Count > 0
                ? BenefitRuleDisplayFormatter.ReviewRequiredDisplayText
                : "";
        }

        return compact
            ? BenefitRuleDisplayFormatter.FormatCompact(rule)
            : BenefitRuleDisplayFormatter.FormatFull(rule);
    }

    private static string CleanComparisonText(string? value, string benefitArea, string? featureName)
    {
        var text = FirstNonBlank(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return "";
        }

        text = Regex.Replace(text, @"\s+", " ").Trim();
        text = RemoveLeadingLabel(text, benefitArea);
        text = RemoveLeadingLabel(text, featureName);

        var salaryMatch = Regex.Match(text, @"(\d+(?:\.\d+)?)\s*times\s+annual\s+earnings", RegexOptions.IgnoreCase);
        if (salaryMatch.Success)
        {
            return $"{salaryMatch.Groups[1].Value}x annual earnings";
        }

        var percentMatch = Regex.Match(text, @"\b(100|[1-9]?\d(?:\.\d+)?)\s*%", RegexOptions.IgnoreCase);
        if (percentMatch.Success && text.Length > 70)
        {
            return $"{percentMatch.Groups[1].Value}%";
        }

        return text;
    }

    private static string RemoveLeadingLabel(string text, string? label)
    {
        if (string.IsNullOrWhiteSpace(label))
        {
            return text;
        }

        var normalized = label.Trim();
        if (text.StartsWith(normalized + ":", StringComparison.OrdinalIgnoreCase))
        {
            return text[(normalized.Length + 1)..].Trim();
        }

        return text;
    }

    private static string NormalizePercent(string value)
    {
        var text = CleanComparisonText(value, "", "");
        var match = Regex.Match(text, @"\b(100|[1-9]?\d(?:\.\d+)?)\s*%");
        return match.Success ? $"{match.Groups[1].Value}%" : text;
    }

    private static bool IsZeroLikeValue(string value)
    {
        var text = value.Trim();
        return text.Equals("0", StringComparison.OrdinalIgnoreCase)
               || text.Equals("$0", StringComparison.OrdinalIgnoreCase)
               || text.Equals("nil", StringComparison.OrdinalIgnoreCase)
               || text.Equals("none", StringComparison.OrdinalIgnoreCase)
               || text.Contains("$0", StringComparison.OrdinalIgnoreCase);
    }

    private static string LimitComparisonText(string value, int maxLength)
    {
        var text = Regex.Replace(FirstNonBlank(value), @"\s+", " ").Trim();
        if (text.Length <= maxLength)
        {
            return text;
        }

        return $"{text[..Math.Max(0, maxLength - 3)].TrimEnd()}...";
    }

    private static (bool IsSame, string Text, Brush Brush) CompareText(string class1Value, string class2Value)
    {
        if (string.IsNullOrWhiteSpace(class2Value) || class2Value.Equals("N/A", StringComparison.OrdinalIgnoreCase))
        {
            return (true, "Single class", Brushes.SeaGreen);
        }

        var same = NormalizeComparable(class1Value).Equals(NormalizeComparable(class2Value), StringComparison.OrdinalIgnoreCase);
        return same
            ? (true, "Same", Brushes.SeaGreen)
            : (false, "Different", Brushes.DarkOrange);
    }

    private static (bool IsSame, string Text, Brush Brush) CompareClassValues(params string[] classValues)
    {
        var comparableValues = classValues
            .Where(value => !string.IsNullOrWhiteSpace(value) && !value.Equals("N/A", StringComparison.OrdinalIgnoreCase))
            .Select(NormalizeComparable)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (comparableValues.Count <= 1)
        {
            return (true, "Same", Brushes.SeaGreen);
        }

        return (false, "Different", Brushes.DarkOrange);
    }

    private static string NormalizeComparable(string value)
    {
        return new string((value ?? "").Where(ch => !char.IsWhiteSpace(ch)).ToArray());
    }

    private static string NormalizeBenefitArea(PlanFeatureRow row)
    {
        var value = string.Join(" ", new[] { row.Benefit, row.Category, row.FeatureName }
            .Where(item => !string.IsNullOrWhiteSpace(item)));
        value = FirstNonBlank(value, "Plan Design");
        value = value.Replace("_", " ", StringComparison.OrdinalIgnoreCase).Trim();
        if (ContainsAny(value, "", "drug"))
        {
            return "Prescription Drugs";
        }

        if (ContainsAny(value, "", "vision"))
        {
            return "Vision";
        }

        if (ContainsAny(value, "", "paramedical", "practitioner"))
        {
            return "Paramedical Services";
        }

        if (ContainsAny(value, "", "life", "ad&d", "accidental death"))
        {
            return "Life & AD&D";
        }

        if (ContainsAny(value, "", "critical illness"))
        {
            return "Critical Illness";
        }

        if (ContainsAny(value, "", "long term disability", "ltd"))
        {
            return "Long Term Disability";
        }

        if (ContainsAny(value, "", "employee assistance", "eap"))
        {
            return "EAP";
        }

        if (ContainsAny(value, "", "dental"))
        {
            return "Dental";
        }

        if (ContainsAny(value, "", "extended health", "health care", "healthcare", "ehc"))
        {
            return "Extended Healthcare (EHC)";
        }

        return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(value.ToLowerInvariant());
    }

    private static int BenefitOrder(string benefit)
    {
        var normalized = benefit.ToLowerInvariant();
        if (normalized.Contains("life", StringComparison.OrdinalIgnoreCase)) return 1;
        if (normalized.Contains("critical", StringComparison.OrdinalIgnoreCase)) return 2;
        if (normalized.Contains("disability", StringComparison.OrdinalIgnoreCase) || normalized.Contains("ltd", StringComparison.OrdinalIgnoreCase)) return 3;
        if (normalized.Contains("health", StringComparison.OrdinalIgnoreCase) || normalized.Contains("ehc", StringComparison.OrdinalIgnoreCase)) return 4;
        if (normalized.Contains("drug", StringComparison.OrdinalIgnoreCase)) return 5;
        if (normalized.Contains("paramedical", StringComparison.OrdinalIgnoreCase)) return 6;
        if (normalized.Contains("vision", StringComparison.OrdinalIgnoreCase)) return 7;
        if (normalized.Contains("dental", StringComparison.OrdinalIgnoreCase)) return 8;
        if (normalized.Contains("eap", StringComparison.OrdinalIgnoreCase)) return 9;
        return 99;
    }

    private static int FeatureOrder(string feature)
    {
        var normalized = feature.ToLowerInvariant();
        if (normalized.Contains("coinsurance", StringComparison.OrdinalIgnoreCase) || normalized.Contains("reimbursement", StringComparison.OrdinalIgnoreCase)) return 1;
        if (normalized.Contains("deductible", StringComparison.OrdinalIgnoreCase)) return 2;
        if (normalized.Contains("maximum", StringComparison.OrdinalIgnoreCase) || normalized.Contains("max", StringComparison.OrdinalIgnoreCase)) return 3;
        if (normalized.Contains("waiting", StringComparison.OrdinalIgnoreCase)) return 4;
        if (normalized.Contains("termination", StringComparison.OrdinalIgnoreCase)) return 5;
        return 50;
    }

    private static string GetClassName(PlanFeatureRow row)
    {
        return FirstNonBlank(row.ClassName, string.IsNullOrWhiteSpace(row.ClassNumber) ? "" : $"Class {row.ClassNumber}", "All Classes");
    }

    private static int CalculateMissingItems(IReadOnlyList<PlanFeatureRow> rows, string sourceDocumentPath)
    {
        var count = 0;
        if (string.IsNullOrWhiteSpace(sourceDocumentPath))
        {
            count++;
        }

        if (!rows.Any(row => ContainsAny(row.Benefit, row.Category, "ehc", "health", "extended")))
        {
            count++;
        }

        if (!rows.Any(row => ContainsAny(row.Benefit, row.Category, "dental")))
        {
            count++;
        }

        return count;
    }

    private static List<PlanDesignFlag> BuildSourceVerificationWarnings(IReadOnlyList<PlanFeatureRow> rows)
    {
        var criticalIllnessRows = rows
            .Where(row => NormalizeBenefitArea(row).Equals("Critical Illness", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var providers = criticalIllnessRows
            .Where(IsSoldPlanSource)
            .Select(FindBenefitProvider)
            .Where(provider => !string.IsNullOrWhiteSpace(provider))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var warnings = new List<PlanDesignFlag>();

        foreach (var provider in providers)
        {
            var confirmedByPolicyOrBooklet = criticalIllnessRows.Any(row =>
                IsPolicyOrBookletSource(row) &&
                ContainsProviderEvidence(row, provider));
            if (confirmedByPolicyOrBooklet)
            {
                continue;
            }

            warnings.Add(new PlanDesignFlag
            {
                Label = $"Critical Illness provider is shown as {provider} in the sold plan only. Verify against a separate {provider} policy, certificate, or master application.",
                Value = "Verify",
                IndicatorBrush = Brushes.DarkOrange
            });
        }

        return warnings;
    }

    private static string FindBenefitProvider(PlanFeatureRow row)
    {
        var evidence = string.Join(" ", row.FeatureValue, row.SourceText, row.FeatureName);
        var matches = BenefitProviderNames
            .Select(provider => new
            {
                provider.DisplayName,
                Index = evidence.IndexOf(provider.SearchText, StringComparison.OrdinalIgnoreCase),
                provider.SearchText.Length
            })
            .Where(match => match.Index >= 0)
            .OrderBy(match => match.Index)
            .ThenByDescending(match => match.Length)
            .ToList();
        if (matches.Count == 0)
        {
            return "";
        }

        var alternateProvider = matches.FirstOrDefault(match =>
            !row.Carrier.Contains(match.DisplayName, StringComparison.OrdinalIgnoreCase) &&
            !match.DisplayName.Contains(row.Carrier, StringComparison.OrdinalIgnoreCase));
        return alternateProvider?.DisplayName ?? matches[0].DisplayName;
    }

    private static bool ContainsProviderEvidence(PlanFeatureRow row, string provider)
    {
        var evidence = string.Join(" ", row.FeatureValue, row.SourceText, row.FeatureName, row.SourceDocument, row.SourceDocumentPath);
        return BenefitProviderNames
            .Where(candidate => candidate.DisplayName.Equals(provider, StringComparison.OrdinalIgnoreCase))
            .Any(candidate => evidence.Contains(candidate.SearchText, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsSoldPlanSource(PlanFeatureRow row)
    {
        var sourceIdentity = $"{row.SourceDocument} {row.SourceDocumentPath}";
        if (ContainsAny(sourceIdentity, "", "sold plan", "sold design", "sold_plan", "sold-design"))
        {
            return true;
        }

        if (ContainsAny(sourceIdentity, row.SourceDocumentType, "booklet", "policy", "contract", "certificate", "master application", "master_application"))
        {
            return false;
        }

        return ContainsAny(row.SourceText, "", "sold plan", "sold design");
    }

    private static bool IsPolicyOrBookletSource(PlanFeatureRow row)
    {
        if (IsSoldPlanSource(row))
        {
            return false;
        }

        return ContainsAny(
            $"{row.SourceDocument} {row.SourceDocumentPath} {row.SourceText}",
            row.SourceDocumentType,
            "booklet",
            "policy",
            "contract",
            "certificate",
            "master application",
            "master_application");
    }

    private static bool ContainsAny(string? left, string? right, params string[] terms)
    {
        var haystack = $"{left} {right}";
        return terms.Any(term => !string.IsNullOrWhiteSpace(term) && haystack.Contains(term, StringComparison.OrdinalIgnoreCase));
    }

    private static string FirstNonBlank(params string?[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return "";
    }

    private static string FirstNonBlank(IEnumerable<string?> values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return "";
    }

    private static string ReadString(SqliteDataReader reader, string columnName)
    {
        var ordinal = TryGetOrdinal(reader, columnName);
        return ordinal >= 0 && !reader.IsDBNull(ordinal)
            ? Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture)?.Trim() ?? ""
            : "";
    }

    private static long ReadLong(SqliteDataReader reader, string columnName)
    {
        var ordinal = TryGetOrdinal(reader, columnName);
        return ordinal >= 0 && !reader.IsDBNull(ordinal)
            ? Convert.ToInt64(reader.GetValue(ordinal), CultureInfo.InvariantCulture)
            : 0L;
    }

    private static long? ReadNullableLong(SqliteDataReader reader, string columnName)
    {
        var ordinal = TryGetOrdinal(reader, columnName);
        return ordinal >= 0 && !reader.IsDBNull(ordinal)
            ? Convert.ToInt64(reader.GetValue(ordinal), CultureInfo.InvariantCulture)
            : null;
    }

    private static IReadOnlyList<PlanBenefitRule> ReadPlanBenefitRules(
        SqliteDataReader reader,
        long clientId,
        long planDesignFeatureId)
    {
        var ruleId = ReadNullableLong(reader, "rule_id");
        if (!ruleId.HasValue)
        {
            return [];
        }

        return
        [
            new PlanBenefitRule
            {
                PlanBenefitRuleId = ruleId,
                PlanDesignFeatureId = planDesignFeatureId,
                ClientId = clientId,
                SourceDocumentId = ReadNullableLong(reader, "rule_source_document_id"),
                RuleType = ParseEnum(ReadString(reader, "rule_type"), BenefitRuleType.TieredReimbursement),
                Tier1ReimbursementPercent = ReadDecimal(reader, "rule_tier1_percent"),
                Tier1ThresholdAmount = ReadDecimal(reader, "rule_tier1_threshold"),
                Tier2ReimbursementPercent = ReadDecimal(reader, "rule_tier2_percent"),
                ThresholdBasis = ParseNullableEnum<BenefitRuleThresholdBasis>(ReadString(reader, "rule_threshold_basis")),
                ResetPeriod = ParseNullableEnum<BenefitRuleResetPeriod>(ReadString(reader, "rule_reset_period")),
                CurrencyCode = ReadString(reader, "rule_currency_code"),
                SourceText = ReadString(reader, "rule_source_text"),
                SourceLocation = ReadString(reader, "rule_source_location"),
                NormalizationStatus = ParseEnum(
                    ReadString(reader, "rule_normalization_status"),
                    BenefitRuleNormalizationStatus.NeedsReview),
                ConfidenceScore = ReadDecimal(reader, "rule_confidence_score"),
                DisplayOrder = ReadNullableInt(reader, "rule_display_order") ?? 10,
                IsActive = (ReadNullableInt(reader, "rule_is_active") ?? 0) == 1,
                EffectiveDate = ReadDateOnly(reader, "rule_effective_date"),
                TerminationDate = ReadDateOnly(reader, "rule_termination_date"),
                CreatedAtUtc = ReadDateTime(reader, "rule_created_at") ?? DateTime.MinValue,
                UpdatedAtUtc = ReadDateTime(reader, "rule_updated_at") ?? DateTime.MinValue
            }
        ];
    }

    private static T ParseEnum<T>(string value, T fallback) where T : struct, Enum
    {
        return Enum.TryParse<T>(value, true, out var parsed) ? parsed : fallback;
    }

    private static T? ParseNullableEnum<T>(string value) where T : struct, Enum
    {
        return Enum.TryParse<T>(value, true, out var parsed) ? parsed : null;
    }

    private static decimal? ReadDecimal(SqliteDataReader reader, string columnName)
    {
        var text = ReadString(reader, columnName);
        return decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;
    }

    private static int? ReadNullableInt(SqliteDataReader reader, string columnName)
    {
        var ordinal = TryGetOrdinal(reader, columnName);
        return ordinal >= 0 && !reader.IsDBNull(ordinal)
            ? Convert.ToInt32(reader.GetValue(ordinal), CultureInfo.InvariantCulture)
            : null;
    }

    private static DateOnly? ReadDateOnly(SqliteDataReader reader, string columnName)
    {
        var text = ReadString(reader, columnName);
        return DateOnly.TryParse(text, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;
    }

    private static DateTime? ReadDateTime(SqliteDataReader reader, string columnName)
    {
        var text = ReadString(reader, columnName);
        return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var value)
            ? value
            : null;
    }

    private static int TryGetOrdinal(SqliteDataReader reader, string columnName)
    {
        for (var index = 0; index < reader.FieldCount; index++)
        {
            if (reader.GetName(index).Equals(columnName, StringComparison.OrdinalIgnoreCase))
            {
                return index;
            }
        }

        return -1;
    }

    private sealed record ClientWorkspaceClient(
        long? ClientId,
        string ClientName,
        string ClientNumber,
        string Carrier,
        DateOnly? EffectiveDate,
        string Industry,
        string Province,
        string Advisor,
        string Status);

    private sealed record ClassComparisonData(
        IReadOnlyList<ClassRateSourceRow> Rates,
        IReadOnlyList<ClassCostSourceRow> Costs,
        int? TotalEligibleEmployees,
        ClassComparisonPreferenceState Preferences,
        IReadOnlyList<ClassComparisonNoteRow> Notes,
        IReadOnlyList<ClassComparisonHistoryRow> History)
    {
        public static ClassComparisonData Empty { get; } = new([], [], null, new ClassComparisonPreferenceState(), [], []);
    }

    private sealed record PlanFeatureRow(
        string Benefit,
        string Category,
        string FeatureName,
        string FeatureValue,
        string Coinsurance,
        string Deductible,
        string Maximum,
        string Frequency,
        string WaitingPeriod,
        string TerminationAge,
        string SourceSection,
        string ClassNumber,
        string ClassName,
        string SourcePage,
        string SourceDocument,
        string SourceDocumentPath,
        string SourceDocumentType,
        string EffectiveDateText,
        string Carrier,
        string SourceText,
        decimal? ConfidenceScore,
        DateTime? DateLoaded,
        long? PlanDesignFeatureId,
        IReadOnlyList<PlanBenefitRule> Rules);
}
