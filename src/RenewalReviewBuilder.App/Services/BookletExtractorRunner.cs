using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Threading;
using RenewalReviewBuilder.App.Connectors;
using RenewalReviewBuilder.App.Connectors.Pdf;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class BookletExtractorRunner
{
    internal const int IdentityCommandArgumentBudget = 5000;
    private const string ExtractorCommandFileName = "Run Booklet Plan Design Extractor.cmd";
    private const string RulesFileName = "booklet_plan_design_rules.json";
    private const string FlexibleBenefitSummariesFileName = "flexible-benefit-summaries.json";

    public async Task<BookletExtractionResult> RunAsync(
        BookletExtractionRequest request,
        IProgress<BookletExtractionProgress>? progress,
        CancellationToken cancellationToken)
    {
        var extractorCommandPath = ResolveExtractorCommandPath();
        if (!File.Exists(extractorCommandPath))
        {
            throw new FileNotFoundException("The tested booklet extractor was not found.", extractorCommandPath);
        }

        var rulesPath = Path.Combine(Path.GetDirectoryName(extractorCommandPath) ?? "", RulesFileName);
        if (!File.Exists(rulesPath))
        {
            throw new FileNotFoundException("The booklet extractor rules file was not found.", rulesPath);
        }

        var inputPaths = request.InputPaths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (inputPaths.Count == 0)
        {
            throw new InvalidOperationException("Add one or more booklet PDFs, or add a folder that contains booklet PDFs.");
        }

        Directory.CreateDirectory(request.OutputFolder);
        var masterPath = Path.Combine(request.OutputFolder, "Benefits_Master_Database.xlsx");
        var reviewPath = Path.Combine(request.OutputFolder, "Booklet_Extraction_Review.xlsx");
        var flexibleBenefitSummariesPath = await WriteFlexibleBenefitSummariesAsync(
            inputPaths,
            request.OutputFolder,
            cancellationToken);
        var totalPdfCount = CountPdfFiles(inputPaths);
        if (totalPdfCount == 0)
        {
            throw new InvalidOperationException("No PDF files were found in the selected booklet inputs.");
        }

        progress?.Report(new BookletExtractionProgress(
            $"Found {totalPdfCount} booklet PDF file(s). Starting extraction...",
            0,
            totalPdfCount,
            0,
            0));

        var output = new StringBuilder();
        var errors = new StringBuilder();
        using var process = new Process
        {
            StartInfo = CreateStartInfo(
                inputPaths,
                masterPath,
                reviewPath,
                extractorCommandPath,
                rulesPath,
                flexibleBenefitSummariesPath),
            EnableRaisingEvents = true
        };

        process.OutputDataReceived += (_, eventArgs) =>
        {
            if (eventArgs.Data is not null)
            {
                output.AppendLine(eventArgs.Data);
            }
        };
        process.ErrorDataReceived += (_, eventArgs) =>
        {
            if (eventArgs.Data is null)
            {
                return;
            }

            if (TryReadProgress(eventArgs.Data, out var status))
            {
                progress?.Report(status);
            }
            else
            {
                errors.AppendLine(eventArgs.Data);
            }
        };

        AppLogger.Info($"Starting booklet extraction for {totalPdfCount} PDF(s).");
        try
        {
            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            await process.WaitForExitAsync(cancellationToken);
        }
        catch (OperationCanceledException)
        {
            TryKill(process);
            throw new OperationCanceledException("Booklet extraction was cancelled.");
        }
        finally
        {
            TryDeleteFile(flexibleBenefitSummariesPath);
        }

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"Booklet extraction failed.{Environment.NewLine}{errors}{Environment.NewLine}{output}");
        }

        var summary = ParseSummary(output.ToString());
        var reviewCounts = CountReviewRows(reviewPath);
        return new BookletExtractionResult
        {
            OutputFolder = request.OutputFolder,
            MasterWorkbookPath = masterPath,
            ReviewWorkbookPath = reviewPath,
            TotalPdfsFound = summary.TotalPdfsFound > 0 ? summary.TotalPdfsFound : totalPdfCount,
            SuccessfullyProcessed = summary.SuccessfullyProcessed,
            ErrorCount = summary.ErrorCount,
            ReviewRequiredCount = reviewCounts.Required,
            ReviewRecommendedCount = reviewCounts.Recommended,
            BenefitAdministrationRecords = summary.BenefitAdministrationRecords,
            RawOutput = output.ToString()
        };
    }

    public async Task<IReadOnlyList<BookletClientIdentity>> DiscoverClientIdentitiesAsync(
        IEnumerable<string> inputPaths,
        CancellationToken cancellationToken)
    {
        var extractorCommandPath = ResolveExtractorCommandPath();
        if (!File.Exists(extractorCommandPath))
        {
            throw new FileNotFoundException("The tested booklet extractor was not found.", extractorCommandPath);
        }

        var rulesPath = Path.Combine(Path.GetDirectoryName(extractorCommandPath) ?? "", RulesFileName);
        if (!File.Exists(rulesPath))
        {
            throw new FileNotFoundException("The booklet extractor rules file was not found.", rulesPath);
        }

        var bookletPaths = inputPaths
            .Where(path => File.Exists(path) && Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (bookletPaths.Count == 0)
        {
            return [];
        }

        var identityInputs = CompactIdentityInputPaths(bookletPaths);
        var inputBatches = BuildIdentityInputBatches(identityInputs);
        var identities = new List<BookletClientIdentity>();
        AppLogger.Info(
            $"Identifying clients for {bookletPaths.Count} booklet PDF(s) using {identityInputs.Count} compact input(s) across {inputBatches.Count} command batch(es).");
        foreach (var inputBatch in inputBatches)
        {
            identities.AddRange(await DiscoverClientIdentityBatchAsync(
                inputBatch,
                extractorCommandPath,
                rulesPath,
                cancellationToken));
        }

        return identities
            .Where(identity => !string.IsNullOrWhiteSpace(identity.SourcePath))
            .GroupBy(identity => Path.GetFullPath(identity.SourcePath), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();
    }

    internal static List<string> CompactIdentityInputPaths(IReadOnlyList<string> bookletPaths)
    {
        var compacted = new List<string>();
        foreach (var group in bookletPaths
                     .Select(Path.GetFullPath)
                     .GroupBy(path => Path.GetDirectoryName(path) ?? "", StringComparer.OrdinalIgnoreCase)
                     .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            var requestedFiles = group.ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(group.Key) && Directory.Exists(group.Key))
            {
                var folderPdfFiles = Directory
                    .EnumerateFiles(group.Key, "*.pdf", SearchOption.AllDirectories)
                    .Select(Path.GetFullPath)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                if (folderPdfFiles.SetEquals(requestedFiles))
                {
                    compacted.Add(group.Key);
                    continue;
                }
            }

            compacted.AddRange(requestedFiles.OrderBy(path => path, StringComparer.OrdinalIgnoreCase));
        }

        return compacted;
    }

    internal static List<List<string>> BuildIdentityInputBatches(
        IReadOnlyList<string> inputPaths,
        int argumentBudget = IdentityCommandArgumentBudget)
    {
        var batches = new List<List<string>>();
        var current = new List<string>();
        var currentLength = 0;
        foreach (var inputPath in inputPaths)
        {
            var estimatedLength = inputPath.Length + 3;
            if (current.Count > 0 && currentLength + estimatedLength > argumentBudget)
            {
                batches.Add(current);
                current = [];
                currentLength = 0;
            }

            current.Add(inputPath);
            currentLength += estimatedLength;
        }

        if (current.Count > 0)
        {
            batches.Add(current);
        }

        return batches;
    }

    private static async Task<IReadOnlyList<BookletClientIdentity>> DiscoverClientIdentityBatchAsync(
        List<string> inputPaths,
        string extractorCommandPath,
        string rulesPath,
        CancellationToken cancellationToken)
    {
        var output = new StringBuilder();
        var errors = new StringBuilder();
        using var process = new Process
        {
            StartInfo = CreateIdentityStartInfo(inputPaths, extractorCommandPath, rulesPath),
            EnableRaisingEvents = true
        };
        process.OutputDataReceived += (_, eventArgs) =>
        {
            if (eventArgs.Data is not null)
            {
                output.AppendLine(eventArgs.Data);
            }
        };
        process.ErrorDataReceived += (_, eventArgs) =>
        {
            if (eventArgs.Data is not null)
            {
                errors.AppendLine(eventArgs.Data);
            }
        };

        try
        {
            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            await process.WaitForExitAsync(cancellationToken);
        }
        catch (OperationCanceledException)
        {
            TryKill(process);
            throw new OperationCanceledException("Booklet client identification was cancelled.");
        }

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"Booklet client identification failed.{Environment.NewLine}{errors}{Environment.NewLine}{output}");
        }

        return ParseClientIdentities(output.ToString());
    }

    public BookletPlanDesignWorkbookResult WritePlanDesignWorkbooksFromFeatures(BookletPlanDesignWorkbookRequest request)
    {
        return BookletPlanDesignWorkbookWriter.Write(request);
    }

    public static int CountPdfFiles(IEnumerable<string> inputPaths)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var inputPath in inputPaths.Where(path => !string.IsNullOrWhiteSpace(path)))
        {
            if (File.Exists(inputPath) && Path.GetExtension(inputPath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                seen.Add(Path.GetFullPath(inputPath));
                continue;
            }

            if (Directory.Exists(inputPath))
            {
                foreach (var pdfPath in Directory.EnumerateFiles(inputPath, "*.pdf", SearchOption.AllDirectories))
                {
                    seen.Add(Path.GetFullPath(pdfPath));
                }
            }
        }

        return seen.Count;
    }

    private static ProcessStartInfo CreateStartInfo(
        List<string> inputPaths,
        string masterPath,
        string reviewPath,
        string extractorCommandPath,
        string rulesPath,
        string flexibleBenefitSummariesPath)
    {
        var commandLineParts = new List<string>
        {
            Quote(extractorCommandPath),
            "--input"
        };
        commandLineParts.AddRange(inputPaths.Select(Quote));
        commandLineParts.AddRange([
            "--master",
            Quote(masterPath),
            "--review",
            Quote(reviewPath),
            "--rules",
            Quote(rulesPath),
            "--json",
            "--progress"
        ]);
        if (!string.IsNullOrWhiteSpace(flexibleBenefitSummariesPath))
        {
            commandLineParts.AddRange([
                "--flexible-benefit-summaries",
                Quote(flexibleBenefitSummariesPath)
            ]);
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe",
            Arguments = $"/d /s /c \"{string.Join(" ", commandLineParts)}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            WorkingDirectory = Path.GetDirectoryName(extractorCommandPath) ?? Environment.CurrentDirectory
        };
        return startInfo;
    }

    private static ProcessStartInfo CreateIdentityStartInfo(
        List<string> inputPaths,
        string extractorCommandPath,
        string rulesPath)
    {
        var commandLineParts = new List<string>
        {
            Quote(extractorCommandPath),
            "--input"
        };
        commandLineParts.AddRange(inputPaths.Select(Quote));
        commandLineParts.AddRange([
            "--rules",
            Quote(rulesPath),
            "--identity-only",
            "--json"
        ]);

        return new ProcessStartInfo
        {
            FileName = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe",
            Arguments = $"/d /s /c \"{string.Join(" ", commandLineParts)}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            WorkingDirectory = Path.GetDirectoryName(extractorCommandPath) ?? Environment.CurrentDirectory
        };
    }

    private static async Task<string> WriteFlexibleBenefitSummariesAsync(
        IEnumerable<string> inputPaths,
        string outputFolder,
        CancellationToken cancellationToken)
    {
        var summaries = new List<FlexibleBenefitSummary>();
        var connector = new NationalHealthClaimFlexibleBenefitsConnector();
        foreach (var pdfPath in EnumeratePotentialFlexibleBenefitPdfs(inputPaths))
        {
            try
            {
                var extraction = await connector.ExtractFlexibleBenefitsAsync(
                    new BenefitsConnectorRequest
                    {
                        InputPath = pdfPath,
                        SourceFolder = Path.GetDirectoryName(pdfPath) ?? "",
                        OutputFolder = outputFolder
                    },
                    cancellationToken);
                if (extraction.Plan.HcsaEnabled != true && extraction.Plan.LsaEnabled != true)
                {
                    continue;
                }

                summaries.Add(CreateFlexibleBenefitSummary(pdfPath, extraction));
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex);
            }
        }

        if (summaries.Count == 0)
        {
            return "";
        }

        var summaryPath = Path.Combine(outputFolder, FlexibleBenefitSummariesFileName);
        var content = JsonSerializer.Serialize(
            new FlexibleBenefitSummaryEnvelope { Documents = summaries },
            new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            });
        await File.WriteAllTextAsync(summaryPath, content, cancellationToken);
        return summaryPath;
    }

    private static IEnumerable<string> EnumeratePotentialFlexibleBenefitPdfs(IEnumerable<string> inputPaths)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var inputPath in inputPaths)
        {
            IEnumerable<string> candidates = File.Exists(inputPath)
                ? [inputPath]
                : Directory.Exists(inputPath)
                    ? Directory.EnumerateFiles(inputPath, "*.pdf", SearchOption.AllDirectories)
                    : [];
            foreach (var candidate in candidates)
            {
                var fileName = Path.GetFileName(candidate);
                if (!fileName.Contains("nhc", StringComparison.OrdinalIgnoreCase) &&
                    !fileName.Contains("national health", StringComparison.OrdinalIgnoreCase) &&
                    !fileName.Contains("lifehealth", StringComparison.OrdinalIgnoreCase) &&
                    !fileName.Contains("hcsa", StringComparison.OrdinalIgnoreCase) &&
                    !fileName.Contains("lsa", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var fullPath = Path.GetFullPath(candidate);
                if (seen.Add(fullPath))
                {
                    yield return fullPath;
                }
            }
        }
    }

    private static FlexibleBenefitSummary CreateFlexibleBenefitSummary(
        string sourcePath,
        FlexibleBenefitExtractionResult extraction)
    {
        var classAmounts = extraction.Classes
            .Select(item => new FlexibleBenefitClassAmount
            {
                ClassName = item.ClassName,
                SingleAmount = item.AnnualSingleAmount,
                CoupleAmount = item.AnnualCoupleAmount,
                FamilyAmount = item.AnnualFamilyAmount,
                SourcePageNumber = item.SourcePageNumber
            })
            .ToList();
        return new FlexibleBenefitSummary
        {
            SourceFile = Path.GetFileName(sourcePath),
            Carrier = extraction.Classification.Administrator,
            HcsaEnabled = extraction.Plan.HcsaEnabled == true,
            LsaEnabled = extraction.Plan.LsaEnabled == true,
            BenefitYear = extraction.Plan.BenefitYearType.ToString(),
            DefaultHcsaPercentage = extraction.Plan.DefaultHcsaPercentage,
            DefaultLsaPercentage = extraction.Plan.DefaultLsaPercentage,
            ClassAmounts = classAmounts
        };
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
        }
    }

    private static string ResolveExtractorCommandPath()
    {
        var packagedPath = Path.Combine(
            AppContext.BaseDirectory,
            "proposal_engine",
            "booklet_plan_design",
            ExtractorCommandFileName);
        if (File.Exists(packagedPath))
        {
            return packagedPath;
        }

        return packagedPath;
    }

    private static bool TryReadProgress(string line, out BookletExtractionProgress progress)
    {
        progress = new BookletExtractionProgress(line, 0, 0, 0, 0);
        if (!line.StartsWith("BOOKLET_PROGRESS|", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var parts = line.Split('|');
        if (parts.Length >= 3 && parts[1].Equals("TOTAL", StringComparison.OrdinalIgnoreCase) && int.TryParse(parts[2], out var total))
        {
            progress = new BookletExtractionProgress($"Found {total} booklet PDF file(s).", 0, total, 0, 0);
            return true;
        }

        if (parts.Length >= 5 &&
            int.TryParse(parts[2], out var index) &&
            int.TryParse(parts[3], out total))
        {
            var fileName = parts[4];
            var message = parts[1].ToUpperInvariant() switch
            {
                "START" => $"Processing {index} of {total} booklets: {fileName}",
                "SUCCESS" => $"Processed {index} of {total} booklets: {fileName}",
                "ERROR" => $"Error on {index} of {total} booklets: {fileName}",
                _ => $"Processing {index} of {total} booklets: {fileName}"
            };
            progress = new BookletExtractionProgress(message, index, total, 0, 0);
            return true;
        }

        return true;
    }

    private static BookletExtractionSummary ParseSummary(string output)
    {
        var start = output.IndexOf('{');
        var end = output.LastIndexOf('}');
        if (start < 0 || end <= start)
        {
            return new BookletExtractionSummary();
        }

        using var document = JsonDocument.Parse(output[start..(end + 1)]);
        var root = document.RootElement;
        var errorCount = root.TryGetProperty("errors", out var errors) && errors.ValueKind == JsonValueKind.Array
            ? errors.GetArrayLength()
            : 0;

        return new BookletExtractionSummary
        {
            TotalPdfsFound = ReadInt(root, "pdfs_found"),
            SuccessfullyProcessed = ReadInt(root, "pdfs_processed"),
            ErrorCount = errorCount,
            BenefitAdministrationRecords = ParseBenefitAdministrationRecords(root)
        };
    }

    private static List<BenefitAdministrationExtractionCandidate> ParseBenefitAdministrationRecords(JsonElement root)
    {
        if (!root.TryGetProperty("benefit_administration_records", out var records) ||
            records.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var parsed = new List<BenefitAdministrationExtractionCandidate>();
        foreach (var record in records.EnumerateArray())
        {
            var providers = new List<string>();
            if (record.TryGetProperty("service_provider_display_names", out var providerValues) &&
                providerValues.ValueKind == JsonValueKind.Array)
            {
                providers.AddRange(providerValues.EnumerateArray()
                    .Where(value => value.ValueKind == JsonValueKind.String)
                    .Select(value => value.GetString() ?? "")
                    .Where(value => !string.IsNullOrWhiteSpace(value)));
            }

            parsed.Add(new BenefitAdministrationExtractionCandidate
            {
                SourceFile = ReadString(record, "source_file"),
                ClientName = ReadString(record, "client_name"),
                BenefitDisplayName = ReadString(record, "benefit_display_name"),
                InsurerDisplayName = ReadString(record, "insurer_display_name"),
                ServiceProviderDisplayNames = providers,
                PolicyNumber = ReadString(record, "policy_number"),
                ProductName = ReadString(record, "product_name"),
                IsOptional = record.TryGetProperty("is_optional", out var isOptional) && isOptional.ValueKind == JsonValueKind.True,
                SourcePageNumber = record.TryGetProperty("source_page_number", out var pageNumber) && pageNumber.TryGetInt32(out var page) ? page : null,
                SourceText = ReadString(record, "source_text"),
                ExtractionConfidence = record.TryGetProperty("extraction_confidence", out var confidence) && confidence.TryGetDecimal(out var value) ? value : null
            });
        }

        return parsed;
    }

    private static IReadOnlyList<BookletClientIdentity> ParseClientIdentities(string output)
    {
        var start = output.IndexOf('{');
        var end = output.LastIndexOf('}');
        if (start < 0 || end <= start)
        {
            return [];
        }

        using var document = JsonDocument.Parse(output[start..(end + 1)]);
        if (!document.RootElement.TryGetProperty("identities", out var identities) ||
            identities.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var result = new List<BookletClientIdentity>();
        foreach (var identity in identities.EnumerateArray())
        {
            result.Add(new BookletClientIdentity
            {
                SourcePath = ReadString(identity, "sourcePath"),
                SourceFile = ReadString(identity, "sourceFile"),
                ClientName = ReadString(identity, "clientName"),
                DocumentClientName = ReadString(identity, "documentClientName"),
                FileNameClientName = ReadString(identity, "fileNameClientName"),
                Carrier = ReadString(identity, "carrier"),
                GroupNumber = ReadString(identity, "groupNumber"),
                PolicyNumber = ReadString(identity, "policyNumber"),
                ConfidenceScore = ReadDecimal(identity, "confidenceScore"),
                IdentitySource = ReadString(identity, "identitySource"),
                IdentityAgreement = ReadString(identity, "identityAgreement"),
                ErrorMessage = ReadString(identity, "errorMessage")
            });
        }

        return result;
    }

    internal static BookletReviewCounts CountReviewRows(string reviewPath)
    {
        try
        {
            if (!File.Exists(reviewPath))
            {
                return new BookletReviewCounts(0, 0);
            }

            using var workbook = new ClosedXML.Excel.XLWorkbook(reviewPath);
            var sheet = workbook.Worksheets.FirstOrDefault(worksheet =>
                worksheet.Name.Equals("Booklet_Extraction_Review", StringComparison.OrdinalIgnoreCase) ||
                worksheet.Name.Equals("Extraction_Review", StringComparison.OrdinalIgnoreCase));
            if (sheet is null)
            {
                return new BookletReviewCounts(0, 0);
            }

            var headerMap = sheet.Row(1)
                .CellsUsed()
                .ToDictionary(
                    cell => cell.GetString(),
                    cell => cell.Address.ColumnNumber,
                    StringComparer.OrdinalIgnoreCase);
            headerMap.TryGetValue("Review required", out var reviewRequiredColumn);
            headerMap.TryGetValue("ReviewRequired", out var reviewRequiredColumnAlt);
            headerMap.TryGetValue("ReviewFlag", out var reviewFlagColumn);
            headerMap.TryGetValue("Confidence", out var confidenceColumn);
            var reviewColumn = reviewRequiredColumn != 0 ? reviewRequiredColumn : reviewRequiredColumnAlt;

            var required = 0;
            var recommended = 0;
            foreach (var row in sheet.RowsUsed().Skip(1))
            {
                var reviewValue = reviewColumn == 0 ? "" : row.Cell(reviewColumn).GetString();
                var flagValue = reviewFlagColumn == 0 ? "" : row.Cell(reviewFlagColumn).GetString();
                var confidenceValue = confidenceColumn == 0 ? "" : row.Cell(confidenceColumn).GetString();
                if (RequiresBlockingReview(reviewValue, flagValue, confidenceValue))
                {
                    required++;
                }
                else if (RecommendsReview(flagValue, confidenceValue))
                {
                    recommended++;
                }
            }

            return new BookletReviewCounts(required, recommended);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return new BookletReviewCounts(0, 0);
        }
    }

    private static bool RequiresBlockingReview(string reviewValue, string flagValue, string confidenceValue)
    {
        if (IsAffirmativeReviewValue(reviewValue) ||
            confidenceValue.Equals("Low", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var flag = flagValue.Trim();
        if (string.IsNullOrWhiteSpace(flag) || IsNegativeReviewValue(flag))
        {
            return false;
        }

        return flag.Equals("Yes", StringComparison.OrdinalIgnoreCase) ||
               flag.Equals("True", StringComparison.OrdinalIgnoreCase) ||
               flag.Contains("required", StringComparison.OrdinalIgnoreCase) ||
               flag.Contains("needs review", StringComparison.OrdinalIgnoreCase) ||
               flag.Contains("low confidence", StringComparison.OrdinalIgnoreCase);
    }

    private static bool RecommendsReview(string flagValue, string confidenceValue)
    {
        if (confidenceValue.Equals("Medium", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var flag = flagValue.Trim();
        return !IsNegativeReviewValue(flag) &&
               (flag.Equals("Review", StringComparison.OrdinalIgnoreCase) ||
                flag.Contains("recommend", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsAffirmativeReviewValue(string value) =>
        value.Equals("Yes", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("True", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("Required", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("Needs Review", StringComparison.OrdinalIgnoreCase);

    private static bool IsNegativeReviewValue(string value) =>
        string.IsNullOrWhiteSpace(value) ||
        value.Equals("No", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("False", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("None", StringComparison.OrdinalIgnoreCase);

    private static int ReadInt(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) && property.TryGetInt32(out var value)
            ? value
            : 0;
    }

    private static string ReadString(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString() ?? ""
            : "";
    }

    private static decimal ReadDecimal(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) &&
               property.ValueKind == JsonValueKind.Number &&
               property.TryGetDecimal(out var value)
            ? value
            : 0m;
    }

    private static string Quote(string value)
    {
        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }

    private static void TryKill(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
        }
    }
}

public sealed class BookletExtractionRequest
{
    public List<string> InputPaths { get; init; } = [];
    public string OutputFolder { get; init; } = "";
}

public sealed class BookletExtractionResult
{
    public string OutputFolder { get; init; } = "";
    public string MasterWorkbookPath { get; init; } = "";
    public string ReviewWorkbookPath { get; init; } = "";
    public int TotalPdfsFound { get; init; }
    public int SuccessfullyProcessed { get; init; }
    public int ErrorCount { get; init; }
    public int ReviewRequiredCount { get; init; }
    public int ReviewRecommendedCount { get; init; }
    public List<BenefitAdministrationExtractionCandidate> BenefitAdministrationRecords { get; init; } = [];
    public string RawOutput { get; init; } = "";
}

internal sealed record BookletReviewCounts(int Required, int Recommended);

public sealed record BookletExtractionProgress(
    string Message,
    int Current,
    int Total,
    int SuccessfullyProcessed,
    int ErrorCount);

internal sealed class BookletExtractionSummary
{
    public int TotalPdfsFound { get; init; }
    public int SuccessfullyProcessed { get; init; }
    public int ErrorCount { get; init; }
    public List<BenefitAdministrationExtractionCandidate> BenefitAdministrationRecords { get; init; } = [];
}

internal sealed class FlexibleBenefitSummaryEnvelope
{
    public List<FlexibleBenefitSummary> Documents { get; init; } = [];
}

internal sealed class FlexibleBenefitSummary
{
    public string SourceFile { get; init; } = "";
    public string Carrier { get; init; } = "National HealthClaim";
    public bool HcsaEnabled { get; init; }
    public bool LsaEnabled { get; init; }
    public string BenefitYear { get; init; } = "";
    public decimal? DefaultHcsaPercentage { get; init; }
    public decimal? DefaultLsaPercentage { get; init; }
    public List<FlexibleBenefitClassAmount> ClassAmounts { get; init; } = [];
}

internal sealed class FlexibleBenefitClassAmount
{
    public string ClassName { get; init; } = "";
    public decimal? SingleAmount { get; init; }
    public decimal? CoupleAmount { get; init; }
    public decimal? FamilyAmount { get; init; }
    public int? SourcePageNumber { get; init; }
}
