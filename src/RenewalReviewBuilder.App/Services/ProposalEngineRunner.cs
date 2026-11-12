using System.Diagnostics;
using System.IO;
using System.Text;

namespace RenewalReviewBuilder.App.Services;

public sealed class ProposalEngineRunner
{
    private const int MaxFocusedCarrierPackages = 5;
    private static readonly HashSet<string> FocusedPdfExtensions = new(StringComparer.OrdinalIgnoreCase) { ".pdf" };
    private static readonly HashSet<string> FocusedExcelExtensions = new(StringComparer.OrdinalIgnoreCase) { ".xlsx", ".xlsm" };
    private static readonly HashSet<string> FocusedWordExtensions = new(StringComparer.OrdinalIgnoreCase) { ".docx" };
    private static readonly HashSet<string> KnownUnsupportedFocusedExtensions = new(StringComparer.OrdinalIgnoreCase) { ".xls", ".csv" };
    private static readonly string[] NonProposalPdfNameKeywords =
    [
        "brochure",
        "booklet",
        "flyer",
        "marketing",
        "rxfood"
    ];

    public async Task<ProposalEngineRunResult> RunComparisonAsync(
        IReadOnlyList<string> proposalPaths,
        string? clientId,
        string? bookletPath,
        string outputFolder,
        string databasePath,
        IProgress<string>? progress = null,
        string? clientName = null)
    {
        var focusedInputs = ResolveFocusedComparisonInputs(proposalPaths);
        if (!focusedInputs.HasRunnableInputs)
        {
            return ProposalEngineRunResult.Fail(
                "No supported proposal files were found. For this comparison workflow, use PDFs or an Excel/Word quote package in the same carrier folder.",
                outputFolder,
                databasePath);
        }

        var arguments = new List<string> { "-m", "proposal_engine.quote_test_extractor" };
        foreach (var package in focusedInputs.CarrierPackages)
        {
            arguments.Add("--carrier-package-comparison");
            arguments.AddRange(package.PdfPaths);
            AddOptional(arguments, package.ExcelPath);
            AddOptional(arguments, package.WordPath);
        }

        AddOptional(arguments, "--output-folder", outputFolder);
        AddOptional(arguments, "--database-path", databasePath);
        AddOptional(arguments, "--client-id", clientId);
        AddOptional(arguments, "--client-name", clientName);
        AddOptional(arguments, "--booklet-path", bookletPath);

        var notes = BuildFocusedComparisonNotes(focusedInputs, bookletPath);
        progress?.Report("Running focused carrier proposal comparison...");
        return await RunPythonAsync(arguments, outputFolder, databasePath, progress, notes);
    }

    public async Task<ProposalEngineRunResult> RunTestAsync(
        string proposalFile,
        string outputFolder,
        string databasePath,
        IProgress<string>? progress = null)
    {
        var arguments = new List<string>
        {
            "-m",
            "proposal_engine.test_runner",
            proposalFile,
            "--output-folder",
            outputFolder,
            "--database-path",
            databasePath
        };

        progress?.Report("Testing one carrier proposal...");
        return await RunPythonAsync(arguments, outputFolder, databasePath, progress);
    }

    private static async Task<ProposalEngineRunResult> RunPythonAsync(
        IReadOnlyList<string> arguments,
        string outputFolder,
        string databasePath,
        IProgress<string>? progress,
        IReadOnlyList<string>? preRunNotes = null)
    {
        var projectRoot = FindProjectRoot();
        if (projectRoot is null)
        {
            return ProposalEngineRunResult.Fail("Could not find the proposal_engine folder in the desktop project.", outputFolder, databasePath);
        }

        var pythonPath = FindPythonPath();
        if (string.IsNullOrWhiteSpace(pythonPath))
        {
            return ProposalEngineRunResult.Fail("Python was not found. Install Python or run from the Codex desktop environment.", outputFolder, databasePath);
        }

        Directory.CreateDirectory(outputFolder);
        Directory.CreateDirectory(Path.GetDirectoryName(databasePath) ?? outputFolder);

        var startInfo = new ProcessStartInfo
        {
            FileName = pythonPath,
            WorkingDirectory = projectRoot,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        var existingPythonPath = Environment.GetEnvironmentVariable("PYTHONPATH");
        startInfo.Environment["PYTHONPATH"] = string.IsNullOrWhiteSpace(existingPythonPath)
            ? projectRoot
            : projectRoot + Path.PathSeparator + existingPythonPath;
        startInfo.Environment["PYTHONIOENCODING"] = "utf-8";

        using var process = new Process { StartInfo = startInfo };
        process.Start();
        var outputTask = process.StandardOutput.ReadToEndAsync();
        var errorTask = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();

        var output = await outputTask;
        var error = await errorTask;
        if (!string.IsNullOrWhiteSpace(output))
        {
            progress?.Report("Proposal engine finished.");
        }

        return new ProposalEngineRunResult
        {
            Succeeded = process.ExitCode == 0,
            ExitCode = process.ExitCode,
            Output = CombineOutput(preRunNotes, output),
            Error = error.Trim(),
            OutputFolder = outputFolder,
            DatabasePath = databasePath
        };
    }

    private static FocusedComparisonInputs ResolveFocusedComparisonInputs(IReadOnlyList<string> proposalPaths)
    {
        var allFiles = ExpandProposalInputFiles(proposalPaths)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var packages = new List<CarrierFilePackage>();
        var skipped = new List<string>();

        foreach (var group in allFiles.GroupBy(path => Path.GetDirectoryName(path) ?? "", StringComparer.OrdinalIgnoreCase))
        {
            var excelFiles = group
                .Where(path => FocusedExcelExtensions.Contains(Path.GetExtension(path)))
                .OrderByDescending(GetFocusedExcelScore)
                .ThenBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var wordFiles = group
                .Where(path => FocusedWordExtensions.Contains(Path.GetExtension(path)))
                .OrderByDescending(GetFocusedWordScore)
                .ThenBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var pdfFiles = group
                .Where(path => FocusedPdfExtensions.Contains(Path.GetExtension(path)))
                .Where(path =>
                {
                    if (IsLikelySupplementalPdf(path))
                    {
                        skipped.Add($"{path} (skipped because it looks like a supplemental brochure, not the carrier quote)");
                        return false;
                    }

                    return true;
                })
                .OrderByDescending(GetFocusedPdfScore)
                .ThenBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

            string? excelPath = null;
            string? wordPath = null;
            if (excelFiles.Count > 0 && wordFiles.Count > 0)
            {
                excelPath = excelFiles[0];
                wordPath = wordFiles[0];
            }

            foreach (var path in excelFiles.Skip(excelPath is null ? 0 : 1))
            {
                skipped.Add(excelPath is null
                    ? $"{path} (Excel quote skipped because no matching Word document was found in the same folder)"
                    : $"{path} (Excel quote skipped because this carrier folder already has a selected Excel quote workbook)");
            }

            foreach (var path in wordFiles.Skip(wordPath is null ? 0 : 1))
            {
                skipped.Add(wordPath is null
                    ? $"{path} (Word quote skipped because no matching Excel workbook was found in the same folder)"
                    : $"{path} (Word quote skipped because this carrier folder already has a selected Word proposal)");
            }

            foreach (var path in group.Where(path => KnownUnsupportedFocusedExtensions.Contains(Path.GetExtension(path))))
            {
                skipped.Add($"{path} (save as .xlsx, or pair an .xlsx workbook with a .docx proposal)");
            }

            if (pdfFiles.Count > 0 || (excelPath is not null && wordPath is not null))
            {
                packages.Add(new CarrierFilePackage(pdfFiles, excelPath, wordPath));
            }
        }

        var selectedPackages = packages.Take(MaxFocusedCarrierPackages).ToList();
        foreach (var package in packages.Skip(MaxFocusedCarrierPackages))
        {
            skipped.Add($"{package.DisplayName} (skipped because this comparison workbook supports up to {MaxFocusedCarrierPackages} carrier packages)");
        }

        return new FocusedComparisonInputs(selectedPackages, skipped);
    }

    private static IEnumerable<string> ExpandProposalInputFiles(IReadOnlyList<string> proposalPaths)
    {
        foreach (var proposalPath in proposalPaths.Where(path => !string.IsNullOrWhiteSpace(path)))
        {
            if (Directory.Exists(proposalPath))
            {
                foreach (var filePath in Directory.EnumerateFiles(proposalPath, "*.*", SearchOption.AllDirectories)
                    .Where(IsFocusedComparisonCandidate))
                {
                    yield return filePath;
                }

                continue;
            }

            if (File.Exists(proposalPath) && IsFocusedComparisonCandidate(proposalPath))
            {
                yield return proposalPath;
            }
        }
    }

    private static bool IsFocusedComparisonCandidate(string path)
    {
        var extension = Path.GetExtension(path);
        return FocusedPdfExtensions.Contains(extension)
            || FocusedExcelExtensions.Contains(extension)
            || FocusedWordExtensions.Contains(extension)
            || KnownUnsupportedFocusedExtensions.Contains(extension);
    }

    private static bool IsLikelySupplementalPdf(string path)
    {
        var fileName = Path.GetFileNameWithoutExtension(path);
        return NonProposalPdfNameKeywords.Any(keyword =>
            fileName.Contains(keyword, StringComparison.OrdinalIgnoreCase));
    }

    private static int GetFocusedPdfScore(string path)
    {
        return GetProposalFileScore(path, isPdf: true);
    }

    private static int GetFocusedExcelScore(string path)
    {
        return GetProposalFileScore(path, isPdf: false);
    }

    private static int GetFocusedWordScore(string path)
    {
        return GetProposalFileScore(path, isPdf: false);
    }

    private static int GetProposalFileScore(string path, bool isPdf)
    {
        var name = Path.GetFileNameWithoutExtension(path);
        var score = 0;
        if (name.Contains("client proposal", StringComparison.OrdinalIgnoreCase))
        {
            score += 100;
        }

        if (name.Contains("quote proposal", StringComparison.OrdinalIgnoreCase))
        {
            score += 100;
        }

        if (name.Contains("proposal", StringComparison.OrdinalIgnoreCase))
        {
            score += 50;
        }

        if (name.Contains("quote", StringComparison.OrdinalIgnoreCase))
        {
            score += 40;
        }

        if (name.Contains("final tmc", StringComparison.OrdinalIgnoreCase))
        {
            score += 40;
        }

        if (name.Contains("assureplus", StringComparison.OrdinalIgnoreCase))
        {
            score += 20;
        }

        if (isPdf && name.Contains("package", StringComparison.OrdinalIgnoreCase))
        {
            score += 10;
        }

        if (name.Contains("advisor", StringComparison.OrdinalIgnoreCase))
        {
            score -= 100;
        }

        if (name.Contains("healthconnector", StringComparison.OrdinalIgnoreCase))
        {
            score -= 100;
        }

        if (name.Contains("brochure", StringComparison.OrdinalIgnoreCase))
        {
            score -= 100;
        }

        return score;
    }

    private static IReadOnlyList<string> BuildFocusedComparisonNotes(FocusedComparisonInputs inputs, string? bookletPath)
    {
        var notes = new List<string>
        {
            "Focused comparison extractor used: proposal_engine.quote_test_extractor.",
            $"Carrier packages: {inputs.CarrierPackages.Count}.",
            "Files in the same carrier folder are read together as one carrier column.",
            "Persistence target: shared desktop SQLite after workbook generation."
        };

        if (!string.IsNullOrWhiteSpace(bookletPath))
        {
            notes.Add("Benefits booklet validation: structured plan design facts will be compared and saved with source evidence.");
        }

        if (inputs.SkippedFiles.Count > 0)
        {
            notes.Add("Skipped files:");
            notes.AddRange(inputs.SkippedFiles.Select(path => $"- {path}"));
        }

        return notes;
    }

    private static string CombineOutput(IReadOnlyList<string>? preRunNotes, string output)
    {
        var lines = new List<string>();
        if (preRunNotes is not null)
        {
            lines.AddRange(preRunNotes.Where(note => !string.IsNullOrWhiteSpace(note)));
        }

        if (!string.IsNullOrWhiteSpace(output))
        {
            lines.Add(output.Trim());
        }

        return string.Join(Environment.NewLine, lines).Trim();
    }

    private static void AddOptional(List<string> arguments, string name, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        arguments.Add(name);
        arguments.Add(value);
    }

    private static void AddOptional(List<string> arguments, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            arguments.Add(value);
        }
    }

    public static string? FindProjectRoot()
    {
        var candidates = new List<string>();
        foreach (var start in new[] { AppContext.BaseDirectory, Directory.GetCurrentDirectory() })
        {
            var directory = new DirectoryInfo(start);
            while (directory is not null)
            {
                if (Directory.Exists(Path.Combine(directory.FullName, "proposal_engine")))
                {
                    candidates.Add(directory.FullName);
                }

                directory = directory.Parent;
            }
        }

        var sourceRoot = candidates.FirstOrDefault(path => File.Exists(Path.Combine(path, "RenewalReviewBuilder.sln")));
        if (!string.IsNullOrWhiteSpace(sourceRoot))
        {
            return sourceRoot;
        }

        if (candidates.Count > 0)
        {
            return candidates[0];
        }

        var desktopRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "Build Renewal");
        return Directory.Exists(Path.Combine(desktopRoot, "proposal_engine")) ? desktopRoot : null;
    }

    private static string? FindPythonPath()
    {
        var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var bundledPython = Path.Combine(
            userProfile,
            ".cache",
            "codex-runtimes",
            "codex-primary-runtime",
            "dependencies",
            "python",
            "python.exe");

        if (File.Exists(bundledPython))
        {
            return bundledPython;
        }

        return "python";
    }
}

public sealed record CarrierFilePackage(IReadOnlyList<string> PdfPaths, string? ExcelPath, string? WordPath)
{
    public string DisplayName
    {
        get
        {
            var paths = new List<string>();
            paths.AddRange(PdfPaths);
            if (!string.IsNullOrWhiteSpace(ExcelPath))
            {
                paths.Add(ExcelPath);
            }

            if (!string.IsNullOrWhiteSpace(WordPath))
            {
                paths.Add(WordPath);
            }

            return paths.Count == 0
                ? "Empty carrier package"
                : Path.GetDirectoryName(paths[0]) ?? paths[0];
        }
    }
}

public sealed record FocusedComparisonInputs(
    IReadOnlyList<CarrierFilePackage> CarrierPackages,
    IReadOnlyList<string> SkippedFiles)
{
    public bool HasRunnableInputs => CarrierPackages.Count > 0;
}

public sealed class ProposalEngineRunResult
{
    public bool Succeeded { get; init; }
    public int ExitCode { get; init; }
    public string Output { get; init; } = "";
    public string Error { get; init; } = "";
    public string OutputFolder { get; init; } = "";
    public string DatabasePath { get; init; } = "";

    public string Message
    {
        get
        {
            var text = string.Join(
                Environment.NewLine,
                new[] { Output, Error }.Where(value => !string.IsNullOrWhiteSpace(value)));
            return string.IsNullOrWhiteSpace(text) ? "Proposal engine finished with no console output." : text;
        }
    }

    public static ProposalEngineRunResult Fail(string message, string outputFolder, string databasePath)
    {
        return new ProposalEngineRunResult
        {
            Succeeded = false,
            ExitCode = -1,
            Error = message,
            OutputFolder = outputFolder,
            DatabasePath = databasePath
        };
    }
}
