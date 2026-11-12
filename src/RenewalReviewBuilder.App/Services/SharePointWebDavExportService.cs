using System.Threading;

namespace RenewalReviewBuilder.App.Services;

public static class SharePointWebDavPathPolicy
{
    private const string SharePointWebDavMarker = ".sharepoint.com@SSL\\DavWWWRoot\\";

    public static bool IsDirectSharePointWebDavPath(string? path)
    {
        var candidate = path?.Trim();
        return !string.IsNullOrWhiteSpace(candidate) &&
               candidate.StartsWith(@"\\", StringComparison.Ordinal) &&
               candidate.Contains(SharePointWebDavMarker, StringComparison.OrdinalIgnoreCase);
    }

    public static bool CanEnumeratePresentationRoot(string? path) =>
        !IsDirectSharePointWebDavPath(path);

    public static string CreateLocalStagingFolder(
        string documentsFolder,
        DateTime createdAt,
        string? runId = null)
    {
        if (string.IsNullOrWhiteSpace(documentsFolder))
        {
            throw new ArgumentException("A local Documents folder is required.", nameof(documentsFolder));
        }

        var token = string.IsNullOrWhiteSpace(runId)
            ? Guid.NewGuid().ToString("N")[..8]
            : new string(runId.Where(char.IsLetterOrDigit).ToArray());
        if (string.IsNullOrWhiteSpace(token))
        {
            token = Guid.NewGuid().ToString("N")[..8];
        }

        return Path.Combine(
            documentsFolder,
            "Renewal Review Builder",
            "SharePoint Staging",
            $"Renewal Review {createdAt:yyyyMMdd_HHmmss}_{token}");
    }
}

public sealed record SharePointPublishResult(
    string DestinationFolder,
    int PublishedFileCount);

public sealed class SharePointStagedExportPublisher
{
    private readonly TimeSpan _directoryTimeout;
    private readonly TimeSpan _fileTimeout;
    private readonly int _maxAttempts;

    public SharePointStagedExportPublisher(
        TimeSpan? directoryTimeout = null,
        TimeSpan? fileTimeout = null,
        int maxAttempts = 2)
    {
        _directoryTimeout = directoryTimeout ?? TimeSpan.FromSeconds(30);
        _fileTimeout = fileTimeout ?? TimeSpan.FromMinutes(2);
        _maxAttempts = Math.Max(1, maxAttempts);
    }

    public async Task<SharePointPublishResult> PublishAsync(
        string stagingFolder,
        string destinationRoot,
        IProgress<string>? progress = null)
    {
        if (!Directory.Exists(stagingFolder))
        {
            throw new DirectoryNotFoundException($"The local staging folder was not found: {stagingFolder}");
        }

        if (string.IsNullOrWhiteSpace(destinationRoot))
        {
            throw new ArgumentException("A SharePoint destination folder is required.", nameof(destinationRoot));
        }

        var stagingName = Path.GetFileName(
            stagingFolder.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        var destinationFolder = Path.Combine(destinationRoot.Trim(), stagingName);
        await CreateDirectoryWithTimeoutAsync(destinationFolder);

        var files = Directory
            .EnumerateFiles(stagingFolder, "*", SearchOption.AllDirectories)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        for (var index = 0; index < files.Count; index++)
        {
            var sourcePath = files[index];
            var relativePath = Path.GetRelativePath(stagingFolder, sourcePath);
            var destinationPath = Path.Combine(destinationFolder, relativePath);
            var destinationDirectory = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrWhiteSpace(destinationDirectory))
            {
                await CreateDirectoryWithTimeoutAsync(destinationDirectory);
            }

            progress?.Report($"Publishing file {index + 1} of {files.Count} to SharePoint...");
            await CopyFileWithRetryAsync(sourcePath, destinationPath);
        }

        return new SharePointPublishResult(destinationFolder, files.Count);
    }

    public static string RebasePublishedPath(
        string? sourcePath,
        string stagingFolder,
        string destinationFolder)
    {
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            return "";
        }

        var relativePath = Path.GetRelativePath(stagingFolder, sourcePath);
        return Path.Combine(destinationFolder, relativePath);
    }

    private async Task CreateDirectoryWithTimeoutAsync(string path)
    {
        try
        {
            await Task.Run(() => Directory.CreateDirectory(path)).WaitAsync(_directoryTimeout);
        }
        catch (TimeoutException exception)
        {
            throw new TimeoutException(
                $"SharePoint did not respond while creating the destination folder within {_directoryTimeout.TotalSeconds:0} seconds.",
                exception);
        }
    }

    private async Task CopyFileWithRetryAsync(string sourcePath, string destinationPath)
    {
        Exception? lastError = null;
        for (var attempt = 1; attempt <= _maxAttempts; attempt++)
        {
            try
            {
                using var timeout = new CancellationTokenSource(_fileTimeout);
                await Task.Run(
                        async () =>
                        {
                            await using var source = new FileStream(
                                sourcePath,
                                FileMode.Open,
                                FileAccess.Read,
                                FileShare.Read,
                                128 * 1024,
                                FileOptions.Asynchronous | FileOptions.SequentialScan);
                            await using var destination = new FileStream(
                                destinationPath,
                                FileMode.Create,
                                FileAccess.Write,
                                FileShare.None,
                                128 * 1024,
                                FileOptions.Asynchronous);
                            await source.CopyToAsync(destination, 128 * 1024, timeout.Token);
                            await destination.FlushAsync(timeout.Token);

                            if (destination.Length != source.Length)
                            {
                                throw new IOException(
                                    $"SharePoint copy verification failed for {Path.GetFileName(sourcePath)}.");
                            }
                        },
                        timeout.Token)
                    .WaitAsync(timeout.Token);
                return;
            }
            catch (OperationCanceledException exception)
            {
                throw new TimeoutException(
                    $"SharePoint did not finish saving {Path.GetFileName(sourcePath)} within {_fileTimeout.TotalSeconds:0} seconds. The local safety copy was retained.",
                    exception);
            }
            catch (IOException exception) when (attempt < _maxAttempts)
            {
                lastError = exception;
                await Task.Delay(TimeSpan.FromSeconds(attempt));
            }
            catch (UnauthorizedAccessException exception) when (attempt < _maxAttempts)
            {
                lastError = exception;
                await Task.Delay(TimeSpan.FromSeconds(attempt));
            }
        }

        throw new IOException(
            $"SharePoint could not save {Path.GetFileName(sourcePath)} after {_maxAttempts} attempt(s). The local safety copy was retained.",
            lastError);
    }
}
