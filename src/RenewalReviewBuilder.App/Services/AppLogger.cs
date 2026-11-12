namespace RenewalReviewBuilder.App.Services;

public static class AppLogger
{
    public static void Info(string message)
    {
        Write("INFO", message);
    }

    public static void Error(Exception exception)
    {
        Write("ERROR", exception.ToString());
    }

    private static void Write(string level, string message)
    {
        try
        {
            AppPaths.EnsureCreated();
            var logPath = Path.Combine(AppPaths.LogsFolder, $"renewal-review-{DateTime.Today:yyyyMMdd}.log");
            File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}{Environment.NewLine}");
        }
        catch
        {
            // Logging should never block the user from creating a review pack.
        }
    }
}

