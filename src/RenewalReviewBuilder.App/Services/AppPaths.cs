namespace RenewalReviewBuilder.App.Services;

public static class AppPaths
{
    public static string AppDataFolder
    {
        get
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(localAppData, "Westland", "RenewalReviewBuilder");
        }
    }

    public static string LogsFolder => Path.Combine(AppDataFolder, "Logs");

    public static string DataFolder => Path.Combine(AppDataFolder, "data");

    public static string BenefitsIntelligenceDatabasePath => Path.Combine(DataFolder, "Benefits_Intelligence.db");

    public static string SettingsPath => Path.Combine(AppDataFolder, "settings.json");

    public static void EnsureCreated()
    {
        Directory.CreateDirectory(AppDataFolder);
        Directory.CreateDirectory(LogsFolder);
        Directory.CreateDirectory(DataFolder);
    }
}
