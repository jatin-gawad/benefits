using System.Text.Json;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class UserSettingsStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public UserSettings Load()
    {
        try
        {
            AppPaths.EnsureCreated();
            if (!File.Exists(AppPaths.SettingsPath))
            {
                return new UserSettings();
            }

            var json = File.ReadAllText(AppPaths.SettingsPath);
            return JsonSerializer.Deserialize<UserSettings>(json) ?? new UserSettings();
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return new UserSettings();
        }
    }

    public void Save(UserSettings settings)
    {
        try
        {
            AppPaths.EnsureCreated();
            File.WriteAllText(AppPaths.SettingsPath, JsonSerializer.Serialize(settings, JsonOptions));
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
        }
    }
}

