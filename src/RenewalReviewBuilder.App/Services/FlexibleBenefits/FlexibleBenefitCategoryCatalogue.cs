using System.Text.Json;
using System.Text.Json.Serialization;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.FlexibleBenefits;

public sealed class FlexibleBenefitCategoryCatalogue
{
    private readonly string _cataloguePath;

    public FlexibleBenefitCategoryCatalogue(string? cataloguePath = null)
    {
        _cataloguePath = string.IsNullOrWhiteSpace(cataloguePath)
            ? Path.Combine(AppContext.BaseDirectory, "config", "flexible_benefits", "lsa_categories.json")
            : Path.GetFullPath(cataloguePath);
    }

    public IReadOnlyList<FlexibleBenefitCategoryDefinition> Load()
    {
        if (!File.Exists(_cataloguePath))
        {
            throw new FileNotFoundException("The flexible-benefit category catalogue is missing.", _cataloguePath);
        }

        var options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        };
        options.Converters.Add(new JsonStringEnumConverter());
        var definitions = JsonSerializer.Deserialize<List<FlexibleBenefitCategoryDefinition>>(
                              File.ReadAllText(_cataloguePath),
                              options) ?? [];
        var duplicates = definitions
            .GroupBy(item => item.CategoryCode, StringComparer.OrdinalIgnoreCase)
            .Where(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1)
            .Select(group => group.Key)
            .ToList();
        if (duplicates.Count > 0)
        {
            throw new InvalidDataException("The flexible-benefit category catalogue contains missing or duplicate category codes.");
        }

        return definitions;
    }
}
