using RenewalReviewBuilder.App.Models.History;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

/// <summary>
/// Formats only canonical values already present on a history event for the shared details drawer.
/// </summary>
public static class ClientHistoryDrawerContentFormatter
{
    public static IReadOnlyList<string> Details(ClientHistoryEventReadModel historyEvent)
    {
        ArgumentNullException.ThrowIfNull(historyEvent);

        var details = new List<string>();
        Add(details, "Summary", historyEvent.Summary);
        Add(details, "Business rationale", historyEvent.BusinessRationale);
        Add(details, "Category", historyEvent.Category);
        Add(details, "Requested by", historyEvent.RequestedBy);
        Add(details, "Approved by", historyEvent.ApprovedBy);
        Add(details, "Notes", historyEvent.Notes);
        return details;
    }

    public static IReadOnlyList<string> Evidence(ClientHistoryEventReadModel historyEvent)
    {
        ArgumentNullException.ThrowIfNull(historyEvent);

        return historyEvent.Evidence.Select(evidence =>
        {
            var values = new List<string>();
            if (evidence.DocumentId.HasValue)
            {
                values.Add($"Document {evidence.DocumentId}");
            }

            var location = string.Join(" ", new[] { evidence.SheetName, evidence.CellRange }
                .Where(value => !string.IsNullOrWhiteSpace(value)));
            if (!string.IsNullOrWhiteSpace(location))
            {
                values.Add(location);
            }

            if (!string.IsNullOrWhiteSpace(evidence.ExtractedTextReference))
            {
                values.Add($"Source text: {evidence.ExtractedTextReference}");
            }

            values.Add(evidence.ValidationStatus.ToString());
            return string.Join("  •  ", values);
        }).ToList();
    }

    private static void Add(ICollection<string> target, string label, string value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            target.Add($"{label}: {value}");
        }
    }
}
