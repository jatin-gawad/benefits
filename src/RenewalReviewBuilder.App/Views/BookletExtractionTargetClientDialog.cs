using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Views;

internal sealed class BookletExtractionTargetClientDialog : Window
{
    private readonly ComboBox _clientComboBox;
    private readonly Button _confirmButton;
    private readonly TextBlock _destinationTextBlock;

    public BookletExtractionTargetClientDialog(
        IReadOnlyList<ClientRecord> clients,
        IEnumerable<string> bookletFileNames)
    {
        Title = "Confirm Booklet Destination";
        Width = 700;
        Height = 470;
        MinWidth = 620;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        Background = new SolidColorBrush(Color.FromRgb(248, 250, 252));

        var choices = clients
            .Where(client => client.ClientId > 0 && !string.IsNullOrWhiteSpace(client.ClientName))
            .OrderBy(client => client.ClientName, StringComparer.OrdinalIgnoreCase)
            .Select(client => new ClientChoice(client))
            .ToList();
        var files = bookletFileNames
            .Where(fileName => !string.IsNullOrWhiteSpace(fileName))
            .Select(fileName => $"- {fileName.Trim()}")
            .ToList();

        var root = new Grid { Margin = new Thickness(24) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        root.Children.Add(BenefitAdministrationEditDialog.Header(
            "Choose the destination client",
            "The selected client must match the client identified inside every booklet. A mismatch will create the workbook but block the database save."));

        var filesPanel = BenefitAdministrationEditDialog.FieldPanel("Selected Booklets", 1);
        filesPanel.Children.Add(new TextBlock
        {
            Text = files.Count == 0 ? "No booklet file names are available." : string.Join(Environment.NewLine, files),
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(71, 85, 105)),
            MaxHeight = 92
        });
        root.Children.Add(filesPanel);

        var clientPanel = BenefitAdministrationEditDialog.FieldPanel("Destination Client (required)", 2);
        _clientComboBox = new ComboBox
        {
            ItemsSource = choices,
            DisplayMemberPath = nameof(ClientChoice.DisplayName),
            IsEditable = true,
            IsTextSearchEnabled = true,
            StaysOpenOnEdit = true,
            Height = 40,
            MaxDropDownHeight = 260,
            Margin = new Thickness(0, 6, 0, 0)
        };
        _clientComboBox.SelectionChanged += (_, _) => UpdateSelection();
        clientPanel.Children.Add(_clientComboBox);
        root.Children.Add(clientPanel);

        _destinationTextBlock = new TextBlock
        {
            Text = "No destination selected. Nothing will be written to SQLite.",
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(153, 27, 27)),
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 2, 0, 8)
        };
        Grid.SetRow(_destinationTextBlock, 3);
        root.Children.Add(_destinationTextBlock);

        var actions = BenefitAdministrationEditDialog.ActionPanel(4);
        actions.Children.Add(new Button
        {
            Content = "Cancel",
            MinWidth = 90,
            Height = 38,
            IsCancel = true,
            Margin = new Thickness(0, 0, 10, 0)
        });
        _confirmButton = BenefitAdministrationEditDialog.PrimaryButton("Confirm Destination", 155);
        _confirmButton.IsDefault = true;
        _confirmButton.IsEnabled = false;
        _confirmButton.Click += (_, _) =>
        {
            if (SelectedClient is not null)
            {
                DialogResult = true;
            }
        };
        actions.Children.Add(_confirmButton);
        root.Children.Add(actions);

        Content = root;
        Loaded += (_, _) => _clientComboBox.Focus();
    }

    public ClientRecord? SelectedClient => (_clientComboBox.SelectedItem as ClientChoice)?.Client;

    private void UpdateSelection()
    {
        var selected = SelectedClient;
        _confirmButton.IsEnabled = selected is not null;
        _destinationTextBlock.Text = selected is null
            ? "No destination selected. Nothing will be written to SQLite."
            : $"Plan Design destination: {selected.ClientName} | Client {selected.ClientId}" +
              (string.IsNullOrWhiteSpace(selected.ClientNumber) ? "" : $" | Policy/Group {selected.ClientNumber}");
        _destinationTextBlock.Foreground = new SolidColorBrush(
            selected is null ? Color.FromRgb(153, 27, 27) : Color.FromRgb(0, 109, 109));
    }

    private sealed record ClientChoice(ClientRecord Client)
    {
        public string DisplayName =>
            $"{Client.ClientName} | Client {Client.ClientId}" +
            (string.IsNullOrWhiteSpace(Client.ClientNumber) ? "" : $" | {Client.ClientNumber}") +
            (string.IsNullOrWhiteSpace(Client.Carrier) ? "" : $" | {Client.Carrier}");
    }
}
