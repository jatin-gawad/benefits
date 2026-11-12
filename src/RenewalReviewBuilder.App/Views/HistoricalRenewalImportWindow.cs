using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Views;

/// <summary>
/// Preview-first import for a completed older renewal. It deliberately uses the selected database
/// client and the dedicated historical persistence path so current client metadata is never upserted.
/// </summary>
public sealed class HistoricalRenewalImportWindow : Window
{
    private readonly long _clientId;
    private readonly DatabaseService _database = new();
    private readonly ReviewPackBuilder _reviewPackBuilder;
    private readonly StackPanel _stage = new() { Margin = new Thickness(22, 16, 22, 16) };
    private readonly StackPanel _stepStrip = new() { Orientation = Orientation.Horizontal, Margin = new Thickness(22, 14, 22, 0) };
    private readonly TextBlock _protectionBanner = new() { FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(Color.FromRgb(18, 100, 60)) };
    private readonly TextBlock _status = new() { Foreground = new SolidColorBrush(Color.FromRgb(86, 103, 130)), VerticalAlignment = VerticalAlignment.Center };
    private readonly Button _back = new() { Content = "Back", MinWidth = 92, Padding = new Thickness(14, 7, 14, 7), Margin = new Thickness(0, 0, 8, 0) };
    private readonly Button _next = new() { Content = "Preview sources", MinWidth = 150, Padding = new Thickness(14, 7, 14, 7) };
    private readonly TextBox _pdfPath = ReadOnlyInput();
    private readonly TextBox _workbookPath = ReadOnlyInput();
    private readonly TextBox _carrier = Input();
    private readonly TextBox _effectiveDate = Input();
    private readonly TextBox _renewalYear = Input();
    private readonly TextBox _requestedIncrease = Input();
    private readonly TextBox _finalAcceptedIncrease = Input();
    private readonly TextBox _annualSavings = Input();
    private readonly CheckBox _confirm = new()
    {
        Content = "I confirm that only the historical renewal records shown above will be added.",
        Margin = new Thickness(0, 16, 0, 0)
    };
    private ClientRecord? _client;
    private HistoricalRenewalImportRequest? _request;
    private HistoricalRenewalImportPreview? _preview;
    private string _extractionSummary = "";
    private int _step;
    private bool _isBusy;

    public HistoricalRenewalImportWindow(long clientId)
    {
        _clientId = clientId;
        _reviewPackBuilder = new ReviewPackBuilder(_database);
        Title = "Add Historical Renewal";
        Width = 960;
        Height = 760;
        MinWidth = 760;
        MinHeight = 620;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = new SolidColorBrush(Color.FromRgb(246, 248, 252));

        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var header = BuildHeader();
        Grid.SetRow(header, 0);
        root.Children.Add(header);
        Grid.SetRow(_stepStrip, 1);
        root.Children.Add(_stepStrip);
        var scroller = new ScrollViewer { Content = _stage, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
        Grid.SetRow(scroller, 2);
        root.Children.Add(scroller);
        var footer = BuildFooter();
        Grid.SetRow(footer, 3);
        root.Children.Add(footer);
        Content = root;

        _back.Click += Back_Click;
        _next.Click += Next_Click;
        Loaded += Window_Loaded;
        BuildStepStrip();
        RenderStep();
    }

    public bool ImportSucceeded { get; private set; }

    private Border BuildHeader()
    {
        var title = new TextBlock
        {
            Text = "Add an older renewal without changing the current renewal",
            FontSize = 20,
            FontWeight = FontWeights.SemiBold,
            Foreground = Brushes.White
        };
        var subtitle = new TextBlock
        {
            Text = "The client identity and current renewal remain locked throughout this import.",
            Foreground = new SolidColorBrush(Color.FromRgb(218, 231, 250)),
            Margin = new Thickness(0, 5, 0, 0)
        };
        var panel = new StackPanel { Margin = new Thickness(22, 18, 22, 18) };
        panel.Children.Add(title);
        panel.Children.Add(subtitle);
        panel.Children.Add(_protectionBanner);
        return new Border { Background = new SolidColorBrush(Color.FromRgb(11, 48, 105)), Child = panel };
    }

    private Border BuildFooter()
    {
        var actions = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        actions.Children.Add(_back);
        actions.Children.Add(_next);
        var grid = new Grid { Margin = new Thickness(22, 12, 22, 12) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.Children.Add(_status);
        Grid.SetColumn(actions, 1);
        grid.Children.Add(actions);
        return new Border
        {
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(220, 228, 240)),
            BorderThickness = new Thickness(0, 1, 0, 0),
            Child = grid
        };
    }

    private void BuildStepStrip()
    {
        _stepStrip.Children.Clear();
        var labels = new[] { "1  Select files", "2  Review", "3  Safety checks", "4  Confirm", "5  Complete" };
        for (var index = 0; index < labels.Length; index++)
        {
            var text = new TextBlock
            {
                Text = labels[index],
                Padding = new Thickness(10, 7, 10, 7),
                Margin = new Thickness(0, 0, 8, 0),
                FontWeight = index == _step ? FontWeights.SemiBold : FontWeights.Normal,
                Foreground = index == _step
                    ? new SolidColorBrush(Color.FromRgb(11, 93, 232))
                    : new SolidColorBrush(Color.FromRgb(97, 113, 140)),
                Background = index == _step
                    ? new SolidColorBrush(Color.FromRgb(232, 240, 255))
                    : Brushes.Transparent
            };
            _stepStrip.Children.Add(text);
        }
    }

    private async void Window_Loaded(object sender, RoutedEventArgs e)
    {
        try
        {
            _client = await _database.GetClientAsync(_clientId);
            if (_client is null)
            {
                MessageBox.Show(this, "The selected persisted client no longer exists.", Title, MessageBoxButton.OK, MessageBoxImage.Warning);
                Close();
                return;
            }

            var versions = await _database.GetRenewalVersionsAsync(_clientId);
            var currentDate = versions.Where(item => item.EffectiveDate.HasValue).Select(item => item.EffectiveDate!.Value).DefaultIfEmpty().Max();
            if (currentDate == default && _client.EffectiveDate.HasValue) currentDate = _client.EffectiveDate.Value;
            _protectionBanner.Text = currentDate == default
                ? $"Client locked: {_client.ClientName}. Current renewal date must be verified before import."
                : $"Client locked: {_client.ClientName}. Current renewal protected: {currentDate:MMM d, yyyy}.";
            _protectionBanner.Margin = new Thickness(0, 9, 0, 0);
            _carrier.Text = _client.Carrier;
            RenderStep();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Client details could not be loaded: {ex.Message}", Title, MessageBoxButton.OK, MessageBoxImage.Error);
            Close();
        }
    }

    private async void Next_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy) return;
        switch (_step)
        {
            case 0:
                if (await ExtractSourcesAsync()) MoveToStep(1);
                break;
            case 1:
                if (!TryBuildRequest(out var request, out var message))
                {
                    MessageBox.Show(this, message, "Review Historical Renewal", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                await RunSafetyChecksAsync(request);
                MoveToStep(2);
                break;
            case 2:
                if (_preview?.CanImport == true) MoveToStep(3);
                else MessageBox.Show(this, _preview?.Message ?? "Run the safety checks before continuing.", "Historical Renewal Blocked", MessageBoxButton.OK, MessageBoxImage.Warning);
                break;
            case 3:
                if (_confirm.IsChecked != true)
                {
                    MessageBox.Show(this, "Confirm the protected changes before importing.", "Confirm Historical Renewal", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                await ImportAsync();
                break;
            case 4:
                Close();
                break;
        }
    }

    private void Back_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy || _step == 0 || _step == 4) return;
        MoveToStep(_step - 1);
    }

    private async Task<bool> ExtractSourcesAsync()
    {
        if (string.IsNullOrWhiteSpace(_pdfPath.Text) || !File.Exists(_pdfPath.Text))
        {
            MessageBox.Show(this, "Select an existing historical renewal PDF.", "Select Historical Renewal", MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }

        SetBusy(true, "Reading historical renewal sources...");
        try
        {
            var analysis = await _reviewPackBuilder.AnalyzeAsync(
                new RenewalCase
                {
                    ClientName = _client?.ClientName ?? "",
                    PolicyNumber = _client?.ClientNumber ?? "",
                    PdfPath = _pdfPath.Text,
                    WorkbookPath = _workbookPath.Text
                });
            _carrier.Text = FirstNonBlank(analysis.Pdf.Carrier, _client?.Carrier);
            var date = analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate;
            if (date.HasValue)
            {
                _effectiveDate.Text = date.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                _renewalYear.Text = date.Value.Year.ToString(CultureInfo.InvariantCulture);
            }
            if (analysis.MonthlyPremiumChangePercent.HasValue)
            {
                _requestedIncrease.Text = analysis.MonthlyPremiumChangePercent.Value.ToString("0.0", CultureInfo.InvariantCulture);
            }
            _extractionSummary = string.Join(" ", analysis.SourceChecks
                .Where(item => item.Status.Equals("Failed", StringComparison.OrdinalIgnoreCase))
                .Select(item => $"{item.Area}: {item.Detail}"));
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"The sources could not be previewed: {ex.Message}", "Preview Historical Renewal", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
        finally
        {
            SetBusy(false, "Source preview complete.");
        }
    }

    private async Task RunSafetyChecksAsync(HistoricalRenewalImportRequest request)
    {
        SetBusy(true, "Running protected-import checks...");
        try
        {
            _request = request;
            _preview = await _database.PreviewHistoricalRenewalImportAsync(request);
        }
        finally
        {
            SetBusy(false, _preview?.Message ?? "Safety checks did not complete.");
        }
    }

    private async Task ImportAsync()
    {
        if (_request is null || _preview?.CanImport != true) return;
        SetBusy(true, "Importing the completed historical renewal...");
        try
        {
            var result = await _database.ImportHistoricalRenewalAsync(_request);
            if (!result.Succeeded)
            {
                MessageBox.Show(this, result.Message, "Historical Renewal Not Imported", MessageBoxButton.OK, MessageBoxImage.Warning);
                _status.Text = result.Message;
                return;
            }
            ImportSucceeded = true;
            _status.Text = result.Message;
            MoveToStep(4);
        }
        finally
        {
            SetBusy(false, _status.Text);
        }
    }

    private bool TryBuildRequest(out HistoricalRenewalImportRequest request, out string message)
    {
        request = new HistoricalRenewalImportRequest();
        message = "";
        if (!DateOnly.TryParseExact(_effectiveDate.Text.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            message = "Enter the effective date as YYYY-MM-DD.";
            return false;
        }
        if (!int.TryParse(_renewalYear.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var year) || year != date.Year)
        {
            message = "Renewal year must match the effective-date year.";
            return false;
        }
        if (!TryParsePercent(_finalAcceptedIncrease.Text, required: true, out var finalAccepted))
        {
            message = "Enter the confirmed final accepted increase as a percentage, for example 7.0%.";
            return false;
        }
        if (!TryParsePercent(_requestedIncrease.Text, required: false, out var requested))
        {
            message = "Requested increase must be a percentage, for example 12.5%.";
            return false;
        }
        if (!TryParseMoney(_annualSavings.Text, out var savings))
        {
            message = "Negotiated annual savings must be a valid non-negative dollar amount.";
            return false;
        }
        request = new HistoricalRenewalImportRequest
        {
            ClientId = _clientId,
            RenewalPdfPath = _pdfPath.Text,
            SupportingWorkbookPath = _workbookPath.Text,
            Carrier = _carrier.Text.Trim(),
            EffectiveDate = date,
            RenewalYear = year,
            RequestedIncreasePercent = requested,
            FinalAcceptedIncreasePercent = finalAccepted!.Value,
            NegotiatedSavingsAnnual = savings,
            ImportedBy = Environment.UserName
        };
        return true;
    }

    private void RenderStep()
    {
        _stage.Children.Clear();
        BuildStepStrip();
        _back.IsEnabled = _step is > 0 and < 4 && !_isBusy;
        _next.IsEnabled = !_isBusy;
        _status.Text = $"Step {_step + 1} of 5";
        switch (_step)
        {
            case 0:
                AddHeading("Select the historical source documents", "The files are parsed into a preview first. Nothing is saved at this step.");
                AddClientLock();
                _stage.Children.Add(FilePicker("Historical renewal PDF", _pdfPath, "PDF files (*.pdf)|*.pdf"));
                _stage.Children.Add(FilePicker("Supporting carrier workbook (optional)", _workbookPath, "Excel workbooks (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"));
                _next.Content = "Preview sources";
                break;
            case 1:
                AddHeading("Review the extracted renewal", "Confirm the historical facts. The final accepted increase is required and is never fabricated from a proposal.");
                AddClientLock();
                var fields = new Grid();
                fields.ColumnDefinitions.Add(new ColumnDefinition());
                fields.ColumnDefinitions.Add(new ColumnDefinition());
                AddField(fields, 0, 0, "Carrier", _carrier);
                AddField(fields, 0, 1, "Renewal effective date (YYYY-MM-DD)", _effectiveDate);
                AddField(fields, 1, 0, "Renewal year", _renewalYear);
                AddField(fields, 1, 1, "Carrier requested increase (%)", _requestedIncrease);
                AddField(fields, 2, 0, "Final accepted increase (%)", _finalAcceptedIncrease);
                AddField(fields, 2, 1, "Negotiated annual savings ($, optional)", _annualSavings);
                _stage.Children.Add(fields);
                if (!string.IsNullOrWhiteSpace(_extractionSummary)) AddWarning(_extractionSummary);
                _next.Content = "Run safety checks";
                break;
            case 2:
                AddHeading("Protected-import safety checks", _preview?.Message ?? "Safety checks are unavailable.");
                foreach (var check in _preview?.Checks ?? []) AddCheck(check, _preview?.CanImport == true);
                _next.Content = "Review changes";
                _next.IsEnabled = _preview?.CanImport == true;
                break;
            case 3:
                AddHeading("Confirm exactly what will change", "Only additive historical records will be committed.");
                AddChange("Historical renewal PDF", "Add", "New source document and evidence link");
                if (!string.IsNullOrWhiteSpace(_request?.SupportingWorkbookPath)) AddChange("Supporting workbook", "Add", "Historical supporting evidence");
                AddChange($"{_request?.RenewalYear} completed renewal", "Add", "Final accepted increase included in the KPI denominator");
                AddChange("Current renewal and selected versions", "Keep unchanged", "Remain current");
                AddChange("Client carrier and effective date", "Keep unchanged", "No client upsert");
                _confirm.IsChecked = false;
                DetachFromParent(_confirm);
                _stage.Children.Add(_confirm);
                _next.Content = "Import historical renewal";
                break;
            case 4:
                AddHeading("Historical renewal added", $"The {_request?.RenewalYear} completed renewal is now available in Client History. The current renewal remains unchanged.");
                AddCheck("Client History will reload its canonical timeline and calculate the average using distinct completed renewals.", true);
                _next.Content = "Close";
                _back.IsEnabled = false;
                break;
        }
    }

    private void MoveToStep(int step)
    {
        _step = Math.Clamp(step, 0, 4);
        RenderStep();
    }

    private void SetBusy(bool busy, string status)
    {
        _isBusy = busy;
        _back.IsEnabled = !busy && _step is > 0 and < 4;
        _next.IsEnabled = !busy;
        _status.Text = status;
    }

    private void AddHeading(string heading, string detail)
    {
        _stage.Children.Add(new TextBlock { Text = heading, FontSize = 18, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(Color.FromRgb(11, 31, 68)) });
        _stage.Children.Add(new TextBlock { Text = detail, TextWrapping = TextWrapping.Wrap, Foreground = new SolidColorBrush(Color.FromRgb(95, 112, 140)), Margin = new Thickness(0, 5, 0, 16) });
    }

    private void AddClientLock()
    {
        var text = _client is null ? "Loading the selected persisted client..." : $"Client locked: {_client.ClientName} (database client {_client.ClientId}).";
        _stage.Children.Add(Callout(text, Color.FromRgb(231, 247, 239), Color.FromRgb(18, 100, 60)));
    }

    private void AddWarning(string text) => _stage.Children.Add(Callout(text, Color.FromRgb(255, 245, 216), Color.FromRgb(112, 71, 0)));

    private void AddCheck(string text, bool passed)
    {
        _stage.Children.Add(Callout($"{(passed ? "✓" : "•")}  {text}", passed ? Color.FromRgb(234, 247, 239) : Color.FromRgb(244, 247, 251), passed ? Color.FromRgb(18, 100, 60) : Color.FromRgb(73, 88, 112)));
    }

    private void AddChange(string record, string action, string result)
    {
        var grid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(3, GridUnitType.Star) });
        grid.Children.Add(new TextBlock { Text = record, TextWrapping = TextWrapping.Wrap });
        var actionText = new TextBlock { Text = action, FontWeight = FontWeights.SemiBold, Foreground = action.Equals("Add", StringComparison.Ordinal) ? new SolidColorBrush(Color.FromRgb(18, 100, 60)) : new SolidColorBrush(Color.FromRgb(11, 82, 189)) };
        Grid.SetColumn(actionText, 1);
        grid.Children.Add(actionText);
        var resultText = new TextBlock { Text = result, TextWrapping = TextWrapping.Wrap };
        Grid.SetColumn(resultText, 2);
        grid.Children.Add(resultText);
        _stage.Children.Add(new Border { BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 241)), BorderThickness = new Thickness(0, 0, 0, 1), Padding = new Thickness(8, 9, 8, 9), Child = grid });
    }

    private static Border Callout(string text, Color background, Color foreground) => new()
    {
        Background = new SolidColorBrush(background),
        CornerRadius = new CornerRadius(6),
        Padding = new Thickness(12),
        Margin = new Thickness(0, 0, 0, 12),
        Child = new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap, Foreground = new SolidColorBrush(foreground) }
    };

    private static FrameworkElement FilePicker(string label, TextBox textBox, string filter)
    {
        var browse = new Button { Content = "Browse...", MinWidth = 90, Margin = new Thickness(8, 0, 0, 0) };
        browse.Click += (_, _) =>
        {
            var dialog = new OpenFileDialog { Filter = filter, CheckFileExists = true };
            if (dialog.ShowDialog() == true) textBox.Text = dialog.FileName;
        };
        var row = new Grid();
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        DetachFromParent(textBox);
        row.Children.Add(textBox);
        Grid.SetColumn(browse, 1);
        row.Children.Add(browse);
        var panel = new StackPanel { Margin = new Thickness(0, 0, 0, 14) };
        panel.Children.Add(new TextBlock { Text = label, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 5) });
        panel.Children.Add(row);
        return panel;
    }

    private static void AddField(Grid grid, int row, int column, string label, TextBox textBox)
    {
        while (grid.RowDefinitions.Count <= row) grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        var panel = new StackPanel { Margin = new Thickness(column == 0 ? 0 : 8, 0, column == 0 ? 8 : 0, 12) };
        panel.Children.Add(new TextBlock { Text = label, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 5) });
        DetachFromParent(textBox);
        panel.Children.Add(textBox);
        Grid.SetRow(panel, row);
        Grid.SetColumn(panel, column);
        grid.Children.Add(panel);
    }

    private static void DetachFromParent(FrameworkElement element)
    {
        switch (LogicalTreeHelper.GetParent(element))
        {
            case Panel panel:
                panel.Children.Remove(element);
                break;
            case Decorator decorator when ReferenceEquals(decorator.Child, element):
                decorator.Child = null;
                break;
            case ContentControl contentControl when ReferenceEquals(contentControl.Content, element):
                contentControl.Content = null;
                break;
            case ContentPresenter presenter when ReferenceEquals(presenter.Content, element):
                presenter.Content = null;
                break;
        }
    }

    private static TextBox Input() => new() { MinHeight = 34, Padding = new Thickness(8, 5, 8, 5) };
    private static TextBox ReadOnlyInput() => new() { MinHeight = 34, Padding = new Thickness(8, 5, 8, 5), IsReadOnly = true, Background = new SolidColorBrush(Color.FromRgb(240, 243, 248)) };

    private static bool TryParsePercent(string text, bool required, out decimal? value)
    {
        value = null;
        var cleaned = text.Trim().TrimEnd('%').Trim();
        if (cleaned.Length == 0) return !required;
        if (!decimal.TryParse(cleaned, NumberStyles.Number | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var percentage) &&
            !decimal.TryParse(cleaned, NumberStyles.Number | NumberStyles.AllowLeadingSign, CultureInfo.CurrentCulture, out percentage)) return false;
        value = decimal.Round(percentage / 100m, 6);
        return value is >= -1m and <= 5m;
    }

    private static bool TryParseMoney(string text, out decimal? value)
    {
        value = null;
        var cleaned = text.Trim().TrimStart('$').Replace(",", "", StringComparison.Ordinal).Trim();
        if (cleaned.Length == 0) return true;
        if (!decimal.TryParse(cleaned, NumberStyles.Number, CultureInfo.InvariantCulture, out var amount) || amount < 0m) return false;
        value = amount;
        return true;
    }

    private static string FirstNonBlank(params string?[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
}
