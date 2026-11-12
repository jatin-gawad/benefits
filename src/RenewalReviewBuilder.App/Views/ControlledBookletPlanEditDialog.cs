using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Views;

internal sealed class ControlledBookletPlanEditDialog : Window
{
    private readonly string _comparisonKey;
    private readonly TextBox _percentageTextBox;
    private readonly TextBox _maximumTextBox;
    private readonly TextBox _frequencyTextBox;
    private readonly TextBox _waitingPeriodTextBox;
    private readonly TextBox _terminationAgeTextBox;
    private readonly TextBox _coverageDescriptionTextBox;
    private readonly TextBox _notesTextBox;
    private readonly TextBox _reasonTextBox;
    private readonly TextBlock _validationTextBlock;

    public ControlledBookletPlanEditDialog(PlanDesignRecord row, string comparisonKey)
    {
        _comparisonKey = comparisonKey;
        Title = "Edit Plan Provision";
        Width = 720;
        Height = 760;
        MinHeight = 620;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = new SolidColorBrush(Color.FromRgb(248, 250, 252));

        var root = new Grid { Margin = new Thickness(24) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition());
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var header = BenefitAdministrationEditDialog.Header(
            string.IsNullOrWhiteSpace(row.ProvisionLabel) ? row.BenefitName : row.ProvisionLabel,
            $"{row.BenefitCategory} | {row.ClassName}" +
            (row.SourcePage.HasValue ? $" | Source page {row.SourcePage}" : ""));
        root.Children.Add(header);

        var fields = new StackPanel { Margin = new Thickness(0, 4, 8, 0) };
        fields.Children.Add(ReadOnlyIdentity(row));

        var percentageAndMaximum = TwoColumnGrid();
        _percentageTextBox = CreateTextBox(
            row.ReimbursementPercent.HasValue
                ? row.ReimbursementPercent.Value.ToString("0.####", CultureInfo.CurrentCulture)
                : row.ReimbursementText.TrimEnd('%'));
        percentageAndMaximum.Children.Add(Field("Reimbursement %", _percentageTextBox));
        _maximumTextBox = CreateTextBox(row.MaximumAmount);
        var maximumField = Field("Maximum", _maximumTextBox, leftMargin: 8);
        Grid.SetColumn(maximumField, 1);
        percentageAndMaximum.Children.Add(maximumField);
        fields.Children.Add(percentageAndMaximum);

        _frequencyTextBox = CreateTextBox(row.Frequency);
        fields.Children.Add(Field("Frequency", _frequencyTextBox));

        var waitingAndTermination = TwoColumnGrid();
        _waitingPeriodTextBox = CreateTextBox(row.WaitingPeriod);
        waitingAndTermination.Children.Add(Field("Waiting Period", _waitingPeriodTextBox));
        _terminationAgeTextBox = CreateTextBox(row.TerminationAge);
        var terminationField = Field("Termination Age", _terminationAgeTextBox, leftMargin: 8);
        Grid.SetColumn(terminationField, 1);
        waitingAndTermination.Children.Add(terminationField);
        fields.Children.Add(waitingAndTermination);

        _coverageDescriptionTextBox = CreateTextBox(row.CoverageDescription, 110, acceptsReturn: true);
        fields.Children.Add(Field("Coverage Details", _coverageDescriptionTextBox));
        _notesTextBox = CreateTextBox(row.Notes, 70, acceptsReturn: true);
        fields.Children.Add(Field("Notes", _notesTextBox));
        _reasonTextBox = CreateTextBox("", 70, acceptsReturn: true);
        fields.Children.Add(Field("Reason for Correction (required)", _reasonTextBox));

        var scrollViewer = new ScrollViewer
        {
            Content = fields,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        };
        Grid.SetRow(scrollViewer, 1);
        root.Children.Add(scrollViewer);

        _validationTextBlock = new TextBlock
        {
            Foreground = new SolidColorBrush(Color.FromRgb(153, 27, 27)),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 8, 0, 0)
        };
        Grid.SetRow(_validationTextBlock, 2);
        root.Children.Add(_validationTextBlock);

        var actions = BenefitAdministrationEditDialog.ActionPanel(3);
        actions.Children.Add(new Button
        {
            Content = "Cancel",
            MinWidth = 90,
            Height = 38,
            IsCancel = true,
            Margin = new Thickness(0, 0, 10, 0)
        });
        var save = BenefitAdministrationEditDialog.PrimaryButton("Save Correction", 130);
        save.IsDefault = true;
        save.Click += Save_Click;
        actions.Children.Add(save);
        root.Children.Add(actions);

        Content = root;
    }

    public ControlledBookletStagedEditRequest? Request { get; private set; }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (!TryReadPercentage(_percentageTextBox.Text, out var percentage, out var percentageError))
        {
            _validationTextBlock.Text = percentageError;
            _percentageTextBox.Focus();
            return;
        }

        var reason = _reasonTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(reason))
        {
            _validationTextBlock.Text = "Enter a reason so this correction is retained in the audit history.";
            _reasonTextBox.Focus();
            return;
        }

        Request = new ControlledBookletStagedEditRequest
        {
            ComparisonKey = _comparisonKey,
            ReimbursementPercent = percentage,
            MaximumAmount = _maximumTextBox.Text.Trim(),
            Frequency = _frequencyTextBox.Text.Trim(),
            WaitingPeriod = _waitingPeriodTextBox.Text.Trim(),
            TerminationAge = _terminationAgeTextBox.Text.Trim(),
            CoverageDescription = _coverageDescriptionTextBox.Text.Trim(),
            Notes = _notesTextBox.Text.Trim(),
            Reason = reason,
            EditedBy = Environment.UserName
        };
        DialogResult = true;
    }

    private static bool TryReadPercentage(string text, out decimal? percentage, out string error)
    {
        percentage = null;
        error = "";
        var value = text.Trim().TrimEnd('%').Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            return true;
        }

        if (!decimal.TryParse(value, NumberStyles.Number, CultureInfo.CurrentCulture, out var parsed) &&
            !decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out parsed))
        {
            error = "Enter one reimbursement percentage, such as 80 or 50. Leave it blank when the provision has no percentage.";
            return false;
        }

        if (parsed is < 0 or > 100)
        {
            error = "Reimbursement must be between 0 and 100.";
            return false;
        }

        percentage = parsed;
        return true;
    }

    private static Border ReadOnlyIdentity(PlanDesignRecord row)
    {
        var panel = new StackPanel();
        panel.Children.Add(new TextBlock
        {
            Text = "Provision Identity",
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(51, 65, 85))
        });
        panel.Children.Add(new TextBlock
        {
            Text = $"{row.BenefitCategory} | {row.ProvisionLabel} | {row.ClassName}",
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 5, 0, 0)
        });
        return new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(241, 245, 249)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(203, 213, 225)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(12),
            Margin = new Thickness(0, 0, 0, 14),
            Child = panel
        };
    }

    private static Grid TwoColumnGrid()
    {
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition());
        grid.ColumnDefinitions.Add(new ColumnDefinition());
        return grid;
    }

    private static StackPanel Field(string label, Control control, double leftMargin = 0)
    {
        var panel = new StackPanel { Margin = new Thickness(leftMargin, 0, 0, 14) };
        panel.Children.Add(new TextBlock { Text = label, FontWeight = FontWeights.SemiBold });
        panel.Children.Add(control);
        return panel;
    }

    private static TextBox CreateTextBox(string value, double height = 38, bool acceptsReturn = false)
    {
        return new TextBox
        {
            Text = value,
            Height = height,
            AcceptsReturn = acceptsReturn,
            VerticalScrollBarVisibility = acceptsReturn ? ScrollBarVisibility.Auto : ScrollBarVisibility.Disabled,
            TextWrapping = acceptsReturn ? TextWrapping.Wrap : TextWrapping.NoWrap,
            Padding = new Thickness(10, 7, 10, 7),
            Margin = new Thickness(0, 6, 0, 0)
        };
    }
}
