using System.Collections.Generic;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Media;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Views;

internal sealed class FlexibleBenefitCorrectionDialog : Window
{
    private readonly TextBox _valueTextBox;
    private readonly TextBox _reasonTextBox;

    public FlexibleBenefitCorrectionDialog(FlexibleBenefitSourceFieldDisplayRow sourceField)
    {
        Title = "Edit and Confirm Extracted Value";
        Width = 620;
        Height = 470;
        MinWidth = 500;
        MinHeight = 400;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.CanResize;
        Background = new SolidColorBrush(Color.FromRgb(248, 250, 252));

        var root = new Grid { Margin = new Thickness(24) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition());
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var title = new TextBlock
        {
            Text = $"{sourceField.EntityType} - {sourceField.FieldName}",
            FontSize = 20,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(15, 23, 42)),
            TextWrapping = TextWrapping.Wrap
        };
        root.Children.Add(title);

        var sourcePanel = new StackPanel { Margin = new Thickness(0, 18, 0, 14) };
        sourcePanel.Children.Add(LabelText("Source PDF value"));
        sourcePanel.Children.Add(new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(sourceField.RawValue) ? "Not confirmed" : sourceField.RawValue,
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(71, 85, 105)),
            Margin = new Thickness(0, 5, 0, 0)
        });
        Grid.SetRow(sourcePanel, 1);
        root.Children.Add(sourcePanel);

        var valuePanel = new StackPanel { Margin = new Thickness(0, 0, 0, 14) };
        valuePanel.Children.Add(LabelText("Confirmed normalized value"));
        _valueTextBox = new TextBox
        {
            Text = sourceField.NormalizedValue,
            MinHeight = 40,
            Padding = new Thickness(10, 8, 10, 8),
            Margin = new Thickness(0, 5, 0, 0),
            TextWrapping = TextWrapping.Wrap,
            AcceptsReturn = true
        };
        AutomationProperties.SetName(_valueTextBox, "Confirmed normalized value");
        valuePanel.Children.Add(_valueTextBox);
        Grid.SetRow(valuePanel, 2);
        root.Children.Add(valuePanel);

        var reasonPanel = new StackPanel { Margin = new Thickness(0, 0, 0, 14) };
        reasonPanel.Children.Add(LabelText("Correction reason"));
        _reasonTextBox = new TextBox
        {
            Text = "Confirmed against the signed source document.",
            MinHeight = 58,
            Padding = new Thickness(10, 8, 10, 8),
            Margin = new Thickness(0, 5, 0, 0),
            TextWrapping = TextWrapping.Wrap,
            AcceptsReturn = true
        };
        AutomationProperties.SetName(_reasonTextBox, "Correction reason");
        reasonPanel.Children.Add(_reasonTextBox);
        Grid.SetRow(reasonPanel, 3);
        root.Children.Add(reasonPanel);

        var note = new TextBlock
        {
            Text = "The source PDF and raw extracted value are preserved. This changes only the normalized operational value and records a correction audit.",
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
            VerticalAlignment = VerticalAlignment.Top
        };
        Grid.SetRow(note, 4);
        root.Children.Add(note);

        var actions = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 18, 0, 0)
        };
        var cancelButton = new Button { Content = "Cancel", MinWidth = 92, Height = 38, Margin = new Thickness(0, 0, 10, 0), IsCancel = true };
        var saveButton = new Button
        {
            Content = "Save Correction",
            MinWidth = 132,
            Height = 38,
            IsDefault = true,
            Background = new SolidColorBrush(Color.FromRgb(29, 78, 216)),
            Foreground = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(29, 78, 216))
        };
        saveButton.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_valueTextBox.Text))
            {
                MessageBox.Show(this, "Enter a confirmed value.", Title, MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            DialogResult = true;
        };
        actions.Children.Add(cancelButton);
        actions.Children.Add(saveButton);
        Grid.SetRow(actions, 5);
        root.Children.Add(actions);

        Content = root;
        Loaded += (_, _) =>
        {
            _valueTextBox.Focus();
            _valueTextBox.SelectAll();
        };
    }

    public string ConfirmedValue => _valueTextBox.Text.Trim();
    public string Reason => _reasonTextBox.Text.Trim();

    private static TextBlock LabelText(string text)
    {
        return new TextBlock
        {
            Text = text,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(30, 41, 59))
        };
    }
}

internal sealed class FlexibleBenefitVersionComparisonDialog : Window
{
    public FlexibleBenefitVersionComparisonDialog(
        int currentVersion,
        int priorVersion,
        IReadOnlyList<FlexibleBenefitVersionDifferenceRow> differences)
    {
        Title = $"Flexible Benefits Version {currentVersion} vs Version {priorVersion}";
        Width = 900;
        Height = 620;
        MinWidth = 700;
        MinHeight = 460;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = new SolidColorBrush(Color.FromRgb(248, 250, 252));

        var root = new Grid { Margin = new Thickness(22) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition());
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var header = new StackPanel { Margin = new Thickness(0, 0, 0, 16) };
        header.Children.Add(new TextBlock
        {
            Text = "Flexible Benefits Version Comparison",
            FontSize = 21,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(15, 23, 42))
        });
        header.Children.Add(new TextBlock
        {
            Text = differences.Count == 0
                ? "No normalized plan-design differences were found."
                : $"{differences.Count:#,0} normalized difference(s) found.",
            Foreground = new SolidColorBrush(Color.FromRgb(71, 85, 105)),
            Margin = new Thickness(0, 5, 0, 0)
        });
        root.Children.Add(header);

        var grid = new DataGrid
        {
            ItemsSource = differences,
            AutoGenerateColumns = false,
            IsReadOnly = true,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            HeadersVisibility = DataGridHeadersVisibility.Column,
            GridLinesVisibility = DataGridGridLinesVisibility.Horizontal,
            AlternatingRowBackground = new SolidColorBrush(Color.FromRgb(248, 250, 252)),
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
            BorderThickness = new Thickness(1)
        };
        grid.Columns.Add(new DataGridTextColumn { Header = "Area", Binding = new System.Windows.Data.Binding(nameof(FlexibleBenefitVersionDifferenceRow.Area)), Width = 150 });
        grid.Columns.Add(new DataGridTextColumn { Header = "Field", Binding = new System.Windows.Data.Binding(nameof(FlexibleBenefitVersionDifferenceRow.Field)), Width = new DataGridLength(1.4, DataGridLengthUnitType.Star) });
        grid.Columns.Add(new DataGridTextColumn { Header = $"Version {priorVersion}", Binding = new System.Windows.Data.Binding(nameof(FlexibleBenefitVersionDifferenceRow.PriorValue)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        grid.Columns.Add(new DataGridTextColumn { Header = $"Version {currentVersion}", Binding = new System.Windows.Data.Binding(nameof(FlexibleBenefitVersionDifferenceRow.CurrentValue)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        Grid.SetRow(grid, 1);
        root.Children.Add(grid);

        var closeButton = new Button
        {
            Content = "Close",
            MinWidth = 100,
            Height = 38,
            IsDefault = true,
            IsCancel = true,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 16, 0, 0)
        };
        Grid.SetRow(closeButton, 2);
        root.Children.Add(closeButton);
        Content = root;
    }
}
