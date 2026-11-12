using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Views;

internal sealed class BenefitAdministrationEditDialog : Window
{
    private readonly ComboBox _insurerComboBox;
    private readonly TextBox _policyNumberTextBox;

    public BenefitAdministrationEditDialog(
        BenefitAdministrationWorkspaceRow row,
        IReadOnlyList<OrganizationRecord> organizations)
    {
        Title = $"Edit {row.Benefit}";
        Width = 560;
        Height = 360;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        Background = new SolidColorBrush(Color.FromRgb(248, 250, 252));

        var insurerOptions = new List<OrganizationChoice>
        {
            new(null, "No insurance company")
        };
        insurerOptions.AddRange(organizations
            .Where(item => item.OrganizationTypeCode is BenefitAdministrationOrganizationTypes.Insurer or BenefitAdministrationOrganizationTypes.Carrier)
            .Select(item => new OrganizationChoice(item.OrganizationId, item.DisplayName)));

        var root = BuildRootGrid();
        root.Children.Add(Header(row.Benefit, "Correct the administrative assignment without changing benefit provisions."));

        var insurerPanel = FieldPanel("Insurance Company", 1);
        _insurerComboBox = new ComboBox
        {
            ItemsSource = insurerOptions,
            DisplayMemberPath = nameof(OrganizationChoice.DisplayName),
            Height = 38,
            Margin = new Thickness(0, 6, 0, 0)
        };
        _insurerComboBox.SelectedItem = insurerOptions.FirstOrDefault(item =>
            item.DisplayName.Equals(row.InsuranceCompany, StringComparison.OrdinalIgnoreCase)) ?? insurerOptions[0];
        insurerPanel.Children.Add(_insurerComboBox);
        root.Children.Add(insurerPanel);

        var policyPanel = FieldPanel("Policy Number", 2);
        _policyNumberTextBox = new TextBox
        {
            Text = row.PolicyNumber.Equals("Not saved", StringComparison.OrdinalIgnoreCase) ? "" : row.PolicyNumber,
            Height = 38,
            Padding = new Thickness(10, 7, 10, 7),
            Margin = new Thickness(0, 6, 0, 0)
        };
        policyPanel.Children.Add(_policyNumberTextBox);
        root.Children.Add(policyPanel);

        var actions = ActionPanel(3);
        actions.Children.Add(new Button { Content = "Cancel", MinWidth = 90, Height = 38, IsCancel = true, Margin = new Thickness(0, 0, 10, 0) });
        var save = PrimaryButton("Save Changes", 120);
        save.IsDefault = true;
        save.Click += (_, _) => DialogResult = true;
        actions.Children.Add(save);
        root.Children.Add(actions);
        Content = root;
    }

    public long? SelectedInsurerOrganizationId => (_insurerComboBox.SelectedItem as OrganizationChoice)?.OrganizationId;
    public string PolicyNumber => _policyNumberTextBox.Text.Trim();

    private sealed record OrganizationChoice(long? OrganizationId, string DisplayName);

    internal static Grid BuildRootGrid()
    {
        var grid = new Grid { Margin = new Thickness(24) };
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        return grid;
    }

    internal static StackPanel Header(string title, string subtitle)
    {
        var panel = new StackPanel { Margin = new Thickness(0, 0, 0, 18) };
        panel.Children.Add(new TextBlock
        {
            Text = title,
            FontSize = 20,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(15, 23, 42)),
            TextWrapping = TextWrapping.Wrap
        });
        panel.Children.Add(new TextBlock
        {
            Text = subtitle,
            Foreground = new SolidColorBrush(Color.FromRgb(71, 85, 105)),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 6, 0, 0)
        });
        return panel;
    }

    internal static StackPanel FieldPanel(string label, int row)
    {
        var panel = new StackPanel { Margin = new Thickness(0, 0, 0, 14) };
        panel.Children.Add(new TextBlock { Text = label, FontWeight = FontWeights.SemiBold });
        Grid.SetRow(panel, row);
        return panel;
    }

    internal static StackPanel ActionPanel(int row)
    {
        var panel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0)
        };
        Grid.SetRow(panel, row);
        return panel;
    }

    internal static Button PrimaryButton(string content, double minWidth)
    {
        return new Button
        {
            Content = content,
            MinWidth = minWidth,
            Height = 38,
            Background = new SolidColorBrush(Color.FromRgb(0, 128, 128)),
            Foreground = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(0, 128, 128))
        };
    }
}

internal enum BenefitProviderDialogAction
{
    None,
    Save,
    EndDate,
    Remove
}

internal sealed class BenefitProviderAssignmentDialog : Window
{
    private readonly ComboBox _providerComboBox;
    private readonly ComboBox _roleComboBox;
    private readonly DatePicker _effectiveDatePicker;
    private readonly DatePicker _terminationDatePicker;

    public BenefitProviderAssignmentDialog(
        string benefitName,
        IReadOnlyList<OrganizationRecord> organizations,
        BenefitServiceProviderAssignmentRecord? assignment = null)
    {
        Title = assignment is null ? $"Add Provider - {benefitName}" : $"Manage Provider - {benefitName}";
        Width = 620;
        Height = 510;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        Background = new SolidColorBrush(Color.FromRgb(248, 250, 252));

        var root = new Grid { Margin = new Thickness(24) };
        for (var index = 0; index < 6; index++)
        {
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        }
        root.Children.Add(BenefitAdministrationEditDialog.Header(
            assignment is null ? "Add Service Provider" : "Manage Service Provider",
            "Provider assignments remain separate from the insurance company and benefit provisions."));

        var providerPanel = BenefitAdministrationEditDialog.FieldPanel("Service Provider", 1);
        _providerComboBox = new ComboBox
        {
            ItemsSource = organizations,
            DisplayMemberPath = nameof(OrganizationRecord.DisplayName),
            Height = 38,
            Margin = new Thickness(0, 6, 0, 0)
        };
        _providerComboBox.SelectedItem = assignment is null
            ? organizations.FirstOrDefault()
            : organizations.FirstOrDefault(item => item.OrganizationId == assignment.ProviderOrganizationId);
        providerPanel.Children.Add(_providerComboBox);
        root.Children.Add(providerPanel);

        var rolePanel = BenefitAdministrationEditDialog.FieldPanel("Provider Role", 2);
        _roleComboBox = new ComboBox
        {
            ItemsSource = new[]
            {
                BenefitAdministrationProviderRoles.BenefitAdministrator,
                BenefitAdministrationProviderRoles.ClaimsAdministrator,
                BenefitAdministrationProviderRoles.DisabilityManagement,
                BenefitAdministrationProviderRoles.TravelAssistance,
                BenefitAdministrationProviderRoles.MedicalSecondOpinion,
                BenefitAdministrationProviderRoles.VirtualCare,
                BenefitAdministrationProviderRoles.OptionalBenefitAdministrator,
                BenefitAdministrationProviderRoles.Other
            },
            Height = 38,
            Margin = new Thickness(0, 6, 0, 0)
        };
        _roleComboBox.SelectedItem = assignment?.ProviderRoleCode ?? BenefitAdministrationProviderRoles.Other;
        rolePanel.Children.Add(_roleComboBox);
        root.Children.Add(rolePanel);

        var dates = new Grid { Margin = new Thickness(0, 0, 0, 14) };
        dates.ColumnDefinitions.Add(new ColumnDefinition());
        dates.ColumnDefinitions.Add(new ColumnDefinition());
        var effectivePanel = new StackPanel { Margin = new Thickness(0, 0, 8, 0) };
        effectivePanel.Children.Add(new TextBlock { Text = "Effective Date", FontWeight = FontWeights.SemiBold });
        _effectiveDatePicker = new DatePicker { SelectedDate = assignment?.EffectiveDate, Height = 38, Margin = new Thickness(0, 6, 0, 0) };
        effectivePanel.Children.Add(_effectiveDatePicker);
        dates.Children.Add(effectivePanel);
        var terminationPanel = new StackPanel { Margin = new Thickness(8, 0, 0, 0) };
        terminationPanel.Children.Add(new TextBlock { Text = "Termination Date", FontWeight = FontWeights.SemiBold });
        _terminationDatePicker = new DatePicker { SelectedDate = assignment?.TerminationDate, Height = 38, Margin = new Thickness(0, 6, 0, 0) };
        terminationPanel.Children.Add(_terminationDatePicker);
        Grid.SetColumn(terminationPanel, 1);
        dates.Children.Add(terminationPanel);
        Grid.SetRow(dates, 3);
        root.Children.Add(dates);

        var note = new TextBlock
        {
            Text = "Changes return the assignment to Needs Review. Removing or end-dating does not delete its history.",
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
            Margin = new Thickness(0, 4, 0, 10)
        };
        Grid.SetRow(note, 4);
        root.Children.Add(note);

        var actions = BenefitAdministrationEditDialog.ActionPanel(5);
        if (assignment is not null)
        {
            var remove = new Button { Content = "Remove", MinWidth = 88, Height = 38, Margin = new Thickness(0, 0, 8, 0) };
            remove.Click += (_, _) => Complete(BenefitProviderDialogAction.Remove);
            actions.Children.Add(remove);
            var end = new Button { Content = "End Assignment", MinWidth = 118, Height = 38, Margin = new Thickness(0, 0, 8, 0) };
            end.Click += (_, _) => Complete(BenefitProviderDialogAction.EndDate);
            actions.Children.Add(end);
        }
        actions.Children.Add(new Button { Content = "Cancel", MinWidth = 88, Height = 38, IsCancel = true, Margin = new Thickness(0, 0, 8, 0) });
        var save = BenefitAdministrationEditDialog.PrimaryButton(assignment is null ? "Add Provider" : "Save Changes", 116);
        save.IsDefault = true;
        save.Click += (_, _) => Complete(BenefitProviderDialogAction.Save);
        actions.Children.Add(save);
        root.Children.Add(actions);

        Content = root;
        ExistingAssignment = assignment;
    }

    public BenefitServiceProviderAssignmentRecord? ExistingAssignment { get; }
    public BenefitProviderDialogAction RequestedAction { get; private set; }
    public OrganizationRecord? SelectedProvider => _providerComboBox.SelectedItem as OrganizationRecord;
    public string SelectedRole => _roleComboBox.SelectedItem?.ToString() ?? BenefitAdministrationProviderRoles.Other;
    public DateTime? EffectiveDate => _effectiveDatePicker.SelectedDate;
    public DateTime? TerminationDate => _terminationDatePicker.SelectedDate;

    private void Complete(BenefitProviderDialogAction action)
    {
        if (action == BenefitProviderDialogAction.Save && SelectedProvider is null)
        {
            MessageBox.Show(this, "Select a service provider.", Title, MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (action == BenefitProviderDialogAction.EndDate && !TerminationDate.HasValue)
        {
            MessageBox.Show(this, "Select a termination date.", Title, MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        RequestedAction = action;
        DialogResult = true;
    }
}

internal sealed class BenefitProviderChoiceDialog : Window
{
    private readonly ComboBox _assignmentComboBox;

    public BenefitProviderChoiceDialog(IReadOnlyList<BenefitServiceProviderAssignmentRecord> assignments)
    {
        Title = "Select Service Provider";
        Width = 500;
        Height = 230;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        Background = new SolidColorBrush(Color.FromRgb(248, 250, 252));

        var panel = new StackPanel { Margin = new Thickness(24) };
        panel.Children.Add(new TextBlock { Text = "Select the provider assignment to manage", FontSize = 18, FontWeight = FontWeights.SemiBold });
        _assignmentComboBox = new ComboBox
        {
            ItemsSource = assignments,
            SelectedIndex = 0,
            Height = 38,
            Margin = new Thickness(0, 18, 0, 18)
        };
        _assignmentComboBox.ItemTemplate = new DataTemplate
        {
            VisualTree = BuildAssignmentTemplate()
        };
        panel.Children.Add(_assignmentComboBox);
        var actions = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        actions.Children.Add(new Button { Content = "Cancel", MinWidth = 88, Height = 38, IsCancel = true, Margin = new Thickness(0, 0, 8, 0) });
        var select = BenefitAdministrationEditDialog.PrimaryButton("Manage", 96);
        select.IsDefault = true;
        select.Click += (_, _) => DialogResult = true;
        actions.Children.Add(select);
        panel.Children.Add(actions);
        Content = panel;
    }

    public BenefitServiceProviderAssignmentRecord? SelectedAssignment => _assignmentComboBox.SelectedItem as BenefitServiceProviderAssignmentRecord;

    private static FrameworkElementFactory BuildAssignmentTemplate()
    {
        var panel = new FrameworkElementFactory(typeof(StackPanel));
        panel.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);
        var provider = new FrameworkElementFactory(typeof(TextBlock));
        provider.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding(nameof(BenefitServiceProviderAssignmentRecord.ProviderName)));
        provider.SetValue(TextBlock.FontWeightProperty, FontWeights.SemiBold);
        panel.AppendChild(provider);
        var separator = new FrameworkElementFactory(typeof(TextBlock));
        separator.SetValue(TextBlock.TextProperty, " - ");
        panel.AppendChild(separator);
        var role = new FrameworkElementFactory(typeof(TextBlock));
        role.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding(nameof(BenefitServiceProviderAssignmentRecord.ProviderRoleCode)));
        panel.AppendChild(role);
        return panel;
    }
}
