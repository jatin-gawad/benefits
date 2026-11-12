using System.Windows;
using System.Windows.Controls;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Views;

public partial class ProposalComparisonWorkspaceView : UserControl
{
    private readonly ProposalComparisonDataService _dataService = new();
    private ProposalComparisonViewModel _model = new();
    private long? _clientId;
    private string _clientName = "";
    private string _carrier = "";

    public event EventHandler? RunExistingComparisonRequested;
    public event EventHandler? ClientsBreadcrumbRequested;
    public event EventHandler? ViewPlanDesignRequested;

    public ProposalComparisonWorkspaceView()
    {
        InitializeComponent();
        DataContext = _model;
    }

    public async Task LoadClientAsync(long? clientId, string? clientName = null, string? carrier = null)
    {
        _clientId = clientId;
        _clientName = clientName ?? "";
        _carrier = carrier ?? "";

        if (clientId is not > 0)
        {
            _model = new ProposalComparisonViewModel
            {
                ClientName = string.IsNullOrWhiteSpace(_clientName) ? "No client selected" : _clientName,
                CurrentCarrier = string.IsNullOrWhiteSpace(_carrier) ? "Not available" : _carrier,
                IsEmpty = true,
                StatusText = "Select a saved client before loading proposal comparison data."
            };
            DataContext = _model;
            return;
        }

        try
        {
            _model = new ProposalComparisonViewModel
            {
                ClientDatabaseId = clientId,
                ClientName = _clientName,
                CurrentCarrier = string.IsNullOrWhiteSpace(_carrier) ? "Not available" : _carrier,
                IsLoading = true,
                IsEmpty = false,
                StatusText = "Loading saved Proposal Intelligence data."
            };
            DataContext = _model;
            _model = await _dataService.LoadAsync(clientId.Value, _clientName, _carrier);
            DataContext = _model;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            _model = new ProposalComparisonViewModel
            {
                ClientDatabaseId = clientId,
                ClientName = _clientName,
                CurrentCarrier = string.IsNullOrWhiteSpace(_carrier) ? "Not available" : _carrier,
                HasError = true,
                IsEmpty = false,
                StatusText = $"Saved proposal comparison data could not load: {ex.Message}"
            };
            DataContext = _model;
        }
    }

    public Task RefreshAsync()
    {
        return LoadClientAsync(_clientId, _clientName, _carrier);
    }

    private async void Refresh_Click(object sender, RoutedEventArgs e)
    {
        await RefreshAsync();
    }

    private void MoreActions_Click(object sender, RoutedEventArgs e)
    {
        RunExistingComparisonRequested?.Invoke(this, EventArgs.Empty);
    }

    private void ProposalBreadcrumbClients_Click(object sender, RoutedEventArgs e)
    {
        ClientsBreadcrumbRequested?.Invoke(this, EventArgs.Empty);
    }

    private void ViewFullPlanDesign_Click(object sender, RoutedEventArgs e)
    {
        ViewPlanDesignRequested?.Invoke(this, EventArgs.Empty);
    }
}
