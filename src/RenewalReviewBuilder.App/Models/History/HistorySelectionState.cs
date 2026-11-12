using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace RenewalReviewBuilder.App.Models.History;
public sealed class HistorySelectionState : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;
    private string _selectedHistoryEventId = ""; private string _selectedEvolutionStepId = ""; private long? _selectedRenewalId; private int? _selectedYear; private long? _selectedCarrierId; private IReadOnlyList<long> _selectedDocumentIds = []; private HistoryMode _activeHistoryMode; private bool _isDetailsDrawerOpen; private DateOnly? _replayAsOfDate; private bool _isReplayActive;
    public string SelectedHistoryEventId { get => _selectedHistoryEventId; set => Set(ref _selectedHistoryEventId, value); } public string SelectedEvolutionStepId { get => _selectedEvolutionStepId; set => Set(ref _selectedEvolutionStepId, value); } public long? SelectedRenewalId { get => _selectedRenewalId; set => Set(ref _selectedRenewalId, value); } public int? SelectedYear { get => _selectedYear; set => Set(ref _selectedYear, value); } public long? SelectedCarrierId { get => _selectedCarrierId; set => Set(ref _selectedCarrierId, value); } public IReadOnlyList<long> SelectedDocumentIds { get => _selectedDocumentIds; set => Set(ref _selectedDocumentIds, value); } public HistoryMode ActiveHistoryMode { get => _activeHistoryMode; set => Set(ref _activeHistoryMode, value); } public bool IsDetailsDrawerOpen { get => _isDetailsDrawerOpen; set => Set(ref _isDetailsDrawerOpen, value); } public DateOnly? ReplayAsOfDate { get => _replayAsOfDate; set => Set(ref _replayAsOfDate, value); } public bool IsReplayActive { get => _isReplayActive; set => Set(ref _isReplayActive, value); } public void ClearSelection() { SelectedHistoryEventId = ""; SelectedEvolutionStepId = ""; SelectedRenewalId = null; SelectedYear = null; SelectedCarrierId = null; SelectedDocumentIds = []; IsDetailsDrawerOpen = false; }
    private void Set<T>(ref T field, T value, [CallerMemberName] string? name = null) { if (EqualityComparer<T>.Default.Equals(field, value)) return; field = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name)); }
}
