using System.Collections.ObjectModel;
using System.ComponentModel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using OutlookAligner.App.Diagnostics;
using OutlookAligner.App.ViewModels;

namespace OutlookAligner.App;

public sealed partial class MainWindow : Window
{
    private const int MaximumVisibleLogEntries = 500;

    private readonly AppEventLog _eventLog = AppEventLog.Current;
    private AppLogLevel? _selectedLogLevel;
    private bool _initialRefreshStarted;
    private string? _lastNoticeLogged;
    private string? _lastStatusLogged;

    public MainWindow()
    {
        InitializeComponent();
        ViewModel = new MainViewModel();
        RootLayout.DataContext = ViewModel;
        RootNavigation.SelectedItem = RootNavigation.MenuItems[0];

        LogLevelFilters =
        [
            new LogLevelChoice("All", null),
            new LogLevelChoice("Debug", AppLogLevel.Debug),
            new LogLevelChoice("Info", AppLogLevel.Info),
            new LogLevelChoice("Warning", AppLogLevel.Warning),
            new LogLevelChoice("Error", AppLogLevel.Error),
        ];

        RebuildVisibleLogEntries();
        _eventLog.EntryAdded += OnLogEntryAdded;
        ViewModel.PropertyChanged += OnViewModelPropertyChanged;
        Closed += OnClosed;
        _eventLog.Info("Application", "WinUI session started.");
    }

    public MainViewModel ViewModel { get; }

    public ObservableCollection<AppLogEntry> VisibleLogEntries { get; } = [];

    public IReadOnlyList<LogLevelChoice> LogLevelFilters { get; }

    public string LogFilePathDisplay => $"Persistent log: {_eventLog.FilePath}";

    private void OnRootLoaded(object sender, RoutedEventArgs e)
    {
        _ = sender;
        _ = e;

        if (_initialRefreshStarted)
        {
            return;
        }

        _initialRefreshStarted = true;
        _eventLog.Info("Calendar", "Automatic Outlook refresh requested at startup.");
        if (ViewModel.RefreshCommand.CanExecute(null))
        {
            ViewModel.RefreshCommand.Execute(null);
        }
    }

    private void OnNavigationSelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
    {
        _ = sender;

        var tag = args.SelectedItemContainer?.Tag as string;
        CalendarPanel.Visibility = tag == "calendar" ? Visibility.Visible : Visibility.Collapsed;
        AlignmentPanel.Visibility = tag == "alignment" ? Visibility.Visible : Visibility.Collapsed;
        SettingsPanel.Visibility = tag == "settings" ? Visibility.Visible : Visibility.Collapsed;
        DiagnosticsPanel.Visibility = tag == "diagnostics" ? Visibility.Visible : Visibility.Collapsed;

        if (!string.IsNullOrWhiteSpace(tag))
        {
            _eventLog.Debug("Navigation", $"Opened {tag} view.");
        }
    }

    private void OnCalendarSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        _ = e;
        if (sender is not ListView { SelectedItem: CalendarEventRowViewModel selected })
        {
            return;
        }

        _eventLog.Debug(
            "Calendar",
            "Calendar event selected.",
            BuildEventContext(selected),
            BuildEventTechnicalDetails(selected));
    }

    private void OnAlignmentSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        _ = e;
        if (sender is not ListView { SelectedItem: AlignmentGroupViewModel selected })
        {
            return;
        }

        _eventLog.Debug(
            "Alignment",
            $"Alignment row selected: {selected.State}.",
            selected.Subject,
            selected.Explanation);
    }

    private void OnRefreshClicked(object sender, RoutedEventArgs e)
    {
        _ = sender;
        _ = e;
        _eventLog.Info("Calendar", "Manual Outlook refresh requested.");
    }

    private void OnForwardActionClicked(object sender, RoutedEventArgs e)
    {
        _ = e;
        if (sender is not Button button || ViewModel.SelectedEvent is not { } selected)
        {
            return;
        }

        var action = button.Tag as string ?? "forward-action";
        var target = ViewModel.SelectedTargetAccount?.SmtpAddress;
        var details = BuildEventTechnicalDetails(selected);
        if (!string.IsNullOrWhiteSpace(target))
        {
            details += Environment.NewLine + $"Target account: {target}";
        }

        _eventLog.Info(
            "Forward",
            $"User requested {action}.",
            BuildEventContext(selected),
            details);
    }

    private void OnLogLevelFilterChanged(object sender, SelectionChangedEventArgs e)
    {
        _ = e;
        _selectedLogLevel = (sender as ComboBox)?.SelectedItem is LogLevelChoice choice
            ? choice.Level
            : null;
        RebuildVisibleLogEntries();
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        _ = sender;

        if (e.PropertyName == nameof(MainViewModel.NoticeMessage)
            && ViewModel.HasNotice
            && !string.IsNullOrWhiteSpace(ViewModel.NoticeMessage)
            && !string.Equals(_lastNoticeLogged, ViewModel.NoticeMessage, StringComparison.Ordinal))
        {
            _lastNoticeLogged = ViewModel.NoticeMessage;
            _eventLog.Warning(
                "UI",
                ViewModel.NoticeMessage,
                ViewModel.SelectedEvent is null ? null : BuildEventContext(ViewModel.SelectedEvent),
                ViewModel.LastDiagnostics);
        }

        if (e.PropertyName == nameof(MainViewModel.Status)
            && !string.IsNullOrWhiteSpace(ViewModel.Status)
            && !string.Equals(_lastStatusLogged, ViewModel.Status, StringComparison.Ordinal))
        {
            _lastStatusLogged = ViewModel.Status;
            var context = ViewModel.SelectedEvent is null ? null : BuildEventContext(ViewModel.SelectedEvent);
            if (ViewModel.Status.Contains("failed", StringComparison.OrdinalIgnoreCase))
            {
                _eventLog.Error("UI", ViewModel.Status, context, ViewModel.LastDiagnostics);
            }
            else
            {
                _eventLog.Debug("UI", ViewModel.Status, context);
            }
        }
    }

    private void OnLogEntryAdded(AppLogEntry entry)
    {
        _ = DispatcherQueue.TryEnqueue(() => AddVisibleEntryIfMatched(entry));
    }

    private void AddVisibleEntryIfMatched(AppLogEntry entry)
    {
        if (_selectedLogLevel is not null && entry.Level != _selectedLogLevel)
        {
            return;
        }

        VisibleLogEntries.Insert(0, entry);
        if (VisibleLogEntries.Count > MaximumVisibleLogEntries)
        {
            VisibleLogEntries.RemoveAt(VisibleLogEntries.Count - 1);
        }
    }

    private void RebuildVisibleLogEntries()
    {
        VisibleLogEntries.Clear();
        foreach (var entry in _eventLog.Snapshot()
                     .Where(entry => _selectedLogLevel is null || entry.Level == _selectedLogLevel)
                     .Reverse()
                     .Take(MaximumVisibleLogEntries))
        {
            VisibleLogEntries.Add(entry);
        }
    }

    private void OnClosed(object sender, WindowEventArgs args)
    {
        _ = sender;
        _ = args;
        _eventLog.Info("Application", "WinUI session closed.");
        _eventLog.EntryAdded -= OnLogEntryAdded;
        ViewModel.PropertyChanged -= OnViewModelPropertyChanged;
        Closed -= OnClosed;
    }

    private static string BuildEventContext(CalendarEventRowViewModel selected)
        => $"{selected.Subject} | {selected.SourceSmtp} | {selected.TimeDisplay} | {selected.MeetingStatusDisplay} | "
            + (selected.CalendarEvent.IsRecurring ? "recurring" : "single");

    private static string BuildEventTechnicalDetails(CalendarEventRowViewModel selected)
        => string.Join(
            Environment.NewLine,
            $"GlobalAppointmentID: {selected.CalendarEvent.GlobalAppointmentId ?? "(unavailable)"}",
            $"EntryID: {selected.CalendarEvent.EntryId ?? "(unavailable)"}",
            $"StoreID: {selected.Account.StoreId ?? "(unavailable)"}",
            $"RecurrenceState: {selected.CalendarEvent.RecurrenceState}",
            $"MeetingStatus: {selected.CalendarEvent.MeetingStatus}",
            $"ForwardEligibility: {selected.ForwardEligibilityMessage}");
}
