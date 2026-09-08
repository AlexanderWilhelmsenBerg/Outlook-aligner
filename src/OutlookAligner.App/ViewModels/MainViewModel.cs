using System.Collections.ObjectModel;
using System.Text.Json;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OutlookAligner.App.Services;
using OutlookAligner.Core.Alignment;
using OutlookAligner.Outlook.Contracts;

namespace OutlookAligner.App.ViewModels;

public sealed class MainViewModel : ObservableObject
{
    private readonly OutlookHostProcessClient _hostClient = new();
    private readonly Dictionary<string, string> _authorityByGroupKey = new(StringComparer.Ordinal);
    private CalendarEventRowViewModel? _selectedEvent;
    private AlignmentGroupViewModel? _selectedAlignmentGroup;
    private OutlookAccountChoice? _selectedTargetAccount;
    private AuthorityChoice? _selectedAuthority;
    private double _scanDays = 90;
    private bool _isBusy;
    private bool _hasNotice;
    private string _status = "Ready to read Classic Outlook.";
    private string _forwardCapability = "Not checked";
    private string _lastDiagnostics = "No diagnostics yet.";
    private string _selectedEventDiagnostics = "Select a calendar item to inspect its Outlook identity.";
    private string _alignmentSummary = "Refresh Outlook to build the alignment view.";
    private string _authorityStatus = "Select a logical event to inspect authority.";
    private string _noticeMessage = string.Empty;

    public MainViewModel()
    {
        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => !IsBusy);
        ProbeForwardCommand = new AsyncRelayCommand(ProbeForwardAsync, CanProbeForward);
        PrepareForwardCommand = new AsyncRelayCommand(PrepareForwardAsync, CanPrepareForward);
    }

    public ObservableCollection<CalendarEventRowViewModel> Events { get; } = [];

    public ObservableCollection<OutlookAccountChoice> Accounts { get; } = [];

    public ObservableCollection<OutlookAccountChoice> TargetAccounts { get; } = [];

    public ObservableCollection<AlignmentGroupViewModel> AlignmentGroups { get; } = [];

    public ObservableCollection<AuthorityChoice> AuthorityChoices { get; } = [];

    public AsyncRelayCommand RefreshCommand { get; }

    public AsyncRelayCommand ProbeForwardCommand { get; }

    public AsyncRelayCommand PrepareForwardCommand { get; }

    public CalendarEventRowViewModel? SelectedEvent
    {
        get => _selectedEvent;
        set
        {
            if (!SetProperty(ref _selectedEvent, value))
            {
                return;
            }

            ForwardCapability = "Not checked";
            SelectedEventDiagnostics = BuildSelectedEventDiagnostics(value);
            UpdateTargetAccounts();
            NotifyCommandStateChanged();
        }
    }

    public AlignmentGroupViewModel? SelectedAlignmentGroup
    {
        get => _selectedAlignmentGroup;
        set
        {
            if (!SetProperty(ref _selectedAlignmentGroup, value))
            {
                return;
            }

            UpdateAuthorityChoices();
        }
    }

    public OutlookAccountChoice? SelectedTargetAccount
    {
        get => _selectedTargetAccount;
        set
        {
            if (SetProperty(ref _selectedTargetAccount, value))
            {
                NotifyCommandStateChanged();
            }
        }
    }

    public AuthorityChoice? SelectedAuthority
    {
        get => _selectedAuthority;
        set
        {
            if (!SetProperty(ref _selectedAuthority, value))
            {
                return;
            }

            var group = SelectedAlignmentGroup;
            if (group is null)
            {
                AuthorityStatus = "Select a logical event to inspect authority.";
                return;
            }

            if (value is null)
            {
                AuthorityStatus = group.CanChooseAuthority
                    ? "Authority is unknown. Choose the account that should control alignment."
                    : BuildUnavailableAuthorityStatus(group);
                return;
            }

            _authorityByGroupKey[group.GroupKey] = value.AccountKey;
            AuthorityStatus = $"Authority: {value.AccountKey} · User selected for this session.";
        }
    }

    public double ScanDays
    {
        get => _scanDays;
        set
        {
            var normalized = double.IsNaN(value) ? 90 : Math.Clamp(value, 1, 3650);
            SetProperty(ref _scanDays, normalized);
        }
    }

    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            if (SetProperty(ref _isBusy, value))
            {
                NotifyCommandStateChanged();
            }
        }
    }

    public bool HasNotice
    {
        get => _hasNotice;
        private set => SetProperty(ref _hasNotice, value);
    }

    public string Status
    {
        get => _status;
        private set => SetProperty(ref _status, value);
    }

    public string ForwardCapability
    {
        get => _forwardCapability;
        private set => SetProperty(ref _forwardCapability, value);
    }

    public string LastDiagnostics
    {
        get => _lastDiagnostics;
        private set => SetProperty(ref _lastDiagnostics, value);
    }

    public string SelectedEventDiagnostics
    {
        get => _selectedEventDiagnostics;
        private set => SetProperty(ref _selectedEventDiagnostics, value);
    }

    public string AlignmentSummary
    {
        get => _alignmentSummary;
        private set => SetProperty(ref _alignmentSummary, value);
    }

    public string AuthorityStatus
    {
        get => _authorityStatus;
        private set => SetProperty(ref _authorityStatus, value);
    }

    public string NoticeMessage
    {
        get => _noticeMessage;
        private set => SetProperty(ref _noticeMessage, value);
    }

    private async Task RefreshAsync()
    {
        IsBusy = true;
        ClearNotice();
        Status = "Reading Outlook accounts and calendars…";
        ForwardCapability = "Not checked";

        try
        {
            var days = (int)Math.Clamp(Math.Round(ScanDays), 1, 3650);
            var result = await _hostClient.ScanAsync(days);

            Accounts.Clear();
            foreach (var account in result.Accounts)
            {
                Accounts.Add(new OutlookAccountChoice(account));
            }

            Events.Clear();
            var eventRows = result.Accounts
                .SelectMany(account => account.Events.Select(calendarEvent => new CalendarEventRowViewModel(account, calendarEvent)))
                .OrderBy(row => row.CalendarEvent.StartLocal)
                .ThenBy(row => row.Subject, StringComparer.CurrentCultureIgnoreCase)
                .ToArray();

            foreach (var eventRow in eventRows)
            {
                Events.Add(eventRow);
            }

            RebuildAlignmentGroups(result, eventRows);
            SelectedEvent = Events.FirstOrDefault();
            LastDiagnostics = BuildScanDiagnostics(result);
            Status = $"Loaded {Events.Count} events from {Accounts.Count} Outlook accounts for {days} days.";

            var failedAccounts = result.Accounts.Count(account => !string.IsNullOrWhiteSpace(account.Error));
            if (failedAccounts > 0 || result.Warnings.Count > 0)
            {
                ShowNotice($"Outlook was loaded with partial results ({failedAccounts} account errors, {result.Warnings.Count} warnings). See Diagnostics for details.");
            }
        }
        catch (Exception exception) when (exception is InvalidOperationException or FileNotFoundException or JsonException)
        {
            Status = "Outlook refresh failed.";
            LastDiagnostics = exception.Message;
            AlignmentSummary = "Alignment could not be rebuilt because Outlook refresh failed.";
            ShowNotice("Outlook calendars could not be refreshed. Technical details are available in Diagnostics.");
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task ProbeForwardAsync()
    {
        if (!TryGetForwardSelection(out var sourceSmtp, out var globalId, out var entryId))
        {
            return;
        }

        IsBusy = true;
        ClearNotice();
        Status = $"Checking native Forward for {SelectedEvent!.Subject}…";

        try
        {
            var result = await _hostClient.ProbeForwardAsync(sourceSmtp, globalId, entryId);
            LastDiagnostics = result.DiagnosticText;
            ForwardCapability = result.Success ? "Available" : "Unavailable";
            Status = result.Success
                ? "Native Outlook Forward is available for the selected meeting."
                : "Native Outlook Forward is unavailable for the selected meeting.";

            if (!result.Success)
            {
                ShowNotice("Outlook does not currently expose a safe native Forward action for this meeting. Nothing was changed.");
            }
        }
        catch (Exception exception) when (exception is InvalidOperationException or FileNotFoundException)
        {
            ForwardCapability = "Check failed";
            Status = "Forward capability check failed.";
            LastDiagnostics = exception.Message;
            ShowNotice("Forward capability could not be checked. Nothing was changed; see Diagnostics for details.");
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task PrepareForwardAsync()
    {
        if (!TryGetForwardSelection(out var sourceSmtp, out var globalId, out var entryId)
            || string.IsNullOrWhiteSpace(SelectedTargetAccount?.SmtpAddress))
        {
            return;
        }

        var recipient = SelectedTargetAccount.SmtpAddress;
        IsBusy = true;
        ClearNotice();
        Status = $"Checking and preparing native Forward to {recipient}…";

        try
        {
            var capability = await _hostClient.ProbeForwardAsync(sourceSmtp, globalId, entryId);
            if (!capability.Success)
            {
                ForwardCapability = "Unavailable";
                LastDiagnostics = capability.DiagnosticText;
                Status = "Native Outlook Forward is unavailable; nothing was prepared.";
                ShowNotice("Outlook no longer reports Forward as available for the selected meeting. The prepare action stopped safely.");
                return;
            }

            ForwardCapability = "Available";
            var prepare = await _hostClient.PrepareForwardAsync(sourceSmtp, globalId, entryId, recipient);
            LastDiagnostics = prepare.DiagnosticText;
            Status = prepare.Success
                ? $"Native Forward prepared for {recipient} and discarded unsent."
                : "Forward preparation failed safely.";

            if (!prepare.Success)
            {
                ShowNotice("The native Forward could not be prepared safely. It was not sent; see Diagnostics for details.");
            }
        }
        catch (Exception exception) when (exception is InvalidOperationException or FileNotFoundException)
        {
            Status = "Forward preparation failed.";
            LastDiagnostics = exception.Message;
            ShowNotice("Forward preparation failed before completion. Nothing was intentionally sent; see Diagnostics for details.");
        }
        finally
        {
            IsBusy = false;
        }
    }

    private bool CanProbeForward()
        => !IsBusy && SelectedEvent?.HasForwardIdentity == true;

    private bool CanPrepareForward()
        => CanProbeForward() && !string.IsNullOrWhiteSpace(SelectedTargetAccount?.SmtpAddress);

    private bool TryGetForwardSelection(out string sourceSmtp, out string globalId, out string entryId)
    {
        sourceSmtp = SelectedEvent?.Account.SmtpAddress ?? string.Empty;
        globalId = SelectedEvent?.CalendarEvent.GlobalAppointmentId ?? string.Empty;
        entryId = SelectedEvent?.CalendarEvent.EntryId ?? string.Empty;

        if (!string.IsNullOrWhiteSpace(sourceSmtp)
            && !string.IsNullOrWhiteSpace(globalId)
            && !string.IsNullOrWhiteSpace(entryId))
        {
            return true;
        }

        Status = "The selected item does not expose enough Outlook identity for native Forward.";
        ForwardCapability = "Unavailable";
        ShowNotice("This item cannot be forwarded safely because Outlook did not provide the required native meeting identity.");
        return false;
    }

    private void UpdateTargetAccounts()
    {
        var previousSmtp = SelectedTargetAccount?.SmtpAddress;
        var sourceSmtp = SelectedEvent?.Account.SmtpAddress;

        TargetAccounts.Clear();
        foreach (var account in Accounts.Where(account =>
                     !string.IsNullOrWhiteSpace(account.SmtpAddress)
                     && !string.Equals(account.SmtpAddress, sourceSmtp, StringComparison.OrdinalIgnoreCase)))
        {
            TargetAccounts.Add(account);
        }

        SelectedTargetAccount = TargetAccounts.FirstOrDefault(account =>
                                    string.Equals(account.SmtpAddress, previousSmtp, StringComparison.OrdinalIgnoreCase))
                                ?? TargetAccounts.FirstOrDefault();
    }

    private void UpdateAuthorityChoices()
    {
        AuthorityChoices.Clear();
        _selectedAuthority = null;
        OnPropertyChanged(nameof(SelectedAuthority));

        var group = SelectedAlignmentGroup;
        if (group is null)
        {
            AuthorityStatus = "Select a logical event to inspect authority.";
            return;
        }

        if (!group.CanChooseAuthority)
        {
            AuthorityStatus = BuildUnavailableAuthorityStatus(group);
            return;
        }

        foreach (var accountKey in group.Group.Members
                     .Select(member => member.AccountKey)
                     .Distinct(StringComparer.OrdinalIgnoreCase)
                     .OrderBy(account => account, StringComparer.OrdinalIgnoreCase))
        {
            AuthorityChoices.Add(new AuthorityChoice(accountKey));
        }

        if (_authorityByGroupKey.TryGetValue(group.GroupKey, out var savedAccountKey))
        {
            SelectedAuthority = AuthorityChoices.FirstOrDefault(choice =>
                string.Equals(choice.AccountKey, savedAccountKey, StringComparison.OrdinalIgnoreCase));
        }

        if (SelectedAuthority is null)
        {
            AuthorityStatus = "Authority is unknown. Choose the account that should control alignment.";
        }
    }

    private void RebuildAlignmentGroups(
        OutlookProbeResult result,
        CalendarEventRowViewModel[] eventRows)
    {
        var expectedAccounts = result.Accounts
            .Where(account => account.CalendarAvailable && !string.IsNullOrWhiteSpace(account.SmtpAddress))
            .Select(account => account.SmtpAddress!)
            .ToArray();

        var observations = eventRows.Select(row => new ObservedCalendarEvent(
            row.Account.SmtpAddress ?? row.Account.DisplayName,
            row.CalendarEvent.EntryId ?? $"unlocated:{row.CalendarEvent.StartLocal:O}:{row.Subject}",
            row.CalendarEvent.GlobalAppointmentId,
            row.CalendarEvent.StartLocal,
            row.CalendarEvent.EndLocal,
            row.CalendarEvent.IsAllDay,
            row.CalendarEvent.IsRecurring,
            row.CalendarEvent.BusyStatus,
            row.CalendarEvent.Sensitivity,
            row.CalendarEvent.Subject,
            row.CalendarEvent.Location));

        var groups = EventCorrelation.BuildGroups(observations, expectedAccounts);

        AlignmentGroups.Clear();
        foreach (var group in groups
                     .OrderBy(group => group.State == AlignmentState.Aligned)
                     .ThenBy(group => group.Members.Min(member => member.StartLocal)))
        {
            AlignmentGroups.Add(new AlignmentGroupViewModel(group));
        }

        SelectedAlignmentGroup = AlignmentGroups.FirstOrDefault(group => group.NeedsAttention)
                                 ?? AlignmentGroups.FirstOrDefault();

        var attention = groups.Count(group => group.State != AlignmentState.Aligned);
        var aligned = groups.Count - attention;
        AlignmentSummary = $"{groups.Count} logical rows: {attention} need attention, {aligned} currently aligned. Recurring items remain deliberately unresolved.";
    }

    private void ClearNotice()
    {
        NoticeMessage = string.Empty;
        HasNotice = false;
    }

    private void ShowNotice(string message)
    {
        NoticeMessage = message;
        HasNotice = true;
    }

    private void NotifyCommandStateChanged()
    {
        RefreshCommand.NotifyCanExecuteChanged();
        ProbeForwardCommand.NotifyCanExecuteChanged();
        PrepareForwardCommand.NotifyCanExecuteChanged();
    }

    private static string BuildScanDiagnostics(OutlookProbeResult result)
    {
        var accountLines = result.Accounts.Select(account =>
            $"{account.DisplayName}: {account.EventCount} events; calendar={(account.CalendarAvailable ? "available" : "unavailable")}; error={account.Error ?? "none"}");
        var warningLines = result.Warnings.Count == 0
            ? ["Warnings: none"]
            : result.Warnings.Select(warning => "Warning: " + warning);

        return string.Join(Environment.NewLine, accountLines.Concat(warningLines));
    }

    private static string BuildSelectedEventDiagnostics(CalendarEventRowViewModel? selectedEvent)
    {
        if (selectedEvent is null)
        {
            return "No calendar item is selected.";
        }

        return string.Join(
            Environment.NewLine,
            $"Subject: {selectedEvent.Subject}",
            $"Source account: {selectedEvent.SourceSmtp}",
            $"GlobalAppointmentID: {selectedEvent.CalendarEvent.GlobalAppointmentId ?? "(unavailable)"}",
            $"EntryID: {selectedEvent.CalendarEvent.EntryId ?? "(unavailable)"}",
            $"StoreID: {selectedEvent.Account.StoreId ?? "(unavailable)"}",
            $"Recurrence: {selectedEvent.RecurrenceDisplay}");
    }

    private static string BuildUnavailableAuthorityStatus(AlignmentGroupViewModel group)
        => group.Group.State switch
        {
            AlignmentState.RecurrenceIdentityUnresolved => "Authority is disabled until recurrence identity is proven.",
            AlignmentState.Uncorrelated => "Authority is disabled because this item is not safely correlated.",
            AlignmentState.Duplicate => "Authority is disabled while duplicate copies are unresolved.",
            _ => "Authority cannot be selected for this logical event yet.",
        };
}
