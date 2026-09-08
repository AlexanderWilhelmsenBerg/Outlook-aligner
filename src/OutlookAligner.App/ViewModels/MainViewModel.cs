using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OutlookAligner.App.Services;
using OutlookAligner.Outlook.Contracts;

namespace OutlookAligner.App.ViewModels;

public sealed class MainViewModel : ObservableObject
{
    private readonly OutlookHostProcessClient _hostClient = new();
    private CalendarEventRowViewModel? _selectedEvent;
    private OutlookAccountChoice? _selectedTargetAccount;
    private double _scanDays = 90;
    private bool _isBusy;
    private string _status = "Ready to read Classic Outlook.";
    private string _forwardCapability = "Not checked";
    private string _lastDiagnostics = "No diagnostics yet.";

    public MainViewModel()
    {
        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => !IsBusy);
        ProbeForwardCommand = new AsyncRelayCommand(ProbeForwardAsync, CanProbeForward);
        PrepareForwardCommand = new AsyncRelayCommand(PrepareForwardAsync, CanPrepareForward);
    }

    public ObservableCollection<CalendarEventRowViewModel> Events { get; } = [];

    public ObservableCollection<OutlookAccountChoice> Accounts { get; } = [];

    public ObservableCollection<OutlookAccountChoice> TargetAccounts { get; } = [];

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
            UpdateTargetAccounts();
            NotifyCommandStateChanged();
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

    private async Task RefreshAsync()
    {
        IsBusy = true;
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
                .ThenBy(row => row.Subject, StringComparer.CurrentCultureIgnoreCase);

            foreach (var eventRow in eventRows)
            {
                Events.Add(eventRow);
            }

            SelectedEvent = Events.FirstOrDefault();
            LastDiagnostics = BuildScanDiagnostics(result);
            Status = $"Loaded {Events.Count} events from {Accounts.Count} Outlook accounts for {days} days.";
        }
        catch (Exception exception) when (exception is InvalidOperationException or FileNotFoundException or JsonException)
        {
            Status = "Outlook refresh failed. Open Diagnostics for details.";
            LastDiagnostics = exception.Message;
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
        Status = $"Checking native Forward for {SelectedEvent!.Subject}…";

        try
        {
            var result = await _hostClient.ProbeForwardAsync(sourceSmtp, globalId, entryId);
            LastDiagnostics = result.DiagnosticText;
            ForwardCapability = result.Success ? "Available" : "Unavailable";
            Status = result.Success
                ? "Native Outlook Forward is available for the selected meeting."
                : "Native Outlook Forward is unavailable for the selected meeting.";
        }
        catch (Exception exception) when (exception is InvalidOperationException or FileNotFoundException)
        {
            ForwardCapability = "Check failed";
            Status = "Forward capability check failed. Open Diagnostics for details.";
            LastDiagnostics = exception.Message;
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
        Status = $"Checking and preparing native Forward to {recipient}…";

        try
        {
            var capability = await _hostClient.ProbeForwardAsync(sourceSmtp, globalId, entryId);
            if (!capability.Success)
            {
                ForwardCapability = "Unavailable";
                LastDiagnostics = capability.DiagnosticText;
                Status = "Native Outlook Forward is unavailable; nothing was prepared.";
                return;
            }

            ForwardCapability = "Available";
            var prepare = await _hostClient.PrepareForwardAsync(sourceSmtp, globalId, entryId, recipient);
            LastDiagnostics = prepare.DiagnosticText;
            Status = prepare.Success
                ? $"Native Forward prepared for {recipient} and discarded unsent."
                : "Forward preparation failed safely. Open Diagnostics for details.";
        }
        catch (Exception exception) when (exception is InvalidOperationException or FileNotFoundException)
        {
            Status = "Forward preparation failed. Open Diagnostics for details.";
            LastDiagnostics = exception.Message;
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
}
