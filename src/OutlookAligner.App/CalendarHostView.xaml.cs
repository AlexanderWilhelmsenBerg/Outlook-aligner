using System.ComponentModel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.Web.WebView2.Core;
using OutlookAligner.App.Diagnostics;
using OutlookAligner.App.Presentation;
using OutlookAligner.App.ViewModels;

namespace OutlookAligner.App;

public sealed partial class CalendarHostView : UserControl, IDisposable
{
    private const string CalendarVirtualHost = "calendar.outlook-aligner.local";

    private readonly AppEventLog _eventLog;
    private readonly CalendarPresentationSession _presentationSession = new();
    private readonly MainViewModel _viewModel;
    private bool _disposed;
    private bool _webReady;

    public CalendarHostView(MainViewModel viewModel, AppEventLog eventLog)
    {
        InitializeComponent();
        _viewModel = viewModel;
        _eventLog = eventLog;
        DataContext = viewModel;
        Loaded += OnLoaded;
        _viewModel.PropertyChanged += OnViewModelPropertyChanged;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        Loaded -= OnLoaded;
        _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
        if (CalendarWebView.CoreWebView2 is not null)
        {
            CalendarWebView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
            CalendarWebView.CoreWebView2.NavigationStarting -= OnNavigationStarting;
        }
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        _ = sender;
        _ = e;
        Loaded -= OnLoaded;
        await InitializeWebAsync();
    }

    private async Task InitializeWebAsync()
    {
        try
        {
            var assetRoot = Path.Combine(AppContext.BaseDirectory, "CalendarWeb");
            var indexPath = Path.Combine(assetRoot, "index.html");
            if (!File.Exists(indexPath))
            {
                throw new FileNotFoundException("Packaged calendar index.html was not found.", indexPath);
            }

            await CalendarWebView.EnsureCoreWebView2Async();
            CalendarWebView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                CalendarVirtualHost,
                assetRoot,
                CoreWebView2HostResourceAccessKind.Allow);
            CalendarWebView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
            CalendarWebView.CoreWebView2.NavigationStarting += OnNavigationStarting;
            CalendarWebView.Source = new Uri($"https://{CalendarVirtualHost}/index.html");
            _eventLog.Debug("Calendar", "WebView2 month calendar initialized from packaged assets.");
        }
        catch (Exception ex)
        {
            ShowWebFailure(ex);
        }
    }

    private void OnNavigationStarting(CoreWebView2 sender, CoreWebView2NavigationStartingEventArgs args)
    {
        _ = sender;
        if (!Uri.TryCreate(args.Uri, UriKind.Absolute, out var uri)
            || !string.Equals(uri.Host, CalendarVirtualHost, StringComparison.OrdinalIgnoreCase))
        {
            args.Cancel = true;
            _eventLog.Warning("Calendar", "Blocked unexpected WebView top-level navigation.", args.Uri);
        }
    }

    private void OnWebMessageReceived(CoreWebView2 sender, CoreWebView2WebMessageReceivedEventArgs args)
    {
        _ = sender;
        string json;
        try
        {
            json = args.TryGetWebMessageAsString();
        }
        catch (Exception ex)
        {
            _eventLog.Warning("Calendar", "Ignored a non-string WebView message.", details: ex.ToString());
            return;
        }

        if (!CalendarWebProtocol.TryParseClientMessage(json, out var message) || message is null)
        {
            _eventLog.Warning("Calendar", "Ignored malformed or unsupported calendar WebView message.");
            return;
        }

        switch (message.Kind)
        {
            case CalendarClientMessageKind.Ready:
                _webReady = true;
                CalendarWebFallback.Visibility = Visibility.Collapsed;
                TodayButton.IsEnabled = true;
                PreviousMonthButton.IsEnabled = true;
                NextMonthButton.IsEnabled = true;
                SendCurrentObservations();
                break;
            case CalendarClientMessageKind.RangeChanged:
                MonthTitle.Text = message.Title ?? "Month";
                EmptyMessage.Visibility = message.VisibleCount == 0 && !_viewModel.IsBusy
                    ? Visibility.Visible
                    : Visibility.Collapsed;
                break;
            case CalendarClientMessageKind.ObservationSelected:
                if (message.PresentationId is not null
                    && _presentationSession.TryResolve(message.PresentationId, out var selected)
                    && selected is not null)
                {
                    _viewModel.SelectedEvent = selected;
                    _eventLog.Debug("Calendar", "Calendar observation selected.", BuildEventContext(selected));
                }
                else
                {
                    _eventLog.Debug("Calendar", "Ignored stale or unknown calendar presentation ID.");
                }

                break;
        }
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        _ = sender;
        if (e.PropertyName == nameof(MainViewModel.IsBusy) && !_viewModel.IsBusy)
        {
            SendCurrentObservations();
        }
    }

    private void SendCurrentObservations()
    {
        if (!_webReady || CalendarWebView.CoreWebView2 is null)
        {
            return;
        }

        var observations = _presentationSession.Reset(_viewModel.Events);
        CalendarWebView.CoreWebView2.PostWebMessageAsString(CalendarWebProtocol.SerializeRender(observations));
        EmptyMessage.Visibility = observations.Count == 0 && !_viewModel.IsBusy
            ? Visibility.Visible
            : Visibility.Collapsed;
    }

    private void SendNavigation(string action)
    {
        if (_webReady && CalendarWebView.CoreWebView2 is not null)
        {
            CalendarWebView.CoreWebView2.PostWebMessageAsString(CalendarWebProtocol.SerializeNavigation(action));
        }
    }

    private void OnTodayClicked(object sender, RoutedEventArgs e)
    {
        _ = sender;
        _ = e;
        SendNavigation("today");
    }

    private void OnPreviousClicked(object sender, RoutedEventArgs e)
    {
        _ = sender;
        _ = e;
        SendNavigation("previous");
    }

    private void OnNextClicked(object sender, RoutedEventArgs e)
    {
        _ = sender;
        _ = e;
        SendNavigation("next");
    }

    private void ShowWebFailure(Exception ex)
    {
        _webReady = false;
        TodayButton.IsEnabled = false;
        PreviousMonthButton.IsEnabled = false;
        NextMonthButton.IsEnabled = false;
        CalendarWebFallback.Visibility = Visibility.Visible;
        _eventLog.Error("Calendar", "WebView2 month calendar failed.", details: ex.ToString());
    }

    private static string BuildEventContext(CalendarEventRowViewModel selected)
        => $"{selected.Subject} | {selected.SourceSmtp} | {selected.TimeDisplay} | {selected.MeetingStatusDisplay} | "
            + (selected.CalendarEvent.IsRecurring ? "recurring" : "single");
}
