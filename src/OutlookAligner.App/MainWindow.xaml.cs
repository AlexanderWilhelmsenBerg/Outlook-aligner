using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using OutlookAligner.App.ViewModels;

namespace OutlookAligner.App;

public sealed partial class MainWindow : Window
{
    private bool _initialRefreshStarted;

    public MainWindow()
    {
        InitializeComponent();
        ViewModel = new MainViewModel();
        RootLayout.DataContext = ViewModel;
        RootNavigation.SelectedItem = RootNavigation.MenuItems[0];
    }

    public MainViewModel ViewModel { get; }

    private void OnRootLoaded(object sender, RoutedEventArgs e)
    {
        _ = sender;
        _ = e;

        if (_initialRefreshStarted)
        {
            return;
        }

        _initialRefreshStarted = true;
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
        DiagnosticsPanel.Visibility = tag == "diagnostics" ? Visibility.Visible : Visibility.Collapsed;

        var isPlaceholder = tag is "alignment" or "settings";
        PlaceholderPanel.Visibility = isPlaceholder ? Visibility.Visible : Visibility.Collapsed;
        PlaceholderTitle.Text = tag switch
        {
            "alignment" => "Alignment",
            "settings" => "Settings",
            _ => string.Empty,
        };
    }
}
