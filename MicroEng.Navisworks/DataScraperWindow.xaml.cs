using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using Autodesk.Navisworks.Api;
using WpfUiControls = Wpf.Ui.Controls;

namespace MicroEng.Navisworks
{
    public partial class DataScraperWindow : Window
    {
        static DataScraperWindow()
        {
            AssemblyResolver.EnsureRegistered();
        }

        private readonly DataScraperService _service = new();
        private List<ScrapedPropertyView> _currentProperties = new();
        private ScrapeScopeType _lastScope = ScrapeScopeType.EntireModel;
        private string _lastSetName;
        private bool? _lastSetIsSearch;
        private List<RawRow> _rawRows = new();
        private ScrapeSession _selectedSession;
        private Guid _loadedRawSessionId = Guid.Empty;
        private bool _rawEntriesTruncated;
        private bool _rawEntriesStored;
        private Brush _statusReadyBrush;
        private ICollectionView _propertiesView;
        private ICollectionView _rawRowsView;
        private List<SetListItem> _availableSets = new();
        private bool _updatingSetLists;

        public DataScraperWindow(string initialProfile = null)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(this);
            ApplyRunScrapeButtonTheme();
            MicroEngWpfUiTheme.ThemeChanged += OnThemeChanged;
            DataScraperCache.CacheChanged += OnCacheChanged;
            Closed += (_, __) =>
            {
                MicroEngWpfUiTheme.ThemeChanged -= OnThemeChanged;
                DataScraperCache.CacheChanged -= OnCacheChanged;
            };
            _statusReadyBrush = StatusText?.Foreground;
            try
            {
                if (SetSearchMenu != null)
                {
                    SetSearchMenu.Opened += (s, e) => LoadSets();
                    SetSearchMenu.AddHandler(MenuItem.ClickEvent, new RoutedEventHandler(OnSetSearchMenuItemClick));
                }

                SelectedRunSummary.Text = "Select a run to view data.";
                if (!string.IsNullOrWhiteSpace(initialProfile))
                {
                    ProfileNameBox.Text = initialProfile;
                }
                LoadSets();
                RefreshHistory();
                LoadProfiles();
                if (!string.IsNullOrWhiteSpace(ProfileNameBox.Text))
                {
                    SelectLatestRunForProfile(ProfileNameBox.Text.Trim());
                }
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"DataScraper init failed: {ex.Message}");
                StatusText.Text = ex.Message;
            }
        }

        private void OnCacheChanged()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke((Action)OnCacheChanged);
                return;
            }

            RefreshHistory();
            LoadProfiles();
            var last = DataScraperCache.LastSession;
            if (last != null)
            {
                HistoryList.SelectedItem = last;
                ShowProperties(last);
            }
            else
            {
                HistoryList.SelectedItem = null;
                PropertiesGrid.ItemsSource = null;
                RawDataGrid.ItemsSource = null;
                _selectedSession = null;
                _loadedRawSessionId = Guid.Empty;
                PropertySummary.Text = "0 properties.";
                RawSummaryText.Text = "0 raw entries.";
                SelectedRunSummary.Text = "Select a run to view data.";
            }
        }

        private void OnThemeChanged(MicroEngThemeMode theme)
        {
            if (Dispatcher.CheckAccess())
            {
                ApplyRunScrapeButtonTheme();
            }
            else
            {
                Dispatcher.BeginInvoke((Action)ApplyRunScrapeButtonTheme);
            }
        }

        private void ApplyRunScrapeButtonTheme()
        {
            if (RunScrapeButton == null)
            {
                return;
            }

            var isDark = MicroEngWpfUiTheme.CurrentTheme == MicroEngThemeMode.Dark;
            var background = isDark ? Brushes.White : Brushes.Black;
            var foreground = isDark ? Brushes.Black : Brushes.White;

            RunScrapeButton.Background = background;
            RunScrapeButton.BorderBrush = background;
            RunScrapeButton.Foreground = foreground;

            if (RunScrapeLabel != null)
            {
                RunScrapeLabel.Foreground = foreground;
            }
        }

        private void LoadSets()
        {
            try
            {
                var lists = GetSelectionSetLists();
                _availableSets = lists.Sets;

                var selectedName = SetSearchNameBox?.Text?.Trim();
                if (!string.IsNullOrWhiteSpace(selectedName))
                {
                    var match = _availableSets.FirstOrDefault(s => string.Equals(s.Name, selectedName, StringComparison.OrdinalIgnoreCase));
                    if (match != null)
                    {
                        _lastSetIsSearch = match.IsSearch;
                        _lastSetName = match.Name;
                    }
                }

                _updatingSetLists = true;
                try
                {
                    UpdateMenuItems(SetSearchMenu, _availableSets, selectedName, _lastSetIsSearch);
                }
                finally
                {
                    _updatingSetLists = false;
                }
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"DataScraper: failed to load sets: {ex.Message}");
            }
        }

        private void RefreshHistory()
        {
            HistoryList.ItemsSource = DataScraperCache.AllSessions
                .OrderByDescending(s => s.Timestamp)
                .ToList();
        }

        private void LoadProfiles()
        {
            var profiles = DataScraperCache.AllSessions
                .Select(s => string.IsNullOrWhiteSpace(s.ProfileName) ? "Default" : s.ProfileName.Trim())
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                .ToList();

            ExistingProfileCombo.ItemsSource = profiles;

            if (!profiles.Any())
            {
                ExistingProfileCombo.SelectedItem = null;
                return;
            }

            var desired = ProfileNameBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(desired))
            {
                ExistingProfileCombo.SelectedItem = profiles[0];
                return;
            }

            var match = profiles.FirstOrDefault(p => string.Equals(p, desired, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                ExistingProfileCombo.SelectedItem = match;
            }
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            SetRunScrapeLoading(true);
            FlushUi();
            try
            {
                var profile = string.IsNullOrWhiteSpace(ProfileNameBox.Text) ? "Default" : ProfileNameBox.Text.Trim();
                var scope = GetScopeType();
                var selectionSetName = string.Empty;
                var searchSetName = string.Empty;

                if (scope == ScrapeScopeType.SelectionSet || scope == ScrapeScopeType.SearchSet)
                {
                    var setName = _lastSetName ?? SetSearchNameBox?.Text?.Trim();
                    if (scope == ScrapeScopeType.SearchSet)
                    {
                        searchSetName = setName ?? string.Empty;
                    }
                    else
                    {
                        selectionSetName = setName ?? string.Empty;
                    }
                }

                var items = _service.ResolveScope(scope, selectionSetName, searchSetName, out var description).ToList();
                if (!items.Any())
                {
                    MessageBox.Show("No items found for the selected scope.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                MicroEngActions.Log($"[DataScraper] Run start. Profile='{profile}', Scope={scope}, Items={items.Count}");

                StatusText.Text = "Scraping...";
                var session = _service.Scrape(profile, scope, description, items);

                StatusText.Text = $"Scraped {session.ItemsScanned} items, {session.Properties.Count} properties.";
                var logMessage = $"Data Scraper: {session.ItemsScanned} items scanned in {description}. {session.Properties.Count} properties cached for profile '{profile}'.";
                MicroEngActions.Log(logMessage);
                RefreshHistory();
                LoadProfiles();
                SelectLatestRunForProfile(profile);
                ShowProperties(session);
                MicroEngActions.Log($"[DataScraper] Run complete. Items={session.ItemsScanned}, Properties={session.Properties.Count}");
                FlashSuccess(RunScrapeButton);
                ShowSnackbar("Scrape complete", $"{session.ItemsScanned} items, {session.Properties.Count} properties cached.", WpfUiControls.ControlAppearance.Success, WpfUiControls.SymbolRegular.CheckmarkCircle24);
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"[DataScraper] Run failed: {ex}");
                StatusText.Text = ex.Message;
                MessageBox.Show(ex.Message, "MicroEng", MessageBoxButton.OK, MessageBoxImage.Error);
                ShowSnackbar("Scrape failed", ex.Message, WpfUiControls.ControlAppearance.Danger, WpfUiControls.SymbolRegular.ErrorCircle24);
            }
            finally
            {
                SetRunScrapeLoading(false);
            }
        }

        private ScrapeScopeType GetScopeType()
        {
            if (SingleItemRadio.IsChecked == true) return ScrapeScopeType.SingleItem;
            if (CurrentSelectionRadio.IsChecked == true) return ScrapeScopeType.CurrentSelection;
            if (SetSearchRadio.IsChecked == true)
            {
                return _lastSetIsSearch == true ? ScrapeScopeType.SearchSet : ScrapeScopeType.SelectionSet;
            }
            if (EntireModelRadio.IsChecked == true) return ScrapeScopeType.EntireModel;
            return ScrapeScopeType.EntireModel;
        }

        private void HistoryList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (HistoryList.SelectedItem is ScrapeSession session)
            {
                ShowProperties(session);
            }
        }

        private void ShowProperties(ScrapeSession session)
        {
            if (_selectedSession != null && _selectedSession.Id != session.Id)
            {
                DataScraperCache.ReleaseRawEntries(_selectedSession);
            }

            _selectedSession = session;

            _currentProperties = session.Properties
                .Select(p => new ScrapedPropertyView(p))
                .ToList();
            _propertiesView = CollectionViewSource.GetDefaultView(_currentProperties);
            PropertiesGrid.ItemsSource = _propertiesView;
            PropertySummary.Text = $"{_currentProperties.Count} properties from session {session.Timestamp:T}";

            _loadedRawSessionId = Guid.Empty;
            _rawRows = new List<RawRow>();
            _rawRowsView = null;
            RawDataGrid.ItemsSource = null;

            _rawEntriesStored = session.RawEntryCount > 0;
            _rawEntriesTruncated = session.RawEntriesTruncated;

            RawDataNotStoredHint.Visibility = Visibility.Collapsed;
            if (_rawEntriesTruncated)
            {
                RawDataNotStoredHint.Text = _rawEntriesStored
                    ? "Raw rows were truncated."
                    : "Raw rows were not stored in memory.";
                RawDataNotStoredHint.Visibility = Visibility.Visible;
            }

            var rawSummary = $"{session.RawEntryCount} raw entries";
            if (_rawEntriesTruncated)
            {
                rawSummary += _rawEntriesStored ? " (preview)" : " (not stored)";
            }
            RawSummaryText.Text = rawSummary + ".";

            SelectedRunSummary.Text = $"{session.ProfileName} - {session.ScopeType} - {session.ItemsScanned} items - {session.Timestamp:G}";

            if (RawDataTab?.IsSelected == true)
            {
                EnsureRawRowsLoaded(session);
            }
        }

        private void EnsureRawRowsLoaded(ScrapeSession session)
        {
            if (session == null)
            {
                return;
            }

            if (_loadedRawSessionId == session.Id && _rawRowsView != null)
            {
                return;
            }

            var rawEntries = session.RawEntries ?? new List<RawEntry>();
            _rawRows = rawEntries
                .Select(r => new RawRow
                {
                    Profile = r.Profile,
                    Scope = r.Scope,
                    ItemPath = r.ItemPath,
                    Category = r.Category,
                    Name = r.Name,
                    DataType = r.DataType,
                    Value = r.Value
                }).ToList();
            _rawRowsView = CollectionViewSource.GetDefaultView(_rawRows);
            RawDataGrid.ItemsSource = _rawRowsView;
            _loadedRawSessionId = session.Id;

            _rawEntriesStored = _rawRows.Count > 0;

            RawDataNotStoredHint.Visibility = Visibility.Collapsed;
            if (_rawEntriesTruncated)
            {
                RawDataNotStoredHint.Text = _rawEntriesStored
                    ? "Raw rows were truncated."
                    : "Raw rows were not stored in memory.";
                RawDataNotStoredHint.Visibility = Visibility.Visible;
            }
            else if (session.RawEntryCount > 0 && _rawRows.Count == 0)
            {
                RawDataNotStoredHint.Text = "Raw rows are unavailable for this run.";
                RawDataNotStoredHint.Visibility = Visibility.Visible;
            }

            var rawSummary = $"{_rawRows.Count} raw entries";
            if (_rawEntriesTruncated)
            {
                rawSummary += _rawEntriesStored ? " (preview)" : " (not stored)";
            }
            RawSummaryText.Text = rawSummary + ".";
        }

        private void DataViewTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!ReferenceEquals(e.Source, DataViewTabControl))
            {
                return;
            }

            if (RawDataTab?.IsSelected == true && _selectedSession != null)
            {
                EnsureRawRowsLoaded(_selectedSession);
            }
        }

        private void FilterBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_propertiesView == null) return;
            var text = FilterBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(text))
            {
                _propertiesView.Filter = null;
            }
            else
            {
                _propertiesView.Filter = obj =>
                {
                    if (obj is ScrapedPropertyView p)
                    {
                        return (p.Name?.IndexOf(text, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0 ||
                               (p.Category?.IndexOf(text, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0;
                    }

                    return false;
                };
            }

            _propertiesView.Refresh();
        }

        private void RawFilterBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_rawRowsView == null) return;
            var text = RawFilterBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(text))
            {
                _rawRowsView.Filter = null;
            }
            else
            {
                _rawRowsView.Filter = obj =>
                {
                    if (obj is RawRow r)
                    {
                        return (r.Category?.IndexOf(text, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0
                               || (r.Name?.IndexOf(text, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0
                               || (r.Value?.IndexOf(text, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0
                               || (r.ItemPath?.IndexOf(text, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0;
                    }

                    return false;
                };
            }

            _rawRowsView.Refresh();
            var count = _rawRowsView.Cast<object>().Count();
            var summary = $"{count} raw entries";
            if (_rawEntriesTruncated)
            {
                summary += _rawEntriesStored ? " (preview)" : " (not stored)";
            }
            RawSummaryText.Text = summary + ".";
        }

        private void Rerun_Click(object sender, RoutedEventArgs e)
        {
            if (HistoryList.SelectedItem is ScrapeSession session)
            {
                _lastScope = Enum.TryParse<ScrapeScopeType>(session.ScopeType, out var scope) ? scope : ScrapeScopeType.EntireModel;
                Run_Click(sender, e);
            }
        }

        private void ViewProperties_Click(object sender, RoutedEventArgs e)
        {
            if (HistoryList.SelectedItem is ScrapeSession session)
            {
                ShowProperties(session);
            }
        }

        private void DeleteHistorySession_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not FrameworkElement element || element.DataContext is not ScrapeSession session)
            {
                return;
            }

            e.Handled = true;
            var profile = string.IsNullOrWhiteSpace(session.ProfileName) ? "Default" : session.ProfileName.Trim();
            var confirm = MessageBox.Show(
                $"Delete cached run '{profile}' from {session.Timestamp:G}?",
                "MicroEng",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            if (!DataScraperCache.RemoveSession(session.Id))
            {
                ShowSnackbar("Delete failed", "Could not remove the selected cached run.", WpfUiControls.ControlAppearance.Danger, WpfUiControls.SymbolRegular.ErrorCircle24);
                return;
            }

            ShowSnackbar("Cache removed", $"Removed '{profile}' run from {session.Timestamp:t}.", WpfUiControls.ControlAppearance.Success, WpfUiControls.SymbolRegular.CheckmarkCircle24);
        }

        private void RefreshSets_Click(object sender, RoutedEventArgs e)
        {
            LoadSets();
        }

        private void SelectLatestRunForProfile(string profile)
        {
            if (string.IsNullOrWhiteSpace(profile))
            {
                return;
            }

            var latest = DataScraperCache.AllSessions
                .Where(s => string.Equals(s.ProfileName ?? "Default", profile, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(s => s.Timestamp)
                .FirstOrDefault();

            if (latest != null)
            {
                HistoryList.SelectedItem = latest;
                HistoryList.ScrollIntoView(latest);
                ShowProperties(latest);
            }
        }

        private void OnSetSearchMenuItemClick(object sender, RoutedEventArgs e)
        {
            if (_updatingSetLists)
            {
                return;
            }

            if (e.OriginalSource is not MenuItem menuItem)
            {
                return;
            }

            if (menuItem.DataContext is not SetListItem item)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(item.Name))
            {
                return;
            }

            _lastSetName = item.Name;
            _lastSetIsSearch = item.IsSearch;
            SetSearchNameBox.Text = item.Name;
            SetSearchRadio.IsChecked = true;
        }

        private static void UpdateMenuItems(ContextMenu menu, List<SetListItem> items, string selectedName, bool? selectedIsSearch)
        {
            if (menu == null)
            {
                return;
            }

            var name = string.IsNullOrWhiteSpace(selectedName) ? null : selectedName.Trim();
            var hasName = !string.IsNullOrWhiteSpace(name);

            var menuItems = items?
                .Select(item => item.WithSelection(
                    hasName
                    && string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)
                    && (!selectedIsSearch.HasValue || item.IsSearch == selectedIsSearch.Value)))
                .ToList()
                ?? new List<SetListItem>();

            menu.ItemsSource = menuItems;
        }

        private sealed class SelectionSetLists
        {
            public List<SetListItem> Sets { get; } = new();
        }

        private sealed class SetListItem
        {
            public SetListItem(string name, bool isSearch, bool isSelected = false)
            {
                Name = name;
                IsSearch = isSearch;
                IsSelected = isSelected;
                DisplayName = $"{name} ({(isSearch ? "Search" : "Selection")})";
            }

            public string Name { get; }
            public bool IsSearch { get; }
            public bool IsSelected { get; }
            public string DisplayName { get; }

            public SetListItem WithSelection(bool isSelected)
            {
                return new SetListItem(Name, IsSearch, isSelected);
            }
        }

        private static SelectionSetLists GetSelectionSetLists()
        {
            var lists = new SelectionSetLists();
            var doc = Autodesk.Navisworks.Api.Application.ActiveDocument;

            foreach (var name in NavisworksSelectionSetUtils.GetSelectionSetNames(doc))
            {
                lists.Sets.Add(new SetListItem(name, isSearch: false));
            }

            foreach (var name in NavisworksSelectionSetUtils.GetSearchSetNames(doc))
            {
                lists.Sets.Add(new SetListItem(name, isSearch: true));
            }

            lists.Sets.Sort((a, b) => string.Compare(a.DisplayName, b.DisplayName, StringComparison.OrdinalIgnoreCase));
            return lists;
        }


        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void SetRunScrapeLoading(bool isLoading)
        {
            if (RunScrapeButton == null) return;
            RunScrapeButton.IsEnabled = !isLoading;
            if (RunScrapeLabel != null) RunScrapeLabel.Visibility = isLoading ? Visibility.Collapsed : Visibility.Visible;
            if (RunScrapeSpinnerIcon != null) RunScrapeSpinnerIcon.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
            if (StatusText != null)
            {
                if (isLoading)
                {
                    StatusText.Text = "Processing - Please Wait";
                    StatusText.Foreground = Brushes.DarkOrange;
                }
                else
                {
                    if (string.Equals(StatusText.Text, "Processing - Please Wait", StringComparison.OrdinalIgnoreCase))
                    {
                        StatusText.Text = "Ready";
                    }

                    if (_statusReadyBrush != null)
                    {
                        StatusText.Foreground = _statusReadyBrush;
                    }
                }
            }
        }

        private void FlashSuccess(System.Windows.Controls.Button button)
        {
            if (button == null) return;

            var flashBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(76, 175, 80));
            var animation = new ColorAnimation
            {
                From = flashBrush.Color,
                To = System.Windows.Media.Colors.White,
                Duration = TimeSpan.FromMilliseconds(6000),
                BeginTime = TimeSpan.FromSeconds(1),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            animation.Completed += (_, _) => button.ClearValue(BackgroundProperty);
            button.Background = flashBrush;
            flashBrush.BeginAnimation(SolidColorBrush.ColorProperty, animation);
        }

        private void FlushUi()
        {
            Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);
            Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Background);
        }

        private void ShowSnackbar(string title, string message, WpfUiControls.ControlAppearance appearance, WpfUiControls.SymbolRegular icon)
        {
            if (SnackbarPresenter == null) return;

            var snackbar = new WpfUiControls.Snackbar(SnackbarPresenter)
            {
                Title = title,
                Content = message,
                Appearance = appearance,
                Icon = new WpfUiControls.SymbolIcon(WpfUiControls.SymbolRegular.PresenceAvailable24)
                {
                    Filled = true,
                    FontSize = 25
                },
                Foreground = System.Windows.Media.Brushes.Black,
                ContentForeground = System.Windows.Media.Brushes.Black,
                Timeout = TimeSpan.FromSeconds(4),
                IsCloseButtonEnabled = false
            };

            snackbar.Show();
        }

        private void ProfileNameBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Profile name text is used when running a scrape; it no longer filters the Profile tab.
        }

        private void SetSearchNameBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _lastSetName = SetSearchNameBox?.Text?.Trim();
            if (string.IsNullOrWhiteSpace(_lastSetName))
            {
                return;
            }

            var match = _availableSets?
                .FirstOrDefault(s => string.Equals(s.Name, _lastSetName, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                _lastSetIsSearch = match.IsSearch;
            }

            UpdateMenuItems(SetSearchMenu, _availableSets, _lastSetName, _lastSetIsSearch);
        }

        private void ExistingProfileCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ExistingProfileCombo.SelectedItem is string profile)
            {
                ProfileNameBox.Text = profile;
                SelectLatestRunForProfile(profile);
            }
        }

        private void UseLatestProfile_Click(object sender, RoutedEventArgs e)
        {
            var profile = (ExistingProfileCombo.SelectedItem as string) ?? ProfileNameBox.Text?.Trim();
            if (!string.IsNullOrWhiteSpace(profile))
            {
                SelectLatestRunForProfile(profile);
            }
        }

        private void PickSingleItem_Click(object sender, RoutedEventArgs e)
        {
            SingleItemRadio.IsChecked = true;
            try
            {
                var doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                var item = doc?.CurrentSelection?.SelectedItems?.FirstOrDefault();
                SingleItemLabel.Text = item != null ? item.DisplayName : "No item selected";
            }
            catch
            {
                SingleItemLabel.Text = "No item selected";
            }
        }

        private void SetSearchRadio_Checked(object sender, RoutedEventArgs e)
        {
            LoadSets();
        }

        private void OtherScope_Checked(object sender, RoutedEventArgs e)
        {
            // No-op; set/search controls are enabled via binding.
        }
    }

    internal class ScrapedPropertyView
    {
        public ScrapedPropertyView(ScrapedProperty p)
        {
            Category = p.Category;
            Name = p.Name;
            DataType = p.DataType;
            ItemCount = p.ItemCount;
            DistinctValueCount = p.DistinctValueCount;
            SampleValuesString = string.Join("; ", p.SampleValues ?? new List<string>());
        }
        public string Category { get; }
        public string Name { get; }
        public string DataType { get; }
        public int ItemCount { get; }
        public int DistinctValueCount { get; }
        public string SampleValuesString { get; }
    }

    internal class RawRow
    {
        public string Profile { get; set; }
        public string Scope { get; set; }
        public string ItemPath { get; set; }
        public string Category { get; set; }
        public string Name { get; set; }
        public string DataType { get; set; }
        public string Value { get; set; }
    }
}

