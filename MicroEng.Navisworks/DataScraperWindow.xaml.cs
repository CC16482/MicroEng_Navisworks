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
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using Microsoft.Win32;
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
        private bool _rawEntriesTruncated;
        private bool _rawEntriesStored;
        private ScrapeSession _activeSession;
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
                WireExportOutputValidation();
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
                if (!ValidateExportOutput())
                {
                    return;
                }

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

                MicroEngActions.Log($"[DataScraper] Run start. Profile='{profile}', Scope={scope}, Items={items.Count}, ExportEnabled={ExportEnabledCheckBox?.IsChecked == true}");

                var exportSettings = BuildExportSettingsFromUi();

                StatusText.Text = "Scraping...";
                ScrapeSession session = null;
                session = _service.Scrape(profile, scope, description, items, exportSettings);

                var status = $"Scraped {session.ItemsScanned} items, {session.Properties.Count} properties.";
                if (session.JsonlExportEnabled)
                {
                    var exportInfo = $" Exported {session.JsonlLinesWritten} lines";
                    if (!string.IsNullOrWhiteSpace(session.JsonlExportPath))
                    {
                        exportInfo += $" to {session.JsonlExportPath}";
                    }
                    exportInfo += ".";
                    status += exportInfo;
                }

                StatusText.Text = status;
                var logMessage = $"Data Scraper: {session.ItemsScanned} items scanned in {description}. {session.Properties.Count} properties cached for profile '{profile}'.";
                if (session.JsonlExportEnabled)
                {
                    logMessage += $" Exported {session.JsonlLinesWritten} lines to {session.JsonlExportPath}.";
                }
                MicroEngActions.Log(logMessage);
                RefreshHistory();
                LoadProfiles();
                SelectLatestRunForProfile(profile);
                ShowProperties(session);
                MicroEngActions.Log($"[DataScraper] Run complete. Items={session.ItemsScanned}, Properties={session.Properties.Count}, ExportLines={session.JsonlLinesWritten}");
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
            _activeSession = session;
            _currentProperties = session.Properties
                .Select(p => new ScrapedPropertyView(p))
                .ToList();
            _propertiesView = CollectionViewSource.GetDefaultView(_currentProperties);
            PropertiesGrid.ItemsSource = _propertiesView;
            PropertySummary.Text = $"{_currentProperties.Count} properties from session {session.Timestamp:T}";

            _rawRows = session.RawEntries
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

            _rawEntriesStored = session.RawEntries.Count > 0;
            _rawEntriesTruncated = session.RawEntriesTruncated;

            RawDataNotStoredHint.Visibility = Visibility.Collapsed;
            if (_rawEntriesTruncated)
            {
                RawDataNotStoredHint.Text = _rawEntriesStored
                    ? "Raw rows were truncated to the preview limit."
                    : "Raw rows weren’t stored (memory saver). Export still contains full data.";
                RawDataNotStoredHint.Visibility = Visibility.Visible;
            }

            var rawSummary = $"{_rawRows.Count} raw entries";
            if (_rawEntriesTruncated)
            {
                rawSummary += _rawEntriesStored ? " (preview)" : " (not stored)";
            }
            RawSummaryText.Text = rawSummary + ".";

            SelectedRunSummary.Text = $"{session.ProfileName} · {session.ScopeType} · {session.ItemsScanned} items · {session.Timestamp:G}";
            UpdateExportSummary(session);
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

        private void UpdateExportSummary(ScrapeSession session)
        {
            if (session == null)
            {
                ExportSummaryText.Text = string.Empty;
                return;
            }

            if (!session.JsonlExportEnabled)
            {
                ExportSummaryText.Text = "No export recorded for this run.";
                return;
            }

            var message = $"Last export: {session.JsonlLinesWritten} line(s)";
            if (!string.IsNullOrWhiteSpace(session.JsonlExportPath))
            {
                message += $" · {session.JsonlExportPath}";
            }

            ExportSummaryText.Text = message;
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
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                var root = doc?.SelectionSets?.RootItem;
                if (root != null)
                {
                    CollectSelectionSets(root, lists.Sets, seen);
                }
            }
            catch
            {
                // ignore managed API enumeration failures
            }

            if (lists.Sets.Count == 0)
            {
                try
                {
                    var state = ComBridge.State;
                    var root = state.SelectionSetsEx();
                    CollectSelectionSets(root, lists.Sets, seen, null);
                }
                catch
                {
                    // ignore COM enumeration failures
                }
            }

            lists.Sets.Sort((a, b) => string.Compare(a.DisplayName, b.DisplayName, StringComparison.OrdinalIgnoreCase));
            return lists;
        }

        private static void CollectSelectionSets(FolderItem folder, List<SetListItem> items, HashSet<string> seen)
        {
            if (folder?.Children == null)
            {
                return;
            }

            foreach (SavedItem child in folder.Children)
            {
                if (child is SelectionSet set)
                {
                    var name = set.DisplayName;
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    AddSetItem(items, seen, name, set.HasSearch);
                }
                else if (child is FolderItem subFolder)
                {
                    CollectSelectionSets(subFolder, items, seen);
                }
            }
        }

        private static void CollectSelectionSets(ComApi.InwSelectionSetExColl collection,
            List<SetListItem> items,
            HashSet<string> seen,
            string folderHint)
        {
            if (collection == null) return;

            for (int i = 1; i <= collection.Count; i++)
            {
                object item;
                try
                {
                    item = collection[i];
                }
                catch
                {
                    continue;
                }

                if (item is ComApi.InwOpSelectionSet set)
                {
                    var name = TryGetName(set);
                    if (string.IsNullOrWhiteSpace(name)) continue;

                    var isSearch = TryGetSearchFlag(set)
                        ?? string.Equals(folderHint, "search", StringComparison.OrdinalIgnoreCase);
                    AddSetItem(items, seen, name, isSearch);
                }
                else if (item is ComApi.InwSelectionSetFolder folder)
                {
                    var folderName = TryGetName(folder);
                    var nextHint = folderHint;
                    if (!string.IsNullOrWhiteSpace(folderName))
                    {
                        var lowered = folderName.ToLowerInvariant();
                        if (lowered.Contains("search"))
                        {
                            nextHint = "search";
                        }
                        else if (lowered.Contains("selection"))
                        {
                            nextHint = "selection";
                        }
                    }

                    try
                    {
                        CollectSelectionSets(folder.SelectionSets(), items, seen, nextHint);
                    }
                    catch
                    {
                        // ignore folder enumeration failures
                    }
                }
            }
        }

        private static void AddSetItem(List<SetListItem> items, HashSet<string> seen, string name, bool isSearch)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            var key = $"{name}|{(isSearch ? "S" : "N")}";
            if (!seen.Add(key))
            {
                return;
            }

            items.Add(new SetListItem(name, isSearch));
        }

        private static string TryGetName(object obj)
        {
            if (obj == null) return null;

            try
            {
                var prop = obj.GetType().GetProperty("name");
                if (prop != null && prop.PropertyType == typeof(string))
                {
                    return prop.GetValue(obj) as string;
                }
            }
            catch
            {
            }

            try
            {
                var prop = obj.GetType().GetProperty("Name");
                if (prop != null && prop.PropertyType == typeof(string))
                {
                    return prop.GetValue(obj) as string;
                }
            }
            catch
            {
            }

            return null;
        }

        private static bool? TryGetSearchFlag(object obj)
        {
            return TryGetBoolProperty(obj, "IsSearch", "IsSearchSet", "IsSearchSelection");
        }

        private static bool? TryGetBoolProperty(object obj, params string[] names)
        {
            if (obj == null || names == null) return null;

            foreach (var name in names)
            {
                try
                {
                    var prop = obj.GetType().GetProperty(name);
                    if (prop != null && prop.PropertyType == typeof(bool))
                    {
                        return (bool)prop.GetValue(obj);
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private void BrowseExport_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "JSONL (*.jsonl)|*.jsonl|GZip JSONL (*.jsonl.gz)|*.jsonl.gz|All files (*.*)|*.*",
                DefaultExt = ExportGzipCheckBox?.IsChecked == true ? ".jsonl.gz" : ".jsonl",
                FileName = string.IsNullOrWhiteSpace(ProfileNameBox.Text) ? "DataScraper.jsonl" : $"{ProfileNameBox.Text.Trim()}.jsonl"
            };

            if (dialog.ShowDialog(this) == true)
            {
                ExportPathBox.Text = dialog.FileName;
            }
        }

        private void ExportNow_Click(object sender, RoutedEventArgs e)
        {
            var scrollOffset = MainScrollViewer?.VerticalOffset ?? 0;
            SetExportNowLoading(true);
            FlushUi();
            try
            {
                if (_activeSession == null)
                {
                    MessageBox.Show("Select a run in History before exporting.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (ExportEnabledCheckBox?.IsChecked != true)
                {
                    MessageBox.Show("Enable JSONL export to use Export Now.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (!ValidateExportOutput())
                {
                    return;
                }

                var exportSettings = BuildExportSettingsFromUi();
                if (!exportSettings.ExportRawRows && !exportSettings.ExportItemDocuments)
                {
                    if (exportSettings.ExportSource != DataScraperExportSource.PropertySummaries)
                    {
                        MessageBox.Show("Select Raw Rows and/or Item Documents to export.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }
                }

                if (exportSettings.ExportSource == DataScraperExportSource.FullRaw)
                {
                    StatusText.Text = "Rescraping for full export...";
                    exportSettings.StreamDuringScrape = true;
                    exportSettings.KeepRawEntriesInMemory = false;
                    exportSettings.PreviewRawRowLimit = 0;

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
                        StatusText.Text = "Ready";
                        return;
                    }

                    MicroEngActions.Log($"[DataScraper] ExportNow full rescrape start. Profile='{profile}', Scope={scope}, Items={items.Count}");
                    ScrapeSession session = null;
                    session = _service.Scrape(profile, scope, description, items, exportSettings);

                    StatusText.Text = $"Exported {session.JsonlLinesWritten} line(s).";
                    RefreshHistory();
                    LoadProfiles();
                    SelectLatestRunForProfile(profile);
                    ShowProperties(session);
                    UpdateExportSummary(session);
                    MicroEngActions.Log($"Data Scraper: re-scraped and exported {session.JsonlLinesWritten} JSONL line(s) for '{profile}'.");
                    FlashSuccess(ExportNowButton);
                    ShowSnackbar("Export complete", $"Exported {session.JsonlLinesWritten} line(s).", WpfUiControls.ControlAppearance.Success, WpfUiControls.SymbolRegular.CheckmarkCircle24);
                    return;
                }

                if (exportSettings.ExportSource != DataScraperExportSource.PropertySummaries
                    && (_activeSession.RawEntries == null || _activeSession.RawEntries.Count == 0))
                {
                    MessageBox.Show("No raw rows are stored in memory for this run. Enable 'Keep raw rows in memory' and re-run the scrape.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (exportSettings.ExportSource != DataScraperExportSource.PropertySummaries && _activeSession.RawEntriesTruncated)
                {
                    var result = MessageBox.Show(
                        "Raw rows are truncated for this run. Export Now will only write the preview rows. Continue?",
                        "MicroEng",
                        MessageBoxButton.OKCancel,
                        MessageBoxImage.Warning);
                    if (result != MessageBoxResult.OK)
                    {
                        return;
                    }
                }

                StatusText.Text = "Exporting...";
                MicroEngActions.Log($"[DataScraper] ExportNow from cache. Profile='{_activeSession.ProfileName}', Source={exportSettings.ExportSource}, Raw={exportSettings.ExportRawRows}, Items={exportSettings.ExportItemDocuments}");
                var lines = _service.ExportFromSession(_activeSession, exportSettings);
                StatusText.Text = $"Exported {lines} line(s).";
                UpdateExportSummary(_activeSession);
                MicroEngActions.Log($"Data Scraper: exported {lines} JSONL line(s) from cached run '{_activeSession.ProfileName}'.");
                FlashSuccess(ExportNowButton);
                ShowSnackbar("Export complete", $"Exported {lines} line(s).", WpfUiControls.ControlAppearance.Success, WpfUiControls.SymbolRegular.CheckmarkCircle24);
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"[DataScraper] ExportNow failed: {ex}");
                StatusText.Text = ex.Message;
                MessageBox.Show(ex.Message, "MicroEng", MessageBoxButton.OK, MessageBoxImage.Error);
                ShowSnackbar("Export failed", ex.Message, WpfUiControls.ControlAppearance.Danger, WpfUiControls.SymbolRegular.ErrorCircle24);
            }
            finally
            {
                SetExportNowLoading(false);
                RestoreScrollOffset(scrollOffset);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void RestoreScrollOffset(double offset)
        {
            if (MainScrollViewer == null)
            {
                return;
            }

            Dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                new Action(() => MainScrollViewer.ScrollToVerticalOffset(offset)));
        }

        private DataScraperJsonlExportSettings BuildExportSettingsFromUi()
        {
            var exportEnabled = ExportEnabledCheckBox?.IsChecked == true;
            var exportFormat = GetSelectedTag(ExportFormatComboBox, "RawOnly");
            var exportRaw = string.Equals(exportFormat, "RawOnly", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(exportFormat, "Both", StringComparison.OrdinalIgnoreCase);
            var exportItems = string.Equals(exportFormat, "ItemsOnly", StringComparison.OrdinalIgnoreCase)
                              || string.Equals(exportFormat, "Both", StringComparison.OrdinalIgnoreCase);
            if (!exportRaw && !exportItems)
            {
                exportEnabled = (GetSelectedEnum(ExportSourceComboBox, DataScraperExportSource.PreviewRaw) == DataScraperExportSource.PropertySummaries)
                    && ExportEnabledCheckBox?.IsChecked == true;
            }
            return new DataScraperJsonlExportSettings
            {
                Enabled = exportEnabled,
                StreamDuringScrape = exportEnabled && ExportStreamDuringScrapeCheckBox?.IsChecked == true,
                Mode = exportItems && !exportRaw ? DataScraperJsonlExportMode.ItemDocuments : DataScraperJsonlExportMode.RawRows,
                ExportSource = GetSelectedEnum(ExportSourceComboBox, DataScraperExportSource.PreviewRaw),
                ExportRawRows = exportRaw,
                ExportItemDocuments = exportItems,
                Gzip = ExportGzipCheckBox?.IsChecked == true,
                OutputPath = ExportPathBox?.Text?.Trim() ?? string.Empty,
                PrimaryKeyMode = GetSelectedEnum(PrimaryKeyComboBox, DataScraperPrimaryKeyMode.SourceFilePlusBestExternalId),
                SourceFileKeyMode = GetSelectedEnum(SourceFileKeyModeComboBox, DataScraperPathKeyMode.FileNameOnly),
                IncludeDocumentFile = IncludeDocumentFileCheckBox?.IsChecked == true,
                IncludeSourceFile = IncludeSourceFileCheckBox?.IsChecked == true,
                IncludeNavisworksInstanceGuid = IncludeNavisInstanceGuidCheckBox?.IsChecked == true,
                IncludeRevitIds = IncludeRevitIdsCheckBox?.IsChecked == true,
                IncludeIfcIds = IncludeIfcIdsCheckBox?.IsChecked == true,
                IncludeDwgIds = IncludeDwgIdsCheckBox?.IsChecked == true,
                NormalizeIds = NormalizeIdsCheckBox?.IsChecked == true,
                KeepRawEntriesInMemory = KeepRawEntriesCheckBox?.IsChecked == true,
                PreviewRawRowLimit = ParsePreviewLimit()
            };
        }

        private int ParsePreviewLimit()
        {
            if (int.TryParse(PreviewRawRowLimitBox?.Text, out var limit) && limit > 0)
            {
                return limit;
            }

            return 50000;
        }

        private static TEnum GetSelectedEnum<TEnum>(ComboBox combo, TEnum fallback) where TEnum : struct
        {
            var tag = (combo?.SelectedItem as ComboBoxItem)?.Tag as string
                      ?? (combo?.SelectedItem as ComboBoxItem)?.Content as string;

            if (!string.IsNullOrWhiteSpace(tag) && Enum.TryParse(tag, out TEnum parsed))
            {
                return parsed;
            }

            return fallback;
        }

        private static string GetSelectedTag(ComboBox combo, string fallback)
        {
            var tag = (combo?.SelectedItem as ComboBoxItem)?.Tag as string
                      ?? (combo?.SelectedItem as ComboBoxItem)?.Content as string;
            if (!string.IsNullOrWhiteSpace(tag))
            {
                return tag.Trim();
            }

            return fallback;
        }

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

        private void SetExportNowLoading(bool isLoading)
        {
            if (ExportNowButton == null) return;
            ExportNowButton.IsEnabled = !isLoading;
            if (ExportNowLabel != null) ExportNowLabel.Visibility = isLoading ? Visibility.Collapsed : Visibility.Visible;
            if (ExportNowSpinnerIcon != null) ExportNowSpinnerIcon.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
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
            if (!isLoading)
            {
                UpdateExportNowEnabled();
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

        private void WireExportOutputValidation()
        {
            if (ExportPathBox != null)
            {
                ExportPathBox.TextChanged += (_, __) => UpdateExportNowEnabled();
            }

            if (ExportEnabledCheckBox != null)
            {
                ExportEnabledCheckBox.Checked += (_, __) => UpdateExportNowEnabled();
                ExportEnabledCheckBox.Unchecked += (_, __) => UpdateExportNowEnabled();
            }

            UpdateExportNowEnabled();
        }

        private void UpdateExportNowEnabled()
        {
            if (ExportNowButton == null)
            {
                return;
            }

            var enabled = ExportEnabledCheckBox?.IsChecked == true;
            var hasPath = !string.IsNullOrWhiteSpace(ExportPathBox?.Text);
            ExportNowButton.IsEnabled = enabled && hasPath;
        }

        private bool ValidateExportOutput()
        {
            if (ExportEnabledCheckBox?.IsChecked != true)
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(ExportPathBox?.Text))
            {
                return true;
            }

            const string message = "Output path is required for JSONL export. Specify Output or uncheck Enable JSONL export.";
            StatusText.Text = message;
            MessageBox.Show(message, "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
            ShowSnackbar("Output required",
                "Specify Output or disable export.",
                WpfUiControls.ControlAppearance.Danger,
                WpfUiControls.SymbolRegular.ErrorCircle24);
            return false;
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

        private void NewProfile_Click(object sender, RoutedEventArgs e)
        {
            ExistingProfileCombo.SelectedItem = null;
            ProfileNameBox.Text = string.Empty;
            ProfileNameBox.Focus();
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
