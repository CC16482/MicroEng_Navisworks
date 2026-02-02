using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows.Input;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Autodesk.Navisworks.Api;
using Microsoft.Win32;
using NavisApp = Autodesk.Navisworks.Api.Application;
using WpfUiControls = Wpf.Ui.Controls;

namespace MicroEng.Navisworks
{
    public partial class DataMatrixControl : UserControl
    {
        static DataMatrixControl()
        {
            AssemblyResolver.EnsureRegistered();
        }

        private readonly DataMatrixRowBuilder _builder = new();
        private readonly IDataMatrixPresetManager _presetManager = new FilePresetManager();
        private readonly DataMatrixExporter _exporter = new();

        private List<DataMatrixAttributeDefinition> _attributes = new();
        private List<DataMatrixAttributeDefinition> _attributeCatalogAll = new();
        private List<DataMatrixRow> _allRows = new();
        private ObservableCollection<DataMatrixRow> _viewRows = new();
        private readonly Dictionary<Guid, ModelItem> _itemCache = new();

        private bool _syncEnabled;
        private bool _suppressSelectionSync;
        private DataMatrixViewPreset _currentPreset;
        private ScrapeSession _currentSession;
        private List<string> _visibleColumnOverride;
        private bool _updatingScope;

        public DataMatrixControl()
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            Loaded += (_, __) => InitializeUi();
        }

        private void InitializeUi()
        {
            RefreshProfiles();
            RefreshProfilesButton.Click += (s, e) => RefreshProfiles();
            ProfileCombo.SelectionChanged += (s, e) => LoadProfileSession();

            ScopeCombo.SelectedIndex = 0;
            ScopeCombo.SelectionChanged += (s, e) => OnScopeChanged();
            ScopeSetCombo.SelectionChanged += (s, e) => OnScopeSetChanged();
            RefreshScopeSetsButton.Click += (s, e) =>
            {
                RefreshScopeSets();
                RebuildMatrixFromCurrentState();
            };

            SaveViewButton.Click += (s, e) => SavePreset(saveAs: false);
            SaveViewAsButton.Click += (s, e) => SavePreset(saveAs: true);
            DeleteViewButton.Click += (s, e) => DeletePreset();
            ViewPresetCombo.SelectionChanged += (s, e) => ApplyPresetSelection();

            ColumnsButton.Click += (s, e) => ShowColumnsDialog();
            SyncToggle.Checked += (s, e) => _syncEnabled = true;
            SyncToggle.Unchecked += (s, e) => _syncEnabled = false;
            SelectedOnlyCheck.Checked += (s, e) => ApplyFiltersAndSorts();
            SelectedOnlyCheck.Unchecked += (s, e) => ApplyFiltersAndSorts();

            ExportFilteredMenu.Click += (s, e) => ExportCsv(filtered: true);
            ExportAllMenu.Click += (s, e) => ExportCsv(filtered: false);
            ExportJsonlItemFilteredMenu.Click += (s, e) => ExportJsonl(filtered: true, DataMatrixJsonlMode.ItemDocuments);
            ExportJsonlItemAllMenu.Click += (s, e) => ExportJsonl(filtered: false, DataMatrixJsonlMode.ItemDocuments);
            ExportJsonlRawFilteredMenu.Click += (s, e) => ExportJsonl(filtered: true, DataMatrixJsonlMode.RawRows);
            ExportJsonlRawAllMenu.Click += (s, e) => ExportJsonl(filtered: false, DataMatrixJsonlMode.RawRows);

        }

        private void RefreshProfiles()
        {
            ProfileCombo.Items.Clear();
            var profiles = DataScraperCache.AllSessions
                .Select(s => string.IsNullOrWhiteSpace(s.ProfileName) ? "Default" : s.ProfileName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(p => p)
                .ToList();
            if (!profiles.Any()) profiles.Add("Default");
            foreach (var p in profiles) ProfileCombo.Items.Add(p);
            ProfileCombo.SelectedIndex = 0;
            LoadProfileSession();
        }

        private void LoadProfileSession()
        {
            var profile = ProfileCombo.SelectedItem?.ToString() ?? "Default";
            var session = DataScraperCache.AllSessions
                .Where(s => string.Equals(s.ProfileName ?? "Default", profile, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(s => s.Timestamp)
                .FirstOrDefault();
            if (session == null)
            {
                MatrixGrid.ItemsSource = null;
                _attributes.Clear();
                _attributeCatalogAll.Clear();
                _allRows.Clear();
                _currentSession = null;
                _visibleColumnOverride = null;
                RowsStatus.Text = "Rows: 0";
                ProfileStatus.Text = $"Profile: {profile} (no sessions)";
                UpdateScopeUiText();
                UpdateSelectionStatus();
                return;
            }

            DataScraperCache.LastSession = session;
            _currentSession = session;
            _attributeCatalogAll = _builder.BuildAttributeCatalog(session);
            _itemCache.Clear();
            LoadPresets(profile);
            ApplyPresetSelection();
            ProfileStatus.Text = $"Profile: {profile} @ {session.Timestamp:t}";
        }

        private void BuildGridColumns(IList<string> visibleIds)
        {
            MatrixGrid.Columns.Clear();

            if (visibleIds == null)
            {
                var catalog = _attributeCatalogAll.Count > 0 ? _attributeCatalogAll : _attributes;
                visibleIds = catalog
                    .Where(a => a.IsVisibleByDefault)
                    .Select(a => a.Id)
                    .ToList();
            }

            var ids = (visibleIds ?? new List<string>())
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .ToList();

            if (ids.Count > 0)
            {
                MatrixGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = "Element",
                    Binding = new System.Windows.Data.Binding("ElementDisplayName"),
                    Width = TryGetElementColumnWidth()
                });
            }

            foreach (var id in ids)
            {
                var attr = _attributes.FirstOrDefault(a => string.Equals(a.Id, id, StringComparison.OrdinalIgnoreCase));
                if (attr == null) continue;
                MatrixGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = attr.DisplayName ?? attr.PropertyName,
                    Binding = new System.Windows.Data.Binding($"Values[{attr.Id}]"),
                    Width = TryGetPresetWidth(attr.Id) ?? (attr.DefaultWidth > 0 ? attr.DefaultWidth : 140)
                });
            }
        }

        private List<string> GetVisibleAttributeIds()
        {
            if (_visibleColumnOverride != null)
            {
                return _visibleColumnOverride
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            if (_currentPreset?.VisibleAttributeIds != null && _currentPreset.VisibleAttributeIds.Count > 0)
            {
                return _currentPreset.VisibleAttributeIds
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            return new List<string>();
        }

        private void RebuildMatrixFromCurrentState()
        {
            if (_currentSession == null)
            {
                MatrixGrid.ItemsSource = null;
                _attributes.Clear();
                _allRows.Clear();
                RowsStatus.Text = "Rows: 0";
                ViewStatus.Text = $"View: {_currentPreset?.Name ?? "(Default)"}";
                UpdateScopeUiText();
                UpdateSelectionStatus();
                return;
            }

            var visibleIds = GetVisibleAttributeIds();
            if (visibleIds.Count == 0)
            {
                _attributes = new List<DataMatrixAttributeDefinition>();
                _allRows = new List<DataMatrixRow>();
                _itemCache.Clear();

                BuildGridColumns(visibleIds);
                _viewRows = new ObservableCollection<DataMatrixRow>();
                MatrixGrid.ItemsSource = _viewRows;
                RowsStatus.Text = "Rows: 0 (choose columns)";
                ViewStatus.Text = $"View: {_currentPreset?.Name ?? "(Default)"}";
                UpdateScopeUiText();
                UpdateSelectionStatus();
                return;
            }

            var scopeKeys = ResolveScopeItemKeys();
            var joinMulti = _currentPreset?.JoinMultiValues ?? false;
            var separator = string.IsNullOrWhiteSpace(_currentPreset?.MultiValueSeparator) ? "; " : _currentPreset.MultiValueSeparator;

            var built = _builder.Build(_currentSession, visibleIds, scopeKeys, joinMulti, separator);
            _attributes = built.Attributes;
            _allRows = built.Rows;
            _itemCache.Clear();

            BuildGridColumns(visibleIds);
            ApplyFiltersAndSorts();
        }

        private DataMatrixScopeKind GetScopeKindFromUi()
        {
            var tag = (ScopeCombo.SelectedItem as ComboBoxItem)?.Tag as string
                      ?? (ScopeCombo.SelectedItem as ComboBoxItem)?.Content as string;
            if (!string.IsNullOrWhiteSpace(tag) && Enum.TryParse(tag, out DataMatrixScopeKind parsed))
            {
                return parsed;
            }

            return DataMatrixScopeKind.EntireSession;
        }

        private string GetSelectedScopeSetName()
        {
            return ScopeSetCombo.SelectedItem?.ToString() ?? string.Empty;
        }

        private void SetScopeUi(DataMatrixScopeKind kind, string setName)
        {
            _updatingScope = true;
            try
            {
                foreach (var item in ScopeCombo.Items)
                {
                    if (item is ComboBoxItem combo && string.Equals(combo.Tag as string, kind.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        ScopeCombo.SelectedItem = combo;
                        break;
                    }
                }

                RefreshScopeSets(setName);
                UpdateScopeUiText();
            }
            finally
            {
                _updatingScope = false;
            }
        }

        private void OnScopeChanged()
        {
            if (_updatingScope) return;

            RefreshScopeSets();
            UpdateScopeUiText();

            var kind = GetScopeKindFromUi();
            if (kind != DataMatrixScopeKind.SelectionSet && kind != DataMatrixScopeKind.SearchSet)
            {
                RebuildMatrixFromCurrentState();
            }
            else if (ScopeSetCombo.Items.Count == 0)
            {
                RebuildMatrixFromCurrentState();
            }
        }

        private void OnScopeSetChanged()
        {
            if (_updatingScope) return;
            UpdateScopeUiText();
            RebuildMatrixFromCurrentState();
        }

        private void RefreshScopeSets(string desiredSelection = null)
        {
            var kind = GetScopeKindFromUi();
            ScopeSetCombo.Items.Clear();
            ScopeSetCombo.SelectedItem = null;

            var enableSets = kind == DataMatrixScopeKind.SelectionSet || kind == DataMatrixScopeKind.SearchSet;
            ScopeSetCombo.IsEnabled = enableSets;

            // Only show the Set picker when it is relevant.
            ScopeSetLabel.Visibility = enableSets ? Visibility.Visible : Visibility.Collapsed;
            ScopeSetCombo.Visibility = enableSets ? Visibility.Visible : Visibility.Collapsed;
            RefreshScopeSetsButton.Visibility = enableSets ? Visibility.Visible : Visibility.Collapsed;

            if (!enableSets)
            {
                return;
            }

            var doc = NavisApp.ActiveDocument;
            var names = kind == DataMatrixScopeKind.SelectionSet
                ? NavisworksSelectionSetUtils.GetSelectionSetNames(doc)
                : NavisworksSelectionSetUtils.GetSearchSetNames(doc);

            foreach (var name in names)
            {
                ScopeSetCombo.Items.Add(name);
            }

            var target = desiredSelection;
            if (string.IsNullOrWhiteSpace(target))
            {
                target = _currentPreset?.ScopeSetName ?? ScopeSetCombo.SelectedItem?.ToString();
            }

            if (!string.IsNullOrWhiteSpace(target))
            {
                var match = ScopeSetCombo.Items
                    .Cast<object>()
                    .FirstOrDefault(item => string.Equals(item?.ToString(), target, StringComparison.OrdinalIgnoreCase));
                if (match != null)
                {
                    ScopeSetCombo.SelectedItem = match;
                    return;
                }
            }

            if (ScopeSetCombo.Items.Count > 0)
            {
                ScopeSetCombo.SelectedIndex = 0;
            }
        }


        private void UpdateScopeUiText()
        {
            try
            {
                var kind = GetScopeKindFromUi();
                var label = GetScopeLabel(kind);

                if (kind == DataMatrixScopeKind.SelectionSet || kind == DataMatrixScopeKind.SearchSet)
                {
                    var set = GetSelectedScopeSetName();
                    if (!string.IsNullOrWhiteSpace(set))
                    {
                        label = $"{label}: {set}";
                    }
                }

                var text = $"Scope: {label}";
                if (ScopeBadgeText != null) ScopeBadgeText.Text = text;
                if (ScopeStatus != null) ScopeStatus.Text = text;
            }
            catch
            {
                // Never let a UI string update crash Navisworks.
            }
        }

        private static string GetScopeLabel(DataMatrixScopeKind kind)
        {
            switch (kind)
            {
                case DataMatrixScopeKind.EntireSession:
                    return "Entire session";
                case DataMatrixScopeKind.CurrentSelection:
                    return "Current selection";
                case DataMatrixScopeKind.SelectionSet:
                    return "Selection set";
                case DataMatrixScopeKind.SearchSet:
                    return "Search set";
                case DataMatrixScopeKind.SingleItem:
                    return "Single item";
                default:
                    return kind.ToString();
            }
        }

        private void UpdateSelectionStatus()
        {
            if (SelectionStatus == null) return;
            SelectionStatus.Text = $"Selected: {MatrixGrid?.SelectedItems?.Count ?? 0}";
        }

        private double TryGetElementColumnWidth()
        {
            var w = _currentPreset?.ElementColumnWidth ?? 0;
            return w > 20 ? w : 180;
        }

        private double? TryGetPresetWidth(string attributeId)
        {
            if (string.IsNullOrWhiteSpace(attributeId)) return null;

            var dict = _currentPreset?.ColumnWidths;
            if (dict == null || dict.Count == 0) return null;

            if (dict.TryGetValue(attributeId, out var w) && w > 20) return w;

            // Be defensive about casing differences.
            foreach (var kv in dict)
            {
                if (string.Equals(kv.Key, attributeId, StringComparison.OrdinalIgnoreCase) && kv.Value > 20)
                {
                    return kv.Value;
                }
            }

            return null;
        }


        private ISet<string> ResolveScopeItemKeys()
        {
            var kind = GetScopeKindFromUi();
            if (kind == DataMatrixScopeKind.EntireSession)
            {
                return null;
            }

            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            IEnumerable<ModelItem> items = Enumerable.Empty<ModelItem>();
            switch (kind)
            {
                case DataMatrixScopeKind.CurrentSelection:
                    items = doc.CurrentSelection?.SelectedItems?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
                    break;
                case DataMatrixScopeKind.SingleItem:
                    var single = doc.CurrentSelection?.SelectedItems?.FirstOrDefault();
                    items = single != null ? new[] { single } : Enumerable.Empty<ModelItem>();
                    break;
                case DataMatrixScopeKind.SelectionSet:
                    items = NavisworksSelectionSetUtils.GetItemsFromSet(doc, GetSelectedScopeSetName(), expectSearchSet: false, out _);
                    break;
                case DataMatrixScopeKind.SearchSet:
                    items = NavisworksSelectionSetUtils.GetItemsFromSet(doc, GetSelectedScopeSetName(), expectSearchSet: true, out _);
                    break;
            }

            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in items)
            {
                if (item == null) continue;
                try
                {
                    keys.Add(item.InstanceGuid.ToString("D"));
                }
                catch
                {
                    // ignore
                }
            }

            return keys;
        }

        private HashSet<string> GetAvailableAttributeIdsForScope(ISet<string> itemKeys)
        {
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (_currentSession == null)
            {
                return ids;
            }

            if (itemKeys == null)
            {
                foreach (var attr in _attributeCatalogAll)
                {
                    ids.Add(attr.Id);
                }
                return ids;
            }

            foreach (var entry in _currentSession.RawEntries ?? Enumerable.Empty<RawEntry>())
            {
                if (string.IsNullOrWhiteSpace(entry.ItemKey)) continue;
                if (!itemKeys.Contains(entry.ItemKey)) continue;
                var id = $"{entry.Category}|{entry.Name}";
                ids.Add(id);
            }

            return ids;
        }

        private void ApplyFiltersAndSorts()
        {
            IEnumerable<DataMatrixRow> rows = _allRows;

            if (SelectedOnlyCheck.IsChecked == true && MatrixGrid.SelectedItems.Count > 0)
            {
                var selectedKeys = MatrixGrid.SelectedItems.Cast<DataMatrixRow>().Select(r => r.ItemKey).ToHashSet(StringComparer.OrdinalIgnoreCase);
                rows = rows.Where(r => selectedKeys.Contains(r.ItemKey));
            }

            // Basic sorting: rely on DataGrid sorting event to reorder _viewRows
            _viewRows = new ObservableCollection<DataMatrixRow>(rows);
            MatrixGrid.ItemsSource = _viewRows;
            RowsStatus.Text = $"Rows: {_viewRows.Count} (from {_allRows.Count})";
            ViewStatus.Text = $"View: {_currentPreset?.Name ?? "(Default)"}";
            UpdateScopeUiText();
            UpdateSelectionStatus();
        }

        private void MatrixGrid_Sorting(object sender, DataGridSortingEventArgs e)
        {
            e.Handled = true;
            var col = e.Column;
            var direction = col.SortDirection != System.ComponentModel.ListSortDirection.Ascending
                ? System.ComponentModel.ListSortDirection.Ascending
                : System.ComponentModel.ListSortDirection.Descending;
            col.SortDirection = direction;

            var binding = (col as DataGridBoundColumn)?.Binding as System.Windows.Data.Binding;
            var path = binding?.Path?.Path;
            if (string.IsNullOrWhiteSpace(path)) return;

            var sorted = direction == System.ComponentModel.ListSortDirection.Ascending
                ? _viewRows.OrderBy(r => GetValue(r, path)).ToList()
                : _viewRows.OrderByDescending(r => GetValue(r, path)).ToList();
            _viewRows = new ObservableCollection<DataMatrixRow>(sorted);
            MatrixGrid.ItemsSource = _viewRows;
        }

        private object GetValue(DataMatrixRow row, string path)
        {
            if (path == "ElementDisplayName") return row.ElementDisplayName ?? string.Empty;
            if (path.StartsWith("Values["))
            {
                var key = path.Substring("Values[".Length).TrimEnd(']');
                if (row.Values.TryGetValue(key, out var val)) return val ?? string.Empty;
            }
            return string.Empty;
        }

        private void MatrixGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateSelectionStatus();

            if (_suppressSelectionSync || !_syncEnabled) return;
            var items = MatrixGrid.SelectedItems.Cast<DataMatrixRow>().ToList();
            if (!items.Any()) return;

            var doc = NavisApp.ActiveDocument;
            if (doc == null) return;

            var collection = new ModelItemCollection();
            foreach (var row in items)
            {
                var mi = ResolveModelItem(doc, row);
                if (mi != null) collection.Add(mi);
            }

            if (collection.Any())
            {
                doc.CurrentSelection.CopyFrom(collection);
            }
        }


        private void MatrixGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Typical Windows behavior: right-click focuses/selects the row under the mouse.
            var dep = e.OriginalSource as DependencyObject;
            var row = FindAncestor<DataGridRow>(dep);
            if (row == null) return;

            if (!row.IsSelected)
            {
                _suppressSelectionSync = true;
                try
                {
                    MatrixGrid.SelectedItems.Clear();
                    row.IsSelected = true;
                }
                finally
                {
                    _suppressSelectionSync = false;
                }
            }
        }

        private void MatrixGrid_ContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            var cm = sender as ContextMenu;
            if (cm == null) return;

            var selectedRows = MatrixGrid.SelectedItems.Cast<DataMatrixRow>().ToList();
            var hasRows = selectedRows.Count > 0;

            var canCopyCell = MatrixGrid.CurrentColumn != null && (MatrixGrid.CurrentItem as DataMatrixRow ?? MatrixGrid.SelectedItem as DataMatrixRow) != null;
            var canCopyGuid = selectedRows.Any(r => Guid.TryParse(r?.ItemKey, out _));
            var canSelectInModel = hasRows && NavisApp.ActiveDocument != null;

            foreach (var item in cm.Items)
            {
                if (item is not MenuItem mi) continue;
                var tag = mi.Tag as string;

                switch (tag)
                {
                    case "CopyCell":
                        mi.IsEnabled = canCopyCell;
                        break;
                    case "CopyRowsCsv":
                        mi.IsEnabled = hasRows;
                        break;
                    case "CopyGuid":
                        mi.IsEnabled = canCopyGuid;
                        break;
                    case "SelectInModel":
                        mi.IsEnabled = canSelectInModel;
                        break;
                }
            }
        }

        private void CopyCellMenu_Click(object sender, RoutedEventArgs e)
        {
            var row = MatrixGrid.CurrentItem as DataMatrixRow ?? MatrixGrid.SelectedItem as DataMatrixRow;
            var col = MatrixGrid.CurrentColumn;

            if (row == null || col == null) return;

            var text = GetCellText(row, col) ?? string.Empty;
            SafeClipboardSetText(text);

            ShowSnackbar("Copied", "Cell copied to clipboard.", WpfUiControls.ControlAppearance.Success, WpfUiControls.SymbolRegular.CheckmarkCircle24);
        }

        private void CopyRowsCsvMenu_Click(object sender, RoutedEventArgs e)
        {
            var rows = MatrixGrid.SelectedItems.Cast<DataMatrixRow>().ToList();
            if (rows.Count == 0) return;

            var cols = GetGridColumnsInDisplayOrder(includeElement: true);

            var sb = new StringBuilder();
            sb.AppendLine(string.Join(",", cols.Select(c => CsvEscape(c.Header?.ToString() ?? string.Empty))));

            foreach (var row in rows)
            {
                sb.AppendLine(string.Join(",", cols.Select(c => CsvEscape(GetCellText(row, c) ?? string.Empty))));
            }

            SafeClipboardSetText(sb.ToString());
            ShowSnackbar("Copied", $"Copied {rows.Count} row(s) as CSV.", WpfUiControls.ControlAppearance.Success, WpfUiControls.SymbolRegular.CheckmarkCircle24);
        }

        private void CopyGuidMenu_Click(object sender, RoutedEventArgs e)
        {
            var row = MatrixGrid.SelectedItems.Cast<DataMatrixRow>().FirstOrDefault();
            if (row == null) return;

            if (!Guid.TryParse(row.ItemKey, out var guid))
            {
                ShowSnackbar("Copy failed", "Selected row does not have a valid GUID key.", WpfUiControls.ControlAppearance.Caution, WpfUiControls.SymbolRegular.ErrorCircle24);
                return;
            }

            SafeClipboardSetText(guid.ToString());
            ShowSnackbar("Copied", "Item GUID copied to clipboard.", WpfUiControls.ControlAppearance.Success, WpfUiControls.SymbolRegular.CheckmarkCircle24);
        }

        private void SelectInModelMenu_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null) return;

            var rows = MatrixGrid.SelectedItems.Cast<DataMatrixRow>().ToList();
            if (rows.Count == 0) return;

            var collection = new ModelItemCollection();
            foreach (var r in rows)
            {
                var mi = ResolveModelItem(doc, r);
                if (mi != null) collection.Add(mi);
            }

            if (!collection.Any())
            {
                ShowSnackbar("Selection failed", "Could not resolve selected rows to model items.", WpfUiControls.ControlAppearance.Caution, WpfUiControls.SymbolRegular.ErrorCircle24);
                return;
            }

            doc.CurrentSelection.CopyFrom(collection);
        }

        private static void SafeClipboardSetText(string text)
        {
            try
            {
                Clipboard.SetText(text ?? string.Empty);
            }
            catch
            {
                // Clipboard can fail in RDP/locked desktop scenarios. Ignore.
            }
        }

        private static string CsvEscape(string s)
        {
            s ??= string.Empty;
            var needsQuotes = s.Contains(",") || s.Contains("\n") || s.Contains("\r") || s.Contains("\"");

            if (!needsQuotes) return s;

            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }

        private static string GetCellText(DataMatrixRow row, DataGridColumn col)
        {
            if (row == null || col == null) return string.Empty;

            var path = GetColumnBindingPath(col);
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;

            if (string.Equals(path, "ElementDisplayName", StringComparison.OrdinalIgnoreCase))
            {
                return row.ElementDisplayName ?? string.Empty;
            }

            if (TryExtractAttributeId(path, out var id))
            {
                if (row.Values == null) return string.Empty;

                if (row.Values.TryGetValue(id, out var v)) return v ?? string.Empty;

                foreach (var kv in row.Values)
                {
                    if (string.Equals(kv.Key, id, StringComparison.OrdinalIgnoreCase))
                    {
                        return kv.Value ?? string.Empty;
                    }
                }
            }

            return string.Empty;
        }

        private static bool TryExtractAttributeId(string bindingPath, out string attributeId)
        {
            attributeId = null;
            if (string.IsNullOrWhiteSpace(bindingPath)) return false;

            // Values[Some.Id]
            if (bindingPath.StartsWith("Values[", StringComparison.OrdinalIgnoreCase) && bindingPath.EndsWith("]"))
            {
                attributeId = bindingPath.Substring("Values[".Length, bindingPath.Length - "Values[".Length - 1);
                return !string.IsNullOrWhiteSpace(attributeId);
            }

            return false;
        }

        private static string GetColumnBindingPath(DataGridColumn col)
        {
            if (col is not DataGridBoundColumn bc) return null;
            if (bc.Binding is not System.Windows.Data.Binding b) return null;
            return b.Path?.Path;
        }

        private List<DataGridColumn> GetGridColumnsInDisplayOrder(bool includeElement)
        {
            return MatrixGrid.Columns
                .Where(c =>
                {
                    if (c.Visibility != Visibility.Visible) return false;
                    var path = GetColumnBindingPath(c);
                    if (string.IsNullOrWhiteSpace(path)) return false;
                    if (!includeElement && string.Equals(path, "ElementDisplayName", StringComparison.OrdinalIgnoreCase)) return false;
                    return true;
                })
                .OrderBy(c => c.DisplayIndex)
                .ToList();
        }

        private static double GetColumnWidthForPersist(DataGridColumn col)
        {
            if (col == null) return 0;

            var w = col.ActualWidth;
            if (w > 0) return w;

            if (col.Width.IsAbsolute) return col.Width.Value;
            return 0;
        }

        private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T t) return t;
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }


        private ModelItem ResolveModelItem(Document doc, DataMatrixRow row)
        {
            if (row?.ModelItem != null) return row.ModelItem;
            if (Guid.TryParse(row?.ItemKey, out var guid))
            {
                if (_itemCache.TryGetValue(guid, out var cached)) return cached;
                var found = FindByGuid(doc, guid);
                if (found != null)
                {
                    _itemCache[guid] = found;
                    return found;
                }
            }
            return null;
        }

        private static ModelItem FindByGuid(Document doc, Guid guid)
        {
            foreach (var item in Traverse(doc.Models.RootItems))
            {
                try
                {
                    if (item.InstanceGuid == guid) return item;
                }
                catch { }
            }
            return null;
        }

        private static IEnumerable<ModelItem> Traverse(IEnumerable<ModelItem> items)
        {
            foreach (ModelItem item in items)
            {
                yield return item;
                if (item.Children != null && item.Children.Any())
                {
                    foreach (var child in Traverse(item.Children))
                        yield return child;
                }
            }
        }

        private void LoadPresets(string profile)
        {
            ViewPresetCombo.Items.Clear();
            ViewPresetCombo.Items.Add("(Default)");
            var presets = _presetManager.GetPresets(profile).ToList();
            foreach (var p in presets) ViewPresetCombo.Items.Add(new ComboBoxItem { Content = p.Name, Tag = p });
            ViewPresetCombo.SelectedIndex = 0;
            _currentPreset = null;
            _visibleColumnOverride = null;
        }

        private void ApplyPresetSelection()
        {
            if (ViewPresetCombo.SelectedItem is ComboBoxItem item && item.Tag is DataMatrixViewPreset preset)
            {
                _currentPreset = preset;
            }
            else
            {
                _currentPreset = null;
            }

            _visibleColumnOverride = null;

            if (_currentPreset != null)
            {
                SetScopeUi(_currentPreset.ScopeKind, _currentPreset.ScopeSetName);
            }
            else
            {
                SetScopeUi(DataMatrixScopeKind.EntireSession, null);
            }

            RebuildMatrixFromCurrentState();
        }

        private void SavePreset(bool saveAs)
        {
            var profile = ProfileCombo.SelectedItem?.ToString() ?? "Default";
            var name = saveAs || _currentPreset == null
                ? PromptText("View name:", "Save View")
                : _currentPreset.Name;
            if (string.IsNullOrWhiteSpace(name)) return;

            var preset = _currentPreset ?? new DataMatrixViewPreset { Id = Guid.NewGuid().ToString(), ScraperProfileName = profile };
            preset.Name = name;

            // Persist visible columns in the same order the user sees them.
            var cols = GetGridColumnsInDisplayOrder(includeElement: true);

            preset.VisibleAttributeIds = cols
                .Select(GetColumnBindingPath)
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Where(path => !string.Equals(path, "ElementDisplayName", StringComparison.OrdinalIgnoreCase))
                .Select(path => TryExtractAttributeId(path, out var id) ? id : null)
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .ToList();

            // Persist column widths (nice QoL when building schedules).
            preset.ElementColumnWidth = GetColumnWidthForPersist(
                cols.FirstOrDefault(c => string.Equals(GetColumnBindingPath(c), "ElementDisplayName", StringComparison.OrdinalIgnoreCase)));

            if (preset.ElementColumnWidth <= 20)
            {
                preset.ElementColumnWidth = 180;
            }

            preset.ColumnWidths = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in cols)
            {
                var path = GetColumnBindingPath(c);
                if (TryExtractAttributeId(path, out var id))
                {
                    var w = GetColumnWidthForPersist(c);
                    if (w > 20)
                    {
                        preset.ColumnWidths[id] = w;
                    }
                }
            }

            var scopeKind = GetScopeKindFromUi();
            preset.ScopeKind = scopeKind;
            preset.ScopeSetName = scopeKind == DataMatrixScopeKind.SelectionSet || scopeKind == DataMatrixScopeKind.SearchSet
                ? GetSelectedScopeSetName()
                : string.Empty;

            preset.JoinMultiValues = _currentPreset?.JoinMultiValues ?? preset.JoinMultiValues;
            preset.MultiValueSeparator = _currentPreset?.MultiValueSeparator ?? preset.MultiValueSeparator;

            _presetManager.SavePreset(preset);
            LoadPresets(profile);
            _currentPreset = preset;
            SelectPresetInCombo(preset);
        }

        private void SelectPresetInCombo(DataMatrixViewPreset preset)
        {
            foreach (var obj in ViewPresetCombo.Items)
            {
                if (obj is ComboBoxItem item && item.Tag is DataMatrixViewPreset p && p.Id == preset.Id)
                {
                    ViewPresetCombo.SelectedItem = item;
                    break;
                }
            }
        }

        private void DeletePreset()
        {
            if (_currentPreset == null) return;
            _presetManager.DeletePreset(_currentPreset.Id);
            LoadPresets(ProfileCombo.SelectedItem?.ToString() ?? "Default");
        }

        private void ShowColumnsDialog()
        {
            if (_currentSession == null)
            {
                return;
            }

            var currentOrder = GetVisibleColumnIdsFromGrid();
            var scopeKeys = ResolveScopeItemKeys();
            var availableIds = GetAvailableAttributeIdsForScope(scopeKeys);

            var attrById = _attributeCatalogAll
                .ToDictionary(a => a.Id, StringComparer.OrdinalIgnoreCase);

            var ordered = new List<DataMatrixAttributeDefinition>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var id in currentOrder)
            {
                if (!availableIds.Contains(id)) continue;
                if (attrById.TryGetValue(id, out var attr))
                {
                    ordered.Add(attr);
                    seen.Add(id);
                }
            }

            var remaining = _attributeCatalogAll
                .Where(a => availableIds.Contains(a.Id) && !seen.Contains(a.Id))
                .OrderBy(a => a.Category)
                .ThenBy(a => a.PropertyName)
                .ToList();

            ordered.AddRange(remaining);

            var dialog = new DataMatrixColumnBuilderWindow(ordered, currentOrder)
            {
                Owner = Window.GetWindow(this)
            };
            if (dialog.ShowDialog() == true)
            {
                var visibleIds = dialog.VisibleAttributeIds.ToList();
                _visibleColumnOverride = visibleIds;
                RebuildMatrixFromCurrentState();
            }
        }

        private List<string> GetVisibleColumnIdsFromGrid()
        {
            return GetGridColumnsInDisplayOrder(includeElement: false)
                .Select(GetColumnBindingPath)
                .Select(path => TryExtractAttributeId(path, out var id) ? id : null)
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .ToList();
        }

        private void ExportCsv(bool filtered)
        {
            if (!_allRows.Any()) return;
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV Files|*.csv",
                FileName = "DataMatrix.csv"
            };
            if (dlg.ShowDialog() != true) return;

            var rows = filtered ? _viewRows.ToList() : _allRows;

            // Respect the user's current column order.
            var columns = GetVisibleColumnIdsFromGrid()
                .Select(id => _attributes.FirstOrDefault(a => string.Equals(a.Id, id, StringComparison.OrdinalIgnoreCase)))
                .Where(a => a != null)
                .ToList();

            var session = _currentSession ?? DataScraperCache.LastSession ?? DataScraperCache.AllSessions.FirstOrDefault();
            if (session == null)
            {
                return;
            }
            try
            {
                _exporter.ExportCsv(dlg.FileName, columns, rows, session, _currentPreset);
                ShowSnackbar("Export complete",
                    $"Wrote {rows.Count} row(s).",
                    WpfUiControls.ControlAppearance.Success,
                    WpfUiControls.SymbolRegular.CheckmarkCircle24);
            }
            catch (Exception ex)
            {
                ShowSnackbar("Export failed",
                    ex.Message,
                    WpfUiControls.ControlAppearance.Danger,
                    WpfUiControls.SymbolRegular.ErrorCircle24);
            }
        }

        private void ExportJsonl(bool filtered, DataMatrixJsonlMode mode)
        {
            if (!_allRows.Any()) return;
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "JSONL Files|*.jsonl|GZip JSONL|*.jsonl.gz|All files (*.*)|*.*",
                FileName = "DataMatrix.jsonl",
                DefaultExt = ".jsonl"
            };
            if (dlg.ShowDialog() != true) return;

            var rows = filtered ? _viewRows.ToList() : _allRows;

            // Respect the user's current column order.
            var columns = GetVisibleColumnIdsFromGrid()
                .Select(id => _attributes.FirstOrDefault(a => string.Equals(a.Id, id, StringComparison.OrdinalIgnoreCase)))
                .Where(attr => attr != null)
                .ToList();

            var session = _currentSession ?? DataScraperCache.LastSession ?? DataScraperCache.AllSessions.FirstOrDefault();
            if (session == null)
            {
                return;
            }

            try
            {
                _exporter.ExportJsonl(dlg.FileName, columns, rows, session, _currentPreset, mode);
                ShowSnackbar("Export complete",
                    $"Wrote {rows.Count} row(s).",
                    WpfUiControls.ControlAppearance.Success,
                    WpfUiControls.SymbolRegular.CheckmarkCircle24);
            }
            catch (Exception ex)
            {
                ShowSnackbar("Export failed",
                    ex.Message,
                    WpfUiControls.ControlAppearance.Danger,
                    WpfUiControls.SymbolRegular.ErrorCircle24);
            }
        }

        private static string PromptText(string caption, string title)
        {
            var window = new Window
            {
                Title = title,
                Width = 360,
                Height = 160,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.ToolWindow,
                ResizeMode = ResizeMode.NoResize
            };
            MicroEngWpfUiTheme.ApplyTo(window);

            var panel = new Grid { Margin = new Thickness(12) };
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var label = new TextBlock { Text = caption, Margin = new Thickness(0, 0, 0, 6) };
            var box = new Wpf.Ui.Controls.TextBox();
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
            var ok = new Wpf.Ui.Controls.Button { Content = "OK", Width = 70, Margin = new Thickness(4), IsDefault = true };
            var cancel = new Wpf.Ui.Controls.Button { Content = "Cancel", Width = 70, Margin = new Thickness(4), IsCancel = true };
            ok.Click += (_, __) => window.DialogResult = true;
            cancel.Click += (_, __) => window.DialogResult = false;
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);

            panel.Children.Add(label);
            panel.Children.Add(box);
            panel.Children.Add(buttons);
            Grid.SetRow(label, 0);
            Grid.SetRow(box, 1);
            Grid.SetRow(buttons, 2);

            window.Content = panel;
            return window.ShowDialog() == true ? box.Text : null;
        }

        private void ShowSnackbar(string title, string message, WpfUiControls.ControlAppearance appearance, WpfUiControls.SymbolRegular icon)
        {
            if (SnackbarPresenter == null)
            {
                return;
            }

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

        private class ColumnsDialog : Window
        {
            private readonly List<CheckBox> _boxes = new();
            private readonly CheckBox _selectAll;
            private readonly Wpf.Ui.Controls.TextBox _searchBox;
            private bool _updatingSelectAll;

            public IEnumerable<string> VisibleAttributeIds => _boxes.Where(cb => cb.IsChecked == true).Select(cb => cb.Tag as string);

            public ColumnsDialog(IEnumerable<DataMatrixAttributeDefinition> attributes, IList<string> currentVisibleOrder)
            {
                // Ensure WPF-UI resources exist before controls are created (so implicit styles apply).
                Resources.MergedDictionaries.Add(new ResourceDictionary { Source = MicroEngResourceUris.WpfUiRoot });

                Title = "Columns";
                Width = 420;
                Height = 560;
                WindowStartupLocation = WindowStartupLocation.CenterScreen;
                WindowStyle = WindowStyle.ToolWindow;
                ResizeMode = ResizeMode.CanResize;
                SetResourceReference(BackgroundProperty, "ApplicationBackgroundBrush");
                SetResourceReference(ForegroundProperty, "TextFillColorPrimaryBrush");

                var ordered = currentVisibleOrder ?? attributes.Select(a => a.Id).ToList();
                var orderedSet = new HashSet<string>(ordered ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
                var sorted = (attributes ?? Enumerable.Empty<DataMatrixAttributeDefinition>()).ToList();

                _searchBox = new Wpf.Ui.Controls.TextBox
                {
                    PlaceholderText = "Search columns...",
                    Margin = new Thickness(0, 0, 0, 8)
                };
                _searchBox.TextChanged += (_, __) => ApplySearchFilter();

                _selectAll = new CheckBox
                {
                    Content = "Select all",
                    IsThreeState = true,
                    Margin = new Thickness(0, 0, 0, 8)
                };
                _selectAll.Click += (_, __) =>
                {
                    if (_updatingSelectAll) return;
                    if (_selectAll.IsChecked == true) SetAll(true);
                    else if (_selectAll.IsChecked == false) SetAll(false);
                    else
                    {
                        // When indeterminate and user clicks, treat as "select all"
                        SetAll(true);
                        _selectAll.IsChecked = true;
                    }
                };

                var stack = new StackPanel { Orientation = Orientation.Vertical };
                foreach (var attr in sorted)
                {
                    var cb = new CheckBox
                    {
                        Content = $"{attr.Category}: {attr.PropertyName}",
                        Tag = attr.Id,
                        IsChecked = orderedSet.Count > 0 ? orderedSet.Contains(attr.Id) : attr.IsVisibleByDefault,
                        Margin = new Thickness(0, 2, 0, 2)
                    };
                    cb.Checked += (_, __) => UpdateSelectAllState();
                    cb.Unchecked += (_, __) => UpdateSelectAllState();
                    _boxes.Add(cb);
                    stack.Children.Add(cb);
                }

                var scroll = new ScrollViewer
                {
                    Content = stack,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                };

                var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 8, 0, 0) };
                var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(4), IsDefault = true };
                var cancel = new Button { Content = "Cancel", Width = 80, Margin = new Thickness(4), IsCancel = true };
                // Use implicit WPF-UI styling; avoid referencing optional custom resource keys.

                ok.Click += (_, __) => DialogResult = true;
                cancel.Click += (_, __) => DialogResult = false;

                buttons.Children.Add(ok);
                buttons.Children.Add(cancel);

                var top = new DockPanel { LastChildFill = true };
                DockPanel.SetDock(_searchBox, Dock.Top);
                top.Children.Add(_searchBox);
                DockPanel.SetDock(_selectAll, Dock.Top);
                top.Children.Add(_selectAll);
                top.Children.Add(scroll);

                // CardControl uses a 3-column header/content template and will offset Content to the right.
                // For general container layout we use Card (Gallery pattern).
                var card = new Wpf.Ui.Controls.Card
                {
                    Padding = new Thickness(12),
                    Margin = new Thickness(8),
                    Content = top
                };

                var root = new DockPanel { Margin = new Thickness(0) };
                DockPanel.SetDock(buttons, Dock.Bottom);
                root.Children.Add(buttons);
                root.Children.Add(card);
                Content = root;

                MicroEngWpfUiTheme.ApplyTo(this);
                UpdateSelectAllState();
            }

            private void ApplySearchFilter()
            {
                var query = (_searchBox?.Text ?? string.Empty).Trim();
                if (query.Length == 0)
                {
                    foreach (var cb in _boxes) cb.Visibility = Visibility.Visible;
                    return;
                }

                foreach (var cb in _boxes)
                {
                    var text = cb.Content as string ?? string.Empty;
                    cb.Visibility = text.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0
                        ? Visibility.Visible
                        : Visibility.Collapsed;
                }
            }

            private void SetAll(bool value)
            {
                foreach (var cb in _boxes)
                {
                    cb.IsChecked = value;
                }
                UpdateSelectAllState();
            }

            private void UpdateSelectAllState()
            {
                try
                {
                    _updatingSelectAll = true;
                    if (_boxes.Count == 0)
                    {
                        _selectAll.IsChecked = false;
                        return;
                    }

                    var checkedCount = _boxes.Count(b => b.IsChecked == true);
                    if (checkedCount == 0) _selectAll.IsChecked = false;
                    else if (checkedCount == _boxes.Count) _selectAll.IsChecked = true;
                    else _selectAll.IsChecked = null;
                }
                finally
                {
                    _updatingSelectAll = false;
                }
            }
        }
    }
}
