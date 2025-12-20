using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Navisworks.Api;
using Microsoft.Win32;
using NavisApp = Autodesk.Navisworks.Api.Application;

namespace MicroEng.Navisworks
{
    public partial class DataMatrixControl : UserControl
    {
        static DataMatrixControl()
        {
            AssemblyResolver.EnsureRegistered();
        }

        private readonly DataMatrixRowBuilder _builder = new();
        private readonly IDataMatrixPresetManager _presetManager = new InMemoryPresetManager();
        private readonly DataMatrixExporter _exporter = new();

        private List<DataMatrixAttributeDefinition> _attributes = new();
        private List<DataMatrixRow> _allRows = new();
        private ObservableCollection<DataMatrixRow> _viewRows = new();
        private readonly Dictionary<Guid, ModelItem> _itemCache = new();

        private bool _syncEnabled;
        private DataMatrixViewPreset _currentPreset;

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

            MatrixGrid.SelectionChanged += MatrixGrid_SelectionChanged;
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
                _allRows.Clear();
                RowsStatus.Text = "Rows: 0";
                ProfileStatus.Text = $"Profile: {profile} (no sessions)";
                return;
            }

            DataScraperCache.LastSession = session;
            var built = _builder.Build(session);
            _attributes = built.Attributes;
            _allRows = built.Rows;
            _itemCache.Clear();
            BuildColumns();
            ApplyFiltersAndSorts();
            LoadPresets(profile);
            ProfileStatus.Text = $"Profile: {profile} @ {session.Timestamp:t}";
        }

        private void BuildColumns()
        {
            MatrixGrid.Columns.Clear();
            MatrixGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Element",
                Binding = new System.Windows.Data.Binding("ElementDisplayName"),
                Width = 180
            });

            var visibleIds = _currentPreset?.VisibleAttributeIds?.ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var attr in _attributes.Where(a => visibleIds == null ? a.IsVisibleByDefault : visibleIds.Contains(a.Id)))
            {
                MatrixGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = attr.DisplayName ?? attr.PropertyName,
                    Binding = new System.Windows.Data.Binding($"Values[{attr.Id}]"),
                    Width = attr.DefaultWidth > 0 ? attr.DefaultWidth : 140
                });
            }
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
            if (!_syncEnabled) return;
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
            BuildColumns();
            ApplyFiltersAndSorts();
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
            preset.VisibleAttributeIds = MatrixGrid.Columns
                .Skip(1) // skip Element column
                .Select(c => ((System.Windows.Data.Binding)((DataGridBoundColumn)c).Binding).Path.Path.Replace("Values[", string.Empty).TrimEnd(']'))
                .ToList();

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
            var dialog = new ColumnsDialog(_attributes, MatrixGrid.Columns.Skip(1).Select(c => ((System.Windows.Data.Binding)((DataGridBoundColumn)c).Binding).Path.Path.Replace("Values[", string.Empty).TrimEnd(']')).ToList());
            if (dialog.ShowDialog() == true)
            {
                var visibleIds = dialog.VisibleAttributeIds.ToList();
                MatrixGrid.Columns.Clear();
                MatrixGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = "Element",
                    Binding = new System.Windows.Data.Binding("ElementDisplayName"),
                    Width = 180
                });
                foreach (var id in visibleIds)
                {
                    var attr = _attributes.FirstOrDefault(a => string.Equals(a.Id, id, StringComparison.OrdinalIgnoreCase));
                    if (attr == null) continue;
                    MatrixGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = attr.DisplayName ?? attr.PropertyName,
                        Binding = new System.Windows.Data.Binding($"Values[{attr.Id}]"),
                        Width = attr.DefaultWidth > 0 ? attr.DefaultWidth : 140
                    });
                }
            }
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
            var columns = MatrixGrid.Columns
                .Skip(1)
                .Select(c => ((System.Windows.Data.Binding)((DataGridBoundColumn)c).Binding).Path.Path.Replace("Values[", string.Empty).TrimEnd(']'))
                .Select(id => _attributes.First(a => string.Equals(a.Id, id, StringComparison.OrdinalIgnoreCase)))
                .ToList();

            var session = DataScraperCache.LastSession ?? DataScraperCache.AllSessions.FirstOrDefault();
            _exporter.ExportCsv(dlg.FileName, columns, rows, session, _currentPreset);
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
                var sorted = attributes.OrderBy(a => a.Category).ThenBy(a => a.PropertyName).ToList();

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
