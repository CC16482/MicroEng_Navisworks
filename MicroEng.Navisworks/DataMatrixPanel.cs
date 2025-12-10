using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using NavisApp = Autodesk.Navisworks.Api.Application;

namespace MicroEng.Navisworks
{
    internal class DataMatrixPanel : UserControl
    {
        private readonly DataMatrixRowBuilder _builder = new();
        private readonly IDataMatrixPresetManager _presetManager = new InMemoryPresetManager();
        private readonly DataMatrixExporter _exporter = new();

        private readonly ToolStripComboBox _profileCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
        private readonly ToolStripComboBox _presetCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160 };
        private readonly ToolStripButton _refreshButton = new("Refresh");
        private readonly ToolStripButton _savePresetButton = new("Save View");
        private readonly ToolStripButton _saveAsPresetButton = new("Save View As...");
        private readonly ToolStripButton _deletePresetButton = new("Delete View");
        private readonly ToolStripButton _columnsButton = new("Columns...");
        private readonly ToolStripButton _syncButton = new("Sync selection") { CheckOnClick = true };
        private readonly ToolStripButton _selectedOnlyButton = new("Selected only") { CheckOnClick = true };
        private readonly ToolStripDropDownButton _exportDropDown = new("Export");

        private readonly DataGridView _grid = new();
        private readonly StatusStrip _statusStrip = new();
        private readonly ToolStripStatusLabel _statusRows = new();
        private readonly ToolStripStatusLabel _statusProfile = new();
        private readonly ToolStripStatusLabel _statusPreset = new();
        private readonly Label _emptyLabel = new()
        {
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Text = "Run Data Scraper to populate data.",
            ForeColor = System.Drawing.Color.DarkSlateGray,
            Visible = false
        };

        private ScrapeSession _session;
        private List<DataMatrixAttributeDefinition> _attributes = new();
        private List<DataMatrixRow> _allRows = new();
        private List<DataMatrixRow> _viewRows = new();
        private readonly Dictionary<string, DataMatrixColumnFilter> _filters = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<DataMatrixSortDefinition> _sorts = new();
        private readonly Dictionary<string, ModelItem> _modelCache = new(StringComparer.OrdinalIgnoreCase);
        private bool _suppressSelectionSync;

        public DataMatrixPanel()
        {
            Dock = DockStyle.Fill;
            DataScraperCache.SessionAdded += OnSessionAdded;
            BuildUi();
            LoadProfiles();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                DataScraperCache.SessionAdded -= OnSessionAdded;
            }
            base.Dispose(disposing);
        }

        private void BuildUi()
        {
            var toolStrip = new ToolStrip
            {
                GripStyle = ToolStripGripStyle.Hidden,
                Dock = DockStyle.Top,
                RenderMode = ToolStripRenderMode.System
            };

            toolStrip.Items.Add(new ToolStripLabel("Data Source:"));
            toolStrip.Items.Add(_profileCombo);
            toolStrip.Items.Add(_refreshButton);
            toolStrip.Items.Add(new ToolStripSeparator());
            toolStrip.Items.Add(new ToolStripLabel("View:"));
            toolStrip.Items.Add(_presetCombo);
            toolStrip.Items.Add(_savePresetButton);
            toolStrip.Items.Add(_saveAsPresetButton);
            toolStrip.Items.Add(_deletePresetButton);
            toolStrip.Items.Add(new ToolStripSeparator());
            toolStrip.Items.Add(_columnsButton);
            toolStrip.Items.Add(new ToolStripSeparator());
            toolStrip.Items.Add(_syncButton);
            toolStrip.Items.Add(_selectedOnlyButton);
            toolStrip.Items.Add(new ToolStripSeparator());
            toolStrip.Items.Add(_exportDropDown);

            _profileCombo.SelectedIndexChanged += (s, e) => OnProfileChanged();
            _refreshButton.Click += (s, e) => LoadProfiles();
            _presetCombo.SelectedIndexChanged += (s, e) => OnPresetChanged();
            _savePresetButton.Click += (s, e) => SavePreset();
            _saveAsPresetButton.Click += (s, e) => SavePreset(asNew: true);
            _deletePresetButton.Click += (s, e) => DeletePreset();
            _columnsButton.Click += (s, e) => ShowColumnsDialog();
            _syncButton.CheckedChanged += (s, e) => SyncSelectionToNavisworks();
            _selectedOnlyButton.CheckedChanged += (s, e) => ApplyFiltersAndSorts();

            _exportDropDown.DropDownItems.Add("Export filtered to CSV...", null, (s, e) => ExportCsv(filtered: true));
            _exportDropDown.DropDownItems.Add("Export all to CSV...", null, (s, e) => ExportCsv(filtered: false));

            _grid.Dock = DockStyle.Fill;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.AutoGenerateColumns = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _grid.RowHeadersVisible = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = true;
            _grid.BackgroundColor = System.Drawing.Color.White;
            _grid.BorderStyle = BorderStyle.None;
            _grid.CellFormatting += Grid_CellFormatting;
            _grid.CellEndEdit += Grid_CellEndEdit;
            _grid.ColumnHeaderMouseClick += Grid_ColumnHeaderMouseClick;
            _grid.SelectionChanged += Grid_SelectionChanged;

            _statusStrip.Items.AddRange(new ToolStripItem[]
            {
                _statusRows,
                new ToolStripStatusLabel { Spring = true },
                _statusProfile,
                new ToolStripStatusLabel { Spring = true },
                _statusPreset
            });

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            layout.Controls.Add(toolStrip, 0, 0);
            var gridHost = new Panel { Dock = DockStyle.Fill };
            gridHost.Controls.Add(_grid);
            gridHost.Controls.Add(_emptyLabel);
            layout.Controls.Add(gridHost, 0, 1);
            layout.Controls.Add(_statusStrip, 0, 2);

            Controls.Add(layout);
        }

        private void LoadProfiles()
        {
            var profiles = DataScraperCache.AllSessions
                .Select(s => string.IsNullOrWhiteSpace(s.ProfileName) ? "Default" : s.ProfileName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(p => p)
                .ToList();

            _profileCombo.Items.Clear();
            foreach (var p in profiles)
            {
                _profileCombo.Items.Add(p);
            }

            var preferred = DataScraperCache.LastSession?.ProfileName ?? profiles.FirstOrDefault();
            if (preferred != null)
            {
                var idx = _profileCombo.Items.IndexOf(preferred);
                if (idx >= 0)
                    _profileCombo.SelectedIndex = idx;
                else if (_profileCombo.Items.Count > 0)
                    _profileCombo.SelectedIndex = 0;
            }

            var hasData = profiles.Any();
            ToggleEmptyState(!hasData);
            SetToolbarEnabled(hasData);

            if (hasData)
            {
                OnProfileChanged();
            }
        }

        private void OnSessionAdded(ScrapeSession session)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => OnSessionAdded(session)));
                return;
            }

            LoadProfiles();
        }

        private void ToggleEmptyState(bool showEmpty)
        {
            _emptyLabel.Visible = showEmpty;
            _grid.Visible = !showEmpty;
        }

        private void SetToolbarEnabled(bool enabled)
        {
            var items = _profileCombo.Owner != null
                ? _profileCombo.Owner.Items.Cast<ToolStripItem>()
                : Enumerable.Empty<ToolStripItem>();
            foreach (var item in items)
            {
                item.Enabled = enabled;
            }
            _statusStrip.Enabled = enabled;
        }

        private void OnProfileChanged()
        {
            if (_profileCombo.SelectedItem == null) return;
            var profile = _profileCombo.SelectedItem.ToString();
            _session = DataScraperCache.AllSessions
                .Where(s => string.Equals(s.ProfileName ?? "Default", profile, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(s => s.Timestamp)
                .FirstOrDefault();

            DataScraperCache.LastSession = _session;
            LoadPresets(profile);
            RebuildFromSession();
        }

        private void LoadPresets(string profile)
        {
            _presetCombo.Items.Clear();
            _presetCombo.Items.Add("(None)");
            foreach (var preset in _presetManager.GetPresets(profile))
            {
                _presetCombo.Items.Add(new PresetComboItem(preset));
            }
            _presetCombo.SelectedIndex = 0;
        }

        private void OnPresetChanged()
        {
            if (_session == null) return;
            ApplyPresetFromSelection();
        }

        private void ApplyPresetFromSelection()
        {
            DataMatrixViewPreset preset = null;
            if (_presetCombo.SelectedItem is PresetComboItem item)
            {
                preset = item.Preset;
            }
            _filters.Clear();
            _sorts.Clear();
            if (preset != null)
            {
                foreach (var filter in preset.Filters ?? new List<DataMatrixColumnFilter>())
                {
                    if (!string.IsNullOrWhiteSpace(filter.AttributeId))
                    {
                        _filters[filter.AttributeId] = filter;
                    }
                }
                if (preset.SortDefinitions != null)
                {
                    _sorts.AddRange(preset.SortDefinitions);
                }
            }
            BuildGridColumns(preset);
            ApplyFiltersAndSorts(preset);
        }

        private void RebuildFromSession()
        {
            if (_session == null)
            {
                ToggleEmptyState(true);
                return;
            }

            ToggleEmptyState(false);
            _modelCache.Clear();
            _filters.Clear();
            _sorts.Clear();

            var result = _builder.Build(_session);
            _attributes = result.Attributes;
            _allRows = result.Rows;

            BuildGridColumns(null);
            ApplyFiltersAndSorts();
        }

        private void BuildGridColumns(DataMatrixViewPreset preset)
        {
            BuildGridColumns(preset, preset?.VisibleAttributeIds);
        }

        private void BuildGridColumns(DataMatrixViewPreset preset, IList<string> orderedVisible)
        {
            _grid.Columns.Clear();
            var visibleIds = orderedVisible?.ToHashSet(StringComparer.OrdinalIgnoreCase)
                             ?? preset?.VisibleAttributeIds?.ToHashSet(StringComparer.OrdinalIgnoreCase)
                             ?? new HashSet<string>(_attributes.Where(a => a.IsVisibleByDefault).Select(a => a.Id), StringComparer.OrdinalIgnoreCase);

            var orderMap = orderedVisible != null
                ? orderedVisible.Select((id, idx) => (id, idx)).ToDictionary(t => t.id, t => t.idx, StringComparer.OrdinalIgnoreCase)
                : null;

            var itemCol = new DataGridViewTextBoxColumn
            {
                Name = "Element",
                HeaderText = "Element",
                DataPropertyName = nameof(DataMatrixRow.ElementDisplayName),
                SortMode = DataGridViewColumnSortMode.Programmatic,
                ReadOnly = true,
                Width = 180,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                FillWeight = 1
            };
            _grid.Columns.Add(itemCol);

            var keyCol = new DataGridViewTextBoxColumn
            {
                Name = "ItemKey",
                HeaderText = "Item Key",
                DataPropertyName = nameof(DataMatrixRow.ItemKey),
                Visible = false,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                FillWeight = 1
            };
            _grid.Columns.Add(keyCol);

            IEnumerable<DataMatrixAttributeDefinition> orderedAttrs = _attributes;
            if (orderMap != null && orderMap.Count > 0)
            {
                orderedAttrs = _attributes.OrderBy(a => orderMap.ContainsKey(a.Id) ? orderMap[a.Id] : int.MaxValue);
            }
            else
            {
                orderedAttrs = _attributes.OrderBy(a => a.DisplayOrder);
            }

            foreach (var attr in orderedAttrs)
            {
            var col = new DataGridViewTextBoxColumn
            {
                Name = attr.Id,
                HeaderText = string.IsNullOrWhiteSpace(attr.Category)
                    ? (attr.DisplayName ?? attr.PropertyName)
                    : $"{attr.Category}: {attr.DisplayName ?? attr.PropertyName}",
                Tag = attr,
                SortMode = DataGridViewColumnSortMode.Programmatic,
                ReadOnly = !attr.IsEditable,
                Width = attr.DefaultWidth > 0 ? attr.DefaultWidth : 140,
                    Visible = visibleIds.Contains(attr.Id),
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                    FillWeight = 1
                };
                _grid.Columns.Add(col);
            }
        }

        private void ApplyFiltersAndSorts(DataMatrixViewPreset preset = null)
        {
            if (_session == null)
            {
                _grid.DataSource = null;
                return;
            }

            IEnumerable<DataMatrixRow> query = _allRows;

            if (_selectedOnlyButton.Checked)
            {
                var selectedKeys = GetSelectedItemKeys();
                query = query.Where(r => selectedKeys.Contains(r.ItemKey));
            }

            foreach (var filter in _filters.Values)
            {
                query = query.Where(r => MatchesFilter(r, filter));
            }

            var sorts = preset?.SortDefinitions?.Any() == true ? preset.SortDefinitions : _sorts;

            IOrderedEnumerable<DataMatrixRow> ordered = null;
            foreach (var sort in sorts.OrderBy(s => s.Priority))
            {
                Func<DataMatrixRow, object> keySelector = r => GetValueForSort(r, sort.AttributeId);
                if (ordered == null)
                {
                    ordered = sort.Direction == SortDirection.Descending
                        ? query.OrderByDescending(keySelector)
                        : query.OrderBy(keySelector);
                }
                else
                {
                    ordered = sort.Direction == SortDirection.Descending
                        ? ordered.ThenByDescending(keySelector)
                        : ordered.ThenBy(keySelector);
                }
            }

            _viewRows = (ordered ?? query).ToList();

            _suppressSelectionSync = true;
            _grid.DataSource = new BindingList<DataMatrixRow>(_viewRows);
            _suppressSelectionSync = false;

            _statusRows.Text = $"Rows: {_viewRows.Count} (from {_allRows.Count})";
            _statusProfile.Text = $"Profile: {_session.ProfileName}";
            var presetName = preset?.Name ?? (_presetCombo.SelectedItem as PresetComboItem)?.Preset.Name;
            _statusPreset.Text = $"View: {presetName ?? "Default"}";
        }

        private bool MatchesFilter(DataMatrixRow row, DataMatrixColumnFilter filter)
        {
            object valObj = null;
            if (string.Equals(filter.AttributeId, "Element", StringComparison.OrdinalIgnoreCase))
            {
                valObj = row.ElementDisplayName;
            }
            else if (row.Values.TryGetValue(filter.AttributeId, out var value))
            {
                valObj = value;
            }

            var valStr = valObj?.ToString() ?? string.Empty;

            switch (filter.Operator)
            {
                case DataMatrixFilterOperator.None:
                    return true;
                case DataMatrixFilterOperator.Contains:
                    return valStr.IndexOf(filter.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase) >= 0;
                case DataMatrixFilterOperator.Equals:
                    return string.Equals(valStr, filter.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
                case DataMatrixFilterOperator.NotEquals:
                    return !string.Equals(valStr, filter.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
                case DataMatrixFilterOperator.Blank:
                    return string.IsNullOrWhiteSpace(valStr);
                case DataMatrixFilterOperator.NotBlank:
                    return !string.IsNullOrWhiteSpace(valStr);
                case DataMatrixFilterOperator.InList:
                    var set = filter.ValuesList?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    return set.Contains(valStr);
                case DataMatrixFilterOperator.GreaterThan:
                case DataMatrixFilterOperator.LessThan:
                    if (double.TryParse(valStr, out var num) && double.TryParse(filter.Value, out var target))
                    {
                        return filter.Operator == DataMatrixFilterOperator.GreaterThan ? num > target : num < target;
                    }
                    return false;
                default:
                    return true;
            }
        }

        private object GetValueForSort(DataMatrixRow row, string attributeId)
        {
            if (string.Equals(attributeId, "Element", StringComparison.OrdinalIgnoreCase))
            {
                return row.ElementDisplayName ?? string.Empty;
            }
            if (row.Values.TryGetValue(attributeId, out var val))
            {
                return val ?? string.Empty;
            }
            return string.Empty;
        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.RowIndex < 0) return;
            var column = _grid.Columns[e.ColumnIndex];
            if (column.Tag is DataMatrixAttributeDefinition attr && _grid.Rows[e.RowIndex].DataBoundItem is DataMatrixRow row)
            {
                if (row.Values.TryGetValue(attr.Id, out var value))
                {
                    e.Value = value;
                    e.FormattingApplied = true;
                }
            }
        }

        private void Grid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var column = _grid.Columns[e.ColumnIndex];
            if (column.Tag is not DataMatrixAttributeDefinition attr) return;
            if (_grid.Rows[e.RowIndex].DataBoundItem is not DataMatrixRow row) return;
            var newValue = _grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            row.Values[attr.Id] = newValue;
            // TODO: push edits back to Navisworks model if required.
        }

        private void Grid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ShowFilterMenu(e.ColumnIndex, e.Location);
                return;
            }

            var column = _grid.Columns[e.ColumnIndex];
            var attrId = column.Tag is DataMatrixAttributeDefinition attr ? attr.Id : column.Name;

            var existing = _sorts.FirstOrDefault(s => string.Equals(s.AttributeId, attrId, StringComparison.OrdinalIgnoreCase));
            var nextDirection = SortDirection.Ascending;
            if (existing != null)
            {
                nextDirection = existing.Direction == SortDirection.Ascending ? SortDirection.Descending : SortDirection.None;
            }

            if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
            {
                _sorts.RemoveAll(s => string.Equals(s.AttributeId, attrId, StringComparison.OrdinalIgnoreCase));
                if (nextDirection != SortDirection.None)
                {
                    _sorts.Add(new DataMatrixSortDefinition { AttributeId = attrId, Direction = nextDirection, Priority = _sorts.Count });
                }
            }
            else
            {
                _sorts.Clear();
                if (nextDirection != SortDirection.None)
                {
                    _sorts.Add(new DataMatrixSortDefinition { AttributeId = attrId, Direction = nextDirection, Priority = 0 });
                }
            }

            ApplyFiltersAndSorts();
        }

        private void ShowFilterMenu(int columnIndex, Point location)
        {
            var column = _grid.Columns[columnIndex];
            var attrId = column.Tag is DataMatrixAttributeDefinition attr ? attr.Id : column.Name;

            var menu = new ContextMenuStrip();
            menu.Items.Add("Filter...", null, (s, e) => OpenFilterDialog(attrId, column.HeaderText));
            menu.Items.Add("Clear filter", null, (s, e) =>
            {
                if (_filters.Remove(attrId))
                {
                    ApplyFiltersAndSorts();
                }
            });
            menu.Show(_grid, location);
        }

        private void OpenFilterDialog(string attributeId, string header)
        {
            _filters.TryGetValue(attributeId, out var existing);
            using var dialog = new DataMatrixFilterDialog(header, existing);
            if (dialog.ShowDialog() == DialogResult.OK && dialog.ResultFilter != null)
            {
                var filter = dialog.ResultFilter;
                filter.AttributeId = attributeId;
                if (filter.Operator == DataMatrixFilterOperator.None)
                {
                    _filters.Remove(attributeId);
                }
                else
                {
                    _filters[attributeId] = filter;
                }
                ApplyFiltersAndSorts();
            }
        }

        private void Grid_SelectionChanged(object sender, EventArgs e)
        {
            if (_suppressSelectionSync) return;
            if (!_syncButton.Checked) return;
            SyncSelectionToNavisworks();
        }

        private void SyncSelectionToNavisworks()
        {
            if (!_syncButton.Checked || _session == null) return;
            var doc = NavisApp.ActiveDocument;
            if (doc == null) return;

            var selectedRows = _grid.SelectedRows
                .Cast<DataGridViewRow>()
                .Select(r => r.DataBoundItem as DataMatrixRow)
                .Where(r => r != null)
                .ToList();

            var collection = new ModelItemCollection();
            foreach (var row in selectedRows)
            {
                var item = ResolveModelItem(row.ItemKey);
                if (item != null)
                {
                    collection.Add(item);
                }
            }

            doc.CurrentSelection.CopyFrom(collection);
        }

        private HashSet<string> GetSelectedItemKeys()
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var doc = NavisApp.ActiveDocument;
            if (doc?.CurrentSelection?.SelectedItems != null)
            {
                foreach (ModelItem item in doc.CurrentSelection.SelectedItems)
                {
                    keys.Add(item.InstanceGuid.ToString());
                }
            }
            return keys;
        }

        private ModelItem ResolveModelItem(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return null;
            if (_modelCache.TryGetValue(key, out var cached)) return cached;
            if (!Guid.TryParse(key, out var guid)) return null;

            var doc = NavisApp.ActiveDocument;
            if (doc == null) return null;

            ModelItem Find(IEnumerable<ModelItem> items)
            {
                foreach (ModelItem item in items)
                {
                    if (item.InstanceGuid == guid) return item;
                    var child = Find(item.Children);
                    if (child != null) return child;
                }
                return null;
            }

            var found = Find(doc.Models.RootItems);
            if (found != null)
            {
                _modelCache[key] = found;
            }
            return found;
        }

        private void ShowColumnsDialog()
        {
            var visibleOrder = _grid.Columns.Cast<DataGridViewColumn>()
                .Where(c => c.Tag is DataMatrixAttributeDefinition && c.Visible)
                .Select(c => ((DataMatrixAttributeDefinition)c.Tag).Id)
                .ToList();

            using var dialog = new DataMatrixColumnsDialog(_attributes, visibleOrder);
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var visible = dialog.VisibleAttributeIds.ToList();
                BuildGridColumns(new DataMatrixViewPreset { VisibleAttributeIds = visible }, visible);
                ApplyFiltersAndSorts();
            }
        }

        private void SavePreset(bool asNew = false)
        {
            if (_session == null) return;
            var preset = (_presetCombo.SelectedItem as PresetComboItem)?.Preset;
            if (preset == null || asNew)
            {
                var name = Prompt.ShowDialog("Preset name", "Save View");
                if (string.IsNullOrWhiteSpace(name)) return;
                preset = new DataMatrixViewPreset
                {
                    Id = Guid.NewGuid().ToString(),
                    Name = name,
                    ScraperProfileName = _session.ProfileName
                };
            }

            preset.VisibleAttributeIds = _grid.Columns
                .Cast<DataGridViewColumn>()
                .Where(c => c.Tag is DataMatrixAttributeDefinition && c.Visible)
                .Select(c => ((DataMatrixAttributeDefinition)c.Tag).Id)
                .ToList();
            preset.SortDefinitions = _sorts.ToList();
            preset.Filters = _filters.Values.ToList();

            _presetManager.SavePreset(preset);
            LoadPresets(_session.ProfileName);
            var idx = _presetCombo.Items.Cast<object>().ToList().FindIndex(i => i is PresetComboItem pci && pci.Preset.Id == preset.Id);
            if (idx >= 0) _presetCombo.SelectedIndex = idx;
        }

        private void DeletePreset()
        {
            if (_presetCombo.SelectedItem is not PresetComboItem item) return;
            _presetManager.DeletePreset(item.Preset.Id);
            LoadPresets(_session.ProfileName);
        }

        private void ExportCsv(bool filtered)
        {
            if (_session == null) return;
            using var dialog = new SaveFileDialog
            {
                Filter = "CSV files|*.csv",
                FileName = $"{_session.ProfileName}_DataMatrix.csv"
            };
            if (dialog.ShowDialog() != DialogResult.OK) return;
            var visibleCols = _grid.Columns.Cast<DataGridViewColumn>()
                .Where(c => c.Tag is DataMatrixAttributeDefinition)
                .Where(c => c.Visible)
                .Select(c => (DataMatrixAttributeDefinition)c.Tag)
                .ToList();
            var rows = filtered ? _viewRows : _allRows;
            _exporter.ExportCsv(dialog.FileName, visibleCols, rows, _session, (_presetCombo.SelectedItem as PresetComboItem)?.Preset);
        }

        private class PresetComboItem
        {
            public PresetComboItem(DataMatrixViewPreset preset) => Preset = preset;
            public DataMatrixViewPreset Preset { get; }
            public override string ToString() => Preset.Name;
        }
    }

    internal class DataMatrixFilterDialog : Form
    {
        private readonly ComboBox _operatorBox = new() { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox _valueBox = new() { Width = 220 };
        private readonly Label _valueLabel = new() { Text = "Value:", AutoSize = true };

        public DataMatrixColumnFilter ResultFilter { get; private set; }

        public DataMatrixFilterDialog(string header, DataMatrixColumnFilter existing)
        {
            Text = $"Filter - {header}";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            Width = 360;
            Height = 180;

            _operatorBox.Items.AddRange(Enum.GetNames(typeof(DataMatrixFilterOperator)));
            _operatorBox.SelectedItem = existing?.Operator.ToString() ?? DataMatrixFilterOperator.Contains.ToString();
            if (_operatorBox.SelectedIndex < 0) _operatorBox.SelectedIndex = 0;
            _valueBox.Text = existing?.Value ?? string.Empty;
            _valueBox.ReadOnly = false;
            _valueBox.Enabled = true;
            _valueBox.BackColor = System.Drawing.Color.White;

            _operatorBox.SelectedIndexChanged += (s, e) => UpdateOperatorState();
            UpdateOperatorState();

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(6) };
            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 80 };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 80 };
            buttons.Controls.Add(ok);
            buttons.Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            layout.Controls.Add(new Label { Text = "Operator:", AutoSize = true }, 0, 0);
            layout.Controls.Add(_operatorBox, 1, 0);
            layout.Controls.Add(_valueLabel, 0, 1);
            layout.Controls.Add(_valueBox, 1, 1);

            Controls.Add(layout);
            Controls.Add(buttons);

            AcceptButton = ok;
            CancelButton = cancel;

            ok.Click += (s, e) =>
            {
                var op = (DataMatrixFilterOperator)Enum.Parse(typeof(DataMatrixFilterOperator), _operatorBox.SelectedItem.ToString());
                var filter = new DataMatrixColumnFilter
                {
                    Operator = op,
                    Value = _valueBox.Text?.Trim()
                };
                if (op == DataMatrixFilterOperator.InList && !string.IsNullOrWhiteSpace(filter.Value))
                {
                    filter.ValuesList = filter.Value.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(v => v.Trim()).ToList();
                }
                ResultFilter = filter;
                DialogResult = DialogResult.OK;
            };

            Shown += (s, e) =>
            {
                _valueBox.Focus();
                _valueBox.SelectAll();
            };
        }

        private void UpdateOperatorState()
        {
            if (_operatorBox.SelectedItem == null) return;
            var op = (DataMatrixFilterOperator)Enum.Parse(typeof(DataMatrixFilterOperator), _operatorBox.SelectedItem.ToString());
            var needsValue = op != DataMatrixFilterOperator.Blank && op != DataMatrixFilterOperator.NotBlank;
            _valueBox.Enabled = needsValue;
            _valueLabel.Enabled = needsValue;
            _valueBox.Visible = needsValue;
            _valueLabel.Visible = needsValue;
        }
    }

    internal class DataMatrixColumnsDialog : Form
    {
        private readonly TreeView _tree = new() { Dock = DockStyle.Fill, CheckBoxes = true, HideSelection = false };
        private readonly TextBox _searchBox = new() { Dock = DockStyle.Fill };
        private readonly Button _selectAll = new() { Text = "Select All", Width = 90 };
        private readonly Button _deselectAll = new() { Text = "Deselect All", Width = 90 };
        private readonly ListView _selectedList = new() { Dock = DockStyle.Fill, View = System.Windows.Forms.View.Details, FullRowSelect = true, HideSelection = false };
        private readonly Button _moveUp = new() { Text = "Up", Width = 60 };
        private readonly Button _moveDown = new() { Text = "Down", Width = 60 };
        private readonly Button _remove = new() { Text = "Remove", Width = 70 };
        private readonly List<DataMatrixAttributeDefinition> _attributes;
        private readonly HashSet<string> _initialVisible;
        private bool _suppressCheck;

        public IEnumerable<string> VisibleAttributeIds => _selectedList.Items.Cast<ListViewItem>().Select(i => i.Tag as string);

        public DataMatrixColumnsDialog(IEnumerable<DataMatrixAttributeDefinition> attributes, IList<string> currentVisibleOrder)
        {
            Text = "Columns";
            Width = 720;
            Height = 520;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;

            _attributes = attributes.ToList();
            _initialVisible = new HashSet<string>(currentVisibleOrder ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);

            var splitter = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = 360 };
            splitter.Panel1.Padding = new Padding(6);
            splitter.Panel2.Padding = new Padding(6);

            // Left side: search + tree + toggles
            var leftLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 3 };
            leftLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
            leftLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            leftLayout.Controls.Add(new Label { Text = "Search:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            leftLayout.SetColumnSpan(_searchBox, 2);
            leftLayout.Controls.Add(_searchBox, 0, 1);

            var togglePanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill };
            togglePanel.Controls.Add(_selectAll);
            togglePanel.Controls.Add(_deselectAll);
            leftLayout.Controls.Add(togglePanel, 0, 2);
            leftLayout.SetColumnSpan(togglePanel, 2);
            leftLayout.Controls.Add(_tree, 0, 3);
            leftLayout.SetColumnSpan(_tree, 2);
            leftLayout.RowCount = 4;
            leftLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            splitter.Panel1.Controls.Add(leftLayout);

            // Right side: selected list + controls
            _selectedList.Columns.Add("Visible Columns (order)", 280);
            var rightLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2 };
            rightLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 75));
            rightLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            rightLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rightLayout.Controls.Add(_selectedList, 0, 0);
            rightLayout.SetColumnSpan(_selectedList, 2);

            var reorderPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill };
            reorderPanel.Controls.Add(_moveUp);
            reorderPanel.Controls.Add(_moveDown);
            reorderPanel.Controls.Add(_remove);
            rightLayout.Controls.Add(reorderPanel, 0, 1);
            rightLayout.SetColumnSpan(reorderPanel, 2);

            splitter.Panel2.Controls.Add(rightLayout);

            // Bottom buttons
            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, Padding = new Padding(6) };
            buttons.Controls.Add(new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 80 });
            buttons.Controls.Add(new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 80 });

            Controls.Add(splitter);
            Controls.Add(buttons);

            _tree.AfterCheck += Tree_AfterCheck;
            _searchBox.TextChanged += (s, e) => BuildTree();
            _selectAll.Click += (s, e) => SetAllChecks(true);
            _deselectAll.Click += (s, e) => SetAllChecks(false);
            _moveUp.Click += (s, e) => MoveSelected(-1);
            _moveDown.Click += (s, e) => MoveSelected(1);
            _remove.Click += (s, e) => RemoveSelected();
            _selectedList.ItemSelectionChanged += (s, e) => SyncTreeFromSelected();

            BuildTree();
            BuildSelectedList(currentVisibleOrder);
        }

        private void BuildTree()
        {
            _suppressCheck = true;
            _tree.BeginUpdate();
            _tree.Nodes.Clear();
            var term = _searchBox.Text?.Trim() ?? string.Empty;
            var grouped = _attributes.GroupBy(a => a.Category ?? string.Empty).OrderBy(g => g.Key);
            foreach (var group in grouped)
            {
                var catNode = new TreeNode(string.IsNullOrWhiteSpace(group.Key) ? "(No Category)" : group.Key) { Tag = group.Key };
                foreach (var attr in group.OrderBy(a => a.PropertyName))
                {
                    if (!string.IsNullOrEmpty(term) &&
                        attr.PropertyName.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0 &&
                        attr.Category.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    var node = new TreeNode(attr.PropertyName) { Tag = attr, Checked = _initialVisible.Contains(attr.Id) };
                    catNode.Nodes.Add(node);
                }
                if (catNode.Nodes.Count > 0)
                {
                    catNode.Checked = catNode.Nodes.Cast<TreeNode>().All(n => n.Checked);
                    _tree.Nodes.Add(catNode);
                }
            }
            _tree.EndUpdate();
            _suppressCheck = false;
        }

        private void BuildSelectedList(IList<string> currentVisibleOrder)
        {
            _selectedList.Items.Clear();
            var order = currentVisibleOrder ?? _initialVisible.ToList();
            foreach (var id in order)
            {
                var attr = _attributes.FirstOrDefault(a => string.Equals(a.Id, id, StringComparison.OrdinalIgnoreCase));
                if (attr == null) continue;
                var item = new ListViewItem($"{attr.Category}: {attr.PropertyName}") { Tag = attr.Id };
                _selectedList.Items.Add(item);
            }
        }

        private void Tree_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (_suppressCheck) return;
            _suppressCheck = true;

            if (e.Node.Tag is string) // category node
            {
                foreach (TreeNode child in e.Node.Nodes)
                {
                    child.Checked = e.Node.Checked;
                    UpdateSelectedList(child, child.Checked);
                }
            }
            else if (e.Node.Tag is DataMatrixAttributeDefinition)
            {
                UpdateSelectedList(e.Node, e.Node.Checked);
                var parent = e.Node.Parent;
                if (parent != null)
                {
                    parent.Checked = parent.Nodes.Cast<TreeNode>().All(n => n.Checked);
                }
            }

            _suppressCheck = false;
        }

        private void UpdateSelectedList(TreeNode node, bool isChecked)
        {
            if (node.Tag is not DataMatrixAttributeDefinition attr) return;
            var existing = _selectedList.Items.Cast<ListViewItem>().FirstOrDefault(i => string.Equals(i.Tag as string, attr.Id, StringComparison.OrdinalIgnoreCase));
            if (isChecked && existing == null)
            {
                var item = new ListViewItem($"{attr.Category}: {attr.PropertyName}") { Tag = attr.Id };
                _selectedList.Items.Add(item);
            }
            else if (!isChecked && existing != null)
            {
                _selectedList.Items.Remove(existing);
            }
        }

        private void SetAllChecks(bool check)
        {
            _suppressCheck = true;
            foreach (TreeNode cat in _tree.Nodes)
            {
                cat.Checked = check;
                foreach (TreeNode child in cat.Nodes)
                {
                    child.Checked = check;
                }
            }
            _suppressCheck = false;
            if (check)
            {
                _selectedList.Items.Clear();
                foreach (var attr in _attributes)
                {
                    var item = new ListViewItem($"{attr.Category}: {attr.PropertyName}") { Tag = attr.Id };
                    _selectedList.Items.Add(item);
                }
            }
            else
            {
                _selectedList.Items.Clear();
            }
        }

        private void MoveSelected(int delta)
        {
            if (_selectedList.SelectedItems.Count == 0) return;
            var item = _selectedList.SelectedItems[0];
            var index = item.Index;
            var newIndex = index + delta;
            if (newIndex < 0 || newIndex >= _selectedList.Items.Count) return;
            _selectedList.Items.RemoveAt(index);
            _selectedList.Items.Insert(newIndex, item);
            item.Selected = true;
        }

        private void RemoveSelected()
        {
            foreach (ListViewItem item in _selectedList.SelectedItems)
            {
                _selectedList.Items.Remove(item);
                UncheckTreeNode(item.Tag as string);
            }
        }

        private void UncheckTreeNode(string id)
        {
            foreach (TreeNode cat in _tree.Nodes)
            {
                foreach (TreeNode child in cat.Nodes)
                {
                    if (child.Tag is DataMatrixAttributeDefinition attr &&
                        string.Equals(attr.Id, id, StringComparison.OrdinalIgnoreCase))
                    {
                        child.Checked = false;
                        cat.Checked = cat.Nodes.Cast<TreeNode>().All(n => n.Checked);
                        return;
                    }
                }
            }
        }

        private void SyncTreeFromSelected()
        {
            // Keep tree checks aligned with selected list if user removes items via remove button
            var selectedIds = _selectedList.Items.Cast<ListViewItem>().Select(i => i.Tag as string).ToHashSet(StringComparer.OrdinalIgnoreCase);
            _suppressCheck = true;
            foreach (TreeNode cat in _tree.Nodes)
            {
                foreach (TreeNode child in cat.Nodes)
                {
                    if (child.Tag is DataMatrixAttributeDefinition attr)
                    {
                        child.Checked = selectedIds.Contains(attr.Id);
                    }
                }
                cat.Checked = cat.Nodes.Cast<TreeNode>().All(n => n.Checked);
            }
            _suppressCheck = false;
        }
    }

    internal static class Prompt
    {
        public static string ShowDialog(string text, string caption)
        {
            using var form = new Form
            {
                Width = 400,
                Height = 160,
                Text = caption,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = false
            };
            var label = new Label { Left = 10, Top = 10, Text = text, AutoSize = true };
            var textBox = new TextBox { Left = 10, Top = 36, Width = 360 };
            var ok = new Button { Text = "OK", Left = 210, Width = 80, Top = 70, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Cancel", Left = 300, Width = 80, Top = 70, DialogResult = DialogResult.Cancel };
            form.Controls.Add(label);
            form.Controls.Add(textBox);
            form.Controls.Add(ok);
            form.Controls.Add(cancel);
            form.AcceptButton = ok;
            form.CancelButton = cancel;
            return form.ShowDialog() == DialogResult.OK ? textBox.Text : string.Empty;
        }
    }
}
