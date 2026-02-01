using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace MicroEng.Navisworks
{
    internal partial class DataMatrixColumnBuilderWindow : Window
    {
        private readonly ColumnBuilderViewModel _viewModel;

        public DataMatrixColumnBuilderWindow(IEnumerable<DataMatrixAttributeDefinition> attributes, IList<string> currentVisibleOrder)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            _viewModel = new ColumnBuilderViewModel(attributes, currentVisibleOrder);
            DataContext = _viewModel;
        }

        public IReadOnlyList<string> VisibleAttributeIds => _viewModel.VisibleAttributeIds;

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void CategoryCheckBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is System.Windows.Controls.CheckBox checkbox
                && checkbox.DataContext is ColumnBuilderCategoryNode node)
            {
                node.ApplyUserToggle();
                e.Handled = true;
            }
        }

        private void CategoryCheckBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Space && e.Key != Key.Enter)
            {
                return;
            }

            if (sender is System.Windows.Controls.CheckBox checkbox
                && checkbox.DataContext is ColumnBuilderCategoryNode node)
            {
                node.ApplyUserToggle();
                e.Handled = true;
            }
        }
    }

    internal sealed class ColumnBuilderViewModel : INotifyPropertyChanged
    {
        private readonly List<ColumnBuilderPropertyNode> _allProperties = new();
        private ObservableCollection<ColumnBuilderPropertyNode> _selectedProperties;
        private bool _suppressSelectionChanged;
        private bool _pendingRebuild;
        private string _filterText;
        private string _selectedCountText = "0 selected";

        public ColumnBuilderViewModel(IEnumerable<DataMatrixAttributeDefinition> attributes, IList<string> currentVisibleOrder)
        {
            Categories = new ObservableCollection<ColumnBuilderCategoryNode>();
            SelectedProperties = new ObservableCollection<ColumnBuilderPropertyNode>();

            var orderMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (currentVisibleOrder != null)
            {
                for (var i = 0; i < currentVisibleOrder.Count; i++)
                {
                    var id = currentVisibleOrder[i];
                    if (string.IsNullOrWhiteSpace(id)) continue;
                    if (!orderMap.ContainsKey(id))
                    {
                        orderMap[id] = i;
                    }
                }
            }

            var available = attributes?.ToList() ?? new List<DataMatrixAttributeDefinition>();
            var byCategory = available
                .GroupBy(a => a.Category ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var group in byCategory)
            {
                var category = new ColumnBuilderCategoryNode(group.Key);
                category.SelectionChanged += OnSelectionChanged;

                foreach (var attr in group.OrderBy(a => a.PropertyName, StringComparer.OrdinalIgnoreCase))
                {
                    var id = attr.Id ?? $"{attr.Category}|{attr.PropertyName}";
                    var propName = string.IsNullOrWhiteSpace(attr.DisplayName) ? attr.PropertyName : attr.DisplayName;
                    if (string.IsNullOrWhiteSpace(propName)) propName = id;

                    var node = new ColumnBuilderPropertyNode(id, propName, group.Key)
                    {
                        SortIndex = orderMap.TryGetValue(id, out var index) ? index : int.MaxValue,
                        IsChecked = orderMap.ContainsKey(id)
                    };
                    node.SelectionChanged += OnSelectionChanged;
                    node.Parent = category;
                    category.Properties.Add(node);
                    _allProperties.Add(node);
                }

                category.RefreshFromChildren();
                Categories.Add(category);
            }

            RebuildSelected();
            ApplyFilter();
        }

        public ObservableCollection<ColumnBuilderCategoryNode> Categories { get; }

        public ObservableCollection<ColumnBuilderPropertyNode> SelectedProperties
        {
            get => _selectedProperties;
            private set
            {
                if (_selectedProperties == value) return;
                _selectedProperties = value;
                OnPropertyChanged();
            }
        }

        public IReadOnlyList<string> VisibleAttributeIds => SelectedProperties.Select(p => p.Id).ToList();

        public string SelectedCountText
        {
            get => _selectedCountText;
            private set
            {
                if (_selectedCountText == value) return;
                _selectedCountText = value;
                OnPropertyChanged();
            }
        }

        public string FilterText
        {
            get => _filterText;
            set
            {
                if (_filterText == value) return;
                _filterText = value;
                OnPropertyChanged();
                ApplyFilter();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnSelectionChanged()
        {
            if (_suppressSelectionChanged) return;
            ScheduleRebuildSelected();
        }

        private void ScheduleRebuildSelected()
        {
            if (_pendingRebuild) return;
            _pendingRebuild = true;

            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null)
            {
                _pendingRebuild = false;
                RebuildSelected();
                return;
            }

            dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
            {
                _pendingRebuild = false;
                RebuildSelected();
            }));
        }

        private void RebuildSelected()
        {
            _suppressSelectionChanged = true;
            try
            {
                var ordered = _allProperties
                    .Where(p => p.IsChecked == true)
                    .OrderBy(p => p.SortIndex)
                    .ThenBy(p => p.CategoryName, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(p => p.PropertyName, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                SelectedProperties = new ObservableCollection<ColumnBuilderPropertyNode>(ordered);
                SelectedCountText = $"{SelectedProperties.Count} selected";
            }
            finally
            {
                _suppressSelectionChanged = false;
            }
        }

        private void ApplyFilter()
        {
            var text = _filterText?.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                foreach (var category in Categories)
                {
                    category.IsVisible = true;
                    foreach (var prop in category.Properties)
                    {
                        prop.IsVisible = true;
                    }
                }
                return;
            }

            foreach (var category in Categories)
            {
                var categoryMatch = category.Name.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0;
                var anyVisible = false;
                foreach (var prop in category.Properties)
                {
                    var match = categoryMatch
                                || prop.PropertyName.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0
                                || prop.Id.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0;
                    prop.IsVisible = match;
                    if (match)
                    {
                        anyVisible = true;
                    }
                }

                category.IsVisible = categoryMatch || anyVisible;
                if (category.IsVisible)
                {
                    category.IsExpanded = true;
                }
            }
        }

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

    }

    public sealed class ColumnBuilderCategoryNode : INotifyPropertyChanged
    {
        private bool? _isChecked = false;
        private bool _suppressChildren;
        private bool _isVisible = true;

        public ColumnBuilderCategoryNode(string name)
        {
            Name = string.IsNullOrWhiteSpace(name) ? "Uncategorized" : name;
            Properties = new ObservableCollection<ColumnBuilderPropertyNode>();
            IsExpanded = true;
        }

        public string Name { get; }

        public ObservableCollection<ColumnBuilderPropertyNode> Properties { get; }

        public bool IsExpanded { get; set; }

        public bool IsVisible
        {
            get => _isVisible;
            set => SetField(ref _isVisible, value);
        }

        public bool? IsChecked
        {
            get => _isChecked;
            set
            {
                if (SetField(ref _isChecked, value))
                {
                    if (_suppressChildren) return;
                    if (value == null)
                    {
                        SelectionChanged?.Invoke();
                        return;
                    }

                    _suppressChildren = true;
                    foreach (var child in Properties)
                    {
                        child.SetCheckedFromParent(value == true);
                    }
                    _suppressChildren = false;
                    SelectionChanged?.Invoke();
                }
            }
        }

        public event Action SelectionChanged;
        public event PropertyChangedEventHandler PropertyChanged;

        public void RefreshFromChildren()
        {
            if (_suppressChildren) return;

            var allTrue = Properties.All(p => p.IsChecked == true);
            var allFalse = Properties.All(p => p.IsChecked == false);
            var newValue = allTrue ? (bool?)true : allFalse ? (bool?)false : null;

            _suppressChildren = true;
            SetField(ref _isChecked, newValue);
            _suppressChildren = false;
        }

        public void ApplyUserToggle()
        {
            var target = _isChecked != true;
            SetCheckedFromUser(target);
        }

        private void SetCheckedFromUser(bool isChecked)
        {
            if (_suppressChildren) return;

            _suppressChildren = true;
            SetField(ref _isChecked, isChecked);
            _suppressChildren = false;

            foreach (var child in Properties)
            {
                child.SetCheckedFromParent(isChecked);
            }

            SelectionChanged?.Invoke();
        }

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string name = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
            {
                return false;
            }

            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
            return true;
        }
    }

    public sealed class ColumnBuilderPropertyNode : INotifyPropertyChanged
    {
        private bool? _isChecked;
        private bool _isVisible = true;

        public ColumnBuilderPropertyNode(string id, string propertyName, string categoryName)
        {
            Id = id;
            PropertyName = propertyName;
            CategoryName = categoryName;
        }

        public string Id { get; }
        public string PropertyName { get; }
        public string CategoryName { get; }
        public int SortIndex { get; set; }
        public ColumnBuilderCategoryNode Parent { get; set; }

        public bool IsVisible
        {
            get => _isVisible;
            set => SetField(ref _isVisible, value);
        }

        public bool? IsChecked
        {
            get => _isChecked;
            set
            {
                var normalized = value == true;
                if (SetField(ref _isChecked, normalized))
                {
                    Parent?.RefreshFromChildren();
                    SelectionChanged?.Invoke();
                }
            }
        }

        public event Action SelectionChanged;
        public event PropertyChangedEventHandler PropertyChanged;

        public void SetCheckedFromParent(bool isChecked)
        {
            if (SetField(ref _isChecked, isChecked))
            {
                SelectionChanged?.Invoke();
            }
        }

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string name = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
            {
                return false;
            }

            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
            return true;
        }
    }
}
