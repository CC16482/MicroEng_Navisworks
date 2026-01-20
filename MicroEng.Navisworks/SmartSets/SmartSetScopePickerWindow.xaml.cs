using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using Autodesk.Navisworks.Api;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.SmartSets
{
    public partial class SmartSetScopePickerWindow : Window
    {
        public SmartSetScopePickerViewModel VM { get; }
        public SmartSetScopePickerResult Result { get; private set; }

        public SmartSetScopePickerWindow(Document doc, IEnumerable<ScrapedPropertyDescriptor> properties, SmartSetRecipe recipe)
        {
            InitializeComponent();
            VM = new SmartSetScopePickerViewModel(doc, properties, recipe);
            DataContext = VM;

            try
            {
                MicroEngWpfUiTheme.ApplyTo(this);
            }
            catch
            {
                // ignore theme failures
            }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!VM.TryBuildResult(out var result))
            {
                return;
            }

            Result = result;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }

    public sealed class SmartSetScopePickerResult
    {
        public SmartSetSearchInMode SearchInMode { get; set; }
        public SmartSetScopeMode ScopeMode { get; set; }
        public List<List<string>> ModelPaths { get; set; } = new List<List<string>>();
        public string FilterCategory { get; set; } = "";
        public string FilterProperty { get; set; } = "";
        public string FilterValue { get; set; } = "";
    }

    public sealed class SmartSetScopePickerViewModel : INotifyPropertyChanged
    {
        private readonly Document _doc;
        private readonly List<ScrapedPropertyDescriptor> _properties;
        private SmartSetSearchInMode _selectedSearchInMode;
        private string _selectedCategory = "";
        private string _selectedProperty = "";
        private string _selectedValue = "";

        public ObservableCollection<SearchInModeOption> SearchInModeOptions { get; } = new ObservableCollection<SearchInModeOption>();
        public ObservableCollection<SmartSetScopeTreeNode> RootNodes { get; } = new ObservableCollection<SmartSetScopeTreeNode>();
        public ObservableCollection<string> Categories { get; } = new ObservableCollection<string>();
        public ObservableCollection<string> Properties { get; } = new ObservableCollection<string>();
        public ObservableCollection<string> Values { get; } = new ObservableCollection<string>();

        public SmartSetScopePickerViewModel(
            Document doc,
            IEnumerable<ScrapedPropertyDescriptor> properties,
            SmartSetRecipe recipe)
        {
            _doc = doc;
            _properties = (properties ?? Enumerable.Empty<ScrapedPropertyDescriptor>()).ToList();

            SearchInModeOptions.Add(new SearchInModeOption(SmartSetSearchInMode.Standard, "Standard"));
            SearchInModeOptions.Add(new SearchInModeOption(SmartSetSearchInMode.Compact, "Compact"));
            SearchInModeOptions.Add(new SearchInModeOption(SmartSetSearchInMode.Properties, "Properties"));

            BuildCategoryList();
            BuildRootNodes();

            var initialSearchIn = recipe?.SearchInMode ?? SmartSetSearchInMode.Standard;
            SelectedSearchInMode = initialSearchIn;

            if (recipe?.ScopeMode == SmartSetScopeMode.PropertyFilter && recipe.HasScopePropertyFilter)
            {
                SetPropertyFilterSelection(recipe.ScopeFilterCategory, recipe.ScopeFilterProperty, recipe.ScopeFilterValue);
            }
            else
            {
                ApplySearchInDefaults();
            }
        }

        public SmartSetSearchInMode SelectedSearchInMode
        {
            get => _selectedSearchInMode;
            set
            {
                if (_selectedSearchInMode == value)
                {
                    return;
                }

                _selectedSearchInMode = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsTreeMode));
                OnPropertyChanged(nameof(IsPropertyFilterMode));
                OnPropertyChanged(nameof(StatusText));
                ApplySearchInDefaults();
                OnPropertyChanged(nameof(CanAccept));
            }
        }

        public bool IsTreeMode => SelectedSearchInMode == SmartSetSearchInMode.Standard;
        public bool IsPropertyFilterMode => SelectedSearchInMode != SmartSetSearchInMode.Standard;
        public bool IsPropertyFilterAvailable => _properties.Count > 0;
        public string PropertyFilterStatusText => "Run Data Scraper to use Compact/Properties scope.";

        public string StatusText
        {
            get
            {
                if (IsPropertyFilterMode && !IsPropertyFilterAvailable)
                {
                    return "Run Data Scraper to use Compact/Properties scope.";
                }

                return "";
            }
        }

        public string SelectedCategory
        {
            get => _selectedCategory;
            set
            {
                var next = value ?? "";
                if (string.Equals(_selectedCategory, next, StringComparison.Ordinal))
                {
                    return;
                }

                _selectedCategory = next;
                EnsureCategoryExists(next);
                OnPropertyChanged();
                RefreshProperties();
                OnPropertyChanged(nameof(CanAccept));
            }
        }

        public string SelectedProperty
        {
            get => _selectedProperty;
            set
            {
                var next = value ?? "";
                if (string.Equals(_selectedProperty, next, StringComparison.Ordinal))
                {
                    return;
                }

                _selectedProperty = next;
                EnsurePropertyExists(next);
                OnPropertyChanged();
                RefreshValues();
                OnPropertyChanged(nameof(CanAccept));
            }
        }

        public string SelectedValue
        {
            get => _selectedValue;
            set
            {
                var next = value ?? "";
                if (string.Equals(_selectedValue, next, StringComparison.Ordinal))
                {
                    return;
                }

                _selectedValue = next;
                EnsureValueExists(next);
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanAccept));
            }
        }

        public bool CanAccept
        {
            get
            {
                if (IsTreeMode)
                {
                    return GetCheckedPaths().Count > 0;
                }

                if (!IsPropertyFilterAvailable)
                {
                    return false;
                }

                return !string.IsNullOrWhiteSpace(SelectedCategory)
                    && !string.IsNullOrWhiteSpace(SelectedProperty)
                    && !string.IsNullOrWhiteSpace(SelectedValue);
            }
        }

        public bool TryBuildResult(out SmartSetScopePickerResult result)
        {
            result = null;

            if (IsTreeMode)
            {
                var paths = GetCheckedPaths();
                if (paths.Count == 0)
                {
                    return false;
                }

                result = new SmartSetScopePickerResult
                {
                    ScopeMode = SmartSetScopeMode.ModelTree,
                    SearchInMode = SmartSetSearchInMode.Standard,
                    ModelPaths = paths
                };
                return true;
            }

            if (!IsPropertyFilterAvailable)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(SelectedCategory)
                || string.IsNullOrWhiteSpace(SelectedProperty)
                || string.IsNullOrWhiteSpace(SelectedValue))
            {
                return false;
            }

            result = new SmartSetScopePickerResult
            {
                ScopeMode = SmartSetScopeMode.PropertyFilter,
                SearchInMode = SelectedSearchInMode,
                FilterCategory = SelectedCategory,
                FilterProperty = SelectedProperty,
                FilterValue = SelectedValue
            };
            return true;
        }

        private void BuildCategoryList()
        {
            Categories.Clear();
            foreach (var cat in _properties.Select(p => p.Category ?? "")
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(c => c, StringComparer.OrdinalIgnoreCase))
            {
                Categories.Add(cat);
            }
        }

        private void BuildRootNodes()
        {
            RootNodes.Clear();

            var roots = _doc?.Models?.RootItems;
            if (roots == null)
            {
                return;
            }

            foreach (var child in roots)
            {
                if (child != null)
                {
                    RootNodes.Add(new SmartSetScopeTreeNode(child, null, OnTreeSelectionChanged));
                }
            }
        }

        private void OnTreeSelectionChanged()
        {
            OnPropertyChanged(nameof(CanAccept));
        }

        private void ApplySearchInDefaults()
        {
            if (!IsPropertyFilterAvailable || !IsPropertyFilterMode)
            {
                return;
            }

            if (SelectedSearchInMode == SmartSetSearchInMode.Compact)
            {
                if (string.IsNullOrWhiteSpace(SelectedCategory))
                {
                    SelectedCategory = Categories.FirstOrDefault(c => string.Equals(c, "Element", StringComparison.OrdinalIgnoreCase))
                        ?? Categories.FirstOrDefault();
                }

                if (string.IsNullOrWhiteSpace(SelectedProperty))
                {
                    SelectedProperty = Properties.FirstOrDefault(p => string.Equals(p, "Category", StringComparison.OrdinalIgnoreCase))
                        ?? Properties.FirstOrDefault();
                }
            }
            else if (SelectedSearchInMode == SmartSetSearchInMode.Properties)
            {
                if (string.IsNullOrWhiteSpace(SelectedCategory))
                {
                    SelectedCategory = Categories.FirstOrDefault();
                }

                if (string.IsNullOrWhiteSpace(SelectedProperty))
                {
                    SelectedProperty = Properties.FirstOrDefault();
                }
            }
        }

        private void SetPropertyFilterSelection(string category, string property, string value)
        {
            if (!string.IsNullOrWhiteSpace(category))
            {
                SelectedCategory = category;
            }

            if (!string.IsNullOrWhiteSpace(property))
            {
                SelectedProperty = property;
            }

            if (!string.IsNullOrWhiteSpace(value))
            {
                SelectedValue = value;
            }
        }

        private void RefreshProperties()
        {
            Properties.Clear();

            if (string.IsNullOrWhiteSpace(SelectedCategory))
            {
                return;
            }

            foreach (var name in _properties
                .Where(p => string.Equals(p.Category, SelectedCategory, StringComparison.OrdinalIgnoreCase))
                .Select(p => p.Name ?? "")
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase))
            {
                Properties.Add(name);
            }

            if (!Properties.Any(p => string.Equals(p, SelectedProperty, StringComparison.OrdinalIgnoreCase)))
            {
                SelectedProperty = Properties.FirstOrDefault() ?? "";
            }
            else
            {
                RefreshValues();
            }
        }

        private void RefreshValues()
        {
            Values.Clear();

            if (string.IsNullOrWhiteSpace(SelectedCategory) || string.IsNullOrWhiteSpace(SelectedProperty))
            {
                return;
            }

            var samples = _properties
                .Where(p => string.Equals(p.Category, SelectedCategory, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(p.Name, SelectedProperty, StringComparison.OrdinalIgnoreCase))
                .SelectMany(p => p.SampleValues ?? Array.Empty<string>())
                .Select(v => v ?? "")
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var value in samples)
            {
                Values.Add(value);
            }

            if (!Values.Any(v => string.Equals(v, SelectedValue, StringComparison.OrdinalIgnoreCase)))
            {
                SelectedValue = Values.FirstOrDefault() ?? "";
            }
        }

        private void EnsureCategoryExists(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            if (!Categories.Any(c => string.Equals(c, value, StringComparison.OrdinalIgnoreCase)))
            {
                Categories.Add(value);
            }
        }

        private void EnsurePropertyExists(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            if (!Properties.Any(p => string.Equals(p, value, StringComparison.OrdinalIgnoreCase)))
            {
                Properties.Add(value);
            }
        }

        private void EnsureValueExists(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            if (!Values.Any(v => string.Equals(v, value, StringComparison.OrdinalIgnoreCase)))
            {
                Values.Add(value);
            }
        }

        private List<List<string>> GetCheckedPaths()
        {
            var results = new List<List<string>>();
            foreach (var node in RootNodes)
            {
                CollectCheckedPaths(node, results);
            }

            return results;
        }

        private static void CollectCheckedPaths(SmartSetScopeTreeNode node, List<List<string>> results)
        {
            if (node == null || node.IsPlaceholder)
            {
                return;
            }

            if (node.IsChecked == true)
            {
                results.Add(node.GetPathSegments().ToList());
            }

            foreach (var child in node.Children)
            {
                CollectCheckedPaths(child, results);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public sealed class SmartSetScopeTreeNode : INotifyPropertyChanged
    {
        private readonly ModelItem _item;
        private readonly Action _selectionChanged;
        private bool _childrenLoaded;
        private bool _isExpanded;
        private bool? _isChecked = false;

        public SmartSetScopeTreeNode(ModelItem item, SmartSetScopeTreeNode parent, Action selectionChanged)
        {
            _item = item;
            Parent = parent;
            _selectionChanged = selectionChanged;
            DisplayName = BuildLabel(item);
            Children = new ObservableCollection<SmartSetScopeTreeNode>();

            if (ItemHasChildren(item))
            {
                Children.Add(new SmartSetScopeTreeNode());
            }
        }

        private SmartSetScopeTreeNode()
        {
            IsPlaceholder = true;
            DisplayName = "";
            Children = new ObservableCollection<SmartSetScopeTreeNode>();
        }

        public SmartSetScopeTreeNode Parent { get; }
        public string DisplayName { get; }
        public ObservableCollection<SmartSetScopeTreeNode> Children { get; }
        public bool IsPlaceholder { get; }

        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                if (_isExpanded == value)
                {
                    return;
                }

                _isExpanded = value;
                OnPropertyChanged();

                if (_isExpanded)
                {
                    EnsureChildrenLoaded();
                }
            }
        }

        public bool? IsChecked
        {
            get => _isChecked;
            set
            {
                var next = value ?? false;
                if (_isChecked == next)
                {
                    return;
                }

                // Keep user toggles two-state; indeterminate is only set by parent aggregation.
                SetIsChecked(next, updateChildren: true, updateParent: true);
            }
        }

        public IReadOnlyList<string> GetPathSegments()
        {
            var stack = new Stack<string>();
            var current = this;
            while (current != null)
            {
                if (!string.IsNullOrWhiteSpace(current.DisplayName))
                {
                    stack.Push(current.DisplayName);
                }

                current = current.Parent;
            }

            return stack.ToList();
        }

        private void EnsureChildrenLoaded()
        {
            if (_childrenLoaded || _item == null)
            {
                return;
            }

            _childrenLoaded = true;
            Children.Clear();

            try
            {
                var children = _item.Children?.Cast<ModelItem>()
                    .Where(c => c != null)
                    .OrderBy(c => BuildLabel(c), StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (children == null)
                {
                    return;
                }

                foreach (var child in children)
                {
                    Children.Add(new SmartSetScopeTreeNode(child, this, _selectionChanged));
                }

                if (_isChecked.HasValue)
                {
                    foreach (var child in Children)
                    {
                        if (!child.IsPlaceholder)
                        {
                            child.SetIsChecked(_isChecked, updateChildren: true, updateParent: false);
                        }
                    }
                }
            }
            catch
            {
                // ignore child loading failures
            }
        }

        private static bool ItemHasChildren(ModelItem item)
        {
            try
            {
                if (item?.Children == null)
                {
                    return false;
                }

                var enumerator = item.Children.GetEnumerator();
                return enumerator.MoveNext();
            }
            catch
            {
                return false;
            }
        }

        private static string BuildLabel(ModelItem item)
        {
            if (item == null)
            {
                return "";
            }

            var label = item.DisplayName ?? "";
            if (!string.IsNullOrWhiteSpace(label))
            {
                return label;
            }

            label = item.ClassDisplayName ?? "";
            if (!string.IsNullOrWhiteSpace(label))
            {
                return label;
            }

            label = item.ClassName ?? "";
            if (!string.IsNullOrWhiteSpace(label))
            {
                return label;
            }

            return item.InstanceGuid.ToString();
        }

        private void SetIsChecked(bool? value, bool updateChildren, bool updateParent)
        {
            if (_isChecked == value)
            {
                return;
            }

            _isChecked = value;
            OnPropertyChanged(nameof(IsChecked));
            _selectionChanged?.Invoke();

            if (updateChildren && value.HasValue)
            {
                foreach (var child in Children)
                {
                    if (!child.IsPlaceholder)
                    {
                        child.SetIsChecked(value, updateChildren: true, updateParent: false);
                    }
                }
            }

            if (updateParent && Parent != null)
            {
                Parent.UpdateParentCheckState();
            }
        }

        private void UpdateParentCheckState()
        {
            var children = Children.Where(c => !c.IsPlaceholder).ToList();
            if (children.Count == 0)
            {
                return;
            }

            var allChecked = children.All(c => c.IsChecked == true);
            var allUnchecked = children.All(c => c.IsChecked == false);
            bool? newValue = allChecked ? true : allUnchecked ? false : null;

            if (_isChecked == newValue)
            {
                return;
            }

            _isChecked = newValue;
            OnPropertyChanged(nameof(IsChecked));
            _selectionChanged?.Invoke();

            Parent?.UpdateParentCheckState();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
