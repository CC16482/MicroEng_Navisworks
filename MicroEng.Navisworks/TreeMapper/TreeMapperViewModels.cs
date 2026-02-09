using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using MicroEng.Navisworks.QuickColour;

namespace MicroEng.Navisworks.TreeMapper
{
    internal sealed class TreeMapperLevelViewModel : NotifyBase
    {
        private readonly Func<string, IEnumerable<string>> _propertyResolver;
        private readonly Action _onChanged;
        private TreeMapperNodeType _nodeType;
        private string _category;
        private string _propertyName;
        private string _missingLabel;
        private TreeMapperSortMode _sortMode;

        public TreeMapperLevelViewModel(
            ObservableCollection<string> categoryOptions,
            Func<string, IEnumerable<string>> propertyResolver,
            Action onChanged)
        {
            CategoryOptions = categoryOptions ?? new ObservableCollection<string>();
            _propertyResolver = propertyResolver ?? (_ => Enumerable.Empty<string>());
            _onChanged = onChanged;
            PropertyOptions = new ObservableCollection<string>();
            _missingLabel = "(Missing)";
            _sortMode = TreeMapperSortMode.Alpha;
            _nodeType = TreeMapperNodeType.Group;

            CategoryOptions.CollectionChanged += (_, __) =>
            {
                RefreshPropertyOptions();
                OnPropertyChanged(nameof(CategoryOptions));
                OnPropertyChanged(nameof(PropertyOptions));
            };
        }

        public ObservableCollection<string> CategoryOptions { get; }
        public ObservableCollection<string> PropertyOptions { get; }

        private static readonly IReadOnlyList<TreeMapperNodeType> NodeTypeOrder = new List<TreeMapperNodeType>
        {
            TreeMapperNodeType.Model,
            TreeMapperNodeType.Layer,
            TreeMapperNodeType.Group,
            TreeMapperNodeType.Composite,
            TreeMapperNodeType.Insert,
            TreeMapperNodeType.Geometry,
            TreeMapperNodeType.Instance,
            TreeMapperNodeType.Collection,
            TreeMapperNodeType.Item
        };

        private static readonly IReadOnlyList<TreeMapperNodeType> NodeTypeOrderWithoutFile = NodeTypeOrder
            .Where(t => t != TreeMapperNodeType.Model)
            .ToList();

        private static readonly IReadOnlyList<TreeMapperNodeType> NodeTypeFileOnly = new List<TreeMapperNodeType>
        {
            TreeMapperNodeType.Model
        };

        public IReadOnlyList<TreeMapperNodeType> NodeTypeOptions => IsFileHeader ? NodeTypeFileOnly : NodeTypeOrderWithoutFile;
        public IReadOnlyList<TreeMapperSortMode> SortModeOptions { get; } = Enum.GetValues(typeof(TreeMapperSortMode)).Cast<TreeMapperSortMode>().ToList();

        public TreeMapperNodeType NodeType
        {
            get => _nodeType;
            set
            {
                if (SetField(ref _nodeType, value))
                {
                    _onChanged?.Invoke();
                }
            }
        }

        private bool _isFileHeader;
        public bool IsFileHeader
        {
            get => _isFileHeader;
            set
            {
                if (SetField(ref _isFileHeader, value))
                {
                    OnPropertyChanged(nameof(NodeTypeOptions));
                    if (_isFileHeader && _nodeType != TreeMapperNodeType.Model)
                    {
                        NodeType = TreeMapperNodeType.Model;
                    }
                }
            }
        }

        public string Category
        {
            get => _category;
            set
            {
                if (SetField(ref _category, value))
                {
                    RefreshPropertyOptions();
                    _onChanged?.Invoke();
                }
            }
        }

        public string PropertyName
        {
            get => _propertyName;
            set
            {
                if (SetField(ref _propertyName, value))
                {
                    _onChanged?.Invoke();
                }
            }
        }

        public string MissingLabel
        {
            get => _missingLabel;
            set
            {
                if (SetField(ref _missingLabel, value))
                {
                    _onChanged?.Invoke();
                }
            }
        }

        public TreeMapperSortMode SortMode
        {
            get => _sortMode;
            set
            {
                if (SetField(ref _sortMode, value))
                {
                    _onChanged?.Invoke();
                }
            }
        }

        public void RefreshPropertyOptions()
        {
            PropertyOptions.Clear();
            foreach (var prop in _propertyResolver(Category) ?? Enumerable.Empty<string>())
            {
                PropertyOptions.Add(prop);
            }

            if (!string.IsNullOrWhiteSpace(PropertyName))
            {
                if (!PropertyOptions.Any(p => string.Equals(p, PropertyName, StringComparison.OrdinalIgnoreCase)))
                {
                    PropertyOptions.Insert(0, PropertyName);
                }
            }
            else
            {
                PropertyName = PropertyOptions.FirstOrDefault();
            }
            OnPropertyChanged(nameof(PropertyOptions));
        }

        public TreeMapperLevel ToModel()
        {
            return new TreeMapperLevel
            {
                NodeType = NodeType,
                Category = Category,
                PropertyName = PropertyName,
                MissingLabel = string.IsNullOrWhiteSpace(MissingLabel) ? "(Missing)" : MissingLabel,
                SortMode = SortMode
            };
        }

        public static TreeMapperLevelViewModel FromModel(
            TreeMapperLevel model,
            ObservableCollection<string> categoryOptions,
            Func<string, IEnumerable<string>> propertyResolver,
            Action onChanged)
        {
            var vm = new TreeMapperLevelViewModel(categoryOptions, propertyResolver, onChanged)
            {
                NodeType = model?.NodeType ?? TreeMapperNodeType.Group,
                Category = model?.Category,
                PropertyName = model?.PropertyName,
                MissingLabel = string.IsNullOrWhiteSpace(model?.MissingLabel) ? "(Missing)" : model?.MissingLabel,
                SortMode = model?.SortMode ?? TreeMapperSortMode.Alpha
            };

            vm.RefreshPropertyOptions();
            if (!string.IsNullOrWhiteSpace(vm.PropertyName) && vm.PropertyOptions.Contains(vm.PropertyName))
            {
                vm.PropertyName = model.PropertyName;
            }

            return vm;
        }
    }
}
