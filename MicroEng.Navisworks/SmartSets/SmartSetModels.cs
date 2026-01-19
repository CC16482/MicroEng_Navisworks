using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;

namespace MicroEng.Navisworks.SmartSets
{
    public enum SmartSetOperator
    {
        Equals,
        NotEquals,
        Contains,
        Wildcard,
        Defined,
        Undefined
    }

    public enum SmartSetOutputType
    {
        SearchSet,
        SelectionSet,
        Both,
        FolderOnly
    }

    public enum SmartSetSearchSetMode
    {
        Single,
        SplitByValue,
        ExpandValuesSingleSet
    }

    public enum SmartSetSearchInMode
    {
        Standard,
        Compact,
        Properties
    }

    public enum SmartSetScopeMode
    {
        AllModel,
        CurrentSelection,
        SavedSelectionSet
    }

    [DataContract]
    public sealed class SmartSetRule : INotifyPropertyChanged
    {
        private string _groupId = "A";
        private string _category = "";
        private string _property = "";
        private SmartSetOperator _operator = SmartSetOperator.Defined;
        private string _value = "";
        private bool _enabled = true;

        [DataMember(Order = 1)]
        public string GroupId
        {
            get => _groupId;
            set
            {
                _groupId = value ?? "A";
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 2)]
        public string Category
        {
            get => _category;
            set
            {
                _category = value ?? "";
                OnPropertyChanged();
                OnPropertyChanged(nameof(Key));
            }
        }

        [DataMember(Order = 3)]
        public string Property
        {
            get => _property;
            set
            {
                _property = value ?? "";
                OnPropertyChanged();
                OnPropertyChanged(nameof(Key));
            }
        }

        [DataMember(Order = 4)]
        public SmartSetOperator Operator
        {
            get => _operator;
            set
            {
                if (_operator == value)
                {
                    return;
                }

                _operator = value;
                if (!IsValueEnabled && !string.IsNullOrWhiteSpace(_value))
                {
                    _value = "";
                    OnPropertyChanged(nameof(Value));
                }

                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValueEnabled));
            }
        }

        [DataMember(Order = 5)]
        public string Value
        {
            get => _value;
            set
            {
                _value = value ?? "";
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 6)]
        public bool Enabled
        {
            get => _enabled;
            set
            {
                _enabled = value;
                OnPropertyChanged();
            }
        }

        [IgnoreDataMember]
        public ObservableCollection<string> SampleValues { get; } = new ObservableCollection<string>();

        [IgnoreDataMember]
        public ObservableCollection<string> PropertyOptions { get; } = new ObservableCollection<string>();

        public string Key => $"{Category}::{Property}";

        [IgnoreDataMember]
        public bool IsValueEnabled => Operator != SmartSetOperator.Defined && Operator != SmartSetOperator.Undefined;

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    [DataContract]
    public sealed class SmartSetGroupingSpec : INotifyPropertyChanged
    {
        private bool _enableSmartGrouping;
        private string _groupByCategory = "";
        private string _groupByProperty = "";
        private bool _useThenBy;
        private string _thenByCategory = "";
        private string _thenByProperty = "";
        private int _maxGroups = 50;
        private int _minCount = 5;
        private bool _includeBlanks;

        [DataMember(Order = 1)]
        public bool EnableSmartGrouping
        {
            get => _enableSmartGrouping;
            set
            {
                _enableSmartGrouping = value;
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 2)]
        public string GroupByCategory
        {
            get => _groupByCategory;
            set
            {
                _groupByCategory = value ?? "";
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 3)]
        public string GroupByProperty
        {
            get => _groupByProperty;
            set
            {
                _groupByProperty = value ?? "";
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 4)]
        public bool UseThenBy
        {
            get => _useThenBy;
            set
            {
                _useThenBy = value;
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 5)]
        public string ThenByCategory
        {
            get => _thenByCategory;
            set
            {
                _thenByCategory = value ?? "";
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 6)]
        public string ThenByProperty
        {
            get => _thenByProperty;
            set
            {
                _thenByProperty = value ?? "";
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 7)]
        public int MaxGroups
        {
            get => _maxGroups;
            set
            {
                _maxGroups = Math.Max(1, value);
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 8)]
        public int MinCount
        {
            get => _minCount;
            set
            {
                _minCount = Math.Max(1, value);
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 9)]
        public bool IncludeBlanks
        {
            get => _includeBlanks;
            set
            {
                _includeBlanks = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    [DataContract]
    public sealed class SmartSetRecipe : INotifyPropertyChanged
    {
        private string _name = "New Recipe";
        private string _description = "";
        private string _dataScraperProfile = "";
        private SmartSetOutputType _outputType = SmartSetOutputType.SearchSet;
        private string _folderPath = "MicroEng/Smart Sets";
        private SmartSetSearchSetMode _searchSetMode = SmartSetSearchSetMode.Single;
        private bool _generateMultipleSearchSets;
        private SmartSetSearchInMode _searchInMode = SmartSetSearchInMode.Standard;
        private SmartSetScopeMode _scopeMode = SmartSetScopeMode.AllModel;
        private string _scopeSelectionSetName = "";
        private string _scopeSummary = "Entire model";

        [DataMember(Order = 1)]
        public int Version { get; set; } = 1;

        [DataMember(Order = 2)]
        public string Name
        {
            get => _name;
            set
            {
                _name = value ?? "";
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 3)]
        public string Description
        {
            get => _description;
            set
            {
                _description = value ?? "";
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 4)]
        public string DataScraperProfile
        {
            get => _dataScraperProfile;
            set
            {
                _dataScraperProfile = value ?? "";
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 5)]
        public SmartSetOutputType OutputType
        {
            get => _outputType;
            set
            {
                _outputType = value;
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 6)]
        public string FolderPath
        {
            get => _folderPath;
            set
            {
                _folderPath = value ?? "";
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 7)]
        public SmartSetSearchSetMode SearchSetMode
        {
            get => _searchSetMode;
            set
            {
                _searchSetMode = value;
                _generateMultipleSearchSets = value == SmartSetSearchSetMode.SplitByValue;
                OnPropertyChanged();
                OnPropertyChanged(nameof(GenerateMultipleSearchSets));
            }
        }

        [DataMember(Order = 8)]
        public bool GenerateMultipleSearchSets
        {
            get => _generateMultipleSearchSets;
            set
            {
                _generateMultipleSearchSets = value;
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 9)]
        public ObservableCollection<SmartSetRule> Rules { get; set; } = new ObservableCollection<SmartSetRule>();

        [DataMember(Order = 10)]
        public SmartSetGroupingSpec Grouping { get; set; } = new SmartSetGroupingSpec();

        [DataMember(Order = 11)]
        public DateTime CreatedUtc { get; set; } = DateTime.UtcNow;

        [DataMember(Order = 12)]
        public DateTime UpdatedUtc { get; set; } = DateTime.UtcNow;

        [DataMember(Order = 13)]
        public SmartSetSearchInMode SearchInMode
        {
            get => _searchInMode;
            set
            {
                _searchInMode = value;
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 14)]
        public SmartSetScopeMode ScopeMode
        {
            get => _scopeMode;
            set
            {
                _scopeMode = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsScopeConstrained));
            }
        }

        [DataMember(Order = 15)]
        public string ScopeSelectionSetName
        {
            get => _scopeSelectionSetName;
            set
            {
                _scopeSelectionSetName = value ?? "";
                OnPropertyChanged();
            }
        }

        [DataMember(Order = 16)]
        public string ScopeSummary
        {
            get => _scopeSummary;
            set
            {
                _scopeSummary = value ?? "";
                OnPropertyChanged();
            }
        }

        [IgnoreDataMember]
        public bool IsScopeConstrained => ScopeMode != SmartSetScopeMode.AllModel;

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        [OnDeserialized]
        private void OnDeserialized(StreamingContext context)
        {
            if (_searchSetMode == SmartSetSearchSetMode.Single && _generateMultipleSearchSets)
            {
                _searchSetMode = SmartSetSearchSetMode.SplitByValue;
            }

            if (string.IsNullOrWhiteSpace(_scopeSummary))
            {
                _scopeSummary = _scopeMode == SmartSetScopeMode.AllModel ? "Entire model" : "Scope active";
            }

            _scopeSelectionSetName ??= "";
        }
    }

    public sealed class SmartGroupRow
    {
        public string Value1 { get; set; }
        public string Value2 { get; set; }
        public int Count { get; set; }

        public string DisplayKey => string.IsNullOrWhiteSpace(Value2) ? Value1 : $"{Value1} / {Value2}";
    }

    public sealed class ScrapedPropertyDescriptor
    {
        public string Category { get; }
        public string Name { get; }
        public string Type { get; }
        public int ItemCount { get; }
        public int DistinctCount { get; }
        public IReadOnlyList<string> SampleValues { get; }

        public ScrapedPropertyDescriptor(
            string category,
            string name,
            string type,
            int itemCount,
            int distinctCount,
            IReadOnlyList<string> sampleValues)
        {
            Category = category ?? "";
            Name = name ?? "";
            Type = type ?? "";
            ItemCount = itemCount;
            DistinctCount = distinctCount;
            SampleValues = sampleValues ?? Array.Empty<string>();
        }

        public string Key => $"{Category}::{Name}";

        public string SampleValuesPreview
            => SampleValues == null || SampleValues.Count == 0
                ? ""
                : string.Join(", ", SampleValues);
    }

    public sealed class FastPreviewResult
    {
        public int EstimatedMatchCount { get; set; }
        public List<string> SampleItemPaths { get; set; } = new List<string>();
        public bool UsedCache { get; set; } = true;
        public string Notes { get; set; } = "";
        public string SessionLabel { get; set; } = "";
    }
}
