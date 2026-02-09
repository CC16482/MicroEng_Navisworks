
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using Autodesk.Navisworks.Api;
using MicroEng.Navisworks.Colour;
using MicroEng.Navisworks;
using MicroEng.Navisworks.QuickColour.Profiles;
using MicroEng.Navisworks.SmartSets;
using Microsoft.Win32;
using NavisApp = Autodesk.Navisworks.Api.Application;
using WpfUiControls = Wpf.Ui.Controls;

namespace MicroEng.Navisworks.QuickColour
{
    public partial class QuickColourControl : UserControl, INotifyPropertyChanged
    {
        public sealed class ScraperProfileOption
        {
            internal ScrapeSession Session { get; }
            public string Label { get; }

            internal ScraperProfileOption(ScrapeSession session)
            {
                Session = session;
                var profile = string.IsNullOrWhiteSpace(session?.ProfileName) ? "Unknown" : session.ProfileName;
                var stamp = session?.Timestamp == default ? "" : session.Timestamp.ToString("yyyy-MM-dd HH:mm");
                Label = string.IsNullOrWhiteSpace(stamp) ? profile : $"{profile} @ {stamp}";
            }
        }

        public sealed class ScopeOption
        {
            public QuickColourScope Value { get; }
            public string Label { get; }

            public ScopeOption(QuickColourScope value, string label)
            {
                Value = value;
                Label = label ?? value.ToString();
            }
        }

        public sealed class PaletteOption
        {
            public MicroEngPaletteKind Value { get; }
            public string Label { get; }

            public PaletteOption(MicroEngPaletteKind value, string label)
            {
                Value = value;
                Label = label ?? value.ToString();
            }
        }

        private const string DefaultCategoryDisciplineMapJson = @"{
  ""version"": 1,
  ""fallbackGroup"": ""Other"",
  ""rules"": [
    { ""group"": ""Architecture"", ""match"": [
      { ""type"": ""exact"", ""value"": ""Doors"" },
      { ""type"": ""exact"", ""value"": ""Windows"" },
      { ""type"": ""contains"", ""value"": ""Wall"" },
      { ""type"": ""exact"", ""value"": ""Floors"" },
      { ""type"": ""exact"", ""value"": ""Ceilings"" },
      { ""type"": ""exact"", ""value"": ""Rooms"" },
      { ""type"": ""wildcard"", ""value"": ""OST_*Door*"" },
      { ""type"": ""wildcard"", ""value"": ""OST_*Window*"" }
    ]},
    { ""group"": ""Mechanical"", ""match"": [
      { ""type"": ""contains"", ""value"": ""Duct"" },
      { ""type"": ""contains"", ""value"": ""Pipe"" },
      { ""type"": ""contains"", ""value"": ""Plumbing"" },
      { ""type"": ""contains"", ""value"": ""Mechanical Equipment"" }
    ]},
    { ""group"": ""Electrical"", ""match"": [
      { ""type"": ""contains"", ""value"": ""Conduit"" },
      { ""type"": ""contains"", ""value"": ""Cable Tray"" },
      { ""type"": ""contains"", ""value"": ""Lighting"" },
      { ""type"": ""contains"", ""value"": ""Electrical"" },
      { ""type"": ""contains"", ""value"": ""Data"" }
    ]},
    { ""group"": ""Fire"", ""match"": [
      { ""type"": ""contains"", ""value"": ""Fire"" },
      { ""type"": ""contains"", ""value"": ""Sprinkler"" }
    ]}
  ]
}";

        private readonly QuickColourNavisworksService _service = new QuickColourNavisworksService();
        private readonly QuickColourHueGroupAutoAssignService _autoAssign = new QuickColourHueGroupAutoAssignService();
        private readonly MicroEngColourProfileStore _colourProfileStore = MicroEngColourProfileStore.CreateDefault();

        private DisciplineMapFile _disciplineMap;
        private List<HueGroupAutoAssignPreviewRow> _lastPreviewRows = new List<HueGroupAutoAssignPreviewRow>();
        private readonly Dictionary<string, List<string>> _hierarchyPropertyOptionsByCategory =
            new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<string>> _quickColourPropertyOptionsByCategory =
            new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        private bool _isApplyingHierarchyColours;
        private bool _skipHierarchyRecolor;
        private string _lastQuickColourProfileId;
        private string _lastHierarchyProfileId;

        public ObservableCollection<ScraperProfileOption> ScraperProfiles { get; } =
            new ObservableCollection<ScraperProfileOption>();

        public ObservableCollection<PaletteOption> QuickColourPaletteOptions { get; } =
            new ObservableCollection<PaletteOption>();

        public ObservableCollection<string> QuickColourCategoryOptions { get; } =
            new ObservableCollection<string>();

        public ObservableCollection<string> QuickColourPropertyOptions { get; } =
            new ObservableCollection<string>();

        public ObservableCollection<QuickColourValueRow> QuickColourValues { get; } =
            new ObservableCollection<QuickColourValueRow>();

        public ObservableCollection<string> HierarchyCategoryOptions { get; } =
            new ObservableCollection<string>();

        public ObservableCollection<string> HierarchyL1PropertyOptions { get; } =
            new ObservableCollection<string>();

        public ObservableCollection<string> HierarchyL2PropertyOptions { get; } =
            new ObservableCollection<string>();

        public ObservableCollection<QuickColourHierarchyGroup> HierarchyGroups { get; } =
            new ObservableCollection<QuickColourHierarchyGroup>();

        public ObservableCollection<QuickColourHueGroup> HueGroups { get; } =
            new ObservableCollection<QuickColourHueGroup>();

        public ObservableCollection<string> HueGroupOptions { get; } =
            new ObservableCollection<string>();

        public ObservableCollection<QuickColourPaletteStyle> PaletteStyleOptions { get; } =
            new ObservableCollection<QuickColourPaletteStyle>();

        public ObservableCollection<QuickColourTypeSortMode> TypeSortOptions { get; } =
            new ObservableCollection<QuickColourTypeSortMode>();

        public ObservableCollection<ScopeOption> ScopeOptions { get; } =
            new ObservableCollection<ScopeOption>();

        public ObservableCollection<string> ScopeSelectionSetOptions { get; } =
            new ObservableCollection<string>();

        public ObservableCollection<HueGroupAutoAssignPreviewRow> AutoAssignPreviewRows { get; } =
            new ObservableCollection<HueGroupAutoAssignPreviewRow>();

        private ScraperProfileOption _selectedScraperProfile;
        public ScraperProfileOption SelectedScraperProfile
        {
            get => _selectedScraperProfile;
            set
            {
                if (SetField(ref _selectedScraperProfile, value))
                {
                    UpdateCurrentSessionForProfile();
                }
            }
        }

        private string _quickColourCategory = "";
        public string QuickColourCategory
        {
            get => _quickColourCategory;
            set
            {
                if (SetField(ref _quickColourCategory, value ?? ""))
                {
                    EnsureQuickColourCategoryOption(_quickColourCategory);
                    UpdateQuickColourPropertyOptionsFor(_quickColourCategory, QuickColourPropertyOptions, QuickColourProperty);
                }
            }
        }

        private string _quickColourProperty = "";
        public string QuickColourProperty
        {
            get => _quickColourProperty;
            set
            {
                if (SetField(ref _quickColourProperty, value ?? ""))
                {
                    EnsurePropertyOption(QuickColourPropertyOptions, _quickColourProperty);
                }
            }
        }

        private MicroEngPaletteKind _quickColourPaletteKind = MicroEngPaletteKind.Deep;
        public MicroEngPaletteKind QuickColourPaletteKind
        {
            get => _quickColourPaletteKind;
            set
            {
                if (SetField(ref _quickColourPaletteKind, value))
                {
                    OnPropertyChanged(nameof(IsQuickColourCustomPalette));
                    OnPropertyChanged(nameof(QuickColourCustomBaseColorBrush));
                    AssignQuickColourPalette();
                }
            }
        }

        private string _quickColourCustomBaseColorHex = "#FF6699";
        public string QuickColourCustomBaseColorHex
        {
            get => _quickColourCustomBaseColorHex;
            set
            {
                if (SetField(ref _quickColourCustomBaseColorHex, value ?? "#FF6699"))
                {
                    OnPropertyChanged(nameof(QuickColourCustomBaseColorBrush));
                    if (IsQuickColourCustomPalette)
                    {
                        AssignQuickColourPalette();
                    }
                }
            }
        }

        public Brush QuickColourCustomBaseColorBrush
            => new SolidColorBrush(ColourPaletteGenerator.ParseHexOrDefault(QuickColourCustomBaseColorHex, Colors.HotPink));

        public bool IsQuickColourCustomPalette =>
            QuickColourPaletteKind == MicroEngPaletteKind.CustomHue
            || QuickColourPaletteKind == MicroEngPaletteKind.Shades;

        private bool _quickColourStableColors;
        public bool QuickColourStableColors
        {
            get => _quickColourStableColors;
            set
            {
                if (SetField(ref _quickColourStableColors, value))
                {
                    AssignQuickColourPalette();
                }
            }
        }

        private string _quickColourStableSeed = "";
        public string QuickColourStableSeed
        {
            get => _quickColourStableSeed;
            set
            {
                if (SetField(ref _quickColourStableSeed, value ?? ""))
                {
                    if (QuickColourStableColors)
                    {
                        AssignQuickColourPalette();
                    }
                }
            }
        }

        private QuickColourScope _quickColourScope = QuickColourScope.EntireModel;
        public QuickColourScope QuickColourScope
        {
            get => _quickColourScope;
            set
            {
                if (SetField(ref _quickColourScope, value))
                {
                    OnPropertyChanged(nameof(IsQuickColourSelectionSetScope));
                    OnPropertyChanged(nameof(IsQuickColourScopePickerEnabled));
                    EnsureQuickColourScopeSelectionSet();
                    UpdateQuickColourScopeSummary();
                }
            }
        }

        private string _quickColourScopeSelectionSetName = "";
        public string QuickColourScopeSelectionSetName
        {
            get => _quickColourScopeSelectionSetName;
            set
            {
                if (SetField(ref _quickColourScopeSelectionSetName, value ?? ""))
                {
                    EnsureScopeSelectionSetOption(_quickColourScopeSelectionSetName);
                    UpdateQuickColourScopeSummary();
                }
            }
        }

        private List<List<string>> _quickColourScopeModelPaths = new List<List<string>>();
        public List<List<string>> QuickColourScopeModelPaths
        {
            get => _quickColourScopeModelPaths;
            set
            {
                if (SetField(ref _quickColourScopeModelPaths, value ?? new List<List<string>>()))
                {
                    UpdateQuickColourScopeSummary();
                }
            }
        }

        private string _quickColourScopeFilterCategory = "";
        public string QuickColourScopeFilterCategory
        {
            get => _quickColourScopeFilterCategory;
            set
            {
                if (SetField(ref _quickColourScopeFilterCategory, value ?? ""))
                {
                    UpdateQuickColourScopeSummary();
                }
            }
        }

        private string _quickColourScopeFilterProperty = "";
        public string QuickColourScopeFilterProperty
        {
            get => _quickColourScopeFilterProperty;
            set
            {
                if (SetField(ref _quickColourScopeFilterProperty, value ?? ""))
                {
                    UpdateQuickColourScopeSummary();
                }
            }
        }

        private string _quickColourScopeFilterValue = "";
        public string QuickColourScopeFilterValue
        {
            get => _quickColourScopeFilterValue;
            set
            {
                if (SetField(ref _quickColourScopeFilterValue, value ?? ""))
                {
                    UpdateQuickColourScopeSummary();
                }
            }
        }

        private string _quickColourScopeSummary = "Entire model";
        public string QuickColourScopeSummary
        {
            get => _quickColourScopeSummary;
            set => SetField(ref _quickColourScopeSummary, value ?? "");
        }

        public bool IsQuickColourSelectionSetScope => QuickColourScope == QuickColourScope.SavedSelectionSet;

        public bool IsQuickColourScopePickerEnabled =>
            QuickColourScope == QuickColourScope.ModelTree || QuickColourScope == QuickColourScope.PropertyFilter;

        private bool _quickColourCreateSearchSets;
        public bool QuickColourCreateSearchSets
        {
            get => _quickColourCreateSearchSets;
            set => SetField(ref _quickColourCreateSearchSets, value);
        }

        private bool _quickColourCreateSnapshots;
        public bool QuickColourCreateSnapshots
        {
            get => _quickColourCreateSnapshots;
            set => SetField(ref _quickColourCreateSnapshots, value);
        }

        private string _quickColourFolderPath = "MicroEng/Quick Colour";
        public string QuickColourFolderPath
        {
            get => _quickColourFolderPath;
            set => SetField(ref _quickColourFolderPath, value ?? "");
        }

        private string _quickColourProfileName = "Quick Colour";
        public string QuickColourProfileName
        {
            get => _quickColourProfileName;
            set => SetField(ref _quickColourProfileName, value ?? "");
        }

        private string _hierarchyL1Category = "";
        public string HierarchyL1Category
        {
            get => _hierarchyL1Category;
            set
            {
                if (SetField(ref _hierarchyL1Category, value ?? ""))
                {
                    EnsureCategoryOption(_hierarchyL1Category);
                    UpdateHierarchyPropertyOptionsFor(_hierarchyL1Category, HierarchyL1PropertyOptions, HierarchyL1Property);
                    HierarchyL1Label = FormatPropertyLabel(_hierarchyL1Category, _hierarchyL1Property);
                }
            }
        }

        private string _hierarchyL1Property = "";
        public string HierarchyL1Property
        {
            get => _hierarchyL1Property;
            set
            {
                if (SetField(ref _hierarchyL1Property, value ?? ""))
                {
                    EnsurePropertyOption(HierarchyL1PropertyOptions, _hierarchyL1Property);
                    HierarchyL1Label = FormatPropertyLabel(_hierarchyL1Category, _hierarchyL1Property);
                }
            }
        }

        private string _hierarchyL1Label = "Pick property...";
        public string HierarchyL1Label
        {
            get => _hierarchyL1Label;
            set => SetField(ref _hierarchyL1Label, value ?? "");
        }

        private string _hierarchyL2Category = "";
        public string HierarchyL2Category
        {
            get => _hierarchyL2Category;
            set
            {
                if (SetField(ref _hierarchyL2Category, value ?? ""))
                {
                    EnsureCategoryOption(_hierarchyL2Category);
                    UpdateHierarchyPropertyOptionsFor(_hierarchyL2Category, HierarchyL2PropertyOptions, HierarchyL2Property);
                    HierarchyL2Label = FormatPropertyLabel(_hierarchyL2Category, _hierarchyL2Property);
                }
            }
        }

        private string _hierarchyL2Property = "";
        public string HierarchyL2Property
        {
            get => _hierarchyL2Property;
            set
            {
                if (SetField(ref _hierarchyL2Property, value ?? ""))
                {
                    EnsurePropertyOption(HierarchyL2PropertyOptions, _hierarchyL2Property);
                    HierarchyL2Label = FormatPropertyLabel(_hierarchyL2Category, _hierarchyL2Property);
                }
            }
        }

        private string _hierarchyL2Label = "Pick property...";
        public string HierarchyL2Label
        {
            get => _hierarchyL2Label;
            set => SetField(ref _hierarchyL2Label, value ?? "");
        }

        private MicroEngPaletteKind _hierarchyPaletteKind = MicroEngPaletteKind.Deep;
        public MicroEngPaletteKind HierarchyPaletteKind
        {
            get => _hierarchyPaletteKind;
            set
            {
                if (SetField(ref _hierarchyPaletteKind, value))
                {
                    OnPropertyChanged(nameof(IsHierarchyHueEditEnabled));
                    AutoAssignBaseColours();
                    ApplyHierarchyColours();
                }
            }
        }

        private double _hierarchyShadeSpreadPct = 60;
        public double HierarchyShadeSpreadPct
        {
            get => _hierarchyShadeSpreadPct;
            set
            {
                if (SetField(ref _hierarchyShadeSpreadPct, value))
                {
                    ApplyHierarchyColours();
                }
            }
        }

        private QuickColourTypeSortMode _hierarchyTypeSortMode = QuickColourTypeSortMode.Count;
        public QuickColourTypeSortMode HierarchyTypeSortMode
        {
            get => _hierarchyTypeSortMode;
            set
            {
                if (SetField(ref _hierarchyTypeSortMode, value))
                {
                    ApplyHierarchyColours();
                }
            }
        }

        private bool _hierarchyIncludeBlanks;
        public bool HierarchyIncludeBlanks
        {
            get => _hierarchyIncludeBlanks;
            set => SetField(ref _hierarchyIncludeBlanks, value);
        }

        private int _hierarchyMaxCategories = 60;
        public int HierarchyMaxCategories
        {
            get => _hierarchyMaxCategories;
            set => SetField(ref _hierarchyMaxCategories, Math.Max(1, value));
        }

        private int _hierarchyMaxTypesPerCategory = 40;
        public int HierarchyMaxTypesPerCategory
        {
            get => _hierarchyMaxTypesPerCategory;
            set => SetField(ref _hierarchyMaxTypesPerCategory, Math.Max(1, value));
        }

        private bool _hierarchySingleHueMode;
        public bool HierarchySingleHueMode
        {
            get => _hierarchySingleHueMode;
            set
            {
                if (SetField(ref _hierarchySingleHueMode, value))
                {
                    OnPropertyChanged(nameof(HierarchyHueSwatchBrush));
                    OnPropertyChanged(nameof(IsHierarchyHueEditEnabled));
                    ApplyHierarchyColours();
                }
            }
        }

        private string _hierarchyHueHex = "#00A000";
        public string HierarchyHueHex
        {
            get => _hierarchyHueHex;
            set
            {
                if (SetField(ref _hierarchyHueHex, value ?? "#00A000"))
                {
                    OnPropertyChanged(nameof(HierarchyHueSwatchBrush));
                    if (HierarchySingleHueMode
                        || HierarchyPaletteKind == MicroEngPaletteKind.CustomHue
                        || HierarchyPaletteKind == MicroEngPaletteKind.Shades)
                    {
                        AutoAssignBaseColours();
                        ApplyHierarchyColours();
                    }
                }
            }
        }

        private double _hierarchyCategoryContrastPct = 70;
        public double HierarchyCategoryContrastPct
        {
            get => _hierarchyCategoryContrastPct;
            set
            {
                if (SetField(ref _hierarchyCategoryContrastPct, value))
                {
                    if (HierarchySingleHueMode || HierarchyUseHueGroups)
                    {
                        ApplyHierarchyColours();
                    }
                }
            }
        }

        public Brush HierarchyHueSwatchBrush => new SolidColorBrush(ParseHueColor());
        public bool IsHierarchyHueEditEnabled =>
            HierarchySingleHueMode
            || HierarchyPaletteKind == MicroEngPaletteKind.CustomHue
            || HierarchyPaletteKind == MicroEngPaletteKind.Shades;
        private bool _hierarchyUseHueGroups;
        public bool HierarchyUseHueGroups
        {
            get => _hierarchyUseHueGroups;
            set
            {
                if (SetField(ref _hierarchyUseHueGroups, value))
                {
                    RefreshHueGroupOptions();
                    EnsureCategoryHueGroupDefaults();
                    ApplyHierarchyColours();
                }
            }
        }

        private bool _hierarchyShadeHueGroups = true;
        public bool HierarchyShadeHueGroups
        {
            get => _hierarchyShadeHueGroups;
            set
            {
                if (SetField(ref _hierarchyShadeHueGroups, value))
                {
                    ApplyHierarchyColours();
                }
            }
        }

        private bool _createFoldersByHueGroup;
        public bool CreateFoldersByHueGroup
        {
            get => _createFoldersByHueGroup;
            set => SetField(ref _createFoldersByHueGroup, value);
        }

        private QuickColourScope _selectedScope = QuickColourScope.EntireModel;
        public QuickColourScope SelectedScope
        {
            get => _selectedScope;
            set
            {
                if (SetField(ref _selectedScope, value))
                {
                    OnPropertyChanged(nameof(IsHierarchySelectionSetScope));
                    OnPropertyChanged(nameof(IsHierarchyScopePickerEnabled));
                    EnsureHierarchyScopeSelectionSet();
                    UpdateHierarchyScopeSummary();
                }
            }
        }

        private string _hierarchyScopeSelectionSetName = "";
        public string HierarchyScopeSelectionSetName
        {
            get => _hierarchyScopeSelectionSetName;
            set
            {
                if (SetField(ref _hierarchyScopeSelectionSetName, value ?? ""))
                {
                    EnsureScopeSelectionSetOption(_hierarchyScopeSelectionSetName);
                    UpdateHierarchyScopeSummary();
                }
            }
        }

        private List<List<string>> _hierarchyScopeModelPaths = new List<List<string>>();
        public List<List<string>> HierarchyScopeModelPaths
        {
            get => _hierarchyScopeModelPaths;
            set
            {
                if (SetField(ref _hierarchyScopeModelPaths, value ?? new List<List<string>>()))
                {
                    UpdateHierarchyScopeSummary();
                }
            }
        }

        private string _hierarchyScopeFilterCategory = "";
        public string HierarchyScopeFilterCategory
        {
            get => _hierarchyScopeFilterCategory;
            set
            {
                if (SetField(ref _hierarchyScopeFilterCategory, value ?? ""))
                {
                    UpdateHierarchyScopeSummary();
                }
            }
        }

        private string _hierarchyScopeFilterProperty = "";
        public string HierarchyScopeFilterProperty
        {
            get => _hierarchyScopeFilterProperty;
            set
            {
                if (SetField(ref _hierarchyScopeFilterProperty, value ?? ""))
                {
                    UpdateHierarchyScopeSummary();
                }
            }
        }

        private string _hierarchyScopeFilterValue = "";
        public string HierarchyScopeFilterValue
        {
            get => _hierarchyScopeFilterValue;
            set
            {
                if (SetField(ref _hierarchyScopeFilterValue, value ?? ""))
                {
                    UpdateHierarchyScopeSummary();
                }
            }
        }

        private string _hierarchyScopeSummary = "Entire model";
        public string HierarchyScopeSummary
        {
            get => _hierarchyScopeSummary;
            set => SetField(ref _hierarchyScopeSummary, value ?? "");
        }

        public bool IsHierarchySelectionSetScope => SelectedScope == QuickColourScope.SavedSelectionSet;

        public bool IsHierarchyScopePickerEnabled =>
            SelectedScope == QuickColourScope.ModelTree || SelectedScope == QuickColourScope.PropertyFilter;

        private bool _createSearchSets;
        public bool CreateSearchSets
        {
            get => _createSearchSets;
            set => SetField(ref _createSearchSets, value);
        }

        private bool _createSnapshots;
        public bool CreateSnapshots
        {
            get => _createSnapshots;
            set => SetField(ref _createSnapshots, value);
        }

        private string _outputFolderPath = "MicroEng/Quick Colour";
        public string OutputFolderPath
        {
            get => _outputFolderPath;
            set => SetField(ref _outputFolderPath, value ?? "");
        }

        private string _outputProfileName = "Quick Colour";
        public string OutputProfileName
        {
            get => _outputProfileName;
            set => SetField(ref _outputProfileName, value ?? "");
        }

        private string _statusText = "Ready";
        public string StatusText
        {
            get => _statusText;
            set => SetField(ref _statusText, value ?? "");
        }

        private QuickColourHueGroup _selectedHueGroup;
        public QuickColourHueGroup SelectedHueGroup
        {
            get => _selectedHueGroup;
            set => SetField(ref _selectedHueGroup, value);
        }

        private string _disciplineMapPath;
        public string DisciplineMapPath
        {
            get => _disciplineMapPath;
            set => SetField(ref _disciplineMapPath, value ?? "");
        }

        private bool _autoAssignOnlyFallback = true;
        public bool AutoAssignOnlyFallback
        {
            get => _autoAssignOnlyFallback;
            set => SetField(ref _autoAssignOnlyFallback, value);
        }

        private bool _autoAssignSkipLocked;
        public bool AutoAssignSkipLocked
        {
            get => _autoAssignSkipLocked;
            set => SetField(ref _autoAssignSkipLocked, value);
        }

        private bool _autoAssignCreateMissingGroups;
        public bool AutoAssignCreateMissingGroups
        {
            get => _autoAssignCreateMissingGroups;
            set => SetField(ref _autoAssignCreateMissingGroups, value);
        }

        private string _autoAssignPreviewSummary = "";
        public string AutoAssignPreviewSummary
        {
            get => _autoAssignPreviewSummary;
            set => SetField(ref _autoAssignPreviewSummary, value ?? "");
        }

        private bool _previewShowChangesOnly;
        public bool PreviewShowChangesOnly
        {
            get => _previewShowChangesOnly;
            set
            {
                if (SetField(ref _previewShowChangesOnly, value))
                {
                    ApplyPreviewFilter();
                }
            }
        }

        public QuickColourControl()
        {
            InitializeComponent();
            DataContext = this;

            try
            {
                MicroEngWpfUiTheme.ApplyTo(this);
            }
            catch
            {
                // ignore theme failures
            }

            HierarchyGroups.CollectionChanged += OnHierarchyGroupsChanged;
            HueGroups.CollectionChanged += OnHueGroupsChanged;

            foreach (var option in BuildPaletteOptions())
            {
                QuickColourPaletteOptions.Add(option);
            }

            foreach (QuickColourPaletteStyle style in Enum.GetValues(typeof(QuickColourPaletteStyle)))
            {
                PaletteStyleOptions.Add(style);
            }

            foreach (QuickColourTypeSortMode mode in Enum.GetValues(typeof(QuickColourTypeSortMode)))
            {
                TypeSortOptions.Add(mode);
            }

            foreach (var option in BuildScopeOptions())
            {
                ScopeOptions.Add(option);
            }

            EnsureDefaultHueGroups();
            RefreshHueGroupOptions();
            RefreshProfiles();
            InitDisciplineMap();
            RefreshScopeSelectionSetOptions();
            UpdateQuickColourScopeSummary();
            UpdateHierarchyScopeSummary();

            if (ProfilesPage != null)
            {
                ProfilesPage.ApplyRequested += OnProfileApplyRequested;
            }

            DataScraperCache.SessionAdded += OnSessionAdded;
            DataScraperCache.CacheChanged += OnCacheChanged;
            Unloaded += (_, __) =>
            {
                DataScraperCache.SessionAdded -= OnSessionAdded;
                DataScraperCache.CacheChanged -= OnCacheChanged;
            };
        }

        private void OnSessionAdded(ScrapeSession session)
        {
            RefreshProfiles(session?.Id);
        }

        private void OnCacheChanged()
        {
            RefreshProfiles();
        }

        private void RefreshProfiles(Guid? preferredSessionId = null)
        {
            var targetSessionId = preferredSessionId ?? SelectedScraperProfile?.Session?.Id;
            var existingById = new Dictionary<Guid, ScraperProfileOption>();
            foreach (var option in ScraperProfiles)
            {
                var session = option?.Session;
                if (session == null)
                {
                    continue;
                }

                if (!existingById.ContainsKey(session.Id))
                {
                    existingById[session.Id] = option;
                }
            }

            ScraperProfiles.Clear();

            var sessions = DataScraperCache.AllSessions
                .OrderByDescending(s => s.Timestamp)
                .ToList();

            foreach (var session in sessions)
            {
                if (session == null)
                {
                    continue;
                }

                if (!existingById.TryGetValue(session.Id, out var option))
                {
                    option = new ScraperProfileOption(session);
                }

                ScraperProfiles.Add(option);
            }

            ScraperProfileOption selected = null;
            if (targetSessionId.HasValue)
            {
                selected = ScraperProfiles.FirstOrDefault(o => o?.Session?.Id == targetSessionId.Value);
            }

            if (selected == null)
            {
                var lastSessionId = DataScraperCache.LastSession?.Id;
                if (lastSessionId.HasValue)
                {
                    selected = ScraperProfiles.FirstOrDefault(o => o?.Session?.Id == lastSessionId.Value);
                }
            }

            selected ??= ScraperProfiles.FirstOrDefault();

            if (selected == null && DataScraperCache.LastSession != null)
            {
                selected = new ScraperProfileOption(DataScraperCache.LastSession);
                ScraperProfiles.Add(selected);
            }

            if (!ReferenceEquals(SelectedScraperProfile, selected))
            {
                SelectedScraperProfile = selected;
            }

            RefreshScopeSelectionSetOptions();
        }

        private void UpdateCurrentSessionForProfile()
        {
            RefreshQuickColourPropertyOptions();
            RefreshHierarchyPropertyOptions();
            EnsureDefaultHierarchyProperties();
            StatusText = SelectedScraperProfile?.Label ?? "Ready";
        }

        private ScrapeSession GetCurrentSession()
        {
            return SelectedScraperProfile?.Session ?? DataScraperCache.LastSession;
        }

        private void RefreshQuickColourPropertyOptions()
        {
            _quickColourPropertyOptionsByCategory.Clear();
            var propertyNamesByCategory = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            var session = GetCurrentSession();
            var properties = session?.Properties ?? Enumerable.Empty<ScrapedProperty>();

            foreach (var prop in properties)
            {
                var category = prop?.Category ?? "";
                var name = prop?.Name ?? "";
                if (string.IsNullOrWhiteSpace(category) || string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (!_quickColourPropertyOptionsByCategory.TryGetValue(category, out var list))
                {
                    list = new List<string>();
                    _quickColourPropertyOptionsByCategory[category] = list;
                    propertyNamesByCategory[category] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }

                if (propertyNamesByCategory.TryGetValue(category, out var names) && names.Add(name))
                {
                    list.Add(name);
                }
            }

            foreach (var list in _quickColourPropertyOptionsByCategory.Values)
            {
                list.Sort(StringComparer.OrdinalIgnoreCase);
            }

            var categories = new HashSet<string>(_quickColourPropertyOptionsByCategory.Keys, StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(QuickColourCategory))
            {
                categories.Add(QuickColourCategory);
            }

            var orderedCategories = categories
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .OrderBy(c => c, StringComparer.OrdinalIgnoreCase)
                .ToList();

            QuickColourCategoryOptions.Clear();
            foreach (var cat in orderedCategories)
            {
                QuickColourCategoryOptions.Add(cat);
            }

            UpdateQuickColourPropertyOptionsFor(QuickColourCategory, QuickColourPropertyOptions, QuickColourProperty);
        }

        private void UpdateQuickColourPropertyOptionsFor(string category, ObservableCollection<string> targetOptions, string selectedValue)
        {
            targetOptions.Clear();

            if (!string.IsNullOrWhiteSpace(category)
                && _quickColourPropertyOptionsByCategory.TryGetValue(category, out var options))
            {
                foreach (var option in options)
                {
                    targetOptions.Add(option);
                }
            }

            EnsurePropertyOption(targetOptions, selectedValue);
        }

        private void EnsureQuickColourCategoryOption(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
            {
                return;
            }

            if (ContainsIgnoreCase(QuickColourCategoryOptions, category))
            {
                return;
            }

            QuickColourCategoryOptions.Add(category);
        }

        private void LoadQuickColourValues_Click(object sender, RoutedEventArgs e)
        {
            LoadQuickColourValues();
        }

        private void LoadQuickColourValues()
        {
            var session = GetCurrentSession();
            if (session == null)
            {
                StatusText = "No Data Scraper session available. Run Data Scraper first.";
                return;
            }

            if (string.IsNullOrWhiteSpace(QuickColourCategory) || string.IsNullOrWhiteSpace(QuickColourProperty))
            {
                StatusText = "Pick Category and Property first.";
                return;
            }

            if (!TryGetScopeItemKeys(
                session,
                QuickColourScope,
                QuickColourScopeSelectionSetName,
                QuickColourScopeModelPaths,
                QuickColourScopeFilterCategory,
                QuickColourScopeFilterProperty,
                QuickColourScopeFilterValue,
                out var allowedKeys,
                out var scopeError))
            {
                StatusText = scopeError;
                return;
            }

            var rows = QuickColourValueBuilder.BuildValues(session, QuickColourCategory, QuickColourProperty, allowedKeys);
            QuickColourValues.Clear();
            foreach (var row in rows)
            {
                QuickColourValues.Add(row);
            }

            AssignQuickColourPalette();
            if (rows.Count == 0)
            {
                StatusText = session.RawEntriesTruncated || (session.RawEntries?.Count ?? 0) == 0
                    ? "No values found. Raw rows are not fully cached. Re-run Data Scraper with 'Keep raw rows in memory' enabled."
                    : "No values found for this property.";
            }
            else
            {
                StatusText = session.RawEntriesTruncated
                    ? $"Loaded {rows.Count} value(s) (preview only). Re-run Data Scraper with 'Keep raw rows in memory' enabled for full values."
                    : $"Loaded {rows.Count} value(s).";
            }
        }

        private void AutoAssignQuickColours_Click(object sender, RoutedEventArgs e)
        {
            AssignQuickColourPalette();
        }

        private bool TryGetScopeItemKeys(
            ScrapeSession session,
            QuickColourScope scope,
            string selectionSetName,
            List<List<string>> modelPaths,
            string filterCategory,
            string filterProperty,
            string filterValue,
            out HashSet<string> keys,
            out string error)
        {
            keys = null;
            error = "";

            if (scope == QuickColourScope.EntireModel)
            {
                return true;
            }

            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                error = "No active document.";
                return false;
            }

            switch (scope)
            {
                case QuickColourScope.CurrentSelection:
                    var sel = doc.CurrentSelection?.SelectedItems;
                    if (sel == null || sel.Count == 0)
                    {
                        error = "Scope is Current Selection, but no items are selected.";
                        return false;
                    }

                    keys = CollectItemKeys(sel);
                    if (keys.Count == 0)
                    {
                        error = "Scope selection has no items.";
                        return false;
                    }

                    return true;

                case QuickColourScope.SavedSelectionSet:
                    if (string.IsNullOrWhiteSpace(selectionSetName))
                    {
                        error = "Scope is Saved Selection Set, but no set is selected.";
                        return false;
                    }

                    var setItems = GetSelectionSetItems(doc, selectionSetName);
                    if (setItems == null || setItems.Count == 0)
                    {
                        error = $"Scope selection set not found or empty: {selectionSetName}.";
                        return false;
                    }

                    keys = CollectItemKeys(setItems);
                    if (keys.Count == 0)
                    {
                        error = $"Scope selection set returned no items: {selectionSetName}.";
                        return false;
                    }

                    return true;

                case QuickColourScope.ModelTree:
                    if (modelPaths == null || modelPaths.Count == 0)
                    {
                        error = "Scope is Tree Selection, but no paths are selected.";
                        return false;
                    }

                    var treeItems = ResolveScopeModelItems(doc, modelPaths);
                    if (treeItems == null || treeItems.Count == 0)
                    {
                        error = "Scope tree selection could not be resolved.";
                        return false;
                    }

                    keys = CollectItemKeys(treeItems);
                    if (keys.Count == 0)
                    {
                        error = "Scope tree selection returned no items.";
                        return false;
                    }

                    return true;

                case QuickColourScope.PropertyFilter:
                    if (string.IsNullOrWhiteSpace(filterCategory)
                        || string.IsNullOrWhiteSpace(filterProperty)
                        || string.IsNullOrWhiteSpace(filterValue))
                    {
                        error = "Scope is Property Filter, but it is incomplete.";
                        return false;
                    }

                    if (session == null)
                    {
                        error = "No Data Scraper session available. Run Data Scraper first.";
                        return false;
                    }

                    keys = BuildPropertyFilterKeys(session, filterCategory, filterProperty, filterValue);
                    return true;
            }

            return true;
        }

        private bool ValidateScopeForApply(
            QuickColourScope scope,
            string selectionSetName,
            List<List<string>> modelPaths,
            string filterCategory,
            string filterProperty,
            string filterValue,
            out string error)
        {
            error = "";

            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                error = "No active document.";
                return false;
            }

            switch (scope)
            {
                case QuickColourScope.EntireModel:
                    return true;
                case QuickColourScope.CurrentSelection:
                    var sel = doc.CurrentSelection?.SelectedItems;
                    if (sel == null || sel.Count == 0)
                    {
                        error = "Scope is Current Selection, but no items are selected.";
                        return false;
                    }
                    return true;
                case QuickColourScope.SavedSelectionSet:
                    if (string.IsNullOrWhiteSpace(selectionSetName))
                    {
                        error = "Scope is Saved Selection Set, but no set is selected.";
                        return false;
                    }
                    var setItems = GetSelectionSetItems(doc, selectionSetName);
                    if (setItems == null || setItems.Count == 0)
                    {
                        error = $"Scope selection set not found or empty: {selectionSetName}.";
                        return false;
                    }
                    return true;
                case QuickColourScope.ModelTree:
                    if (modelPaths == null || modelPaths.Count == 0)
                    {
                        error = "Scope is Tree Selection, but no paths are selected.";
                        return false;
                    }
                    var resolved = ResolveScopeModelItems(doc, modelPaths);
                    if (resolved == null || resolved.Count == 0)
                    {
                        error = "Scope tree selection could not be resolved.";
                        return false;
                    }
                    return true;
                case QuickColourScope.PropertyFilter:
                    if (string.IsNullOrWhiteSpace(filterCategory)
                        || string.IsNullOrWhiteSpace(filterProperty)
                        || string.IsNullOrWhiteSpace(filterValue))
                    {
                        error = "Scope is Property Filter, but it is incomplete.";
                        return false;
                    }
                    return true;
                default:
                    return true;
            }
        }

        private void PickQuickColourBaseColor_Click(object sender, RoutedEventArgs e)
        {
            var current = ColourPaletteGenerator.ParseHexOrDefault(QuickColourCustomBaseColorHex, Colors.HotPink);
            using (var dlg = new System.Windows.Forms.ColorDialog
            {
                FullOpen = true,
                AnyColor = true,
                Color = System.Drawing.Color.FromArgb(current.R, current.G, current.B)
            })
            {
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    QuickColourCustomBaseColorHex = $"#{dlg.Color.R:X2}{dlg.Color.G:X2}{dlg.Color.B:X2}";
                }
            }
        }

        private void PickQuickColourValueColor_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is FrameworkElement element))
            {
                return;
            }

            if (!(element.DataContext is QuickColourValueRow row))
            {
                return;
            }

            var current = row.Color;
            using (var dlg = new System.Windows.Forms.ColorDialog
            {
                FullOpen = true,
                AnyColor = true,
                Color = System.Drawing.Color.FromArgb(current.R, current.G, current.B)
            })
            {
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    row.Color = System.Windows.Media.Color.FromRgb(dlg.Color.R, dlg.Color.G, dlg.Color.B);
                }
            }
        }

        private void AssignQuickColourPalette()
        {
            if (QuickColourValues == null || QuickColourValues.Count == 0)
            {
                return;
            }

            if (QuickColourPaletteKind == MicroEngPaletteKind.Shades)
            {
                var baseColor = ColourPaletteGenerator.ParseHexOrDefault(
                    QuickColourCustomBaseColorHex,
                    Colors.HotPink);
                QuickColourPalette.AssignShades(
                    QuickColourValues,
                    baseColor,
                    QuickColourStableColors,
                    QuickColourStableSeed);
                return;
            }

            if (QuickColourPaletteKind == MicroEngPaletteKind.CustomHue)
            {
                var baseStyle = QuickColourPaletteStyle.Deep;
                var baseRows = new List<QuickColourValueRow>(QuickColourValues.Count);
                foreach (var row in QuickColourValues)
                {
                    baseRows.Add(new QuickColourValueRow
                    {
                        Value = row?.Value ?? ""
                    });
                }

                if (QuickColourStableColors)
                {
                    QuickColourPalette.AssignStableColors(baseRows, baseStyle, QuickColourStableSeed);
                }
                else
                {
                    QuickColourPalette.AssignPalette(baseRows, baseStyle);
                }

                var baseColor = ColourPaletteGenerator.ParseHexOrDefault(
                    QuickColourCustomBaseColorHex,
                    Colors.HotPink);
                ColourPaletteGenerator.RgbToHsl(baseColor, out _, out _, out var targetL);
                var baseL = baseStyle == QuickColourPaletteStyle.Pastel ? 0.78 : 0.55;
                var deltaL = targetL - baseL;

                for (var i = 0; i < QuickColourValues.Count && i < baseRows.Count; i++)
                {
                    var row = QuickColourValues[i];
                    var baseRow = baseRows[i];
                    if (row == null || baseRow == null)
                    {
                        continue;
                    }

                    ColourPaletteGenerator.RgbToHsl(baseRow.Color, out var h, out var s, out var l);
                    var newL = Math.Max(0.18, Math.Min(0.92, l + deltaL));
                    row.Color = ColourPaletteGenerator.HslToRgb(h, s, newL);
                }

                return;
            }

            if (QuickColourPaletteKind == MicroEngPaletteKind.Deep && QuickColourStableColors)
            {
                QuickColourPalette.AssignStableColors(
                    QuickColourValues,
                    QuickColourPaletteStyle.Deep,
                    QuickColourStableSeed);
                return;
            }

            var palette = QuickColourPalette.GeneratePalette(QuickColourValues.Count, QuickColourPaletteKind);
            QuickColourPalette.AssignPalette(
                QuickColourValues,
                palette,
                QuickColourStableColors,
                QuickColourStableSeed);
        }

        private void ApplyQuickColourTemporary_Click(object sender, RoutedEventArgs e)
        {
            ApplyQuickColour(permanent: false, sender as System.Windows.Controls.Button);
        }

        private void ApplyQuickColourPermanent_Click(object sender, RoutedEventArgs e)
        {
            ApplyQuickColour(permanent: true, sender as System.Windows.Controls.Button);
        }

        private void ApplyQuickColour(bool permanent,
            System.Windows.Controls.Button sourceButton = null,
            string categoryOverride = null,
            string propertyOverride = null,
            MicroEngColourProfile profileOverride = null)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                StatusText = "No active document.";
                return;
            }

            if (QuickColourValues == null || QuickColourValues.Count == 0)
            {
                StatusText = "No values loaded.";
                return;
            }

            var category = string.IsNullOrWhiteSpace(categoryOverride) ? QuickColourCategory : categoryOverride;
            var property = string.IsNullOrWhiteSpace(propertyOverride) ? QuickColourProperty : propertyOverride;

            if (string.IsNullOrWhiteSpace(category) || string.IsNullOrWhiteSpace(property))
            {
                StatusText = profileOverride != null
                    ? "Profile is missing Category/Property info."
                    : "Pick Category and Property first.";
                return;
            }

            if (!ValidateScopeForApply(
                QuickColourScope,
                QuickColourScopeSelectionSetName,
                QuickColourScopeModelPaths,
                QuickColourScopeFilterCategory,
                QuickColourScopeFilterProperty,
                QuickColourScopeFilterValue,
                out var scopeError))
            {
                StatusText = scopeError;
                return;
            }

            try
            {
                var expected = QuickColourValues
                    .Where(v => v != null && v.Enabled)
                    .Sum(v => Math.Max(0, v.Count));
                if (IsQuickColourTraceEnabled())
                {
                    Log($"QuickColour: expected count from cache={expected} for {category}.{property}.");
                }

                var count = _service.ApplyBySingleProperty(
                    doc,
                    category,
                    property,
                    QuickColourValues,
                    QuickColourScope,
                    QuickColourScopeSelectionSetName,
                    QuickColourScopeModelPaths,
                    QuickColourScopeFilterCategory,
                    QuickColourScopeFilterProperty,
                    QuickColourScopeFilterValue,
                    permanent,
                    QuickColourCreateSearchSets,
                    QuickColourCreateSnapshots,
                    QuickColourFolderPath,
                    QuickColourProfileName,
                    Log);

                Log($"QuickColour: cached entries={expected}, coloured items={count}.");

                StatusText = $"Applied {count} item(s).";
                FlashSuccess(sourceButton);
                ShowSnackbar(permanent ? "Applied permanent colours" : "Applied temporary colours",
                    $"Applied to {count} item(s).",
                    WpfUiControls.ControlAppearance.Success,
                    WpfUiControls.SymbolRegular.CheckmarkCircle24);
            }
            catch (Exception ex)
            {
                StatusText = "Apply failed: " + ex.Message;
                ShowSnackbar("Apply failed",
                    ex.Message,
                    WpfUiControls.ControlAppearance.Danger,
                    WpfUiControls.SymbolRegular.ErrorCircle24);
            }
        }

        private void ExportQuickColourLegend_Click(object sender, RoutedEventArgs e)
        {
            if (QuickColourValues == null || QuickColourValues.Count == 0)
            {
                StatusText = "No values to export.";
                return;
            }

            var dlg = new SaveFileDialog
            {
                Title = "Export Quick Colour Legend",
                FileName = string.IsNullOrWhiteSpace(QuickColourProfileName) ? "QuickColourLegend" : QuickColourProfileName,
                Filter = "CSV (*.csv)|*.csv|JSON (*.json)|*.json"
            };

            if (dlg.ShowDialog() != true)
            {
                return;
            }

            var ext = Path.GetExtension(dlg.FileName) ?? "";
            if (string.Equals(ext, ".json", StringComparison.OrdinalIgnoreCase))
            {
                QuickColourLegendExporter.ExportJson(
                    dlg.FileName,
                    QuickColourProfileName,
                    QuickColourCategory,
                    QuickColourProperty,
                    QuickColourScope,
                    QuickColourValues);
            }
            else
            {
                QuickColourLegendExporter.ExportCsv(
                    dlg.FileName,
                    QuickColourProfileName,
                    QuickColourCategory,
                    QuickColourProperty,
                    QuickColourScope,
                    QuickColourValues);
            }

            StatusText = "Legend exported.";
        }

        private void RefreshHierarchyPropertyOptions()
        {
            _hierarchyPropertyOptionsByCategory.Clear();
            var propertyNamesByCategory = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            var session = GetCurrentSession();
            var properties = session?.Properties ?? Enumerable.Empty<ScrapedProperty>();

            foreach (var prop in properties)
            {
                var category = prop?.Category ?? "";
                var name = prop?.Name ?? "";
                if (string.IsNullOrWhiteSpace(category) || string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (!_hierarchyPropertyOptionsByCategory.TryGetValue(category, out var list))
                {
                    list = new List<string>();
                    _hierarchyPropertyOptionsByCategory[category] = list;
                    propertyNamesByCategory[category] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }

                if (propertyNamesByCategory.TryGetValue(category, out var names) && names.Add(name))
                {
                    list.Add(name);
                }
            }

            foreach (var list in _hierarchyPropertyOptionsByCategory.Values)
            {
                list.Sort(StringComparer.OrdinalIgnoreCase);
            }

            var categories = new HashSet<string>(_hierarchyPropertyOptionsByCategory.Keys, StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(HierarchyL1Category))
            {
                categories.Add(HierarchyL1Category);
            }
            if (!string.IsNullOrWhiteSpace(HierarchyL2Category))
            {
                categories.Add(HierarchyL2Category);
            }

            var orderedCategories = categories
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .OrderBy(c => c, StringComparer.OrdinalIgnoreCase)
                .ToList();

            HierarchyCategoryOptions.Clear();
            foreach (var cat in orderedCategories)
            {
                HierarchyCategoryOptions.Add(cat);
            }

            UpdateHierarchyPropertyOptions();
        }

        private void UpdateHierarchyPropertyOptions()
        {
            UpdateHierarchyPropertyOptionsFor(HierarchyL1Category, HierarchyL1PropertyOptions, HierarchyL1Property);
            UpdateHierarchyPropertyOptionsFor(HierarchyL2Category, HierarchyL2PropertyOptions, HierarchyL2Property);
        }

        private void UpdateHierarchyPropertyOptionsFor(string category, ObservableCollection<string> targetOptions, string selectedValue)
        {
            targetOptions.Clear();

            if (!string.IsNullOrWhiteSpace(category)
                && _hierarchyPropertyOptionsByCategory.TryGetValue(category, out var options))
            {
                foreach (var option in options)
                {
                    targetOptions.Add(option);
                }
            }

            EnsurePropertyOption(targetOptions, selectedValue);
        }

        private void EnsureCategoryOption(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
            {
                return;
            }

            if (ContainsIgnoreCase(HierarchyCategoryOptions, category))
            {
                return;
            }

            HierarchyCategoryOptions.Add(category);
        }

        private static void EnsurePropertyOption(ObservableCollection<string> options, string value)
        {
            if (options == null || string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            if (!ContainsIgnoreCase(options, value))
            {
                options.Insert(0, value);
            }
        }

        private static MicroEngPaletteKind ParsePaletteKind(string name, MicroEngPaletteKind fallback)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return fallback;
            }

            var trimmed = name.Trim();

            if (string.Equals(trimmed, "Default", StringComparison.OrdinalIgnoreCase))
            {
                return MicroEngPaletteKind.Deep;
            }

            if (string.Equals(trimmed, "Custom", StringComparison.OrdinalIgnoreCase)
                || string.Equals(trimmed, "Custom (Hue)", StringComparison.OrdinalIgnoreCase)
                || string.Equals(trimmed, "Custom Hue", StringComparison.OrdinalIgnoreCase))
            {
                return MicroEngPaletteKind.CustomHue;
            }

            if (Enum.TryParse(trimmed, true, out MicroEngPaletteKind directParsed))
            {
                return directParsed;
            }

            var normalized = new string(trimmed.Where(char.IsLetterOrDigit).ToArray());
            if (!string.IsNullOrWhiteSpace(normalized)
                && Enum.TryParse(normalized, true, out MicroEngPaletteKind parsed))
            {
                return parsed;
            }

            return fallback;
        }

        private static PaletteOption[] BuildPaletteOptions()
        {
            return new[]
            {
                new PaletteOption(MicroEngPaletteKind.Deep, "Default"),
                new PaletteOption(MicroEngPaletteKind.CustomHue, "Custom"),
                new PaletteOption(MicroEngPaletteKind.Shades, "Shades"),
                new PaletteOption(MicroEngPaletteKind.Beach, "Beach"),
                new PaletteOption(MicroEngPaletteKind.OceanBreeze, "Ocean Breeze"),
                new PaletteOption(MicroEngPaletteKind.Vibrant, "Vibrant"),
                new PaletteOption(MicroEngPaletteKind.Pastel, "Pastel"),
                new PaletteOption(MicroEngPaletteKind.Autumn, "Autumn"),
                new PaletteOption(MicroEngPaletteKind.RedSunset, "Red Sunset"),
                new PaletteOption(MicroEngPaletteKind.ForestHues, "Forest Hues"),
                new PaletteOption(MicroEngPaletteKind.PurpleRaindrops, "Purple Raindrops"),
                new PaletteOption(MicroEngPaletteKind.LightSteel, "Light Steel"),
                new PaletteOption(MicroEngPaletteKind.EarthyBrown, "Earthy Brown"),
                new PaletteOption(MicroEngPaletteKind.EarthyGreen, "Earthy Green"),
                new PaletteOption(MicroEngPaletteKind.WarmNeutrals1, "Warm Neutrals 1"),
                new PaletteOption(MicroEngPaletteKind.WarmNeutrals2, "Warm Neutrals 2"),
                new PaletteOption(MicroEngPaletteKind.CandyPop, "Candy Pop")
            };
        }

        private static ScopeOption[] BuildScopeOptions()
        {
            return new[]
            {
                new ScopeOption(QuickColourScope.EntireModel, "All model"),
                new ScopeOption(QuickColourScope.CurrentSelection, "Current selection"),
                new ScopeOption(QuickColourScope.SavedSelectionSet, "Saved selection set"),
                new ScopeOption(QuickColourScope.ModelTree, "Tree selection"),
                new ScopeOption(QuickColourScope.PropertyFilter, "Property filter")
            };
        }

        private static SmartSetScopeMode ToSmartSetScopeMode(QuickColourScope scope)
        {
            return scope switch
            {
                QuickColourScope.CurrentSelection => SmartSetScopeMode.CurrentSelection,
                QuickColourScope.SavedSelectionSet => SmartSetScopeMode.SavedSelectionSet,
                QuickColourScope.ModelTree => SmartSetScopeMode.ModelTree,
                QuickColourScope.PropertyFilter => SmartSetScopeMode.PropertyFilter,
                _ => SmartSetScopeMode.AllModel
            };
        }

        private static QuickColourScope ToQuickColourScope(SmartSetScopeMode scope)
        {
            return scope switch
            {
                SmartSetScopeMode.CurrentSelection => QuickColourScope.CurrentSelection,
                SmartSetScopeMode.SavedSelectionSet => QuickColourScope.SavedSelectionSet,
                SmartSetScopeMode.ModelTree => QuickColourScope.ModelTree,
                SmartSetScopeMode.PropertyFilter => QuickColourScope.PropertyFilter,
                _ => QuickColourScope.EntireModel
            };
        }

        private static string BuildScopeKind(QuickColourScope scope)
        {
            return scope switch
            {
                QuickColourScope.CurrentSelection => "CurrentSelection",
                QuickColourScope.SavedSelectionSet => "SavedSelectionSet",
                QuickColourScope.ModelTree => "ModelTree",
                QuickColourScope.PropertyFilter => "PropertyFilter",
                _ => "EntireModel"
            };
        }

        private static QuickColourScope ParseScope(string kind)
        {
            if (string.Equals(kind, "CurrentSelection", StringComparison.OrdinalIgnoreCase))
            {
                return QuickColourScope.CurrentSelection;
            }

            if (string.Equals(kind, "AllModel", StringComparison.OrdinalIgnoreCase)
                || string.Equals(kind, "EntireModel", StringComparison.OrdinalIgnoreCase))
            {
                return QuickColourScope.EntireModel;
            }

            if (string.Equals(kind, "SavedSelectionSet", StringComparison.OrdinalIgnoreCase)
                || string.Equals(kind, "SelectionSet", StringComparison.OrdinalIgnoreCase))
            {
                return QuickColourScope.SavedSelectionSet;
            }

            if (string.Equals(kind, "ModelTree", StringComparison.OrdinalIgnoreCase)
                || string.Equals(kind, "TreeSelection", StringComparison.OrdinalIgnoreCase))
            {
                return QuickColourScope.ModelTree;
            }

            if (string.Equals(kind, "PropertyFilter", StringComparison.OrdinalIgnoreCase))
            {
                return QuickColourScope.PropertyFilter;
            }

            return QuickColourScope.EntireModel;
        }

        private void UpdateQuickColourScopeSummary()
        {
            QuickColourScopeSummary = BuildScopeSummary(
                QuickColourScope,
                QuickColourScopeSelectionSetName,
                QuickColourScopeModelPaths,
                QuickColourScopeFilterCategory,
                QuickColourScopeFilterProperty,
                QuickColourScopeFilterValue);
        }

        private void UpdateHierarchyScopeSummary()
        {
            HierarchyScopeSummary = BuildScopeSummary(
                SelectedScope,
                HierarchyScopeSelectionSetName,
                HierarchyScopeModelPaths,
                HierarchyScopeFilterCategory,
                HierarchyScopeFilterProperty,
                HierarchyScopeFilterValue);
        }

        private string BuildScopeSummary(
            QuickColourScope scope,
            string selectionSetName,
            IReadOnlyList<List<string>> modelPaths,
            string filterCategory,
            string filterProperty,
            string filterValue)
        {
            switch (scope)
            {
                case QuickColourScope.EntireModel:
                    return "Entire model";
                case QuickColourScope.CurrentSelection:
                    var count = NavisApp.ActiveDocument?.CurrentSelection?.SelectedItems?.Count ?? 0;
                    return count > 0 ? $"Current selection ({count} items)" : "Current selection";
                case QuickColourScope.SavedSelectionSet:
                    return string.IsNullOrWhiteSpace(selectionSetName)
                        ? "Saved selection set"
                        : $"Saved selection set: {selectionSetName}";
                case QuickColourScope.ModelTree:
                    return BuildTreeScopeSummary(modelPaths);
                case QuickColourScope.PropertyFilter:
                    return BuildFilterScopeSummary(filterCategory, filterProperty, filterValue);
                default:
                    return "Scope active";
            }
        }

        private static string BuildTreeScopeSummary(IReadOnlyList<List<string>> paths)
        {
            if (paths == null || paths.Count == 0)
            {
                return "Tree selection";
            }

            if (paths.Count == 1)
            {
                var path = paths[0] ?? new List<string>();
                return path.Count == 0 ? "Tree selection" : $"Tree selection: {string.Join(" > ", path)}";
            }

            return $"Tree selection ({paths.Count} items)";
        }

        private static string BuildFilterScopeSummary(string category, string property, string value)
        {
            if (string.IsNullOrWhiteSpace(category)
                || string.IsNullOrWhiteSpace(property)
                || string.IsNullOrWhiteSpace(value))
            {
                return "Property filter";
            }

            return $"Property filter: {category} / {property} = {value}";
        }

        private void RefreshScopeSelectionSetOptions()
        {
            ScopeSelectionSetOptions.Clear();

            var doc = NavisApp.ActiveDocument;
            foreach (var name in GetSavedSelectionSetNames(doc))
            {
                ScopeSelectionSetOptions.Add(name);
            }

            EnsureQuickColourScopeSelectionSet();
            EnsureHierarchyScopeSelectionSet();
        }

        private void EnsureQuickColourScopeSelectionSet()
        {
            if (QuickColourScope != QuickColourScope.SavedSelectionSet)
            {
                return;
            }

            if (ScopeSelectionSetOptions.Count == 0)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(QuickColourScopeSelectionSetName)
                || !ContainsIgnoreCase(ScopeSelectionSetOptions, QuickColourScopeSelectionSetName))
            {
                QuickColourScopeSelectionSetName = ScopeSelectionSetOptions.FirstOrDefault() ?? "";
            }
        }

        private void EnsureHierarchyScopeSelectionSet()
        {
            if (SelectedScope != QuickColourScope.SavedSelectionSet)
            {
                return;
            }

            if (ScopeSelectionSetOptions.Count == 0)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(HierarchyScopeSelectionSetName)
                || !ContainsIgnoreCase(ScopeSelectionSetOptions, HierarchyScopeSelectionSetName))
            {
                HierarchyScopeSelectionSetName = ScopeSelectionSetOptions.FirstOrDefault() ?? "";
            }
        }

        private void EnsureScopeSelectionSetOption(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            if (ContainsIgnoreCase(ScopeSelectionSetOptions, name))
            {
                return;
            }

            ScopeSelectionSetOptions.Add(name);
        }

        private static string BuildScopePath(
            QuickColourScope scope,
            string selectionSetName,
            IReadOnlyList<List<string>> modelPaths,
            string filterCategory,
            string filterProperty,
            string filterValue)
        {
            switch (scope)
            {
                case QuickColourScope.SavedSelectionSet:
                    return selectionSetName ?? "";
                case QuickColourScope.ModelTree:
                    return EncodeModelPaths(modelPaths);
                case QuickColourScope.PropertyFilter:
                    return EncodeScopeFilter(filterCategory, filterProperty, filterValue);
                default:
                    return "";
            }
        }

        private static string EncodeModelPaths(IReadOnlyList<List<string>> paths)
        {
            if (paths == null || paths.Count == 0)
            {
                return "";
            }

            var tokens = new List<string>();
            foreach (var path in paths)
            {
                if (path == null || path.Count == 0)
                {
                    continue;
                }

                tokens.Add(string.Join(">", path));
            }

            return string.Join("||", tokens);
        }

        private static List<List<string>> DecodeModelPaths(string path)
        {
            var results = new List<List<string>>();
            if (string.IsNullOrWhiteSpace(path))
            {
                return results;
            }

            var paths = path.Split(new[] { "||" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var raw in paths)
            {
                var segments = raw.Split(new[] { ">" }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .ToList();

                if (segments.Count > 0)
                {
                    results.Add(segments);
                }
            }

            return results;
        }

        private static string EncodeScopeFilter(string category, string property, string value)
        {
            return string.Join("||", new[]
            {
                category ?? "",
                property ?? "",
                value ?? ""
            });
        }

        private static void DecodeScopeFilter(string path, out string category, out string property, out string value)
        {
            category = "";
            property = "";
            value = "";

            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            var parts = path.Split(new[] { "||" }, StringSplitOptions.None);
            if (parts.Length > 0) category = parts[0];
            if (parts.Length > 1) property = parts[1];
            if (parts.Length > 2) value = parts[2];
        }

        private void TryPickProperty(bool isLevel1)
        {
            var session = GetCurrentSession();
            if (session == null)
            {
                StatusText = "No Data Scraper session available. Run Data Scraper first.";
                return;
            }

            var props = new DataScraperSessionAdapter(session).Properties.ToList();
            var vm = new PropertyPickerViewModel(props);
            var window = new PropertyPickerWindow(vm)
            {
                Owner = Window.GetWindow(this)
            };

            if (window.ShowDialog() == true && window.Selected != null)
            {
                if (isLevel1)
                {
                    HierarchyL1Category = window.Selected.Category ?? "";
                    HierarchyL1Property = window.Selected.Name ?? "";
                }
                else
                {
                    HierarchyL2Category = window.Selected.Category ?? "";
                    HierarchyL2Property = window.Selected.Name ?? "";
                }
            }
        }

        private void PickQuickColourScope_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                StatusText = "No active document.";
                return;
            }

            var session = GetCurrentSession();
            if (session == null)
            {
                StatusText = "No Data Scraper session available. Run Data Scraper first.";
                return;
            }

            var properties = new DataScraperSessionAdapter(session).Properties.ToList();
            var temp = new SmartSetRecipe
            {
                ScopeMode = ToSmartSetScopeMode(QuickColourScope),
                ScopeSelectionSetName = QuickColourScopeSelectionSetName ?? "",
                ScopeModelPaths = QuickColourScopeModelPaths?.Select(p => p?.ToList() ?? new List<string>()).ToList()
                    ?? new List<List<string>>(),
                ScopeFilterCategory = QuickColourScopeFilterCategory ?? "",
                ScopeFilterProperty = QuickColourScopeFilterProperty ?? "",
                ScopeFilterValue = QuickColourScopeFilterValue ?? ""
            };

            var picker = new SmartSetScopePickerWindow(doc, properties, temp)
            {
                Owner = Window.GetWindow(this)
            };

            if (picker.ShowDialog() == true && picker.Result != null)
            {
                ApplyQuickColourScopePickerResult(picker.Result);
            }
        }

        private void PickHierarchyScope_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                StatusText = "No active document.";
                return;
            }

            var session = GetCurrentSession();
            if (session == null)
            {
                StatusText = "No Data Scraper session available. Run Data Scraper first.";
                return;
            }

            var properties = new DataScraperSessionAdapter(session).Properties.ToList();
            var temp = new SmartSetRecipe
            {
                ScopeMode = ToSmartSetScopeMode(SelectedScope),
                ScopeSelectionSetName = HierarchyScopeSelectionSetName ?? "",
                ScopeModelPaths = HierarchyScopeModelPaths?.Select(p => p?.ToList() ?? new List<string>()).ToList()
                    ?? new List<List<string>>(),
                ScopeFilterCategory = HierarchyScopeFilterCategory ?? "",
                ScopeFilterProperty = HierarchyScopeFilterProperty ?? "",
                ScopeFilterValue = HierarchyScopeFilterValue ?? ""
            };

            var picker = new SmartSetScopePickerWindow(doc, properties, temp)
            {
                Owner = Window.GetWindow(this)
            };

            if (picker.ShowDialog() == true && picker.Result != null)
            {
                ApplyHierarchyScopePickerResult(picker.Result);
            }
        }

        private void ApplyQuickColourScopePickerResult(SmartSetScopePickerResult result)
        {
            if (result == null)
            {
                return;
            }

            QuickColourScopeSelectionSetName = "";
            QuickColourScopeModelPaths = result.ModelPaths ?? new List<List<string>>();
            QuickColourScopeFilterCategory = result.FilterCategory ?? "";
            QuickColourScopeFilterProperty = result.FilterProperty ?? "";
            QuickColourScopeFilterValue = result.FilterValue ?? "";
            QuickColourScope = ToQuickColourScope(result.ScopeMode);
        }

        private void ApplyHierarchyScopePickerResult(SmartSetScopePickerResult result)
        {
            if (result == null)
            {
                return;
            }

            HierarchyScopeSelectionSetName = "";
            HierarchyScopeModelPaths = result.ModelPaths ?? new List<List<string>>();
            HierarchyScopeFilterCategory = result.FilterCategory ?? "";
            HierarchyScopeFilterProperty = result.FilterProperty ?? "";
            HierarchyScopeFilterValue = result.FilterValue ?? "";
            SelectedScope = ToQuickColourScope(result.ScopeMode);
        }

        private void OpenDataScraper_Click(object sender, RoutedEventArgs e)
        {
            MicroEngActions.TryShowDataScraper(null, out _);
        }

        private void RefreshProfiles_Click(object sender, RoutedEventArgs e)
        {
            RefreshProfiles();
        }

        private void OnProfileApplyRequested(MicroEngColourProfile profile, MicroEngColourApplyMode mode)
        {
            if (profile == null) return;

            if (profile.Source == MicroEngColourProfileSource.HierarchyBuilder)
            {
                if (!LoadHierarchyProfile(profile, out var error))
                {
                    StatusText = error;
                    return;
                }

                _skipHierarchyRecolor = true;
                var l1Category = profile.Generator?.CategoryName ?? "";
                var l1Property = profile.Generator?.PropertyName ?? "";
                var notes = profile.Generator?.Notes ?? "";
                TryParseHierarchyNotes(notes, out var l2Category, out var l2Property);
                ApplyHierarchy(permanent: mode == MicroEngColourApplyMode.Permanent,
                    sourceButton: null,
                    l1CategoryOverride: l1Category,
                    l1PropertyOverride: l1Property,
                    l2CategoryOverride: l2Category,
                    l2PropertyOverride: l2Property,
                    profileOverride: profile);
                return;
            }

            LoadQuickColourProfile(profile);
            var category = profile.Generator?.CategoryName ?? "";
            var property = profile.Generator?.PropertyName ?? "";
            ApplyQuickColour(permanent: mode == MicroEngColourApplyMode.Permanent,
                sourceButton: null,
                categoryOverride: category,
                propertyOverride: property,
                profileOverride: profile);
        }

        private void SaveQuickColourProfile_Click(object sender, RoutedEventArgs e)
        {
            SaveQuickColourProfile(preserveId: true);
        }

        private void SaveQuickColourProfileAs_Click(object sender, RoutedEventArgs e)
        {
            SaveQuickColourProfile(preserveId: false);
        }

        private void SaveHierarchyProfile_Click(object sender, RoutedEventArgs e)
        {
            SaveHierarchyProfile(preserveId: true);
        }

        private void SaveHierarchyProfileAs_Click(object sender, RoutedEventArgs e)
        {
            SaveHierarchyProfile(preserveId: false);
        }

        private void SaveQuickColourProfile(bool preserveId)
        {
            if (QuickColourValues == null || QuickColourValues.Count == 0)
            {
                StatusText = "No values to save.";
                return;
            }

            if (string.IsNullOrWhiteSpace(QuickColourCategory) || string.IsNullOrWhiteSpace(QuickColourProperty))
            {
                StatusText = "Pick Category and Property first.";
                return;
            }

            var profile = BuildQuickColourProfile(preserveId);
            if (preserveId && !string.IsNullOrWhiteSpace(_lastQuickColourProfileId))
            {
                _colourProfileStore.Delete(_lastQuickColourProfileId);
            }

            _colourProfileStore.Save(profile);
            _lastQuickColourProfileId = profile.Id;
            RefreshProfilesPage();

            StatusText = $"Profile saved: {profile.Name}";
        }

        private void SaveHierarchyProfile(bool preserveId)
        {
            if (HierarchyGroups == null || HierarchyGroups.Count == 0)
            {
                StatusText = "No hierarchy loaded.";
                return;
            }

            if (string.IsNullOrWhiteSpace(_hierarchyL1Category)
                || string.IsNullOrWhiteSpace(_hierarchyL1Property)
                || string.IsNullOrWhiteSpace(_hierarchyL2Category)
                || string.IsNullOrWhiteSpace(_hierarchyL2Property))
            {
                StatusText = "Pick Level 1 and Level 2 properties first.";
                return;
            }

            var profile = BuildHierarchyProfile(preserveId);
            if (preserveId && !string.IsNullOrWhiteSpace(_lastHierarchyProfileId))
            {
                _colourProfileStore.Delete(_lastHierarchyProfileId);
            }

            _colourProfileStore.Save(profile);
            _lastHierarchyProfileId = profile.Id;
            RefreshProfilesPage();

            StatusText = $"Profile saved: {profile.Name}";
        }

        private void LoadQuickColourProfile(MicroEngColourProfile profile)
        {
            if (profile == null) return;

            _lastQuickColourProfileId = profile.Id;
            var generator = profile.Generator;

            QuickColourCategory = generator?.CategoryName ?? "";
            QuickColourProperty = generator?.PropertyName ?? "";
            var paletteKind = generator?.PaletteKind ?? MicroEngPaletteKind.Deep;
            paletteKind = ParsePaletteKind(generator?.PaletteName, paletteKind);
            QuickColourPaletteKind = paletteKind;
            QuickColourCustomBaseColorHex = generator?.CustomBaseColorHex ?? "#FF6699";
            QuickColourStableColors = generator?.StableColors ?? false;
            QuickColourStableSeed = generator?.Seed ?? "";
            QuickColourCreateSearchSets = profile.Outputs?.CreateSearchSets ?? false;
            QuickColourCreateSnapshots = profile.Outputs?.CreateSnapshots ?? false;
            QuickColourFolderPath = profile.Outputs?.FolderPath ?? "MicroEng/Quick Colour";
            QuickColourProfileName = string.IsNullOrWhiteSpace(profile.Name) ? "Quick Colour" : profile.Name;

            ApplyQuickColourScopeFromProfile(profile.Scope);

            EnsureQuickColourCategoryOption(QuickColourCategory);
            UpdateQuickColourPropertyOptionsFor(QuickColourCategory, QuickColourPropertyOptions, QuickColourProperty);

            QuickColourValues.Clear();
            foreach (var rule in profile.Rules ?? new List<MicroEngColourRule>())
            {
                var row = new QuickColourValueRow
                {
                    Value = rule.Value ?? "",
                    Enabled = rule.Enabled,
                    Count = rule.Count ?? 0
                };
                row.ColorHex = rule.ColorHex ?? "#D3D3D3";
                QuickColourValues.Add(row);
            }
        }

        private void ApplyQuickColourScopeFromProfile(MicroEngColourScope scope)
        {
            var parsed = ParseScope(scope?.Kind);
            var path = scope?.Path ?? "";

            QuickColourScopeSelectionSetName = "";
            QuickColourScopeModelPaths = new List<List<string>>();
            QuickColourScopeFilterCategory = "";
            QuickColourScopeFilterProperty = "";
            QuickColourScopeFilterValue = "";

            switch (parsed)
            {
                case QuickColourScope.SavedSelectionSet:
                    QuickColourScopeSelectionSetName = path;
                    break;
                case QuickColourScope.ModelTree:
                    QuickColourScopeModelPaths = DecodeModelPaths(path);
                    break;
                case QuickColourScope.PropertyFilter:
                    DecodeScopeFilter(path, out var cat, out var prop, out var val);
                    QuickColourScopeFilterCategory = cat;
                    QuickColourScopeFilterProperty = prop;
                    QuickColourScopeFilterValue = val;
                    break;
            }

            QuickColourScope = parsed;
        }

        private bool LoadHierarchyProfile(MicroEngColourProfile profile, out string error)
        {
            error = "";

            _lastHierarchyProfileId = profile.Id;
            var l1Category = profile.Generator?.CategoryName ?? "";
            var l1Property = profile.Generator?.PropertyName ?? "";
            var l2Category = profile.Generator?.L2CategoryName ?? "";
            var l2Property = profile.Generator?.L2PropertyName ?? "";
            var notes = profile.Generator?.Notes ?? "";

            if (string.IsNullOrWhiteSpace(l2Category) || string.IsNullOrWhiteSpace(l2Property))
            {
                if (!TryParseHierarchyNotes(notes, out l2Category, out l2Property))
                {
                    error = "Hierarchy profile missing Level 2 property info.";
                    return false;
                }
            }

            if (string.IsNullOrWhiteSpace(l1Category) || string.IsNullOrWhiteSpace(l1Property))
            {
                error = "Hierarchy profile missing Level 1 property info.";
                return false;
            }

            HierarchyL1Category = l1Category;
            HierarchyL1Property = l1Property;
            HierarchyL2Category = l2Category;
            HierarchyL2Property = l2Property;
            ApplyHierarchyScopeFromProfile(profile.Scope);
            var paletteKind = profile.Generator?.PaletteKind
                              ?? ParsePaletteKind(profile.Generator?.PaletteName, HierarchyPaletteKind);
            HierarchyPaletteKind = paletteKind;

            CreateSearchSets = profile.Outputs?.CreateSearchSets ?? false;
            CreateSnapshots = profile.Outputs?.CreateSnapshots ?? false;
            OutputFolderPath = profile.Outputs?.FolderPath ?? "MicroEng/Quick Colour";
            OutputProfileName = string.IsNullOrWhiteSpace(profile.Name) ? "Quick Colour" : profile.Name;

            EnsureCategoryOption(l1Category);
            EnsureCategoryOption(l2Category);
            UpdateHierarchyPropertyOptions();

            HierarchyGroups.Clear();

            var groups = new Dictionary<string, QuickColourHierarchyGroup>(StringComparer.OrdinalIgnoreCase);
            foreach (var rule in profile.Rules ?? new List<MicroEngColourRule>())
            {
                if (!TrySplitHierarchyValue(rule.Value, out var l1Value, out var l2Value))
                {
                    continue;
                }

                if (!groups.TryGetValue(l1Value, out var group))
                {
                    group = new QuickColourHierarchyGroup
                    {
                        Value = l1Value,
                        Enabled = true
                    };
                    groups[l1Value] = group;
                }

                var typeRow = new QuickColourHierarchyTypeRow
                {
                    Value = l2Value,
                    Enabled = rule.Enabled,
                    Count = rule.Count ?? 0
                };
                typeRow.ColorHex = rule.ColorHex ?? "#D3D3D3";
                group.Types.Add(typeRow);
            }

            foreach (var group in groups.Values.OrderBy(g => g.Value, StringComparer.OrdinalIgnoreCase))
            {
                group.Count = group.Types.Sum(t => t.Count);
                var first = group.Types.FirstOrDefault();
                if (first != null)
                {
                    group.SetComputedBaseColor(first.Color);
                }
                HierarchyGroups.Add(group);
            }

            if (HierarchyGroups.Count == 0)
            {
                error = "Hierarchy profile has no rules.";
                return false;
            }

            return true;
        }

        private void ApplyHierarchyScopeFromProfile(MicroEngColourScope scope)
        {
            var parsed = ParseScope(scope?.Kind);
            var path = scope?.Path ?? "";

            HierarchyScopeSelectionSetName = "";
            HierarchyScopeModelPaths = new List<List<string>>();
            HierarchyScopeFilterCategory = "";
            HierarchyScopeFilterProperty = "";
            HierarchyScopeFilterValue = "";

            switch (parsed)
            {
                case QuickColourScope.SavedSelectionSet:
                    HierarchyScopeSelectionSetName = path;
                    break;
                case QuickColourScope.ModelTree:
                    HierarchyScopeModelPaths = DecodeModelPaths(path);
                    break;
                case QuickColourScope.PropertyFilter:
                    DecodeScopeFilter(path, out var cat, out var prop, out var val);
                    HierarchyScopeFilterCategory = cat;
                    HierarchyScopeFilterProperty = prop;
                    HierarchyScopeFilterValue = val;
                    break;
            }

            SelectedScope = parsed;
        }

        private MicroEngColourProfile BuildQuickColourProfile(bool preserveId)
        {
            var name = string.IsNullOrWhiteSpace(QuickColourProfileName) ? "Quick Colour" : QuickColourProfileName.Trim();

            var profile = new MicroEngColourProfile
            {
                Id = preserveId && !string.IsNullOrWhiteSpace(_lastQuickColourProfileId)
                    ? _lastQuickColourProfileId
                    : Guid.NewGuid().ToString("N"),
                Name = name,
                Source = MicroEngColourProfileSource.QuickColour,
                Scope = new MicroEngColourScope
                {
                    Kind = BuildScopeKind(QuickColourScope),
                    Path = BuildScopePath(
                        QuickColourScope,
                        QuickColourScopeSelectionSetName,
                        QuickColourScopeModelPaths,
                        QuickColourScopeFilterCategory,
                        QuickColourScopeFilterProperty,
                        QuickColourScopeFilterValue)
                },
                Generator = new MicroEngColourGenerator
                {
                    CategoryName = QuickColourCategory ?? "",
                    PropertyName = QuickColourProperty ?? "",
                    PaletteName = QuickColourPaletteKind.ToString(),
                    PaletteKind = QuickColourPaletteKind,
                    CustomBaseColorHex = QuickColourCustomBaseColorHex ?? "#FF6699",
                    StableColors = QuickColourStableColors,
                    Seed = QuickColourStableSeed ?? "",
                    Notes = ""
                },
                Outputs = new MicroEngColourOutputs
                {
                    CreateSearchSets = QuickColourCreateSearchSets,
                    CreateSnapshots = QuickColourCreateSnapshots,
                    FolderPath = QuickColourFolderPath ?? "MicroEng/Quick Colour",
                    SetNamePrefix = ""
                }
            };

            foreach (var row in QuickColourValues)
            {
                profile.Rules.Add(new MicroEngColourRule
                {
                    Enabled = row.Enabled,
                    Value = row.Value ?? "",
                    ColorHex = row.ColorHex ?? "#D3D3D3",
                    Count = row.Count
                });
            }

            return profile;
        }

        private MicroEngColourProfile BuildHierarchyProfile(bool preserveId)
        {
            var name = string.IsNullOrWhiteSpace(OutputProfileName) ? "Quick Colour" : OutputProfileName.Trim();
            var notes = BuildHierarchyNotes(_hierarchyL2Category, _hierarchyL2Property);

            var profile = new MicroEngColourProfile
            {
                Id = preserveId && !string.IsNullOrWhiteSpace(_lastHierarchyProfileId)
                    ? _lastHierarchyProfileId
                    : Guid.NewGuid().ToString("N"),
                Name = name,
                Source = MicroEngColourProfileSource.HierarchyBuilder,
                Scope = new MicroEngColourScope
                {
                    Kind = BuildScopeKind(SelectedScope),
                    Path = BuildScopePath(
                        SelectedScope,
                        HierarchyScopeSelectionSetName,
                        HierarchyScopeModelPaths,
                        HierarchyScopeFilterCategory,
                        HierarchyScopeFilterProperty,
                        HierarchyScopeFilterValue)
                },
                Generator = new MicroEngColourGenerator
                {
                    CategoryName = _hierarchyL1Category ?? "",
                    PropertyName = _hierarchyL1Property ?? "",
                    L2CategoryName = _hierarchyL2Category ?? "",
                    L2PropertyName = _hierarchyL2Property ?? "",
                    PaletteName = HierarchyPaletteKind.ToString(),
                    PaletteKind = HierarchyPaletteKind,
                    StableColors = false,
                    Seed = "",
                    Notes = notes
                },
                Outputs = new MicroEngColourOutputs
                {
                    CreateSearchSets = CreateSearchSets,
                    CreateSnapshots = CreateSnapshots,
                    FolderPath = OutputFolderPath ?? "MicroEng/Quick Colour",
                    SetNamePrefix = ""
                }
            };

            foreach (var group in HierarchyGroups.Where(g => g != null))
            {
                foreach (var type in group.Types.Where(t => t != null))
                {
                    profile.Rules.Add(new MicroEngColourRule
                    {
                        Enabled = group.Enabled && type.Enabled,
                        Value = BuildHierarchyValue(group.Value, type.Value),
                        ColorHex = type.ColorHex ?? "#D3D3D3",
                        Count = type.Count
                    });
                }
            }

            return profile;
        }

        private static string BuildHierarchyNotes(string l2Category, string l2Property)
        {
            l2Category ??= "";
            l2Property ??= "";
            return $"L2={l2Category}|{l2Property}";
        }

        private static bool TryParseHierarchyNotes(string notes, out string l2Category, out string l2Property)
        {
            l2Category = "";
            l2Property = "";

            if (string.IsNullOrWhiteSpace(notes))
            {
                return false;
            }

            notes = notes.Trim();
            if (!notes.StartsWith("L2=", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var payload = notes.Substring(3);
            var parts = payload.Split(new[] { '|' }, 2);
            if (parts.Length != 2)
            {
                return false;
            }

            l2Category = parts[0].Trim();
            l2Property = parts[1].Trim();
            return !string.IsNullOrWhiteSpace(l2Category) && !string.IsNullOrWhiteSpace(l2Property);
        }

        private static string BuildHierarchyValue(string l1Value, string l2Value)
        {
            return $"{l1Value ?? ""}||{l2Value ?? ""}";
        }

        private static bool TrySplitHierarchyValue(string value, out string l1Value, out string l2Value)
        {
            l1Value = "";
            l2Value = "";

            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var parts = value.Split(new[] { "||" }, StringSplitOptions.None);
            if (parts.Length < 2)
            {
                return false;
            }

            l1Value = parts[0];
            l2Value = parts[1];
            return true;
        }

        private void RefreshProfilesPage()
        {
            ProfilesPage?.Refresh();
        }

        private void LoadHierarchy_Click(object sender, RoutedEventArgs e)
        {
            LoadHierarchy();
        }

        private void LoadHierarchy()
        {
            try
            {
                var session = GetCurrentSession();
                if (session == null)
                {
                    StatusText = "No Data Scraper session available. Run Data Scraper first.";
                    return;
                }

                if (string.IsNullOrWhiteSpace(_hierarchyL1Category)
                    || string.IsNullOrWhiteSpace(_hierarchyL1Property)
                    || string.IsNullOrWhiteSpace(_hierarchyL2Category)
                    || string.IsNullOrWhiteSpace(_hierarchyL2Property))
                {
                    StatusText = "Pick Level 1 and Level 2 properties first.";
                    return;
                }

                if (!TryGetScopeItemKeys(
                    session,
                    SelectedScope,
                    HierarchyScopeSelectionSetName,
                    HierarchyScopeModelPaths,
                    HierarchyScopeFilterCategory,
                    HierarchyScopeFilterProperty,
                    HierarchyScopeFilterValue,
                    out var allowedKeys,
                    out var scopeError))
                {
                    StatusText = scopeError;
                    return;
                }

                var l1 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var l2 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                foreach (var entry in session.RawEntries ?? Enumerable.Empty<RawEntry>())
                {
                    if (entry == null)
                    {
                        continue;
                    }

                    if (allowedKeys != null && !allowedKeys.Contains(entry.ItemKey ?? ""))
                    {
                        continue;
                    }

                    var itemKey = entry.ItemKey;
                    if (string.IsNullOrWhiteSpace(itemKey))
                    {
                        continue;
                    }

                    if (IsMatch(entry, _hierarchyL1Category, _hierarchyL1Property))
                    {
                        l1[itemKey] = entry.Value ?? "";
                    }

                    if (IsMatch(entry, _hierarchyL2Category, _hierarchyL2Property))
                    {
                        l2[itemKey] = entry.Value ?? "";
                    }
                }

                var pairs = new List<(string L1, string L2)>();
                foreach (var kvp in l1)
                {
                    if (!l2.TryGetValue(kvp.Key, out var l2Value))
                    {
                        continue;
                    }

                    var l1Value = kvp.Value ?? "";
                    l2Value = l2Value ?? "";

                    if (!HierarchyIncludeBlanks)
                    {
                        if (string.IsNullOrWhiteSpace(l1Value) || string.IsNullOrWhiteSpace(l2Value))
                        {
                            continue;
                        }
                    }

                    pairs.Add((l1Value, l2Value));
                }

                var maxCategories = HierarchyMaxCategories > 0 ? HierarchyMaxCategories : int.MaxValue;
                var maxTypes = HierarchyMaxTypesPerCategory > 0 ? HierarchyMaxTypesPerCategory : int.MaxValue;

                var groups = pairs
                    .GroupBy(p => p.L1 ?? "", StringComparer.OrdinalIgnoreCase)
                    .Select(g => new
                    {
                        Value = g.Key ?? "",
                        Count = g.Count(),
                        Types = g.GroupBy(x => x.L2 ?? "", StringComparer.OrdinalIgnoreCase)
                                 .Select(t => new { Value = t.Key ?? "", Count = t.Count() })
                                 .OrderByDescending(t => t.Count)
                                 .ThenBy(t => t.Value, StringComparer.OrdinalIgnoreCase)
                                 .Take(maxTypes)
                                 .ToList()
                    })
                    .OrderByDescending(g => g.Count)
                    .ThenBy(g => g.Value, StringComparer.OrdinalIgnoreCase)
                    .Take(maxCategories)
                    .ToList();

                HierarchyGroups.Clear();
                foreach (var g in groups)
                {
                    var group = new QuickColourHierarchyGroup
                    {
                        Value = g.Value,
                        Count = g.Count
                    };

                    foreach (var t in g.Types)
                    {
                        group.Types.Add(new QuickColourHierarchyTypeRow
                        {
                            Value = t.Value,
                            Count = t.Count
                        });
                    }

                    HierarchyGroups.Add(group);
                }

                EnsureCategoryHueGroupDefaults();
                AutoAssignBaseColours();
                ApplyHierarchyColours();

                StatusText = $"Loaded {HierarchyGroups.Count} categories.";
            }
            catch (Exception ex)
            {
                StatusText = "Load hierarchy failed: " + ex.Message;
            }
        }

        private void AutoAssignBaseColours_Click(object sender, RoutedEventArgs e)
        {
            AutoAssignBaseColours();
        }

        private void AutoAssignBaseColours()
        {
            if (HierarchyGroups == null || HierarchyGroups.Count == 0)
            {
                return;
            }

            var groups = HierarchyGroups
                .Where(g => g != null && g.Enabled)
                .OrderByDescending(g => g.Count)
                .ThenBy(g => g.Value, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var palette = GetHierarchyBasePalette(groups.Count);
            for (int i = 0; i < groups.Count; i++)
            {
                var group = groups[i];
                if (group.UseCustomBaseColor)
                {
                    continue;
                }

                var color = palette.Count > i ? palette[i] : Colors.LightGray;
                group.SetComputedBaseColor(color);
            }

            if (!_skipHierarchyRecolor)
            {
                ApplyHierarchyColours();
            }
            _skipHierarchyRecolor = false;
        }

        private List<System.Windows.Media.Color> GetHierarchyBasePalette(int count)
        {
            if (count <= 0)
            {
                return new List<System.Windows.Media.Color>();
            }

            if (HierarchyPaletteKind == MicroEngPaletteKind.Shades)
            {
                var baseColor = ColourPaletteGenerator.ParseHexOrDefault(HierarchyHueHex, Colors.HotPink);
                return QuickColourPalette.GenerateShades(baseColor, count, GetHierarchyPaletteStyle(), 0.85);
            }

            if (HierarchyPaletteKind == MicroEngPaletteKind.CustomHue)
            {
                var baseStyle = QuickColourPaletteStyle.Deep;
                var palette = QuickColourPalette.GeneratePalette(count, baseStyle);
                var baseColor = ColourPaletteGenerator.ParseHexOrDefault(HierarchyHueHex, Colors.HotPink);
                ColourPaletteGenerator.RgbToHsl(baseColor, out _, out _, out var targetL);
                var baseL = baseStyle == QuickColourPaletteStyle.Pastel ? 0.78 : 0.55;
                var deltaL = targetL - baseL;

                for (int i = 0; i < palette.Count; i++)
                {
                    ColourPaletteGenerator.RgbToHsl(palette[i], out var h, out var s, out var l);
                    var newL = Math.Max(0.18, Math.Min(0.92, l + deltaL));
                    palette[i] = ColourPaletteGenerator.HslToRgb(h, s, newL);
                }

                return palette;
            }

            return QuickColourPalette.GeneratePalette(count, HierarchyPaletteKind);
        }

        private QuickColourPaletteStyle GetHierarchyPaletteStyle()
        {
            return HierarchyPaletteKind == MicroEngPaletteKind.Pastel
                ? QuickColourPaletteStyle.Pastel
                : QuickColourPaletteStyle.Deep;
        }

        private void RegenerateTypeShades_Click(object sender, RoutedEventArgs e)
        {
            ApplyHierarchyColours();
        }

        private void ApplyTemporary_Click(object sender, RoutedEventArgs e)
        {
            ApplyHierarchy(permanent: false, sender as System.Windows.Controls.Button);
        }

        private void ApplyPermanent_Click(object sender, RoutedEventArgs e)
        {
            ApplyHierarchy(permanent: true, sender as System.Windows.Controls.Button);
        }

        private void ClearTemporary_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                StatusText = "No active document.";
                return;
            }

            doc.Models.ResetAllTemporaryMaterials();
            StatusText = "Temporary colors cleared.";
        }

        private void ResetPermanent_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                StatusText = "No active document.";
                return;
            }

            doc.Models.ResetAllPermanentMaterials();
            StatusText = "Permanent colors reset.";
        }

        private void ApplyHierarchy(bool permanent,
            System.Windows.Controls.Button sourceButton = null,
            string l1CategoryOverride = null,
            string l1PropertyOverride = null,
            string l2CategoryOverride = null,
            string l2PropertyOverride = null,
            MicroEngColourProfile profileOverride = null)
        {
            var skipRecolor = _skipHierarchyRecolor;
            _skipHierarchyRecolor = false;

            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                StatusText = "No active document.";
                return;
            }

            if (HierarchyGroups == null || HierarchyGroups.Count == 0)
            {
                StatusText = "No hierarchy loaded.";
                return;
            }

            var l1Category = string.IsNullOrWhiteSpace(l1CategoryOverride) ? _hierarchyL1Category : l1CategoryOverride;
            var l1Property = string.IsNullOrWhiteSpace(l1PropertyOverride) ? _hierarchyL1Property : l1PropertyOverride;
            var l2Category = string.IsNullOrWhiteSpace(l2CategoryOverride) ? _hierarchyL2Category : l2CategoryOverride;
            var l2Property = string.IsNullOrWhiteSpace(l2PropertyOverride) ? _hierarchyL2Property : l2PropertyOverride;

            if (string.IsNullOrWhiteSpace(l1Category)
                || string.IsNullOrWhiteSpace(l1Property)
                || string.IsNullOrWhiteSpace(l2Category)
                || string.IsNullOrWhiteSpace(l2Property))
            {
                StatusText = profileOverride != null
                    ? "Profile is missing Level 1/Level 2 property info."
                    : "Pick Level 1 and Level 2 properties first.";
                return;
            }

            if (!ValidateScopeForApply(
                SelectedScope,
                HierarchyScopeSelectionSetName,
                HierarchyScopeModelPaths,
                HierarchyScopeFilterCategory,
                HierarchyScopeFilterProperty,
                HierarchyScopeFilterValue,
                out var scopeError))
            {
                StatusText = scopeError;
                return;
            }

            if (!skipRecolor)
            {
                ApplyHierarchyColours();
            }

            try
            {
                var expected = HierarchyGroups
                    .Where(g => g != null && g.Enabled)
                    .SelectMany(g => g.Types.Where(t => t != null && t.Enabled))
                    .Sum(t => Math.Max(0, t.Count));
                if (IsQuickColourTraceEnabled())
                {
                    Log($"QuickColour(Hierarchy): expected count from cache={expected} for {l1Category}.{l1Property} -> {l2Category}.{l2Property}.");
                }

                var count = _service.ApplyByHierarchy(
                    doc,
                    l1Category,
                    l1Property,
                    l2Category,
                    l2Property,
                    HierarchyGroups,
                    SelectedScope,
                    HierarchyScopeSelectionSetName,
                    HierarchyScopeModelPaths,
                    HierarchyScopeFilterCategory,
                    HierarchyScopeFilterProperty,
                    HierarchyScopeFilterValue,
                    permanent,
                    CreateSearchSets,
                    CreateSnapshots,
                    OutputFolderPath,
                    OutputProfileName,
                    CreateFoldersByHueGroup && HierarchyUseHueGroups,
                    Log);

                Log($"QuickColour(Hierarchy): cached entries={expected}, coloured items={count}.");

                StatusText = $"Applied {count} item(s).";
                FlashSuccess(sourceButton);
                ShowSnackbar(permanent ? "Applied permanent hierarchy" : "Applied temporary hierarchy",
                    $"Applied to {count} item(s).",
                    WpfUiControls.ControlAppearance.Success,
                    WpfUiControls.SymbolRegular.CheckmarkCircle24);
            }
            catch (Exception ex)
            {
                StatusText = "Apply failed: " + ex.Message;
                ShowSnackbar("Apply failed",
                    ex.Message,
                    WpfUiControls.ControlAppearance.Danger,
                    WpfUiControls.SymbolRegular.ErrorCircle24);
            }
        }
        private void ApplyHierarchyColours()
        {
            if (_isApplyingHierarchyColours)
            {
                return;
            }

            if (HierarchyGroups == null || HierarchyGroups.Count == 0)
            {
                return;
            }

            _isApplyingHierarchyColours = true;
            try
            {
                if (HierarchySingleHueMode)
                {
                    ApplySingleHueColours();
                    return;
                }

                ApplyBasePaletteColours();
            }
            finally
            {
                _isApplyingHierarchyColours = false;
            }
        }

        private void ApplyBasePaletteColours()
        {
            var style = GetHierarchyPaletteStyle();
            var typeUsage = Math.Max(0.15, GetTypeSpread01());

            foreach (var group in HierarchyGroups.Where(g => g != null && g.Enabled))
            {
                var types = GetSortedTypes(group);
                if (types.Count == 0)
                {
                    continue;
                }

                if (!HierarchyShadeHueGroups)
                {
                    foreach (var type in types)
                    {
                        type.Color = group.BaseColor;
                    }
                    continue;
                }

                var shades = QuickColourPalette.GenerateShades(group.BaseColor, types.Count, style, typeUsage);
                for (int i = 0; i < types.Count; i++)
                {
                    types[i].Color = shades[i];
                }
            }
        }

        private void ApplySingleHueColours()
        {
            var style = GetHierarchyPaletteStyle();
            var typeUsage = Math.Max(0.15, GetTypeSpread01());
            var hueColor = ParseHueColor();
            var hue01 = QuickColourPalette.GetHue01(hueColor);
            var sat = QuickColourPalette.GetRecommendedSaturation(style);

            var contrast01 = Clamp01(HierarchyCategoryContrastPct / 100.0);
            var (minL, maxL) = QuickColourPalette.ComputeCategoryLightnessRange(style, contrast01);

            var groups = HierarchyGroups
                .Where(g => g != null && g.Enabled)
                .OrderByDescending(g => g.Count)
                .ThenBy(g => g.Value, StringComparer.OrdinalIgnoreCase)
                .ToList();

            int n = groups.Count;
            if (n == 0)
            {
                return;
            }

            double bandWidth = (maxL - minL) / n;
            double pad = bandWidth * 0.08;

            for (int i = 0; i < n; i++)
            {
                var group = groups[i];
                double bandMin = minL + bandWidth * i + pad;
                double bandMax = minL + bandWidth * (i + 1) - pad;

                if (bandMax <= bandMin)
                {
                    bandMin = minL + bandWidth * i;
                    bandMax = minL + bandWidth * (i + 1);
                }

                double baseL = (bandMin + bandMax) / 2.0;

                if (!group.UseCustomBaseColor)
                {
                    group.SetComputedBaseColor(QuickColourPalette.FromHsl01(hue01, sat, baseL));
                }

                var types = GetSortedTypes(group);
                if (types.Count == 0)
                {
                    continue;
                }

                if (!HierarchyShadeHueGroups)
                {
                    foreach (var type in types)
                    {
                        type.Color = group.BaseColor;
                    }
                    continue;
                }

                double half = ((bandMax - bandMin) / 2.0) * typeUsage;
                double effMin = Math.Max(bandMin, baseL - half);
                double effMax = Math.Min(bandMax, baseL + half);

                var ramp = QuickColourPalette.GenerateHslRamp(hue01, sat, effMin, effMax, types.Count);
                for (int t = 0; t < types.Count; t++)
                {
                    types[t].Color = ramp[t];
                }
            }
        }

        private List<QuickColourHierarchyTypeRow> GetSortedTypes(QuickColourHierarchyGroup group)
        {
            var list = group.Types
                .Where(t => t != null && t.Enabled)
                .ToList();

            switch (HierarchyTypeSortMode)
            {
                case QuickColourTypeSortMode.Name:
                    return list.OrderBy(t => t.Value, StringComparer.OrdinalIgnoreCase).ToList();
                case QuickColourTypeSortMode.Count:
                default:
                    return list.OrderByDescending(t => t.Count)
                        .ThenBy(t => t.Value, StringComparer.OrdinalIgnoreCase)
                        .ToList();
            }
        }

        private double GetTypeSpread01()
        {
            var v = HierarchyShadeSpreadPct / 100.0;
            if (v < 0) v = 0;
            if (v > 1) v = 1;
            return v;
        }

        private static double Clamp01(double x)
        {
            if (x < 0) return 0;
            if (x > 1) return 1;
            return x;
        }

        private System.Windows.Media.Color ParseHueColor()
        {
            return QuickColourPalette.TryParseHex(HierarchyHueHex, out var c)
                ? c
                : Colors.Green;
        }

        private void PickHierarchyHue_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var dlg = new System.Windows.Forms.ColorDialog())
                {
                    dlg.FullOpen = true;

                    if (QuickColourPalette.TryParseHex(HierarchyHueHex, out var current))
                    {
                        dlg.Color = System.Drawing.Color.FromArgb(current.R, current.G, current.B);
                    }

                    if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        HierarchyHueHex = $"#{dlg.Color.R:X2}{dlg.Color.G:X2}{dlg.Color.B:X2}";
                    }
                }
            }
            catch (Exception ex)
            {
                StatusText = "Pick hue failed: " + ex.Message;
            }
        }

        private void EnsureDefaultHierarchyProperties()
        {
            var session = GetCurrentSession();
            if (session == null)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(HierarchyL1Category) || string.IsNullOrWhiteSpace(HierarchyL1Property))
            {
                if (TryPickProperty(session, "Element", "Category", out var cat, out var prop)
                    || TryPickPropertyByName(session, "Category", out cat, out prop))
                {
                    HierarchyL1Category = cat;
                    HierarchyL1Property = prop;
                }
            }

            if (string.IsNullOrWhiteSpace(HierarchyL2Category) || string.IsNullOrWhiteSpace(HierarchyL2Property))
            {
                if (TryPickProperty(session, "Element", "Type", out var cat, out var prop)
                    || TryPickProperty(session, "Element", "Type Name", out cat, out prop)
                    || TryPickPropertyByName(session, "Type", out cat, out prop))
                {
                    HierarchyL2Category = cat;
                    HierarchyL2Property = prop;
                }
            }
        }

        private static bool TryPickProperty(ScrapeSession session, string category, string property, out string cat, out string prop)
        {
            cat = "";
            prop = "";

            if (session?.Properties == null)
            {
                return false;
            }

            foreach (var p in session.Properties)
            {
                if (p == null) continue;
                if (string.Equals(p.Category, category, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(p.Name, property, StringComparison.OrdinalIgnoreCase))
                {
                    cat = p.Category ?? "";
                    prop = p.Name ?? "";
                    return true;
                }
            }

            return false;
        }

        private static bool TryPickPropertyByName(ScrapeSession session, string nameToken, out string cat, out string prop)
        {
            cat = "";
            prop = "";

            if (session?.Properties == null)
            {
                return false;
            }

            var match = session.Properties
                .FirstOrDefault(p => p != null && (p.Name ?? "").IndexOf(nameToken, StringComparison.OrdinalIgnoreCase) >= 0);

            if (match == null)
            {
                return false;
            }

            cat = match.Category ?? "";
            prop = match.Name ?? "";
            return true;
        }

        private static bool IsMatch(RawEntry entry, string category, string property)
        {
            if (entry == null) return false;

            if (!string.Equals(entry.Category ?? "", category ?? "", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return string.Equals(entry.Name ?? "", property ?? "", StringComparison.OrdinalIgnoreCase);
        }

        private static string FormatPropertyLabel(string category, string property)
        {
            if (string.IsNullOrWhiteSpace(category) && string.IsNullOrWhiteSpace(property))
            {
                return "Pick property...";
            }

            if (string.IsNullOrWhiteSpace(category))
            {
                return property ?? "";
            }

            if (string.IsNullOrWhiteSpace(property))
            {
                return category ?? "";
            }

            return $"{category} > {property}";
        }

        private static bool ContainsIgnoreCase(IEnumerable<string> items, string value)
        {
            if (items == null || string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return items.Any(item => string.Equals(item, value, StringComparison.OrdinalIgnoreCase));
        }

        private static HashSet<string> CollectItemKeys(IEnumerable<ModelItem> items)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (items == null)
            {
                return keys;
            }

            foreach (var item in items)
            {
                CollectItemKeys(item, keys);
            }

            return keys;
        }

        private static void CollectItemKeys(ModelItem item, HashSet<string> keys)
        {
            if (item == null || keys == null)
            {
                return;
            }

            var id = item.InstanceGuid;
            if (id != Guid.Empty)
            {
                keys.Add(id.ToString());
            }

            var children = item.Children;
            if (children == null)
            {
                return;
            }

            foreach (var child in children)
            {
                CollectItemKeys(child, keys);
            }
        }

        private static HashSet<string> BuildPropertyFilterKeys(
            ScrapeSession session,
            string category,
            string property,
            string value)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (session?.RawEntries == null)
            {
                return keys;
            }

            foreach (var entry in session.RawEntries)
            {
                if (entry == null)
                {
                    continue;
                }

                if (!string.Equals(entry.Category ?? "", category ?? "", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.Equals(entry.Name ?? "", property ?? "", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.Equals(entry.Value ?? "", value ?? "", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(entry.ItemKey))
                {
                    keys.Add(entry.ItemKey);
                }
            }

            return keys;
        }

        private static IReadOnlyList<string> GetSavedSelectionSetNames(Document doc)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var root = doc?.SelectionSets?.RootItem;
            CollectSelectionSetNames(root, names);

            return names.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static void CollectSelectionSetNames(Autodesk.Navisworks.Api.GroupItem folder, HashSet<string> names)
        {
            if (folder?.Children == null)
            {
                return;
            }

            foreach (SavedItem item in folder.Children)
            {
                if (item is SelectionSet set)
                {
                    var label = set.DisplayName ?? "";
                    if (!string.IsNullOrWhiteSpace(label))
                    {
                        names.Add(label);
                    }
                }
                else if (item is Autodesk.Navisworks.Api.GroupItem group)
                {
                    CollectSelectionSetNames(group, names);
                }
            }
        }

        private static ModelItemCollection GetSelectionSetItems(Document doc, string name)
        {
            try
            {
                var root = doc?.SelectionSets?.RootItem;
                var selectionSet = FindSelectionSetByName(root, name);
                return selectionSet?.GetSelectedItems(doc);
            }
            catch
            {
                return null;
            }
        }

        private static SelectionSet FindSelectionSetByName(Autodesk.Navisworks.Api.GroupItem folder, string name)
        {
            if (folder == null || string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            var children = folder.Children;
            if (children == null)
            {
                return null;
            }

            foreach (SavedItem item in children)
            {
                if (item is SelectionSet set)
                {
                    if (string.Equals(set.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return set;
                    }
                }
                else if (item is Autodesk.Navisworks.Api.GroupItem group)
                {
                    var found = FindSelectionSetByName(group, name);
                    if (found != null)
                    {
                        return found;
                    }
                }
            }

            return null;
        }

        private static ModelItemCollection ResolveScopeModelItems(Document doc, IReadOnlyList<List<string>> paths)
        {
            if (doc?.Models?.RootItems == null || paths == null || paths.Count == 0)
            {
                return null;
            }

            var items = new ModelItemCollection();
            var added = new HashSet<Guid>();
            foreach (var path in paths)
            {
                var resolved = ResolveModelPath(doc.Models.RootItems, path);
                if (resolved == null)
                {
                    continue;
                }

                var id = resolved.InstanceGuid;
                if (id == Guid.Empty || added.Add(id))
                {
                    items.Add(resolved);
                }
            }

            return items;
        }

        private static ModelItem ResolveModelPath(IEnumerable<ModelItem> roots, IReadOnlyList<string> path)
        {
            if (roots == null || path == null || path.Count == 0)
            {
                return null;
            }

            var current = FindRootByLabel(roots, path[0]);
            if (current == null)
            {
                return null;
            }

            for (var i = 1; i < path.Count; i++)
            {
                var segment = path[i];
                if (string.IsNullOrWhiteSpace(segment))
                {
                    return null;
                }

                var next = FindChildByLabel(current, segment);
                if (next == null)
                {
                    return null;
                }

                current = next;
            }

            return current;
        }

        private static ModelItem FindRootByLabel(IEnumerable<ModelItem> roots, string label)
        {
            if (roots == null || string.IsNullOrWhiteSpace(label))
            {
                return null;
            }

            foreach (var root in roots)
            {
                var name = root?.DisplayName ?? "";
                if (string.Equals(name, label, StringComparison.OrdinalIgnoreCase))
                {
                    return root;
                }
            }

            return null;
        }

        private static ModelItem FindChildByLabel(ModelItem parent, string label)
        {
            if (parent?.Children == null || string.IsNullOrWhiteSpace(label))
            {
                return null;
            }

            foreach (var child in parent.Children)
            {
                var name = child?.DisplayName ?? "";
                if (string.Equals(name, label, StringComparison.OrdinalIgnoreCase))
                {
                    return child;
                }
            }

            return null;
        }

        private void OnHierarchyGroupsChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
            {
                foreach (var item in e.OldItems.OfType<QuickColourHierarchyGroup>())
                {
                    item.PropertyChanged -= OnHierarchyGroupPropertyChanged;
                    foreach (var t in item.Types)
                    {
                        t.PropertyChanged -= OnHierarchyTypePropertyChanged;
                    }
                }
            }

            if (e.NewItems != null)
            {
                foreach (var item in e.NewItems.OfType<QuickColourHierarchyGroup>())
                {
                    item.PropertyChanged += OnHierarchyGroupPropertyChanged;
                    foreach (var t in item.Types)
                    {
                        t.PropertyChanged += OnHierarchyTypePropertyChanged;
                    }
                }
            }
        }

        private void OnHierarchyGroupPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(QuickColourHierarchyGroup.BaseHex)
                || e.PropertyName == nameof(QuickColourHierarchyGroup.BaseColor)
                || e.PropertyName == nameof(QuickColourHierarchyGroup.UseCustomBaseColor)
                || e.PropertyName == nameof(QuickColourHierarchyGroup.HueGroupName)
                || e.PropertyName == nameof(QuickColourHierarchyGroup.Enabled))
            {
                ApplyHierarchyColours();
            }
        }

        private void OnHierarchyTypePropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(QuickColourHierarchyTypeRow.Enabled))
            {
                ApplyHierarchyColours();
            }
        }

        private void EnsureDefaultHueGroups()
        {
            if (HueGroups.Count > 0)
            {
                return;
            }

            HueGroups.Add(new QuickColourHueGroup { Name = "Architecture", HueHex = "#00A000" });
            HueGroups.Add(new QuickColourHueGroup { Name = "Mechanical", HueHex = "#2563EB" });
            HueGroups.Add(new QuickColourHueGroup { Name = "Electrical", HueHex = "#7C3AED" });
            HueGroups.Add(new QuickColourHueGroup { Name = "Fire", HueHex = "#DC2626" });
            HueGroups.Add(new QuickColourHueGroup { Name = "Other", HueHex = "#6B7280" });
        }

        private void RefreshHueGroupOptions()
        {
            HueGroupOptions.Clear();
            foreach (var g in HueGroups.Where(x => x != null && x.Enabled).Select(x => x.Name).Where(n => !string.IsNullOrWhiteSpace(n)).Distinct())
            {
                HueGroupOptions.Add(g);
            }

            if (HueGroupOptions.Count == 0)
            {
                HueGroupOptions.Add("Architecture");
            }
        }

        private string GetDefaultHueGroupName()
        {
            if (HierarchyUseHueGroups)
            {
                if (HueGroupOptions.Any(x => string.Equals(x, "Other", StringComparison.OrdinalIgnoreCase)))
                {
                    return HueGroupOptions.First(x => string.Equals(x, "Other", StringComparison.OrdinalIgnoreCase));
                }
            }

            return HueGroupOptions.FirstOrDefault() ?? "Architecture";
        }

        private void EnsureCategoryHueGroupDefaults()
        {
            if (HierarchyGroups == null || HierarchyGroups.Count == 0)
            {
                return;
            }

            var fallback = GetDefaultHueGroupName();
            foreach (var group in HierarchyGroups)
            {
                if (group == null)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(group.HueGroupName)
                    || !HueGroupOptions.Any(o => string.Equals(o, group.HueGroupName, StringComparison.OrdinalIgnoreCase)))
                {
                    group.HueGroupName = fallback;
                }
            }
        }

        private void OnHueGroupsChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
            {
                foreach (var item in e.OldItems.OfType<QuickColourHueGroup>())
                {
                    item.PropertyChanged -= OnHueGroupPropertyChanged;
                }
            }

            if (e.NewItems != null)
            {
                foreach (var item in e.NewItems.OfType<QuickColourHueGroup>())
                {
                    item.PropertyChanged += OnHueGroupPropertyChanged;
                }
            }

            RefreshHueGroupOptions();
            EnsureCategoryHueGroupDefaults();
            ApplyHierarchyColours();
        }

        private void OnHueGroupPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(QuickColourHueGroup.Enabled)
                || e.PropertyName == nameof(QuickColourHueGroup.Name)
                || e.PropertyName == nameof(QuickColourHueGroup.HueHex))
            {
                RefreshHueGroupOptions();
                EnsureCategoryHueGroupDefaults();
                ApplyHierarchyColours();
            }
        }

        private void InitDisciplineMap()
        {
            DisciplineMapPath = MicroEngStorageSettings.GetDataFilePath("QuickColourCategoryDisciplineMap.json");
            TryMigrateLegacyDisciplineMap(DisciplineMapPath);

            DisciplineMapFileIO.EnsureDefaultExists(DisciplineMapPath, DefaultCategoryDisciplineMapJson);
            ReloadDisciplineMap();
        }

        private static void TryMigrateLegacyDisciplineMap(string newPath)
        {
            try
            {
                if (File.Exists(newPath))
                {
                    return;
                }

                var legacyPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "MicroEng",
                    "Navisworks",
                    "QuickColour",
                    "CategoryDisciplineMap.json");
                if (!File.Exists(legacyPath))
                {
                    return;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(newPath) ?? MicroEngStorageSettings.DataStorageDirectory);
                File.Copy(legacyPath, newPath, overwrite: false);
            }
            catch
            {
                // non-fatal migration
            }
        }

        private void ReloadDisciplineMap()
        {
            _disciplineMap = DisciplineMapFileIO.Load(DisciplineMapPath);
            StatusText = $"Loaded category map: {_disciplineMap?.Rules?.Count ?? 0} rules (fallback={_disciplineMap?.FallbackGroup}).";
        }

        private void ApplyPreviewFilter()
        {
            AutoAssignPreviewRows.Clear();

            foreach (var row in _lastPreviewRows)
            {
                if (PreviewShowChangesOnly && !row.WillChange)
                {
                    continue;
                }

                AutoAssignPreviewRows.Add(row);
            }
        }

        private void ShowSnackbar(string title, string message, WpfUiControls.ControlAppearance appearance, WpfUiControls.SymbolRegular icon)
        {
            MicroEngSnackbar.Show(SnackbarPresenter, title, message, appearance, icon);
        }

        private void FlashSuccess(System.Windows.Controls.Button button)
        {
            if (button == null)
            {
                return;
            }

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

        private void Log(string message)
        {
            MicroEngActions.Log(message);
        }

        private static bool IsQuickColourTraceEnabled()
        {
            return string.Equals(
                Environment.GetEnvironmentVariable("MICROENG_QUICKCOLOUR_TRACE"),
                "1",
                StringComparison.OrdinalIgnoreCase);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string name = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
            {
                return false;
            }

            field = value;
            OnPropertyChanged(name);
            return true;
        }
    }
}

