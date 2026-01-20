
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MicroEng.Navisworks;
using MicroEng.Navisworks.SmartSets;
using NavisApp = Autodesk.Navisworks.Api.Application;

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

        private DisciplineMapFile _disciplineMap;
        private List<HueGroupAutoAssignPreviewRow> _lastPreviewRows = new List<HueGroupAutoAssignPreviewRow>();

        public ObservableCollection<ScraperProfileOption> ScraperProfiles { get; } =
            new ObservableCollection<ScraperProfileOption>();

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

        public ObservableCollection<QuickColourScope> ScopeOptions { get; } =
            new ObservableCollection<QuickColourScope>();

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

        private string _hierarchyL1Category = "";
        private string _hierarchyL1Property = "";
        private string _hierarchyL1Label = "Pick property...";

        public string HierarchyL1Label
        {
            get => _hierarchyL1Label;
            set => SetField(ref _hierarchyL1Label, value ?? "");
        }

        private string _hierarchyL2Category = "";
        private string _hierarchyL2Property = "";
        private string _hierarchyL2Label = "Pick property...";

        public string HierarchyL2Label
        {
            get => _hierarchyL2Label;
            set => SetField(ref _hierarchyL2Label, value ?? "");
        }

        private QuickColourPaletteStyle _hierarchyPaletteStyle = QuickColourPaletteStyle.Deep;
        public QuickColourPaletteStyle HierarchyPaletteStyle
        {
            get => _hierarchyPaletteStyle;
            set
            {
                if (SetField(ref _hierarchyPaletteStyle, value))
                {
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
                    if (HierarchySingleHueMode)
                    {
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
            set => SetField(ref _selectedScope, value);
        }

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

            foreach (QuickColourPaletteStyle style in Enum.GetValues(typeof(QuickColourPaletteStyle)))
            {
                PaletteStyleOptions.Add(style);
            }

            foreach (QuickColourTypeSortMode mode in Enum.GetValues(typeof(QuickColourTypeSortMode)))
            {
                TypeSortOptions.Add(mode);
            }

            foreach (QuickColourScope scope in Enum.GetValues(typeof(QuickColourScope)))
            {
                ScopeOptions.Add(scope);
            }

            EnsureDefaultHueGroups();
            RefreshHueGroupOptions();
            RefreshProfiles();
            InitDisciplineMap();

            DataScraperCache.SessionAdded += OnSessionAdded;
        }

        private void OnSessionAdded(ScrapeSession session)
        {
            RefreshProfiles();
        }

        private void RefreshProfiles()
        {
            ScraperProfiles.Clear();

            var sessions = DataScraperCache.AllSessions
                .OrderByDescending(s => s.Timestamp)
                .ToList();

            foreach (var session in sessions)
            {
                ScraperProfiles.Add(new ScraperProfileOption(session));
            }

            SelectedScraperProfile = ScraperProfiles.FirstOrDefault();

            if (SelectedScraperProfile == null && DataScraperCache.LastSession != null)
            {
                SelectedScraperProfile = new ScraperProfileOption(DataScraperCache.LastSession);
                ScraperProfiles.Add(SelectedScraperProfile);
            }
        }

        private void UpdateCurrentSessionForProfile()
        {
            EnsureDefaultHierarchyProperties();
            StatusText = SelectedScraperProfile?.Label ?? "Ready";
        }

        private ScrapeSession GetCurrentSession()
        {
            return SelectedScraperProfile?.Session ?? DataScraperCache.LastSession;
        }

        private void PickHierarchyL1_Click(object sender, RoutedEventArgs e)
        {
            TryPickProperty(true);
        }

        private void PickHierarchyL2_Click(object sender, RoutedEventArgs e)
        {
            TryPickProperty(false);
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
                    _hierarchyL1Category = window.Selected.Category ?? "";
                    _hierarchyL1Property = window.Selected.Name ?? "";
                    HierarchyL1Label = FormatPropertyLabel(_hierarchyL1Category, _hierarchyL1Property);
                }
                else
                {
                    _hierarchyL2Category = window.Selected.Category ?? "";
                    _hierarchyL2Property = window.Selected.Name ?? "";
                    HierarchyL2Label = FormatPropertyLabel(_hierarchyL2Category, _hierarchyL2Property);
                }
            }
        }

        private void OpenDataScraper_Click(object sender, RoutedEventArgs e)
        {
            MicroEngActions.TryShowDataScraper(null, out _);
        }

        private void RefreshProfiles_Click(object sender, RoutedEventArgs e)
        {
            RefreshProfiles();
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

                var l1 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var l2 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                foreach (var entry in session.RawEntries ?? Enumerable.Empty<RawEntry>())
                {
                    if (entry == null)
                    {
                        continue;
                    }

                    var item = entry.ItemPath ?? "";
                    if (string.IsNullOrWhiteSpace(item))
                    {
                        continue;
                    }

                    if (IsMatch(entry, _hierarchyL1Category, _hierarchyL1Property))
                    {
                        l1[item] = entry.Value ?? "";
                    }

                    if (IsMatch(entry, _hierarchyL2Category, _hierarchyL2Property))
                    {
                        l2[item] = entry.Value ?? "";
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

            var palette = QuickColourPalette.GeneratePalette(groups.Count, HierarchyPaletteStyle);
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

            ApplyHierarchyColours();
        }

        private void RegenerateTypeShades_Click(object sender, RoutedEventArgs e)
        {
            ApplyHierarchyColours();
        }

        private void ApplyTemporary_Click(object sender, RoutedEventArgs e)
        {
            ApplyHierarchy(permanent: false);
        }

        private void ApplyPermanent_Click(object sender, RoutedEventArgs e)
        {
            ApplyHierarchy(permanent: true);
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

        private void ApplyHierarchy(bool permanent)
        {
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

            if (string.IsNullOrWhiteSpace(_hierarchyL1Category)
                || string.IsNullOrWhiteSpace(_hierarchyL1Property)
                || string.IsNullOrWhiteSpace(_hierarchyL2Category)
                || string.IsNullOrWhiteSpace(_hierarchyL2Property))
            {
                StatusText = "Pick Level 1 and Level 2 properties first.";
                return;
            }

            if (SelectedScope == QuickColourScope.CurrentSelection)
            {
                var sel = doc.CurrentSelection?.SelectedItems;
                if (sel == null || sel.Count == 0)
                {
                    StatusText = "Scope is Current Selection, but no items are selected.";
                    return;
                }
            }

            ApplyHierarchyColours();

            var count = _service.ApplyByHierarchy(
                doc,
                _hierarchyL1Category,
                _hierarchyL1Property,
                _hierarchyL2Category,
                _hierarchyL2Property,
                HierarchyGroups,
                SelectedScope,
                permanent,
                CreateSearchSets,
                CreateSnapshots,
                OutputFolderPath,
                OutputProfileName,
                CreateFoldersByHueGroup && HierarchyUseHueGroups,
                Log);

            StatusText = $"Applied {count} item(s).";
        }
        private void ApplyHierarchyColours()
        {
            if (HierarchyGroups == null || HierarchyGroups.Count == 0)
            {
                return;
            }

            if (HierarchyUseHueGroups)
            {
                ApplyHueGroupsColours();
                return;
            }

            if (HierarchySingleHueMode)
            {
                ApplySingleHueColours();
                return;
            }

            ApplyBasePaletteColours();
        }

        private void ApplyBasePaletteColours()
        {
            var style = HierarchyPaletteStyle;
            var typeUsage = Math.Max(0.15, GetTypeSpread01());

            foreach (var group in HierarchyGroups.Where(g => g != null && g.Enabled))
            {
                var types = GetSortedTypes(group);
                if (types.Count == 0)
                {
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
            var style = HierarchyPaletteStyle;
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

        private void ApplyHueGroupsColours()
        {
            var style = HierarchyPaletteStyle;
            var typeSpread01 = Math.Max(0.15, GetTypeSpread01());
            var contrast01 = Clamp01(HierarchyCategoryContrastPct / 100.0);

            var (minL, maxL) = QuickColourPalette.ComputeCategoryLightnessRange(style, contrast01);
            var sat = QuickColourPalette.GetRecommendedSaturation(style);

            var enabledGroups = HueGroups
                .Where(h => h != null && h.Enabled && !string.IsNullOrWhiteSpace(h.Name))
                .ToDictionary(h => h.Name, h => h.HueColor, StringComparer.OrdinalIgnoreCase);

            if (enabledGroups.Count == 0)
            {
                return;
            }

            var categories = HierarchyGroups.Where(c => c != null && c.Enabled).ToList();

            foreach (var kvp in enabledGroups)
            {
                var groupName = kvp.Key;
                var hueColor = kvp.Value;
                var hue01 = QuickColourPalette.GetHue01(hueColor);

                var catsInGroup = categories
                    .Where(c => string.Equals(c.HueGroupName, groupName, StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(c => c.Count)
                    .ThenBy(c => c.Value, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (catsInGroup.Count == 0)
                {
                    continue;
                }

                int n = catsInGroup.Count;
                double bandWidth = (maxL - minL) / n;
                double pad = bandWidth * 0.08;

                for (int i = 0; i < n; i++)
                {
                    var cat = catsInGroup[i];

                    double bandMin = minL + bandWidth * i + pad;
                    double bandMax = minL + bandWidth * (i + 1) - pad;
                    if (bandMax <= bandMin)
                    {
                        bandMin = minL + bandWidth * i;
                        bandMax = minL + bandWidth * (i + 1);
                    }

                    double baseL = (bandMin + bandMax) / 2.0;

                    if (!cat.UseCustomBaseColor)
                    {
                        cat.SetComputedBaseColor(QuickColourPalette.FromHsl01(hue01, sat, baseL));
                    }

                    var types = GetSortedTypes(cat);
                    if (types.Count == 0)
                    {
                        continue;
                    }

                    var shades = QuickColourPalette.GenerateShades(cat.BaseColor, types.Count, style, typeSpread01);
                    for (int t = 0; t < types.Count; t++)
                    {
                        types[t].Color = shades[t];
                    }
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

        private Color ParseHueColor()
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

            if (string.IsNullOrWhiteSpace(_hierarchyL1Category) || string.IsNullOrWhiteSpace(_hierarchyL1Property))
            {
                if (TryPickProperty(session, "Element", "Category", out var cat, out var prop)
                    || TryPickPropertyByName(session, "Category", out cat, out prop))
                {
                    _hierarchyL1Category = cat;
                    _hierarchyL1Property = prop;
                    HierarchyL1Label = FormatPropertyLabel(cat, prop);
                }
            }

            if (string.IsNullOrWhiteSpace(_hierarchyL2Category) || string.IsNullOrWhiteSpace(_hierarchyL2Property))
            {
                if (TryPickProperty(session, "Element", "Type", out var cat, out var prop)
                    || TryPickProperty(session, "Element", "Type Name", out cat, out prop)
                    || TryPickPropertyByName(session, "Type", out cat, out prop))
                {
                    _hierarchyL2Category = cat;
                    _hierarchyL2Property = prop;
                    HierarchyL2Label = FormatPropertyLabel(cat, prop);
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

        private void AddHueGroup_Click(object sender, RoutedEventArgs e)
        {
            HueGroups.Add(new QuickColourHueGroup { Name = "New Group", HueHex = "#00A000" });
            RefreshHueGroupOptions();
        }

        private void RemoveHueGroup_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedHueGroup == null)
            {
                return;
            }

            var removedName = SelectedHueGroup.Name;
            HueGroups.Remove(SelectedHueGroup);
            RefreshHueGroupOptions();

            var fallback = GetDefaultHueGroupName();
            foreach (var cat in HierarchyGroups)
            {
                if (cat != null && string.Equals(cat.HueGroupName, removedName, StringComparison.OrdinalIgnoreCase))
                {
                    cat.HueGroupName = fallback;
                }
            }

            ApplyHierarchyColours();
        }

        private void ResetHueGroups_Click(object sender, RoutedEventArgs e)
        {
            HueGroups.Clear();
            EnsureDefaultHueGroups();
            RefreshHueGroupOptions();

            foreach (var cat in HierarchyGroups)
            {
                if (cat != null)
                {
                    cat.HueGroupName = "Architecture";
                }
            }

            ApplyHierarchyColours();
        }
        private void InitDisciplineMap()
        {
            var baseDir = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "MicroEng", "Navisworks", "QuickColour");

            DisciplineMapPath = System.IO.Path.Combine(baseDir, "CategoryDisciplineMap.json");

            DisciplineMapFileIO.EnsureDefaultExists(DisciplineMapPath, DefaultCategoryDisciplineMapJson);
            ReloadDisciplineMap();
        }

        private void ReloadDisciplineMap()
        {
            _disciplineMap = DisciplineMapFileIO.Load(DisciplineMapPath);
            StatusText = $"Loaded category map: {_disciplineMap?.Rules?.Count ?? 0} rules (fallback={_disciplineMap?.FallbackGroup}).";
        }

        private void OpenDisciplineMap_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = DisciplineMapPath,
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(psi);
            }
            catch (Exception ex)
            {
                StatusText = "Open map failed: " + ex.Message;
            }
        }

        private void ReloadDisciplineMap_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ReloadDisciplineMap();
            }
            catch (Exception ex)
            {
                StatusText = "Reload map failed: " + ex.Message;
            }
        }

        private void AutoAssignHueGroups_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_disciplineMap == null)
                {
                    ReloadDisciplineMap();
                }

                if (HierarchyGroups == null || HierarchyGroups.Count == 0)
                {
                    StatusText = "No categories loaded. Run Load Hierarchy first.";
                    return;
                }

                var options = new HueGroupAutoAssignOptions
                {
                    OnlyAssignWhenCurrentlyFallback = AutoAssignOnlyFallback,
                    SkipLockedCategories = AutoAssignSkipLocked,
                    AutoCreateMissingHueGroups = AutoAssignCreateMissingGroups
                };

                var res = _autoAssign.Apply(_disciplineMap, HierarchyGroups.ToList(), HueGroups, options, Log);

                if (res.MissingHueGroups.Count > 0)
                {
                    StatusText = $"Auto assigned. Missing groups in UI: {string.Join(", ", res.MissingHueGroups)} (mapped to fallback).";
                }
                else
                {
                    StatusText = $"Auto assigned {res.Assigned} categories.";
                }

                RefreshHueGroupOptions();
                ApplyHierarchyColours();
            }
            catch (Exception ex)
            {
                StatusText = "Auto assign failed: " + ex.Message;
            }
        }

        private void PreviewHueMap_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_disciplineMap == null)
                {
                    ReloadDisciplineMap();
                }

                if (HierarchyGroups == null || HierarchyGroups.Count == 0)
                {
                    AutoAssignPreviewSummary = "No hierarchy loaded. Run Load Hierarchy first.";
                    AutoAssignPreviewRows.Clear();
                    return;
                }

                var options = new HueGroupAutoAssignOptions
                {
                    OnlyAssignWhenCurrentlyFallback = AutoAssignOnlyFallback,
                    SkipLockedCategories = AutoAssignSkipLocked,
                    AutoCreateMissingHueGroups = AutoAssignCreateMissingGroups
                };

                var rows = new List<HueGroupAutoAssignPreviewRow>();
                var summary = _autoAssign.BuildPreview(
                    _disciplineMap,
                    HierarchyGroups.ToList(),
                    HueGroups.ToList(),
                    options,
                    rows,
                    Log);

                _lastPreviewRows = rows;
                ApplyPreviewFilter();

                AutoAssignPreviewSummary =
                    $"Preview: Total={summary.Total}, Will change={summary.WillChange}, " +
                    $"No change={summary.NoChange}, Skipped locked={summary.SkippedLocked}, " +
                    $"Skipped assigned={summary.SkippedAlreadyAssigned}, Unmatched+fallback={summary.UnmatchedFallback}, " +
                    $"Missing group+fallback={summary.MissingGroupFallback}, Will create group={summary.WillCreateMissingGroup}.";
            }
            catch (Exception ex)
            {
                AutoAssignPreviewSummary = "Preview failed: " + ex.Message;
            }
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

        private void Log(string message)
        {
            MicroEngActions.Log(message);
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
