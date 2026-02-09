using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using Autodesk.Navisworks.Api;
using NavisApp = Autodesk.Navisworks.Api.Application;
using MicroEng.Navisworks;
using WpfUiControls = Wpf.Ui.Controls;

namespace MicroEng.Navisworks.SmartSets
{
    public partial class SmartSetGeneratorControl : UserControl, INotifyPropertyChanged
    {
        private readonly SmartSetGeneratorNavisworksService _navisworksService = new SmartSetGeneratorNavisworksService();
        private readonly SmartSetFastPreviewService _fastPreviewService = new SmartSetFastPreviewService();
        private SmartSetRecipeStore _recipeStore;
        private CancellationTokenSource _previewCts;
        private ModelItemCollection _lastPreviewResults;

        private SmartSetRecipe _currentRecipe;
        private string _selectedScraperProfile;
        private SmartSetRule _selectedRule;
        private bool _useFastPreview = true;
        private string _previewStatusText = "Ready.";
        private string _previewDetailsText = "";
        private RecipeFileItem _selectedRecipeFile;
        private SmartSetSuggestion _selectedSuggestion;
        private SmartSetPackDefinition _selectedPack;
        private readonly Dictionary<string, List<string>> _propertyOptionsByCategory = new(StringComparer.OrdinalIgnoreCase);
        private readonly SmartSetGeneratorQuickBuilderPage _quickBuilderPage;
        private readonly SmartSetGeneratorSmartGroupingPage _smartGroupingPage;
        private readonly SmartSetGeneratorFromSelectionPage _fromSelectionPage;
        private readonly SmartSetGeneratorPacksPage _packsPage;
        private readonly Dictionary<Type, UserControl> _pageMap = new();
        private bool _initialized;
        private bool _dataScraperEventsHooked;

        public SmartSetGeneratorControl()
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

            ConditionOptions = new ObservableCollection<ConditionOption>(BuildConditionOptions());
            SearchSetModeOptions = new ObservableCollection<SearchSetModeOption>(BuildSearchSetModeOptions());
            SearchInModeOptions = new ObservableCollection<SearchInModeOption>(BuildSearchInModeOptions());
            ScopeModeOptions = new ObservableCollection<ScopeModeOption>(BuildScopeModeOptions());
            OutputTypes = new ObservableCollection<SmartSetOutputType>(Enum.GetValues(typeof(SmartSetOutputType)).Cast<SmartSetOutputType>());

            CurrentRecipe = new SmartSetRecipe();
            if (CurrentRecipe.Rules.Count == 0)
            {
                CurrentRecipe.Rules.Add(new SmartSetRule());
            }

            InitRecipeStore();
            RefreshRecipeFiles();
            RefreshScraperProfiles();
            BuildPackList();

            _quickBuilderPage = new SmartSetGeneratorQuickBuilderPage(this);
            _smartGroupingPage = new SmartSetGeneratorSmartGroupingPage(this);
            _fromSelectionPage = new SmartSetGeneratorFromSelectionPage(this);
            _packsPage = new SmartSetGeneratorPacksPage(this);

            _pageMap[typeof(SmartSetGeneratorQuickBuilderPage)] = _quickBuilderPage;
            _pageMap[typeof(SmartSetGeneratorSmartGroupingPage)] = _smartGroupingPage;
            _pageMap[typeof(SmartSetGeneratorFromSelectionPage)] = _fromSelectionPage;
            _pageMap[typeof(SmartSetGeneratorPacksPage)] = _packsPage;

            Loaded += OnLoaded;
            Unloaded += OnUnloaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            EnsureDataScraperEventHandlers();

            if (_initialized)
            {
                return;
            }

            _initialized = true;
            RefreshSavedSelectionSets();
            NavigateToPage(typeof(SmartSetGeneratorQuickBuilderPage));
        }

        private void OnUnloaded(object sender, RoutedEventArgs e)
        {
            if (!_dataScraperEventsHooked)
            {
                return;
            }

            DataScraperCache.SessionAdded -= OnSessionAdded;
            DataScraperCache.CacheChanged -= OnCacheChanged;
            _dataScraperEventsHooked = false;
        }

        private void EnsureDataScraperEventHandlers()
        {
            if (_dataScraperEventsHooked)
            {
                return;
            }

            DataScraperCache.SessionAdded += OnSessionAdded;
            DataScraperCache.CacheChanged += OnCacheChanged;
            _dataScraperEventsHooked = true;
        }

        private void NavigateToPage(Type pageType)
        {
            if (pageType == null)
            {
                return;
            }

            if (_pageMap.TryGetValue(pageType, out var page))
            {
                if (!ReferenceEquals(SmartSetHost.Content, page))
                {
                    SmartSetHost.Content = page;
                }

                UpdateNavButtonStates(pageType);
            }
        }

        private void UpdateNavButtonStates(Type pageType)
        {
            SetNavButtonState(QuickBuilderNavButton, pageType == typeof(SmartSetGeneratorQuickBuilderPage));
            SetNavButtonState(SmartGroupingNavButton, pageType == typeof(SmartSetGeneratorSmartGroupingPage));
            SetNavButtonState(FromSelectionNavButton, pageType == typeof(SmartSetGeneratorFromSelectionPage));
            SetNavButtonState(PacksNavButton, pageType == typeof(SmartSetGeneratorPacksPage));
        }

        private static void SetNavButtonState(Wpf.Ui.Controls.Button button, bool isActive)
        {
            if (button == null)
            {
                return;
            }

            button.Appearance = isActive
                ? Wpf.Ui.Controls.ControlAppearance.Primary
                : Wpf.Ui.Controls.ControlAppearance.Secondary;
            button.FontWeight = isActive ? FontWeights.SemiBold : FontWeights.Normal;
        }

        private void QuickBuilderNav_Click(object sender, RoutedEventArgs e)
        {
            NavigateToPage(typeof(SmartSetGeneratorQuickBuilderPage));
        }

        private void SmartGroupingNav_Click(object sender, RoutedEventArgs e)
        {
            NavigateToPage(typeof(SmartSetGeneratorSmartGroupingPage));
        }

        private void FromSelectionNav_Click(object sender, RoutedEventArgs e)
        {
            NavigateToPage(typeof(SmartSetGeneratorFromSelectionPage));
        }

        private void PacksNav_Click(object sender, RoutedEventArgs e)
        {
            NavigateToPage(typeof(SmartSetGeneratorPacksPage));
        }

        public ObservableCollection<string> ScraperProfiles { get; } = new ObservableCollection<string>();
        public ObservableCollection<string> CategoryOptions { get; } = new ObservableCollection<string>();
        public ObservableCollection<string> GroupByPropertyOptions { get; } = new ObservableCollection<string>();
        public ObservableCollection<string> ThenByPropertyOptions { get; } = new ObservableCollection<string>();
        public ObservableCollection<ConditionOption> ConditionOptions { get; }
        public ObservableCollection<SearchSetModeOption> SearchSetModeOptions { get; }
        public ObservableCollection<SearchInModeOption> SearchInModeOptions { get; }
        public ObservableCollection<ScopeModeOption> ScopeModeOptions { get; }
        public ObservableCollection<SmartSetOutputType> OutputTypes { get; }
        public ObservableCollection<SmartGroupRow> GroupRows { get; } = new ObservableCollection<SmartGroupRow>();
        public ObservableCollection<SmartSetSuggestion> SelectionSuggestions { get; } = new ObservableCollection<SmartSetSuggestion>();
        public ObservableCollection<SmartSetPackDefinition> Packs { get; } = new ObservableCollection<SmartSetPackDefinition>();
        public ObservableCollection<RecipeFileItem> RecipeFiles { get; } = new ObservableCollection<RecipeFileItem>();
        public ObservableCollection<string> SavedSelectionSets { get; } = new ObservableCollection<string>();

        public SmartSetRecipe CurrentRecipe
        {
            get => _currentRecipe;
            set
            {
                if (ReferenceEquals(_currentRecipe, value))
                {
                    return;
                }

                if (_currentRecipe != null)
                {
                    UnhookRecipe(_currentRecipe);
                }

                _currentRecipe = value ?? new SmartSetRecipe();
                _currentRecipe.Rules ??= new ObservableCollection<SmartSetRule>();
                _currentRecipe.Grouping ??= new SmartSetGroupingSpec();
                EnsureScopeDefaults(_currentRecipe);
                EnsureGroupingDefaults(_currentRecipe.Grouping);
                HookRecipe(_currentRecipe);
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanUseFastPreview));
                EnforcePreviewCompatibility(showStatus: false);

                if (_currentRecipe.SearchSetMode == SmartSetSearchSetMode.Single && _currentRecipe.GenerateMultipleSearchSets)
                {
                    _currentRecipe.SearchSetMode = SmartSetSearchSetMode.SplitByValue;
                }

                RefreshCategoryAndPropertyOptions();
            }
        }

        public string SelectedScraperProfile
        {
            get => _selectedScraperProfile;
            set
            {
                if (string.Equals(_selectedScraperProfile, value, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                _selectedScraperProfile = value;
                if (CurrentRecipe != null)
                {
                    CurrentRecipe.DataScraperProfile = value ?? "";
                    if (string.IsNullOrWhiteSpace(CurrentRecipe.FolderPath))
                    {
                        CurrentRecipe.FolderPath = BuildDefaultFolderPath(value);
                    }

                    if (CurrentRecipe.Grouping != null && string.IsNullOrWhiteSpace(CurrentRecipe.Grouping.OutputFolderPath))
                    {
                        CurrentRecipe.Grouping.OutputFolderPath = BuildDefaultFolderPath(value);
                    }
                }
                OnPropertyChanged();
                RefreshCategoryAndPropertyOptions();
            }
        }

        public SmartSetRule SelectedRule
        {
            get => _selectedRule;
            set
            {
                _selectedRule = value;
                OnPropertyChanged();
            }
        }

        public bool UseFastPreview
        {
            get => _useFastPreview;
            set
            {
                if (value && !CanUseFastPreview)
                {
                    _useFastPreview = false;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(CanSelectPreviewResults));
                    return;
                }

                if (_useFastPreview == value)
                {
                    return;
                }

                _useFastPreview = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanSelectPreviewResults));
            }
        }

        public string PreviewStatusText
        {
            get => _previewStatusText;
            set
            {
                _previewStatusText = value ?? "";
                OnPropertyChanged();
            }
        }

        public string PreviewDetailsText
        {
            get => _previewDetailsText;
            set
            {
                _previewDetailsText = value ?? "";
                OnPropertyChanged();
            }
        }

        public RecipeFileItem SelectedRecipeFile
        {
            get => _selectedRecipeFile;
            set
            {
                _selectedRecipeFile = value;
                OnPropertyChanged();
            }
        }

        public SmartSetSuggestion SelectedSuggestion
        {
            get => _selectedSuggestion;
            set
            {
                _selectedSuggestion = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanApplySuggestion));
            }
        }

        public SmartSetPackDefinition SelectedPack
        {
            get => _selectedPack;
            set
            {
                _selectedPack = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanRunPack));
            }
        }

        public bool CanApplySuggestion => SelectedSuggestion != null;
        public bool CanRunPack => SelectedPack != null;
        public bool CanUseFastPreview => CurrentRecipe == null || !CurrentRecipe.IsScopeConstrained;
        public bool CanSelectPreviewResults => !UseFastPreview && _lastPreviewResults != null && _lastPreviewResults.Count > 0;

        private void OnSessionAdded(ScrapeSession session)
        {
            Dispatcher.BeginInvoke(new Action(RefreshScraperProfiles));
        }

        private void OnCacheChanged()
        {
            Dispatcher.BeginInvoke(new Action(RefreshScraperProfiles));
        }

        private void RefreshScraperProfiles()
        {
            var previousSelection = SelectedScraperProfile;
            ScraperProfiles.Clear();
            var profiles = DataScraperCache.AllSessions
                .Where(s => !string.IsNullOrWhiteSpace(s.ProfileName))
                .OrderByDescending(s => s.Timestamp)
                .Select(s => s.ProfileName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var p in profiles)
            {
                ScraperProfiles.Add(p);
            }

            if (profiles.Count > 0)
            {
                var match = profiles.FirstOrDefault(p => string.Equals(p, previousSelection, StringComparison.OrdinalIgnoreCase));
                var nextSelection = match ?? profiles[0];
                if (!string.Equals(SelectedScraperProfile, nextSelection, StringComparison.OrdinalIgnoreCase))
                {
                    SelectedScraperProfile = nextSelection;
                    return;
                }
            }
            else if (!string.IsNullOrWhiteSpace(SelectedScraperProfile))
            {
                SelectedScraperProfile = string.Empty;
                return;
            }

            RefreshCategoryAndPropertyOptions();
        }

        private void RefreshSavedSelectionSets()
        {
            SavedSelectionSets.Clear();

            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                return;
            }

            foreach (var name in _navisworksService.GetSavedSelectionSetNames(doc))
            {
                SavedSelectionSets.Add(name);
            }
        }

        private void HookRecipe(SmartSetRecipe recipe)
        {
            if (recipe == null)
            {
                return;
            }

            recipe.PropertyChanged += OnRecipePropertyChanged;
            if (recipe.Grouping != null)
            {
                recipe.Grouping.PropertyChanged += OnGroupingPropertyChanged;
            }
            if (recipe.Rules == null)
            {
                return;
            }

            recipe.Rules.CollectionChanged += OnRulesCollectionChanged;
            foreach (var rule in recipe.Rules)
            {
                HookRule(rule);
            }
        }

        private void UnhookRecipe(SmartSetRecipe recipe)
        {
            if (recipe == null)
            {
                return;
            }

            recipe.PropertyChanged -= OnRecipePropertyChanged;
            if (recipe.Grouping != null)
            {
                recipe.Grouping.PropertyChanged -= OnGroupingPropertyChanged;
            }
            if (recipe.Rules == null)
            {
                return;
            }

            recipe.Rules.CollectionChanged -= OnRulesCollectionChanged;
            foreach (var rule in recipe.Rules)
            {
                UnhookRule(rule);
            }
        }

        private void OnRecipePropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (sender is not SmartSetRecipe recipe)
            {
                return;
            }

            if (string.Equals(e.PropertyName, nameof(SmartSetRecipe.ScopeMode), StringComparison.Ordinal))
            {
                UpdateScopeSummary(recipe);

                if (recipe.ScopeMode == SmartSetScopeMode.SavedSelectionSet)
                {
                    RefreshSavedSelectionSets();
                }

                OnPropertyChanged(nameof(CanUseFastPreview));
                EnforcePreviewCompatibility(showStatus: true);
            }
            else if (string.Equals(e.PropertyName, nameof(SmartSetRecipe.ScopeSelectionSetName), StringComparison.Ordinal))
            {
                UpdateScopeSummary(recipe);
                OnPropertyChanged(nameof(CanUseFastPreview));
                EnforcePreviewCompatibility(showStatus: false);
            }
            else if (string.Equals(e.PropertyName, nameof(SmartSetRecipe.ScopeModelPaths), StringComparison.Ordinal)
                || string.Equals(e.PropertyName, nameof(SmartSetRecipe.ScopeFilterCategory), StringComparison.Ordinal)
                || string.Equals(e.PropertyName, nameof(SmartSetRecipe.ScopeFilterProperty), StringComparison.Ordinal)
                || string.Equals(e.PropertyName, nameof(SmartSetRecipe.ScopeFilterValue), StringComparison.Ordinal))
            {
                UpdateScopeSummary(recipe);
                OnPropertyChanged(nameof(CanUseFastPreview));
                EnforcePreviewCompatibility(showStatus: false);
            }
        }

        private static void EnsureScopeDefaults(SmartSetRecipe recipe)
        {
            if (recipe == null)
            {
                return;
            }

            recipe.ScopeSelectionSetName ??= "";
            recipe.ScopeModelPaths ??= new List<List<string>>();
            recipe.ScopeFilterCategory ??= "";
            recipe.ScopeFilterProperty ??= "";
            recipe.ScopeFilterValue ??= "";

            UpdateScopeSummary(recipe);
        }

        private static void EnsureGroupingDefaults(SmartSetGroupingSpec grouping)
        {
            if (grouping == null)
            {
                return;
            }

            grouping.OutputName ??= "New Grouping";
            grouping.OutputFolderPath ??= "MicroEng/Smart Sets";
            grouping.ScopeSelectionSetName ??= "";
            grouping.ScopeModelPaths ??= new List<List<string>>();
            grouping.ScopeFilterCategory ??= "";
            grouping.ScopeFilterProperty ??= "";
            grouping.ScopeFilterValue ??= "";

            UpdateGroupingScopeSummary(grouping);
        }

        private static void UpdateScopeSummary(SmartSetRecipe recipe)
        {
            if (recipe == null)
            {
                return;
            }

            switch (recipe.ScopeMode)
            {
                case SmartSetScopeMode.AllModel:
                    recipe.ScopeSummary = "Entire model";
                    break;
                case SmartSetScopeMode.CurrentSelection:
                    if (string.IsNullOrWhiteSpace(recipe.ScopeSummary)
                        || !recipe.ScopeSummary.StartsWith("Current selection", StringComparison.OrdinalIgnoreCase))
                    {
                        recipe.ScopeSummary = "Current selection";
                    }
                    break;
                case SmartSetScopeMode.SavedSelectionSet:
                    recipe.ScopeSummary = string.IsNullOrWhiteSpace(recipe.ScopeSelectionSetName)
                        ? "Saved selection set"
                        : $"Saved selection set: {recipe.ScopeSelectionSetName}";
                    break;
                case SmartSetScopeMode.ModelTree:
                    recipe.ScopeSummary = BuildTreeScopeSummary(recipe.ScopeModelPaths);
                    break;
                case SmartSetScopeMode.PropertyFilter:
                    recipe.ScopeSummary = BuildFilterScopeSummary(recipe);
                    break;
                default:
                    if (string.IsNullOrWhiteSpace(recipe.ScopeSummary))
                    {
                        recipe.ScopeSummary = "Scope active";
                    }
                    break;
            }
        }

        private static void UpdateGroupingScopeSummary(SmartSetGroupingSpec grouping)
        {
            if (grouping == null)
            {
                return;
            }

            switch (grouping.ScopeMode)
            {
                case SmartSetScopeMode.AllModel:
                    grouping.ScopeSummary = "Entire model";
                    break;
                case SmartSetScopeMode.CurrentSelection:
                    if (string.IsNullOrWhiteSpace(grouping.ScopeSummary)
                        || !grouping.ScopeSummary.StartsWith("Current selection", StringComparison.OrdinalIgnoreCase))
                    {
                        grouping.ScopeSummary = "Current selection";
                    }
                    break;
                case SmartSetScopeMode.SavedSelectionSet:
                    grouping.ScopeSummary = string.IsNullOrWhiteSpace(grouping.ScopeSelectionSetName)
                        ? "Saved selection set"
                        : $"Saved selection set: {grouping.ScopeSelectionSetName}";
                    break;
                case SmartSetScopeMode.ModelTree:
                    grouping.ScopeSummary = BuildTreeScopeSummary(grouping.ScopeModelPaths);
                    break;
                case SmartSetScopeMode.PropertyFilter:
                    grouping.ScopeSummary = BuildFilterScopeSummary(grouping);
                    break;
                default:
                    if (string.IsNullOrWhiteSpace(grouping.ScopeSummary))
                    {
                        grouping.ScopeSummary = "Scope active";
                    }
                    break;
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

        private static string BuildFilterScopeSummary(SmartSetRecipe recipe)
        {
            if (recipe == null || !recipe.HasScopePropertyFilter)
            {
                return "Property filter";
            }

            return $"Property filter: {recipe.ScopeFilterCategory} / {recipe.ScopeFilterProperty} = {recipe.ScopeFilterValue}";
        }

        private static string BuildFilterScopeSummary(SmartSetGroupingSpec grouping)
        {
            if (grouping == null || !grouping.HasScopePropertyFilter)
            {
                return "Property filter";
            }

            return $"Property filter: {grouping.ScopeFilterCategory} / {grouping.ScopeFilterProperty} = {grouping.ScopeFilterValue}";
        }

        private void EnforcePreviewCompatibility(bool showStatus)
        {
            if (CurrentRecipe?.IsScopeConstrained == true && UseFastPreview)
            {
                UseFastPreview = false;
                if (showStatus)
                {
                    PreviewStatusText = "Scope active: using live preview.";
                }
            }
        }

        private void OnRulesCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
            {
                foreach (var item in e.OldItems)
                {
                    if (item is SmartSetRule rule)
                    {
                        UnhookRule(rule);
                    }
                }
            }

            if (e.NewItems != null)
            {
                foreach (var item in e.NewItems)
                {
                    if (item is SmartSetRule rule)
                    {
                        HookRule(rule);
                    }
                }
            }
        }

        private void HookRule(SmartSetRule rule)
        {
            if (rule == null)
            {
                return;
            }

            rule.PropertyChanged += OnRulePropertyChanged;
            EnsureCategoryOption(rule.Category);
            UpdateRulePropertyOptions(rule);
        }

        private void UnhookRule(SmartSetRule rule)
        {
            if (rule == null)
            {
                return;
            }

            rule.PropertyChanged -= OnRulePropertyChanged;
        }

        private void OnRulePropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (sender is not SmartSetRule rule)
            {
                return;
            }

            if (string.Equals(e.PropertyName, nameof(SmartSetRule.Category), StringComparison.Ordinal))
            {
                EnsureCategoryOption(rule.Category);
                UpdateRulePropertyOptions(rule);
            }
            else if (string.Equals(e.PropertyName, nameof(SmartSetRule.Property), StringComparison.Ordinal))
            {
                EnsurePropertyOption(rule);
            }
        }

        private void OnGroupingPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (sender is not SmartSetGroupingSpec grouping)
            {
                return;
            }

            if (string.Equals(e.PropertyName, nameof(SmartSetGroupingSpec.GroupByCategory), StringComparison.Ordinal))
            {
                EnsureCategoryOption(grouping.GroupByCategory);
                UpdateGroupingPropertyOptionsFor(grouping.GroupByCategory, grouping.GroupByProperty, GroupByPropertyOptions);
            }
            else if (string.Equals(e.PropertyName, nameof(SmartSetGroupingSpec.ThenByCategory), StringComparison.Ordinal))
            {
                EnsureCategoryOption(grouping.ThenByCategory);
                UpdateGroupingPropertyOptionsFor(grouping.ThenByCategory, grouping.ThenByProperty, ThenByPropertyOptions);
            }
            else if (string.Equals(e.PropertyName, nameof(SmartSetGroupingSpec.GroupByProperty), StringComparison.Ordinal))
            {
                EnsurePropertyOption(GroupByPropertyOptions, grouping.GroupByProperty);
            }
            else if (string.Equals(e.PropertyName, nameof(SmartSetGroupingSpec.ThenByProperty), StringComparison.Ordinal))
            {
                EnsurePropertyOption(ThenByPropertyOptions, grouping.ThenByProperty);
            }
            else if (string.Equals(e.PropertyName, nameof(SmartSetGroupingSpec.ScopeMode), StringComparison.Ordinal)
                || string.Equals(e.PropertyName, nameof(SmartSetGroupingSpec.ScopeSelectionSetName), StringComparison.Ordinal)
                || string.Equals(e.PropertyName, nameof(SmartSetGroupingSpec.ScopeModelPaths), StringComparison.Ordinal)
                || string.Equals(e.PropertyName, nameof(SmartSetGroupingSpec.ScopeFilterCategory), StringComparison.Ordinal)
                || string.Equals(e.PropertyName, nameof(SmartSetGroupingSpec.ScopeFilterProperty), StringComparison.Ordinal)
                || string.Equals(e.PropertyName, nameof(SmartSetGroupingSpec.ScopeFilterValue), StringComparison.Ordinal))
            {
                if (grouping.ScopeMode == SmartSetScopeMode.SavedSelectionSet)
                {
                    RefreshSavedSelectionSets();
                }

                UpdateGroupingScopeSummary(grouping);
            }
        }

        private void RefreshCategoryAndPropertyOptions()
        {
            _propertyOptionsByCategory.Clear();
            var propertyNamesByCategory = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            var session = GetSelectedSession();
            var properties = session?.Properties ?? Enumerable.Empty<ScrapedProperty>();

            foreach (var prop in properties)
            {
                var category = prop?.Category ?? "";
                var name = prop?.Name ?? "";
                if (string.IsNullOrWhiteSpace(category) || string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (!_propertyOptionsByCategory.TryGetValue(category, out var list))
                {
                    list = new List<string>();
                    _propertyOptionsByCategory[category] = list;
                    propertyNamesByCategory[category] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }

                if (propertyNamesByCategory.TryGetValue(category, out var names) && names.Add(name))
                {
                    list.Add(name);
                }
            }

            foreach (var list in _propertyOptionsByCategory.Values)
            {
                list.Sort(StringComparer.OrdinalIgnoreCase);
            }

            var categories = new HashSet<string>(_propertyOptionsByCategory.Keys, StringComparer.OrdinalIgnoreCase);
            if (CurrentRecipe?.Rules != null)
            {
                foreach (var rule in CurrentRecipe.Rules)
                {
                    var cat = rule?.Category;
                    if (!string.IsNullOrWhiteSpace(cat))
                    {
                        categories.Add(cat);
                    }
                }
            }
            if (CurrentRecipe?.Grouping != null)
            {
                if (!string.IsNullOrWhiteSpace(CurrentRecipe.Grouping.GroupByCategory))
                {
                    categories.Add(CurrentRecipe.Grouping.GroupByCategory);
                }

                if (!string.IsNullOrWhiteSpace(CurrentRecipe.Grouping.ThenByCategory))
                {
                    categories.Add(CurrentRecipe.Grouping.ThenByCategory);
                }
            }

            var orderedCategories = categories
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .OrderBy(c => c, StringComparer.OrdinalIgnoreCase)
                .ToList();

            CategoryOptions.Clear();
            foreach (var cat in orderedCategories)
            {
                CategoryOptions.Add(cat);
            }

            if (CurrentRecipe?.Rules != null)
            {
                foreach (var rule in CurrentRecipe.Rules)
                {
                    UpdateRulePropertyOptions(rule);
                }
            }

            UpdateGroupingPropertyOptions();
        }

        private void UpdateRulePropertyOptions(SmartSetRule rule)
        {
            if (rule == null)
            {
                return;
            }

            rule.PropertyOptions.Clear();
            if (!string.IsNullOrWhiteSpace(rule.Category)
                && _propertyOptionsByCategory.TryGetValue(rule.Category, out var options))
            {
                foreach (var option in options)
                {
                    rule.PropertyOptions.Add(option);
                }
            }

            EnsurePropertyOption(rule);
        }

        private void EnsureCategoryOption(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
            {
                return;
            }

            if (ContainsIgnoreCase(CategoryOptions, category))
            {
                return;
            }

            CategoryOptions.Add(category);
        }

        private void EnsurePropertyOption(SmartSetRule rule)
        {
            if (rule == null)
            {
                return;
            }

            var value = rule.Property;
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            if (!ContainsIgnoreCase(rule.PropertyOptions, value))
            {
                rule.PropertyOptions.Insert(0, value);
            }
        }

        private void UpdateGroupingPropertyOptions()
        {
            var grouping = CurrentRecipe?.Grouping;
            if (grouping == null)
            {
                GroupByPropertyOptions.Clear();
                ThenByPropertyOptions.Clear();
                return;
            }

            UpdateGroupingPropertyOptionsFor(grouping.GroupByCategory, grouping.GroupByProperty, GroupByPropertyOptions);
            UpdateGroupingPropertyOptionsFor(grouping.ThenByCategory, grouping.ThenByProperty, ThenByPropertyOptions);
        }

        private void UpdateGroupingPropertyOptionsFor(string category, string selectedValue, ObservableCollection<string> targetOptions)
        {
            targetOptions.Clear();

            if (!string.IsNullOrWhiteSpace(category)
                && _propertyOptionsByCategory.TryGetValue(category, out var options))
            {
                foreach (var option in options)
                {
                    targetOptions.Add(option);
                }
            }

            EnsurePropertyOption(targetOptions, selectedValue);
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

        private static bool ContainsIgnoreCase(IEnumerable<string> items, string value)
        {
            if (items == null || string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return items.Any(item => string.Equals(item, value, StringComparison.OrdinalIgnoreCase));
        }

        private void InitRecipeStore()
        {
            var baseDir = MicroEngStorageSettings.GetDataSubdirectory("SmartSetRecipes");
            var legacyDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "Autodesk",
                "Navisworks Manage 2025",
                "Plugins",
                "MicroEng.Navisworks",
                "SmartSets",
                "Recipes");

            TryMigrateLegacyRecipes(legacyDir, baseDir);

            _recipeStore = new SmartSetRecipeStore(baseDir);
        }

        private static void TryMigrateLegacyRecipes(string legacyDir, string newDir)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(legacyDir) || string.IsNullOrWhiteSpace(newDir))
                {
                    return;
                }

                if (!Directory.Exists(legacyDir))
                {
                    return;
                }

                Directory.CreateDirectory(newDir);
                var hasNewRecipes = Directory.EnumerateFiles(newDir, "*.json", SearchOption.TopDirectoryOnly).Any();
                if (hasNewRecipes)
                {
                    return;
                }

                foreach (var legacyFile in Directory.EnumerateFiles(legacyDir, "*.json", SearchOption.TopDirectoryOnly))
                {
                    var fileName = Path.GetFileName(legacyFile);
                    if (string.IsNullOrWhiteSpace(fileName))
                    {
                        continue;
                    }

                    var targetPath = Path.Combine(newDir, fileName);
                    if (!File.Exists(targetPath))
                    {
                        File.Copy(legacyFile, targetPath, overwrite: false);
                    }
                }
            }
            catch
            {
                // non-fatal migration; keep operating with current directory
            }
        }

        private void RefreshRecipeFiles()
        {
            RecipeFiles.Clear();
            foreach (var file in _recipeStore.ListRecipeFiles())
            {
                RecipeFiles.Add(new RecipeFileItem(file));
            }
        }

        private ScrapeSession GetSelectedSession()
        {
            if (string.IsNullOrWhiteSpace(SelectedScraperProfile))
            {
                return DataScraperCache.LastSession;
            }

            return DataScraperCache.AllSessions
                .Where(s => string.Equals(s.ProfileName, SelectedScraperProfile, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(s => s.Timestamp)
                .FirstOrDefault();
        }

        private static string BuildDefaultFolderPath(string profile)
        {
            if (string.IsNullOrWhiteSpace(profile))
            {
                return "MicroEng/Smart Sets";
            }

            return $"MicroEng/Smart Sets/{profile}";
        }

        internal void AddRule_Click(object sender, RoutedEventArgs e)
        {
            CurrentRecipe.Rules.Add(new SmartSetRule());
        }

        internal void RemoveRule_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedRule != null)
            {
                CurrentRecipe.Rules.Remove(SelectedRule);
            }
            else if (CurrentRecipe.Rules.Count > 0)
            {
                CurrentRecipe.Rules.RemoveAt(CurrentRecipe.Rules.Count - 1);
            }
        }

        internal void DuplicateRule_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedRule == null)
            {
                return;
            }

            CurrentRecipe.Rules.Add(new SmartSetRule
            {
                GroupId = SelectedRule.GroupId,
                Category = SelectedRule.Category,
                Property = SelectedRule.Property,
                Operator = SelectedRule.Operator,
                Value = SelectedRule.Value,
                Enabled = SelectedRule.Enabled
            });
        }

        internal void UseCurrentSelectionScope_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                PreviewStatusText = "No active document.";
                return;
            }

            var selection = doc.CurrentSelection?.SelectedItems;
            if (selection == null || selection.Count == 0)
            {
                PreviewStatusText = "No current selection to use for scope.";
                return;
            }

            CurrentRecipe.ScopeMode = SmartSetScopeMode.CurrentSelection;
            CurrentRecipe.ScopeSelectionSetName = "";
            CurrentRecipe.ScopeModelPaths = new List<List<string>>();
            CurrentRecipe.ScopeFilterCategory = "";
            CurrentRecipe.ScopeFilterProperty = "";
            CurrentRecipe.ScopeFilterValue = "";
            CurrentRecipe.ScopeSummary = $"Current selection ({selection.Count} items)";
            EnforcePreviewCompatibility(showStatus: true);
        }

        internal void ClearScope_Click(object sender, RoutedEventArgs e)
        {
            CurrentRecipe.ScopeMode = SmartSetScopeMode.AllModel;
            CurrentRecipe.ScopeSelectionSetName = "";
            CurrentRecipe.ScopeModelPaths = new List<List<string>>();
            CurrentRecipe.ScopeFilterCategory = "";
            CurrentRecipe.ScopeFilterProperty = "";
            CurrentRecipe.ScopeFilterValue = "";
            CurrentRecipe.ScopeSummary = "Entire model";
            PreviewStatusText = "Scope cleared.";
        }

        internal void PickScope_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                PreviewStatusText = "No active document.";
                return;
            }

            IEnumerable<ScrapedPropertyDescriptor> properties = null;
            var session = GetSelectedSession();
            if (session != null)
            {
                properties = new DataScraperSessionAdapter(session).Properties.ToList();
            }

            var picker = new SmartSetScopePickerWindow(doc, properties, CurrentRecipe)
            {
                Owner = Window.GetWindow(this)
            };

            if (picker.ShowDialog() == true && picker.Result != null)
            {
                ApplyScopePickerResult(picker.Result);
            }
        }

        private void ApplyScopePickerResult(SmartSetScopePickerResult result)
        {
            if (result == null)
            {
                return;
            }

            CurrentRecipe.ScopeSelectionSetName = "";
            CurrentRecipe.ScopeModelPaths = result.ModelPaths ?? new List<List<string>>();
            CurrentRecipe.ScopeFilterCategory = result.FilterCategory ?? "";
            CurrentRecipe.ScopeFilterProperty = result.FilterProperty ?? "";
            CurrentRecipe.ScopeFilterValue = result.FilterValue ?? "";
            CurrentRecipe.SearchInMode = result.SearchInMode;
            CurrentRecipe.ScopeMode = result.ScopeMode;

            UpdateScopeSummary(CurrentRecipe);
            OnPropertyChanged(nameof(CanUseFastPreview));
            EnforcePreviewCompatibility(showStatus: true);
        }

        internal void UseCurrentGroupingScope_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null || CurrentRecipe?.Grouping == null)
            {
                return;
            }

            var selection = doc.CurrentSelection?.SelectedItems;
            if (selection == null || selection.Count == 0)
            {
                return;
            }

            var grouping = CurrentRecipe.Grouping;
            grouping.ScopeMode = SmartSetScopeMode.CurrentSelection;
            grouping.ScopeSelectionSetName = "";
            grouping.ScopeModelPaths = new List<List<string>>();
            grouping.ScopeFilterCategory = "";
            grouping.ScopeFilterProperty = "";
            grouping.ScopeFilterValue = "";
            grouping.ScopeSummary = $"Current selection ({selection.Count} items)";
        }

        internal void ClearGroupingScope_Click(object sender, RoutedEventArgs e)
        {
            var grouping = CurrentRecipe?.Grouping;
            if (grouping == null)
            {
                return;
            }

            grouping.ScopeMode = SmartSetScopeMode.AllModel;
            grouping.ScopeSelectionSetName = "";
            grouping.ScopeModelPaths = new List<List<string>>();
            grouping.ScopeFilterCategory = "";
            grouping.ScopeFilterProperty = "";
            grouping.ScopeFilterValue = "";
            grouping.ScopeSummary = "Entire model";
        }

        internal void PickGroupingScope_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null || CurrentRecipe?.Grouping == null)
            {
                return;
            }

            IEnumerable<ScrapedPropertyDescriptor> properties = null;
            var session = GetSelectedSession();
            if (session != null)
            {
                properties = new DataScraperSessionAdapter(session).Properties.ToList();
            }

            var grouping = CurrentRecipe.Grouping;
            var tempRecipe = new SmartSetRecipe
            {
                SearchInMode = grouping.SearchInMode,
                ScopeMode = grouping.ScopeMode,
                ScopeSelectionSetName = grouping.ScopeSelectionSetName,
                ScopeModelPaths = grouping.ScopeModelPaths?.Select(p => p?.ToList() ?? new List<string>()).ToList()
                    ?? new List<List<string>>(),
                ScopeFilterCategory = grouping.ScopeFilterCategory,
                ScopeFilterProperty = grouping.ScopeFilterProperty,
                ScopeFilterValue = grouping.ScopeFilterValue
            };

            var picker = new SmartSetScopePickerWindow(doc, properties, tempRecipe)
            {
                Owner = Window.GetWindow(this)
            };

            if (picker.ShowDialog() == true && picker.Result != null)
            {
                ApplyGroupingScopePickerResult(picker.Result);
            }
        }

        private void ApplyGroupingScopePickerResult(SmartSetScopePickerResult result)
        {
            if (result == null || CurrentRecipe?.Grouping == null)
            {
                return;
            }

            var grouping = CurrentRecipe.Grouping;
            grouping.ScopeSelectionSetName = "";
            grouping.ScopeModelPaths = result.ModelPaths ?? new List<List<string>>();
            grouping.ScopeFilterCategory = result.FilterCategory ?? "";
            grouping.ScopeFilterProperty = result.FilterProperty ?? "";
            grouping.ScopeFilterValue = result.FilterValue ?? "";
            grouping.SearchInMode = result.SearchInMode;
            grouping.ScopeMode = result.ScopeMode;

            UpdateGroupingScopeSummary(grouping);
        }

internal void RulesGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var cell = FindParent<DataGridCell>(e.OriginalSource as DependencyObject);
            if (cell == null || cell.IsEditing || cell.IsReadOnly)
            {
                return;
            }

            if (cell.Column is DataGridComboBoxColumn && sender is DataGrid grid && !grid.IsReadOnly)
            {
                if (!grid.IsKeyboardFocusWithin)
                {
                    grid.Focus();
                }

                if (!cell.IsFocused)
                {
                    cell.Focus();
                }

                if (grid.CurrentCell.Item != cell.DataContext || grid.CurrentCell.Column != cell.Column)
                {
                    grid.SelectedItem = cell.DataContext;
                    grid.CurrentCell = new DataGridCellInfo(cell.DataContext, cell.Column);
                }

                if (grid.BeginEdit())
                {
                    e.Handled = true;
                }
            }
        }

        internal void RulesGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            if (e.Column is DataGridComboBoxColumn && e.EditingElement is ComboBox combo)
            {
                combo.Focus();
                combo.IsDropDownOpen = true;
            }
        }

        internal void PickGroupBy_Click(object sender, RoutedEventArgs e)
        {
            PickGroupingProperty(isThenBy: false);
        }

        internal void PickThenBy_Click(object sender, RoutedEventArgs e)
        {
            PickGroupingProperty(isThenBy: true);
        }

        private void PickGroupingProperty(bool isThenBy)
        {
            var session = GetSelectedSession();
            if (session == null)
            {
                MessageBox.Show("No Data Scraper session available. Run Data Scraper first.", "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var view = new DataScraperSessionAdapter(session);
            var vm = new PropertyPickerViewModel(view.Properties);
            var win = new PropertyPickerWindow(vm) { Owner = Window.GetWindow(this) };

            if (win.ShowDialog() == true && win.Selected != null)
            {
                if (isThenBy)
                {
                    CurrentRecipe.Grouping.ThenByCategory = win.Selected.Category;
                    CurrentRecipe.Grouping.ThenByProperty = win.Selected.Name;
                }
                else
                {
                    CurrentRecipe.Grouping.GroupByCategory = win.Selected.Category;
                    CurrentRecipe.Grouping.GroupByProperty = win.Selected.Name;
                }
            }
        }

        internal async void Preview_Click(object sender, RoutedEventArgs e)
        {
            _previewCts?.Cancel();
            _previewCts = new CancellationTokenSource();

            var rules = CurrentRecipe.Rules.ToList();
            _lastPreviewResults = null;
            OnPropertyChanged(nameof(CanSelectPreviewResults));

            if (!SmartSetFastPreviewService.IsCompatibleWithFastPreview(CurrentRecipe))
            {
                EnforcePreviewCompatibility(showStatus: true);
            }

            if (UseFastPreview)
            {
                var session = GetSelectedSession();
                if (session == null)
                {
                    PreviewStatusText = "No Data Scraper session available.";
                    return;
                }

                var view = new DataScraperSessionAdapter(session);
                var ct = _previewCts.Token;
                var sw = Stopwatch.StartNew();

                try
                {
                    PreviewStatusText = "Previewing (fast)...";
                    PreviewDetailsText = "";

                    var result = await Task.Run(() =>
                        _fastPreviewService.Evaluate(view, rules, samplePathsToReturn: 25, ct), ct);

                    sw.Stop();
                    PreviewStatusText = $"Matches (fast): {result.EstimatedMatchCount} ({sw.ElapsedMilliseconds} ms)";
                    var sessionLabel = string.IsNullOrWhiteSpace(result.SessionLabel)
                        ? string.Empty
                        : $"Session: {result.SessionLabel}";
                    var samples = string.Join(Environment.NewLine, result.SampleItemPaths);
                    if (!string.IsNullOrWhiteSpace(sessionLabel) && !string.IsNullOrWhiteSpace(samples))
                    {
                        PreviewDetailsText = sessionLabel + Environment.NewLine + samples;
                    }
                    else
                    {
                        PreviewDetailsText = string.IsNullOrWhiteSpace(sessionLabel) ? samples : sessionLabel;
                    }
                    MicroEngActions.Log($"SmartSets preview (fast): {result.EstimatedMatchCount} in {sw.ElapsedMilliseconds} ms");
                }
                catch (OperationCanceledException)
                {
                    PreviewStatusText = "Preview cancelled.";
                }
                catch (Exception ex)
                {
                    PreviewStatusText = "Preview failed: " + ex.Message;
                }

                return;
            }

            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                PreviewStatusText = "No active document.";
                return;
            }

            var realSw = Stopwatch.StartNew();
            PreviewStatusText = "Previewing (live)...";
            PreviewDetailsText = "";

            try
            {
                var preview = _navisworksService.Preview(doc, CurrentRecipe, rules);
                _lastPreviewResults = preview.Results;
                realSw.Stop();

                var filterNote = preview.UsedPostFilter ? " (post-filtered)" : "";
                PreviewStatusText = $"Matches (live): {preview.Count}{filterNote} ({realSw.ElapsedMilliseconds} ms)";
                PreviewDetailsText = preview.UsedPostFilter
                    ? "Some rules required post-filtering. Selection may take longer for large models."
                    : "";

                MicroEngActions.Log($"SmartSets preview (live): {preview.Count} in {realSw.ElapsedMilliseconds} ms");
                OnPropertyChanged(nameof(CanSelectPreviewResults));
            }
            catch (Exception ex)
            {
                PreviewStatusText = "Preview failed: " + ex.Message;
            }
        }

        internal void SelectResults_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null || _lastPreviewResults == null)
            {
                return;
            }

            var selection = doc.CurrentSelection.SelectedItems;
            selection.Clear();
            selection.AddRange(_lastPreviewResults);
        }

        internal void Generate_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("No active document.", "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (string.IsNullOrWhiteSpace(CurrentRecipe.FolderPath))
            {
                CurrentRecipe.FolderPath = BuildDefaultFolderPath(SelectedScraperProfile);
            }

            try
            {
                var mode = CurrentRecipe.SearchSetMode;
                if (mode == SmartSetSearchSetMode.Single && CurrentRecipe.GenerateMultipleSearchSets)
                {
                    mode = SmartSetSearchSetMode.SplitByValue;
                }

                if (mode == SmartSetSearchSetMode.SplitByValue)
                {
                    var session = GetSelectedSession();
                    if (session == null)
                    {
                        MessageBox.Show("No Data Scraper session available. Run Data Scraper first.", "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    var enabledRules = CurrentRecipe.Rules?.Where(r => r != null && r.Enabled).ToList() ?? new List<SmartSetRule>();
                    var definedRules = enabledRules
                        .Where(r => r.Operator == SmartSetOperator.Defined
                            && !string.IsNullOrWhiteSpace(r.Category)
                            && !string.IsNullOrWhiteSpace(r.Property))
                        .ToList();
                    if (definedRules.Count != 1)
                    {
                        MessageBox.Show("Multiple Search Sets requires exactly one Defined rule with Category and Property.", "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    var groupIds = enabledRules
                        .Select(r => string.IsNullOrWhiteSpace(r.GroupId) ? "A" : r.GroupId.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    if (groupIds.Count > 1)
                    {
                        MessageBox.Show("Multiple Search Sets supports a single group only.", "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    var created = _navisworksService.GenerateSplitSearchSets(doc, CurrentRecipe, session, msg => MicroEngActions.Log(msg));
                    MicroEngActions.Log($"SmartSets: generated {created} sets for recipe: {CurrentRecipe.Name}");
                    PreviewStatusText = created > 0 ? "Generated sets." : "No sets generated.";
                    ShowSnackbar("Smart Sets generated",
                        created > 0 ? $"Created {created} sets." : "No sets were created.",
                        created > 0 ? WpfUiControls.ControlAppearance.Success : WpfUiControls.ControlAppearance.Caution,
                        created > 0 ? WpfUiControls.SymbolRegular.CheckmarkCircle24 : WpfUiControls.SymbolRegular.Info24);
                    FlashSuccess(sender as System.Windows.Controls.Button);
                    return;
                }
                if (mode == SmartSetSearchSetMode.ExpandValuesSingleSet)
                {
                    var session = GetSelectedSession();
                    if (session == null)
                    {
                        MessageBox.Show("No Data Scraper session available. Run Data Scraper first.", "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    var enabledRules = CurrentRecipe.Rules?.Where(r => r != null && r.Enabled).ToList() ?? new List<SmartSetRule>();
                    var definedRules = enabledRules
                        .Where(r => r.Operator == SmartSetOperator.Defined
                            && !string.IsNullOrWhiteSpace(r.Category)
                            && !string.IsNullOrWhiteSpace(r.Property))
                        .ToList();
                    if (definedRules.Count != 1)
                    {
                        MessageBox.Show("Expand Values requires exactly one Defined rule with Category and Property.", "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    var groupIds = enabledRules
                        .Select(r => string.IsNullOrWhiteSpace(r.GroupId) ? "A" : r.GroupId.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    if (groupIds.Count > 1)
                    {
                        MessageBox.Show("Expand Values supports a single group only.", "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }

                    var created = _navisworksService.GenerateExpandedSearchSet(doc, CurrentRecipe, session, msg => MicroEngActions.Log(msg));
                    MicroEngActions.Log($"SmartSets: generated {created} sets for recipe: {CurrentRecipe.Name}");
                    PreviewStatusText = created > 0 ? "Generated sets." : "No sets generated.";
                    ShowSnackbar("Smart Sets generated",
                        created > 0 ? $"Created {created} sets." : "No sets were created.",
                        created > 0 ? WpfUiControls.ControlAppearance.Success : WpfUiControls.ControlAppearance.Caution,
                        created > 0 ? WpfUiControls.SymbolRegular.CheckmarkCircle24 : WpfUiControls.SymbolRegular.Info24);
                    FlashSuccess(sender as System.Windows.Controls.Button);
                    return;
                }

                _navisworksService.Generate(doc, CurrentRecipe, msg => MicroEngActions.Log(msg));
                MicroEngActions.Log("SmartSets: generated sets for recipe: " + CurrentRecipe.Name);
                PreviewStatusText = "Generated sets.";
                ShowSnackbar("Smart Sets generated",
                    "Generated sets for the current recipe.",
                    WpfUiControls.ControlAppearance.Success,
                    WpfUiControls.SymbolRegular.CheckmarkCircle24);
                FlashSuccess(sender as System.Windows.Controls.Button);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate failed: " + ex.Message, "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Error);
                ShowSnackbar("Generate failed",
                    ex.Message,
                    WpfUiControls.ControlAppearance.Danger,
                    WpfUiControls.SymbolRegular.ErrorCircle24);
            }
        }

        internal void PreviewGroups_Click(object sender, RoutedEventArgs e)
        {
            GroupRows.Clear();

            var session = GetSelectedSession();
            if (session == null)
            {
                PreviewStatusText = "No Data Scraper session available.";
                return;
            }

            var grouping = CurrentRecipe.Grouping;
            var rows = SmartSetGroupingEngine.BuildGroups(
                session,
                grouping.GroupByCategory,
                grouping.GroupByProperty,
                grouping.UseThenBy,
                grouping.ThenByCategory,
                grouping.ThenByProperty,
                grouping.MinCount,
                grouping.IncludeBlanks);

            foreach (var row in rows.Take(grouping.MaxGroups))
            {
                GroupRows.Add(row);
            }

            if (GroupRows.Count == 0)
            {
                PreviewStatusText = "No groups found. Adjust Group By/Then By or lower Min Count.";
            }
            else
            {
                PreviewStatusText = $"Found {GroupRows.Count} groups.";
            }
        }

        internal void GenerateGroups_Click(object sender, RoutedEventArgs e)
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("No active document.", "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var grouping = CurrentRecipe.Grouping;
            if (GroupRows.Count == 0)
            {
                PreviewGroups_Click(sender, e);
            }
            if (GroupRows.Count == 0)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(grouping.OutputName))
            {
                grouping.OutputName = "New Grouping";
            }
            if (string.IsNullOrWhiteSpace(grouping.OutputFolderPath))
            {
                grouping.OutputFolderPath = BuildDefaultFolderPath(SelectedScraperProfile);
            }

            _navisworksService.GenerateGroupedSearchSets(
                doc,
                grouping,
                grouping.GroupByCategory,
                grouping.GroupByProperty,
                grouping.UseThenBy,
                grouping.ThenByCategory,
                grouping.ThenByProperty,
                GroupRows.ToList(),
                msg => MicroEngActions.Log(msg));

            MicroEngActions.Log("SmartSets: generated grouped search sets.");
            PreviewStatusText = "Generated grouped search sets.";
            ShowSnackbar("Grouped sets generated",
                $"Generated {GroupRows.Count} grouped set(s).",
                WpfUiControls.ControlAppearance.Success,
                WpfUiControls.SymbolRegular.CheckmarkCircle24);
            FlashSuccess(sender as System.Windows.Controls.Button);
        }

        internal void AnalyzeSelection_Click(object sender, RoutedEventArgs e)
        {
            SelectionSuggestions.Clear();

            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                SelectionSuggestions.Add(new SmartSetSuggestion { Category = "", Property = "", Operator = SmartSetOperator.Equals, Value = "No active document." });
                return;
            }

            var selection = doc.CurrentSelection.SelectedItems;
            var suggestions = SmartSetInferenceEngine.AnalyzeSelection(selection, 10);

            foreach (var s in suggestions)
            {
                SelectionSuggestions.Add(s);
            }

            if (SelectionSuggestions.Count == 0)
            {
                SelectionSuggestions.Add(new SmartSetSuggestion { Category = "", Property = "", Operator = SmartSetOperator.Equals, Value = "No strong suggestions found." });
            }
        }

        internal void ApplySuggestion_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedSuggestion == null)
            {
                return;
            }

            CurrentRecipe.Rules.Add(new SmartSetRule
            {
                Category = SelectedSuggestion.Category,
                Property = SelectedSuggestion.Property,
                Operator = SelectedSuggestion.Operator,
                Value = SelectedSuggestion.Value,
                Enabled = true
            });
        }

        internal void RunPack_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedPack == null)
            {
                return;
            }

            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("No active document.", "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var session = GetSelectedSession();
            var view = session != null ? new DataScraperSessionAdapter(session) : null;
            var missing = SelectedPack.CheckMissingProperties(view);
            if (missing.Count > 0)
            {
                MessageBox.Show("Pack references missing properties: " + string.Join(", ", missing), "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            foreach (var recipe in SelectedPack.BuildRecipes(SelectedScraperProfile))
            {
                _navisworksService.Generate(doc, recipe, msg => MicroEngActions.Log(msg));
            }

            MicroEngActions.Log("SmartSets: pack executed: " + SelectedPack.Name);
            ShowSnackbar("Pack executed",
                $"Generated sets for pack '{SelectedPack.Name}'.",
                WpfUiControls.ControlAppearance.Success,
                WpfUiControls.SymbolRegular.CheckmarkCircle24);
            FlashSuccess(sender as System.Windows.Controls.Button);
        }

        internal void SaveRecipe_Click(object sender, RoutedEventArgs e)
        {
            var path = _recipeStore.GetDefaultPathForRecipe(CurrentRecipe.Name);
            _recipeStore.Save(CurrentRecipe, path);
            RefreshRecipeFiles();
            MicroEngActions.Log("SmartSets: recipe saved: " + path);
        }

        internal void SaveRecipeAs_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save Smart Set Recipe",
                FileName = SmartSetRecipeStore.MakeSafeFileName(CurrentRecipe.Name) + ".json",
                Filter = "Recipe Files (*.json)|*.json",
                InitialDirectory = _recipeStore.RecipesDirectory
            };

            if (dlg.ShowDialog() == true)
            {
                _recipeStore.Save(CurrentRecipe, dlg.FileName);
                RefreshRecipeFiles();
                MicroEngActions.Log("SmartSets: recipe saved: " + dlg.FileName);
            }
        }

        internal void LoadRecipe_Click(object sender, RoutedEventArgs e)
        {
            string path = SelectedRecipeFile?.Path;

            if (string.IsNullOrWhiteSpace(path))
            {
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Load Smart Set Recipe",
                    Filter = "Recipe Files (*.json)|*.json",
                    InitialDirectory = _recipeStore.RecipesDirectory
                };

                if (dlg.ShowDialog() == true)
                {
                    path = dlg.FileName;
                }
            }

            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            var recipe = _recipeStore.Load(path);
            if (recipe == null)
            {
                return;
            }

            if (recipe.Rules == null)
            {
                recipe.Rules = new ObservableCollection<SmartSetRule>();
            }

            if (recipe.Grouping == null)
            {
                recipe.Grouping = new SmartSetGroupingSpec();
            }

            if (recipe.Rules.Count == 0)
            {
                recipe.Rules.Add(new SmartSetRule());
            }

            CurrentRecipe = recipe;
            OnPropertyChanged(nameof(CurrentRecipe));

            if (!string.IsNullOrWhiteSpace(recipe.DataScraperProfile))
            {
                SelectedScraperProfile = recipe.DataScraperProfile;
            }
            else if (string.IsNullOrWhiteSpace(CurrentRecipe.FolderPath))
            {
                CurrentRecipe.FolderPath = BuildDefaultFolderPath(SelectedScraperProfile);
            }

            RefreshRecipeFiles();
            RefreshSavedSelectionSets();
            MicroEngActions.Log("SmartSets: recipe loaded: " + path);
        }

        private void RefreshProfiles_Click(object sender, RoutedEventArgs e)
        {
            RefreshScraperProfiles();
            RefreshSavedSelectionSets();
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

        private void OpenDataScraper_Click(object sender, RoutedEventArgs e)
        {
            MicroEngActions.TryShowDataScraper(null, out _);
        }

        private void BuildPackList()
        {
            Packs.Clear();

            Packs.Add(new SmartSetPackDefinition(
                "QA: Missing Level",
                "Find items with blank Level property.",
                new[]
                {
                    SmartSetPackDefinition.MakeRule("Element", "Level", SmartSetOperator.Undefined, "")
                }));

            Packs.Add(new SmartSetPackDefinition(
                "QA: Missing Category",
                "Find items missing the Element Properties Category.",
                new[]
                {
                    SmartSetPackDefinition.MakeRule("Element Properties", "Category", SmartSetOperator.Undefined, "")
                }));

            Packs.Add(new SmartSetPackDefinition(
                "Handover: Missing Asset Tag",
                "Identify assets with blank Asset Tag value.",
                new[]
                {
                    SmartSetPackDefinition.MakeRule("Element", "Asset Tag", SmartSetOperator.Undefined, "")
                }));

            Packs.Add(new SmartSetPackDefinition(
                "MEP: Valves",
                "Search for items where Type contains Valve.",
                new[]
                {
                    SmartSetPackDefinition.MakeRule("Element", "Type", SmartSetOperator.Contains, "Valve")
                }));

            Packs.Add(new SmartSetPackDefinition(
                "MEP: Dampers",
                "Search for items where Type contains Damper.",
                new[]
                {
                    SmartSetPackDefinition.MakeRule("Element", "Type", SmartSetOperator.Contains, "Damper")
                }));

            Packs.Add(new SmartSetPackDefinition(
                "MEP: Panels",
                "Search for items where Type contains Panel.",
                new[]
                {
                    SmartSetPackDefinition.MakeRule("Element", "Type", SmartSetOperator.Contains, "Panel")
                }));

            Packs.Add(new SmartSetPackDefinition(
                "QA: Missing Mark",
                "Find items with blank Mark value.",
                new[]
                {
                    SmartSetPackDefinition.MakeRule("Element", "Mark", SmartSetOperator.Undefined, "")
                }));

            Packs.Add(new SmartSetPackDefinition(
                "QA: Missing Type",
                "Find items with blank Type value.",
                new[]
                {
                    SmartSetPackDefinition.MakeRule("Element", "Type", SmartSetOperator.Undefined, "")
                }));

            Packs.Add(new SmartSetPackDefinition(
                "QA: Duplicate Candidates",
                "Items with the same Name value (group by Item/Name).",
                new[]
                {
                    SmartSetPackDefinition.MakeRule("Item", "Name", SmartSetOperator.Defined, "")
                }));
        }

        private static ConditionOption[] BuildConditionOptions()
        {
            return new[]
            {
                new ConditionOption(SmartSetOperator.Equals, "="),
                new ConditionOption(SmartSetOperator.NotEquals, "not equals"),
                new ConditionOption(SmartSetOperator.Contains, "Contains"),
                new ConditionOption(SmartSetOperator.Wildcard, "Wildcard"),
                new ConditionOption(SmartSetOperator.Defined, "Defined"),
                new ConditionOption(SmartSetOperator.Undefined, "Undefined")
            };
        }

        private static SearchSetModeOption[] BuildSearchSetModeOptions()
        {
            return new[]
            {
                new SearchSetModeOption(SmartSetSearchSetMode.Single, "Single Search Set"),
                new SearchSetModeOption(SmartSetSearchSetMode.SplitByValue, "Multiple Search Sets (split by value)"),
                new SearchSetModeOption(SmartSetSearchSetMode.ExpandValuesSingleSet, "Single Search Set (expand values)")
            };
        }

        private static SearchInModeOption[] BuildSearchInModeOptions()
        {
            return new[]
            {
                new SearchInModeOption(SmartSetSearchInMode.Standard, "Standard"),
                new SearchInModeOption(SmartSetSearchInMode.Compact, "Compact"),
                new SearchInModeOption(SmartSetSearchInMode.Properties, "Properties")
            };
        }

        private static ScopeModeOption[] BuildScopeModeOptions()
        {
            return new[]
            {
                new ScopeModeOption(SmartSetScopeMode.AllModel, "All model"),
                new ScopeModeOption(SmartSetScopeMode.CurrentSelection, "Current selection"),
                new ScopeModeOption(SmartSetScopeMode.SavedSelectionSet, "Saved selection set"),
                new ScopeModeOption(SmartSetScopeMode.ModelTree, "Tree selection"),
                new ScopeModeOption(SmartSetScopeMode.PropertyFilter, "Property filter")
            };
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private static T FindParent<T>(DependencyObject source) where T : DependencyObject
        {
            var current = source;
            while (current != null)
            {
                if (current is T match)
                {
                    return match;
                }

                current = VisualTreeHelper.GetParent(current);
            }

            return null;
        }
    }

    public sealed class RecipeFileItem
    {
        public RecipeFileItem(string path)
        {
            Path = path;
            DisplayName = System.IO.Path.GetFileNameWithoutExtension(path) ?? path;
        }

        public string Path { get; }
        public string DisplayName { get; }
    }

    public sealed class ConditionOption
    {
        public ConditionOption(SmartSetOperator value, string label)
        {
            Value = value;
            Label = label ?? value.ToString();
        }

        public SmartSetOperator Value { get; }
        public string Label { get; }
    }

    public sealed class SearchSetModeOption
    {
        public SearchSetModeOption(SmartSetSearchSetMode value, string label)
        {
            Value = value;
            Label = label ?? value.ToString();
        }

        public SmartSetSearchSetMode Value { get; }
        public string Label { get; }
    }

    public sealed class SearchInModeOption
    {
        public SearchInModeOption(SmartSetSearchInMode value, string label)
        {
            Value = value;
            Label = label ?? value.ToString();
        }

        public SmartSetSearchInMode Value { get; }
        public string Label { get; }
    }

    public sealed class ScopeModeOption
    {
        public ScopeModeOption(SmartSetScopeMode value, string label)
        {
            Value = value;
            Label = label ?? value.ToString();
        }

        public SmartSetScopeMode Value { get; }
        public string Label { get; }
    }

}

