using System;
using System.Collections.ObjectModel;
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
using Autodesk.Navisworks.Api;
using NavisApp = Autodesk.Navisworks.Api.Application;
using MicroEng.Navisworks;

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

            DataScraperCache.SessionAdded += OnSessionAdded;
            Unloaded += (_, __) => DataScraperCache.SessionAdded -= OnSessionAdded;
        }

        public ObservableCollection<string> ScraperProfiles { get; } = new ObservableCollection<string>();
        public ObservableCollection<ConditionOption> ConditionOptions { get; }
        public ObservableCollection<SmartSetOutputType> OutputTypes { get; }
        public ObservableCollection<SmartGroupRow> GroupRows { get; } = new ObservableCollection<SmartGroupRow>();
        public ObservableCollection<SmartSetSuggestion> SelectionSuggestions { get; } = new ObservableCollection<SmartSetSuggestion>();
        public ObservableCollection<SmartSetPackDefinition> Packs { get; } = new ObservableCollection<SmartSetPackDefinition>();
        public ObservableCollection<RecipeFileItem> RecipeFiles { get; } = new ObservableCollection<RecipeFileItem>();

        public SmartSetRecipe CurrentRecipe
        {
            get => _currentRecipe;
            set
            {
                _currentRecipe = value ?? new SmartSetRecipe();
                OnPropertyChanged();
            }
        }

        public string SelectedScraperProfile
        {
            get => _selectedScraperProfile;
            set
            {
                _selectedScraperProfile = value;
                if (CurrentRecipe != null)
                {
                    CurrentRecipe.DataScraperProfile = value ?? "";
                    if (string.IsNullOrWhiteSpace(CurrentRecipe.FolderPath))
                    {
                        CurrentRecipe.FolderPath = BuildDefaultFolderPath(value);
                    }
                }
                OnPropertyChanged();
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
        public bool CanSelectPreviewResults => !UseFastPreview && _lastPreviewResults != null && _lastPreviewResults.Count > 0;

        private void OnSessionAdded(ScrapeSession session)
        {
            Dispatcher.BeginInvoke(new Action(RefreshScraperProfiles));
        }

        private void RefreshScraperProfiles()
        {
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
                if (string.IsNullOrWhiteSpace(SelectedScraperProfile) || !profiles.Contains(SelectedScraperProfile))
                {
                    SelectedScraperProfile = profiles[0];
                }
            }
        }

        private void InitRecipeStore()
        {
            var baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "Autodesk",
                "Navisworks Manage 2025",
                "Plugins",
                "MicroEng.Navisworks",
                "SmartSets",
                "Recipes");

            _recipeStore = new SmartSetRecipeStore(baseDir);
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

        private void AddRule_Click(object sender, RoutedEventArgs e)
        {
            CurrentRecipe.Rules.Add(new SmartSetRule());
        }

        private void RemoveRule_Click(object sender, RoutedEventArgs e)
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

        private void DuplicateRule_Click(object sender, RoutedEventArgs e)
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

        private void PickProperty_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is FrameworkElement element))
            {
                return;
            }

            if (!(element.DataContext is SmartSetRule rule))
            {
                return;
            }

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
                rule.Category = win.Selected.Category;
                rule.Property = win.Selected.Name;

                rule.SampleValues.Clear();
                foreach (var sample in win.Selected.SampleValues.Take(25))
                {
                    rule.SampleValues.Add(sample);
                }

                if (string.IsNullOrWhiteSpace(rule.Value) && win.Selected.SampleValues.Count > 0)
                {
                    rule.Value = win.Selected.SampleValues[0];
                }
            }
        }

        private void RulesGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var cell = FindParent<DataGridCell>(e.OriginalSource as DependencyObject);
            if (cell == null || cell.IsEditing || cell.IsReadOnly)
            {
                return;
            }

            if (cell.Column is DataGridComboBoxColumn && sender is DataGrid grid && !grid.IsReadOnly)
            {
                if (!cell.IsFocused)
                {
                    cell.Focus();
                }

                grid.BeginEdit();
                e.Handled = true;
            }
        }

        private void PickGroupBy_Click(object sender, RoutedEventArgs e)
        {
            PickGroupingProperty(isThenBy: false);
        }

        private void PickThenBy_Click(object sender, RoutedEventArgs e)
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

        private async void Preview_Click(object sender, RoutedEventArgs e)
        {
            _previewCts?.Cancel();
            _previewCts = new CancellationTokenSource();

            var rules = CurrentRecipe.Rules.ToList();
            _lastPreviewResults = null;
            OnPropertyChanged(nameof(CanSelectPreviewResults));

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
                    PreviewDetailsText = string.Join(Environment.NewLine, result.SampleItemPaths);
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
                var preview = _navisworksService.Preview(doc, rules);
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

        private void SelectResults_Click(object sender, RoutedEventArgs e)
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

        private void Generate_Click(object sender, RoutedEventArgs e)
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
                _navisworksService.Generate(doc, CurrentRecipe, msg => MicroEngActions.Log(msg));
                MicroEngActions.Log("SmartSets: generated sets for recipe: " + CurrentRecipe.Name);
                PreviewStatusText = "Generated sets.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate failed: " + ex.Message, "Smart Set Generator", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PreviewGroups_Click(object sender, RoutedEventArgs e)
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
        }

        private void GenerateGroups_Click(object sender, RoutedEventArgs e)
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

            var folder = string.IsNullOrWhiteSpace(CurrentRecipe.FolderPath)
                ? BuildDefaultFolderPath(SelectedScraperProfile)
                : CurrentRecipe.FolderPath;

            _navisworksService.GenerateGroupedSearchSets(
                doc,
                folder,
                CurrentRecipe.Name,
                grouping.GroupByCategory,
                grouping.GroupByProperty,
                grouping.UseThenBy,
                grouping.ThenByCategory,
                grouping.ThenByProperty,
                GroupRows.ToList());

            MicroEngActions.Log("SmartSets: generated grouped search sets.");
            PreviewStatusText = "Generated grouped search sets.";
        }

        private void AnalyzeSelection_Click(object sender, RoutedEventArgs e)
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

        private void ApplySuggestion_Click(object sender, RoutedEventArgs e)
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

        private void RunPack_Click(object sender, RoutedEventArgs e)
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
        }

        private void SaveRecipe_Click(object sender, RoutedEventArgs e)
        {
            var path = _recipeStore.GetDefaultPathForRecipe(CurrentRecipe.Name);
            _recipeStore.Save(CurrentRecipe, path);
            RefreshRecipeFiles();
            MicroEngActions.Log("SmartSets: recipe saved: " + path);
        }

        private void SaveRecipeAs_Click(object sender, RoutedEventArgs e)
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

        private void LoadRecipe_Click(object sender, RoutedEventArgs e)
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
            MicroEngActions.Log("SmartSets: recipe loaded: " + path);
        }

        private void RefreshProfiles_Click(object sender, RoutedEventArgs e)
        {
            RefreshScraperProfiles();
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
}
