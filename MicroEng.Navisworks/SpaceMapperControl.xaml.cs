using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Wpf.Ui.Abstractions;

namespace MicroEng.Navisworks
{
    public partial class SpaceMapperControl : UserControl
    {
        static SpaceMapperControl()
        {
            AssemblyResolver.EnsureRegistered();
        }

        internal ObservableCollection<SpaceMapperTargetRule> TargetRules { get; } = new();
        internal ObservableCollection<SpaceMapperMappingDefinition> Mappings { get; } = new();
        internal ObservableCollection<ZoneSummary> ZoneSummaries { get; } = new();

        private readonly SpaceMapperTemplateStore _templateStore = new(AppDomain.CurrentDomain.BaseDirectory);
        private readonly SpaceMapperStepSetupPage _setupPage;
        private readonly SpaceMapperStepZonesTargetsPage _zonesPage;
        private readonly SpaceMapperStepProcessingPage _processingPage;
        private readonly SpaceMapperStepMappingPage _mappingPage;
        private readonly SpaceMapperStepResultsPage _resultsPage;
        private bool _handlersWired;
        private bool _initialized;

        public SpaceMapperControl()
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            DataContext = this;
            _setupPage = new SpaceMapperStepSetupPage(this);
            _zonesPage = new SpaceMapperStepZonesTargetsPage(this);
            _processingPage = new SpaceMapperStepProcessingPage(this);
            _mappingPage = new SpaceMapperStepMappingPage(this);
            _resultsPage = new SpaceMapperStepResultsPage(this);

            SpaceMapperNav.SetPageProviderService(new SpaceMapperStepPageProvider(
                _setupPage,
                _zonesPage,
                _processingPage,
                _mappingPage,
                _resultsPage));

            Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            if (_initialized)
            {
                return;
            }

            _initialized = true;
            InitializeUi();
            SpaceMapperNav.Navigate(typeof(SpaceMapperStepSetupPage));
        }

        private void InitializeUi()
        {
            _setupPage.ProfileCombo.Items.Clear();
            var profiles = DataScraperCache.AllSessions
                .Select(s => string.IsNullOrWhiteSpace(s.ProfileName) ? "Default" : s.ProfileName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(p => p)
                .ToList();
            foreach (var p in profiles) _setupPage.ProfileCombo.Items.Add(p);
            if (_setupPage.ProfileCombo.Items.Count == 0) _setupPage.ProfileCombo.Items.Add("Default");
            _setupPage.ProfileCombo.SelectedIndex = 0;

            _setupPage.ScopeCombo.ItemsSource = Enum.GetValues(typeof(SpaceMapperScope));
            _setupPage.ScopeCombo.SelectedItem = SpaceMapperScope.EntireModel;

            _zonesPage.ZoneSourceCombo.ItemsSource = Enum.GetValues(typeof(ZoneSourceType));
            _zonesPage.ZoneSourceCombo.SelectedItem = ZoneSourceType.DataScraperZones;

            _zonesPage.TargetSourceCombo.ItemsSource = Enum.GetValues(typeof(TargetSourceType));
            _zonesPage.TargetSourceCombo.SelectedItem = TargetSourceType.EntireModel;

            _processingPage.ProcessingModeCombo.ItemsSource = new[] { SpaceMapperProcessingMode.CpuNormal };
            _processingPage.ProcessingModeCombo.SelectedItem = SpaceMapperProcessingMode.CpuNormal;
            _processingPage.ProcessingModeCombo.IsEnabled = false;

            if (TargetRules.Count == 0)
                TargetRules.Add(new SpaceMapperTargetRule { Name = "Level 0", TargetType = SpaceMapperTargetType.SelectionTreeLevel, MinTreeLevel = 0, MaxTreeLevel = 0 });
            if (Mappings.Count == 0)
                Mappings.Add(new SpaceMapperMappingDefinition { Name = "Zone Name", ZoneCategory = "Zone", ZonePropertyName = "Name", TargetPropertyName = "Zone Name" });

            if (!_handlersWired)
            {
                AddHandlers();
                _handlersWired = true;
            }
        }

        private void AddHandlers()
        {
            _setupPage.RefreshProfilesButton.Click += (s, e) => InitializeUi();
            _setupPage.RunScraperButton.Click += (s, e) => OpenScraper();
            _setupPage.RunButton.Click += (s, e) => RunSpaceMapper();
            _zonesPage.AddRuleButton.Click += (s, e) => TargetRules.Add(new SpaceMapperTargetRule { Name = $"Rule {TargetRules.Count + 1}" });
            _zonesPage.DeleteRuleButton.Click += (s, e) =>
            {
                if (_zonesPage.TargetRulesGrid.SelectedItem is SpaceMapperTargetRule rule)
                    TargetRules.Remove(rule);
            };
            _mappingPage.AddMappingButton.Click += (s, e) => Mappings.Add(new SpaceMapperMappingDefinition { Name = $"Mapping {Mappings.Count + 1}", TargetPropertyName = "Zone Name" });
            _mappingPage.DeleteMappingButton.Click += (s, e) =>
            {
                if (_mappingPage.MappingGrid.SelectedItem is SpaceMapperMappingDefinition map)
                    Mappings.Remove(map);
            };
            _mappingPage.SaveTemplateButton.Click += (s, e) => SaveTemplate();
            _mappingPage.LoadTemplateButton.Click += (s, e) => LoadTemplate();
            _resultsPage.ExportStatsButton.Click += (s, e) => ExportStats();
        }

        private void OpenScraper()
        {
            var profile = _setupPage.ProfileCombo.SelectedItem?.ToString() ?? "Default";
            try
            {
                MicroEngActions.TryShowDataScraper(profile, out _);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open Data Scraper: {ex.Message}", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RunSpaceMapper()
        {
            try
            {
                var settings = BuildSettings();
                var request = new SpaceMapperRequest
                {
                    ProfileName = _setupPage.ProfileCombo.SelectedItem?.ToString() ?? "Default",
                    Scope = _setupPage.ScopeCombo.SelectedItem is SpaceMapperScope sc ? sc : SpaceMapperScope.EntireModel,
                    ZoneSource = _zonesPage.ZoneSourceCombo.SelectedItem is ZoneSourceType zs ? zs : ZoneSourceType.DataScraperZones,
                    ZoneSetName = _zonesPage.ZoneSetBox.Text,
                    TargetSource = _zonesPage.TargetSourceCombo.SelectedItem is TargetSourceType ts ? ts : TargetSourceType.EntireModel,
                    TargetRules = TargetRules.ToList(),
                    Mappings = Mappings.ToList(),
                    ProcessingSettings = settings
                };

                var service = new SpaceMapperService(MicroEngActions.Log);
                var result = service.Run(request);
                ShowResults(result);
                SpaceMapperNav.Navigate(typeof(SpaceMapperStepResultsPage));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Space Mapper failed: {ex.Message}", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private SpaceMapperProcessingSettings BuildSettings()
        {
            return new SpaceMapperProcessingSettings
            {
                ProcessingMode = _processingPage.ProcessingModeCombo.SelectedItem is SpaceMapperProcessingMode pm ? pm : SpaceMapperProcessingMode.Auto,
                TreatPartialAsContained = _processingPage.TreatPartialCheck.IsChecked == true,
                TagPartialSeparately = _processingPage.TagPartialCheck.IsChecked == true,
                EnableMultipleZones = _processingPage.EnableMultiZoneCheck.IsChecked == true,
                Offset3D = ParseDouble(_processingPage.Offset3DBox.Text),
                OffsetTop = ParseDouble(_processingPage.OffsetTopBox.Text),
                OffsetBottom = ParseDouble(_processingPage.OffsetBottomBox.Text),
                OffsetSides = ParseDouble(_processingPage.OffsetSidesBox.Text),
                Units = _processingPage.UnitsBox.Text,
                OffsetMode = _processingPage.OffsetModeBox.Text,
                MaxThreads = ParseInt(_processingPage.MaxThreadsBox.Text),
                BatchSize = ParseInt(_processingPage.BatchSizeBox.Text)
            };
        }

        private static double ParseDouble(string text)
        {
            return double.TryParse(text, out var d) ? d : 0;
        }

        private static int? ParseInt(string text)
        {
            return int.TryParse(text, out var i) && i > 0 ? i : (int?)null;
        }

        private void ShowResults(SpaceMapperRunResult result)
        {
            ZoneSummaries.Clear();
            if (result?.Stats?.ZoneSummaries != null)
            {
                foreach (var z in result.Stats.ZoneSummaries.OrderByDescending(z => z.ContainedCount + z.PartialCount))
                {
                    ZoneSummaries.Add(z);
                }
            }

            var sb = new StringBuilder();
            sb.AppendLine(result?.Message ?? "No result.");
            if (result?.Stats != null)
            {
                sb.AppendLine($"Zones: {result.Stats.ZonesProcessed}, Targets: {result.Stats.TargetsProcessed}");
                sb.AppendLine($"Contained: {result.Stats.ContainedTagged}, Partial: {result.Stats.PartialTagged}, Multi-Zone: {result.Stats.MultiZoneTagged}, Skipped: {result.Stats.Skipped}");
                sb.AppendLine($"Mode: {result.Stats.ModeUsed}, Time: {result.Stats.Elapsed.TotalSeconds:0.00}s");
            }
            _resultsPage.ResultsSummaryBox.Text = sb.ToString();
        }

        private void SaveTemplate()
        {
            var templates = _templateStore.Load();
            var name = PromptText("Template name:", "Space Mapper");
            if (string.IsNullOrWhiteSpace(name)) return;
            var existing = templates.FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));
            var tpl = existing ?? new SpaceMapperTemplate { Name = name };
            tpl.TargetRules = TargetRules.ToList();
            tpl.Mappings = Mappings.ToList();
            tpl.ProcessingSettings = BuildSettings();

            if (existing == null) templates.Add(tpl);
            _templateStore.Save(templates);
        }

        private void LoadTemplate()
        {
            var templates = _templateStore.Load();
            if (!templates.Any())
            {
                MessageBox.Show("No templates saved yet.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var name = PromptText("Template to load:", "Space Mapper");
            if (string.IsNullOrWhiteSpace(name)) return;
            var tpl = templates.FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));
            if (tpl == null)
            {
                MessageBox.Show($"Template '{name}' not found.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            TargetRules.Clear();
            foreach (var r in tpl.TargetRules) TargetRules.Add(r);

            Mappings.Clear();
            foreach (var m in tpl.Mappings) Mappings.Add(m);

            ApplySettings(tpl.ProcessingSettings);
        }

        private void ApplySettings(SpaceMapperProcessingSettings settings)
        {
            if (settings == null) return;
            _processingPage.ProcessingModeCombo.SelectedItem = settings.ProcessingMode;
            _processingPage.TreatPartialCheck.IsChecked = settings.TreatPartialAsContained;
            _processingPage.TagPartialCheck.IsChecked = settings.TagPartialSeparately;
            _processingPage.EnableMultiZoneCheck.IsChecked = settings.EnableMultipleZones;
            _processingPage.Offset3DBox.Text = settings.Offset3D.ToString();
            _processingPage.OffsetTopBox.Text = settings.OffsetTop.ToString();
            _processingPage.OffsetBottomBox.Text = settings.OffsetBottom.ToString();
            _processingPage.OffsetSidesBox.Text = settings.OffsetSides.ToString();
            _processingPage.UnitsBox.Text = settings.Units;
            _processingPage.OffsetModeBox.Text = settings.OffsetMode;
            _processingPage.MaxThreadsBox.Text = settings.MaxThreads?.ToString() ?? "0";
            _processingPage.BatchSizeBox.Text = settings.BatchSize?.ToString() ?? "0";
        }

        private void ExportStats()
        {
            if (ZoneSummaries.Count == 0)
            {
                MessageBox.Show("No stats to export.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV Files|*.csv",
                FileName = "SpaceMapper_Stats.csv"
            };
            if (dlg.ShowDialog() != true) return;

            using var writer = new StreamWriter(dlg.FileName, false, Encoding.UTF8);
            writer.WriteLine("Zone,Contained,Partial");
            foreach (var z in ZoneSummaries)
            {
                writer.WriteLine($"{Escape(z.ZoneName)},{z.ContainedCount},{z.PartialCount}");
            }
            writer.Flush();
            MessageBox.Show("Export complete.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private static string Escape(string value)
        {
            if (value == null) return string.Empty;
            if (value.Contains(",") || value.Contains("\""))
            {
                return $"\"{value.Replace("\"", "\"\"")}\"";
            }
            return value;
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
            var box = new TextBox();
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
            var ok = new Button { Content = "OK", Width = 70, Margin = new Thickness(4), IsDefault = true };
            var cancel = new Button { Content = "Cancel", Width = 70, Margin = new Thickness(4), IsCancel = true };
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

        private sealed class SpaceMapperStepPageProvider : INavigationViewPageProvider
        {
            private readonly Dictionary<Type, object> _pages;

            public SpaceMapperStepPageProvider(params Page[] pages)
            {
                _pages = pages.ToDictionary(p => p.GetType(), p => (object)p);
            }

            public object GetPage(Type pageType)
            {
                _pages.TryGetValue(pageType, out var page);
                return page;
            }
        }
    }
}
