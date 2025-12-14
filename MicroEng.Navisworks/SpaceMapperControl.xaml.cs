using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace MicroEng.Navisworks
{
    public partial class SpaceMapperControl : UserControl
    {
        internal ObservableCollection<SpaceMapperTargetRule> TargetRules { get; } = new();
        internal ObservableCollection<SpaceMapperMappingDefinition> Mappings { get; } = new();
        internal ObservableCollection<ZoneSummary> ZoneSummaries { get; } = new();

        private readonly SpaceMapperTemplateStore _templateStore = new(AppDomain.CurrentDomain.BaseDirectory);

        public SpaceMapperControl()
        {
            InitializeComponent();
            DataContext = this;
            Loaded += (_, __) => InitializeUi();
        }

        private void InitializeUi()
        {
            ProfileCombo.Items.Clear();
            var profiles = DataScraperCache.AllSessions
                .Select(s => string.IsNullOrWhiteSpace(s.ProfileName) ? "Default" : s.ProfileName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(p => p)
                .ToList();
            foreach (var p in profiles) ProfileCombo.Items.Add(p);
            if (ProfileCombo.Items.Count == 0) ProfileCombo.Items.Add("Default");
            ProfileCombo.SelectedIndex = 0;

            ScopeCombo.ItemsSource = Enum.GetValues(typeof(SpaceMapperScope));
            ScopeCombo.SelectedItem = SpaceMapperScope.EntireModel;

            ZoneSourceCombo.ItemsSource = Enum.GetValues(typeof(ZoneSourceType));
            ZoneSourceCombo.SelectedItem = ZoneSourceType.DataScraperZones;

            TargetSourceCombo.ItemsSource = Enum.GetValues(typeof(TargetSourceType));
            TargetSourceCombo.SelectedItem = TargetSourceType.EntireModel;

            ProcessingModeCombo.ItemsSource = new[] { SpaceMapperProcessingMode.CpuNormal };
            ProcessingModeCombo.SelectedItem = SpaceMapperProcessingMode.CpuNormal;
            ProcessingModeCombo.IsEnabled = false;

            if (TargetRules.Count == 0)
                TargetRules.Add(new SpaceMapperTargetRule { Name = "Level 0", TargetType = SpaceMapperTargetType.SelectionTreeLevel, MinTreeLevel = 0, MaxTreeLevel = 0 });
            if (Mappings.Count == 0)
                Mappings.Add(new SpaceMapperMappingDefinition { Name = "Zone Name", ZoneCategory = "Zone", ZonePropertyName = "Name", TargetPropertyName = "Zone Name" });

            AddHandlers();
        }

        private void AddHandlers()
        {
            RefreshProfilesButton.Click += (s, e) => InitializeUi();
            RunScraperButton.Click += (s, e) => OpenScraper();
            RunButton.Click += (s, e) => RunSpaceMapper();
            AddRuleButton.Click += (s, e) => TargetRules.Add(new SpaceMapperTargetRule { Name = $"Rule {TargetRules.Count + 1}" });
            DeleteRuleButton.Click += (s, e) =>
            {
                if (TargetRulesGrid.SelectedItem is SpaceMapperTargetRule rule)
                    TargetRules.Remove(rule);
            };
            AddMappingButton.Click += (s, e) => Mappings.Add(new SpaceMapperMappingDefinition { Name = $"Mapping {Mappings.Count + 1}", TargetPropertyName = "Zone Name" });
            DeleteMappingButton.Click += (s, e) =>
            {
                if (MappingGrid.SelectedItem is SpaceMapperMappingDefinition map)
                    Mappings.Remove(map);
            };
            SaveTemplateButton.Click += (s, e) => SaveTemplate();
            LoadTemplateButton.Click += (s, e) => LoadTemplate();
            ExportStatsButton.Click += (s, e) => ExportStats();
        }

        private void OpenScraper()
        {
            var profile = ProfileCombo.SelectedItem?.ToString() ?? "Default";
            try
            {
                var win = new DataScraperWindow(profile);
                System.Windows.Forms.Integration.ElementHost.EnableModelessKeyboardInterop(win);
                win.Show();
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
                    ProfileName = ProfileCombo.SelectedItem?.ToString() ?? "Default",
                    Scope = ScopeCombo.SelectedItem is SpaceMapperScope sc ? sc : SpaceMapperScope.EntireModel,
                    ZoneSource = ZoneSourceCombo.SelectedItem is ZoneSourceType zs ? zs : ZoneSourceType.DataScraperZones,
                    ZoneSetName = ZoneSetBox.Text,
                    TargetSource = TargetSourceCombo.SelectedItem is TargetSourceType ts ? ts : TargetSourceType.EntireModel,
                    TargetRules = TargetRules.ToList(),
                    Mappings = Mappings.ToList(),
                    ProcessingSettings = settings
                };

                var service = new SpaceMapperService(MicroEngActions.Log);
                var result = service.Run(request);
                ShowResults(result);
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
                ProcessingMode = ProcessingModeCombo.SelectedItem is SpaceMapperProcessingMode pm ? pm : SpaceMapperProcessingMode.Auto,
                TreatPartialAsContained = TreatPartialCheck.IsChecked == true,
                TagPartialSeparately = TagPartialCheck.IsChecked == true,
                EnableMultipleZones = EnableMultiZoneCheck.IsChecked == true,
                Offset3D = ParseDouble(Offset3DBox.Text),
                OffsetTop = ParseDouble(OffsetTopBox.Text),
                OffsetBottom = ParseDouble(OffsetBottomBox.Text),
                OffsetSides = ParseDouble(OffsetSidesBox.Text),
                Units = UnitsBox.Text,
                OffsetMode = OffsetModeBox.Text,
                MaxThreads = ParseInt(MaxThreadsBox.Text),
                BatchSize = ParseInt(BatchSizeBox.Text)
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
            ResultsSummaryBox.Text = sb.ToString();
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
            ProcessingModeCombo.SelectedItem = settings.ProcessingMode;
            TreatPartialCheck.IsChecked = settings.TreatPartialAsContained;
            TagPartialCheck.IsChecked = settings.TagPartialSeparately;
            EnableMultiZoneCheck.IsChecked = settings.EnableMultipleZones;
            Offset3DBox.Text = settings.Offset3D.ToString();
            OffsetTopBox.Text = settings.OffsetTop.ToString();
            OffsetBottomBox.Text = settings.OffsetBottom.ToString();
            OffsetSidesBox.Text = settings.OffsetSides.ToString();
            UnitsBox.Text = settings.Units;
            OffsetModeBox.Text = settings.OffsetMode;
            MaxThreadsBox.Text = settings.MaxThreads?.ToString() ?? "0";
            BatchSizeBox.Text = settings.BatchSize?.ToString() ?? "0";
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
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                WindowStyle = WindowStyle.ToolWindow,
                ResizeMode = ResizeMode.NoResize
            };

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
    }
}
