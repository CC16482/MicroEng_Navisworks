using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Autodesk.Navisworks.Api;
using Wpf.Ui.Abstractions;
using WpfNavigationView = Wpf.Ui.Controls.NavigationView;
using WpfNavigatingCancelEventArgs = Wpf.Ui.Controls.NavigatingCancelEventArgs;
using WpfFlyout = Wpf.Ui.Controls.Flyout;
using NavisApp = Autodesk.Navisworks.Api.Application;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

namespace MicroEng.Navisworks
{
    public partial class SpaceMapperControl : UserControl
    {
        static SpaceMapperControl()
        {
            AssemblyResolver.EnsureRegistered();
        }

        public ObservableCollection<SpaceMapperTargetRule> TargetRules { get; } = new();
        public ObservableCollection<SpaceMapperMappingDefinition> Mappings { get; } = new();
        public ObservableCollection<ZoneSummary> ZoneSummaries { get; } = new();

        private readonly SpaceMapperTemplateStore _templateStore = new(
            Path.GetDirectoryName(typeof(SpaceMapperControl).Assembly.Location) ?? AppDomain.CurrentDomain.BaseDirectory);
        private List<SpaceMapperTemplate> _templates = new();
        private readonly SpaceMapperStepSetupPage _setupPage;
        private readonly SpaceMapperStepZonesTargetsPage _zonesPage;
        private readonly SpaceMapperStepProcessingPage _processingPage;
        private readonly SpaceMapperStepMappingPage _mappingPage;
        private readonly SpaceMapperStepResultsPage _resultsPage;
        private Stopwatch _processingNavStopwatch;
        private CancellationTokenSource _runCts;
        private bool _handlersWired;
        private bool _initialized;
        private bool _applyingTemplate;
        private bool _updatingSetLists;
        private bool? _lastMeshAccurate;
        private const string NotApplicableText = "N/A";
        private List<ModelItem> _lastTargetsWithoutBounds = new();
        private List<ModelItem> _lastTargetsUnmatched = new();

        public SpaceMapperControl()
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            DataContext = this;
            GeometryExtractor.SetUiDispatcher(Dispatcher);
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

            _processingPage.Loaded += OnProcessingPageLoaded;
            SpaceMapperNav.Navigating += OnSpaceMapperNavNavigating;
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

        private void OnSpaceMapperNavNavigating(WpfNavigationView sender, WpfNavigatingCancelEventArgs args)
        {
            RefreshSelectionSetLists();
            if (!ReferenceEquals(args.Page, _processingPage))
            {
                return;
            }

            _processingNavStopwatch = Stopwatch.StartNew();
            MicroEngActions.Log("SpaceMapper: navigating to Processing page...");
        }

        private void OnProcessingPageLoaded(object sender, RoutedEventArgs e)
        {
            if (_processingNavStopwatch == null || !_processingNavStopwatch.IsRunning)
            {
                return;
            }

            _processingNavStopwatch.Stop();
            MicroEngActions.Log($"SpaceMapper: Processing page loaded in {_processingNavStopwatch.ElapsedMilliseconds}ms");
            _processingNavStopwatch = null;
        }

        private void InitializeUi()
        {
            _applyingTemplate = true;
            try
            {
                _zonesPage.ZoneSourceCombo.ItemsSource = Enum.GetValues(typeof(ZoneSourceType));
                _zonesPage.RuleTargetDefinitionCombo.SelectedValue = SpaceMapperTargetDefinition.EntireModel;
                _zonesPage.RuleMembershipCombo.ItemsSource = Enum.GetValues(typeof(SpaceMembershipMode));
                _zonesPage.RuleMembershipCombo.SelectedItem = SpaceMembershipMode.ContainedAndPartial;
                _zonesPage.RuleEnabledCheckBox.IsChecked = true;
                UpdateRuleDefinitionInputs(clearText: true, focusSetName: false);

                RefreshTemplateList();
                RefreshScraperProfiles();

                if (!_handlersWired)
                {
                    AddHandlers();
                    _handlersWired = true;
                }

                LoadSelectedTemplate();
                RefreshSelectionSetLists();
            }
            finally
            {
                _applyingTemplate = false;
            }

            UpdateReadinessText();
        }

        private void AddHandlers()
        {
            RunSpaceMapperButtonControl.Click += (s, e) => RunSpaceMapper();
            _setupPage.NewTemplateButton.Click += (s, e) => NewTemplate();
            _setupPage.SaveTemplateButton.Click += (s, e) => SaveCurrentTemplate();
            _setupPage.SaveAsTemplateButton.Click += (s, e) => SaveTemplateAs();
            _setupPage.DeleteTemplateButton.Click += (s, e) => DeleteTemplate();
            _setupPage.TemplateCombo.SelectionChanged += (s, e) => OnTemplateSelectionChanged();
            _setupPage.RefreshScraperButton.Click += (s, e) => RefreshScraperProfiles();
            _setupPage.OpenScraperButton.Click += (s, e) => OpenScraper();
            _setupPage.ScraperProfileCombo.SelectionChanged += (s, e) => OnScraperProfileChanged();
            _setupPage.ValidateButton.Click += (s, e) => ValidateReadiness();
            _setupPage.GoToZonesButton.Click += (s, e) => SpaceMapperNav.Navigate(typeof(SpaceMapperStepZonesTargetsPage));
            _zonesPage.AddRuleButton.Click += (s, e) => AddTargetRule();
            _zonesPage.DeleteRuleButton.Click += (s, e) => DeleteSelectedRule();
            _mappingPage.AddMappingButton.Click += (s, e) => AddMapping();
            _mappingPage.DeleteMappingButton.Click += (s, e) => DeleteSelectedMapping();
            _mappingPage.SaveTemplateButton.Click += (s, e) => SaveTemplate();
            _mappingPage.LoadTemplateButton.Click += (s, e) => LoadTemplate();
            _mappingPage.ZonePropertyPickerButton.Click += (s, e) => ToggleZonePropertyPicker();
            _mappingPage.ZonePropertyTreeView.SelectedItemChanged += OnZonePropertyTreeSelectionChanged;
            _resultsPage.ExportStatsButton.Click += (s, e) => ExportStats();
            if (RunHealthDetailsButtonControl != null)
            {
                RunHealthDetailsButtonControl.Click += (s, e) => ToggleRunHealthFlyout(RunHealthDetailsFlyoutControl);
            }
            if (_resultsPage.RunHealthDetailsButton != null)
            {
                _resultsPage.RunHealthDetailsButton.Click += (s, e) => ToggleRunHealthFlyout(_resultsPage.RunHealthDetailsFlyout);
            }
            if (RunHealthCreateMissingBoundsButtonControl != null)
            {
                RunHealthCreateMissingBoundsButtonControl.Click += (s, e) => CreateDiagnosticsSelectionSet("Targets without bounds", _lastTargetsWithoutBounds);
            }
            if (_resultsPage.RunHealthCreateMissingBoundsButton != null)
            {
                _resultsPage.RunHealthCreateMissingBoundsButton.Click += (s, e) => CreateDiagnosticsSelectionSet("Targets without bounds", _lastTargetsWithoutBounds);
            }
            if (RunHealthCreateUnmatchedButtonControl != null)
            {
                RunHealthCreateUnmatchedButtonControl.Click += (s, e) => CreateDiagnosticsSelectionSet("Unmatched targets", _lastTargetsUnmatched);
            }
            if (_resultsPage.RunHealthCreateUnmatchedButton != null)
            {
                _resultsPage.RunHealthCreateUnmatchedButton.Click += (s, e) => CreateDiagnosticsSelectionSet("Unmatched targets", _lastTargetsUnmatched);
            }

            _processingPage.ZoneBoundsSlider.ValueChanged += (s, e) => OnBoundsModesChanged();
            _processingPage.TargetBoundsSlider.ValueChanged += (s, e) => OnBoundsModesChanged();
            _processingPage.ZoneKDopVariantCombo.SelectionChanged += (s, e) => OnBoundsModesChanged();
            _processingPage.TargetKDopVariantCombo.SelectionChanged += (s, e) => OnBoundsModesChanged();
            _processingPage.TargetMidpointModeCombo.SelectionChanged += (s, e) => OnBoundsModesChanged();
            if (_processingPage.ZoneContainmentEngineCombo != null)
            {
                _processingPage.ZoneContainmentEngineCombo.SelectionChanged += (s, e) => OnBoundsModesChanged();
            }
            if (_processingPage.ZoneResolutionStrategyCombo != null)
            {
                _processingPage.ZoneResolutionStrategyCombo.SelectionChanged += (s, e) => OnBoundsModesChanged();
            }
            if (_processingPage.VariationCheckButton != null)
            {
                _processingPage.VariationCheckButton.Click += (s, e) => RunVariationCheckReport();
            }

            _processingPage.TreatPartialCheck.Checked += (s, e) =>
            {
                UpdateFastTraversalUi();
                TriggerLivePreflight();
            };
            _processingPage.TreatPartialCheck.Unchecked += (s, e) =>
            {
                UpdateFastTraversalUi();
                TriggerLivePreflight();
            };
            _processingPage.TagPartialCheck.Checked += (s, e) =>
            {
                UpdateFastTraversalUi();
                TriggerLivePreflight();
            };
            _processingPage.TagPartialCheck.Unchecked += (s, e) =>
            {
                UpdateFastTraversalUi();
                TriggerLivePreflight();
            };
            if (_processingPage.WriteZoneBehaviorCheck != null)
            {
                _processingPage.WriteZoneBehaviorCheck.Checked += (s, e) =>
                {
                    UpdateZoneBehaviorInputs();
                    UpdateFastTraversalUi();
                    TriggerLivePreflight();
                };
                _processingPage.WriteZoneBehaviorCheck.Unchecked += (s, e) =>
                {
                    UpdateZoneBehaviorInputs();
                    UpdateFastTraversalUi();
                    TriggerLivePreflight();
                };
            }
            if (_processingPage.WriteContainmentPercentCheck != null)
            {
                _processingPage.WriteContainmentPercentCheck.Checked += (s, e) =>
                {
                    UpdateFastTraversalUi();
                    TriggerLivePreflight();
                };
                _processingPage.WriteContainmentPercentCheck.Unchecked += (s, e) =>
                {
                    UpdateFastTraversalUi();
                    TriggerLivePreflight();
                };
            }
            if (_processingPage.ContainmentCalculationCombo != null)
            {
                _processingPage.ContainmentCalculationCombo.SelectionChanged += (s, e) => TriggerLivePreflight();
            }
            _processingPage.EnableMultiZoneCheck.Checked += (s, e) => TriggerLivePreflight();
            _processingPage.EnableMultiZoneCheck.Unchecked += (s, e) => TriggerLivePreflight();
            if (_processingPage.ExcludeZonesFromTargetsCheck != null)
            {
                _processingPage.ExcludeZonesFromTargetsCheck.Checked += (s, e) => TriggerLivePreflight();
                _processingPage.ExcludeZonesFromTargetsCheck.Unchecked += (s, e) => TriggerLivePreflight();
            }
            _processingPage.Offset3DBox.TextChanged += (s, e) => TriggerLivePreflight();
            _processingPage.OffsetTopBox.TextChanged += (s, e) => TriggerLivePreflight();
            _processingPage.OffsetBottomBox.TextChanged += (s, e) => TriggerLivePreflight();
            _processingPage.OffsetSidesBox.TextChanged += (s, e) => TriggerLivePreflight();
            _processingPage.UnitsBox.TextChanged += (s, e) => TriggerLivePreflight();
            _processingPage.OffsetModeBox.TextChanged += (s, e) => TriggerLivePreflight();
            if (_processingPage.EnableZoneOffsetsCheck != null)
            {
                _processingPage.EnableZoneOffsetsCheck.Checked += (s, e) => TriggerLivePreflight();
                _processingPage.EnableZoneOffsetsCheck.Unchecked += (s, e) => TriggerLivePreflight();
            }
            if (_processingPage.EnableOffsetAreaPassCheck != null)
            {
                _processingPage.EnableOffsetAreaPassCheck.Checked += (s, e) => TriggerLivePreflight();
                _processingPage.EnableOffsetAreaPassCheck.Unchecked += (s, e) => TriggerLivePreflight();
            }
            if (_processingPage.WriteOffsetMatchPropertyCheck != null)
            {
                _processingPage.WriteOffsetMatchPropertyCheck.Checked += (s, e) => TriggerLivePreflight();
                _processingPage.WriteOffsetMatchPropertyCheck.Unchecked += (s, e) => TriggerLivePreflight();
            }

            _zonesPage.ZoneSourceCombo.SelectionChanged += (s, e) =>
            {
                UpdateReadinessText();
                TriggerLivePreflight();
            };
            _zonesPage.ZoneSetBox.TextChanged += (s, e) =>
            {
                UpdateReadinessText();
                TriggerLivePreflight();
            };
            _zonesPage.ZoneSetMenu.Opened += (s, e) => RefreshSelectionSetLists();
            _zonesPage.ZoneSetMenu.AddHandler(MenuItem.ClickEvent, new RoutedEventHandler(OnZoneSetMenuItemClick));
            _zonesPage.RuleTargetDefinitionCombo.SelectionChanged += (s, e) =>
            {
                UpdateRuleDefinitionInputs(clearText: true, focusSetName: true);
            };
            _zonesPage.RuleSetMenu.Opened += (s, e) => RefreshSelectionSetLists();
            _zonesPage.RuleSetMenu.AddHandler(MenuItem.ClickEvent, new RoutedEventHandler(OnRuleSetMenuItemClick));
            _zonesPage.TargetRulesGrid.CellEditEnding += (s, e) =>
            {
                if (e.Row?.Item is SpaceMapperTargetRule rule)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        NormalizeTargetRule(rule);
                        _zonesPage.TargetRulesGrid.Items.Refresh();
                    }), DispatcherPriority.Background);
                }
                UpdateReadinessText();
                TriggerLivePreflight();
            };
            _mappingPage.MappingGrid.CellEditEnding += (s, e) =>
            {
                UpdateReadinessText();
                TriggerLivePreflight();
            };
            TargetRules.CollectionChanged += (s, e) =>
            {
                UpdateReadinessText();
                TriggerLivePreflight();
            };
            Mappings.CollectionChanged += (s, e) =>
            {
                UpdateReadinessText();
                TriggerLivePreflight();
            };
        }

        private void OnTemplateSelectionChanged()
        {
            if (_applyingTemplate)
            {
                return;
            }

            LoadSelectedTemplate();
        }

        private void RefreshTemplateList(string selectName = null)
        {
            _templates = _templateStore.Load() ?? new List<SpaceMapperTemplate>();
            if (_templates.Count == 0)
            {
                var fallback = SpaceMapperTemplateStore.CreateDefault();
                _templates.Add(fallback);
                _templateStore.Save(_templates);
            }

            _setupPage.TemplateCombo.Items.Clear();
            foreach (var template in _templates.OrderBy(t => t.Name ?? string.Empty))
            {
                _setupPage.TemplateCombo.Items.Add(template.Name);
            }

            var nameToSelect = selectName;
            if (string.IsNullOrWhiteSpace(nameToSelect))
            {
                nameToSelect = _templates.FirstOrDefault()?.Name;
            }

            if (!string.IsNullOrWhiteSpace(nameToSelect) && _setupPage.TemplateCombo.Items.Contains(nameToSelect))
            {
                _setupPage.TemplateCombo.SelectedItem = nameToSelect;
            }
            else if (_setupPage.TemplateCombo.Items.Count > 0)
            {
                _setupPage.TemplateCombo.SelectedIndex = 0;
            }
        }

        private SpaceMapperTemplate GetSelectedTemplate()
        {
            var name = _setupPage.TemplateCombo.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            return _templates.FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        private void LoadSelectedTemplate()
        {
            var template = GetSelectedTemplate();
            if (template == null)
            {
                return;
            }

            ApplyTemplate(template);
        }

        private void ApplyTemplate(SpaceMapperTemplate template)
        {
            if (template == null)
            {
                return;
            }

            _applyingTemplate = true;
            try
            {
                TargetRules.Clear();
                if (template.TargetRules != null)
                {
                    foreach (var rule in template.TargetRules)
                    {
                        TargetRules.Add(rule);
                    }
                }

                if (TargetRules.Count == 0)
                {
                    TargetRules.Add(new SpaceMapperTargetRule
                    {
                        Name = "All Targets",
                        TargetDefinition = SpaceMapperTargetDefinition.EntireModel,
                        MembershipMode = SpaceMembershipMode.ContainedAndPartial,
                        Enabled = true
                    });
                }
                else
                {
                    foreach (var rule in TargetRules)
                    {
                        NormalizeTargetRule(rule);
                    }
                }

                Mappings.Clear();
                if (template.Mappings != null)
                {
                    foreach (var map in template.Mappings)
                    {
                        Mappings.Add(map);
                    }
                }

                if (Mappings.Count == 0)
                {
                    Mappings.Add(new SpaceMapperMappingDefinition
                    {
                        Name = "Zone Name",
                        ZoneCategory = "Zone",
                        ZonePropertyName = "Name",
                        TargetPropertyName = "Zone Name"
                    });
                }

                _zonesPage.ZoneSourceCombo.SelectedItem = template.ZoneSource;
                _zonesPage.ZoneSetBox.Text = template.ZoneSetName ?? string.Empty;

                ApplySettings(template.ProcessingSettings ?? new SpaceMapperProcessingSettings());
                SelectScraperProfile(template.PreferredScraperProfileName ?? "Default");
            }
            finally
            {
                _applyingTemplate = false;
            }

            UpdateReadinessText();
        }

        private SpaceMapperTemplate BuildTemplateFromUi(string name)
        {
            return new SpaceMapperTemplate
            {
                Name = name,
                TargetRules = TargetRules.ToList(),
                Mappings = Mappings.ToList(),
                ProcessingSettings = BuildSettings(),
                PreferredScraperProfileName = GetSelectedScraperProfileName() ?? "Default",
                ZoneSource = _zonesPage.ZoneSourceCombo.SelectedItem is ZoneSourceType zs ? zs : ZoneSourceType.DataScraperZones,
                ZoneSetName = _zonesPage.ZoneSetBox.Text
            };
        }

        private void NewTemplate()
        {
            var name = PromptText("New Space Mapper profile name:", "Space Mapper");
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            var templates = _templateStore.Load();
            if (templates.Any(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show($"A profile named '{name}' already exists.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var template = SpaceMapperTemplateStore.CreateDefault();
            template.Name = name;
            templates.Add(template);
            _templateStore.Save(templates);
            RefreshTemplateList(name);
            ApplyTemplate(template);
        }

        private void SaveCurrentTemplate()
        {
            var name = _setupPage.TemplateCombo.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show("Select a Space Mapper profile to save.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var templates = _templateStore.Load();
            var existing = templates.FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));
            var updated = BuildTemplateFromUi(name);

            if (existing == null)
            {
                templates.Add(updated);
            }
            else
            {
                var index = templates.IndexOf(existing);
                templates[index] = updated;
            }

            _templateStore.Save(templates);
            RefreshTemplateList(name);
        }

        private void SaveTemplateAs()
        {
            var name = PromptText("Save Space Mapper profile as:", "Space Mapper");
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            var templates = _templateStore.Load();
            var existing = templates.FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                var overwrite = MessageBox.Show($"Overwrite existing profile '{name}'?", "MicroEng", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (overwrite != MessageBoxResult.Yes)
                {
                    return;
                }
            }

            var updated = BuildTemplateFromUi(name);
            if (existing == null)
            {
                templates.Add(updated);
            }
            else
            {
                var index = templates.IndexOf(existing);
                templates[index] = updated;
            }

            _templateStore.Save(templates);
            RefreshTemplateList(name);
        }

        private void DeleteTemplate()
        {
            var name = _setupPage.TemplateCombo.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            var templates = _templateStore.Load();
            if (templates.Count <= 1)
            {
                MessageBox.Show("At least one Space Mapper profile is required.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var confirm = MessageBox.Show($"Delete profile '{name}'?", "MicroEng", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            templates = templates
                .Where(t => !string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (templates.Count == 0)
            {
                templates.Add(SpaceMapperTemplateStore.CreateDefault());
            }

            _templateStore.Save(templates);
            var nextName = templates.FirstOrDefault()?.Name;
            RefreshTemplateList(nextName);
        }

        private void RefreshScraperProfiles(string selectName = null)
        {
            var profiles = DataScraperCache.AllSessions
                .Select(s => string.IsNullOrWhiteSpace(s.ProfileName) ? "Default" : s.ProfileName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(p => p)
                .ToList();

            _setupPage.ScraperProfileCombo.Items.Clear();
            foreach (var profile in profiles)
            {
                _setupPage.ScraperProfileCombo.Items.Add(profile);
            }

            if (_setupPage.ScraperProfileCombo.Items.Count == 0)
            {
                _setupPage.ScraperProfileCombo.Items.Add("Default");
            }

            var nameToSelect = selectName;
            if (string.IsNullOrWhiteSpace(nameToSelect))
            {
                nameToSelect = _setupPage.ScraperProfileCombo.SelectedItem?.ToString();
            }

            if (string.IsNullOrWhiteSpace(nameToSelect))
            {
                _setupPage.ScraperProfileCombo.SelectedIndex = 0;
            }
            else if (_setupPage.ScraperProfileCombo.Items.Contains(nameToSelect))
            {
                _setupPage.ScraperProfileCombo.SelectedItem = nameToSelect;
            }
            else
            {
                _setupPage.ScraperProfileCombo.SelectedIndex = 0;
            }

            var session = GetLatestScrapeSession(GetSelectedScraperProfileName());
            DataScraperCache.LastSession = session;
            UpdateScraperSummary(session);

            if (!_applyingTemplate)
            {
                UpdateReadinessText();
            }
        }

        private void SelectScraperProfile(string profileName)
        {
            if (string.IsNullOrWhiteSpace(profileName))
            {
                return;
            }

            if (_setupPage.ScraperProfileCombo.Items.Contains(profileName))
            {
                _setupPage.ScraperProfileCombo.SelectedItem = profileName;
            }
            else
            {
                RefreshScraperProfiles(profileName);
            }
        }

        private void OnScraperProfileChanged()
        {
            var profile = GetSelectedScraperProfileName();
            var session = GetLatestScrapeSession(profile);
            DataScraperCache.LastSession = session;

            UpdateScraperSummary(session);

            if (_applyingTemplate)
            {
                return;
            }

            UpdateReadinessText();
            TriggerLivePreflight();
        }

        private string GetSelectedScraperProfileName()
        {
            var name = _setupPage.ScraperProfileCombo.SelectedItem?.ToString();
            return string.IsNullOrWhiteSpace(name) ? null : name.Trim();
        }

        private ScrapeSession GetLatestScrapeSession(string profileName)
        {
            if (string.IsNullOrWhiteSpace(profileName))
            {
                return null;
            }

            return DataScraperCache.AllSessions
                .Where(s => string.Equals(s.ProfileName ?? "Default", profileName, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(s => s.Timestamp)
                .FirstOrDefault();
        }

        private void UpdateScraperSummary(ScrapeSession session)
        {
            if (session == null)
            {
                _setupPage.ScraperSummaryText.Text = DataScraperCache.AllSessions.Count == 0
                    ? "No Data Scraper sessions found."
                    : "No Data Scraper session selected.";
                return;
            }

            var scope = string.Join(" ", new[] { session.ScopeType, session.ScopeDescription }
                .Where(s => !string.IsNullOrWhiteSpace(s)));
            if (string.IsNullOrWhiteSpace(scope))
            {
                scope = "Unknown scope";
            }

            var propertyCount = session.Properties?.Count ?? 0;
            _setupPage.ScraperSummaryText.Text = $"Last scrape: {scope} | {session.ItemsScanned:N0} items | {propertyCount:N0} properties | {session.Timestamp:g}";
        }

        private void UpdateReadinessText()
        {
            if (_setupPage?.ReadinessText == null)
            {
                return;
            }

            var zonesReady = AreZonesConfigured();
            var targetsReady = AreTargetsConfigured();
            var mappingsReady = AreMappingsConfigured();

            var lines = new List<string>
            {
                $"Zones configured: {(zonesReady ? "Yes" : "No")}",
                $"Targets configured: {(targetsReady ? "Yes" : "No")}",
                $"Mappings configured: {(mappingsReady ? "Yes" : "No")}"
            };

            _setupPage.ReadinessText.Text = string.Join(Environment.NewLine, lines);
        }

        private void ValidateReadiness()
        {
            UpdateReadinessText();
            var missing = new List<string>();
            if (!AreZonesConfigured()) missing.Add("Zones");
            if (!AreTargetsConfigured()) missing.Add("Targets");
            if (!AreMappingsConfigured()) missing.Add("Mappings");

            if (missing.Count == 0)
            {
                MessageBox.Show("All required sections are configured.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var message = $"Missing configuration: {string.Join(", ", missing)}.";
            MessageBox.Show(message, "MicroEng", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private bool AreZonesConfigured()
        {
            if (_zonesPage.ZoneSourceCombo.SelectedItem == null)
            {
                return false;
            }

            var zoneSource = _zonesPage.ZoneSourceCombo.SelectedItem is ZoneSourceType zs
                ? zs
                : ZoneSourceType.DataScraperZones;

            if (zoneSource == ZoneSourceType.DataScraperZones)
            {
                var profile = GetSelectedScraperProfileName();
                return !string.IsNullOrWhiteSpace(profile) && SpaceMapperService.GetSession(profile) != null;
            }

            if ((zoneSource == ZoneSourceType.ZoneSelectionSet || zoneSource == ZoneSourceType.ZoneSearchSet)
                && string.IsNullOrWhiteSpace(_zonesPage.ZoneSetBox.Text))
            {
                return false;
            }

            return true;
        }

        private bool AreTargetsConfigured()
        {
            if (TargetRules.Count == 0)
            {
                return false;
            }

            var enabledRules = TargetRules.Where(r => r.Enabled).ToList();
            if (enabledRules.Count == 0)
            {
                return false;
            }

            return enabledRules.All(IsRuleConfigured);
        }

        private bool AreMappingsConfigured()
        {
            return Mappings.Count > 0;
        }

        private void AddTargetRule()
        {
            if (!EnsureZoneSourceReady())
            {
                return;
            }

            var targetDefinition = _zonesPage.RuleTargetDefinitionCombo.SelectedValue is SpaceMapperTargetDefinition rt
                ? rt
                : SpaceMapperTargetDefinition.EntireModel;

            var ruleSetName = _zonesPage.RuleSetNameBox.Text?.Trim();
            if (UsesSetSearchName(targetDefinition) && (string.IsNullOrWhiteSpace(ruleSetName) || IsNotApplicable(ruleSetName)))
            {
                MessageBox.Show("Enter a Set/Search Name for the rule when using Selection/Search Set.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Warning);
                _zonesPage.RuleSetNameBox.Focus();
                return;
            }
            if (!UsesSetSearchName(targetDefinition))
            {
                ruleSetName = null;
            }

            var name = _zonesPage.RuleNameBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                name = $"Rule {TargetRules.Count + 1}";
            }

            int? minLevel = null;
            int? maxLevel = null;
            if (UsesLevels(targetDefinition))
            {
                if (!TryParseOptionalNonNegativeInt(_zonesPage.RuleMinLevelBox.Text, out minLevel))
                {
                    ShowValidationMessage("Min Level must be a whole number (0 or greater).",
                        typeof(SpaceMapperStepZonesTargetsPage),
                        () => _zonesPage.RuleMinLevelBox.Focus());
                    return;
                }

                if (!TryParseOptionalNonNegativeInt(_zonesPage.RuleMaxLevelBox.Text, out maxLevel))
                {
                    ShowValidationMessage("Max Level must be a whole number (0 or greater).",
                        typeof(SpaceMapperStepZonesTargetsPage),
                        () => _zonesPage.RuleMaxLevelBox.Focus());
                    return;
                }

                if (minLevel == null && maxLevel == null)
                {
                    minLevel = 0;
                    maxLevel = 0;
                }
                else if (minLevel == null)
                {
                    minLevel = maxLevel;
                }
                else if (maxLevel == null)
                {
                    maxLevel = minLevel;
                }

                if (minLevel < 0 || maxLevel < 0)
                {
                    ShowValidationMessage("Min/Max Level must be 0 or greater.",
                        typeof(SpaceMapperStepZonesTargetsPage),
                        () => _zonesPage.RuleMinLevelBox.Focus());
                    return;
                }

                if (minLevel > maxLevel)
                {
                    ShowValidationMessage("Min Level cannot be greater than Max Level.",
                        typeof(SpaceMapperStepZonesTargetsPage),
                        () => _zonesPage.RuleMinLevelBox.Focus());
                    return;
                }
            }
            var membership = _zonesPage.RuleMembershipCombo.SelectedItem is SpaceMembershipMode mm
                ? mm
                : SpaceMembershipMode.ContainedAndPartial;

            var rule = new SpaceMapperTargetRule
            {
                Name = name,
                TargetDefinition = targetDefinition,
                MinLevel = minLevel,
                MaxLevel = maxLevel,
                SetSearchName = ruleSetName,
                CategoryFilter = _zonesPage.RuleCategoryFilterBox.Text?.Trim(),
                MembershipMode = membership,
                Enabled = _zonesPage.RuleEnabledCheckBox.IsChecked == true
            };
            NormalizeTargetRule(rule);
            TargetRules.Add(rule);
            _zonesPage.TargetRulesGrid.SelectedItem = rule;
            _zonesPage.TargetRulesGrid.ScrollIntoView(rule);
            _zonesPage.TargetRulesGrid.Items.Refresh();
        }

        private void DeleteSelectedRule()
        {
            if (TargetRules.Count == 0)
            {
                MessageBox.Show("There are no rules to delete.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (_zonesPage.TargetRulesGrid.SelectedItem is SpaceMapperTargetRule rule)
            {
                TargetRules.Remove(rule);
                _zonesPage.TargetRulesGrid.Items.Refresh();
                return;
            }

            MessageBox.Show("Select a rule to delete.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private bool EnsureZoneSourceReady(bool showMessage = true)
        {
            var zoneSource = _zonesPage.ZoneSourceCombo.SelectedItem is ZoneSourceType zs
                ? zs
                : ZoneSourceType.DataScraperZones;

            if ((zoneSource == ZoneSourceType.ZoneSelectionSet || zoneSource == ZoneSourceType.ZoneSearchSet)
                && string.IsNullOrWhiteSpace(_zonesPage.ZoneSetBox.Text))
            {
                if (showMessage)
                {
                    MessageBox.Show("Enter a Set/Search Name for Zones when using Selection/Search Set.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _zonesPage.ZoneSetBox.Focus();
                }
                return false;
            }

            return true;
        }

        private static bool UsesSetSearchName(SpaceMapperTargetDefinition definition)
        {
            return definition == SpaceMapperTargetDefinition.SelectionSet
                || definition == SpaceMapperTargetDefinition.SearchSet;
        }

        private static bool UsesLevels(SpaceMapperTargetDefinition definition)
        {
            return definition == SpaceMapperTargetDefinition.SelectionTreeLevel;
        }

        private static bool IsNotApplicable(string value)
        {
            return string.Equals(value?.Trim(), NotApplicableText, StringComparison.OrdinalIgnoreCase);
        }

        private void UpdateZoneBehaviorInputs()
        {
            if (_processingPage == null)
            {
                return;
            }

            var enabled = _processingPage.WriteZoneBehaviorCheck?.IsChecked == true;
            _processingPage.ZoneBehaviorCategoryBox.IsEnabled = enabled;
            _processingPage.ZoneBehaviorPropertyBox.IsEnabled = enabled;
            _processingPage.ZoneBehaviorContainedBox.IsEnabled = enabled;
            _processingPage.ZoneBehaviorPartialBox.IsEnabled = enabled;
            _processingPage.UpdateProcessingUiState();
        }

        private void RefreshSelectionSetLists()
        {
            if (_zonesPage == null)
            {
                return;
            }

            var lists = GetSelectionSetLists();

            _updatingSetLists = true;
            try
            {
                var zoneSelectedName = _zonesPage.ZoneSetBox.Text;
                bool? zoneIsSearch = _zonesPage.ZoneSourceCombo.SelectedItem is ZoneSourceType zoneSource
                    ? zoneSource switch
                    {
                        ZoneSourceType.ZoneSelectionSet => false,
                        ZoneSourceType.ZoneSearchSet => true,
                        _ => (bool?)null
                    }
                    : null;

                var ruleSelectedName = _zonesPage.RuleSetNameBox.Text;
                bool? ruleIsSearch = _zonesPage.RuleTargetDefinitionCombo.SelectedValue is SpaceMapperTargetDefinition ruleDef
                    ? ruleDef switch
                    {
                        SpaceMapperTargetDefinition.SearchSet => true,
                        SpaceMapperTargetDefinition.SelectionSet => (bool?)null,
                        _ => (bool?)null
                    }
                    : null;

                UpdateMenuItems(_zonesPage.ZoneSetMenu, lists.Sets, zoneSelectedName, zoneIsSearch);
                UpdateMenuItems(_zonesPage.RuleSetMenu, lists.Sets, ruleSelectedName, ruleIsSearch);
            }
            finally
            {
                _updatingSetLists = false;
            }
        }

        private void ToggleZonePropertyPicker()
        {
            if (_mappingPage?.ZonePropertyPickerFlyout == null)
            {
                return;
            }

            if (_mappingPage.ZonePropertyPickerFlyout.IsOpen)
            {
                _mappingPage.ZonePropertyPickerFlyout.Hide();
                return;
            }

            RefreshZonePropertyTree();
            _mappingPage.ZonePropertyPickerFlyout.Show();
        }

        private void RefreshZonePropertyTree()
        {
            if (_mappingPage?.ZonePropertyTreeView == null)
            {
                return;
            }

            var profile = GetSelectedScraperProfileName();
            var session = GetLatestScrapeSession(profile);
            var items = BuildZonePropertyTree(session);

            _mappingPage.ZonePropertyTreeView.ItemsSource = items;

            if (_mappingPage.ZonePropertyPickerStatusText != null)
            {
                if (session?.Properties == null || session.Properties.Count == 0)
                {
                    if (DataScraperCache.AllSessions.Count == 0)
                    {
                        _mappingPage.ZonePropertyPickerStatusText.Text = "No Data Scraper sessions found.";
                    }
                    else
                    {
                        var label = string.IsNullOrWhiteSpace(profile) ? "Default" : profile;
                        _mappingPage.ZonePropertyPickerStatusText.Text = $"No Data Scraper session for '{label}'.";
                    }
                }
                else
                {
                    var label = string.IsNullOrWhiteSpace(profile) ? "Default" : profile;
                    _mappingPage.ZonePropertyPickerStatusText.Text = $"{label} · {session.Properties.Count:N0} properties";
                }
            }
        }

        private void OnZonePropertyTreeSelectionChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (_mappingPage == null)
            {
                return;
            }

            if (e.NewValue is not ZonePropertyTreeNode node || !node.IsProperty)
            {
                return;
            }

            _mappingPage.ZoneCategoryBox.Text = node.CategoryName;
            _mappingPage.ZonePropertyBox.Text = node.PropertyName;
            _mappingPage.ZonePropertyPickerFlyout?.Hide();
        }

        private static List<ZonePropertyTreeNode> BuildZonePropertyTree(ScrapeSession session)
        {
            var nodes = new List<ZonePropertyTreeNode>();
            if (session?.Properties == null || session.Properties.Count == 0)
            {
                return nodes;
            }

            var groups = session.Properties
                .Where(p => !string.IsNullOrWhiteSpace(p.Category) && !string.IsNullOrWhiteSpace(p.Name))
                .GroupBy(p => p.Category, StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase);

            foreach (var group in groups)
            {
                var categoryNode = new ZonePropertyTreeNode(group.Key, group.Key, null, isExpanded: false);
                var props = group.Select(p => p.Name)
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase);

                foreach (var propName in props)
                {
                    categoryNode.Children.Add(new ZonePropertyTreeNode(propName, group.Key, propName));
                }

                nodes.Add(categoryNode);
            }

            return nodes;
        }

        private void OnZoneSetMenuItemClick(object sender, RoutedEventArgs e)
        {
            if (_updatingSetLists)
            {
                return;
            }

            if (e.OriginalSource is not MenuItem menuItem)
            {
                return;
            }

            if (menuItem.DataContext is not SetListItem item)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(item.Name))
            {
                return;
            }

            _zonesPage.ZoneSetBox.Text = item.Name;
            _zonesPage.ZoneSourceCombo.SelectedItem = item.IsSearch
                ? ZoneSourceType.ZoneSearchSet
                : ZoneSourceType.ZoneSelectionSet;
        }

        private void OnRuleSetMenuItemClick(object sender, RoutedEventArgs e)
        {
            if (_updatingSetLists)
            {
                return;
            }

            if (e.OriginalSource is not MenuItem menuItem)
            {
                return;
            }

            if (menuItem.DataContext is not SetListItem item)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(item.Name))
            {
                return;
            }

            _zonesPage.RuleSetNameBox.Text = item.Name;
            _zonesPage.RuleTargetDefinitionCombo.SelectedValue = SpaceMapperTargetDefinition.SelectionSet;
            UpdateRuleDefinitionInputs(clearText: false, focusSetName: false);
        }

        private static void UpdateMenuItems(ContextMenu menu, List<SetListItem> items, string selectedName, bool? selectedIsSearch)
        {
            if (menu == null)
            {
                return;
            }

            var name = string.IsNullOrWhiteSpace(selectedName) ? null : selectedName.Trim();
            var hasName = !string.IsNullOrWhiteSpace(name);

            var menuItems = items?
                .Select(item => item.WithSelection(
                    hasName
                    && string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)
                    && (!selectedIsSearch.HasValue || item.IsSearch == selectedIsSearch.Value)))
                .ToList()
                ?? new List<SetListItem>();

            menu.ItemsSource = menuItems;
        }

        private sealed class SetListItem
        {
            public SetListItem(string name, bool isSearch, bool isSelected = false)
            {
                Name = name;
                IsSearch = isSearch;
                IsSelected = isSelected;
                DisplayName = $"{name} ({(isSearch ? "Search" : "Selection")})";
            }

            public string Name { get; }
            public bool IsSearch { get; }
            public bool IsSelected { get; }
            public string DisplayName { get; }

            public SetListItem WithSelection(bool isSelected)
            {
                return new SetListItem(Name, IsSearch, isSelected);
            }
        }

        private sealed class ZonePropertyTreeNode
        {
            public ZonePropertyTreeNode(string name, string categoryName, string propertyName, bool isExpanded = false)
            {
                Name = name;
                CategoryName = categoryName;
                PropertyName = propertyName;
                IsExpanded = isExpanded;
            }

            public string Name { get; }
            public string CategoryName { get; }
            public string PropertyName { get; }
            public bool IsExpanded { get; set; }
            public bool IsProperty => !string.IsNullOrWhiteSpace(PropertyName);
            public List<ZonePropertyTreeNode> Children { get; } = new();
        }

        private sealed class SelectionSetLists
        {
            public List<SetListItem> Sets { get; } = new();
        }

        private static SelectionSetLists GetSelectionSetLists()
        {
            var lists = new SelectionSetLists();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var doc = NavisApp.ActiveDocument;
                var root = doc?.SelectionSets?.RootItem;
                if (root != null)
                {
                    CollectSelectionSets(root, lists.Sets, seen);
                }
            }
            catch
            {
                // ignore managed API enumeration failures
            }

            if (lists.Sets.Count == 0)
            {
                try
                {
                    var state = ComBridge.State;
                    var root = state.SelectionSetsEx();
                    CollectSelectionSets(root, lists.Sets, seen, null);
                }
                catch
                {
                    // ignore COM enumeration failures
                }
            }

            lists.Sets.Sort((a, b) => string.Compare(a.DisplayName, b.DisplayName, StringComparison.OrdinalIgnoreCase));
            return lists;
        }

        private static void CollectSelectionSets(FolderItem folder, List<SetListItem> items, HashSet<string> seen)
        {
            if (folder?.Children == null)
            {
                return;
            }

            foreach (SavedItem child in folder.Children)
            {
                if (child is SelectionSet set)
                {
                    var name = set.DisplayName;
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    AddSetItem(items, seen, name, set.HasSearch);
                }
                else if (child is FolderItem subFolder)
                {
                    CollectSelectionSets(subFolder, items, seen);
                }
            }
        }

        private static void CollectSelectionSets(ComApi.InwSelectionSetExColl collection,
            List<SetListItem> items,
            HashSet<string> seen,
            string folderHint)
        {
            if (collection == null) return;

            for (int i = 1; i <= collection.Count; i++)
            {
                object item;
                try
                {
                    item = collection[i];
                }
                catch
                {
                    continue;
                }

                if (item is ComApi.InwOpSelectionSet set)
                {
                    var name = TryGetName(set);
                    if (string.IsNullOrWhiteSpace(name)) continue;

                    var isSearch = TryGetSearchFlag(set)
                        ?? string.Equals(folderHint, "search", StringComparison.OrdinalIgnoreCase);
                    AddSetItem(items, seen, name, isSearch);
                }
                else if (item is ComApi.InwSelectionSetFolder folder)
                {
                    var folderName = TryGetName(folder);
                    var nextHint = folderHint;
                    if (!string.IsNullOrWhiteSpace(folderName))
                    {
                        var lowered = folderName.ToLowerInvariant();
                        if (lowered.Contains("search"))
                        {
                            nextHint = "search";
                        }
                        else if (lowered.Contains("selection"))
                        {
                            nextHint = "selection";
                        }
                    }

                    try
                    {
                        CollectSelectionSets(folder.SelectionSets(), items, seen, nextHint);
                    }
                    catch
                    {
                        // ignore folder enumeration failures
                    }
                }
            }
        }

        private static void AddSetItem(List<SetListItem> items, HashSet<string> seen, string name, bool isSearch)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            var key = $"{name}|{(isSearch ? "S" : "N")}";
            if (!seen.Add(key))
            {
                return;
            }

            items.Add(new SetListItem(name, isSearch));
        }

        private static string TryGetName(object obj)
        {
            if (obj == null) return null;

            try
            {
                var prop = obj.GetType().GetProperty("name");
                if (prop != null && prop.PropertyType == typeof(string))
                {
                    return prop.GetValue(obj) as string;
                }
            }
            catch
            {
            }

            try
            {
                var prop = obj.GetType().GetProperty("Name");
                if (prop != null && prop.PropertyType == typeof(string))
                {
                    return prop.GetValue(obj) as string;
                }
            }
            catch
            {
            }

            return null;
        }

        private static bool? TryGetSearchFlag(object obj)
        {
            return TryGetBoolProperty(obj, "IsSearch", "IsSearchSet", "IsSearchSelection");
        }

        private static bool? TryGetBoolProperty(object obj, params string[] names)
        {
            if (obj == null || names == null) return null;

            foreach (var name in names)
            {
                try
                {
                    var prop = obj.GetType().GetProperty(name);
                    if (prop != null && prop.PropertyType == typeof(bool))
                    {
                        return (bool)prop.GetValue(obj);
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private void UpdateRuleDefinitionInputs(bool clearText, bool focusSetName = false)
        {
            if (_zonesPage?.RuleTargetDefinitionCombo == null)
            {
                return;
            }

            var definition = _zonesPage.RuleTargetDefinitionCombo.SelectedValue is SpaceMapperTargetDefinition def
                ? def
                : SpaceMapperTargetDefinition.EntireModel;

            var usesSet = UsesSetSearchName(definition);
            var usesLevels = UsesLevels(definition);

            _zonesPage.RuleSetNameBox.IsEnabled = usesSet;
            _zonesPage.RuleSetDropDownButton.IsEnabled = usesSet;
            _zonesPage.RuleMinLevelBox.IsEnabled = usesLevels;
            _zonesPage.RuleMaxLevelBox.IsEnabled = usesLevels;

            if (_zonesPage.RuleSetNameHintText != null)
            {
                _zonesPage.RuleSetNameHintText.Visibility = usesSet ? Visibility.Collapsed : Visibility.Visible;
            }

            if (_zonesPage.RuleMinMaxHintText != null)
            {
                _zonesPage.RuleMinMaxHintText.Visibility = usesLevels ? Visibility.Collapsed : Visibility.Visible;
            }

            if (clearText)
            {
                if (!usesSet)
                {
                    _zonesPage.RuleSetNameBox.Text = NotApplicableText;
                }
                else if (IsNotApplicable(_zonesPage.RuleSetNameBox.Text))
                {
                    _zonesPage.RuleSetNameBox.Text = string.Empty;
                }

                if (!usesLevels)
                {
                    _zonesPage.RuleMinLevelBox.Text = NotApplicableText;
                    _zonesPage.RuleMaxLevelBox.Text = NotApplicableText;
                }
                else
                {
                    if (IsNotApplicable(_zonesPage.RuleMinLevelBox.Text))
                    {
                        _zonesPage.RuleMinLevelBox.Text = string.Empty;
                    }

                    if (IsNotApplicable(_zonesPage.RuleMaxLevelBox.Text))
                    {
                        _zonesPage.RuleMaxLevelBox.Text = string.Empty;
                    }

                    if (string.IsNullOrWhiteSpace(_zonesPage.RuleMinLevelBox.Text))
                    {
                        _zonesPage.RuleMinLevelBox.Text = "0";
                    }

                    if (string.IsNullOrWhiteSpace(_zonesPage.RuleMaxLevelBox.Text))
                    {
                        _zonesPage.RuleMaxLevelBox.Text = "0";
                    }
                }
            }

            if (focusSetName && usesSet)
            {
                _zonesPage.RuleSetNameBox.Focus();
            }
        }

        private static void NormalizeTargetRule(SpaceMapperTargetRule rule)
        {
            if (rule == null)
            {
                return;
            }

            if (rule.TargetDefinition == SpaceMapperTargetDefinition.SearchSet)
            {
                rule.TargetDefinition = SpaceMapperTargetDefinition.SelectionSet;
            }

            if (UsesSetSearchName(rule.TargetDefinition))
            {
                rule.MinLevel = null;
                rule.MaxLevel = null;
                return;
            }

            if (UsesLevels(rule.TargetDefinition))
            {
                if (rule.MinLevel == null && rule.MaxLevel == null)
                {
                    rule.MinLevel = 0;
                    rule.MaxLevel = 0;
                }
                else if (rule.MinLevel == null)
                {
                    rule.MinLevel = rule.MaxLevel;
                }
                else if (rule.MaxLevel == null)
                {
                    rule.MaxLevel = rule.MinLevel;
                }
                return;
            }

            rule.SetSearchName = null;
            rule.MinLevel = null;
            rule.MaxLevel = null;
        }

        private static bool IsRuleConfigured(SpaceMapperTargetRule rule)
        {
            if (rule == null)
            {
                return false;
            }

            if (UsesSetSearchName(rule.TargetDefinition))
            {
                return !string.IsNullOrWhiteSpace(rule.SetSearchName);
            }

            if (UsesLevels(rule.TargetDefinition))
            {
                if (rule.MinLevel == null || rule.MaxLevel == null)
                {
                    return false;
                }

                if (rule.MinLevel < 0 || rule.MaxLevel < 0)
                {
                    return false;
                }

                return rule.MinLevel <= rule.MaxLevel;
            }

            return true;
        }

        private void AddMapping()
        {
            if (!ValidateMappingBuilder())
            {
                return;
            }

            var name = _mappingPage.MappingNameBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                name = $"Mapping {Mappings.Count + 1}";
            }

            var zoneCategory = string.IsNullOrWhiteSpace(_mappingPage.ZoneCategoryBox.Text)
                ? null
                : _mappingPage.ZoneCategoryBox.Text.Trim();
            var zoneProperty = string.IsNullOrWhiteSpace(_mappingPage.ZonePropertyBox.Text)
                ? null
                : _mappingPage.ZonePropertyBox.Text.Trim();
            var targetCategory = string.IsNullOrWhiteSpace(_mappingPage.TargetCategoryBox.Text)
                ? null
                : _mappingPage.TargetCategoryBox.Text.Trim();
            var targetProperty = string.IsNullOrWhiteSpace(_mappingPage.TargetPropertyBox.Text)
                ? null
                : _mappingPage.TargetPropertyBox.Text.Trim();
            var writeMode = _mappingPage.WriteModeCombo.SelectedItem is WriteMode wm
                ? wm
                : WriteMode.Overwrite;
            var multiZone = _mappingPage.MultiZoneCombo.SelectedItem is MultiZoneCombineMode mz
                ? mz
                : MultiZoneCombineMode.First;
            var appendSeparator = _mappingPage.AppendSeparatorBox.Text ?? string.Empty;
            var editable = _mappingPage.EditableCheckBox.IsChecked == true;

            var mapping = new SpaceMapperMappingDefinition
            {
                Name = name,
                ZoneCategory = zoneCategory,
                ZonePropertyName = zoneProperty,
                TargetCategory = targetCategory,
                TargetPropertyName = targetProperty,
                WriteMode = writeMode,
                MultiZoneCombineMode = multiZone,
                AppendSeparator = appendSeparator,
                IsEditable = editable
            };
            Mappings.Add(mapping);
            _mappingPage.MappingGrid.SelectedItem = mapping;
            _mappingPage.MappingGrid.ScrollIntoView(mapping);
        }

        private void DeleteSelectedMapping()
        {
            if (Mappings.Count == 0)
            {
                MessageBox.Show("There are no mappings to delete.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (_mappingPage.MappingGrid.SelectedItem is SpaceMapperMappingDefinition map)
            {
                Mappings.Remove(map);
                return;
            }

            MessageBox.Show("Select a mapping to delete.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void OpenScraper()
        {
            var profile = GetSelectedScraperProfileName() ?? "Default";
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
            DockPaneVisibilitySnapshot paneSnapshot = null;
            SpaceMapperRunProgressHost progressHost = null;
            SpaceMapperRunProgressState runProgress = null;
            if (!ValidateRunInputs())
            {
                return;
            }

            _runCts?.Cancel();
            _runCts?.Dispose();
            _runCts = new CancellationTokenSource();

            try
            {
                var request = BuildRequest();
                if (request == null)
                {
                    runProgress?.MarkFailed(new InvalidOperationException("Unable to build Space Mapper request."));
                    return;
                }

                var useMesh = request.ProcessingSettings?.ZoneContainmentEngine == SpaceMapperZoneContainmentEngine.MeshAccurate;
                if (_lastMeshAccurate.HasValue && _lastMeshAccurate.Value != useMesh)
                {
                    GeometryExtractor.ClearCache();
                }
                _lastMeshAccurate = useMesh;

                if (request.ProcessingSettings?.CloseDockPanesDuringRun == true)
                {
                    paneSnapshot = NavisworksDockPaneManager.HideDockPanes(settleDelay: TimeSpan.Zero);
                    WaitForDockPaneSettle(request.ProcessingSettings);
                }

                runProgress = new SpaceMapperRunProgressState();
                progressHost = SpaceMapperRunProgressHost.Show(runProgress, () => _runCts.Cancel());

                var service = new SpaceMapperService(MicroEngActions.Log);
                var result = service.RunWithProgress(request, null, runProgress, _runCts.Token);

                progressHost?.Close();
                ShowResults(result);
                SpaceMapperNav.Navigate(typeof(SpaceMapperStepResultsPage));
            }
            catch (OperationCanceledException)
            {
                runProgress?.MarkCancelled();
            }
            catch (Exception ex)
            {
                runProgress?.MarkFailed(ex);
                MessageBox.Show($"Space Mapper failed: {ex.Message}", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                progressHost?.Close();
                progressHost?.Dispose();
                NavisworksDockPaneManager.RestoreDockPanes(paneSnapshot);
            }
        }

        private void RunVariationCheckReport()
        {
            DockPaneVisibilitySnapshot paneSnapshot = null;
            SpaceMapperRunProgressHost progressHost = null;
            SpaceMapperRunProgressState runProgress = null;
            if (!ValidateVariationCheckInputs())
            {
                return;
            }

            _runCts?.Cancel();
            _runCts?.Dispose();
            _runCts = new CancellationTokenSource();

            try
            {
                runProgress = new SpaceMapperRunProgressState();
                runProgress.Start();
                runProgress.SetStage(SpaceMapperRunStage.ResolvingInputs, "Resolving zones and targets...");

                var request = BuildRequest();
                if (request == null)
                {
                    runProgress.MarkFailed(new InvalidOperationException("Unable to build Space Mapper request."));
                    return;
                }

                var doc = NavisApp.ActiveDocument;
                if (doc == null)
                {
                    runProgress.MarkFailed(new InvalidOperationException("No active document."));
                    return;
                }

                var session = SpaceMapperService.GetSession(request.ScraperProfileName);
                var requiresSession = request.ZoneSource == ZoneSourceType.DataScraperZones;
                if (requiresSession && session == null)
                {
                    runProgress.MarkFailed(new InvalidOperationException(
                        $"No Data Scraper sessions found for profile '{request.ScraperProfileName ?? "Default"}'."));
                    return;
                }

                if (session != null)
                {
                    DataScraperCache.LastSession = session;
                }

                if (request.ProcessingSettings?.CloseDockPanesDuringRun == true)
                {
                    paneSnapshot = NavisworksDockPaneManager.HideDockPanes(settleDelay: TimeSpan.Zero);
                    WaitForDockPaneSettle(request.ProcessingSettings);
                }

                progressHost = SpaceMapperRunProgressHost.Show(runProgress, () => _runCts.Cancel());

                var resolved = SpaceMapperService.ResolveData(request, doc, session);
                runProgress.SetTotals(resolved.ZoneModels.Count, resolved.TargetModels.Count);

                if (!resolved.ZoneModels.Any())
                {
                    runProgress.MarkFailed(new InvalidOperationException("No zones found."));
                    return;
                }

                if (!resolved.TargetModels.Any())
                {
                    runProgress.MarkFailed(new InvalidOperationException("No targets found for the selected rules."));
                    return;
                }

                runProgress.SetStage(SpaceMapperRunStage.ExtractingGeometry, "Preparing geometry...");

                var dataset = SpaceMapperService.BuildGeometryData(
                    resolved,
                    request.ProcessingSettings,
                    buildTargetGeometry: true,
                    _runCts.Token,
                    runProgress);

                if (!dataset.Zones.Any())
                {
                    runProgress.MarkFailed(new InvalidOperationException("No zones with bounding boxes found."));
                    return;
                }

                if (!dataset.TargetsForEngine.Any())
                {
                    runProgress.MarkFailed(new InvalidOperationException("No targets with bounding boxes found."));
                    return;
                }

                Interlocked.Exchange(ref runProgress.ZonesProcessed, runProgress.ZonesTotal);
                Interlocked.Exchange(ref runProgress.TargetsProcessed, runProgress.TargetsTotal);

                runProgress.SetStage(SpaceMapperRunStage.ComputingIntersections, "Running variation check...", "Preparing variants...");

                var options = new SpaceMapperComparisonOptions
                {
                    IncludeAllBoundsVariants = true,
                    IncludeCpuGpuComparison = true,
                    BenchmarkMode = SpaceMapperBenchmarkMode.ComputeOnly,
                    MaxZones = int.MaxValue,
                    MaxTargets = int.MaxValue
                };

                var input = new SpaceMapperComparisonInput
                {
                    Request = request,
                    Dataset = dataset,
                    Options = options,
                    ModelName = GetDocumentTitle(doc)
                };

                var progress = new Progress<SpaceMapperComparisonProgress>(update =>
                {
                    if (update == null)
                    {
                        return;
                    }

                    var detail = string.IsNullOrWhiteSpace(update.Stage) ? "Running variants..." : update.Stage;
                    if (update.Percentage >= 0)
                    {
                        detail = $"{detail} ({update.Percentage}%)";
                    }

                    runProgress.SetStage(SpaceMapperRunStage.ComputingIntersections, "Running variation check...", detail);
                });

                var runner = new SpaceMapperComparisonRunner(MicroEngActions.Log);
                var output = runner.Run(input, progress, _runCts.Token);

                runProgress.SetStage(SpaceMapperRunStage.Finalizing, "Writing report...");

                var reportPath = BuildVariationReportPath(doc, request.TemplateName);
                var reportDir = Path.GetDirectoryName(reportPath);
                if (!string.IsNullOrWhiteSpace(reportDir))
                {
                    Directory.CreateDirectory(reportDir);
                }

                File.WriteAllText(reportPath, output?.Markdown ?? string.Empty, Encoding.UTF8);

                if (!string.IsNullOrWhiteSpace(output?.Json))
                {
                    var jsonPath = Path.ChangeExtension(reportPath, ".json");
                    File.WriteAllText(jsonPath, output.Json, Encoding.UTF8);
                }

                runProgress.MarkCompleted();
                progressHost?.Close();

                MessageBox.Show($"Variation check report saved to:\n{reportPath}", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (OperationCanceledException)
            {
                runProgress?.MarkCancelled();
            }
            catch (Exception ex)
            {
                runProgress?.MarkFailed(ex);
                MessageBox.Show($"Variation check failed: {ex.Message}", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                progressHost?.Close();
                progressHost?.Dispose();
                NavisworksDockPaneManager.RestoreDockPanes(paneSnapshot);
            }
        }

        private SpaceMapperProcessingSettings BuildSettings()
        {
            var containmentEngine = GetSelectedZoneContainmentEngine();
            var resolutionStrategy = GetSelectedZoneResolutionStrategy();
            var zoneBoundsMode = GetSelectedZoneBoundsMode();
            var targetBoundsMode = GetSelectedTargetBoundsMode();
            var treatPartial = _processingPage.TreatPartialCheck.IsChecked == true;
            var tagPartial = _processingPage.TagPartialCheck.IsChecked == true;
            var writeBehavior = _processingPage.WriteZoneBehaviorCheck?.IsChecked == true;
            var writeContainmentPercent = _processingPage.WriteContainmentPercentCheck?.IsChecked == true;
            var containmentCalculationMode = GetSelectedContainmentCalculationMode();
            var gpuRayCount = GetSelectedGpuRayCount();
            var enableZoneOffsets = _processingPage.EnableZoneOffsetsCheck?.IsChecked == true;
            var enableOffsetAreaPass = _processingPage.EnableOffsetAreaPassCheck?.IsChecked == true;
            var writeOffsetMatch = _processingPage.WriteOffsetMatchPropertyCheck?.IsChecked == true;

            if (!enableZoneOffsets)
            {
                enableOffsetAreaPass = false;
                writeOffsetMatch = false;
            }
            else if (!enableOffsetAreaPass)
            {
                writeOffsetMatch = false;
            }

            // Tier A/B: user controls target representation even when MeshAccurate is selected.
            if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                treatPartial = false;
                tagPartial = false;
                writeContainmentPercent = false;
            }

            var settings = new SpaceMapperProcessingSettings
            {
                ProcessingMode = SpaceMapperProcessingMode.Auto,
                TreatPartialAsContained = treatPartial,
                TagPartialSeparately = tagPartial,
                EnableMultipleZones = _processingPage.EnableMultiZoneCheck.IsChecked == true,
                ExcludeZonesFromTargets = _processingPage.ExcludeZonesFromTargetsCheck?.IsChecked == true,
                Offset3D = ParseDouble(_processingPage.Offset3DBox.Text),
                OffsetTop = ParseDouble(_processingPage.OffsetTopBox.Text),
                OffsetBottom = ParseDouble(_processingPage.OffsetBottomBox.Text),
                OffsetSides = ParseDouble(_processingPage.OffsetSidesBox.Text),
                Units = _processingPage.UnitsBox.Text,
                OffsetMode = _processingPage.OffsetModeBox.Text,
                MaxThreads = null,
                BatchSize = null,
                IndexGranularity = 0,
                PerformancePreset = InferPresetFromBounds(zoneBoundsMode, targetBoundsMode, containmentEngine),
                FastTraversalMode = SpaceMapperFastTraversalMode.Auto,
                WriteZoneBehaviorProperty = writeBehavior,
                WriteZoneContainmentPercentProperty = writeContainmentPercent,
                ContainmentCalculationMode = containmentCalculationMode,
                ZoneBehaviorCategory = _processingPage.ZoneBehaviorCategoryBox.Text?.Trim(),
                ZoneBehaviorPropertyName = _processingPage.ZoneBehaviorPropertyBox.Text?.Trim(),
                ZoneBehaviorContainedValue = _processingPage.ZoneBehaviorContainedBox.Text?.Trim(),
                ZoneBehaviorPartialValue = _processingPage.ZoneBehaviorPartialBox.Text?.Trim(),
                WritebackStrategy = SpaceMapperWritebackStrategy.OptimizedSingleCategory,
                ShowInternalPropertiesDuringWriteback = _processingPage.ShowInternalWritebackCheck?.IsChecked == true,
                CloseDockPanesDuringRun = _processingPage.CloseDockPanesCheck?.IsChecked == true,
                DockPaneCloseDelaySeconds = Math.Max(0, ParseDouble(_processingPage.DockPaneDelayBox?.Text)),
                SkipUnchangedWriteback = _processingPage.SkipUnchangedWritebackCheck?.IsChecked == true,
                PackWritebackProperties = _processingPage.PackWritebackCheck?.IsChecked == true,
                ZoneBoundsMode = zoneBoundsMode,
                ZoneKDopVariant = GetSelectedZoneKDopVariant(),
                TargetBoundsMode = targetBoundsMode,
                TargetKDopVariant = GetSelectedTargetKDopVariant(),
                TargetMidpointMode = GetSelectedTargetMidpointMode(),
                ZoneContainmentEngine = containmentEngine,
                ZoneResolutionStrategy = resolutionStrategy,
                UseOriginPointOnly = targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint,
                GpuRayCount = gpuRayCount,
                EnableZoneOffsets = enableZoneOffsets,
                EnableOffsetAreaPass = enableOffsetAreaPass,
                WriteZoneOffsetMatchProperty = writeOffsetMatch
            };

            if (!settings.EnableZoneOffsets)
            {
                settings.Offset3D = 0;
                settings.OffsetTop = 0;
                settings.OffsetBottom = 0;
                settings.OffsetSides = 0;
                settings.OffsetMode = "None";
            }

            return settings;
        }

        private static TimeSpan GetDockPaneDelay(SpaceMapperProcessingSettings settings)
        {
            if (settings == null)
            {
                return TimeSpan.Zero;
            }

            var delaySeconds = settings.DockPaneCloseDelaySeconds;
            if (delaySeconds <= 0)
            {
                return TimeSpan.Zero;
            }

            return TimeSpan.FromSeconds(delaySeconds);
        }

        private static void WaitForDockPaneSettle(SpaceMapperProcessingSettings settings)
        {
            var delay = GetDockPaneDelay(settings);
            if (delay <= TimeSpan.Zero)
            {
                return;
            }

            var frame = new DispatcherFrame();
            var timer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = delay
            };
            timer.Tick += (_, __) =>
            {
                timer.Stop();
                frame.Continue = false;
            };
            timer.Start();
            Dispatcher.PushFrame(frame);
        }

        private SpaceMapperRequest BuildRequest()
        {
            var settings = BuildSettings();
            return new SpaceMapperRequest
            {
                TemplateName = _setupPage.TemplateCombo.SelectedItem?.ToString(),
                ScraperProfileName = GetSelectedScraperProfileName(),
                Scope = SpaceMapperScope.EntireModel,
                ZoneSource = _zonesPage.ZoneSourceCombo.SelectedItem is ZoneSourceType zs ? zs : ZoneSourceType.DataScraperZones,
                ZoneSetName = _zonesPage.ZoneSetBox.Text,
                TargetRules = TargetRules.ToList(),
                Mappings = Mappings.ToList(),
                ProcessingSettings = settings
            };
        }

        private void TriggerLivePreflight()
        {
            // Preflight/benchmark UI removed; keep as a no-op for existing change hooks.
        }

        private SpaceMapperZoneBoundsMode GetSelectedZoneBoundsMode()
        {
            var v = (int)Math.Round(_processingPage.ZoneBoundsSlider.Value);
            return v switch
            {
                1 => SpaceMapperZoneBoundsMode.Obb,
                2 => SpaceMapperZoneBoundsMode.KDop,
                3 => SpaceMapperZoneBoundsMode.Hull,
                _ => SpaceMapperZoneBoundsMode.Aabb
            };
        }

        private SpaceMapperTargetBoundsMode GetSelectedTargetBoundsMode()
        {
            var v = (int)Math.Round(_processingPage.TargetBoundsSlider.Value);
            return v switch
            {
                0 => SpaceMapperTargetBoundsMode.Midpoint,
                1 => SpaceMapperTargetBoundsMode.Aabb,
                2 => SpaceMapperTargetBoundsMode.Obb,
                3 => SpaceMapperTargetBoundsMode.KDop,
                4 => SpaceMapperTargetBoundsMode.Hull,
                _ => SpaceMapperTargetBoundsMode.Aabb
            };
        }

        private SpaceMapperZoneContainmentEngine GetSelectedZoneContainmentEngine()
        {
            if (_processingPage.ZoneContainmentEngineCombo == null)
            {
                return SpaceMapperZoneContainmentEngine.BoundsFast;
            }

            return _processingPage.ZoneContainmentEngineCombo.SelectedIndex switch
            {
                1 => SpaceMapperZoneContainmentEngine.MeshAccurate,
                _ => SpaceMapperZoneContainmentEngine.BoundsFast
            };
        }

        private SpaceMapperZoneResolutionStrategy GetSelectedZoneResolutionStrategy()
        {
            if (_processingPage.ZoneResolutionStrategyCombo == null)
            {
                return SpaceMapperZoneResolutionStrategy.MostSpecific;
            }

            return _processingPage.ZoneResolutionStrategyCombo.SelectedIndex switch
            {
                1 => SpaceMapperZoneResolutionStrategy.LargestOverlap,
                2 => SpaceMapperZoneResolutionStrategy.FirstMatch,
                _ => SpaceMapperZoneResolutionStrategy.MostSpecific
            };
        }

        private SpaceMapperContainmentCalculationMode GetSelectedContainmentCalculationMode()
        {
            if (_processingPage.ContainmentCalculationCombo == null)
            {
                return SpaceMapperContainmentCalculationMode.Auto;
            }

            return _processingPage.ContainmentCalculationCombo.SelectedIndex switch
            {
                1 => SpaceMapperContainmentCalculationMode.SamplePoints,
                2 => SpaceMapperContainmentCalculationMode.SamplePointsDense,
                3 => SpaceMapperContainmentCalculationMode.TargetGeometry,
                4 => SpaceMapperContainmentCalculationMode.TargetGeometryGpu,
                5 => SpaceMapperContainmentCalculationMode.BoundsOverlap,
                _ => SpaceMapperContainmentCalculationMode.Auto
            };
        }

        private int GetSelectedGpuRayCount()
        {
            if (_processingPage.GpuRayAccuracyCombo == null)
            {
                return 2;
            }

            return _processingPage.GpuRayAccuracyCombo.SelectedIndex switch
            {
                0 => 1,
                _ => 2
            };
        }

        private SpaceMapperKDopVariant GetSelectedZoneKDopVariant()
        {
            return _processingPage.ZoneKDopVariantCombo.SelectedIndex switch
            {
                0 => SpaceMapperKDopVariant.KDop8,
                2 => SpaceMapperKDopVariant.KDop18,
                _ => SpaceMapperKDopVariant.KDop14
            };
        }

        private SpaceMapperKDopVariant GetSelectedTargetKDopVariant()
        {
            return _processingPage.TargetKDopVariantCombo.SelectedIndex switch
            {
                0 => SpaceMapperKDopVariant.KDop8,
                2 => SpaceMapperKDopVariant.KDop18,
                _ => SpaceMapperKDopVariant.KDop14
            };
        }

        private SpaceMapperMidpointMode GetSelectedTargetMidpointMode()
        {
            return _processingPage.TargetMidpointModeCombo.SelectedIndex switch
            {
                1 => SpaceMapperMidpointMode.BoundingBoxBottomCenter,
                _ => SpaceMapperMidpointMode.BoundingBoxCenter
            };
        }

        private static SpaceMapperPerformancePreset InferPresetFromBounds(
            SpaceMapperZoneBoundsMode zoneMode,
            SpaceMapperTargetBoundsMode targetMode,
            SpaceMapperZoneContainmentEngine containmentEngine = SpaceMapperZoneContainmentEngine.BoundsFast)
        {
            if (containmentEngine == SpaceMapperZoneContainmentEngine.MeshAccurate)
            {
                return SpaceMapperPerformancePreset.Accurate;
            }

            if (targetMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                return SpaceMapperPerformancePreset.Fast;
            }

            if (zoneMode == SpaceMapperZoneBoundsMode.Hull
                || targetMode == SpaceMapperTargetBoundsMode.Hull
                || zoneMode == SpaceMapperZoneBoundsMode.KDop
                || targetMode == SpaceMapperTargetBoundsMode.KDop
                || zoneMode == SpaceMapperZoneBoundsMode.Obb
                || targetMode == SpaceMapperTargetBoundsMode.Obb)
            {
                return SpaceMapperPerformancePreset.Accurate;
            }

            return SpaceMapperPerformancePreset.Normal;
        }

        private void OnBoundsModesChanged()
        {
            var containmentEngine = GetSelectedZoneContainmentEngine();
            var resolutionStrategy = GetSelectedZoneResolutionStrategy();

            if (_processingPage.TargetBoundsSlider != null)
            {
                // Tier A/B: allow user to choose target bounds even when MeshAccurate is selected.
                _processingPage.TargetBoundsSlider.IsEnabled = true;
            }

            var zoneMode = GetSelectedZoneBoundsMode();
            var targetMode = GetSelectedTargetBoundsMode();

            if (_processingPage.ZoneKDopVariantRow != null)
            {
                _processingPage.ZoneKDopVariantRow.Visibility = zoneMode == SpaceMapperZoneBoundsMode.KDop
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            }

            if (_processingPage.TargetKDopVariantRow != null)
            {
                _processingPage.TargetKDopVariantRow.Visibility = targetMode == SpaceMapperTargetBoundsMode.KDop
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            }

            if (_processingPage.TargetMidpointModeRow != null)
            {
                _processingPage.TargetMidpointModeRow.Visibility = targetMode == SpaceMapperTargetBoundsMode.Midpoint
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            }

            if (targetMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                if (_processingPage.TreatPartialCheck.IsChecked == true)
                {
                    _processingPage.TreatPartialCheck.IsChecked = false;
                }
                if (_processingPage.TagPartialCheck.IsChecked == true)
                {
                    _processingPage.TagPartialCheck.IsChecked = false;
                }
                _processingPage.TreatPartialCheck.IsEnabled = false;
                _processingPage.TagPartialCheck.IsEnabled = false;
            }
            else
            {
                _processingPage.TreatPartialCheck.IsEnabled = true;
                _processingPage.TagPartialCheck.IsEnabled = true;
            }

            if (_processingPage.ZoneContainmentHintText != null)
            {
                _processingPage.ZoneContainmentHintText.Text = GetZoneContainmentHintText(containmentEngine, resolutionStrategy);
            }

            UpdateZoneBehaviorInputs();
            UpdateFastTraversalUi();
            TriggerLivePreflight();
        }

        private void UpdateFastTraversalUi()
        {
            // Fast traversal UI removed; keep no-op for existing call sites.
        }

        private static string GetZoneContainmentHintText(
            SpaceMapperZoneContainmentEngine engine,
            SpaceMapperZoneResolutionStrategy strategy)
        {
            var prefix = engine == SpaceMapperZoneContainmentEngine.MeshAccurate
                ? "Mesh accurate engine: Tier A = Midpoint (single point); Tier B = AABB/OBB/k-DOP/Hull (multi-sample)."
                : "Bounds engine: uses planes/AABB checks; Midpoint reduces targets to a single point.";

            return strategy switch
            {
                SpaceMapperZoneResolutionStrategy.MostSpecific =>
                    $"{prefix} Most specific picks the smallest enclosing zone.",
                SpaceMapperZoneResolutionStrategy.LargestOverlap =>
                    $"{prefix} Largest overlap uses AABB overlap volume to choose a single zone.",
                SpaceMapperZoneResolutionStrategy.FirstMatch =>
                    $"{prefix} First match picks the first zone in order.",
                _ => prefix
            };
        }

        private static string FormatDuration(TimeSpan span)
        {
            if (span.TotalSeconds >= 1)
            {
                return $"{span.TotalSeconds:0.00}s";
            }
            return $"{span.TotalMilliseconds:0}ms";
        }

        private sealed class ValidationIssue
        {
            public string Message { get; set; }
            public Action FocusAction { get; set; }
            public Type PageType { get; set; }
        }

        private bool ValidateVariationCheckInputs()
        {
            var issues = new List<ValidationIssue>();
            ValidateSetupInputs(issues);
            ValidateZonesTargetsInputs(issues);
            ValidateProcessingInputs(issues);

            if (issues.Count == 0)
            {
                return true;
            }

            ShowValidationIssues(issues);
            NavigateToIssue(issues[0]);
            return false;
        }

        private bool ValidateRunInputs()
        {
            var issues = new List<ValidationIssue>();
            ValidateSetupInputs(issues);
            ValidateZonesTargetsInputs(issues);
            ValidateProcessingInputs(issues);
            ValidateMappingInputs(issues);

            if (issues.Count == 0)
            {
                return true;
            }

            ShowValidationIssues(issues);
            NavigateToIssue(issues[0]);
            return false;
        }

        private void ValidateSetupInputs(List<ValidationIssue> issues)
        {
            if (string.IsNullOrWhiteSpace(_setupPage.TemplateCombo.SelectedItem?.ToString()))
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepSetupPage),
                    Message = "Select a Space Mapper profile.",
                    FocusAction = () => _setupPage.TemplateCombo.Focus()
                });
            }
        }

        private void ValidateZonesTargetsInputs(List<ValidationIssue> issues)
        {
            if (_zonesPage.ZoneSourceCombo.SelectedItem == null)
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepZonesTargetsPage),
                    Message = "Select a zone source.",
                    FocusAction = () => _zonesPage.ZoneSourceCombo.Focus()
                });
            }

            var zoneSource = _zonesPage.ZoneSourceCombo.SelectedItem is ZoneSourceType zs
                ? zs
                : ZoneSourceType.DataScraperZones;

            if ((zoneSource == ZoneSourceType.ZoneSelectionSet || zoneSource == ZoneSourceType.ZoneSearchSet)
                && string.IsNullOrWhiteSpace(_zonesPage.ZoneSetBox.Text))
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepZonesTargetsPage),
                    Message = "Enter a Zone Set/Search name.",
                    FocusAction = () => _zonesPage.ZoneSetBox.Focus()
                });
            }

            if (zoneSource == ZoneSourceType.DataScraperZones)
            {
                var profile = GetSelectedScraperProfileName();
                if (string.IsNullOrWhiteSpace(profile))
                {
                    issues.Add(new ValidationIssue
                    {
                        PageType = typeof(SpaceMapperStepSetupPage),
                        Message = "Select a Data Scraper profile or change the Zone Source.",
                        FocusAction = () => _setupPage.ScraperProfileCombo.Focus()
                    });
                }
                else if (SpaceMapperService.GetSession(profile) == null)
                {
                    issues.Add(new ValidationIssue
                    {
                        PageType = typeof(SpaceMapperStepSetupPage),
                        Message = $"Run Data Scraper to create a profile for '{profile}'.",
                        FocusAction = () => _setupPage.OpenScraperButton.Focus()
                    });
                }
            }

            if (TargetRules.Count == 0)
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepZonesTargetsPage),
                    Message = "Add at least one target rule.",
                    FocusAction = () => _zonesPage.AddRuleButton.Focus()
                });
                return;
            }

            var enabledRules = TargetRules.Where(r => r.Enabled).ToList();
            if (enabledRules.Count == 0)
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepZonesTargetsPage),
                    Message = "Enable at least one target rule.",
                    FocusAction = () => _zonesPage.TargetRulesGrid.Focus()
                });
                return;
            }

            foreach (var rule in enabledRules)
            {
                if (string.IsNullOrWhiteSpace(rule.Name))
                {
                    issues.Add(new ValidationIssue
                    {
                        PageType = typeof(SpaceMapperStepZonesTargetsPage),
                        Message = "Each enabled rule needs a name.",
                        FocusAction = () => FocusTargetRule(rule)
                    });
                }

                if (UsesSetSearchName(rule.TargetDefinition)
                    && string.IsNullOrWhiteSpace(rule.SetSearchName))
                {
                    issues.Add(new ValidationIssue
                    {
                        PageType = typeof(SpaceMapperStepZonesTargetsPage),
                        Message = $"Rule '{rule.Name ?? "Unnamed"}' needs a Set/Search name.",
                        FocusAction = () => FocusTargetRule(rule)
                    });
                }

                if (UsesLevels(rule.TargetDefinition))
                {
                    if (rule.MinLevel == null || rule.MaxLevel == null)
                    {
                        issues.Add(new ValidationIssue
                        {
                            PageType = typeof(SpaceMapperStepZonesTargetsPage),
                            Message = $"Rule '{rule.Name ?? "Unnamed"}' needs Min/Max Level.",
                            FocusAction = () => FocusTargetRule(rule)
                        });
                    }
                    else if (rule.MinLevel < 0 || rule.MaxLevel < 0)
                    {
                        issues.Add(new ValidationIssue
                        {
                            PageType = typeof(SpaceMapperStepZonesTargetsPage),
                            Message = $"Rule '{rule.Name ?? "Unnamed"}' has a negative Level.",
                            FocusAction = () => FocusTargetRule(rule)
                        });
                    }
                    else if (rule.MinLevel > rule.MaxLevel)
                    {
                        issues.Add(new ValidationIssue
                        {
                            PageType = typeof(SpaceMapperStepZonesTargetsPage),
                            Message = $"Rule '{rule.Name ?? "Unnamed"}' has Min Level greater than Max Level.",
                            FocusAction = () => FocusTargetRule(rule)
                        });
                    }
                }
            }
        }

        private void ValidateProcessingInputs(List<ValidationIssue> issues)
        {
            if (!TryParseDoubleInput(_processingPage.Offset3DBox.Text, out _))
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepProcessingPage),
                    Message = "3D Offset must be a number.",
                    FocusAction = () => _processingPage.Offset3DBox.Focus()
                });
            }

            if (!TryParseDoubleInput(_processingPage.OffsetSidesBox.Text, out _))
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepProcessingPage),
                    Message = "Offset Sides must be a number.",
                    FocusAction = () => _processingPage.OffsetSidesBox.Focus()
                });
            }

            if (!TryParseDoubleInput(_processingPage.OffsetBottomBox.Text, out _))
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepProcessingPage),
                    Message = "Offset Bottom must be a number.",
                    FocusAction = () => _processingPage.OffsetBottomBox.Focus()
                });
            }

            if (!TryParseDoubleInput(_processingPage.OffsetTopBox.Text, out _))
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepProcessingPage),
                    Message = "Offset Top must be a number.",
                    FocusAction = () => _processingPage.OffsetTopBox.Focus()
                });
            }

            if (string.IsNullOrWhiteSpace(_processingPage.UnitsBox.Text))
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepProcessingPage),
                    Message = "Units cannot be blank.",
                    FocusAction = () => _processingPage.UnitsBox.Focus()
                });
            }

            if (string.IsNullOrWhiteSpace(_processingPage.OffsetModeBox.Text))
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepProcessingPage),
                    Message = "Offset Mode cannot be blank.",
                    FocusAction = () => _processingPage.OffsetModeBox.Focus()
                });
            }

            if (_processingPage.WriteZoneBehaviorCheck?.IsChecked == true)
            {
                if (string.IsNullOrWhiteSpace(_processingPage.ZoneBehaviorCategoryBox.Text))
                {
                    issues.Add(new ValidationIssue
                    {
                        PageType = typeof(SpaceMapperStepProcessingPage),
                        Message = "Zone Behaviour Category cannot be blank.",
                        FocusAction = () => _processingPage.ZoneBehaviorCategoryBox.Focus()
                    });
                }

                if (string.IsNullOrWhiteSpace(_processingPage.ZoneBehaviorPropertyBox.Text))
                {
                    issues.Add(new ValidationIssue
                    {
                        PageType = typeof(SpaceMapperStepProcessingPage),
                        Message = "Zone Behaviour Property cannot be blank.",
                        FocusAction = () => _processingPage.ZoneBehaviorPropertyBox.Focus()
                    });
                }

                if (string.IsNullOrWhiteSpace(_processingPage.ZoneBehaviorContainedBox.Text))
                {
                    issues.Add(new ValidationIssue
                    {
                        PageType = typeof(SpaceMapperStepProcessingPage),
                        Message = "Contained Value cannot be blank.",
                        FocusAction = () => _processingPage.ZoneBehaviorContainedBox.Focus()
                    });
                }

                if (string.IsNullOrWhiteSpace(_processingPage.ZoneBehaviorPartialBox.Text))
                {
                    issues.Add(new ValidationIssue
                    {
                        PageType = typeof(SpaceMapperStepProcessingPage),
                        Message = "Partial Value cannot be blank.",
                        FocusAction = () => _processingPage.ZoneBehaviorPartialBox.Focus()
                    });
                }
            }

        }

        private void ValidateMappingInputs(List<ValidationIssue> issues)
        {
            if (Mappings.Count == 0)
            {
                issues.Add(new ValidationIssue
                {
                    PageType = typeof(SpaceMapperStepMappingPage),
                    Message = "Add at least one attribute mapping.",
                    FocusAction = () => _mappingPage.AddMappingButton.Focus()
                });
                return;
            }

            for (int i = 0; i < Mappings.Count; i++)
            {
                var mapping = Mappings[i];
                if (string.IsNullOrWhiteSpace(mapping.ZonePropertyName))
                {
                    issues.Add(new ValidationIssue
                    {
                        PageType = typeof(SpaceMapperStepMappingPage),
                        Message = $"Mapping row {i + 1}: Zone Property is required.",
                        FocusAction = () => FocusMappingRow(mapping)
                    });
                }

                if (string.IsNullOrWhiteSpace(mapping.TargetCategory))
                {
                    issues.Add(new ValidationIssue
                    {
                        PageType = typeof(SpaceMapperStepMappingPage),
                        Message = $"Mapping row {i + 1}: Target Category is required.",
                        FocusAction = () => FocusMappingRow(mapping)
                    });
                }

                if (string.IsNullOrWhiteSpace(mapping.TargetPropertyName))
                {
                    issues.Add(new ValidationIssue
                    {
                        PageType = typeof(SpaceMapperStepMappingPage),
                        Message = $"Mapping row {i + 1}: Target Property is required.",
                        FocusAction = () => FocusMappingRow(mapping)
                    });
                }
            }
        }

        private bool ValidateMappingBuilder()
        {
            if (string.IsNullOrWhiteSpace(_mappingPage.ZonePropertyBox.Text))
            {
                ShowValidationMessage("Enter a Zone Property to read.",
                    typeof(SpaceMapperStepMappingPage),
                    () => _mappingPage.ZonePropertyBox.Focus());
                return false;
            }

            if (string.IsNullOrWhiteSpace(_mappingPage.TargetCategoryBox.Text))
            {
                ShowValidationMessage("Enter a Target Category to write to.",
                    typeof(SpaceMapperStepMappingPage),
                    () => _mappingPage.TargetCategoryBox.Focus());
                return false;
            }

            if (string.IsNullOrWhiteSpace(_mappingPage.TargetPropertyBox.Text))
            {
                ShowValidationMessage("Enter a Target Property to write to.",
                    typeof(SpaceMapperStepMappingPage),
                    () => _mappingPage.TargetPropertyBox.Focus());
                return false;
            }

            if (_mappingPage.WriteModeCombo.SelectedItem == null)
            {
                ShowValidationMessage("Select a Write Mode.",
                    typeof(SpaceMapperStepMappingPage),
                    () => _mappingPage.WriteModeCombo.Focus());
                return false;
            }

            if (_mappingPage.MultiZoneCombo.SelectedItem == null)
            {
                ShowValidationMessage("Select a Multi-Zone mode.",
                    typeof(SpaceMapperStepMappingPage),
                    () => _mappingPage.MultiZoneCombo.Focus());
                return false;
            }

            return true;
        }

        private void ShowValidationIssues(IEnumerable<ValidationIssue> issues)
        {
            var messages = issues
                .Select(i => i.Message)
                .Where(m => !string.IsNullOrWhiteSpace(m))
                .Distinct()
                .ToList();

            if (messages.Count == 0)
            {
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("Fix the following before running Space Mapper:");
            foreach (var msg in messages)
            {
                sb.AppendLine($"- {msg}");
            }

            MessageBox.Show(sb.ToString(), "MicroEng", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void ShowValidationMessage(string message, Type pageType, Action focusAction)
        {
            MessageBox.Show(message, "MicroEng", MessageBoxButton.OK, MessageBoxImage.Warning);
            NavigateToIssue(new ValidationIssue
            {
                PageType = pageType,
                Message = message,
                FocusAction = focusAction
            });
        }

        private void NavigateToIssue(ValidationIssue issue)
        {
            if (issue == null)
            {
                return;
            }

            if (issue.PageType != null)
            {
                SpaceMapperNav.Navigate(issue.PageType);
            }

            if (issue.FocusAction != null)
            {
                Dispatcher.BeginInvoke(issue.FocusAction, DispatcherPriority.Background);
            }
        }

        private void FocusTargetRule(SpaceMapperTargetRule rule)
        {
            if (rule == null)
            {
                return;
            }

            _zonesPage.TargetRulesGrid.SelectedItem = rule;
            _zonesPage.TargetRulesGrid.ScrollIntoView(rule);
            _zonesPage.TargetRulesGrid.Focus();
        }

        private void FocusMappingRow(SpaceMapperMappingDefinition mapping)
        {
            if (mapping == null)
            {
                return;
            }

            _mappingPage.MappingGrid.SelectedItem = mapping;
            _mappingPage.MappingGrid.ScrollIntoView(mapping);
            _mappingPage.MappingGrid.Focus();
        }

        private static bool TryParseDoubleInput(string text, out double value)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                value = 0;
                return true;
            }

            if (!double.TryParse(text, out value))
            {
                return false;
            }

            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private static bool TryParseOptionalNonNegativeInt(string text, out int? value)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                value = null;
                return true;
            }

            if (int.TryParse(text, out var parsed) && parsed >= 0)
            {
                value = parsed;
                return true;
            }

            value = null;
            return false;
        }

        private static double ParseDouble(string text)
        {
            return double.TryParse(text, out var d) ? d : 0;
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
                sb.AppendLine($"Candidates: {result.Stats.CandidatePairs:N0} (avg {result.Stats.AvgCandidatesPerZone:0.##}/zone, max {result.Stats.MaxCandidatesPerZone:N0})");
                sb.AppendLine($"Preset used: {result.Stats.PresetUsed}");
                sb.AppendLine($"Mode: {result.Stats.ModeUsed}, Time: {result.Stats.Elapsed.TotalSeconds:0.00}s");
                sb.AppendLine($"Stages: resolve {FormatDuration(result.Stats.ResolveTime)}, build {FormatDuration(result.Stats.BuildGeometryTime)}, index {FormatDuration(result.Stats.BuildIndexTime)}, candidates {FormatDuration(result.Stats.CandidateQueryTime)}, narrow {FormatDuration(result.Stats.NarrowPhaseTime)}, write {FormatDuration(result.Stats.WriteBackTime)}");
            }
            if (!string.IsNullOrWhiteSpace(result?.ReportPath))
            {
                sb.AppendLine($"Report: {result.ReportPath}");
            }
            _resultsPage.ResultsSummaryBox.Text = sb.ToString();
            UpdateRunHealthUi(result);
        }

        private void UpdateRunHealthUi(SpaceMapperRunResult result)
        {
            var stats = result?.Stats;
            _lastTargetsWithoutBounds = result?.TargetsWithoutBounds?.Where(i => i != null).Distinct().ToList()
                ?? new List<ModelItem>();
            _lastTargetsUnmatched = result?.TargetsUnmatched?.Where(i => i != null).Distinct().ToList()
                ?? new List<ModelItem>();

            var missingBounds = _lastTargetsWithoutBounds.Count;
            var unmatched = _lastTargetsUnmatched.Count;
            var totalTargets = stats?.TargetsTotal ?? 0;
            var missingRatio = totalTargets > 0 ? (double)missingBounds / totalTargets : 0d;
            var unmatchedRatio = totalTargets > 0 ? (double)unmatched / totalTargets : 0d;
            var showStrip = (missingBounds > 0 && (missingBounds >= 10 || missingRatio >= 0.0005))
                || (unmatched > 0 && (unmatched >= 10 || unmatchedRatio >= 0.0005));
            var hasIssues = missingBounds > 0 || unmatched > 0;

            if (RunHealthStripControl != null)
            {
                RunHealthStripControl.Visibility = showStrip ? Visibility.Visible : Visibility.Collapsed;
            }
            if (RunHealthDetailsButtonControl != null)
            {
                RunHealthDetailsButtonControl.Visibility = showStrip ? Visibility.Visible : Visibility.Collapsed;
            }
            if (RunHealthTextControl != null)
            {
                RunHealthTextControl.Text = BuildRunHealthSummaryText(missingBounds, unmatched, missingRatio);
            }
            var warnIcon = missingRatio >= 0.001 || missingBounds >= 10;
            if (RunHealthIconInfoControl != null)
            {
                RunHealthIconInfoControl.Visibility = warnIcon ? Visibility.Collapsed : Visibility.Visible;
            }
            if (RunHealthIconWarningControl != null)
            {
                RunHealthIconWarningControl.Visibility = warnIcon ? Visibility.Visible : Visibility.Collapsed;
            }

            if (_resultsPage?.RunHealthChipPanel != null)
            {
                _resultsPage.RunHealthChipPanel.Visibility = hasIssues ? Visibility.Visible : Visibility.Collapsed;
            }
            if (_resultsPage?.RunHealthMissingBoundsChip != null)
            {
                _resultsPage.RunHealthMissingBoundsChip.Visibility = missingBounds > 0 ? Visibility.Visible : Visibility.Collapsed;
            }
            if (_resultsPage?.RunHealthMissingBoundsText != null)
            {
                _resultsPage.RunHealthMissingBoundsText.Text = $"Missing bounds: {missingBounds:N0}";
            }
            if (_resultsPage?.RunHealthUnmatchedChip != null)
            {
                _resultsPage.RunHealthUnmatchedChip.Visibility = unmatched > 0 ? Visibility.Visible : Visibility.Collapsed;
            }
            if (_resultsPage?.RunHealthUnmatchedText != null)
            {
                _resultsPage.RunHealthUnmatchedText.Text = $"Unmatched targets: {unmatched:N0}";
            }
            if (_resultsPage?.RunHealthDetailsButton != null)
            {
                _resultsPage.RunHealthDetailsButton.Visibility = hasIssues ? Visibility.Visible : Visibility.Collapsed;
            }

            var impactText = BuildRunHealthImpactText(stats, missingBounds, unmatched);
            UpdateRunHealthFlyout(
                RunHealthTargetsTotalTextControl,
                RunHealthTargetsWithBoundsTextControl,
                RunHealthTargetsWithoutBoundsTextControl,
                RunHealthTargetsSampledTextControl,
                RunHealthTargetsSampleNoBoundsTextControl,
                RunHealthTargetsSampleNoGeometryTextControl,
                RunHealthTargetsUnmatchedTextControl,
                RunHealthImpactTextControl,
                RunHealthCreateMissingBoundsButtonControl,
                RunHealthCreateUnmatchedButtonControl,
                stats,
                missingBounds,
                unmatched,
                impactText);

            UpdateRunHealthFlyout(
                _resultsPage?.RunHealthTargetsTotalText,
                _resultsPage?.RunHealthTargetsWithBoundsText,
                _resultsPage?.RunHealthTargetsWithoutBoundsText,
                _resultsPage?.RunHealthTargetsSampledText,
                _resultsPage?.RunHealthTargetsSampleNoBoundsText,
                _resultsPage?.RunHealthTargetsSampleNoGeometryText,
                _resultsPage?.RunHealthTargetsUnmatchedText,
                _resultsPage?.RunHealthImpactText,
                _resultsPage?.RunHealthCreateMissingBoundsButton,
                _resultsPage?.RunHealthCreateUnmatchedButton,
                stats,
                missingBounds,
                unmatched,
                impactText);
        }

        private static string BuildRunHealthSummaryText(int missingBounds, int unmatched, double missingRatio)
        {
            var parts = new List<string>();
            if (missingBounds > 0)
            {
                parts.Add($"{missingBounds:N0} targets missing bounds (skipped)");
            }
            if (unmatched > 0)
            {
                parts.Add($"{unmatched:N0} unmatched targets");
            }

            if (parts.Count == 0)
            {
                return "Run health: ready";
            }

            var message = "Run health: " + string.Join("; ", parts);
            if (missingRatio >= 0.02)
            {
                message += ". Accuracy reduced for those targets.";
            }

            return message;
        }

        private static string BuildRunHealthImpactText(SpaceMapperRunStats stats, int missingBounds, int unmatched)
        {
            var impact = new List<string>();
            if (missingBounds > 0)
            {
                impact.Add("Targets without bounds were skipped from containment checks.");
                if (stats != null)
                {
                    if (stats.TargetsSampleSkippedNoBounds > 0 || stats.TargetsSampleSkippedNoGeometry > 0)
                    {
                        impact.Add(
                            $"Sampling applied to {stats.TargetsSampled:N0} targets; skipped for {stats.TargetsSampleSkippedNoBounds:N0} missing bounds and {stats.TargetsSampleSkippedNoGeometry:N0} missing geometry.");
                    }
                }
            }
            if (unmatched > 0)
            {
                impact.Add("Unmatched targets had no zone match.");
            }

            return string.Join(" ", impact);
        }

        private static void UpdateRunHealthFlyout(
            TextBlock targetsTotalText,
            TextBlock targetsWithBoundsText,
            TextBlock targetsWithoutBoundsText,
            TextBlock targetsSampledText,
            TextBlock targetsSampleNoBoundsText,
            TextBlock targetsSampleNoGeometryText,
            TextBlock targetsUnmatchedText,
            TextBlock impactText,
            Wpf.Ui.Controls.Button createMissingBoundsButton,
            Wpf.Ui.Controls.Button createUnmatchedButton,
            SpaceMapperRunStats stats,
            int missingBounds,
            int unmatched,
            string impact)
        {
            if (targetsTotalText != null)
            {
                targetsTotalText.Text = FormatCount(stats?.TargetsTotal);
            }
            if (targetsWithBoundsText != null)
            {
                targetsWithBoundsText.Text = FormatCount(stats?.TargetsWithBounds);
            }
            if (targetsWithoutBoundsText != null)
            {
                targetsWithoutBoundsText.Text = FormatCount(stats?.TargetsWithoutBounds ?? missingBounds);
            }
            if (targetsSampledText != null)
            {
                targetsSampledText.Text = FormatCount(stats?.TargetsSampled);
            }
            if (targetsSampleNoBoundsText != null)
            {
                targetsSampleNoBoundsText.Text = FormatCount(stats?.TargetsSampleSkippedNoBounds);
            }
            if (targetsSampleNoGeometryText != null)
            {
                targetsSampleNoGeometryText.Text = FormatCount(stats?.TargetsSampleSkippedNoGeometry);
            }
            if (targetsUnmatchedText != null)
            {
                targetsUnmatchedText.Text = FormatCount(unmatched);
            }
            if (impactText != null)
            {
                impactText.Text = string.IsNullOrWhiteSpace(impact) ? string.Empty : impact;
                impactText.Visibility = string.IsNullOrWhiteSpace(impact) ? Visibility.Collapsed : Visibility.Visible;
            }
            if (createMissingBoundsButton != null)
            {
                createMissingBoundsButton.Visibility = missingBounds > 0 ? Visibility.Visible : Visibility.Collapsed;
            }
            if (createUnmatchedButton != null)
            {
                createUnmatchedButton.Visibility = unmatched > 0 ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private static string FormatCount(int? value)
        {
            return value.HasValue ? value.Value.ToString("N0") : NotApplicableText;
        }

        private static void ToggleRunHealthFlyout(WpfFlyout flyout)
        {
            if (flyout == null)
            {
                return;
            }

            if (flyout.IsOpen)
            {
                flyout.Hide();
            }
            else
            {
                flyout.Show();
            }
        }

        private void CreateDiagnosticsSelectionSet(string label, IReadOnlyCollection<ModelItem> items)
        {
            if (items == null || items.Count == 0)
            {
                MessageBox.Show("No items available for selection set creation.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("No active document found.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var collection = new ModelItemCollection();
            foreach (var item in items)
            {
                if (item != null)
                {
                    collection.Add(item);
                }
            }

            if (collection.Count == 0)
            {
                MessageBox.Show("No valid items available for selection set creation.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var folder = EnsureSelectionSetFolder(doc, "Space Mapper Diagnostics");
            var set = new SelectionSet(collection)
            {
                DisplayName = MakeUniqueName(folder, label)
            };
            doc.SelectionSets.AddCopy(folder, set);
            MessageBox.Show($"Selection set created: {set.DisplayName}", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private static Autodesk.Navisworks.Api.GroupItem EnsureSelectionSetFolder(Document doc, string folderPath)
        {
            var parts = (folderPath ?? string.Empty)
                .Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);

            Autodesk.Navisworks.Api.GroupItem current = doc.SelectionSets.RootItem;
            foreach (var part in parts)
            {
                var existing = current.Children
                    .OfType<Autodesk.Navisworks.Api.GroupItem>()
                    .FirstOrDefault(x => string.Equals(x.DisplayName, part, StringComparison.OrdinalIgnoreCase));

                if (existing != null)
                {
                    current = existing;
                    continue;
                }

                var folder = new FolderItem { DisplayName = part };
                doc.SelectionSets.AddCopy(current, folder);
                current = current.Children
                    .OfType<Autodesk.Navisworks.Api.GroupItem>()
                    .FirstOrDefault(x => string.Equals(x.DisplayName, part, StringComparison.OrdinalIgnoreCase))
                          ?? current;
            }

            return current;
        }

        private static string MakeUniqueName(Autodesk.Navisworks.Api.GroupItem folder, string desired)
        {
            desired = SanitizeName(desired);
            if (string.IsNullOrWhiteSpace(desired))
            {
                desired = "Set";
            }

            var names = new HashSet<string>(
                folder.Children.Select(x => x.DisplayName),
                StringComparer.OrdinalIgnoreCase);

            if (!names.Contains(desired))
            {
                return desired;
            }

            for (var i = 2; i < 9999; i++)
            {
                var candidate = $"{desired} ({i})";
                if (!names.Contains(candidate))
                {
                    return candidate;
                }
            }

            return desired + " (9999)";
        }

        private static string SanitizeName(string value)
        {
            if (value == null)
            {
                return "Set";
            }

            foreach (var c in Path.GetInvalidFileNameChars())
            {
                value = value.Replace(c, '_');
            }

            return value.Trim();
        }

        private void SaveTemplate()
        {
            SaveTemplateAs();
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

            RefreshTemplateList(tpl.Name);
            ApplyTemplate(tpl);
        }

        private void ApplySettings(SpaceMapperProcessingSettings settings)
        {
            if (settings == null) return;
            _processingPage.TreatPartialCheck.IsChecked = settings.TreatPartialAsContained;
            _processingPage.TagPartialCheck.IsChecked = settings.TagPartialSeparately;
            var writeBehavior = settings.WriteZoneBehaviorProperty;
            if (!writeBehavior && settings.TagPartialSeparately)
            {
                writeBehavior = true;
            }
            if (_processingPage.WriteZoneBehaviorCheck != null)
            {
                _processingPage.WriteZoneBehaviorCheck.IsChecked = writeBehavior;
            }
            if (_processingPage.WriteContainmentPercentCheck != null)
            {
                _processingPage.WriteContainmentPercentCheck.IsChecked = settings.WriteZoneContainmentPercentProperty;
            }
            if (_processingPage.ContainmentCalculationCombo != null)
            {
                _processingPage.ContainmentCalculationCombo.SelectedIndex = settings.ContainmentCalculationMode switch
                {
                    SpaceMapperContainmentCalculationMode.SamplePoints => 1,
                    SpaceMapperContainmentCalculationMode.SamplePointsDense => 2,
                    SpaceMapperContainmentCalculationMode.TargetGeometry => 3,
                    SpaceMapperContainmentCalculationMode.TargetGeometryGpu => 4,
                    SpaceMapperContainmentCalculationMode.BoundsOverlap => 5,
                    _ => 0
                };
            }
            if (_processingPage.GpuRayAccuracyCombo != null)
            {
                _processingPage.GpuRayAccuracyCombo.SelectedIndex = settings.GpuRayCount >= 2 ? 1 : 0;
            }
            _processingPage.EnableMultiZoneCheck.IsChecked = settings.EnableMultipleZones;
            if (_processingPage.ExcludeZonesFromTargetsCheck != null)
            {
                _processingPage.ExcludeZonesFromTargetsCheck.IsChecked = settings.ExcludeZonesFromTargets;
            }
            if (_processingPage.EnableZoneOffsetsCheck != null)
            {
                _processingPage.EnableZoneOffsetsCheck.IsChecked = settings.EnableZoneOffsets;
            }
            if (_processingPage.EnableOffsetAreaPassCheck != null)
            {
                _processingPage.EnableOffsetAreaPassCheck.IsChecked = settings.EnableOffsetAreaPass;
            }
            if (_processingPage.WriteOffsetMatchPropertyCheck != null)
            {
                _processingPage.WriteOffsetMatchPropertyCheck.IsChecked = settings.WriteZoneOffsetMatchProperty;
            }
            _processingPage.Offset3DBox.Text = settings.Offset3D.ToString();
            _processingPage.OffsetTopBox.Text = settings.OffsetTop.ToString();
            _processingPage.OffsetBottomBox.Text = settings.OffsetBottom.ToString();
            _processingPage.OffsetSidesBox.Text = settings.OffsetSides.ToString();
            _processingPage.UnitsBox.Text = settings.Units;
            _processingPage.OffsetModeBox.Text = settings.OffsetMode;
            _processingPage.ZoneBehaviorCategoryBox.Text = settings.ZoneBehaviorCategory ?? "ME_SpaceInfo";
            _processingPage.ZoneBehaviorPropertyBox.Text = settings.ZoneBehaviorPropertyName ?? "Zone Behaviour";
            _processingPage.ZoneBehaviorContainedBox.Text = settings.ZoneBehaviorContainedValue ?? "Contained";
            _processingPage.ZoneBehaviorPartialBox.Text = settings.ZoneBehaviorPartialValue ?? "Partial";
            UpdateZoneBehaviorInputs();
            if (_processingPage.ShowInternalWritebackCheck != null)
            {
                _processingPage.ShowInternalWritebackCheck.IsChecked = settings.ShowInternalPropertiesDuringWriteback;
            }
            if (_processingPage.SkipUnchangedWritebackCheck != null)
            {
                _processingPage.SkipUnchangedWritebackCheck.IsChecked = settings.SkipUnchangedWriteback;
            }
            if (_processingPage.PackWritebackCheck != null)
            {
                _processingPage.PackWritebackCheck.IsChecked = settings.PackWritebackProperties;
            }
            if (_processingPage.CloseDockPanesCheck != null)
            {
                _processingPage.CloseDockPanesCheck.IsChecked = settings.CloseDockPanesDuringRun;
            }
            if (_processingPage.DockPaneDelayBox != null)
            {
                _processingPage.DockPaneDelayBox.Text = settings.DockPaneCloseDelaySeconds.ToString();
            }
            ApplyBoundsSettings(settings);
            _processingPage.UpdateProcessingUiState();
        }

        private void ApplyBoundsSettings(SpaceMapperProcessingSettings settings)
        {
            if (settings == null || _processingPage == null)
            {
                return;
            }

            ApplyLegacyBoundsFromPreset(settings);

            _processingPage.ZoneBoundsSlider.Value = settings.ZoneBoundsMode switch
            {
                SpaceMapperZoneBoundsMode.Obb => 1,
                SpaceMapperZoneBoundsMode.KDop => 2,
                SpaceMapperZoneBoundsMode.Hull => 3,
                _ => 0
            };

            _processingPage.TargetBoundsSlider.Value = settings.TargetBoundsMode switch
            {
                SpaceMapperTargetBoundsMode.Midpoint => 0,
                SpaceMapperTargetBoundsMode.Aabb => 1,
                SpaceMapperTargetBoundsMode.Obb => 2,
                SpaceMapperTargetBoundsMode.KDop => 3,
                SpaceMapperTargetBoundsMode.Hull => 4,
                _ => 1
            };

            _processingPage.ZoneKDopVariantCombo.SelectedIndex = settings.ZoneKDopVariant switch
            {
                SpaceMapperKDopVariant.KDop8 => 0,
                SpaceMapperKDopVariant.KDop18 => 2,
                _ => 1
            };

            _processingPage.TargetKDopVariantCombo.SelectedIndex = settings.TargetKDopVariant switch
            {
                SpaceMapperKDopVariant.KDop8 => 0,
                SpaceMapperKDopVariant.KDop18 => 2,
                _ => 1
            };

            _processingPage.TargetMidpointModeCombo.SelectedIndex = settings.TargetMidpointMode switch
            {
                SpaceMapperMidpointMode.BoundingBoxBottomCenter => 1,
                _ => 0
            };

            if (_processingPage.ZoneContainmentEngineCombo != null)
            {
                _processingPage.ZoneContainmentEngineCombo.SelectedIndex = settings.ZoneContainmentEngine switch
                {
                    SpaceMapperZoneContainmentEngine.MeshAccurate => 1,
                    _ => 0
                };
            }

            if (_processingPage.ZoneResolutionStrategyCombo != null)
            {
                _processingPage.ZoneResolutionStrategyCombo.SelectedIndex = settings.ZoneResolutionStrategy switch
                {
                    SpaceMapperZoneResolutionStrategy.LargestOverlap => 1,
                    SpaceMapperZoneResolutionStrategy.FirstMatch => 2,
                    _ => 0
                };
            }

            OnBoundsModesChanged();
        }

        private static void ApplyLegacyBoundsFromPreset(SpaceMapperProcessingSettings settings)
        {
            if (settings == null)
            {
                return;
            }

            var boundsExplicit = settings.ZoneBoundsMode != SpaceMapperZoneBoundsMode.Aabb
                || settings.TargetBoundsMode != SpaceMapperTargetBoundsMode.Aabb
                || settings.ZoneKDopVariant != SpaceMapperKDopVariant.KDop14
                || settings.TargetKDopVariant != SpaceMapperKDopVariant.KDop14
                || settings.TargetMidpointMode != SpaceMapperMidpointMode.BoundingBoxCenter;

            if (boundsExplicit)
            {
                return;
            }

            switch (settings.PerformancePreset)
            {
                case SpaceMapperPerformancePreset.Fast:
                    settings.ZoneBoundsMode = SpaceMapperZoneBoundsMode.Aabb;
                    settings.TargetBoundsMode = SpaceMapperTargetBoundsMode.Midpoint;
                    settings.TargetMidpointMode = SpaceMapperMidpointMode.BoundingBoxCenter;
                    break;
                case SpaceMapperPerformancePreset.Accurate:
                    settings.ZoneBoundsMode = SpaceMapperZoneBoundsMode.Hull;
                    settings.TargetBoundsMode = SpaceMapperTargetBoundsMode.Hull;
                    break;
                default:
                    settings.ZoneBoundsMode = SpaceMapperZoneBoundsMode.Aabb;
                    settings.TargetBoundsMode = SpaceMapperTargetBoundsMode.Aabb;
                    break;
            }
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

        private static string BuildVariationReportPath(Document doc, string templateName)
        {
            var baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "Autodesk",
                "Navisworks Manage 2025",
                "Plugins",
                "MicroEng.Navisworks",
                "Reports",
                "SpaceMapper");

            var modelName = GetDocumentTitle(doc);
            var template = string.IsNullOrWhiteSpace(templateName) ? null : templateName.Trim();
            var safeModel = SanitizeFileName(modelName);
            var safeTemplate = SanitizeFileName(template);
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            var suffixParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(safeModel)) suffixParts.Add(safeModel);
            if (!string.IsNullOrWhiteSpace(safeTemplate)) suffixParts.Add(safeTemplate);
            var suffix = suffixParts.Count == 0 ? string.Empty : "_" + string.Join("_", suffixParts);

            var fileName = $"SpaceMapper_VariationCheck_{stamp}{suffix}.md";
            return Path.Combine(baseDir, fileName);
        }

        private static string GetDocumentTitle(Document doc)
        {
            if (doc == null)
            {
                return "Untitled";
            }

            var title = doc.Title;
            if (!string.IsNullOrWhiteSpace(title))
            {
                return title.Trim();
            }

            if (!string.IsNullOrWhiteSpace(doc.FileName))
            {
                return Path.GetFileNameWithoutExtension(doc.FileName);
            }

            return "Untitled";
        }

        private static string SanitizeFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var invalid = Path.GetInvalidFileNameChars();
            var cleaned = new string(value.Where(ch => !invalid.Contains(ch)).ToArray());
            return cleaned.Length > 60 ? cleaned.Substring(0, 60) : cleaned;
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
