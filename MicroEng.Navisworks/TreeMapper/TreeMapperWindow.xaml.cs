using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.TreeMapper
{
    internal partial class TreeMapperWindow : Window, INotifyPropertyChanged
    {
        private readonly TreeMapperProfileStore _store = new TreeMapperProfileStore();
        private readonly TreeMapperPublishStore _publishStore = new TreeMapperPublishStore();
        private TreeMapperProfileStoreData _storeData;
        private TreeMapperProfile _activeProfile;
        private TreeMapperProfile _selectedProfile;
        private readonly ObservableCollection<string> _categoryOptions = new ObservableCollection<string>();
        private readonly ObservableCollection<ScrapeSessionOption> _sessionOptions = new ObservableCollection<ScrapeSessionOption>();
        private readonly ObservableCollection<TreeMapperProfile> _profiles = new ObservableCollection<TreeMapperProfile>();
        private readonly Dictionary<string, List<string>> _propertyOptions = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        private readonly DispatcherTimer _previewDebounce;
        private CancellationTokenSource _previewCts;
        private bool _isLoading;
        private bool _isScrapeRunning;
        private bool _isPublishEnabled;
        private bool _isDirty;
        private string _profileName;
        private string _cacheStatusText;
        private string _scrapeButtonText = "Run Data Scraper (Entire Model)";
        private bool _isScrapeEnabled = true;
        private TreeMapperLevelViewModel _selectedLevel;
        private ScrapeSessionOption _selectedSession;

        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<TreeMapperLevelViewModel> Levels { get; } = new ObservableCollection<TreeMapperLevelViewModel>();
        public ObservableCollection<TreeMapperPreviewNode> PreviewRoots { get; } = new ObservableCollection<TreeMapperPreviewNode>();
        public ObservableCollection<string> CategoryOptions => _categoryOptions;
        public ObservableCollection<ScrapeSessionOption> SessionOptions => _sessionOptions;
        public ObservableCollection<TreeMapperProfile> Profiles => _profiles;

        public TreeMapperLevelViewModel SelectedLevel
        {
            get => _selectedLevel;
            set => SetField(ref _selectedLevel, value);
        }

        public TreeMapperProfile SelectedProfile
        {
            get => _selectedProfile;
            set
            {
                if (SetField(ref _selectedProfile, value))
                {
                    if (!_isLoading && value != null)
                    {
                        if (!ConfirmProfileSwitch())
                        {
                            _isLoading = true;
                            _selectedProfile = _activeProfile;
                            OnPropertyChanged(nameof(SelectedProfile));
                            _isLoading = false;
                            return;
                        }

                        SwitchProfile(value);
                    }
                }
            }
        }

        public ScrapeSessionOption SelectedSession
        {
            get => _selectedSession;
            set
            {
                if (SetField(ref _selectedSession, value))
                {
                    if (!_isLoading)
                    {
                        var session = _selectedSession?.Session;
                        RefreshCategoryOptions(session);
                        UpdateCacheStatus(session);
                        SchedulePreviewRebuild();
                    }
                }
            }
        }

        public string ProfileName
        {
            get => _profileName;
            set
            {
                if (SetField(ref _profileName, value))
                {
                    if (!_isLoading)
                    {
                        MarkDirty();
                    }
                }
            }
        }

        public string CacheStatusText
        {
            get => _cacheStatusText;
            private set => SetField(ref _cacheStatusText, value);
        }

        public string ScrapeButtonText
        {
            get => _scrapeButtonText;
            private set => SetField(ref _scrapeButtonText, value);
        }

        public bool IsScrapeEnabled
        {
            get => _isScrapeEnabled;
            private set => SetField(ref _isScrapeEnabled, value);
        }

        public bool IsPublishEnabled
        {
            get => _isPublishEnabled;
            private set => SetField(ref _isPublishEnabled, value);
        }

        public TreeMapperWindow()
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(this);
            DataContext = this;

            _previewDebounce = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(250)
            };
            _previewDebounce.Tick += (_, __) =>
            {
                _previewDebounce.Stop();
                _ = RebuildPreviewAsync();
            };

            Levels.CollectionChanged += OnLevelsChanged;
            DataScraperCache.SessionAdded += OnSessionAdded;
            Closed += (_, __) => DataScraperCache.SessionAdded -= OnSessionAdded;

            LoadProfiles();
            RefreshSessionOptions(DataScraperCache.LastSession);
            UpdatePublishEnabled();
            SchedulePreviewRebuild();
        }

        private void OnSessionAdded(ScrapeSession session)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => OnSessionAdded(session)));
                return;
            }

            RefreshSessionOptions(session);
            SchedulePreviewRebuild();
        }

        private void OnLevelsChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (_isLoading)
            {
                return;
            }

            SchedulePreviewRebuild();
            MarkDirty();
            UpdatePublishEnabled();
        }

        private void LoadProfiles()
        {
            _storeData = _store.Load();
            if (_storeData.Profiles == null)
            {
                _storeData.Profiles = new List<TreeMapperProfile>();
            }

            _activeProfile = null;
            if (!string.IsNullOrWhiteSpace(_storeData.ActiveProfileId))
            {
                _activeProfile = _storeData.Profiles.FirstOrDefault(p => p.Id == _storeData.ActiveProfileId);
            }

            if (_activeProfile == null && _storeData.Profiles.Count > 0)
            {
                _activeProfile = _storeData.Profiles[0];
            }

            if (_activeProfile == null)
            {
                _activeProfile = new TreeMapperProfile();
                _storeData.Profiles.Add(_activeProfile);
            }

            if (_activeProfile.Levels == null)
            {
                _activeProfile.Levels = new List<TreeMapperLevel>();
            }

            if (_activeProfile.Levels.Count == 0)
            {
                _activeProfile.Levels.Add(new TreeMapperLevel());
            }

            _storeData.ActiveProfileId = _activeProfile.Id;

            RefreshProfileList();
            ApplyProfileToUi(_activeProfile);
        }

        private TreeMapperLevelViewModel CreateLevelViewModel(TreeMapperLevel model)
        {
            return TreeMapperLevelViewModel.FromModel(
                model ?? new TreeMapperLevel(),
                _categoryOptions,
                ResolveProperties,
                OnLevelChanged);
        }

        private void OnLevelChanged()
        {
            if (_isLoading)
            {
                return;
            }

            SchedulePreviewRebuild();
            MarkDirty();
            UpdatePublishEnabled();
        }

        private void RefreshSessionOptions(ScrapeSession preferredSession)
        {
            _isLoading = true;
            _sessionOptions.Clear();
            var sessions = DataScraperCache.AllSessions
                .OrderByDescending(s => s.Timestamp)
                .ToList();
            foreach (var session in sessions)
            {
                _sessionOptions.Add(new ScrapeSessionOption(session));
            }

            var target = preferredSession ?? DataScraperCache.LastSession;
            SelectedSession = _sessionOptions.FirstOrDefault(o => o.Session == target) ?? _sessionOptions.FirstOrDefault();
            _isLoading = false;
            var selected = SelectedSession?.Session;
            RefreshCategoryOptions(selected);
            UpdateCacheStatus(selected);
        }

        private void RefreshCategoryOptions(ScrapeSession session)
        {
            _categoryOptions.Clear();
            _propertyOptions.Clear();

            var properties = session?.Properties;
            if (properties != null && properties.Count > 0)
            {
                foreach (var group in properties
                             .Where(p => !string.IsNullOrWhiteSpace(p.Category) && !string.IsNullOrWhiteSpace(p.Name))
                             .GroupBy(p => p.Category, StringComparer.OrdinalIgnoreCase)
                             .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
                {
                    _categoryOptions.Add(group.Key);
                    var props = group.Select(p => p.Name)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    _propertyOptions[group.Key] = props;
                }
            }
            else if (session?.RawEntries != null && session.RawEntries.Count > 0)
            {
                foreach (var group in session.RawEntries
                             .Where(p => !string.IsNullOrWhiteSpace(p.Category) && !string.IsNullOrWhiteSpace(p.Name))
                             .GroupBy(p => p.Category, StringComparer.OrdinalIgnoreCase)
                             .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
                {
                    _categoryOptions.Add(group.Key);
                    var props = group.Select(p => p.Name)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    _propertyOptions[group.Key] = props;
                }
            }

            EnsureCategoriesForLevels();

            _isLoading = true;
            foreach (var level in Levels)
            {
                level.RefreshPropertyOptions();
            }
            _isLoading = false;
            UpdatePublishEnabled();
        }

        private void RefreshProfileList()
        {
            _profiles.Clear();
            foreach (var profile in _storeData.Profiles)
            {
                _profiles.Add(profile);
            }

            _isLoading = true;
            SelectedProfile = _profiles.FirstOrDefault(p => p.Id == _storeData.ActiveProfileId) ?? _profiles.FirstOrDefault();
            _isLoading = false;
        }

        private void SwitchProfile(TreeMapperProfile profile)
        {
            if (profile == null || ReferenceEquals(profile, _activeProfile))
            {
                return;
            }

            _activeProfile = profile;
            _storeData.ActiveProfileId = profile.Id;
            ApplyProfileToUi(profile);
            SchedulePreviewRebuild();
        }

        private void ApplyProfileToUi(TreeMapperProfile profile)
        {
            if (profile == null)
            {
                return;
            }

            if (profile.Levels == null)
            {
                profile.Levels = new List<TreeMapperLevel>();
            }

            if (profile.Levels.Count == 0)
            {
                profile.Levels.Add(new TreeMapperLevel());
            }

            _isLoading = true;
            ProfileName = profile.Name;
            Levels.Clear();
            foreach (var level in profile.Levels)
            {
                Levels.Add(CreateLevelViewModel(level));
            }
            SelectedLevel = Levels.FirstOrDefault();
            EnsureCategoriesForLevels();
            _isLoading = false;
            _isDirty = false;
            UpdatePublishEnabled();
        }

        private IEnumerable<string> ResolveProperties(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
            {
                return Enumerable.Empty<string>();
            }

            return _propertyOptions.TryGetValue(category, out var props) ? props : Enumerable.Empty<string>();
        }

        private void EnsureCategoriesForLevels()
        {
            foreach (var level in Levels)
            {
                var category = level.Category;
                if (string.IsNullOrWhiteSpace(category))
                {
                    continue;
                }

                if (!_categoryOptions.Any(c => string.Equals(c, category, StringComparison.OrdinalIgnoreCase)))
                {
                    _categoryOptions.Add(category);
                }

                if (!_propertyOptions.TryGetValue(category, out var props))
                {
                    props = new List<string>();
                    _propertyOptions[category] = props;
                }

                var propertyName = level.PropertyName;
                if (!string.IsNullOrWhiteSpace(propertyName) &&
                    !props.Any(p => string.Equals(p, propertyName, StringComparison.OrdinalIgnoreCase)))
                {
                    props.Add(propertyName);
                }
            }
        }

        private void UpdateCacheStatus(ScrapeSession session)
        {
            if (session == null)
            {
                CacheStatusText = "No Data Scraper cache found. Run an entire model scrape to populate the Tree Mapper.";
                ScrapeButtonText = "Run Data Scraper (Entire Model)";
                UpdatePublishEnabled();
                return;
            }

            CacheStatusText = $"Last scrape: {session.Timestamp:G} | Items scanned: {session.ItemsScanned} | Profile: {session.ProfileName}";
            ScrapeButtonText = "Refresh Data Scraper Cache";
            UpdatePublishEnabled();
        }

        private void UpdatePublishEnabled()
        {
            var session = SelectedSession?.Session ?? DataScraperCache.LastSession;
            var hasSession = session != null;
            var hasLevels = Levels.Any(l =>
                !string.IsNullOrWhiteSpace(l.Category) && !string.IsNullOrWhiteSpace(l.PropertyName));
            IsPublishEnabled = hasSession && hasLevels;
        }

        private void SchedulePreviewRebuild()
        {
            if (_isLoading)
            {
                return;
            }

            _previewDebounce.Stop();
            _previewDebounce.Start();
        }

        private async Task RebuildPreviewAsync()
        {
            var session = SelectedSession?.Session ?? DataScraperCache.LastSession;
            var levels = Levels.Select(l => l.ToModel()).ToList();

            _previewCts?.Cancel();
            _previewCts = new CancellationTokenSource();
            var token = _previewCts.Token;

            List<TreeMapperPreviewNode> preview;
            try
            {
                preview = await Task.Run(() => TreeMapperEngine.BuildPreview(session, levels, token), token);
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper preview failed: {ex.Message}");
                return;
            }

            if (token.IsCancellationRequested)
            {
                return;
            }

            PreviewRoots.Clear();
            foreach (var node in preview)
            {
                PreviewRoots.Add(node);
            }
        }

        private void SaveNow()
        {
            if (_activeProfile == null || _storeData == null)
            {
                return;
            }

            CommitLevelEdits();
            var name = string.IsNullOrWhiteSpace(ProfileName) ? "TreeMapper" : ProfileName.Trim();
            _activeProfile.Name = name;
            _activeProfile.Levels = Levels.Select(l => l.ToModel()).ToList();
            _activeProfile.UpdatedUtc = DateTime.UtcNow;
            _storeData.ActiveProfileId = _activeProfile.Id;

            try
            {
                _store.Save(_storeData);
                var index = _profiles.IndexOf(_activeProfile);
                if (index >= 0)
                {
                    _profiles[index] = _activeProfile;
                }

                _isDirty = false;
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper profile save failed: {ex.Message}");
            }
        }

        private void AddLevel_Click(object sender, RoutedEventArgs e)
        {
            var vm = CreateLevelViewModel(new TreeMapperLevel());
            Levels.Add(vm);
            SelectedLevel = vm;
        }

        private void RemoveLevel_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedLevel == null)
            {
                return;
            }

            var index = Levels.IndexOf(SelectedLevel);
            if (index < 0)
            {
                return;
            }

            Levels.RemoveAt(index);
            SelectedLevel = Levels.Count > 0 ? Levels[Math.Min(index, Levels.Count - 1)] : null;
        }

        private void MoveLevelUp_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedLevel == null)
            {
                return;
            }

            var index = Levels.IndexOf(SelectedLevel);
            if (index <= 0)
            {
                return;
            }

            Levels.Move(index, index - 1);
        }

        private void MoveLevelDown_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedLevel == null)
            {
                return;
            }

            var index = Levels.IndexOf(SelectedLevel);
            if (index < 0 || index >= Levels.Count - 1)
            {
                return;
            }

            Levels.Move(index, index + 1);
        }

        private void RunScrape_Click(object sender, RoutedEventArgs e)
        {
            if (_isScrapeRunning)
            {
                return;
            }

            _isScrapeRunning = true;
            IsScrapeEnabled = false;
            var previousButtonText = ScrapeButtonText;
            ScrapeButtonText = "Scraping...";

            try
            {
                MicroEngActions.Log("TreeMapper: running Data Scraper (entire model)");
                var service = new DataScraperService();
                var items = service.ResolveScope(ScrapeScopeType.EntireModel, null, null, out var description);
            var session = service.Scrape("TreeMapper", ScrapeScopeType.EntireModel, description, items);
            RefreshCategoryOptions(session);
            RefreshSessionOptions(session);
            SchedulePreviewRebuild();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper scrape failed: {ex}");
                MessageBox.Show($"Tree Mapper failed to run Data Scraper: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _isScrapeRunning = false;
                IsScrapeEnabled = true;
                if (ScrapeButtonText == "Scraping...")
                {
                    ScrapeButtonText = previousButtonText;
                }
            }
        }

        private void RefreshSessions_Click(object sender, RoutedEventArgs e)
        {
            RefreshSessionOptions(SelectedSession?.Session ?? DataScraperCache.LastSession);
        }

        private void OpenDataScraper_Click(object sender, RoutedEventArgs e)
        {
            MicroEngActions.TryShowDataScraper(null, out _);
        }

        private void OpenDataMatrix_Click(object sender, RoutedEventArgs e)
        {
            MicroEngActions.DataMatrix();
        }

        private void SaveProfile_Click(object sender, RoutedEventArgs e)
        {
            CommitLevelEdits();
            SaveNow();
        }

        private void SaveProfileAs_Click(object sender, RoutedEventArgs e)
        {
            CommitLevelEdits();
            var name = PromptText("Save profile as:", "Tree Mapper");
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            var trimmed = name.Trim();
            var existing = _storeData.Profiles.FirstOrDefault(p =>
                string.Equals(p.Name, trimmed, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                var result = MessageBox.Show(
                    "A profile with that name already exists. Overwrite it?",
                    "Tree Mapper",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (result != MessageBoxResult.Yes)
                {
                    return;
                }

                existing.Name = trimmed;
                existing.Levels = Levels.Select(l => l.ToModel()).ToList();
                existing.UpdatedUtc = DateTime.UtcNow;
                _activeProfile = existing;
            }
            else
            {
                var profile = new TreeMapperProfile
                {
                    Id = Guid.NewGuid().ToString(),
                    Name = trimmed,
                    Levels = Levels.Select(l => l.ToModel()).ToList(),
                    UpdatedUtc = DateTime.UtcNow
                };
                _storeData.Profiles.Add(profile);
                _activeProfile = profile;
            }

            _storeData.ActiveProfileId = _activeProfile.Id;
            SaveNow();
            RefreshProfileList();
            SelectedProfile = _profiles.FirstOrDefault(p => p.Id == _activeProfile.Id) ?? _profiles.FirstOrDefault();
        }

        private void ReloadProfile_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedProfile == null)
            {
                return;
            }

            var latest = _store.Load();
            var profile = latest.Profiles?.FirstOrDefault(p => p.Id == _selectedProfile.Id) ??
                          latest.Profiles?.FirstOrDefault(p => string.Equals(p.Name, _selectedProfile.Name, StringComparison.OrdinalIgnoreCase));
            if (profile == null)
            {
                MessageBox.Show("Saved profile not found on disk. Save it first to reload.",
                    "Tree Mapper", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _storeData = latest;
            _activeProfile = profile;
            _storeData.ActiveProfileId = profile.Id;
            RefreshProfileList();
            ApplyProfileToUi(profile);
            _isDirty = false;
        }

        private void Publish_Click(object sender, RoutedEventArgs e)
        {
            CommitLevelEdits();
            SaveNow();

            var session = SelectedSession?.Session ?? DataScraperCache.LastSession;
            if (session == null)
            {
                MessageBox.Show("No Data Scraper cache found. Run an entire model scrape first.",
                    "Tree Mapper", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var levels = Levels.Select(l => l.ToModel()).ToList();
            if (levels.All(l => string.IsNullOrWhiteSpace(l.Category) || string.IsNullOrWhiteSpace(l.PropertyName)))
            {
                MessageBox.Show("Add at least one level with Category and Property before publishing.",
                    "Tree Mapper", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                var profileId = _activeProfile?.Id ?? Guid.NewGuid().ToString();
                var profileName = string.IsNullOrWhiteSpace(ProfileName) ? "TreeMapper" : ProfileName.Trim();
                var snapshot = TreeMapperEngine.BuildPublishedTree(session, levels, profileId, profileName, CancellationToken.None);
                EnsureSnapshotDocumentKey(snapshot);
                _publishStore.SaveActiveProfile(_activeProfile ?? new TreeMapperProfile { Id = profileId, Name = profileName, Levels = levels });
                _publishStore.SavePublishedTree(snapshot);

                MicroEngActions.Log($"TreeMapper published: nodes={snapshot.Nodes.Count} file={_publishStore.PublishedTreePath}");
                NavisworksDockPaneManager.RefreshSelectionTreePane();
                MessageBox.Show("Published Tree Mapper to Selection Tree. Restart Navisworks if the dropdown does not refresh.",
                    "Tree Mapper", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"TreeMapper publish failed: {ex}");
                MessageBox.Show($"Tree Mapper publish failed: {ex.Message}", "Tree Mapper",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static void EnsureSnapshotDocumentKey(TreeMapperPublishedTree snapshot)
        {
            if (snapshot == null)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(snapshot.DocumentFileKey))
            {
                return;
            }

            try
            {
                var docFile = Autodesk.Navisworks.Api.Application.ActiveDocument?.FileName ?? string.Empty;
                if (string.IsNullOrWhiteSpace(docFile))
                {
                    return;
                }

                snapshot.DocumentFile = docFile;
                snapshot.DocumentFileKey = Path.GetFileName(docFile) ?? string.Empty;
            }
            catch
            {
                // swallow to avoid publish failure on document access issues
            }
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

        private void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
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

            var panel = new System.Windows.Controls.Grid { Margin = new Thickness(12) };
            panel.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            panel.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            panel.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

            var label = new System.Windows.Controls.TextBlock { Text = caption, Margin = new Thickness(0, 0, 0, 6) };
            var box = new System.Windows.Controls.TextBox();
            var buttons = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };
            var ok = new System.Windows.Controls.Button { Content = "OK", Width = 70, Margin = new Thickness(4), IsDefault = true };
            var cancel = new System.Windows.Controls.Button { Content = "Cancel", Width = 70, Margin = new Thickness(4), IsCancel = true };
            ok.Click += (_, __) => window.DialogResult = true;
            cancel.Click += (_, __) => window.DialogResult = false;
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);

            panel.Children.Add(label);
            panel.Children.Add(box);
            panel.Children.Add(buttons);
            System.Windows.Controls.Grid.SetRow(label, 0);
            System.Windows.Controls.Grid.SetRow(box, 1);
            System.Windows.Controls.Grid.SetRow(buttons, 2);

            window.Content = panel;
            return window.ShowDialog() == true ? box.Text : null;
        }

        private void CommitLevelEdits()
        {
            if (LevelsGrid == null)
            {
                return;
            }

            LevelsGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            LevelsGrid.CommitEdit(DataGridEditingUnit.Row, true);
        }

        private void MarkDirty()
        {
            _isDirty = true;
        }

        private bool ConfirmProfileSwitch()
        {
            if (!_isDirty)
            {
                return true;
            }

            var result = MessageBox.Show(
                "You have unsaved changes. Save them before switching profiles?",
                "Tree Mapper",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel)
            {
                return false;
            }

            if (result == MessageBoxResult.Yes)
            {
                SaveNow();
            }

            _isDirty = false;
            return true;
        }

        internal sealed class ScrapeSessionOption
        {
            public ScrapeSession Session { get; }
            public string Label { get; }

            public ScrapeSessionOption(ScrapeSession session)
            {
                Session = session;
                var profile = string.IsNullOrWhiteSpace(session?.ProfileName) ? "Unknown" : session.ProfileName;
                var stamp = session?.Timestamp == default ? string.Empty : session.Timestamp.ToString("yyyy-MM-dd HH:mm");
                Label = string.IsNullOrWhiteSpace(stamp) ? profile : $"{profile} @ {stamp}";
            }
        }
    }
}
