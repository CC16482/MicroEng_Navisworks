using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Navisworks.Api;
using NavisApp = Autodesk.Navisworks.Api.Application;
using System.Windows.Media;
using System.Windows.Interop;
using System.Windows.Media.Animation;
using WpfUiControls = Wpf.Ui.Controls;

namespace MicroEng.Navisworks
{
    public partial class AppendIntegrateDialog : Window
    {
        static AppendIntegrateDialog()
        {
            AssemblyResolver.EnsureRegistered();
        }

        private readonly AppendIntegrateTemplateStore _store;
        private ObservableCollection<AppendIntegrateTemplate> _templates;
        private AppendIntegrateTemplate _currentTemplate;
        private ObservableCollection<AppendIntegrateRow> _rowBinding;
        private List<string> _dataProfiles = new();

        public Action<string> LogAction { get; set; }

        public AppendIntegrateDialog()
        {
            // Force software rendering to avoid GPU/driver issues inside Navisworks host.
            RenderOptions.ProcessRenderMode = RenderMode.SoftwareOnly;

            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(this);
            _store = new AppendIntegrateTemplateStore(AppDomain.CurrentDomain.BaseDirectory);
            _templates = new ObservableCollection<AppendIntegrateTemplate>(_store.Load());
            _currentTemplate = _templates.FirstOrDefault() ?? AppendIntegrateTemplate.CreateDefault("Default");
            _rowBinding = new ObservableCollection<AppendIntegrateRow>(_currentTemplate.Rows ?? new System.Collections.Generic.List<AppendIntegrateRow>());

            RowsGrid.ItemsSource = _rowBinding;
            TemplateCombo.ItemsSource = _templates;
            TemplateCombo.DisplayMemberPath = nameof(AppendIntegrateTemplate.Name);
            TemplateCombo.SelectedItem = _currentTemplate;

            LoadTemplateIntoUi(_currentTemplate);
            ModeColumn.ItemsSource = Enum.GetValues(typeof(AppendValueMode));
            OptionColumn.ItemsSource = Enum.GetValues(typeof(AppendValueOption));
            RefreshDataProfiles();
        }

        private void TemplateCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TemplateCombo.SelectedItem is AppendIntegrateTemplate selected)
            {
                CaptureUiIntoTemplate();
                LoadTemplateIntoUi(selected);
            }
        }

        private void LoadTemplateIntoUi(AppendIntegrateTemplate template)
        {
            _currentTemplate = template;
            TargetTabTextBox.Text = template.TargetTabName;
            _rowBinding = new ObservableCollection<AppendIntegrateRow>(template.Rows ?? new System.Collections.Generic.List<AppendIntegrateRow>());
            RowsGrid.ItemsSource = _rowBinding;

            ApplyItemsRadio.IsChecked = template.ApplyPropertyTo == ApplyPropertyTarget.Items;
            ApplyGroupsRadio.IsChecked = template.ApplyPropertyTo == ApplyPropertyTarget.Groups;
            ApplyBothRadio.IsChecked = template.ApplyPropertyTo == ApplyPropertyTarget.ItemsAndGroups;

            CreateTabCheck.IsChecked = template.CreateTargetTabIfMissing;
            KeepTabsCheck.IsChecked = template.KeepExistingTabs;
            UpdateTabCheck.IsChecked = template.UpdateExistingTargetTab;
            DeleteBlankCheck.IsChecked = template.DeletePropertyIfBlank;
            DeleteTabBlankCheck.IsChecked = template.DeleteTargetTabIfAllBlank;
            ApplySelectionCheck.IsChecked = template.ApplyToSelectionOnly;
            ShowInternalNamesCheck.IsChecked = template.ShowInternalPropertyNames;
        }

        private void CaptureUiIntoTemplate()
        {
            if (_currentTemplate == null) return;
            _currentTemplate.TargetTabName = string.IsNullOrWhiteSpace(TargetTabTextBox.Text) ? "MicroEng" : TargetTabTextBox.Text.Trim();
            _currentTemplate.Rows = _rowBinding.ToList();
            _currentTemplate.ApplyPropertyTo = ApplyItemsRadio.IsChecked == true
                ? ApplyPropertyTarget.Items
                : ApplyGroupsRadio.IsChecked == true
                    ? ApplyPropertyTarget.Groups
                    : ApplyPropertyTarget.ItemsAndGroups;
            _currentTemplate.CreateTargetTabIfMissing = CreateTabCheck.IsChecked == true;
            _currentTemplate.KeepExistingTabs = KeepTabsCheck.IsChecked == true;
            _currentTemplate.UpdateExistingTargetTab = UpdateTabCheck.IsChecked == true;
            _currentTemplate.DeletePropertyIfBlank = DeleteBlankCheck.IsChecked == true;
            _currentTemplate.DeleteTargetTabIfAllBlank = DeleteTabBlankCheck.IsChecked == true;
            _currentTemplate.ApplyToSelectionOnly = ApplySelectionCheck.IsChecked == true;
            _currentTemplate.ShowInternalPropertyNames = ShowInternalNamesCheck.IsChecked == true;
        }

        private void NewTemplate_Click(object sender, RoutedEventArgs e)
        {
            var name = Prompt("New template name:");
            if (string.IsNullOrWhiteSpace(name)) return;
            var template = AppendIntegrateTemplate.CreateDefault(name.Trim());
            _templates.Add(template);
            TemplateCombo.ItemsSource = _templates;
            TemplateCombo.SelectedItem = template;
            LoadTemplateIntoUi(template);
        }

        private void CopyTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTemplate == null) return;
            var name = Prompt("Copy template as:");
            if (string.IsNullOrWhiteSpace(name)) return;
            var clone = new AppendIntegrateTemplate
            {
                Name = name.Trim(),
                TargetTabName = _currentTemplate.TargetTabName,
                ApplyPropertyTo = _currentTemplate.ApplyPropertyTo,
                CreateTargetTabIfMissing = _currentTemplate.CreateTargetTabIfMissing,
                KeepExistingTabs = _currentTemplate.KeepExistingTabs,
                UpdateExistingTargetTab = _currentTemplate.UpdateExistingTargetTab,
                DeletePropertyIfBlank = _currentTemplate.DeletePropertyIfBlank,
                DeleteTargetTabIfAllBlank = _currentTemplate.DeleteTargetTabIfAllBlank,
                ApplyToSelectionOnly = _currentTemplate.ApplyToSelectionOnly,
                ShowInternalPropertyNames = _currentTemplate.ShowInternalPropertyNames,
                Rows = _currentTemplate.Rows.Select(r => new AppendIntegrateRow
                {
                    TargetPropertyName = r.TargetPropertyName,
                    Mode = r.Mode,
                    SourcePropertyPath = r.SourcePropertyPath,
                    SourcePropertyLabel = r.SourcePropertyLabel,
                    StaticOrExpressionValue = r.StaticOrExpressionValue,
                    Option = r.Option,
                    Enabled = r.Enabled
                }).ToList()
            };
            _templates.Add(clone);
            TemplateCombo.SelectedItem = clone;
            LoadTemplateIntoUi(clone);
        }

        private void RenameTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTemplate == null) return;
            var name = Prompt("Rename template to:", _currentTemplate.Name);
            if (string.IsNullOrWhiteSpace(name)) return;
            _currentTemplate.Name = name.Trim();
            TemplateCombo.Items.Refresh();
        }

        private void DeleteTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTemplate == null) return;
            if (MessageBox.Show($"Delete template '{_currentTemplate.Name}'?", "Confirm", MessageBoxButton.YesNo) != MessageBoxResult.Yes) return;
            _templates.Remove(_currentTemplate);
            if (!_templates.Any())
            {
                _templates.Add(AppendIntegrateTemplate.CreateDefault("Default"));
            }
            TemplateCombo.SelectedItem = _templates.First();
            LoadTemplateIntoUi(_templates.First());
        }

        private void SaveTemplate_Click(object sender, RoutedEventArgs e)
        {
            CaptureUiIntoTemplate();
            _store.Save(_templates);
            StatusText.Text = $"Saved '{_currentTemplate?.Name}'.";
            LogAction?.Invoke($"Template '{_currentTemplate?.Name}' saved.");
        }

        private void AddRow_Click(object sender, RoutedEventArgs e) => _rowBinding.Add(new AppendIntegrateRow());

        private void DeleteRow_Click(object sender, RoutedEventArgs e)
        {
            foreach (var row in RowsGrid.SelectedItems.Cast<AppendIntegrateRow>().ToList())
            {
                _rowBinding.Remove(row);
            }
        }

        private void PickSource_Click(object sender, RoutedEventArgs e)
        {
            if (RowsGrid.SelectedItem is not AppendIntegrateRow row) return;
            var doc = NavisApp.ActiveDocument;
            var sampleItem = doc?.CurrentSelection?.SelectedItems?.FirstOrDefault();
            if (sampleItem == null)
            {
                MessageBox.Show("Select an item in Navisworks to pick properties from.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            using var picker = new PropertyPickerDialog(sampleItem, ShowInternalNamesCheck.IsChecked == true);
            if (picker.ShowDialog() == System.Windows.Forms.DialogResult.OK && picker.SelectedProperty != null)
            {
                row.SourcePropertyPath = picker.SelectedProperty.ToPath();
                row.SourcePropertyLabel = picker.SelectedProperty.ToLabel(ShowInternalNamesCheck.IsChecked == true);
                if (string.IsNullOrWhiteSpace(row.TargetPropertyName))
                {
                    row.TargetPropertyName = picker.SelectedProperty.PropertyDisplayName ?? picker.SelectedProperty.PropertyName;
                }
                RowsGrid.Items.Refresh();
            }
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            CaptureUiIntoTemplate();
            if (string.IsNullOrWhiteSpace(_currentTemplate.TargetTabName))
            {
                MessageBox.Show("Target Tab Name is required.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!_currentTemplate.Rows.Any(r => r.Enabled))
            {
                MessageBox.Show("Add at least one enabled row before running.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            StatusText.Text = "Processing...";
            try
            {
                var executor = new AppendIntegrateExecutor(_currentTemplate, LogAction);
                var result = executor.Execute();
                StatusText.Text = result.Message ?? "Completed.";
                MessageBox.Show(result.Message ?? "Completed.", "MicroEng", MessageBoxButton.OK, MessageBoxImage.Information);
                LogAction?.Invoke(result.Message ?? string.Empty);
                ShowSnackbar("Data Mapper complete",
                    result.Message ?? "Completed.",
                    WpfUiControls.ControlAppearance.Success,
                    WpfUiControls.SymbolRegular.CheckmarkCircle24);
                FlashSuccess(sender as System.Windows.Controls.Button);
            }
            catch (Exception ex)
            {
                StatusText.Text = ex.Message;
                MessageBox.Show(ex.Message, "MicroEng", MessageBoxButton.OK, MessageBoxImage.Error);
                ShowSnackbar("Data Mapper failed",
                    ex.Message,
                    WpfUiControls.ControlAppearance.Danger,
                    WpfUiControls.SymbolRegular.ErrorCircle24);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private static string Prompt(string message, string defaultValue = "")
        {
            var dialog = new Window
            {
                Title = "MicroEng",
                Width = 420,
                Height = 160,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.ToolWindow,
                ResizeMode = ResizeMode.NoResize
            };
            MicroEngWpfUiTheme.ApplyTo(dialog);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(dialog);
            var panel = new Grid { Margin = new Thickness(12) };
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var label = new TextBlock { Text = message, Margin = new Thickness(0, 0, 0, 6) };
            var box = new TextBox { Text = defaultValue ?? string.Empty };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
            var ok = new Button { Content = "OK", Width = 70, Margin = new Thickness(4), IsDefault = true };
            var cancel = new Button { Content = "Cancel", Width = 70, Margin = new Thickness(4), IsCancel = true };
            ok.Click += (_, __) => dialog.DialogResult = true;
            cancel.Click += (_, __) => dialog.DialogResult = false;
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            panel.Children.Add(label);
            panel.Children.Add(box);
            panel.Children.Add(buttons);
            Grid.SetRow(label, 0);
            Grid.SetRow(box, 1);
            Grid.SetRow(buttons, 2);
            dialog.Content = panel;
            return dialog.ShowDialog() == true ? box.Text : null;
        }

        private void ShowSnackbar(string title, string message, WpfUiControls.ControlAppearance appearance, WpfUiControls.SymbolRegular icon)
        {
            if (SnackbarPresenter == null)
            {
                return;
            }

            var snackbar = new WpfUiControls.Snackbar(SnackbarPresenter)
            {
                Title = title,
                Content = message,
                Appearance = appearance,
                Icon = new WpfUiControls.SymbolIcon(WpfUiControls.SymbolRegular.PresenceAvailable24)
                {
                    Filled = true,
                    FontSize = 25
                },
                Foreground = System.Windows.Media.Brushes.Black,
                ContentForeground = System.Windows.Media.Brushes.Black,
                Timeout = TimeSpan.FromSeconds(4),
                IsCloseButtonEnabled = false
            };

            snackbar.Show();
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

        private void RefreshDataProfiles()
        {
            _dataProfiles = DataScraperCache.AllSessions
                .Select(s => string.IsNullOrWhiteSpace(s.ProfileName) ? "Default" : s.ProfileName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n)
                .ToList();
            DataProfileCombo.ItemsSource = _dataProfiles;
            if (_dataProfiles.Any())
            {
                DataProfileCombo.SelectedItem = _dataProfiles.First();
                SetLastSessionFromProfile(_dataProfiles.First());
            }
        }

        private void DataProfileCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = DataProfileCombo.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(selected)) return;
            SetLastSessionFromProfile(selected);
        }

        private void SetLastSessionFromProfile(string profileName)
        {
            var latest = DataScraperCache.AllSessions
                .Where(s => string.Equals(string.IsNullOrWhiteSpace(s.ProfileName) ? "Default" : s.ProfileName, profileName, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(s => s.Timestamp)
                .FirstOrDefault();
            if (latest != null)
            {
                DataScraperCache.LastSession = latest;
                StatusText.Text = $"Using Data Scraper profile '{profileName}' (cached {latest.Timestamp:T}).";
            }
        }
    }
}

