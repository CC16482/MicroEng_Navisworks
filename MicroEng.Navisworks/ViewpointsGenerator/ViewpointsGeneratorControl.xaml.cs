using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using Autodesk.Navisworks.Api;
using MicroEng.Navisworks;
using WpfUiControls = Wpf.Ui.Controls;

namespace MicroEng.Navisworks.ViewpointsGenerator
{
    public partial class ViewpointsGeneratorControl : UserControl, INotifyPropertyChanged
    {
        public ViewpointsGeneratorSettings Settings { get; } = new ViewpointsGeneratorSettings();

        public ObservableCollection<ViewpointsSourceMode> SourceModes { get; } =
            new ObservableCollection<ViewpointsSourceMode>((ViewpointsSourceMode[])Enum.GetValues(typeof(ViewpointsSourceMode)));

        public ObservableCollection<ViewDirectionPreset> DirectionPresets { get; } =
            new ObservableCollection<ViewDirectionPreset>((ViewDirectionPreset[])Enum.GetValues(typeof(ViewDirectionPreset)));

        public ObservableCollection<ProjectionMode> ProjectionModes { get; } =
            new ObservableCollection<ProjectionMode>((ProjectionMode[])Enum.GetValues(typeof(ProjectionMode)));

        public ObservableCollection<SelectionSetPickerItem> SelectionSets { get; } = new ObservableCollection<SelectionSetPickerItem>();
        public ObservableCollection<ViewpointPlanItem> Plan { get; } = new ObservableCollection<ViewpointPlanItem>();

        private string _statusText = "Ready.";

        public string StatusText
        {
            get => _statusText;
            set
            {
                if (string.Equals(_statusText, value, StringComparison.Ordinal))
                {
                    return;
                }

                _statusText = value ?? "";
                OnPropertyChanged();
            }
        }

        public bool IsSelectionSetMode => Settings.SourceMode == ViewpointsSourceMode.SelectionSets;

        static ViewpointsGeneratorControl()
        {
            AssemblyResolver.EnsureRegistered();
        }

        public ViewpointsGeneratorControl()
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            DataContext = this;

            Settings.PropertyChanged += Settings_PropertyChanged;
            Loaded += (_, __) => RefreshSets();
            Unloaded += (_, __) => Settings.PropertyChanged -= Settings_PropertyChanged;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string prop = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        private void Settings_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (string.Equals(e.PropertyName, nameof(ViewpointsGeneratorSettings.SourceMode), StringComparison.Ordinal))
            {
                OnPropertyChanged(nameof(IsSelectionSetMode));
            }
        }

        private void RefreshSets_Click(object sender, RoutedEventArgs e) => RefreshSets();

        private void RefreshSets()
        {
            try
            {
                SelectionSets.Clear();

                var doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                if (doc == null)
                {
                    StatusText = "No active document.";
                    return;
                }

                foreach (var s in ViewpointsGeneratorNavisworksService.LoadSelectionSets(doc))
                {
                    SelectionSets.Add(s);
                }

                StatusText = $"Loaded {SelectionSets.Count} sets.";
            }
            catch (Exception ex)
            {
                StatusText = "Failed to load sets: " + ex.Message;
            }
        }

        private void Preview_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Plan.Clear();

                var doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                if (doc == null)
                {
                    StatusText = "No active document.";
                    return;
                }

                var plan = ViewpointsGeneratorNavisworksService.BuildPlan(doc, Settings, SelectionSets);
                foreach (var p in plan)
                {
                    Plan.Add(p);
                }

                StatusText = $"Plan: {Plan.Count(p => p.Enabled)} viewpoints enabled ({Plan.Count} total).";
            }
            catch (Exception ex)
            {
                StatusText = "Preview failed: " + ex.Message;
            }
        }

        private async void Generate_Click(object sender, RoutedEventArgs e)
        {
            var doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            if (doc == null)
            {
                StatusText = "No active document.";
                return;
            }

            var enabledCount = 0;
            foreach (var item in Plan)
            {
                if (item?.Enabled == true)
                {
                    enabledCount++;
                }
            }

            if (enabledCount == 0)
            {
                StatusText = "Nothing enabled in preview plan.";
                return;
            }

            StatusText = "Generating viewpoints...";

            await Task.Yield();

            try
            {
                ViewpointsGeneratorNavisworksService.GenerateSavedViewpoints(
                    doc,
                    Settings,
                    Plan.ToList(),
                    (done, total, name) =>
                    {
                        StatusText = (done >= total)
                            ? "Done."
                            : $"Creating {done + 1}/{total}: {name}";
                    });

                StatusText = "Done. Viewpoints created.";
                ShowSnackbar("Viewpoints generated",
                    $"Created {enabledCount} viewpoint(s).",
                    WpfUiControls.ControlAppearance.Success,
                    WpfUiControls.SymbolRegular.CheckmarkCircle24);
                FlashSuccess(sender as System.Windows.Controls.Button);
            }
            catch (Exception ex)
            {
                StatusText = "Generate failed: " + ex.Message;
                ShowSnackbar("Generate failed",
                    ex.Message,
                    WpfUiControls.ControlAppearance.Danger,
                    WpfUiControls.SymbolRegular.ErrorCircle24);
            }
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
    }
}

