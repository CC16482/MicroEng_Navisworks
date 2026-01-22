using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Autodesk.Navisworks.Api.Plugins;
using NavisApp = Autodesk.Navisworks.Api.Application;

namespace MicroEng.Navisworks
{
    public partial class MicroEngPanelControl : UserControl
    {
        private bool _settingThemeToggle;
        private DispatcherTimer _toolStateTimer;

        static MicroEngPanelControl()
        {
            AssemblyResolver.EnsureRegistered();
            MicroEngActions.Init();
        }

        public MicroEngPanelControl()
        {
            try
            {
                var diagnosticBypass = string.Equals(
                    Environment.GetEnvironmentVariable("MICROENG_DIAGNOSTIC_BYPASS"),
                    "1",
                    StringComparison.OrdinalIgnoreCase);
                MicroEngActions.Log("Panel: ctor entered");
                // Force software rendering to reduce GPU/driver issues in host.
                System.Windows.Media.RenderOptions.ProcessRenderMode = System.Windows.Interop.RenderMode.SoftwareOnly;
                if (diagnosticBypass)
                {
                    MicroEngActions.Log("Panel: diagnostic bypass active, skipping InitializeComponent");
                    Content = new TextBlock
                    {
                        Text = "MicroEng panel placeholder (diagnostic bypass).",
                        Margin = new Thickness(12)
                    };
                }
                else
                {
                    MicroEngActions.Log("Panel: before InitializeComponent");
                    try
                    {
                        InitializeComponent();
                        MicroEngActions.Log("Panel: after InitializeComponent");
                    }
                    catch (Exception initEx)
                    {
                        MicroEngActions.Log($"Panel: InitializeComponent failed: {initEx}");
                        throw;
                    }
                    try
                    {
                        MicroEngActions.Log("Panel: applying theme");
                        MicroEngWpfUiTheme.ApplyTo(this);
                        MicroEngActions.Log("Panel: theme applied");
                    }
                    catch (Exception themeEx)
                    {
                        MicroEngActions.Log($"Panel: theme apply failed: {themeEx.Message}");
                    }

                    try
                    {
                        InitThemeToggle();
                    }
                    catch (Exception toggleEx)
                    {
                        MicroEngActions.Log($"Panel: theme toggle init failed: {toggleEx.Message}");
                    }
                    if (!DesignerProperties.GetIsInDesignMode(this))
                    {
                        MicroEngActions.LogMessage += LogToPanel;
                        MicroEngActions.Log("Panel: LogMessage handler attached");
                        MicroEngActions.ToolWindowStateChanged += OnToolWindowStateChanged;
                    }

                    UpdateToolButtonStates();
                    StartToolStateTimer();
                }

            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"Panel init failed: {ex}");
                Content = new TextBlock
                {
                    Text = "MicroEng panel failed to load. See log for details.",
                    Margin = new Thickness(12)
                };
            }
        }

        private void InitThemeToggle()
        {
            if (ThemeToggle == null)
            {
                return;
            }

            _settingThemeToggle = true;
            ThemeToggle.IsChecked = MicroEngWpfUiTheme.CurrentTheme == MicroEngThemeMode.Light;
            _settingThemeToggle = false;

            ThemeToggle.Checked += ThemeToggle_Checked;
            ThemeToggle.Unchecked += ThemeToggle_Unchecked;

            MicroEngWpfUiTheme.ThemeChanged += OnThemeChanged;
            Unloaded += (_, __) => MicroEngWpfUiTheme.ThemeChanged -= OnThemeChanged;
        }

        private void StartToolStateTimer()
        {
            if (_toolStateTimer != null)
            {
                return;
            }

            _toolStateTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(750)
            };
            _toolStateTimer.Tick += (_, __) => UpdateToolButtonStates();
            _toolStateTimer.Start();

            Unloaded += (_, __) =>
            {
                try
                {
                    _toolStateTimer?.Stop();
                    _toolStateTimer = null;
                }
                catch
                {
                    // ignore
                }
            };
        }

        private void ThemeToggle_Checked(object sender, RoutedEventArgs e)
        {
            if (_settingThemeToggle)
            {
                return;
            }

            MicroEngWpfUiTheme.SetTheme(MicroEngThemeMode.Light);
        }

        private void ThemeToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            if (_settingThemeToggle)
            {
                return;
            }

            MicroEngWpfUiTheme.SetTheme(MicroEngThemeMode.Dark);
        }

        private void OnThemeChanged(MicroEngThemeMode theme)
        {
            if (ThemeToggle == null)
            {
                return;
            }

            void UpdateToggle()
            {
                _settingThemeToggle = true;
                ThemeToggle.IsChecked = theme == MicroEngThemeMode.Light;
                _settingThemeToggle = false;
            }

            if (Dispatcher.CheckAccess())
            {
                UpdateToggle();
            }
            else
            {
                Dispatcher.BeginInvoke((Action)UpdateToggle, DispatcherPriority.Background);
            }
        }

        private void AppendData_Click(object sender, RoutedEventArgs e)
        {
            LogToPanel("Opening Data Mapper...");
            MicroEngActions.TryShowDataMapper(out _);
            UpdateToolButtonStates();
        }

        private void Reconstruct_Click(object sender, RoutedEventArgs e)
        {
            MicroEngActions.Reconstruct();
            LogToPanel("[Reconstruct] executed.");
        }

        private void ZoneFinder_Click(object sender, RoutedEventArgs e)
        {
            MicroEngActions.ZoneFinder();
            LogToPanel("[Zone Finder] executed.");
        }

        private void DataScraper_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MicroEngActions.TryShowDataScraper(null, out _);
                UpdateToolButtonStates();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"[Data Scraper] failed to open: {ex}");
                LogToPanel($"[Data Scraper] failed to open: {ex.Message}");
                System.Windows.MessageBox.Show($"Data Scraper failed to open: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DataMatrix_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (IsDockPaneVisible("MicroEng.DataMatrix.DockPane.MENG"))
                {
                    UpdateToolButtonStates();
                    return;
                }

                MicroEngActions.DataMatrix();
                UpdateToolButtonStates();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"[Data Matrix] failed to open: {ex}");
                LogToPanel($"[Data Matrix] failed to open: {ex.Message}");
                System.Windows.MessageBox.Show($"Data Matrix failed to open: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ViewpointsGenerator_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (IsDockPaneVisible("MicroEng.ViewpointsGenerator.DockPane.MENG"))
                {
                    UpdateToolButtonStates();
                    return;
                }

                MicroEngActions.ViewpointsGenerator();
                UpdateToolButtonStates();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"[Viewpoints Generator] failed to open: {ex}");
                LogToPanel($"[Viewpoints Generator] failed to open: {ex.Message}");
                System.Windows.MessageBox.Show($"Viewpoints Generator failed to open: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SpaceMapper_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (IsDockPaneVisible("MicroEng.SpaceMapper.DockPane.MENG"))
                {
                    UpdateToolButtonStates();
                    return;
                }

                MicroEngActions.SpaceMapper();
                UpdateToolButtonStates();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"[Space Mapper] failed to open: {ex}");
                LogToPanel($"[Space Mapper] failed to open: {ex.Message}");
                System.Windows.MessageBox.Show($"Space Mapper failed to open: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SmartSetGenerator_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MicroEngActions.SmartSetGenerator();
                UpdateToolButtonStates();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"[Smart Set Generator] failed to open: {ex}");
                LogToPanel($"[Smart Set Generator] failed to open: {ex.Message}");
                System.Windows.MessageBox.Show($"Smart Set Generator failed to open: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void QuickColour_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MicroEngActions.QuickColour();
                UpdateToolButtonStates();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"[Quick Colour] failed to open: {ex}");
                LogToPanel($"[Quick Colour] failed to open: {ex.Message}");
                System.Windows.MessageBox.Show($"Quick Colour failed to open: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Sequence4D_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (IsDockPaneVisible("MicroEng.Sequence4D.DockPane.MENG"))
                {
                    UpdateToolButtonStates();
                    return;
                }

                MicroEngActions.Sequence4D();
                UpdateToolButtonStates();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"[4D Sequence] failed to open: {ex}");
                LogToPanel($"[4D Sequence] failed to open: {ex.Message}");
                System.Windows.MessageBox.Show($"4D Sequence failed to open: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MicroEngActions.TryShowSettings(out _);
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"[Settings] failed to open: {ex}");
                LogToPanel($"[Settings] failed to open: {ex.Message}");
                System.Windows.MessageBox.Show($"Settings failed to open: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LogToPanel(string message)
        {
            if (Dispatcher.CheckAccess())
            {
                LogTextBox.AppendText($"{DateTime.Now:HH:mm:ss} {message}{Environment.NewLine}");
                LogTextBox.ScrollToEnd();
            }
            else
            {
                Dispatcher.Invoke(() => LogToPanel(message));
            }
        }

        private static bool IsDockPaneVisible(string pluginId)
        {
            try
            {
                var record = NavisApp.Plugins.FindPlugin(pluginId);
                if (record?.IsLoaded != true)
                {
                    return false;
                }

                if (record.LoadedPlugin is DockPanePlugin pane)
                {
                    return pane.Visible;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private void UpdateToolButtonStates()
        {
            void SetCardActive(Wpf.Ui.Controls.CardAction card, bool isActive)
            {
                if (card == null)
                {
                    return;
                }

                card.Tag = isActive;
            }

            var dataScraperOpen = MicroEngActions.IsDataScraperOpen;
            var dataMapperOpen = MicroEngActions.IsDataMapperOpen;
            var dataMatrixOpen = IsDockPaneVisible("MicroEng.DataMatrix.DockPane.MENG");
            var smartSetsOpen = MicroEngActions.IsSmartSetGeneratorOpen;
            var viewpointsOpen = IsDockPaneVisible("MicroEng.ViewpointsGenerator.DockPane.MENG");
            var quickColourOpen = MicroEngActions.IsQuickColourOpen;
            var spaceMapperOpen = IsDockPaneVisible("MicroEng.SpaceMapper.DockPane.MENG");
            var sequence4dOpen = IsDockPaneVisible("MicroEng.Sequence4D.DockPane.MENG");

            SetCardActive(DataScraperButton, dataScraperOpen);
            SetCardActive(DataMapperButton, dataMapperOpen);
            SetCardActive(DataMatrixButton, dataMatrixOpen);
            SetCardActive(SmartSetsButton, smartSetsOpen);
            SetCardActive(ViewpointsGeneratorButton, viewpointsOpen);
            SetCardActive(QuickColourButton, quickColourOpen);
            SetCardActive(SpaceMapperButton, spaceMapperOpen);
            SetCardActive(Sequence4DButton, sequence4dOpen);
        }

        private void OnToolWindowStateChanged()
        {
            try
            {
                if (!Dispatcher.CheckAccess())
                {
                    Dispatcher.BeginInvoke(new Action(OnToolWindowStateChanged));
                    return;
                }

                UpdateToolButtonStates();
            }
            catch
            {
                // ignore
            }
        }

        ~MicroEngPanelControl()
        {
            if (!DesignerProperties.GetIsInDesignMode(this))
            {
                MicroEngActions.LogMessage -= LogToPanel;
                MicroEngActions.ToolWindowStateChanged -= OnToolWindowStateChanged;
            }
        }
    }
}
