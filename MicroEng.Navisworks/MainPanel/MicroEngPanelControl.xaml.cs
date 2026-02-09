using System;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Text.RegularExpressions;
using Autodesk.Navisworks.Api.Plugins;
using MicroEng.Navisworks.QuickColour;
using NavisApp = Autodesk.Navisworks.Api.Application;
using System.Linq;
using System.Collections.Generic;

namespace MicroEng.Navisworks
{
    public partial class MicroEngPanelControl : UserControl
    {
        private bool _settingThemeToggle;
        private DispatcherTimer _toolStateTimer;
        private readonly Dictionary<string, bool> _lastToolStates = new(StringComparer.OrdinalIgnoreCase);
        private bool _toolStateBaselineCaptured;
        private const int MaxPanelLogLines = 300;
        private readonly List<PanelLogEntry> _panelLogEntries = new(MaxPanelLogLines);
        private static readonly FontFamily PanelLogFontFamily = new("Consolas");
        private const double PanelLogFontSize = 12d;
        private static readonly Brush LogTimestampBrush = CreateVibrantSyntaxBrush("#FFBE0B", Brushes.DarkGray);
        private static readonly Brush LogErrorBrush = CreateVibrantSyntaxBrush("#FF006E", Brushes.IndianRed);
        private static readonly Brush LogWarningBrush = CreateVibrantSyntaxBrush("#FB5607", Brushes.DarkOrange);
        private static readonly Brush LogSuccessBrush = CreateVibrantSyntaxBrush("#3A86FF", Brushes.MediumSeaGreen);
        private static readonly Brush LogStatsBrush = CreateVibrantSyntaxBrush("#3A86FF", Brushes.Goldenrod);
        private static readonly Brush LogToolBrush = CreateVibrantSyntaxBrush("#FF006E", Brushes.DeepSkyBlue);
        private static readonly Brush LogOpenedBrush = CreateVibrantSyntaxBrush("#33E653", Brushes.LimeGreen);
        private static readonly Regex NumericTokenRegex = new(@"\d+(?:[.,]\d+)*", RegexOptions.Compiled);
        private static readonly Regex OpenCloseTokenRegex = new(@"\b(opened|closed)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private readonly struct PanelLogEntry
        {
            public PanelLogEntry(string timestamp, string message)
            {
                Timestamp = timestamp ?? string.Empty;
                Message = message ?? string.Empty;
            }

            public string Timestamp { get; }
            public string Message { get; }
        }

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
                // Force software rendering to reduce GPU/driver issues in host.
                System.Windows.Media.RenderOptions.ProcessRenderMode = System.Windows.Interop.RenderMode.SoftwareOnly;
                if (diagnosticBypass)
                {
                    Content = new TextBlock
                    {
                        Text = "MicroEng panel placeholder (diagnostic bypass).",
                        Margin = new Thickness(12)
                    };
                }
                else
                {
                    try
                    {
                        InitializeComponent();
                        InitializePanelLog();
                        TryLoadHeaderLogo();
                    }
                    catch (Exception initEx)
                    {
                        MicroEngActions.Log($"Panel: InitializeComponent failed: {initEx}");
                        throw;
                    }
                    try
                    {
                        MicroEngWpfUiTheme.ApplyTo(this);
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

        private void TryLoadHeaderLogo()
        {
            UpdateHeaderLogo(MicroEngWpfUiTheme.CurrentTheme);
        }

        private void UpdateHeaderLogo(MicroEngThemeMode theme)
        {
            if (HeaderLogoImage == null)
            {
                return;
            }

            try
            {
                var assemblyPath = typeof(MicroEngPanelControl).Assembly.Location;
                var assemblyDir = Path.GetDirectoryName(assemblyPath);
                if (string.IsNullOrWhiteSpace(assemblyDir))
                {
                    return;
                }

                var fileName = theme == MicroEngThemeMode.Light
                    ? "microeng-logo2.png"
                    : "microeng-logo3.png";

                var logoPath = Path.Combine(assemblyDir, "Logos", fileName);
                if (!File.Exists(logoPath))
                {
                    MicroEngActions.Log($"Panel: header logo not found at {logoPath}");
                    HeaderLogoImage.Source = null;
                    HeaderLogoImage.Visibility = Visibility.Collapsed;
                    return;
                }

                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(logoPath, UriKind.Absolute);
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();

                HeaderLogoImage.Source = bitmap;
                HeaderLogoImage.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"Panel: header logo load failed: {ex.Message}");
                HeaderLogoImage.Source = null;
                HeaderLogoImage.Visibility = Visibility.Collapsed;
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
            UpdateThemeModeLabel(MicroEngWpfUiTheme.CurrentTheme);

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
                UpdateThemeModeLabel(theme);
                UpdateHeaderLogo(theme);
                RebuildPanelLogForCurrentTheme();
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

        private void UpdateThemeModeLabel(MicroEngThemeMode theme)
        {
            if (ThemeModeLabel == null)
            {
                return;
            }

            ThemeModeLabel.Text = theme == MicroEngThemeMode.Light ? "Light Mode" : "Dark Mode";
        }

        private void AppendData_Click(object sender, RoutedEventArgs e)
        {
            MicroEngActions.TryShowDataMapper(out _);
            UpdateToolButtonStates();
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

        private void TreeMapper_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MicroEngActions.TreeMapper();
                UpdateToolButtonStates();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"[Tree Mapper] failed to open: {ex}");
                LogToPanel($"[Tree Mapper] failed to open: {ex.Message}");
                MessageBox.Show($"Tree Mapper failed to open: {ex.Message}", "MicroEng",
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
            var panelMessage = NormalizePanelLogMessage(message);
            if (string.IsNullOrWhiteSpace(panelMessage))
            {
                return;
            }

            if (Dispatcher.CheckAccess())
            {
                AddPanelLogEntry(panelMessage);
            }
            else
            {
                Dispatcher.BeginInvoke(
                    DispatcherPriority.Background,
                    new Action(() => AddPanelLogEntry(panelMessage)));
            }
        }

        private void CopyLog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var text = GetPanelLogText();
                if (string.IsNullOrWhiteSpace(text))
                {
                    return;
                }

                Clipboard.SetText(text);
                FlashLogCopyFeedback();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not copy log to clipboard: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ClearLog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                InitializePanelLog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not clear log: {ex.Message}", "MicroEng",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void InitializePanelLog()
        {
            if (LogTextBox?.Document == null)
            {
                return;
            }

            LogTextBox.Document.FontFamily = PanelLogFontFamily;
            LogTextBox.Document.FontSize = PanelLogFontSize;

            _panelLogEntries.Clear();
            LogTextBox.Document.Blocks.Clear();
            AddPanelLogEntry("Ready");
        }

        private void AddPanelLogEntry(string message)
        {
            if (LogTextBox?.Document == null)
            {
                return;
            }

            var entry = new PanelLogEntry(DateTime.Now.ToString("HH:mm:ss"), message);
            _panelLogEntries.Add(entry);
            AppendHighlightedLogLine(entry.Timestamp, entry.Message);
            TrimPanelLogLines();
            LogTextBox.ScrollToEnd();
        }

        private void RebuildPanelLogForCurrentTheme()
        {
            if (LogTextBox?.Document == null)
            {
                return;
            }

            LogTextBox.Document.FontFamily = PanelLogFontFamily;
            LogTextBox.Document.FontSize = PanelLogFontSize;
            LogTextBox.Document.Blocks.Clear();

            foreach (var entry in _panelLogEntries)
            {
                AppendHighlightedLogLine(entry.Timestamp, entry.Message);
            }

            LogTextBox.ScrollToEnd();
        }

        private void AppendHighlightedLogLine(string timestamp, string message)
        {
            if (LogTextBox?.Document == null)
            {
                return;
            }

            var paragraph = new Paragraph
            {
                Margin = new Thickness(0),
                TextAlignment = TextAlignment.Left
            };
            paragraph.Inlines.Add(CreatePanelLogRun($"{timestamp} ", ResolvePanelLogTimestampBrush()));

            if (IsStatsPanelMessage(message))
            {
                AppendStatsMessageRuns(paragraph, message);
            }
            else if (IsOpenClosePanelMessage(message))
            {
                AppendOpenCloseMessageRuns(paragraph, message);
            }
            else
            {
                paragraph.Inlines.Add(CreatePanelLogRun(message, ResolvePanelLogMessageBrush(message)));
            }

            LogTextBox.Document.Blocks.Add(paragraph);
        }

        private void AppendStatsMessageRuns(Paragraph paragraph, string message)
        {
            if (paragraph == null || string.IsNullOrEmpty(message))
            {
                return;
            }

            var defaultBrush = ResolvePanelLogDefaultBrush();
            var matches = NumericTokenRegex.Matches(message);

            if (matches.Count == 0)
            {
                paragraph.Inlines.Add(CreatePanelLogRun(message, defaultBrush));
                return;
            }

            var cursor = 0;
            foreach (Match match in matches)
            {
                if (!match.Success)
                {
                    continue;
                }

                if (match.Index > cursor)
                {
                    var prefix = message.Substring(cursor, match.Index - cursor);
                    paragraph.Inlines.Add(CreatePanelLogRun(prefix, defaultBrush));
                }

                paragraph.Inlines.Add(CreatePanelLogRun(match.Value, LogStatsBrush));
                cursor = match.Index + match.Length;
            }

            if (cursor < message.Length)
            {
                paragraph.Inlines.Add(CreatePanelLogRun(message.Substring(cursor), defaultBrush));
            }
        }

        private static Run CreatePanelLogRun(string text, Brush foreground)
        {
            return new Run(text)
            {
                Foreground = foreground,
                FontFamily = PanelLogFontFamily,
                FontSize = PanelLogFontSize,
                FontWeight = FontWeights.Normal
            };
        }

        private void AppendOpenCloseMessageRuns(Paragraph paragraph, string message)
        {
            if (paragraph == null || string.IsNullOrEmpty(message))
            {
                return;
            }

            var defaultBrush = ResolvePanelLogDefaultBrush();
            var matches = OpenCloseTokenRegex.Matches(message);

            if (matches.Count == 0)
            {
                paragraph.Inlines.Add(CreatePanelLogRun(message, defaultBrush));
                return;
            }

            var cursor = 0;
            foreach (Match match in matches)
            {
                if (!match.Success)
                {
                    continue;
                }

                if (match.Index > cursor)
                {
                    paragraph.Inlines.Add(CreatePanelLogRun(
                        message.Substring(cursor, match.Index - cursor),
                        defaultBrush));
                }

                paragraph.Inlines.Add(CreatePanelLogRun(match.Value, ResolveOpenCloseTokenBrush(match.Value)));
                cursor = match.Index + match.Length;
            }

            if (cursor < message.Length)
            {
                paragraph.Inlines.Add(CreatePanelLogRun(message.Substring(cursor), defaultBrush));
            }
        }

        private void TrimPanelLogLines()
        {
            if (LogTextBox?.Document == null)
            {
                return;
            }

            while (_panelLogEntries.Count > MaxPanelLogLines)
            {
                _panelLogEntries.RemoveAt(0);
            }

            while (LogTextBox.Document.Blocks.Count > MaxPanelLogLines)
            {
                LogTextBox.Document.Blocks.Remove(LogTextBox.Document.Blocks.FirstBlock);
            }
        }

        private void FlashLogCopyFeedback()
        {
            if (LogTextBox == null)
            {
                return;
            }

            var original = LogTextBox.BorderBrush as SolidColorBrush;
            var baseColor = original?.Color ?? (MicroEngWpfUiTheme.CurrentTheme == MicroEngThemeMode.Light
                ? Colors.Gray
                : Colors.LightGray);

            var animatedBrush = new SolidColorBrush(baseColor);
            LogTextBox.BorderBrush = animatedBrush;

            var flash = QuickColourPalette.TryParseHex("#33E653", out var flashColor)
                ? flashColor
                : Color.FromRgb(51, 230, 83);

            var animation = new ColorAnimation
            {
                From = flash,
                To = baseColor,
                Duration = TimeSpan.FromMilliseconds(450),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            animatedBrush.BeginAnimation(SolidColorBrush.ColorProperty, animation);
        }

        private string GetPanelLogText()
        {
            if (LogTextBox?.Document == null)
            {
                return string.Empty;
            }

            var range = new TextRange(LogTextBox.Document.ContentStart, LogTextBox.Document.ContentEnd);
            return (range.Text ?? string.Empty).TrimEnd('\r', '\n');
        }

        private Brush ResolvePanelLogTimestampBrush()
        {
            return LogTimestampBrush;
        }

        private Brush ResolvePanelLogMessageBrush(string message)
        {
            var text = message ?? string.Empty;

            if (text.IndexOf("failed", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("exception", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("[Unhandled]", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("[DispatcherUnhandled]", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("[UnobservedTaskException]", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return LogErrorBrush;
            }

            if (text.IndexOf("warning", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("warn", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return LogWarningBrush;
            }

            if (text.StartsWith("[", StringComparison.OrdinalIgnoreCase) ||
                text.StartsWith("Tool:", StringComparison.OrdinalIgnoreCase) ||
                text.StartsWith("Panel:", StringComparison.OrdinalIgnoreCase))
            {
                return LogToolBrush;
            }

            if (text.IndexOf(" opened", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf(" closed", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("complete", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("success", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return LogSuccessBrush;
            }

            return ResolvePanelLogDefaultBrush();
        }

        private Brush ResolvePanelLogDefaultBrush()
        {
            return MicroEngWpfUiTheme.CurrentTheme == MicroEngThemeMode.Light
                ? Brushes.Black
                : Brushes.White;
        }

        private static bool IsStatsPanelMessage(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return false;
            }

            var text = message;
            var hasDigit = text.Any(char.IsDigit);
            var hasPercent = text.IndexOf('%') >= 0;

            if (!hasDigit && !hasPercent)
            {
                return false;
            }

            return hasPercent || ContainsAny(text,
                " item",
                "items",
                " property",
                "properties",
                " cached",
                " scanned",
                " rows",
                " row",
                " count",
                " total",
                " elapsed",
                " duration",
                " profile",
                " scope",
                " selected",
                " hits",
                " zones",
                " targets",
                " skipped",
                " written",
                " matched",
                " ms",
                " sec",
                " seconds");
        }

        private static bool IsOpenClosePanelMessage(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return false;
            }

            return message.IndexOf("opened", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("closed", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private Brush ResolveOpenCloseTokenBrush(string token)
        {
            if (string.Equals(token, "opened", StringComparison.OrdinalIgnoreCase))
            {
                return LogOpenedBrush;
            }

            return LogToolBrush;
        }

        private static Brush CreateVibrantSyntaxBrush(string hex, Brush fallback)
        {
            if (QuickColourPalette.TryParseHex(hex, out var color))
            {
                var brush = new SolidColorBrush(color);
                if (brush.CanFreeze)
                {
                    brush.Freeze();
                }

                return brush;
            }

            return fallback;
        }

        private static string NormalizePanelLogMessage(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return string.Empty;
            }

            var compact = message.Replace("\r\n", "\n").Trim();
            if (!ShouldShowInPanelLog(compact))
            {
                return string.Empty;
            }

            var firstLine = compact.Split('\n').FirstOrDefault()?.Trim() ?? string.Empty;
            if (firstLine.Length > 240)
            {
                firstLine = firstLine.Substring(0, 237) + "...";
            }

            return firstLine;
        }

        private static bool ShouldShowInPanelLog(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return false;
            }

            if (IsHighPriorityPanelMessage(message))
            {
                return true;
            }

            if (StartsWithAny(message,
                "Panel:",
                "DockPane:",
                "DataMatrixDockPane:",
                "SpaceMapperDockPane:",
                "ViewpointsGeneratorDockPane:",
                "SmartSetGeneratorDockPane:",
                "[ColumnBuilder][Perf]",
                "[ColumnBuilder][Choose]",
                "Theme:",
                "Assembly=",
                "DockPanes:",
                "AppendData: launching Data Mapper dialog",
                "AppendData: dialog shown",
                "DataMatrix: locating plugin record",
                "DataMatrix: loading plugin",
                "DataMatrix: setting pane visible",
                "ViewpointsGenerator: locating plugin record",
                "ViewpointsGenerator: loading plugin",
                "ViewpointsGenerator: setting pane visible",
                "SpaceMapper: locating plugin record",
                "SpaceMapper: loading plugin",
                "SpaceMapper: setting pane visible",
                "Sequence4D: locating plugin record",
                "Sequence4D: loading plugin",
                "Sequence4D: setting pane visible",
                "[DataScraper] Run start",
                "[DataScraper] Run complete"))
            {
                return false;
            }

            return true;
        }

        private static bool IsHighPriorityPanelMessage(string message)
        {
            var text = message ?? string.Empty;
            if (text.IndexOf("failed", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (text.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (text.IndexOf("exception", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (text.IndexOf("[Unhandled]", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (text.IndexOf("[DispatcherUnhandled]", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (text.IndexOf("[UnobservedTaskException]", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (text.IndexOf("[CrashReport]", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        private static bool StartsWithAny(string text, params string[] prefixes)
        {
            if (string.IsNullOrEmpty(text) || prefixes == null || prefixes.Length == 0)
            {
                return false;
            }

            foreach (var prefix in prefixes)
            {
                if (string.IsNullOrWhiteSpace(prefix))
                {
                    continue;
                }

                if (text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool ContainsAny(string text, params string[] fragments)
        {
            if (string.IsNullOrEmpty(text) || fragments == null || fragments.Length == 0)
            {
                return false;
            }

            foreach (var fragment in fragments)
            {
                if (string.IsNullOrWhiteSpace(fragment))
                {
                    continue;
                }

                if (text.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
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
            var treeMapperOpen = MicroEngActions.IsTreeMapperOpen;
            var smartSetsOpen = MicroEngActions.IsSmartSetGeneratorOpen;
            var viewpointsOpen = IsDockPaneVisible("MicroEng.ViewpointsGenerator.DockPane.MENG");
            var quickColourOpen = MicroEngActions.IsQuickColourOpen;
            var spaceMapperOpen = IsDockPaneVisible("MicroEng.SpaceMapper.DockPane.MENG");
            var sequence4dOpen = IsDockPaneVisible("MicroEng.Sequence4D.DockPane.MENG");
            var settingsOpen = MicroEngActions.IsSettingsOpen;

            SetCardActive(DataScraperButton, dataScraperOpen);
            SetCardActive(DataMapperButton, dataMapperOpen);
            SetCardActive(DataMatrixButton, dataMatrixOpen);
            SetCardActive(TreeMapperButton, treeMapperOpen);
            SetCardActive(SmartSetsButton, smartSetsOpen);
            SetCardActive(ViewpointsGeneratorButton, viewpointsOpen);
            SetCardActive(QuickColourButton, quickColourOpen);
            SetCardActive(SpaceMapperButton, spaceMapperOpen);
            SetCardActive(Sequence4DButton, sequence4dOpen);

            var currentStates = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase)
            {
                ["Data Scraper"] = dataScraperOpen,
                ["Data Mapper"] = dataMapperOpen,
                ["Data Matrix"] = dataMatrixOpen,
                ["Tree Mapper"] = treeMapperOpen,
                ["Smart Set Generator"] = smartSetsOpen,
                ["Viewpoints Generator"] = viewpointsOpen,
                ["Quick Colour"] = quickColourOpen,
                ["Space Mapper"] = spaceMapperOpen,
                ["4D Sequence"] = sequence4dOpen,
                ["Settings"] = settingsOpen
            };

            LogToolStateTransitions(currentStates);
        }

        private void LogToolStateTransitions(IReadOnlyDictionary<string, bool> currentStates)
        {
            if (currentStates == null || currentStates.Count == 0)
            {
                return;
            }

            if (!_toolStateBaselineCaptured)
            {
                _lastToolStates.Clear();
                foreach (var pair in currentStates)
                {
                    _lastToolStates[pair.Key] = pair.Value;
                }
                _toolStateBaselineCaptured = true;
                return;
            }

            foreach (var pair in currentStates.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
            {
                var toolName = pair.Key;
                var isOpen = pair.Value;

                if (_lastToolStates.TryGetValue(toolName, out var wasOpen) && wasOpen == isOpen)
                {
                    continue;
                }

                _lastToolStates[toolName] = isOpen;
                MicroEngActions.Log($"Tool: {toolName} {(isOpen ? "opened" : "closed")}");
            }

            foreach (var removedKey in _lastToolStates.Keys.Except(currentStates.Keys, StringComparer.OrdinalIgnoreCase).ToList())
            {
                _lastToolStates.Remove(removedKey);
            }
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
            try
            {
                MicroEngActions.LogMessage -= LogToPanel;
                MicroEngActions.ToolWindowStateChanged -= OnToolWindowStateChanged;
            }
            catch
            {
                // Never throw from finalizer.
            }
        }
    }
}
