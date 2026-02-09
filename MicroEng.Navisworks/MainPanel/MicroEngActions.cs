using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using NavisApp = Autodesk.Navisworks.Api.Application;
using ElementHost = System.Windows.Forms.Integration.ElementHost;
using WpfWindow = System.Windows.Window;
using MicroEng.Navisworks.QuickColour;
using MicroEng.Navisworks.SmartSets;
using MicroEng.Navisworks.TreeMapper;

namespace MicroEng.Navisworks
{
    internal static class MicroEngActions
    {
        private static DataScraperWindow _dataScraperWindow;
        private static AppendIntegrateDialog _dataMapperWindow;
        private static MicroEngSettingsWindow _settingsWindow;
        private static SmartSetGeneratorWindow _smartSetGeneratorWindow;
        private static QuickColourWindow _quickColourWindow;
        private static TreeMapperWindow _treeMapperWindow;

        internal static event Action ToolWindowStateChanged;

        private static void RaiseToolWindowStateChanged()
        {
            try
            {
                ToolWindowStateChanged?.Invoke();
            }
            catch
            {
                // never allow UI-state notifications to crash Navisworks
            }
        }

            private static readonly string LogFilePathPrimary = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MicroEng.Navisworks",
            "NavisErrors",
            "MicroEng.log");
        private static readonly string LogFilePathFallback = Path.Combine(Path.GetTempPath(), "MicroEng.log");

        public static event Action<string> LogMessage;

        static MicroEngActions()
        {
            AssemblyResolver.EnsureRegistered();
            SafeWireUnhandledExceptionLogging();
            EnsureLogFiles();
            WriteLogLine(LogFilePathPrimary, $"=== MicroEng init {DateTime.Now:u} (primary)");
            WriteLogLine(LogFilePathFallback, $"=== MicroEng init {DateTime.Now:u} (fallback)");
            try
            {
                using var proc = Process.GetCurrentProcess();
                var exePath = proc.MainModule?.FileName ?? "<unknown>";
                WriteLogLine(LogFilePathPrimary, $"Process={exePath}, PID={proc.Id}, BaseDir={AppDomain.CurrentDomain.BaseDirectory}");
            }
            catch
            {
                // swallow
            }

            LogInstallLocations();
        }

        public static void Init()
        {
            // Intentionally empty: calling this ensures the static ctor runs and log files exist.
        }

        private static void LogInstallLocations()
        {
            try
            {
                var loadedFrom = typeof(MicroEngActions).Assembly.Location;
                Log($"Assembly={loadedFrom}");

                var programDataPlugin = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "Autodesk", "Navisworks Manage 2025", "Plugins", "MicroEng.Navisworks", "MicroEng.Navisworks.dll");

                var programFilesPlugin = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    "Autodesk", "Navisworks Manage 2025", "plugins", "MicroEng.Navisworks", "MicroEng.Navisworks.dll");

                var appDataBundlePlugin = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Autodesk", "ApplicationPlugins", "MENG.MicroEng.Navisworks.bundle", "Contents", "MicroEng.Navisworks.dll");

                var known = new[] { programDataPlugin, programFilesPlugin, appDataBundlePlugin };
                foreach (var path in known)
                {
                    if (!File.Exists(path))
                        continue;

                    if (!string.Equals(path, loadedFrom, StringComparison.OrdinalIgnoreCase))
                    {
                        Log($"[warning] Another MicroEng.Navisworks.dll exists at: {path}");
                    }
                }
            }
            catch
            {
                // swallow
            }
        }

        private static bool _dispatcherHooked;
        private static bool _taskSchedulerHooked;
        private static int _crashReportIndex;

        private static void SafeWireUnhandledExceptionLogging()
        {
            try
            {
                AppDomain.CurrentDomain.UnhandledException -= OnUnhandled;
                AppDomain.CurrentDomain.UnhandledException += OnUnhandled;
                WireDispatcherUnhandledException();
                WireTaskSchedulerUnhandledException();
            }
            catch
            {
                // swallow
            }
        }

        private static void WireDispatcherUnhandledException()
        {
            if (_dispatcherHooked)
            {
                return;
            }

            try
            {
                var app = System.Windows.Application.Current;
                if (app == null)
                {
                    return;
                }

                app.DispatcherUnhandledException -= OnDispatcherUnhandled;
                app.DispatcherUnhandledException += OnDispatcherUnhandled;
                _dispatcherHooked = true;
            }
            catch
            {
                // swallow
            }
        }

        private static void WireTaskSchedulerUnhandledException()
        {
            if (_taskSchedulerHooked)
            {
                return;
            }

            try
            {
                System.Threading.Tasks.TaskScheduler.UnobservedTaskException -= OnUnobservedTaskException;
                System.Threading.Tasks.TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
                _taskSchedulerHooked = true;
            }
            catch
            {
                // swallow
            }
        }

        private static void OnUnhandled(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                var ex = e.ExceptionObject as Exception;
                Log($"[Unhandled] {(e.IsTerminating ? "Terminating" : "Non-terminating")}: {ex}");
                WriteCrashReport("Unhandled", ex, e.IsTerminating);
            }
            catch
            {
                // swallow
            }
        }

        private static void OnDispatcherUnhandled(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            try
            {
                Log($"[DispatcherUnhandled] {e.Exception}");
                WriteCrashReport("DispatcherUnhandled", e.Exception, isTerminating: true);
            }
            catch
            {
                // swallow
            }
        }

        private static void OnUnobservedTaskException(object sender, System.Threading.Tasks.UnobservedTaskExceptionEventArgs e)
        {
            try
            {
                Log($"[UnobservedTaskException] {e.Exception}");
            }
            catch
            {
                // swallow
            }
        }

        private static void WriteCrashReport(string source, Exception ex, bool isTerminating)
        {
            try
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MicroEng.Navisworks",
                    "NavisErrors",
                    "CrashReports");

                Directory.CreateDirectory(dir);

                var index = Interlocked.Increment(ref _crashReportIndex);
                var fileName = $"Crash_{DateTime.Now:yyyyMMdd_HHmmss}_{source}_{index}.txt";
                var path = Path.Combine(dir, fileName);

                var sb = new StringBuilder();
                sb.AppendLine("MicroEng Crash Report");
                sb.AppendLine($"Timestamp: {DateTime.Now:O}");
                sb.AppendLine($"Source: {source}");
                sb.AppendLine($"IsTerminating: {isTerminating}");
                sb.AppendLine($"ThreadId: {Thread.CurrentThread.ManagedThreadId}");

                try
                {
                    using var proc = Process.GetCurrentProcess();
                    sb.AppendLine($"Process: {proc.MainModule?.FileName ?? "<unknown>"}");
                    sb.AppendLine($"ProcessId: {proc.Id}");
                }
                catch
                {
                    // ignore
                }

                try
                {
                    sb.AppendLine($"BaseDir: {AppDomain.CurrentDomain.BaseDirectory}");
                    sb.AppendLine($"Assembly: {typeof(MicroEngActions).Assembly.Location}");
                }
                catch
                {
                    // ignore
                }

                try
                {
                    var doc = NavisApp.ActiveDocument;
                    sb.AppendLine($"Document: {doc?.FileName ?? "<none>"}");
                    sb.AppendLine($"SelectionCount: {doc?.CurrentSelection?.SelectedItems?.Count ?? 0}");
                }
                catch
                {
                    // ignore
                }

                try
                {
                    var builtInSummary = NavisworksDockPaneManager.GetBuiltInDockPaneVisibilitySummary();
                    sb.AppendLine($"BuiltInDockPanes: {builtInSummary}");
                }
                catch
                {
                    // ignore
                }

                sb.AppendLine();
                sb.AppendLine("Exception:");
                sb.AppendLine(ex?.ToString() ?? "<null>");

                sb.AppendLine();
                sb.AppendLine("Recent MicroEng.log:");
                AppendRecentLogLines(sb, LogFilePathPrimary, 200);

                File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
                Log($"[CrashReport] Wrote {path}");
            }
            catch
            {
                // swallow
            }
        }

        private static void AppendRecentLogLines(StringBuilder sb, string path, int maxLines)
        {
            try
            {
                if (!File.Exists(path))
                {
                    sb.AppendLine("(log file missing)");
                    return;
                }

                var lines = File.ReadAllLines(path);
                var start = Math.Max(0, lines.Length - Math.Max(1, maxLines));
                for (var i = start; i < lines.Length; i++)
                {
                    sb.AppendLine(lines[i]);
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"(log read failed: {ex.Message})");
            }
        }

        internal static bool IsDataScraperOpen => _dataScraperWindow?.IsVisible == true;
        internal static bool IsDataMapperOpen => _dataMapperWindow?.IsVisible == true;
        internal static bool IsSmartSetGeneratorOpen => _smartSetGeneratorWindow?.IsVisible == true;
        internal static bool IsQuickColourOpen => _quickColourWindow?.IsVisible == true;
        internal static bool IsTreeMapperOpen => _treeMapperWindow?.IsVisible == true;
        internal static bool IsSettingsOpen => _settingsWindow?.IsVisible == true;

        private static bool TryActivateWindow(WpfWindow window)
        {
            if (window == null || !window.IsVisible)
            {
                return false;
            }

            try
            {
                if (window.WindowState == System.Windows.WindowState.Minimized)
                {
                    window.WindowState = System.Windows.WindowState.Normal;
                }

                window.Activate();
                window.Focus();
                return true;
            }
            catch
            {
                return false;
            }
        }

        internal static bool TryShowDataScraper(string initialProfile, out DataScraperWindow window)
        {
            window = _dataScraperWindow;
            if (TryActivateWindow(window))
            {
                return true;
            }

            try
            {
                var createdWindow = new DataScraperWindow(initialProfile);
                _dataScraperWindow = createdWindow;
                window = createdWindow;
                createdWindow.Closed += (_, __) =>
                {
                    if (ReferenceEquals(_dataScraperWindow, createdWindow))
                    {
                        _dataScraperWindow = null;
                    }
                    RaiseToolWindowStateChanged();
                };
                ElementHost.EnableModelessKeyboardInterop(createdWindow);
                createdWindow.Show();
                TryActivateWindow(createdWindow);
                RaiseToolWindowStateChanged();
                return true;
            }
            catch (Exception ex)
            {
                Log($"Data Scraper failed: {ex}");
                MessageBox.Show($"Data Scraper failed: {ex.Message}", "MicroEng",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        internal static bool TryShowDataMapper(out AppendIntegrateDialog dialog)
        {
            dialog = _dataMapperWindow;
            if (TryActivateWindow(dialog))
            {
                return true;
            }

            try
            {
                var createdDialog = new AppendIntegrateDialog
                {
                    LogAction = Log
                };
                _dataMapperWindow = createdDialog;
                dialog = createdDialog;
                createdDialog.Closed += (_, __) =>
                {
                    if (ReferenceEquals(_dataMapperWindow, createdDialog))
                    {
                        _dataMapperWindow = null;
                    }
                    RaiseToolWindowStateChanged();
                };

                ElementHost.EnableModelessKeyboardInterop(createdDialog);
                createdDialog.Show();
                TryActivateWindow(createdDialog);
                RaiseToolWindowStateChanged();
                return true;
            }
            catch (Exception ex)
            {
                Log($"Data Mapper failed: {ex}");
                MessageBox.Show($"Data Mapper failed: {ex.Message}", "MicroEng",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        internal static bool TryShowSettings(out MicroEngSettingsWindow window)
        {
            window = _settingsWindow;
            if (TryActivateWindow(window))
            {
                return true;
            }

            try
            {
                var createdWindow = new MicroEngSettingsWindow();
                _settingsWindow = createdWindow;
                window = createdWindow;
                createdWindow.Closed += (_, __) =>
                {
                    if (ReferenceEquals(_settingsWindow, createdWindow))
                    {
                        _settingsWindow = null;
                    }
                    RaiseToolWindowStateChanged();
                };
                ElementHost.EnableModelessKeyboardInterop(createdWindow);
                createdWindow.Show();
                TryActivateWindow(createdWindow);
                RaiseToolWindowStateChanged();
                return true;
            }
            catch (Exception ex)
            {
                Log($"Settings failed: {ex}");
                MessageBox.Show($"Settings failed: {ex.Message}", "MicroEng",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        internal static bool TryShowSmartSetGenerator(out SmartSetGeneratorWindow window)
        {
            window = _smartSetGeneratorWindow;
            if (TryActivateWindow(window))
            {
                return true;
            }

            try
            {
                var createdWindow = new SmartSetGeneratorWindow();
                _smartSetGeneratorWindow = createdWindow;
                window = createdWindow;
                createdWindow.Closed += (_, __) =>
                {
                    if (ReferenceEquals(_smartSetGeneratorWindow, createdWindow))
                    {
                        _smartSetGeneratorWindow = null;
                    }
                    RaiseToolWindowStateChanged();
                };
                ElementHost.EnableModelessKeyboardInterop(createdWindow);
                createdWindow.Show();
                TryActivateWindow(createdWindow);
                RaiseToolWindowStateChanged();
                return true;
            }
            catch (Exception ex)
            {
                Log($"Smart Set Generator failed: {ex}");
                MessageBox.Show($"Smart Set Generator failed: {ex.Message}", "MicroEng",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        internal static bool TryShowTreeMapper(out TreeMapperWindow window)
        {
            window = _treeMapperWindow;
            if (TryActivateWindow(window))
            {
                return true;
            }

            try
            {
                var createdWindow = new TreeMapperWindow();
                _treeMapperWindow = createdWindow;
                window = createdWindow;
                createdWindow.Closed += (_, __) =>
                {
                    if (ReferenceEquals(_treeMapperWindow, createdWindow))
                    {
                        _treeMapperWindow = null;
                    }
                    RaiseToolWindowStateChanged();
                };
                ElementHost.EnableModelessKeyboardInterop(createdWindow);
                createdWindow.Show();
                TryActivateWindow(createdWindow);
                RaiseToolWindowStateChanged();
                return true;
            }
            catch (Exception ex)
            {
                Log($"Tree Mapper failed: {ex}");
                MessageBox.Show($"Tree Mapper failed: {ex.Message}", "MicroEng",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        internal static bool TryShowQuickColour(out QuickColourWindow window)
        {
            window = _quickColourWindow;
            if (TryActivateWindow(window))
            {
                return true;
            }

            try
            {
                var createdWindow = new QuickColourWindow();
                _quickColourWindow = createdWindow;
                window = createdWindow;
                createdWindow.Closed += (_, __) =>
                {
                    if (ReferenceEquals(_quickColourWindow, createdWindow))
                    {
                        _quickColourWindow = null;
                    }
                    RaiseToolWindowStateChanged();
                };
                ElementHost.EnableModelessKeyboardInterop(createdWindow);
                createdWindow.Show();
                TryActivateWindow(createdWindow);
                RaiseToolWindowStateChanged();
                return true;
            }
            catch (Exception ex)
            {
                Log($"Quick Colour failed: {ex}");
                MessageBox.Show($"Quick Colour failed: {ex.Message}", "MicroEng",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        internal static void Log(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            if (!ShouldEmitLogMessage(message))
            {
                return;
            }

            System.Diagnostics.Trace.WriteLine($"[MicroEng] {message}");
            LogMessage?.Invoke(message);
            var line = $"{DateTime.Now:HH:mm:ss} {message}";
            WriteLogLine(LogFilePathPrimary, line);
            WriteLogLine(LogFilePathFallback, line);
        }

        private static bool ShouldEmitLogMessage(string message)
        {
            if (IsVerboseLoggingEnabled())
            {
                return true;
            }

            var text = message?.Trim() ?? string.Empty;
            if (text.Length == 0)
            {
                return false;
            }

            if (IsHighPriorityLogMessage(text))
            {
                return true;
            }

            if (text.StartsWith("Tool:", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (IsDiagnosticNoise(text))
            {
                return false;
            }

            return true;
        }

        private static bool IsVerboseLoggingEnabled()
        {
            return string.Equals(
                Environment.GetEnvironmentVariable("MICROENG_VERBOSE_LOG"),
                "1",
                StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsHighPriorityLogMessage(string message)
        {
            var text = message ?? string.Empty;
            return ContainsAny(text,
                "failed",
                "error",
                "exception",
                "[Unhandled]",
                "[DispatcherUnhandled]",
                "[UnobservedTaskException]",
                "[CrashReport]",
                "[warning]");
        }

        private static bool IsDiagnosticNoise(string message)
        {
            var text = message ?? string.Empty;

            if (StartsWithAny(text,
                "Assembly=",
                "Theme: broadcasting",
                "Theme: accentMode=",
                "Theme: datagridGridLineColor=",
                "[ColumnBuilder][Perf]",
                "[ColumnBuilder][Choose]",
                "DockPanes:",
                "SpaceMapper Processing:",
                "SpaceMapper: navigating to Processing page",
                "SpaceMapper: Processing page loaded in",
                "SmartSets preview (fast):",
                "SmartSets preview (live):",
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
                "Sequence4D: setting pane visible"))
            {
                return true;
            }

            return false;
        }

        private static bool StartsWithAny(string text, params string[] prefixes)
        {
            if (string.IsNullOrWhiteSpace(text) || prefixes == null || prefixes.Length == 0)
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

        private static bool ContainsAny(string text, params string[] values)
        {
            if (string.IsNullOrWhiteSpace(text) || values == null || values.Length == 0)
            {
                return false;
            }

            foreach (var value in values)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                if (text.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static void EnsureLogFiles()
        {
            WriteLogLine(LogFilePathPrimary, $"[startup] log path primary={LogFilePathPrimary}");
            WriteLogLine(LogFilePathFallback, $"[startup] log path fallback={LogFilePathFallback}");
        }

        private static void WriteLogLine(string path, string line)
        {
            try
            {
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                File.AppendAllLines(path, new[] { line });
            }
            catch (Exception ex)
            {
                if (!string.Equals(path, LogFilePathFallback, StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        File.AppendAllLines(LogFilePathFallback, new[] { $"[logwrite fail -> {path}] {ex.Message}" });
                    }
                    catch
                    {
                        // swallow
                    }
                }
            }
        }

        public static void ToggleMainPanel()
        {
            const string dockPanePluginId = "MicroEng.DockPane.MENG";
            var record = NavisApp.Plugins.FindPlugin(dockPanePluginId);
            if (record == null)
            {
                MessageBox.Show($"Could not find dock pane plugin '{dockPanePluginId}'.",
                    "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!record.IsLoaded)
            {
                record.LoadPlugin();
            }

            if (record.LoadedPlugin is DockPanePlugin pane)
            {
                pane.Visible = !pane.Visible;
            }
        }

        public static void DataMatrix()
        {
            const string dockPanePluginId = "MicroEng.DataMatrix.DockPane.MENG";
            var record = NavisApp.Plugins.FindPlugin(dockPanePluginId);
            if (record == null)
            {
                MessageBox.Show($"Could not find Data Matrix dock pane plugin '{dockPanePluginId}'.",
                    "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!record.IsLoaded)
            {
                record.LoadPlugin();
            }

            if (record.LoadedPlugin is DockPanePlugin pane)
            {
                pane.Visible = true;
            }
        }

        public static void ViewpointsGenerator()
        {
            const string dockPanePluginId = "MicroEng.ViewpointsGenerator.DockPane.MENG";
            var record = NavisApp.Plugins.FindPlugin(dockPanePluginId);
            if (record == null)
            {
                MessageBox.Show($"Could not find Viewpoints Generator dock pane plugin '{dockPanePluginId}'.",
                    "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!record.IsLoaded)
            {
                record.LoadPlugin();
            }

            if (record.LoadedPlugin is DockPanePlugin pane)
            {
                pane.Visible = true;
            }
        }

        public static void SpaceMapper()
        {
            const string dockPanePluginId = "MicroEng.SpaceMapper.DockPane.MENG";
            var record = NavisApp.Plugins.FindPlugin(dockPanePluginId);
            if (record == null)
            {
                MessageBox.Show($"Could not find Space Mapper dock pane plugin '{dockPanePluginId}'.",
                    "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!record.IsLoaded)
            {
                record.LoadPlugin();
            }

            if (record.LoadedPlugin is DockPanePlugin pane)
            {
                pane.Visible = true;
            }
        }

        public static void SmartSetGenerator()
        {
            TryShowSmartSetGenerator(out _);
        }

        public static void TreeMapper()
        {
            TryShowTreeMapper(out _);
        }

        public static void QuickColour()
        {
            TryShowQuickColour(out _);
        }

        public static void Sequence4D()
        {
            const string dockPanePluginId = "MicroEng.Sequence4D.DockPane.MENG";
            var record = NavisApp.Plugins.FindPlugin(dockPanePluginId);
            if (record == null)
            {
                MessageBox.Show($"Could not find 4D Sequence dock pane plugin '{dockPanePluginId}'.",
                    "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!record.IsLoaded)
            {
                record.LoadPlugin();
            }

            if (record.LoadedPlugin is DockPanePlugin pane)
            {
                pane.Visible = true;
            }
        }
    }
}
