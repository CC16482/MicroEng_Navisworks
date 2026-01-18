using System;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using WpfWindow = System.Windows.Window;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using NavisApp = Autodesk.Navisworks.Api.Application;
using DrawingColor = System.Drawing.Color;
using DrawingColorTranslator = System.Drawing.ColorTranslator;
using ElementHost = System.Windows.Forms.Integration.ElementHost;

namespace MicroEng.Navisworks
{
    internal static class ThemeAssets
    {
        private static readonly string AssetRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logos");
        private static readonly Lazy<Bitmap> _ribbonIcon = new(() => LoadBitmap("microeng_logotray.png"));
        private static readonly Lazy<Bitmap> _headerLogo = new(() => LoadBitmap("microeng-logo2.png"));
        public static DrawingColor BackgroundPanel => DrawingColorTranslator.FromHtml("#f5f7fb");
        public static DrawingColor BackgroundMuted => DrawingColorTranslator.FromHtml("#eef1f4");
        public static DrawingColor Accent => DrawingColorTranslator.FromHtml("#8ba9d9");
        public static DrawingColor AccentStrong => DrawingColorTranslator.FromHtml("#6b89c9");
        public static DrawingColor TextPrimary => DrawingColorTranslator.FromHtml("#111827");
        public static DrawingColor TextSecondary => DrawingColorTranslator.FromHtml("#374151");

        public static Font DefaultFont => new Font("Segoe UI", 9F, FontStyle.Regular);
        public static Bitmap RibbonIcon => _ribbonIcon.Value;
        public static Bitmap HeaderLogo => _headerLogo.Value;

        private static Bitmap LoadBitmap(string fileName)
        {
            try
            {
                var path = Path.Combine(AssetRoot, fileName);
                return File.Exists(path) ? (Bitmap)Image.FromFile(path) : null;
            }
            catch
            {
                return null;
            }
        }
    }

    // ========= Shared logic for the 3 tools =========

    internal static class MicroEngActions
    {
        private static DataScraperWindow _dataScraperWindow;
        private static AppendIntegrateDialog _dataMapperWindow;
        private static MicroEngSettingsWindow _settingsWindow;

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

        private static void SafeWireUnhandledExceptionLogging()
        {
            try
            {
                AppDomain.CurrentDomain.UnhandledException -= OnUnhandled;
                AppDomain.CurrentDomain.UnhandledException += OnUnhandled;
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
            }
            catch
            {
                // swallow
            }
        }

        public static void AppendData()
        {
            TryShowDataMapper(out _);
        }

        internal static bool IsDataScraperOpen => _dataScraperWindow?.IsVisible == true;
        internal static bool IsDataMapperOpen => _dataMapperWindow?.IsVisible == true;

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

            Log("AppendData: launching Data Mapper dialog");
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
                Log("AppendData: dialog shown");
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

        public static void Reconstruct()
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                ShowInfo("No active document.");
                return;
            }

            // TODO: implement your "Reconstruct" logic here
            ShowInfo("[Reconstruct] Placeholder - hook up your reconstruction logic.");
        }

        public static void ZoneFinder()
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                ShowInfo("No active document.");
                return;
            }

            // TODO: implement "Zone Finder" logic here
            ShowInfo("[Zone Finder] Placeholder - implement your zone logic here.");
        }

        private static void ShowInfo(string message)
        {
            MessageBox.Show(message, "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        internal static void Log(string message)
        {
            System.Diagnostics.Trace.WriteLine($"[MicroEng] {message}");
            LogMessage?.Invoke(message);
            var line = $"{DateTime.Now:HH:mm:ss} {message}";
            WriteLogLine(LogFilePathPrimary, line);
            WriteLogLine(LogFilePathFallback, line);
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

        public static void DataMatrix()
        {
            const string dockPanePluginId = "MicroEng.DataMatrix.DockPane.MENG";
            Log("DataMatrix: locating plugin record");
            var record = NavisApp.Plugins.FindPlugin(dockPanePluginId);
            if (record == null)
            {
                MessageBox.Show($"Could not find Data Matrix dock pane plugin '{dockPanePluginId}'.",
                    "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!record.IsLoaded)
            {
                Log("DataMatrix: loading plugin");
                record.LoadPlugin();
            }

            if (record.LoadedPlugin is DockPanePlugin pane)
            {
                Log("DataMatrix: setting pane visible");
                pane.Visible = true;
            }
        }

        public static void SpaceMapper()
        {
            const string dockPanePluginId = "MicroEng.SpaceMapper.DockPane.MENG";
            Log("SpaceMapper: locating plugin record");
            var record = NavisApp.Plugins.FindPlugin(dockPanePluginId);
            if (record == null)
            {
                MessageBox.Show($"Could not find Space Mapper dock pane plugin '{dockPanePluginId}'.",
                    "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!record.IsLoaded)
            {
                Log("SpaceMapper: loading plugin");
                record.LoadPlugin();
            }

            if (record.LoadedPlugin is DockPanePlugin pane)
            {
                Log("SpaceMapper: setting pane visible");
                pane.Visible = true;
            }
        }

        public static void SmartSetGenerator()
        {
            const string dockPanePluginId = "MicroEng.SmartSetGenerator.DockPane.MENG";
            Log("SmartSetGenerator: locating plugin record");
            var record = NavisApp.Plugins.FindPlugin(dockPanePluginId);
            if (record == null)
            {
                MessageBox.Show($"Could not find Smart Set Generator dock pane plugin '{dockPanePluginId}'.",
                    "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!record.IsLoaded)
            {
                Log("SmartSetGenerator: loading plugin");
                record.LoadPlugin();
            }

            if (record.LoadedPlugin is DockPanePlugin pane)
            {
                Log("SmartSetGenerator: setting pane visible");
                pane.Visible = true;
            }
        }

        public static void Sequence4D()
        {
            const string dockPanePluginId = "MicroEng.Sequence4D.DockPane.MENG";
            Log("Sequence4D: locating plugin record");
            var record = NavisApp.Plugins.FindPlugin(dockPanePluginId);
            if (record == null)
            {
                MessageBox.Show($"Could not find 4D Sequence dock pane plugin '{dockPanePluginId}'.",
                    "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!record.IsLoaded)
            {
                Log("Sequence4D: loading plugin");
                record.LoadPlugin();
            }

            if (record.LoadedPlugin is DockPanePlugin pane)
            {
                Log("Sequence4D: setting pane visible");
                pane.Visible = true;
            }
        }
    }

    // ========= Standalone Add-In commands (nice for NavisAddinManager testing) =========

    [Plugin("MicroEng.AppendData", "MENG",
        DisplayName = "MicroEng Data Mapper",
        ToolTip = "Map and append data to the current selection.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class AppendDataAddIn : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.AppendData();
            return 0;
        }
    }

    [Plugin("MicroEng.Reconstruct", "MENG",
        DisplayName = "MicroEng Reconstruct",
        ToolTip = "Reconstruct / rebuild MicroEng structures.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class ReconstructAddIn : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Reconstruct();
            return 0;
        }
    }

    [Plugin("MicroEng.ZoneFinder", "MENG",
        DisplayName = "MicroEng Zone Finder",
        ToolTip = "Find and analyse zones in the model.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class ZoneFinderAddIn : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.ZoneFinder();
            return 0;
        }
    }

    // ========= Dockable panel =========

    [Plugin("MicroEng.DockPane", "MENG",
        DisplayName = "MicroEng Tools",
        ToolTip = "Dockable panel for MicroEng tools.")]
    [DockPanePlugin(800, 600, FixedSize = false, AutoScroll = true, MinimumHeight = 480, MinimumWidth = 360)]
    public class MicroEngDockPane : DockPanePlugin
    {
        public override Control CreateControlPane()
        {
            MicroEngActions.Log("DockPane: CreateControlPane start");
            MicroEngActions.Init();
            try
            {
                var host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new MicroEngPanelControl()
                };
                MicroEngActions.Log("DockPane: CreateControlPane created ElementHost");
                return host;
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"DockPane: CreateControlPane failed: {ex}");
                var fallback = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new System.Windows.Controls.TextBlock
                    {
                        Text = "MicroEng panel failed to load. Check log.",
                        Margin = new System.Windows.Thickness(12)
                    }
                };
                return fallback;
            }
        }

        public override void DestroyControlPane(Control pane)
        {
            pane?.Dispose();
        }
    }

    // Command that toggles the dockable panel
    [Plugin("MicroEng.PanelCommand", "MENG",
        DisplayName = "MicroEng Panel",
        ToolTip = "Show/hide the MicroEng tools panel.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class MicroEngPanelCommand : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Init();
            // ID pattern is: PluginAttribute.Id + "." + VendorId
            // e.g. "MicroEng.DockPane.MENG"
            const string dockPanePluginId = "MicroEng.DockPane.MENG";

            var record = NavisApp.Plugins.FindPlugin(dockPanePluginId);
            if (record == null)
            {
                MessageBox.Show($"Could not find dock pane plugin '{dockPanePluginId}'.",
                    "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }

            if (!record.IsLoaded)
            {
                record.LoadPlugin();
            }

            var pane = record.LoadedPlugin as DockPanePlugin;
            if (pane == null)
            {
                MessageBox.Show("Found plugin record but could not cast to DockPanePlugin.",
                    "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }

            pane.Visible = !pane.Visible;
            return 0;
        }
    }

    internal static class AssemblyResolver
    {
        private static readonly object Sync = new object();
        private static bool _registered;
        private static bool _firstChanceHooked;
        private static int _firstChanceLogged;
        private const int FirstChanceLogLimit = 10;

        private static readonly string AssemblyDir =
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? AppDomain.CurrentDomain.BaseDirectory;

        private static readonly string LogFilePathPrimary = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MicroEng.Navisworks",
            "NavisErrors",
            "MicroEng.log");

        private static readonly string LogFilePathFallback = Path.Combine(Path.GetTempPath(), "MicroEng.log");

        private static readonly HashSet<string> ResolveLogOnce = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        [ThreadStatic]
        private static HashSet<string> _resolving;

        public static void EnsureRegistered()
        {
            if (_registered) return;

            lock (Sync)
            {
                if (_registered) return;
                AppDomain.CurrentDomain.AssemblyResolve += OnResolve;
                if (!_firstChanceHooked)
                {
                    AppDomain.CurrentDomain.FirstChanceException += OnFirstChance;
                    _firstChanceHooked = true;
                }
                _registered = true;
            }
        }

        private static void OnFirstChance(object sender, System.Runtime.ExceptionServices.FirstChanceExceptionEventArgs e)
        {
            try
            {
                // Only log the first few FileLoad/FileNotFound exceptions for our assemblies.
                if (_firstChanceLogged >= FirstChanceLogLimit)
                    return;

                var ex = e.Exception;
                switch (ex)
                {
                    case System.IO.FileLoadException fle:
                        if (IsOurs(fle.FileName))
                        {
                            _firstChanceLogged++;
                            SafeLog($"[FirstChance] FileLoad: {fle.FileName} :: {fle.Message}");
                        }
                        break;
                    case System.IO.FileNotFoundException fnf:
                        if (IsOurs(fnf.FileName))
                        {
                            _firstChanceLogged++;
                            SafeLog($"[FirstChance] FileNotFound: {fnf.FileName} :: {fnf.Message}");
                        }
                        break;
                }
            }
            catch
            {
                // swallow
            }
        }

        private static bool IsOurs(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return false;
            return fileName.IndexOf("MicroEng", StringComparison.OrdinalIgnoreCase) >= 0
                   || fileName.IndexOf("Wpf.Ui", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static Assembly OnResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                var requested = new AssemblyName(args.Name);
                var name = requested.Name ?? string.Empty;
                if (string.IsNullOrEmpty(name))
                    return null;

                // Ignore resource satellites.
                if (name.EndsWith(".resources", StringComparison.OrdinalIgnoreCase))
                    return null;

                // Only handle our own and Wpf.Ui assemblies.
                if (!ShouldHandle(name))
                {
                    return null;
                }

                // If it's already loaded, never load it again (avoids duplicate-load issues that can break WPF pack URIs).
                var alreadyLoaded = TryGetLoaded(name);
                if (alreadyLoaded != null)
                {
                    return alreadyLoaded;
                }

                _resolving ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (!_resolving.Add(name))
                {
                    return null;
                }

                try
                {
                    var candidate = Path.Combine(AssemblyDir, name + ".dll");
                    if (!File.Exists(candidate))
                    {
                        return null;
                    }

                    // Re-check in case another thread loaded it while we were waiting.
                    alreadyLoaded = TryGetLoaded(name);
                    if (alreadyLoaded != null)
                    {
                        return alreadyLoaded;
                    }

                    SafeResolveLogOnce($"[AssemblyResolve] Loading {name} from {candidate}");
                    return Assembly.LoadFrom(candidate);
                }
                finally
                {
                    _resolving.Remove(name);
                }
            }
            catch
            {
                // swallow
            }
            return null;
        }

        private static bool ShouldHandle(string assemblySimpleName)
        {
            return assemblySimpleName.StartsWith("MicroEng", StringComparison.OrdinalIgnoreCase)
                   || assemblySimpleName.StartsWith("Wpf.Ui", StringComparison.OrdinalIgnoreCase);
        }

        private static Assembly TryGetLoaded(string simpleName)
        {
            try
            {
                foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                {
                    var asmName = asm.GetName().Name;
                    if (string.Equals(asmName, simpleName, StringComparison.OrdinalIgnoreCase))
                    {
                        return asm;
                    }
                }
            }
            catch
            {
                // swallow
            }

            return null;
        }

        private static void SafeResolveLogOnce(string message)
        {
            try
            {
                lock (Sync)
                {
                    if (!ResolveLogOnce.Add(message))
                    {
                        return;
                    }
                }

                SafeLog(message);
            }
            catch
            {
                // swallow
            }
        }

        private static void SafeLog(string message)
        {
            var line = $"{DateTime.Now:HH:mm:ss} {message}";
            WriteLogLine(LogFilePathPrimary, line);
            WriteLogLine(LogFilePathFallback, line);
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
            catch
            {
                // swallow
            }
        }
    }
}
