using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
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
        public static event Action<string> LogMessage;

        static MicroEngActions()
        {
            AssemblyResolver.EnsureRegistered();
        }

        public static void AppendData()
        {
            try
            {
                var window = new AppendIntegrateDialog
                {
                    LogAction = Log
                };
                ElementHost.EnableModelessKeyboardInterop(window);
                window.Show();
            }
            catch (Exception ex)
            {
                Log($"Data Mapper failed: {ex}");
                MessageBox.Show($"Data Mapper failed: {ex.Message}", "MicroEng",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            var host = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = new MicroEngPanelControl()
            };
            return host;
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
        private static bool _registered;

        public static void EnsureRegistered()
        {
            if (_registered) return;
            AppDomain.CurrentDomain.AssemblyResolve += OnResolve;
            _registered = true;
        }

        private static Assembly OnResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                var requested = new AssemblyName(args.Name);
                var name = requested.Name ?? string.Empty;
                // Ignore resource satellite probes (we don't ship them).
                if (name.EndsWith(".resources", StringComparison.OrdinalIgnoreCase) ||
                    name.StartsWith("Janus.Windows.UI.v3.resources", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                var asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var candidates = new[]
                {
                    Path.Combine(baseDir ?? string.Empty, requested.Name + ".dll"),
                    Path.Combine(asmDir ?? string.Empty, requested.Name + ".dll")
                };

                foreach (var path in candidates)
                {
                    if (File.Exists(path))
                    {
                        // Only log assemblies we actually load (reduces noise).
                        MicroEngActions.Log($"[AssemblyResolve] Loading {requested.Name} from {path}");
                        return Assembly.LoadFrom(path);
                    }
                }
                // Keep failures quiet for non-MicroEng assemblies to avoid noisy logs.
                if (name.Contains("MicroEng", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("Wpf.Ui", StringComparison.OrdinalIgnoreCase))
                {
                    MicroEngActions.Log($"[AssemblyResolve] Failed to resolve {requested.FullName}. BaseDir={baseDir}, AsmDir={asmDir}");
                }
            }
            catch
            {
                // swallow
            }
            return null;
        }
    }
}
