using System;
using System.Collections.Generic;
using System.Windows.Automation;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.ApplicationParts;
using NavisApp = Autodesk.Navisworks.Api.Application;

namespace MicroEng.Navisworks
{
    internal sealed class DockPaneVisibilitySnapshot
    {
        public DockPaneVisibilitySnapshot(
            Dictionary<string, bool> visibleById,
            Dictionary<string, bool> builtInVisibleByCommandId = null,
            HashSet<string> capturedDisplayNames = null)
        {
            VisibleById = visibleById ?? new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            BuiltInVisibleByCommandId = builtInVisibleByCommandId ?? new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            CapturedDisplayNames = capturedDisplayNames ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        public Dictionary<string, bool> VisibleById { get; }
        public Dictionary<string, bool> BuiltInVisibleByCommandId { get; }
        public HashSet<string> CapturedDisplayNames { get; }
    }

    internal static class NavisworksDockPaneManager
    {
        private static readonly TimeSpan DefaultDockPaneSettleDelay = TimeSpan.FromMilliseconds(500);

        private static readonly HashSet<string> DefaultExcludedPaneIds = new(StringComparer.OrdinalIgnoreCase)
        {
            "MicroEng.SpaceMapper.DockPane.MENG"
        };

        private static readonly HashSet<string> KnownDockPaneDisplayNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "Clash Detective",
            "TimeLiner",
            "Timeliner",
            "Autodesk Rendering",
            "Animator",
            "Scripter",
            "Quantification",
            "Quantification Workbook",
            "Item Catalog",
            "Resource Catalog",
            "Selection Tree",
            "Selection Tree (Compact)",
            "Sets",
            "Properties",
            "Item Properties",
            "Find Items",
            "Selection Inspector",
            "Comments",
            "Find Comments",
            "Set Scale by Measurement",
            "Saved Viewpoints",
            "Tilt",
            "Plan View",
            "Section View",
            "Section Plane Settings",
            "Property Favorites",
            "Civil Alignments",
            "Measure Tools",
            "Appearance Profiler",
            "BIM 360 Shared Views",
            "BIM 360 Glue Shared Views",
            "Shared Views",
            "Sheet Browser",
            "Find Item in Other Sheets and Models"
        };

        private static readonly string[] KnownDockPaneIdKeywords =
        {
            "SelectionTree",
            "Selection Tree",
            "FindItems",
            "Find Items",
            "Properties",
            "Sets",
            "SelectionInspector",
            "Selection Inspector"
        };

        private sealed class BuiltInDockPaneCommand
        {
            public BuiltInDockPaneCommand(string commandId, params string[] names)
            {
                CommandId = commandId;
                Names = names ?? Array.Empty<string>();
            }

            public string CommandId { get; }
            public string[] Names { get; }
        }

        private static readonly BuiltInDockPaneCommand[] BuiltInDockPaneCommands =
        {
            new BuiltInDockPaneCommand(
                "RoamerGUI_OM_VIEW_TREE",
                "Selection Tree",
                "Selection Tree (Compact)"),
            new BuiltInDockPaneCommand(
                "RoamerGUI_OM_FIND_ITEM",
                "Find Items"),
            new BuiltInDockPaneCommand(
                "RoamerGUI_OM_ATTRIB_BAR",
                "Properties",
                "Item Properties"),
            new BuiltInDockPaneCommand(
                "SelectionInspectorCommand.Navisworks",
                "Selection Inspector"),
            new BuiltInDockPaneCommand(
                "ClashWindowCommand.Navisworks",
                "Clash Detective"),
            new BuiltInDockPaneCommand(
                "TimelinerRibbonCommand.Navisworks",
                "TimeLiner",
                "Timeliner"),
            new BuiltInDockPaneCommand(
                "TakeoffWorkbookRibbonCommand.Navisworks",
                "Quantification",
                "Quantification Workbook"),
            new BuiltInDockPaneCommand(
                "RenderBrowserCommand.Navisworks",
                "Autodesk Rendering"),
            new BuiltInDockPaneCommand(
                "navisworks.animator.plugin.Animator",
                "Animator"),
            new BuiltInDockPaneCommand(
                "navisworks.scripter.plugin.Scripter",
                "Scripter"),
            new BuiltInDockPaneCommand(
                "AutoAppearanceCommand.Navisworks",
                "Appearance Profiler"),
            new BuiltInDockPaneCommand(
                "RoamerGUI_OM_MESSAGE_BOARD",
                "Comments"),
            new BuiltInDockPaneCommand(
                "RoamerGUI_OM_FIND_BAR",
                "Find Comments"),
            new BuiltInDockPaneCommand(
                "RoamerGUI_OM_VP_ORG",
                "Saved Viewpoints"),
            new BuiltInDockPaneCommand(
                "RoamerGUI_OM_VIEW_TILT",
                "Tilt",
                "Tilt Bar",
                "Camera Tilt"),
            new BuiltInDockPaneCommand(
                "RoamerGUI_OM_VIEW_PLAN_THUMBNAIL",
                "Plan View"),
            new BuiltInDockPaneCommand(
                "RoamerGUI_OM_VIEW_XSECT_THUMBNAIL",
                "Section View"),
            new BuiltInDockPaneCommand(
                "RoamerGUI_OM_SECTION_PLANES_DIALOG",
                "Section Plane Settings"),
            new BuiltInDockPaneCommand(
                "RoamerGUI_CIVIL_ALIGNMENT_PANEL",
                "Civil Alignments"),
            new BuiltInDockPaneCommand(
                "RoamerGUI_MEASURE_PANEL",
                "Measure Tools"),
            new BuiltInDockPaneCommand(
                "Bim360ViewsPaneCommand.Navisworks",
                "BIM 360 Shared Views",
                "BIM 360 Glue Shared Views",
                "Shared Views"),
            new BuiltInDockPaneCommand(
                "MultiSheetSearchPaneCommand.Navisworks",
                "Find Item in Other Sheets and Models")
        };

        public static DockPaneVisibilitySnapshot HideDockPanes(IEnumerable<string> extraExclusions = null, TimeSpan? settleDelay = null)
        {
            var exclusions = BuildExclusions(extraExclusions);
            var gui = NavisApp.Gui;
            if (gui == null)
            {
                return new DockPaneVisibilitySnapshot(new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
            }

            var snapshot = CaptureDockPanes(gui, exclusions);
            var builtInSnapshot = CaptureBuiltInDockPanes(gui.MainWindow, snapshot.CapturedDisplayNames);
            foreach (var kvp in snapshot.VisibleById)
            {
                if (!kvp.Value)
                {
                    continue;
                }

                try
                {
                    gui.SetDockPanePluginVisibility(kvp.Key, false);
                }
                catch
                {
                    // ignore
                }
            }

            foreach (var kvp in builtInSnapshot)
            {
                if (!kvp.Value)
                {
                    continue;
                }

                TryExecuteBuiltInCommand(kvp.Key);
            }

            var delay = settleDelay ?? DefaultDockPaneSettleDelay;
            if (delay > TimeSpan.Zero)
            {
                Thread.Sleep(delay);
            }

            return new DockPaneVisibilitySnapshot(snapshot.VisibleById, builtInSnapshot, snapshot.CapturedDisplayNames);
        }

        public static void RestoreDockPanes(DockPaneVisibilitySnapshot snapshot)
        {
            if (snapshot == null || snapshot.VisibleById.Count == 0)
            {
                return;
            }

            var gui = NavisApp.Gui;
            if (gui == null)
            {
                return;
            }

            foreach (var kvp in snapshot.VisibleById)
            {
                try
                {
                    gui.SetDockPanePluginVisibility(kvp.Key, kvp.Value);
                }
                catch
                {
                    // ignore
                }
            }

            foreach (var kvp in snapshot.BuiltInVisibleByCommandId)
            {
                if (!kvp.Value)
                {
                    continue;
                }

                TryExecuteBuiltInCommand(kvp.Key);
            }
        }

        private static DockPaneVisibilitySnapshot CaptureDockPanes(IApplicationGui gui, HashSet<string> exclusions)
        {
            var visibleById = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            var capturedDisplayNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var records = NavisApp.Plugins?.PluginRecords;
            if (records == null)
            {
                return new DockPaneVisibilitySnapshot(visibleById, null, capturedDisplayNames);
            }

            foreach (var record in records)
            {
                if (!IsDockPaneRecord(record) && !IsKnownDockPaneName(record))
                {
                    continue;
                }

                var id = record.Id;
                if (string.IsNullOrWhiteSpace(id) || exclusions.Contains(id))
                {
                    continue;
                }

                bool visible;
                try
                {
                    visible = gui.GetDockPanePluginVisibility(id);
                }
                catch
                {
                    try
                    {
                        if (record.LoadedPlugin is DockPanePlugin pane)
                        {
                            visible = pane.Visible;
                        }
                        else
                        {
                            continue;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                visibleById[id] = visible;

                var displayName = record.DisplayName ?? record.Name ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(displayName))
                {
                    capturedDisplayNames.Add(displayName);
                }

                if (!string.IsNullOrWhiteSpace(record.Name))
                {
                    capturedDisplayNames.Add(record.Name);
                }
            }

            return new DockPaneVisibilitySnapshot(visibleById, null, capturedDisplayNames);
        }

        private static bool IsDockPaneRecord(PluginRecord record)
        {
            if (record == null)
            {
                return false;
            }

            if (record is DockPanePluginRecord)
            {
                return true;
            }

            var interfaces = record.InterfaceRecords;
            if (interfaces == null)
            {
                return false;
            }

            var dockToolId = typeof(Autodesk.Navisworks.Api.Interop.IDockToolPlugin).GUID.ToString("D");
            foreach (var iface in interfaces)
            {
                if (iface == null || string.IsNullOrWhiteSpace(iface.Id))
                {
                    continue;
                }

                if (string.Equals(iface.Id, dockToolId, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsKnownDockPaneName(PluginRecord record)
        {
            if (record == null)
            {
                return false;
            }

            var displayName = record.DisplayName ?? record.Name ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(displayName) && KnownDockPaneDisplayNames.Contains(displayName))
            {
                return true;
            }

            var id = record.Id ?? string.Empty;
            foreach (var keyword in KnownDockPaneIdKeywords)
            {
                if (displayName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0
                    || id.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static Dictionary<string, bool> CaptureBuiltInDockPanes(IWin32Window mainWindow, HashSet<string> skipNames)
        {
            var visibleByCommand = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            if (mainWindow == null)
            {
                return visibleByCommand;
            }

            AutomationElement root = null;
            try
            {
                root = AutomationElement.FromHandle(mainWindow.Handle);
            }
            catch
            {
                return visibleByCommand;
            }

            if (root == null)
            {
                return visibleByCommand;
            }

            foreach (var command in BuiltInDockPaneCommands)
            {
                if (command == null || string.IsNullOrWhiteSpace(command.CommandId))
                {
                    continue;
                }

                if (ShouldSkipBuiltInCommand(command, skipNames))
                {
                    continue;
                }

                var isVisible = IsAutomationPaneVisible(root, command.Names);
                visibleByCommand[command.CommandId] = isVisible;
            }

            return visibleByCommand;
        }

        private static bool IsAutomationPaneVisible(AutomationElement root, IEnumerable<string> names)
        {
            if (root == null || names == null)
            {
                return false;
            }

            foreach (var name in names)
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                try
                {
                    var condition = new PropertyCondition(AutomationElement.NameProperty, name);
                    var elements = root.FindAll(TreeScope.Descendants, condition);
                    if (elements == null || elements.Count == 0)
                    {
                        continue;
                    }

                    foreach (AutomationElement element in elements)
                    {
                        if (element == null)
                        {
                            continue;
                        }

                        var type = element.Current.ControlType;
                        if (type != ControlType.Pane && type != ControlType.Window)
                        {
                            continue;
                        }

                        var className = element.Current.ClassName ?? string.Empty;
                        var automationId = element.Current.AutomationId ?? string.Empty;
                        if (className.IndexOf("Ribbon", StringComparison.OrdinalIgnoreCase) >= 0
                            || automationId.IndexOf("Ribbon", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            continue;
                        }

                        if (IsInRibbon(element))
                        {
                            continue;
                        }

                        var rect = element.Current.BoundingRectangle;
                        if (!element.Current.IsOffscreen && rect.Width >= 20 && rect.Height >= 20)
                        {
                            return true;
                        }
                    }
                }
                catch
                {
                    // ignore
                }
            }

            return false;
        }

        private static bool ShouldSkipBuiltInCommand(BuiltInDockPaneCommand command, HashSet<string> skipNames)
        {
            if (command?.Names == null || skipNames == null || skipNames.Count == 0)
            {
                return false;
            }

            return command.Names.Any(name => !string.IsNullOrWhiteSpace(name) && skipNames.Contains(name));
        }

        private static bool IsInRibbon(AutomationElement element)
        {
            if (element == null)
            {
                return false;
            }

            try
            {
                var walker = TreeWalker.ControlViewWalker;
                var parent = walker.GetParent(element);
                while (parent != null)
                {
                    var className = parent.Current.ClassName ?? string.Empty;
                    var automationId = parent.Current.AutomationId ?? string.Empty;
                    var name = parent.Current.Name ?? string.Empty;
                    if (className.IndexOf("Ribbon", StringComparison.OrdinalIgnoreCase) >= 0
                        || automationId.IndexOf("Ribbon", StringComparison.OrdinalIgnoreCase) >= 0
                        || name.IndexOf("Ribbon", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }

                    parent = walker.GetParent(parent);
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        private static void TryExecuteBuiltInCommand(string commandId)
        {
            if (string.IsNullOrWhiteSpace(commandId))
            {
                return;
            }

            try
            {
                Autodesk.Navisworks.Api.Interop.LcRmFrameworkInterface.ExecuteCommand(
                    commandId,
                    Autodesk.Navisworks.Api.Interop.LcUCIPExecutionContext.eMENU);
            }
            catch
            {
                // ignore
            }
        }

        private static HashSet<string> BuildExclusions(IEnumerable<string> extraExclusions)
        {
            var exclusions = new HashSet<string>(DefaultExcludedPaneIds, StringComparer.OrdinalIgnoreCase);
            if (extraExclusions == null)
            {
                return exclusions;
            }

            foreach (var id in extraExclusions)
            {
                if (!string.IsNullOrWhiteSpace(id))
                {
                    exclusions.Add(id);
                }
            }

            return exclusions;
        }
    }
}
