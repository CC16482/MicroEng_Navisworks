using System;
using System.Collections.Generic;
using System.Windows.Automation;
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
            Dictionary<string, bool> builtInVisibleByCommandId = null)
        {
            VisibleById = visibleById ?? new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            BuiltInVisibleByCommandId = builtInVisibleByCommandId ?? new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        }

        public Dictionary<string, bool> VisibleById { get; }
        public Dictionary<string, bool> BuiltInVisibleByCommandId { get; }
    }

    internal static class NavisworksDockPaneManager
    {
        private static readonly HashSet<string> DefaultExcludedPaneIds = new(StringComparer.OrdinalIgnoreCase)
        {
            "MicroEng.SpaceMapper.DockPane.MENG"
        };

        private static readonly HashSet<string> KnownDockPaneDisplayNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "Selection Tree",
            "Selection Tree (Compact)",
            "Sets",
            "Properties",
            "Item Properties",
            "Find Items",
            "Selection Inspector"
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
            public BuiltInDockPaneCommand(string commandId, params string[] automationNames)
            {
                CommandId = commandId;
                AutomationNames = automationNames ?? Array.Empty<string>();
            }

            public string CommandId { get; }
            public string[] AutomationNames { get; }
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
                "Item Properties")
        };

        public static DockPaneVisibilitySnapshot HideDockPanes(IEnumerable<string> extraExclusions = null)
        {
            var exclusions = BuildExclusions(extraExclusions);
            var gui = NavisApp.Gui;
            if (gui == null)
            {
                return new DockPaneVisibilitySnapshot(new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase));
            }

            var snapshot = CaptureDockPanes(gui, exclusions);
            var builtInSnapshot = CaptureBuiltInDockPanes(gui.MainWindow);
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

            return new DockPaneVisibilitySnapshot(snapshot.VisibleById, builtInSnapshot);
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
            var records = NavisApp.Plugins?.PluginRecords;
            if (records == null)
            {
                return new DockPaneVisibilitySnapshot(visibleById);
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
            }

            return new DockPaneVisibilitySnapshot(visibleById);
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

        private static Dictionary<string, bool> CaptureBuiltInDockPanes(IWin32Window mainWindow)
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

                var isVisible = IsAutomationPaneVisible(root, command.AutomationNames);
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
                    var element = root.FindFirst(TreeScope.Descendants, condition);
                    if (element == null)
                    {
                        continue;
                    }

                    var rect = element.Current.BoundingRectangle;
                    if (!element.Current.IsOffscreen && rect.Width > 0 && rect.Height > 0)
                    {
                        return true;
                    }
                }
                catch
                {
                    // ignore
                }
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
