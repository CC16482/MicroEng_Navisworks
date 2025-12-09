using System;
using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

namespace MicroEng.Navisworks
{
    // ========= Shared logic for the 3 tools =========

    internal static class MicroEngActions
    {
        private const string CategoryName = "MicroEng";
        private const string TagPropertyName = "Tag";

        public static event Action<string> LogMessage;

        public static void AppendData()
        {
            var doc = Application.ActiveDocument;
            if (doc == null)
            {
                ShowInfo("No active document.");
                Log("AppendData: no active document.");
                return;
            }

            var selectedItems = doc.CurrentSelection?.SelectedItems;
            if (selectedItems == null || selectedItems.Count == 0)
            {
                ShowInfo("No items selected. Select some items first.");
                Log("AppendData: no selection.");
                return;
            }

            var tagValue = $"ME-AUTO-{DateTime.Now:yyyyMMdd-HHmmss}";
            var processed = 0;
            var updated = 0;

            foreach (ModelItem item in selectedItems)
            {
                processed++;
                if (TryWriteMicroEngTag(item, tagValue))
                {
                    updated++;
                }
            }

            var summary =
                $"Append Data wrote {CategoryName}.{TagPropertyName}='{tagValue}' to {updated}/{processed} item(s).";
            ShowInfo(summary);
            Log(summary);
        }

        public static void Reconstruct()
        {
            var doc = Application.ActiveDocument;
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
            var doc = Application.ActiveDocument;
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

        private static bool TryWriteMicroEngTag(ModelItem item, string tagValue)
        {
            try
            {
                var state = ComBridge.State;
                var path = ComBridge.ToInwOaPath(item);
                var propertyNode = (ComApi.InwGUIPropertyNode2)state.GetGUIPropertyNode(path, true);

                var propertyVector = (ComApi.InwOaPropertyVec)state.ObjectFactory(
                    ComApi.nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

                var tagProperty = (ComApi.InwOaProperty)state.ObjectFactory(
                    ComApi.nwEObjectType.eObjectType_nwOaProperty, null, null);

                tagProperty.name = TagPropertyName;
                tagProperty.UserName = TagPropertyName;
                tagProperty.value = tagValue;

                propertyVector.Properties().Add(tagProperty);

                propertyNode.SetUserDefined(0, CategoryName, CategoryName, propertyVector);
                return true;
            }
            catch (Exception ex)
            {
                Log($"AppendData: failed for '{item.DisplayName}': {ex.Message}");
                return false;
            }
        }

        private static void Log(string message)
        {
            System.Diagnostics.Trace.WriteLine($"[MicroEng] {message}");
            LogMessage?.Invoke(message);
        }
    }

    // ========= Standalone Add-In commands (nice for NavisAddinManager testing) =========

    [Plugin("MicroEng.AppendData", "MENG",
        DisplayName = "MicroEng Append Data",
        ToolTip = "Append data to the current selection.")]
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
    [DockPanePlugin(400, 300)]
    public class MicroEngDockPane : DockPanePlugin
    {
        private MicroEngPanelControl _control;

        public override Control CreateControlPane()
        {
            _control = new MicroEngPanelControl();
            return _control;
        }

        public override void DestroyControlPane(Control pane)
        {
            pane?.Dispose();
            _control = null;
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

            var record = Application.Plugins.FindPlugin(dockPanePluginId);
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

    // ========= WinForms UI for the panel =========

    public class MicroEngPanelControl : UserControl
    {
        private readonly Button _appendDataButton;
        private readonly Button _reconstructButton;
        private readonly Button _zoneFinderButton;
        private readonly TextBox _logTextBox;

        public MicroEngPanelControl()
        {
            Dock = DockStyle.Fill;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                AutoSize = true
            };

            _appendDataButton = new Button
            {
                Text = "Append Data",
                Dock = DockStyle.Top
            };
            _appendDataButton.Click += (s, e) =>
            {
                MicroEngActions.AppendData();
                LogToPanel("[Append Data] executed.");
            };

            _reconstructButton = new Button
            {
                Text = "Reconstruct",
                Dock = DockStyle.Top
            };
            _reconstructButton.Click += (s, e) =>
            {
                MicroEngActions.Reconstruct();
                LogToPanel("[Reconstruct] executed.");
            };

            _zoneFinderButton = new Button
            {
                Text = "Zone Finder",
                Dock = DockStyle.Top
            };
            _zoneFinderButton.Click += (s, e) =>
            {
                MicroEngActions.ZoneFinder();
                LogToPanel("[Zone Finder] executed.");
            };

            _logTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical
            };

            layout.Controls.Add(_appendDataButton, 0, 0);
            layout.Controls.Add(_reconstructButton, 0, 1);
            layout.Controls.Add(_zoneFinderButton, 0, 2);
            layout.Controls.Add(_logTextBox, 0, 3);

            Controls.Add(layout);

            MicroEngActions.LogMessage += LogToPanel;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                MicroEngActions.LogMessage -= LogToPanel;
            }

            base.Dispose(disposing);
        }

        private void LogToPanel(string message)
        {
            if (IsHandleCreated && InvokeRequired)
            {
                BeginInvoke((Action)(() => LogToPanel(message)));
                return;
            }

            _logTextBox.AppendText($"{DateTime.Now:HH:mm:ss} {message}{Environment.NewLine}");
        }
    }
}
