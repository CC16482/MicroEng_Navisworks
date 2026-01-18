using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using NavisApp = Autodesk.Navisworks.Api.Application;
using ElementHost = System.Windows.Forms.Integration.ElementHost;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.SmartSets
{
    [Plugin("MicroEng.SmartSetGenerator.DockPane", "MENG",
        DisplayName = "Smart Set Generator",
        ToolTip = "Generate Search Sets and Selection Sets quickly.")]
    [DockPanePlugin(900, 650, FixedSize = false, AutoScroll = true, MinimumHeight = 420, MinimumWidth = 520)]
    public class SmartSetGeneratorDockPane : DockPanePlugin
    {
        public override Control CreateControlPane()
        {
            MicroEngActions.Init();
            MicroEngActions.Log("SmartSetGeneratorDockPane: CreateControlPane start");

            try
            {
                var host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new SmartSetGeneratorControl()
                };
                MicroEngActions.Log("SmartSetGeneratorDockPane: CreateControlPane created ElementHost");
                return host;
            }
            catch (System.Exception ex)
            {
                MicroEngActions.Log($"SmartSetGeneratorDockPane: CreateControlPane failed: {ex}");
                return new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new System.Windows.Controls.TextBlock
                    {
                        Text = "Smart Set Generator failed to load. See MicroEng.log for details.",
                        Margin = new System.Windows.Thickness(12)
                    }
                };
            }
        }

        public override void DestroyControlPane(Control pane)
        {
            pane?.Dispose();
            base.DestroyControlPane(pane);
        }
    }

    [Plugin("MicroEng.SmartSetGenerator.Command", "MENG",
        DisplayName = "Smart Set Generator",
        ToolTip = "Show/hide the Smart Set Generator panel.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class SmartSetGeneratorCommand : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Init();

            const string paneId = "MicroEng.SmartSetGenerator.DockPane.MENG";
            try
            {
                var record = NavisApp.Plugins.FindPlugin(paneId);
                if (record == null)
                {
                    MessageBox.Show($"Could not find Smart Set Generator dock pane plugin '{paneId}'.", "Smart Set Generator",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 0;
                }

                if (!record.IsLoaded)
                {
                    MicroEngActions.Log("SmartSetGeneratorCommand: loading plugin");
                    record.LoadPlugin();
                }

                if (record.LoadedPlugin is DockPanePlugin pane)
                {
                    MicroEngActions.Log("SmartSetGeneratorCommand: toggling visibility");
                    pane.Visible = !pane.Visible;
                }
            }
            catch (System.Exception ex)
            {
                MicroEngActions.Log($"SmartSetGeneratorCommand failed: {ex}");
                MessageBox.Show($"Smart Set Generator failed: {ex.Message}", "Smart Set Generator", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return 0;
        }
    }
}
