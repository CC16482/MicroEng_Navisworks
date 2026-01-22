using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using NavisApp = Autodesk.Navisworks.Api.Application;
using ElementHost = System.Windows.Forms.Integration.ElementHost;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.ViewpointsGenerator
{
    [Plugin("MicroEng.ViewpointsGenerator.DockPane", "MENG",
        DisplayName = "Viewpoints Generator",
        ToolTip = "Batch-generate Saved Viewpoints.")]
    [DockPanePlugin(900, 650, FixedSize = false, AutoScroll = true, MinimumHeight = 420, MinimumWidth = 520)]
    public class ViewpointsGeneratorDockPane : DockPanePlugin
    {
        public override Control CreateControlPane()
        {
            MicroEngActions.Init();
            MicroEngActions.Log("ViewpointsGeneratorDockPane: CreateControlPane start");

            try
            {
                var host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new ViewpointsGeneratorControl()
                };
                MicroEngActions.Log("ViewpointsGeneratorDockPane: CreateControlPane created ElementHost");
                return host;
            }
            catch (System.Exception ex)
            {
                MicroEngActions.Log($"ViewpointsGeneratorDockPane: CreateControlPane failed: {ex}");
                return new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new System.Windows.Controls.TextBlock
                    {
                        Text = "Viewpoints Generator failed to load. See MicroEng.log for details.",
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

    [Plugin("MicroEng.ViewpointsGenerator.Command", "MENG",
        DisplayName = "Viewpoints Generator",
        ToolTip = "Show the Viewpoints Generator panel.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class ViewpointsGeneratorCommand : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Init();

            const string paneId = "MicroEng.ViewpointsGenerator.DockPane.MENG";
            try
            {
                var record = NavisApp.Plugins.FindPlugin(paneId);
                if (record == null)
                {
                    MessageBox.Show($"Could not find Viewpoints Generator dock pane plugin '{paneId}'.",
                        "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 0;
                }

                if (!record.IsLoaded)
                {
                    MicroEngActions.Log("ViewpointsGeneratorCommand: loading plugin");
                    record.LoadPlugin();
                }

                if (record.LoadedPlugin is DockPanePlugin pane)
                {
                    MicroEngActions.Log("ViewpointsGeneratorCommand: setting pane visible");
                    pane.Visible = true;
                }
            }
            catch (System.Exception ex)
            {
                MicroEngActions.Log($"ViewpointsGeneratorCommand failed: {ex}");
                MessageBox.Show($"Viewpoints Generator failed: {ex.Message}",
                    "Viewpoints Generator", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return 0;
        }
    }
}
