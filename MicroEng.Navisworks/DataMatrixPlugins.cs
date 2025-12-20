using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using NavisApp = Autodesk.Navisworks.Api.Application;
using ElementHost = System.Windows.Forms.Integration.ElementHost;

namespace MicroEng.Navisworks
{
    [Plugin("MicroEng.DataMatrix.DockPane", "MENG",
        DisplayName = "MicroEng Data Matrix",
        ToolTip = "Tabular view of scraped model data.")]
    [DockPanePlugin(900, 600, FixedSize = false, AutoScroll = true)]
    public class DataMatrixDockPane : DockPanePlugin
    {
        public override Control CreateControlPane()
        {
            MicroEngActions.Init();
            MicroEngActions.Log("DataMatrixDockPane: CreateControlPane start");

            try
            {
                var host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new DataMatrixControl()
                };
                MicroEngActions.Log("DataMatrixDockPane: CreateControlPane created ElementHost");
                return host;
            }
            catch (System.Exception ex)
            {
                MicroEngActions.Log($"DataMatrixDockPane: CreateControlPane failed: {ex}");
                return new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new System.Windows.Controls.TextBlock
                    {
                        Text = "Data Matrix failed to load. See MicroEng.log for details.",
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

    [Plugin("MicroEng.DataMatrix.Command", "MENG",
        DisplayName = "Data Matrix",
        ToolTip = "Show/hide the MicroEng Data Matrix panel.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class DataMatrixCommand : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Init();

            const string paneId = "MicroEng.DataMatrix.DockPane.MENG";
            try
            {
                var record = NavisApp.Plugins.FindPlugin(paneId);
                if (record == null)
                {
                    MessageBox.Show($"Could not find Data Matrix dock pane plugin '{paneId}'.", "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 0;
                }

                if (!record.IsLoaded)
                {
                    MicroEngActions.Log("DataMatrixCommand: loading plugin");
                    record.LoadPlugin();
                }

                if (record.LoadedPlugin is DockPanePlugin pane)
                {
                    MicroEngActions.Log("DataMatrixCommand: toggling visibility");
                    pane.Visible = !pane.Visible;
                }
            }
            catch (System.Exception ex)
            {
                MicroEngActions.Log($"DataMatrixCommand failed: {ex}");
                MessageBox.Show($"Data Matrix failed: {ex.Message}", "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return 0;
        }
    }
}
