using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using NavisApp = Autodesk.Navisworks.Api.Application;

namespace MicroEng.Navisworks
{
    [Plugin("MicroEng.DataMatrix.DockPane", "MENG",
        DisplayName = "MicroEng Data Matrix",
        ToolTip = "Tabular view of scraped model data.")]
    [DockPanePlugin(900, 600, FixedSize = false, AutoScroll = true)]
    public class DataMatrixDockPane : DockPanePlugin
    {
        private DataMatrixPanel _panel;

        public override Control CreateControlPane()
        {
            _panel = new DataMatrixPanel();
            return _panel;
        }

        public override void DestroyControlPane(Control pane)
        {
            _panel?.Dispose();
            _panel = null;
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
            const string paneId = "MicroEng.DataMatrix.DockPane.MENG";
            var record = NavisApp.Plugins.FindPlugin(paneId);
            if (record == null)
            {
                MessageBox.Show($"Could not find Data Matrix dock pane plugin '{paneId}'.", "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }

            if (!record.IsLoaded)
            {
                record.LoadPlugin();
            }

            if (record.LoadedPlugin is DockPanePlugin pane)
            {
                pane.Visible = !pane.Visible;
            }

            return 0;
        }
    }
}
