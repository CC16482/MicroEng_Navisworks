using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
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

            try
            {
                var host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new DataMatrixControl()
                };
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

}
