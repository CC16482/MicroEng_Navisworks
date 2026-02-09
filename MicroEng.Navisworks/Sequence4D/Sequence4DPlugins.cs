using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using ElementHost = System.Windows.Forms.Integration.ElementHost;

namespace MicroEng.Navisworks
{
    [Plugin("MicroEng.Sequence4D.DockPane", "MENG",
        DisplayName = "MicroEng 4D Sequence",
        ToolTip = "Build 4D Timeliner sequences from selection.")]
    [DockPanePlugin(760, 700, FixedSize = false, AutoScroll = true, MinimumHeight = 420, MinimumWidth = 520)]
    public class Sequence4DDockPane : DockPanePlugin
    {
        public override Control CreateControlPane()
        {
            MicroEngActions.Init();

            try
            {
                var host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new Sequence4DControl()
                };
                return host;
            }
            catch (System.Exception ex)
            {
                MicroEngActions.Log($"Sequence4DDockPane: CreateControlPane failed: {ex}");
                return new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new System.Windows.Controls.TextBlock
                    {
                        Text = "4D Sequence failed to load. See MicroEng.log for details.",
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
