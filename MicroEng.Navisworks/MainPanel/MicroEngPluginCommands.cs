using System;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using ElementHost = System.Windows.Forms.Integration.ElementHost;

namespace MicroEng.Navisworks
{
    [Plugin("MicroEng.DockPane", "MENG",
        DisplayName = "NavisTools",
        ToolTip = "Dockable panel for MicroEng NavisTools.")]
    [DockPanePlugin(800, 600, FixedSize = false, AutoScroll = true, MinimumHeight = 480, MinimumWidth = 360)]
    public class MicroEngDockPane : DockPanePlugin
    {
        public override Control CreateControlPane()
        {
            MicroEngActions.Init();
            try
            {
                var host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new MicroEngPanelControl()
                };
                return host;
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"DockPane: CreateControlPane failed: {ex}");
                var fallback = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new System.Windows.Controls.TextBlock
                    {
                        Text = "MicroEng panel failed to load. Check log.",
                        Margin = new System.Windows.Thickness(12)
                    }
                };
                return fallback;
            }
        }

        public override void DestroyControlPane(Control pane)
        {
            pane?.Dispose();
        }
    }

}
