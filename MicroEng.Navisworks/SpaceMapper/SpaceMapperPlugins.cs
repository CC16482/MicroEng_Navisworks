using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using ElementHost = System.Windows.Forms.Integration.ElementHost;

namespace MicroEng.Navisworks
{
    [Plugin("MicroEng.SpaceMapper.DockPane", "MENG",
        DisplayName = "MicroEng Space Mapper",
        ToolTip = "Map zone/room/space data onto elements.")]
    [DockPanePlugin(900, 600, FixedSize = false, AutoScroll = true, MinimumHeight = 400, MinimumWidth = 500)]
    public class SpaceMapperDockPane : DockPanePlugin
    {
        public override Control CreateControlPane()
        {
            MicroEngActions.Init();

            try
            {
                var host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    AutoSize = true,
                    Child = new SpaceMapperControl()
                };
                return host;
            }
            catch (System.Exception ex)
            {
                MicroEngActions.Log($"SpaceMapperDockPane: CreateControlPane failed: {ex}");
                return new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new System.Windows.Controls.TextBlock
                    {
                        Text = "Space Mapper failed to load. See MicroEng.log for details.",
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
