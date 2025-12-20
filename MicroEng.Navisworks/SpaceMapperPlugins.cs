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
            MicroEngActions.Log("SpaceMapperDockPane: CreateControlPane start");

            try
            {
                var host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    AutoSize = true,
                    Child = new SpaceMapperControl()
                };
                MicroEngActions.Log("SpaceMapperDockPane: CreateControlPane created ElementHost");
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

    [Plugin("MicroEng.SpaceMapper.Command", "MENG",
        DisplayName = "Space Mapper",
        ToolTip = "Show/hide the MicroEng Space Mapper window.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class SpaceMapperCommand : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Init();

            const string paneId = "MicroEng.SpaceMapper.DockPane.MENG";

            try
            {
                var record = Autodesk.Navisworks.Api.Application.Plugins.FindPlugin(paneId);
                if (record == null)
                {
                    MessageBox.Show($"Could not find Space Mapper dock pane plugin '{paneId}'.",
                        "MicroEng Space Mapper",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return 0;
                }

                if (!record.IsLoaded)
                {
                    MicroEngActions.Log("SpaceMapperCommand: loading plugin");
                    record.LoadPlugin();
                }

                if (record.LoadedPlugin is DockPanePlugin pane)
                {
                    MicroEngActions.Log("SpaceMapperCommand: toggling visibility");
                    pane.Visible = !pane.Visible;
                }
            }
            catch (System.Exception ex)
            {
                MicroEngActions.Log($"SpaceMapperCommand failed: {ex}");
                MessageBox.Show($"Space Mapper failed: {ex.Message}", "MicroEng Space Mapper", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return 0;
        }
    }
}
