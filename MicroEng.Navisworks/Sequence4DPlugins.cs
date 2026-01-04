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
            MicroEngActions.Log("Sequence4DDockPane: CreateControlPane start");

            try
            {
                var host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = new Sequence4DControl()
                };
                MicroEngActions.Log("Sequence4DDockPane: CreateControlPane created ElementHost");
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

    [Plugin("MicroEng.Sequence4D.Command", "MENG",
        DisplayName = "4D Sequence Tool (MicroEng)",
        ToolTip = "Show/hide the MicroEng 4D Sequence dock pane.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class Sequence4DCommand : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Init();

            const string paneId = "MicroEng.Sequence4D.DockPane.MENG";
            try
            {
                var record = Autodesk.Navisworks.Api.Application.Plugins.FindPlugin(paneId);
                if (record == null)
                {
                    MessageBox.Show($"Could not find 4D Sequence dock pane plugin '{paneId}'.", "MicroEng",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 0;
                }

                if (!record.IsLoaded)
                {
                    MicroEngActions.Log("Sequence4DCommand: loading plugin");
                    record.LoadPlugin();
                }

                if (record.LoadedPlugin is DockPanePlugin pane)
                {
                    MicroEngActions.Log("Sequence4DCommand: toggling visibility");
                    pane.Visible = !pane.Visible;
                }
            }
            catch (System.Exception ex)
            {
                MicroEngActions.Log($"Sequence4DCommand failed: {ex}");
                MessageBox.Show($"4D Sequence failed: {ex.Message}", "MicroEng", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return 0;
        }
    }
}
