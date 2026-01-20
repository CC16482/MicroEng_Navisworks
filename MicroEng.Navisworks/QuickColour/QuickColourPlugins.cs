using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.QuickColour
{
    [Plugin("MicroEng.QuickColour.Command", "MENG",
        DisplayName = "Quick Colour",
        ToolTip = "Colorize model items using hierarchy-based profiles.")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class QuickColourCommand : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MicroEngActions.Init();
            try
            {
                MicroEngActions.TryShowQuickColour(out _);
            }
            catch (System.Exception ex)
            {
                MicroEngActions.Log($"Quick Colour failed: {ex}");
                MessageBox.Show($"Quick Colour failed: {ex.Message}", "Quick Colour",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return 0;
        }
    }
}
