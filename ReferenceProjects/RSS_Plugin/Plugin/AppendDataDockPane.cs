using Autodesk.Navisworks.Api.Plugins;
using System.Windows.Forms;                // For Control, ElementHost, DockStyle
using System.Windows.Forms.Integration;    // For ElementHost
using AppendDataTool.UI;                   // Corrected: Using the UI namespace based on RootNamespace "AppendDataTool"

namespace AppendDataTool.Plugin // Consistent with RootNamespace "AppendDataTool" and Plugin subfolder
{
    [Plugin("AppendDataTool.AppendDataDockPane.MyCompany", // Unique Plugin ID (ensure "MyCompany" is your DevId)
            "MyCompany", // Your Developer ID
            ToolTip = "Append Data Custom Properties Pane",
            DisplayName = "Append Data Pane")]
    [DockPanePlugin(800, 600, FixedSize = false, AutoScroll = true, MinimumHeight = 450, MinimumWidth = 350)] // Adjust dimensions as needed
    public class AppendDataDockPane : DockPanePlugin
    {
        // Store the created control to manage its lifecycle if needed, though ElementHost handles child disposal.
        private ElementHost _hostedWpfControlHost;

        public override Control CreateControlPane()
        {
            // Create an ElementHost to host the WPF UserControl
            _hostedWpfControlHost = new ElementHost
            {
                Dock = DockStyle.Fill // Fill the dock pane
            };

            // Create an instance of your WPF UserControl
            // This now correctly refers to AppendDataControl within the AppendDataTool.UI namespace
            AppendDataControl wpfControl = new AppendDataControl();
            _hostedWpfControlHost.Child = wpfControl;

            return _hostedWpfControlHost;
        }

        public override void DestroyControlPane(Control pane)
        {
            // The pane passed in is the ElementHost itself.
            // ElementHost handles disposing its Child when the ElementHost itself is disposed.
            if (pane != null)
            {
                pane.Dispose();
            }
            _hostedWpfControlHost = null; // Clear the reference
        }

        // Optional: If you need to refresh or update the UI when the pane becomes visible/active
        // public override void OnPaneShown(bool visible)
        // {
        //     base.OnPaneShown(visible);
        //     if (visible && _hostedWpfControlHost?.Child is AppendDataControl wpfControl)
        //     {
        //         // Example: wpfControl.RefreshData(); // If you had such a method
        //     }
        // }
    }
}