using Autodesk.Navisworks.Api.Plugins;

namespace MicroEng.Navisworks
{
    [Plugin("MicroEng.Ribbon", "MENG", DisplayName = "MicroEng NavisTools")]
    [RibbonLayout("MicroEngRibbon.xaml")]
    [RibbonTab("ID_MicroEng_TAB", DisplayName = "MicroEng")]
    [Command("ID_MicroEng_OpenPanel",
        DisplayName = "MicroEng NavisTools",
        ToolTip = "Show or hide the MicroEng NavisTools panel.",
        Icon = "Logos\\microeng_navistools_16.png",
        LargeIcon = "Logos\\microeng_navistools_32.png")]
    public sealed class MicroEngRibbonCommandHandler : CommandHandlerPlugin
    {
        public override int ExecuteCommand(string name, params string[] parameters)
        {
            if (name == "ID_MicroEng_OpenPanel")
            {
                MicroEngActions.Init();
                MicroEngActions.ToggleMainPanel();
            }

            return 0;
        }
    }
}
