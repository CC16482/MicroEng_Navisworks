using System.Windows.Controls;

namespace MicroEng.Navisworks
{
    public partial class SpaceMapperStepMappingPage : Page
    {
        public SpaceMapperStepMappingPage(SpaceMapperControl host)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            DataContext = host;
        }

        internal Wpf.Ui.Controls.Button AddMappingButton => AddMappingButtonControl;
        internal Wpf.Ui.Controls.Button DeleteMappingButton => DeleteMappingButtonControl;
        internal Wpf.Ui.Controls.Button SaveTemplateButton => SaveTemplateButtonControl;
        internal Wpf.Ui.Controls.Button LoadTemplateButton => LoadTemplateButtonControl;
        internal Wpf.Ui.Controls.DataGrid MappingGrid => MappingGridControl;
    }
}
