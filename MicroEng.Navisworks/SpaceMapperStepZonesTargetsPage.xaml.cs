using System.Windows.Controls;

namespace MicroEng.Navisworks
{
    public partial class SpaceMapperStepZonesTargetsPage : Page
    {
        public SpaceMapperStepZonesTargetsPage(SpaceMapperControl host)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            DataContext = host;
        }

        internal ComboBox ZoneSourceCombo => ZoneSourceComboBox;
        internal TextBox ZoneSetBox => ZoneSetBoxControl;
        internal ComboBox TargetSourceCombo => TargetSourceComboBox;
        internal Wpf.Ui.Controls.Button AddRuleButton => AddRuleButtonControl;
        internal Wpf.Ui.Controls.Button DeleteRuleButton => DeleteRuleButtonControl;
        internal Wpf.Ui.Controls.DataGrid TargetRulesGrid => TargetRulesGridControl;
    }
}
