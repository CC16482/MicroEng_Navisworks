using System.Windows.Controls;

namespace MicroEng.Navisworks
{
    public partial class SpaceMapperStepSetupPage : Page
    {
        public SpaceMapperStepSetupPage(SpaceMapperControl host)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            DataContext = host;
        }

        internal ComboBox ProfileCombo => ProfileComboBox;
        internal Wpf.Ui.Controls.Button RefreshProfilesButton => RefreshProfilesButtonControl;
        internal ComboBox ScopeCombo => ScopeComboBox;
        internal Wpf.Ui.Controls.Button RunScraperButton => RunScraperButtonControl;
        internal Wpf.Ui.Controls.Button RunButton => RunButtonControl;
    }
}
