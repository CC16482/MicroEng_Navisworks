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

        internal ComboBox TemplateCombo => TemplateComboBox;
        internal Wpf.Ui.Controls.Button NewTemplateButton => NewTemplateButtonControl;
        internal Wpf.Ui.Controls.Button SaveTemplateButton => SaveTemplateButtonControl;
        internal Wpf.Ui.Controls.Button SaveAsTemplateButton => SaveAsTemplateButtonControl;
        internal Wpf.Ui.Controls.Button DeleteTemplateButton => DeleteTemplateButtonControl;
        internal ComboBox ScraperProfileCombo => ScraperProfileComboBox;
        internal Wpf.Ui.Controls.Button RefreshScraperButton => RefreshScraperButtonControl;
        internal Wpf.Ui.Controls.Button OpenScraperButton => OpenScraperButtonControl;
        internal TextBlock ScraperSummaryText => ScraperSummaryTextControl;
        internal TextBlock ReadinessText => ReadinessTextControl;
        internal Wpf.Ui.Controls.Button ValidateButton => ValidateButtonControl;
        internal Wpf.Ui.Controls.Button GoToZonesButton => GoToZonesButtonControl;
    }
}
