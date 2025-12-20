using System.Windows.Controls;

namespace MicroEng.Navisworks
{
    public partial class SpaceMapperStepResultsPage : Page
    {
        public SpaceMapperStepResultsPage(SpaceMapperControl host)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            DataContext = host;
        }

        internal Wpf.Ui.Controls.Button ExportStatsButton => ExportStatsButtonControl;
        internal TextBox ResultsSummaryBox => ResultsSummaryBoxControl;
        internal Wpf.Ui.Controls.DataGrid ResultsGrid => ResultsGridControl;
    }
}
