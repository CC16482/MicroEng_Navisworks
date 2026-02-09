using System.Windows;
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
        internal FrameworkElement RunHealthChipPanel => RunHealthChipPanelControl;
        internal FrameworkElement RunHealthMissingBoundsChip => RunHealthMissingBoundsChipControl;
        internal TextBlock RunHealthMissingBoundsText => RunHealthMissingBoundsTextControl;
        internal FrameworkElement RunHealthUnmatchedChip => RunHealthUnmatchedChipControl;
        internal TextBlock RunHealthUnmatchedText => RunHealthUnmatchedTextControl;
        internal Wpf.Ui.Controls.Button RunHealthDetailsButton => RunHealthDetailsButtonControl;
        internal Wpf.Ui.Controls.Flyout RunHealthDetailsFlyout => RunHealthDetailsFlyoutControl;
        internal TextBlock RunHealthTargetsTotalText => RunHealthTargetsTotalTextControl;
        internal TextBlock RunHealthTargetsWithBoundsText => RunHealthTargetsWithBoundsTextControl;
        internal TextBlock RunHealthTargetsWithoutBoundsText => RunHealthTargetsWithoutBoundsTextControl;
        internal TextBlock RunHealthTargetsSampledText => RunHealthTargetsSampledTextControl;
        internal TextBlock RunHealthTargetsSampleNoBoundsText => RunHealthTargetsSampleNoBoundsTextControl;
        internal TextBlock RunHealthTargetsSampleNoGeometryText => RunHealthTargetsSampleNoGeometryTextControl;
        internal TextBlock RunHealthTargetsUnmatchedText => RunHealthTargetsUnmatchedTextControl;
        internal TextBlock RunHealthImpactText => RunHealthImpactTextControl;
        internal Wpf.Ui.Controls.Button RunHealthCreateMissingBoundsButton => RunHealthCreateMissingBoundsButtonControl;
        internal Wpf.Ui.Controls.Button RunHealthCreateUnmatchedButton => RunHealthCreateUnmatchedButtonControl;
    }
}
