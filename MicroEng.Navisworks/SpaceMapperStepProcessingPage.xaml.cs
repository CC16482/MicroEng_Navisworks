using System.Windows.Controls;

namespace MicroEng.Navisworks
{
    public partial class SpaceMapperStepProcessingPage : Page
    {
        public SpaceMapperStepProcessingPage(SpaceMapperControl host)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            DataContext = host;
        }

        internal CheckBox TreatPartialCheck => TreatPartialCheckControl;
        internal CheckBox TagPartialCheck => TagPartialCheckControl;
        internal CheckBox EnableMultiZoneCheck => EnableMultiZoneCheckControl;
        internal ComboBox ProcessingModeCombo => ProcessingModeComboBox;
        internal TextBox Offset3DBox => Offset3DBoxControl;
        internal TextBox OffsetTopBox => OffsetTopBoxControl;
        internal TextBox OffsetBottomBox => OffsetBottomBoxControl;
        internal TextBox OffsetSidesBox => OffsetSidesBoxControl;
        internal TextBox UnitsBox => UnitsBoxControl;
        internal TextBox OffsetModeBox => OffsetModeBoxControl;
        internal TextBox MaxThreadsBox => MaxThreadsBoxControl;
        internal TextBox BatchSizeBox => BatchSizeBoxControl;
    }
}
