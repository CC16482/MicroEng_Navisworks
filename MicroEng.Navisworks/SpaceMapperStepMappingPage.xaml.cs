using System.Windows;
using System.Windows.Controls;
using WpfFlyout = Wpf.Ui.Controls.Flyout;

namespace MicroEng.Navisworks
{
    public partial class SpaceMapperStepMappingPage : Page
    {
        public SpaceMapperStepMappingPage(SpaceMapperControl host)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            DataContext = host;

            WireHover(MappingNameHelpToggle, MappingNameHelpFlyout);
            WireHover(ZoneCategoryHelpToggle, ZoneCategoryHelpFlyout);
            WireHover(ZonePropertyHelpToggle, ZonePropertyHelpFlyout);
            WireHover(TargetCategoryHelpToggle, TargetCategoryHelpFlyout);
            WireHover(TargetPropertyHelpToggle, TargetPropertyHelpFlyout);
            WireHover(WriteModeHelpToggle, WriteModeHelpFlyout);
            WireHover(MultiZoneHelpToggle, MultiZoneHelpFlyout);
            WireHover(AppendSeparatorHelpToggle, AppendSeparatorHelpFlyout);
            WireHover(EditableHelpToggle, EditableHelpFlyout);
        }

        internal Wpf.Ui.Controls.Button AddMappingButton => AddMappingButtonControl;
        internal Wpf.Ui.Controls.Button DeleteMappingButton => DeleteMappingButtonControl;
        internal Wpf.Ui.Controls.Button SaveTemplateButton => SaveTemplateButtonControl;
        internal Wpf.Ui.Controls.Button LoadTemplateButton => LoadTemplateButtonControl;
        internal Wpf.Ui.Controls.DataGrid MappingGrid => MappingGridControl;
        internal TextBox MappingNameBox => MappingNameBoxControl;
        internal TextBox ZoneCategoryBox => ZoneCategoryBoxControl;
        internal TextBox ZonePropertyBox => ZonePropertyBoxControl;
        internal TextBox TargetCategoryBox => TargetCategoryBoxControl;
        internal TextBox TargetPropertyBox => TargetPropertyBoxControl;
        internal ComboBox WriteModeCombo => WriteModeComboBoxControl;
        internal ComboBox MultiZoneCombo => MultiZoneComboBoxControl;
        internal TextBox AppendSeparatorBox => AppendSeparatorBoxControl;
        internal CheckBox EditableCheckBox => EditableCheckBoxControl;

        private static void WireHover(FrameworkElement trigger, WpfFlyout flyout)
        {
            if (trigger == null || flyout == null)
            {
                return;
            }

            if (flyout.Content is not FrameworkElement content)
            {
                return;
            }

            _ = new HoverFlyoutController(trigger, content, flyout);
        }
    }
}
