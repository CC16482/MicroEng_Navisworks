using System.Windows;
using System.Windows.Controls;
using WpfFlyout = Wpf.Ui.Controls.Flyout;

namespace MicroEng.Navisworks
{
    public partial class SpaceMapperStepZonesTargetsPage : Page
    {
        public SpaceMapperStepZonesTargetsPage(SpaceMapperControl host)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            DataContext = host;

            WireHover(RuleNameHelpToggle, RuleNameHelpFlyout);
            WireHover(RuleTargetDefinitionHelpToggle, RuleTargetDefinitionHelpFlyout);
            WireHover(RuleMinLevelHelpToggle, RuleMinLevelHelpFlyout);
            WireHover(RuleMaxLevelHelpToggle, RuleMaxLevelHelpFlyout);
            WireHover(RuleSetNameHelpToggle, RuleSetNameHelpFlyout);
            WireHover(RuleCategoryFilterHelpToggle, RuleCategoryFilterHelpFlyout);
            WireHover(RuleMembershipHelpToggle, RuleMembershipHelpFlyout);
            WireHover(RuleEnabledHelpToggle, RuleEnabledHelpFlyout);
        }

        internal ComboBox ZoneSourceCombo => ZoneSourceComboBox;
        internal TextBox ZoneSetBox => ZoneSetBoxControl;
        internal Wpf.Ui.Controls.DropDownButton ZoneSetDropDownButton => ZoneSetDropDownButtonControl;
        internal ContextMenu ZoneSetMenu => ZoneSetMenuControl;
        internal TextBox RuleNameBox => RuleNameBoxControl;
        internal ComboBox RuleTargetDefinitionCombo => RuleTargetDefinitionComboBoxControl;
        internal TextBox RuleMinLevelBox => RuleMinLevelBoxControl;
        internal TextBox RuleMaxLevelBox => RuleMaxLevelBoxControl;
        internal TextBox RuleSetNameBox => RuleSetNameBoxControl;
        internal Wpf.Ui.Controls.DropDownButton RuleSetDropDownButton => RuleSetDropDownButtonControl;
        internal ContextMenu RuleSetMenu => RuleSetMenuControl;
        internal TextBox RuleCategoryFilterBox => RuleCategoryFilterBoxControl;
        internal ComboBox RuleMembershipCombo => RuleMembershipComboBoxControl;
        internal CheckBox RuleEnabledCheckBox => RuleEnabledCheckBoxControl;
        internal TextBlock RuleMinMaxHintText => RuleMinMaxHintTextControl;
        internal TextBlock RuleSetNameHintText => RuleSetNameHintTextControl;
        internal Wpf.Ui.Controls.Button AddRuleButton => AddRuleButtonControl;
        internal Wpf.Ui.Controls.Button DeleteRuleButton => DeleteRuleButtonControl;
        internal Wpf.Ui.Controls.DataGrid TargetRulesGrid => TargetRulesGridControl;

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
