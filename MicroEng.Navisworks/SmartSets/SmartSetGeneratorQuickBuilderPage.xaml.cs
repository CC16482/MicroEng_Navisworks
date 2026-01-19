using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace MicroEng.Navisworks.SmartSets
{
    public partial class SmartSetGeneratorQuickBuilderPage : UserControl
    {
        private readonly SmartSetGeneratorControl _owner;

        public SmartSetGeneratorQuickBuilderPage(SmartSetGeneratorControl owner)
        {
            _owner = owner;
            InitializeComponent();
            DataContext = owner;
        }

        private void SaveRecipe_Click(object sender, RoutedEventArgs e)
        {
            _owner.SaveRecipe_Click(sender, e);
        }

        private void SaveRecipeAs_Click(object sender, RoutedEventArgs e)
        {
            _owner.SaveRecipeAs_Click(sender, e);
        }

        private void LoadRecipe_Click(object sender, RoutedEventArgs e)
        {
            _owner.LoadRecipe_Click(sender, e);
        }

        private void AddRule_Click(object sender, RoutedEventArgs e)
        {
            _owner.AddRule_Click(sender, e);
        }

        private void RemoveRule_Click(object sender, RoutedEventArgs e)
        {
            _owner.RemoveRule_Click(sender, e);
        }

        private void DuplicateRule_Click(object sender, RoutedEventArgs e)
        {
            _owner.DuplicateRule_Click(sender, e);
        }

        private void RulesGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _owner.RulesGrid_PreviewMouseLeftButtonDown(sender, e);
        }

        private void RulesGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            _owner.RulesGrid_PreparingCellForEdit(sender, e);
        }

        private void UseCurrentSelectionScope_Click(object sender, RoutedEventArgs e)
        {
            _owner.UseCurrentSelectionScope_Click(sender, e);
        }

        private void ClearScope_Click(object sender, RoutedEventArgs e)
        {
            _owner.ClearScope_Click(sender, e);
        }

        private void Preview_Click(object sender, RoutedEventArgs e)
        {
            _owner.Preview_Click(sender, e);
        }

        private void SelectResults_Click(object sender, RoutedEventArgs e)
        {
            _owner.SelectResults_Click(sender, e);
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            _owner.Generate_Click(sender, e);
        }
    }
}
