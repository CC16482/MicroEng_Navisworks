using System.Windows;
using System.Windows.Controls;

namespace MicroEng.Navisworks.SmartSets
{
    public partial class SmartSetGeneratorFromSelectionPage : UserControl
    {
        private readonly SmartSetGeneratorControl _owner;

        public SmartSetGeneratorFromSelectionPage(SmartSetGeneratorControl owner)
        {
            _owner = owner;
            InitializeComponent();
            DataContext = owner;
        }

        private void AnalyzeSelection_Click(object sender, RoutedEventArgs e)
        {
            _owner.AnalyzeSelection_Click(sender, e);
        }

        private void ApplySuggestion_Click(object sender, RoutedEventArgs e)
        {
            _owner.ApplySuggestion_Click(sender, e);
        }
    }
}
