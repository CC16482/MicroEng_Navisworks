using System.Windows;
using System.Windows.Controls;

namespace MicroEng.Navisworks.SmartSets
{
    public partial class SmartSetGeneratorSmartGroupingPage : UserControl
    {
        private readonly SmartSetGeneratorControl _owner;

        public SmartSetGeneratorSmartGroupingPage(SmartSetGeneratorControl owner)
        {
            _owner = owner;
            InitializeComponent();
            DataContext = owner;
        }

        private void PickGroupBy_Click(object sender, RoutedEventArgs e)
        {
            _owner.PickGroupBy_Click(sender, e);
        }

        private void PickThenBy_Click(object sender, RoutedEventArgs e)
        {
            _owner.PickThenBy_Click(sender, e);
        }

        private void PreviewGroups_Click(object sender, RoutedEventArgs e)
        {
            _owner.PreviewGroups_Click(sender, e);
        }

        private void GenerateGroups_Click(object sender, RoutedEventArgs e)
        {
            _owner.GenerateGroups_Click(sender, e);
        }
    }
}
