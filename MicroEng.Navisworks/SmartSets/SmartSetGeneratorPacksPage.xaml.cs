using System.Windows;
using System.Windows.Controls;

namespace MicroEng.Navisworks.SmartSets
{
    public partial class SmartSetGeneratorPacksPage : UserControl
    {
        private readonly SmartSetGeneratorControl _owner;

        public SmartSetGeneratorPacksPage(SmartSetGeneratorControl owner)
        {
            _owner = owner;
            InitializeComponent();
            DataContext = owner;
        }

        private void RunPack_Click(object sender, RoutedEventArgs e)
        {
            _owner.RunPack_Click(sender, e);
        }
    }
}
