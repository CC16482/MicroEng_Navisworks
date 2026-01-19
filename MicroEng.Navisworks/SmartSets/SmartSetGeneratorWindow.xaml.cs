using System.Windows;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.SmartSets
{
    public partial class SmartSetGeneratorWindow : Window
    {
        static SmartSetGeneratorWindow()
        {
            AssemblyResolver.EnsureRegistered();
        }

        public SmartSetGeneratorWindow()
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(this);
        }
    }
}
