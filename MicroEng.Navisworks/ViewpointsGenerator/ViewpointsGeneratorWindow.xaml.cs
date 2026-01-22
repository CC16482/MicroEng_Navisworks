using System.Windows;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.ViewpointsGenerator
{
    public partial class ViewpointsGeneratorWindow : Window
    {
        static ViewpointsGeneratorWindow()
        {
            AssemblyResolver.EnsureRegistered();
        }

        public ViewpointsGeneratorWindow()
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(this);
        }
    }
}
