using System.Windows;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.QuickColour
{
    public partial class QuickColourWindow : Window
    {
        static QuickColourWindow()
        {
            AssemblyResolver.EnsureRegistered();
        }

        public QuickColourWindow()
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(this);
        }
    }
}
