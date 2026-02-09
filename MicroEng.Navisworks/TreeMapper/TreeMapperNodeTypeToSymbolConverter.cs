using System;
using System.Globalization;
using System.Windows.Data;
using WpfUiControls = Wpf.Ui.Controls;

namespace MicroEng.Navisworks.TreeMapper
{
    internal sealed class TreeMapperNodeTypeToSymbolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not TreeMapperNodeType nodeType)
            {
                return WpfUiControls.SymbolRegular.Box20;
            }

            return nodeType switch
            {
                TreeMapperNodeType.Model => WpfUiControls.SymbolRegular.DocumentTableCube20,
                TreeMapperNodeType.Layer => WpfUiControls.SymbolRegular.Layer20,
                TreeMapperNodeType.Group => WpfUiControls.SymbolRegular.CubeArrowCurveDown20,
                TreeMapperNodeType.Composite => WpfUiControls.SymbolRegular.Box20,
                TreeMapperNodeType.Geometry => WpfUiControls.SymbolRegular.Cube20,
                TreeMapperNodeType.Insert => WpfUiControls.SymbolRegular.CubeArrowCurveDown20,
                TreeMapperNodeType.Instance => WpfUiControls.SymbolRegular.Cube20,
                TreeMapperNodeType.Collection => WpfUiControls.SymbolRegular.Connected20,
                TreeMapperNodeType.Item => WpfUiControls.SymbolRegular.Cube20,
                _ => WpfUiControls.SymbolRegular.Box20
            };
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }
}
