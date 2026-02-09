using System;
using System.Globalization;
using System.Windows.Data;

namespace MicroEng.Navisworks.TreeMapper
{
    internal sealed class TreeMapperNodeTypeToLabelConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not TreeMapperNodeType nodeType)
            {
                return string.Empty;
            }

            return nodeType switch
            {
                TreeMapperNodeType.Model => "File",
                TreeMapperNodeType.Layer => "Layer",
                TreeMapperNodeType.Group => "Group",
                TreeMapperNodeType.Composite => "Composite Object",
                TreeMapperNodeType.Insert => "Insert",
                TreeMapperNodeType.Geometry => "Geometry",
                TreeMapperNodeType.Instance => "Instance",
                TreeMapperNodeType.Collection => "Collection",
                TreeMapperNodeType.Item => "Item",
                _ => nodeType.ToString()
            };
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }
}
