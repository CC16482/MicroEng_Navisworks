using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using MicroEng.Navisworks.QuickColour;

namespace MicroEng.Navisworks.QuickColour.Profiles
{
    public sealed class ColorHexToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var hex = value as string;
            if (QuickColourPalette.TryParseHex(hex, out var color))
            {
                var brush = new SolidColorBrush(color);
                brush.Freeze();
                return brush;
            }

            return Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }
}
