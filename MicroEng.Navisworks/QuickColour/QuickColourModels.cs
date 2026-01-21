namespace MicroEng.Navisworks.QuickColour
{
    public enum QuickColourValueMode
    {
        Categorical,
        NumericGradient
    }

    public enum QuickColourScope
    {
        EntireModel,
        CurrentSelection,
        SavedSelectionSet,
        ModelTree,
        PropertyFilter
    }

    public enum QuickColourPaletteStyle
    {
        Deep,
        Pastel
    }

    public enum QuickColourTypeSortMode
    {
        Count,
        Name
    }

    public sealed class QuickColourValueRow : NotifyBase
    {
        private bool _enabled = true;
        private string _value = "";
        private int _count;
        private System.Windows.Media.Color _color = System.Windows.Media.Colors.LightGray;
        private string _colorHex = "#D3D3D3";

        public bool Enabled { get => _enabled; set => SetField(ref _enabled, value); }

        public string Value
        {
            get => _value;
            set
            {
                if (SetField(ref _value, value ?? ""))
                {
                    OnPropertyChanged(nameof(DisplayValue));
                }
            }
        }

        public int Count { get => _count; set => SetField(ref _count, value); }

        public System.Windows.Media.Color Color
        {
            get => _color;
            set
            {
                if (SetField(ref _color, value))
                {
                    _colorHex = QuickColourPalette.ToHex(value);
                    OnPropertyChanged(nameof(ColorHex));
                    OnPropertyChanged(nameof(SwatchBrush));
                }
            }
        }

        public string ColorHex
        {
            get => _colorHex;
            set
            {
                if (QuickColourPalette.TryParseHex(value, out var parsed))
                {
                    _color = parsed;
                    _colorHex = QuickColourPalette.ToHex(parsed);
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(Color));
                    OnPropertyChanged(nameof(SwatchBrush));
                }
                else
                {
                    _colorHex = QuickColourPalette.ToHex(_color);
                    OnPropertyChanged();
                }
            }
        }

        public string DisplayValue => string.IsNullOrWhiteSpace(Value) ? "<blank>" : Value;

        public System.Windows.Media.Brush SwatchBrush => new System.Windows.Media.SolidColorBrush(Color);
    }
}
