using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Windows.Media;

namespace MicroEng.Navisworks.QuickColour
{
    public sealed class QuickColourHierarchyTypeRow : NotifyBase
    {
        private bool _enabled = true;
        private string _value = "";
        private int _count;
        private Color _color = Colors.LightGray;
        private string _colorHex = "#D3D3D3";

        public bool Enabled { get => _enabled; set => SetField(ref _enabled, value); }
        public string Value { get => _value; set => SetField(ref _value, value ?? ""); }
        public int Count { get => _count; set => SetField(ref _count, value); }

        public Color Color
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

        public Brush SwatchBrush => new SolidColorBrush(Color);
    }

    public sealed class QuickColourHierarchyGroup : NotifyBase
    {
        private bool _enabled = true;
        private string _value = "";
        private int _count;
        private Color _baseColor = Colors.LightGreen;
        private string _baseHex = "#90EE90";
        private string _hueGroupName = "Architecture";
        private bool _useCustomBaseColor;
        private bool _suppressCustomMark;

        public bool Enabled { get => _enabled; set => SetField(ref _enabled, value); }
        public string Value { get => _value; set => SetField(ref _value, value ?? ""); }
        public int Count { get => _count; set => SetField(ref _count, value); }

        public string HueGroupName
        {
            get => _hueGroupName;
            set => SetField(ref _hueGroupName, (value ?? "").Trim());
        }

        public bool UseCustomBaseColor
        {
            get => _useCustomBaseColor;
            set => SetField(ref _useCustomBaseColor, value);
        }

        public Color BaseColor
        {
            get => _baseColor;
            set
            {
                if (SetField(ref _baseColor, value))
                {
                    _baseHex = QuickColourPalette.ToHex(value);
                    OnPropertyChanged(nameof(BaseHex));
                    OnPropertyChanged(nameof(BaseSwatchBrush));

                    if (!_suppressCustomMark)
                    {
                        UseCustomBaseColor = true;
                    }
                }
            }
        }

        public string BaseHex
        {
            get => _baseHex;
            set
            {
                if (QuickColourPalette.TryParseHex(value, out var parsed))
                {
                    _suppressCustomMark = true;
                    try
                    {
                        _baseColor = parsed;
                        _baseHex = QuickColourPalette.ToHex(parsed);
                        OnPropertyChanged();
                        OnPropertyChanged(nameof(BaseColor));
                        OnPropertyChanged(nameof(BaseSwatchBrush));
                    }
                    finally
                    {
                        _suppressCustomMark = false;
                    }

                    UseCustomBaseColor = true;
                }
                else
                {
                    _baseHex = QuickColourPalette.ToHex(_baseColor);
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<QuickColourHierarchyTypeRow> Types { get; } =
            new ObservableCollection<QuickColourHierarchyTypeRow>();

        public int TypeCount => Types?.Count ?? 0;

        public Brush BaseSwatchBrush => new SolidColorBrush(BaseColor);

        public QuickColourHierarchyGroup()
        {
            Types.CollectionChanged += OnTypesChanged;
        }

        internal void SetComputedBaseColor(Color computed)
        {
            _suppressCustomMark = true;
            try
            {
                _baseColor = computed;
                _baseHex = QuickColourPalette.ToHex(computed);
                OnPropertyChanged(nameof(BaseHex));
                OnPropertyChanged(nameof(BaseColor));
                OnPropertyChanged(nameof(BaseSwatchBrush));
            }
            finally
            {
                _suppressCustomMark = false;
            }

            UseCustomBaseColor = false;
        }

        private void OnTypesChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged(nameof(TypeCount));
        }
    }
}
