using System.Windows.Media;

namespace MicroEng.Navisworks.QuickColour
{
    public sealed class QuickColourHueGroup : NotifyBase
    {
        private bool _enabled = true;
        private string _name = "Architecture";
        private Color _hueColor = Colors.Green;
        private string _hueHex = "#00A000";

        public bool Enabled
        {
            get => _enabled;
            set => SetField(ref _enabled, value);
        }

        public string Name
        {
            get => _name;
            set => SetField(ref _name, (value ?? "").Trim());
        }

        public string HueHex
        {
            get => _hueHex;
            set
            {
                var s = (value ?? "").Trim();
                if (QuickColourPalette.TryParseHex(s, out var c))
                {
                    _hueColor = c;
                    _hueHex = QuickColourPalette.ToHex(c);
                    OnPropertyChanged(nameof(HueHex));
                    OnPropertyChanged(nameof(HueColor));
                    OnPropertyChanged(nameof(SwatchBrush));
                }
                else
                {
                    _hueHex = _hueHex ?? "#00A000";
                    OnPropertyChanged(nameof(HueHex));
                }
            }
        }

        public Color HueColor => _hueColor;

        public Brush SwatchBrush => new SolidColorBrush(_hueColor);
    }
}
