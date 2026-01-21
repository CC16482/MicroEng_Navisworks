using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Media;

namespace MicroEng.Navisworks.Colour
{
    internal static class ColourPaletteGenerator
    {
        public static Color ParseHexOrDefault(string hex, Color fallback)
        {
            if (TryParseHex(hex, out var c))
            {
                return c;
            }

            return fallback;
        }

        public static bool TryParseHex(string hex, out Color color)
        {
            color = default;

            if (string.IsNullOrWhiteSpace(hex))
            {
                return false;
            }

            var s = hex.Trim();
            if (s.StartsWith("#", StringComparison.Ordinal))
            {
                s = s.Substring(1);
            }

            if (s.Length == 6)
            {
                if (byte.TryParse(s.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r)
                    && byte.TryParse(s.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g)
                    && byte.TryParse(s.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
                {
                    color = Color.FromRgb(r, g, b);
                    return true;
                }

                return false;
            }

            if (s.Length == 8)
            {
                if (byte.TryParse(s.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var a)
                    && byte.TryParse(s.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r)
                    && byte.TryParse(s.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g)
                    && byte.TryParse(s.Substring(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
                {
                    color = Color.FromArgb(a, r, g, b);
                    return true;
                }

                return false;
            }

            return false;
        }

        public static string ToHexRgb(Color c)
        {
            return $"#{c.R:X2}{c.G:X2}{c.B:X2}";
        }

        public static Color GetCustomHueColor(
            Color baseColor,
            string valueKey,
            int ordinalIndex,
            int totalCount,
            bool stableColors,
            string seed)
        {
            RgbToHsl(baseColor, out var baseH, out var baseS, out var baseL);

            baseS = Clamp(baseS, 0.35, 0.95);
            baseL = Clamp(baseL, 0.25, 0.85);

            double hue;
            if (stableColors)
            {
                var h = StableHashInt($"{seed}|{valueKey}");
                hue = (baseH + (h % 360)) % 360.0;
            }
            else
            {
                var step = 360.0 / Math.Max(1, totalCount);
                hue = (baseH + (ordinalIndex * step)) % 360.0;
            }

            return HslToRgb(hue, baseS, baseL);
        }

        private static int StableHashInt(string s)
        {
            using (var md5 = MD5.Create())
            {
                var bytes = md5.ComputeHash(Encoding.UTF8.GetBytes(s ?? ""));
                var v = BitConverter.ToInt32(bytes, 0);
                return v == int.MinValue ? int.MaxValue : Math.Abs(v);
            }
        }

        private static double Clamp(double v, double min, double max)
        {
            return v < min ? min : (v > max ? max : v);
        }

        public static void RgbToHsl(Color c, out double h, out double s, out double l)
        {
            var r = c.R / 255.0;
            var g = c.G / 255.0;
            var b = c.B / 255.0;

            var max = Math.Max(r, Math.Max(g, b));
            var min = Math.Min(r, Math.Min(g, b));
            var delta = max - min;

            l = (max + min) / 2.0;

            if (delta < 1e-9)
            {
                h = 0;
                s = 0;
                return;
            }

            s = delta / (1.0 - Math.Abs(2.0 * l - 1.0));

            if (Math.Abs(max - r) < 1e-9)
            {
                h = 60.0 * (((g - b) / delta) % 6.0);
            }
            else if (Math.Abs(max - g) < 1e-9)
            {
                h = 60.0 * (((b - r) / delta) + 2.0);
            }
            else
            {
                h = 60.0 * (((r - g) / delta) + 4.0);
            }

            if (h < 0)
            {
                h += 360.0;
            }
        }

        public static Color HslToRgb(double h, double s, double l)
        {
            h = ((h % 360.0) + 360.0) % 360.0;
            s = Clamp(s, 0, 1);
            l = Clamp(l, 0, 1);

            var c = (1.0 - Math.Abs(2.0 * l - 1.0)) * s;
            var x = c * (1.0 - Math.Abs(((h / 60.0) % 2.0) - 1.0));
            var m = l - c / 2.0;

            double rp;
            double gp;
            double bp;

            if (h < 60)
            {
                rp = c;
                gp = x;
                bp = 0;
            }
            else if (h < 120)
            {
                rp = x;
                gp = c;
                bp = 0;
            }
            else if (h < 180)
            {
                rp = 0;
                gp = c;
                bp = x;
            }
            else if (h < 240)
            {
                rp = 0;
                gp = x;
                bp = c;
            }
            else if (h < 300)
            {
                rp = x;
                gp = 0;
                bp = c;
            }
            else
            {
                rp = c;
                gp = 0;
                bp = x;
            }

            var r = (byte)Math.Round((rp + m) * 255.0);
            var g = (byte)Math.Round((gp + m) * 255.0);
            var b = (byte)Math.Round((bp + m) * 255.0);

            return Color.FromRgb(r, g, b);
        }
    }
}
