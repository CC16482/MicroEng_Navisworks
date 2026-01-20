using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Media;
using MediaColor = System.Windows.Media.Color;

namespace MicroEng.Navisworks.QuickColour
{
    public static class QuickColourPalette
    {
        public static bool TryParseHex(string hex, out MediaColor color)
        {
            color = Colors.Transparent;
            if (string.IsNullOrWhiteSpace(hex))
            {
                return false;
            }

            var s = hex.Trim();
            if (s.StartsWith("#", StringComparison.Ordinal))
            {
                s = s.Substring(1);
            }

            if (s.Length != 6)
            {
                return false;
            }

            if (!int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value))
            {
                return false;
            }

            var r = (byte)((value >> 16) & 0xFF);
            var g = (byte)((value >> 8) & 0xFF);
            var b = (byte)(value & 0xFF);
            color = MediaColor.FromRgb(r, g, b);
            return true;
        }

        public static string ToHex(MediaColor color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        public static Autodesk.Navisworks.Api.Color ToNavisworksColor(MediaColor color)
        {
            return Autodesk.Navisworks.Api.Color.FromByteRGB(color.R, color.G, color.B);
        }

        public static List<MediaColor> GeneratePalette(int count, QuickColourPaletteStyle style)
        {
            var colors = new List<MediaColor>(Math.Max(0, count));
            if (count <= 0)
            {
                return colors;
            }

            var sat = GetRecommendedSaturation(style);
            var light = style == QuickColourPaletteStyle.Pastel ? 0.78 : 0.55;

            for (var i = 0; i < count; i++)
            {
                var h = (double)i / Math.Max(1, count);
                colors.Add(HslToRgb(h, sat, light));
            }

            return colors;
        }

        public static List<MediaColor> GenerateShades(MediaColor baseColor, int count, QuickColourPaletteStyle style, double spread01)
        {
            if (count <= 0)
            {
                return new List<MediaColor>();
            }

            RgbToHsl(baseColor, out var h, out var s, out var l);

            if (style == QuickColourPaletteStyle.Pastel)
            {
                s = Math.Min(s, 0.35);
                l = Math.Max(l, 0.70);
            }
            else
            {
                s = Math.Max(s, 0.60);
            }

            var minL = Clamp01(l - spread01 * 0.5);
            var maxL = Clamp01(l + spread01 * 0.5);

            minL = Math.Max(minL, 0.18);
            maxL = Math.Min(maxL, 0.92);

            var shades = new List<MediaColor>(count);

            if (count == 1)
            {
                shades.Add(HslToRgb(h, s, l));
                return shades;
            }

            for (int i = 0; i < count; i++)
            {
                double t = (double)i / (count - 1);
                double li = minL + (maxL - minL) * t;
                shades.Add(HslToRgb(h, s, li));
            }

            return shades;
        }

        public static double GetHue01(MediaColor c)
        {
            RgbToHsl(c, out var h, out _, out _);
            return h;
        }

        public static double GetRecommendedSaturation(QuickColourPaletteStyle style)
        {
            return style == QuickColourPaletteStyle.Pastel ? 0.35 : 0.78;
        }

        public static (double minL, double maxL) ComputeCategoryLightnessRange(QuickColourPaletteStyle style, double contrast01)
        {
            contrast01 = Clamp01(contrast01);

            double center = style == QuickColourPaletteStyle.Pastel ? 0.78 : 0.55;
            double maxHalf = style == QuickColourPaletteStyle.Pastel ? 0.18 : 0.33;

            double half = maxHalf * contrast01;

            double minL = Clamp01(center - half);
            double maxL = Clamp01(center + half);

            minL = Math.Max(minL, 0.18);
            maxL = Math.Min(maxL, 0.92);

            if (maxL < minL)
            {
                var t = minL;
                minL = maxL;
                maxL = t;
            }

            return (minL, maxL);
        }

        public static MediaColor FromHsl01(double h01, double s01, double l01)
        {
            return HslToRgb(Clamp01(h01), Clamp01(s01), Clamp01(l01));
        }

        public static List<MediaColor> GenerateHslRamp(double h01, double s01, double minL, double maxL, int count)
        {
            if (count <= 0)
            {
                return new List<MediaColor>();
            }

            minL = Clamp01(minL);
            maxL = Clamp01(maxL);
            if (maxL < minL)
            {
                var t = minL;
                minL = maxL;
                maxL = t;
            }

            var list = new List<MediaColor>(count);

            if (count == 1)
            {
                list.Add(FromHsl01(h01, s01, (minL + maxL) / 2.0));
                return list;
            }

            for (int i = 0; i < count; i++)
            {
                double t = (double)i / (count - 1);
                double l = minL + (maxL - minL) * t;
                list.Add(FromHsl01(h01, s01, l));
            }

            return list;
        }

        private static void RgbToHsl(MediaColor c, out double h, out double s, out double l)
        {
            double r = c.R / 255.0;
            double g = c.G / 255.0;
            double b = c.B / 255.0;

            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            l = (max + min) * 0.5;

            if (Math.Abs(max - min) < 1e-9)
            {
                h = 0;
                s = 0;
                return;
            }

            double d = max - min;
            s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);

            if (Math.Abs(max - r) < 1e-9)
            {
                h = (g - b) / d + (g < b ? 6 : 0);
            }
            else if (Math.Abs(max - g) < 1e-9)
            {
                h = (b - r) / d + 2;
            }
            else
            {
                h = (r - g) / d + 4;
            }

            h /= 6.0;
        }

        private static MediaColor HslToRgb(double h, double s, double l)
        {
            double r, g, b;

            if (Math.Abs(s) < 1e-9)
            {
                r = g = b = l;
            }
            else
            {
                double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
                double p = 2 * l - q;
                r = HueToRgb(p, q, h + 1.0 / 3.0);
                g = HueToRgb(p, q, h);
                b = HueToRgb(p, q, h - 1.0 / 3.0);
            }

            return MediaColor.FromRgb(
                (byte)Math.Round(r * 255),
                (byte)Math.Round(g * 255),
                (byte)Math.Round(b * 255));
        }

        private static double HueToRgb(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1.0 / 6.0) return p + (q - p) * 6 * t;
            if (t < 1.0 / 2.0) return q;
            if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6;
            return p;
        }

        private static double Clamp01(double x)
        {
            if (x < 0) return 0;
            if (x > 1) return 1;
            return x;
        }
    }
}
