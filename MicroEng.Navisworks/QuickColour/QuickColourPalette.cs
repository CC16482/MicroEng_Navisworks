using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Media;
using MicroEng.Navisworks.QuickColour.Profiles;
using MediaColor = System.Windows.Media.Color;

namespace MicroEng.Navisworks.QuickColour
{
    public static class QuickColourPalette
    {
        private static readonly IReadOnlyDictionary<MicroEngPaletteKind, string[]> FixedPalettes =
            new Dictionary<MicroEngPaletteKind, string[]>
            {
                {
                    MicroEngPaletteKind.Beach,
                    new[]
                    {
                        "#264653", "#2A9D8F", "#E9C46A", "#F4A261", "#E76F51"
                    }
                },
                {
                    MicroEngPaletteKind.OceanBreeze,
                    new[]
                    {
                        "#E63946", "#F1FAEE", "#A8DADC", "#457B9D", "#1D3557"
                    }
                },
                {
                    MicroEngPaletteKind.Vibrant,
                    new[]
                    {
                        "#FFBE0B", "#FB5607", "#FF006E", "#8338EC", "#3A86FF"
                    }
                },
                {
                    MicroEngPaletteKind.Pastel,
                    new[]
                    {
                        "#FFADAD", "#FFD6A5", "#FDFFB6", "#CAFFBF", "#9BF6FF",
                        "#A0C4FF", "#BDB2FF", "#FFC6FF", "#FFFFFC"
                    }
                },
                {
                    MicroEngPaletteKind.Autumn,
                    new[]
                    {
                        "#003049", "#D62828", "#F77F00", "#FCBF49", "#EAE2B7"
                    }
                },
                {
                    MicroEngPaletteKind.RedSunset,
                    new[]
                    {
                        "#03071E", "#370617", "#6A040F", "#9D0208", "#D00000",
                        "#DC2F02", "#E85D04", "#F48C06", "#FAA307", "#FFBA08"
                    }
                },
                {
                    MicroEngPaletteKind.ForestHues,
                    new[]
                    {
                        "#DAD7CD", "#A3B18A", "#588157", "#3A5A40", "#344E41"
                    }
                },
                {
                    MicroEngPaletteKind.PurpleRaindrops,
                    new[]
                    {
                        "#F72585", "#B5179E", "#7209B7", "#560BAD", "#480CA8",
                        "#3A0CA3", "#3F37C9", "#4361EE", "#4895EF", "#4CC9F0"
                    }
                },
                {
                    MicroEngPaletteKind.LightSteel,
                    new[]
                    {
                        "#F8F9FA", "#E9ECEF", "#DEE2E6", "#CED4DA", "#ADB5BD",
                        "#6C757D", "#495057", "#343A40", "#212529"
                    }
                },
                {
                    MicroEngPaletteKind.EarthyBrown,
                    new[]
                    {
                        "#EDE0D4", "#E6CCB2", "#DDB892", "#B08968", "#7F5539",
                        "#9C6644"
                    }
                },
                {
                    MicroEngPaletteKind.EarthyGreen,
                    new[]
                    {
                        "#CAD2C5", "#84A98C", "#52796F", "#354F52", "#2F3E46"
                    }
                },
                {
                    MicroEngPaletteKind.WarmNeutrals1,
                    new[]
                    {
                        "#CB997E", "#EDDCD2", "#FFF1E6", "#F0EFEB", "#DDBEA9",
                        "#A5A58D", "#B7B7A4"
                    }
                },
                {
                    MicroEngPaletteKind.WarmNeutrals2,
                    new[]
                    {
                        "#582F0E", "#7F4F24", "#936639", "#A68A64", "#B6AD90",
                        "#C2C5AA", "#A4AC86", "#656D4A", "#414833", "#333D29"
                    }
                },
                {
                    MicroEngPaletteKind.CandyPop,
                    new[]
                    {
                        "#9B5DE5", "#F15BB5", "#FEE440", "#00BBF9", "#00F5D4"
                    }
                }
            };

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

            if (s.Length != 6 && s.Length != 8)
            {
                return false;
            }

            if (!uint.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value))
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
            if (style == QuickColourPaletteStyle.Pastel)
            {
                return GeneratePalette(count, MicroEngPaletteKind.Pastel);
            }

            var colors = new List<MediaColor>(Math.Max(0, count));
            if (count <= 0)
            {
                return colors;
            }

            var sat = GetRecommendedSaturation(style);
            var light = 0.55;

            for (var i = 0; i < count; i++)
            {
                var h = (double)i / Math.Max(1, count);
                colors.Add(HslToRgb(h, sat, light));
            }

            return colors;
        }

        public static List<MediaColor> GeneratePalette(int count, MicroEngPaletteKind kind)
        {
            if (count <= 0)
            {
                return new List<MediaColor>();
            }

            if (kind == MicroEngPaletteKind.Deep)
            {
                return GeneratePalette(count, QuickColourPaletteStyle.Deep);
            }

            if (FixedPalettes.TryGetValue(kind, out var hexes))
            {
                return BuildFixedPalette(hexes, count);
            }

            return GeneratePalette(count, QuickColourPaletteStyle.Deep);
        }

        private static List<MediaColor> BuildFixedPalette(IReadOnlyList<string> hexes, int count)
        {
            var baseColors = new List<MediaColor>();
            if (hexes != null)
            {
                foreach (var hex in hexes)
                {
                    if (TryParseHex(hex, out var color))
                    {
                        baseColors.Add(color);
                    }
                }
            }

            if (baseColors.Count == 0)
            {
                return GeneratePalette(count, QuickColourPaletteStyle.Deep);
            }

            if (count <= baseColors.Count)
            {
                return SamplePalette(baseColors, count);
            }

            var result = new List<MediaColor>(count);
            result.AddRange(baseColors);

            var extrasNeeded = count - baseColors.Count;
            if (extrasNeeded > 0)
            {
                result.AddRange(BuildShadeVariations(baseColors, extrasNeeded));
            }

            return result;
        }

        private static List<MediaColor> SamplePalette(IReadOnlyList<MediaColor> colors, int count)
        {
            var sampled = new List<MediaColor>(Math.Max(0, count));
            if (colors == null || colors.Count == 0 || count <= 0)
            {
                return sampled;
            }

            if (count == 1)
            {
                sampled.Add(colors[0]);
                return sampled;
            }

            for (int i = 0; i < count; i++)
            {
                double t = (double)i / (count - 1);
                int idx = (int)Math.Round(t * (colors.Count - 1));
                if (idx < 0) idx = 0;
                if (idx >= colors.Count) idx = colors.Count - 1;
                sampled.Add(colors[idx]);
            }

            return sampled;
        }

        private static List<MediaColor> BuildShadeVariations(IReadOnlyList<MediaColor> baseColors, int count)
        {
            var extras = new List<MediaColor>(Math.Max(0, count));
            if (count <= 0 || baseColors == null || baseColors.Count == 0)
            {
                return extras;
            }

            int baseCount = baseColors.Count;
            int levels = (int)Math.Ceiling((double)count / baseCount);
            var offsets = BuildShadeOffsets(levels);

            foreach (var offset in offsets)
            {
                foreach (var color in baseColors)
                {
                    if (extras.Count >= count)
                    {
                        return extras;
                    }

                    extras.Add(ShiftLightness(color, offset));
                }
            }

            return extras;
        }

        private static List<double> BuildShadeOffsets(int levels)
        {
            var offsets = new List<double>(Math.Max(0, levels));
            if (levels <= 0)
            {
                return offsets;
            }

            const double baseOffset = 0.18;
            const double step = 0.08;

            for (int i = 0; i < levels; i++)
            {
                double magnitude = baseOffset + step * (i / 2);
                if (magnitude > 0.35)
                {
                    magnitude = 0.35;
                }

                offsets.Add(i % 2 == 0 ? magnitude : -magnitude);
            }

            return offsets;
        }

        private static MediaColor ShiftLightness(MediaColor color, double offset)
        {
            RgbToHsl(color, out var h, out var s, out var l);
            var newL = Clamp01(l + offset);
            newL = Math.Max(0.12, Math.Min(0.92, newL));
            return HslToRgb(h, s, newL);
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

        public static void AssignPalette(IList<QuickColourValueRow> rows, QuickColourPaletteStyle style)
        {
            if (rows == null || rows.Count == 0)
            {
                return;
            }

            var list = rows.Where(r => r != null).ToList();
            if (list.Count == 0)
            {
                return;
            }

            var palette = GeneratePalette(list.Count, style);
            for (int i = 0; i < list.Count; i++)
            {
                list[i].Color = palette[i];
            }
        }

        public static void AssignStableColors(IList<QuickColourValueRow> rows, QuickColourPaletteStyle style, string seed)
        {
            if (rows == null || rows.Count == 0)
            {
                return;
            }

            var list = rows.Where(r => r != null).ToList();
            if (list.Count == 0)
            {
                return;
            }

            var sat = GetRecommendedSaturation(style);
            var light = style == QuickColourPaletteStyle.Pastel ? 0.78 : 0.55;
            var seedText = seed ?? "";

            foreach (var row in list)
            {
                var h = HashToHue01(row.Value, seedText);
                row.Color = FromHsl01(h, sat, light);
            }
        }

        public static void AssignPalette(
            IList<QuickColourValueRow> rows,
            IList<MediaColor> palette,
            bool stableColors,
            string seed)
        {
            if (rows == null || rows.Count == 0 || palette == null || palette.Count == 0)
            {
                return;
            }

            var list = rows.Where(r => r != null).ToList();
            if (list.Count == 0)
            {
                return;
            }

            if (stableColors)
            {
                var seedText = seed ?? "";
                int paletteCount = palette.Count;
                foreach (var row in list)
                {
                    var idx = (int)Math.Floor(GetStableUnit01(row.Value, seedText) * paletteCount);
                    if (idx >= paletteCount)
                    {
                        idx = paletteCount - 1;
                    }

                    row.Color = palette[idx];
                }

                return;
            }

            for (int i = 0; i < list.Count && i < palette.Count; i++)
            {
                list[i].Color = palette[i];
            }
        }

        public static void AssignShades(IList<QuickColourValueRow> rows, MediaColor baseColor, bool stableColors, string seed)
        {
            if (rows == null || rows.Count == 0)
            {
                return;
            }

            var list = rows.Where(r => r != null).ToList();
            if (list.Count == 0)
            {
                return;
            }

            if (list.Count == 1)
            {
                list[0].Color = baseColor;
                return;
            }

            RgbToHsl(baseColor, out var h, out var s, out var l);
            var (minL, maxL) = ComputeShadeRange(l);

            if (stableColors)
            {
                var seedText = seed ?? "";
                foreach (var row in list)
                {
                    var t = GetStableUnit01(row.Value, seedText);
                    var shadeL = minL + (maxL - minL) * t;
                    row.Color = FromHsl01(h, s, shadeL);
                }

                return;
            }

            var ramp = GenerateHslRamp(h, s, minL, maxL, list.Count);
            for (int i = 0; i < list.Count; i++)
            {
                list[i].Color = ramp[i];
            }
        }

        public static double GetHue01(MediaColor c)
        {
            RgbToHsl(c, out var h, out _, out _);
            return h;
        }

        internal static double GetStableUnit01(string value, string seed)
        {
            return HashToHue01(value, seed);
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

        private static (double minL, double maxL) ComputeShadeRange(double baseL)
        {
            double minL = baseL - 0.5;
            double maxL = baseL + 0.5;

            minL = Clamp01(minL);
            maxL = Clamp01(maxL);

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

        private static double HashToHue01(string value, string seed)
        {
            unchecked
            {
                uint hash = 2166136261;
                var text = (value ?? "") + "|" + (seed ?? "");
                foreach (var ch in text)
                {
                    hash ^= ch;
                    hash *= 16777619;
                }

                return (hash % 360u) / 360.0;
            }
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
