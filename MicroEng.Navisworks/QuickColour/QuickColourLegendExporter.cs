using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace MicroEng.Navisworks.QuickColour
{
    public static class QuickColourLegendExporter
    {
        public static void ExportCsv(
            string path,
            string profileName,
            string category,
            string property,
            QuickColourScope scope,
            IEnumerable<QuickColourValueRow> rows)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("Profile,Category,Property,Scope,Value,Hex,Count,Enabled");

            foreach (var row in rows ?? Enumerable.Empty<QuickColourValueRow>())
            {
                if (row == null)
                {
                    continue;
                }

                sb.Append(EscapeCsv(profileName));
                sb.Append(',');
                sb.Append(EscapeCsv(category));
                sb.Append(',');
                sb.Append(EscapeCsv(property));
                sb.Append(',');
                sb.Append(EscapeCsv(scope.ToString()));
                sb.Append(',');
                sb.Append(EscapeCsv(row.Value));
                sb.Append(',');
                sb.Append(EscapeCsv(row.ColorHex));
                sb.Append(',');
                sb.Append(row.Count.ToString());
                sb.Append(',');
                sb.Append(row.Enabled ? "true" : "false");
                sb.AppendLine();
            }

            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
        }

        public static void ExportJson(
            string path,
            string profileName,
            string category,
            string property,
            QuickColourScope scope,
            IEnumerable<QuickColourValueRow> rows)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            var export = new QuickColourLegendExport
            {
                ProfileName = profileName ?? "",
                Category = category ?? "",
                Property = property ?? "",
                Scope = scope.ToString(),
                ExportedUtc = DateTime.UtcNow,
                Rows = (rows ?? Enumerable.Empty<QuickColourValueRow>())
                    .Where(r => r != null)
                    .Select(r => new QuickColourLegendRow
                    {
                        Value = r.Value ?? "",
                        ColorHex = r.ColorHex ?? "",
                        Count = r.Count,
                        Enabled = r.Enabled
                    })
                    .ToList()
            };

            var ser = new DataContractJsonSerializer(typeof(QuickColourLegendExport));
            using (var ms = new MemoryStream())
            {
                ser.WriteObject(ms, export);
                var json = Encoding.UTF8.GetString(ms.ToArray());
                File.WriteAllText(path, json, new UTF8Encoding(false));
            }
        }

        private static string EscapeCsv(string value)
        {
            var s = value ?? "";
            if (s.Contains("\"") || s.Contains(",") || s.Contains("\n") || s.Contains("\r"))
            {
                s = "\"" + s.Replace("\"", "\"\"") + "\"";
            }

            return s;
        }

        [DataContract]
        private sealed class QuickColourLegendExport
        {
            [DataMember(Name = "profileName")] public string ProfileName { get; set; } = "";
            [DataMember(Name = "category")] public string Category { get; set; } = "";
            [DataMember(Name = "property")] public string Property { get; set; } = "";
            [DataMember(Name = "scope")] public string Scope { get; set; } = "";
            [DataMember(Name = "exportedUtc")] public DateTime ExportedUtc { get; set; }
            [DataMember(Name = "rows")] public List<QuickColourLegendRow> Rows { get; set; } = new List<QuickColourLegendRow>();
        }

        [DataContract]
        private sealed class QuickColourLegendRow
        {
            [DataMember(Name = "value")] public string Value { get; set; } = "";
            [DataMember(Name = "colorHex")] public string ColorHex { get; set; } = "";
            [DataMember(Name = "count")] public int Count { get; set; }
            [DataMember(Name = "enabled")] public bool Enabled { get; set; }
        }
    }
}
