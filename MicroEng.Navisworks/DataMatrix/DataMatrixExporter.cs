using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace MicroEng.Navisworks
{
    internal class DataMatrixExporter
    {
        public void ExportCsv(string path, IEnumerable<DataMatrixAttributeDefinition> columns, IEnumerable<DataMatrixRow> rows, ScrapeSession session, DataMatrixViewPreset preset)
        {
            var colList = columns.ToList();
            var sb = new StringBuilder();
            sb.AppendLine($"# Profile: {session.ProfileName}");
            sb.AppendLine($"# Scope: {session.ScopeDescription} at {session.Timestamp}");
            sb.AppendLine($"# View: {preset?.Name ?? "None"}");
            sb.AppendLine($"# Exported: {DateTime.Now}");
            sb.AppendLine(string.Join(",", colList.Select(c => Escape(c.DisplayName))));

            foreach (var row in rows)
            {
                var vals = new List<string>();
                foreach (var col in colList)
                {
                    row.Values.TryGetValue(col.Id, out var val);
                    vals.Add(Escape(val));
                }
                sb.AppendLine(string.Join(",", vals));
            }
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        public void ExportJsonl(
            string path,
            IEnumerable<DataMatrixAttributeDefinition> columns,
            IEnumerable<DataMatrixRow> rows,
            ScrapeSession session,
            DataMatrixViewPreset preset,
            DataMatrixJsonlMode mode)
        {
            var gzip = path.EndsWith(".gz", StringComparison.OrdinalIgnoreCase);
            var colList = columns.ToList();

            using (var fs = File.Create(path))
            using (var stream = gzip ? (Stream)new GZipStream(fs, CompressionLevel.Optimal) : fs)
            using (var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
            {
                foreach (var row in rows)
                {
                    if (mode == DataMatrixJsonlMode.ItemDocuments)
                    {
                        writer.WriteLine(BuildItemDocJson(row, colList, session, preset));
                    }
                    else
                    {
                        foreach (var col in colList)
                        {
                            if (!row.Values.TryGetValue(col.Id, out var val) || val == null) continue;
                            writer.WriteLine(BuildRawRowJson(row, col, val, session, preset));
                        }
                    }
                }
            }
        }

        private string Escape(object value)
        {
            if (value == null) return "";
            var s = value.ToString() ?? "";
            if (s.Contains(",") || s.Contains("\""))
            {
                s = "\"" + s.Replace("\"", "\"\"") + "\"";
            }
            return s;
        }

        private string BuildItemDocJson(DataMatrixRow row, List<DataMatrixAttributeDefinition> cols, ScrapeSession session, DataMatrixViewPreset preset)
        {
            var sb = new StringBuilder();
            sb.Append('{');
            AppendKV(sb, "profile", session?.ProfileName);
            sb.Append(',');
            AppendKV(sb, "scope", session?.ScopeDescription);
            sb.Append(',');
            AppendKV(sb, "sessionTimestamp", session?.Timestamp.ToString("o"));
            sb.Append(',');
            AppendKV(sb, "view", preset?.Name ?? "(Default)");
            sb.Append(',');
            AppendKV(sb, "exportedAt", DateTime.Now.ToString("o"));
            sb.Append(',');
            AppendKV(sb, "itemKey", row?.ItemKey);
            sb.Append(',');
            AppendKV(sb, "itemPath", row?.ElementDisplayName);
            sb.Append(',');
            sb.Append("\"properties\":{");

            var first = true;
            foreach (var col in cols)
            {
                if (!row.Values.TryGetValue(col.Id, out var val)) continue;

                if (!first) sb.Append(',');
                first = false;

                sb.Append('\"').Append(JsonEscape(col.Id)).Append("\":");
                AppendJsonValue(sb, val);
            }

            sb.Append("}}");
            return sb.ToString();
        }

        private string BuildRawRowJson(DataMatrixRow row, DataMatrixAttributeDefinition col, object val, ScrapeSession session, DataMatrixViewPreset preset)
        {
            var sb = new StringBuilder();
            sb.Append('{');
            AppendKV(sb, "profile", session?.ProfileName);
            sb.Append(',');
            AppendKV(sb, "scope", session?.ScopeDescription);
            sb.Append(',');
            AppendKV(sb, "sessionTimestamp", session?.Timestamp.ToString("o"));
            sb.Append(',');
            AppendKV(sb, "view", preset?.Name ?? "(Default)");
            sb.Append(',');
            AppendKV(sb, "exportedAt", DateTime.Now.ToString("o"));
            sb.Append(',');
            AppendKV(sb, "itemKey", row?.ItemKey);
            sb.Append(',');
            AppendKV(sb, "itemPath", row?.ElementDisplayName);
            sb.Append(',');
            AppendKV(sb, "columnId", col?.Id);
            sb.Append(',');
            AppendKV(sb, "category", col?.Category);
            sb.Append(',');
            AppendKV(sb, "property", col?.PropertyName);
            sb.Append(',');
            sb.Append("\"value\":");
            AppendJsonValue(sb, val);
            sb.Append(',');
            AppendKV(sb, "valueType", val?.GetType().Name);
            sb.Append('}');
            return sb.ToString();
        }

        private void AppendKV(StringBuilder sb, string key, string value)
        {
            sb.Append('\"').Append(JsonEscape(key)).Append("\":");
            if (value == null) sb.Append("null");
            else sb.Append('\"').Append(JsonEscape(value)).Append('\"');
        }

        private void AppendJsonValue(StringBuilder sb, object value)
        {
            if (value == null)
            {
                sb.Append("null");
                return;
            }

            switch (value)
            {
                case bool b:
                    sb.Append(b ? "true" : "false");
                    return;
                case int i:
                    sb.Append(i.ToString(CultureInfo.InvariantCulture));
                    return;
                case double d:
                    sb.Append(d.ToString(CultureInfo.InvariantCulture));
                    return;
                case float f:
                    sb.Append(f.ToString(CultureInfo.InvariantCulture));
                    return;
                case DateTime dt:
                    sb.Append('\"').Append(JsonEscape(dt.ToString("o"))).Append('\"');
                    return;
                default:
                    sb.Append('\"').Append(JsonEscape(value.ToString())).Append('\"');
                    return;
            }
        }

        private string JsonEscape(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;

            var sb = new StringBuilder(s.Length + 16);
            foreach (var ch in s)
            {
                switch (ch)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '\"': sb.Append("\\\""); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (ch < 32) sb.Append("\\u").Append(((int)ch).ToString("x4"));
                        else sb.Append(ch);
                        break;
                }
            }

            return sb.ToString();
        }
    }
}
