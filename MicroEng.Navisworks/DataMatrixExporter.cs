using System;
using System.Collections.Generic;
using System.IO;
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
    }
}
