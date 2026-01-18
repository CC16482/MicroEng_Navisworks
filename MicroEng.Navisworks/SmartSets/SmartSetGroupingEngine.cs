using System;
using System.Collections.Generic;
using System.Linq;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.SmartSets
{
    internal static class SmartSetGroupingEngine
    {
        public static List<SmartGroupRow> BuildGroups(
            ScrapeSession session,
            string groupByCategory,
            string groupByProperty,
            bool useThenBy,
            string thenByCategory,
            string thenByProperty,
            int minCount,
            bool includeBlanks)
        {
            if (session?.RawEntries == null)
            {
                return new List<SmartGroupRow>();
            }

            var map1 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var map2 = useThenBy ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : null;

            foreach (var entry in session.RawEntries)
            {
                var item = entry.ItemPath ?? "";
                if (string.IsNullOrWhiteSpace(item))
                {
                    continue;
                }

                if (IsMatch(entry, groupByCategory, groupByProperty))
                {
                    map1[item] = Normalize(entry.Value);
                }

                if (useThenBy && IsMatch(entry, thenByCategory, thenByProperty))
                {
                    map2[item] = Normalize(entry.Value);
                }
            }

            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var kvp in map1)
            {
                var item = kvp.Key;
                var v1 = kvp.Value ?? "";

                if (!includeBlanks && string.IsNullOrWhiteSpace(v1))
                {
                    continue;
                }

                string key;

                if (useThenBy)
                {
                    map2.TryGetValue(item, out var v2);
                    v2 = v2 ?? "";
                    if (!includeBlanks && string.IsNullOrWhiteSpace(v2))
                    {
                        continue;
                    }

                    key = $"{v1}\u001F{v2}";
                }
                else
                {
                    key = v1;
                }

                counts.TryGetValue(key, out var count);
                counts[key] = count + 1;
            }

            var rows = new List<SmartGroupRow>();

            foreach (var kvp in counts.OrderByDescending(k => k.Value))
            {
                if (kvp.Value < minCount)
                {
                    continue;
                }

                if (!useThenBy)
                {
                    rows.Add(new SmartGroupRow { Value1 = kvp.Key, Value2 = "", Count = kvp.Value });
                }
                else
                {
                    var parts = kvp.Key.Split(new[] { '\u001F' }, 2);
                    rows.Add(new SmartGroupRow
                    {
                        Value1 = parts.Length > 0 ? parts[0] : "",
                        Value2 = parts.Length > 1 ? parts[1] : "",
                        Count = kvp.Value
                    });
                }
            }

            return rows;
        }

        private static bool IsMatch(RawEntry entry, string category, string property)
        {
            return string.Equals(entry.Category ?? "", category ?? "", StringComparison.OrdinalIgnoreCase)
                && string.Equals(entry.Name ?? "", property ?? "", StringComparison.OrdinalIgnoreCase);
        }

        private static string Normalize(string value)
        {
            return value?.Trim() ?? "";
        }
    }
}
