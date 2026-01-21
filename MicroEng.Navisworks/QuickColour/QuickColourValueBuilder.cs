using System;
using System.Collections.Generic;
using System.Linq;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.QuickColour
{
    internal static class QuickColourValueBuilder
    {
        internal static List<QuickColourValueRow> BuildValues(
            ScrapeSession session,
            string category,
            string property,
            ISet<string> allowedItemKeys)
        {
            var rows = new List<QuickColourValueRow>();
            if (session == null || string.IsNullOrWhiteSpace(category) || string.IsNullOrWhiteSpace(property))
            {
                return rows;
            }

            var counts = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (var entry in session.RawEntries ?? Enumerable.Empty<RawEntry>())
            {
                if (entry == null)
                {
                    continue;
                }

                if (!string.Equals(entry.Category ?? "", category ?? "", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.Equals(entry.Name ?? "", property ?? "", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (allowedItemKeys != null)
                {
                    var key = entry.ItemKey ?? "";
                    if (!allowedItemKeys.Contains(key))
                    {
                        continue;
                    }
                }

                var value = entry.Value ?? "";
                if (counts.TryGetValue(value, out var current))
                {
                    counts[value] = current + 1;
                }
                else
                {
                    counts[value] = 1;
                }
            }

            foreach (var kvp in counts)
            {
                rows.Add(new QuickColourValueRow
                {
                    Value = kvp.Key ?? "",
                    Count = kvp.Value,
                    Enabled = true
                });
            }

            return rows
                .OrderByDescending(r => r.Count)
                .ThenBy(r => r.Value, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }
}
