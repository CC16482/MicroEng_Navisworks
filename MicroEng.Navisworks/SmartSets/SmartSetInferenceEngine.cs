using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;

namespace MicroEng.Navisworks.SmartSets
{
    public sealed class SmartSetSuggestion
    {
        public string Category { get; set; }
        public string Property { get; set; }
        public SmartSetOperator Operator { get; set; }
        public string Value { get; set; }
        public int MatchCount { get; set; }
        public int TotalCount { get; set; }

        public string Display
        {
            get
            {
                if (string.IsNullOrWhiteSpace(Category) && string.IsNullOrWhiteSpace(Property))
                {
                    return Value ?? "";
                }

                var valueLabel = string.IsNullOrWhiteSpace(Value) ? "(blank)" : Value;
                return $"{Category} / {Property} {Operator} {valueLabel} ({MatchCount}/{TotalCount})";
            }
        }
    }

    public static class SmartSetInferenceEngine
    {
        public static List<SmartSetSuggestion> AnalyzeSelection(ModelItemCollection selection, int maxSuggestions)
        {
            var suggestions = new List<SmartSetSuggestion>();

            if (selection == null || selection.Count == 0)
            {
                return suggestions;
            }

            var totals = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var values = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in selection)
            {
                if (item?.PropertyCategories == null)
                {
                    continue;
                }

                foreach (PropertyCategory category in item.PropertyCategories)
                {
                    if (category == null)
                    {
                        continue;
                    }

                    foreach (DataProperty prop in category.Properties)
                    {
                        if (prop == null)
                        {
                            continue;
                        }

                        var catName = category.DisplayName ?? category.Name ?? "";
                        var propName = prop.DisplayName ?? prop.Name ?? "";
                        if (string.IsNullOrWhiteSpace(catName) || string.IsNullOrWhiteSpace(propName))
                        {
                            continue;
                        }

                        var key = $"{catName}::{propName}";
                        totals.TryGetValue(key, out var totalCount);
                        totals[key] = totalCount + 1;

                        var displayValue = SafePropertyValue(prop);

                        if (!values.TryGetValue(key, out var valueCounts))
                        {
                            valueCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                            values[key] = valueCounts;
                        }

                        valueCounts.TryGetValue(displayValue, out var valCount);
                        valueCounts[displayValue] = valCount + 1;
                    }
                }
            }

            var selectionCount = selection.Count;

            foreach (var kvp in values)
            {
                var key = kvp.Key;
                var parts = key.Split(new[] { "::" }, 2, StringSplitOptions.None);
                var cat = parts.Length > 0 ? parts[0] : "";
                var prop = parts.Length > 1 ? parts[1] : "";

                if (!totals.TryGetValue(key, out var totalCount))
                {
                    totalCount = selectionCount;
                }

                var ordered = kvp.Value.OrderByDescending(v => v.Value).ToList();
                if (ordered.Count == 0)
                {
                    continue;
                }

                var top = ordered[0];
                var coverage = (double)top.Value / Math.Max(1, selectionCount);
                if (coverage < 0.75)
                {
                    continue;
                }

                suggestions.Add(new SmartSetSuggestion
                {
                    Category = cat,
                    Property = prop,
                    Operator = string.IsNullOrWhiteSpace(top.Key) ? SmartSetOperator.Defined : SmartSetOperator.Equals,
                    Value = top.Key,
                    MatchCount = top.Value,
                    TotalCount = selectionCount
                });
            }

            return suggestions
                .OrderByDescending(s => s.MatchCount)
                .ThenBy(s => s.Category)
                .ThenBy(s => s.Property)
                .Take(Math.Max(1, maxSuggestions))
                .ToList();
        }

        private static string SafePropertyValue(DataProperty prop)
        {
            try
            {
                if (prop.Value == null)
                {
                    return "";
                }

                if (prop.Value.IsDisplayString)
                {
                    return prop.Value.ToDisplayString() ?? "";
                }

                return prop.Value.ToString() ?? "";
            }
            catch
            {
                return "";
            }
        }
    }
}
