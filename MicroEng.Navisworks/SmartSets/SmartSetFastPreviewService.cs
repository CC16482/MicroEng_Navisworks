using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace MicroEng.Navisworks.SmartSets
{
    public interface IDataScrapeSessionView
    {
        string ProfileName { get; }
        IEnumerable<ScrapedPropertyDescriptor> Properties { get; }
        IEnumerable<ScrapedRawEntryView> RawEntries { get; }
    }

    public sealed class ScrapedRawEntryView
    {
        public string ItemPath { get; set; } = "";
        public string Category { get; set; } = "";
        public string Property { get; set; } = "";
        public string Value { get; set; } = "";
    }

    public sealed class SmartSetFastPreviewService
    {
        private readonly Dictionary<string, Dictionary<string, List<string>>> _propertyIndexCache =
            new Dictionary<string, Dictionary<string, List<string>>>(StringComparer.OrdinalIgnoreCase);

        public FastPreviewResult Evaluate(
            IDataScrapeSessionView session,
            IReadOnlyList<SmartSetRule> rules,
            int samplePathsToReturn,
            CancellationToken ct)
        {
            var result = new FastPreviewResult();

            if (session == null)
            {
                result.UsedCache = false;
                result.Notes = "No Data Scraper session available. Run Data Scraper first.";
                return result;
            }

            var enabledRules = rules?.Where(r => r != null && r.Enabled).ToList() ?? new List<SmartSetRule>();
            if (enabledRules.Count == 0)
            {
                result.EstimatedMatchCount = 0;
                result.Notes = "No enabled rules.";
                return result;
            }

            var groupIds = enabledRules
                .Select(r => string.IsNullOrWhiteSpace(r.GroupId) ? "A" : r.GroupId.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var allItems = BuildAllItemSet(session);
            var finalUnion = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var gid in groupIds)
            {
                ct.ThrowIfCancellationRequested();

                var groupRules = enabledRules
                    .Where(r => (r.GroupId ?? "A").Trim().Equals(gid, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (groupRules.Count == 0) continue;

                HashSet<string> groupSet = null;

                foreach (var rule in groupRules)
                {
                    ct.ThrowIfCancellationRequested();
                    var ruleSet = EvaluateRule(session, rule, allItems, ct);

                    if (groupSet == null)
                    {
                        groupSet = ruleSet;
                    }
                    else
                    {
                        groupSet.IntersectWith(ruleSet);
                    }

                    if (groupSet.Count == 0)
                    {
                        break;
                    }
                }

                if (groupSet != null && groupSet.Count > 0)
                {
                    finalUnion.UnionWith(groupSet);
                }
            }

            result.EstimatedMatchCount = finalUnion.Count;
            result.SampleItemPaths = finalUnion.Take(Math.Max(0, samplePathsToReturn)).ToList();
            result.Notes = "Fast preview uses cached scrape entries; results may differ slightly from live Navisworks search.";
            return result;
        }

        private HashSet<string> EvaluateRule(
            IDataScrapeSessionView session,
            SmartSetRule rule,
            HashSet<string> allItems,
            CancellationToken ct)
        {
            var key = $"{rule.Category}::{rule.Property}";

            if (!_propertyIndexCache.TryGetValue(key, out var itemToValues))
            {
                itemToValues = BuildIndexForProperty(session, rule.Category, rule.Property, ct);
                _propertyIndexCache[key] = itemToValues;
            }

            var effectiveOperator = GetEffectiveOperator(rule);
            var hits = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (effectiveOperator == SmartSetOperator.Defined)
            {
                foreach (var item in itemToValues.Keys)
                {
                    hits.Add(item);
                }

                return hits;
            }

            if (effectiveOperator == SmartSetOperator.Undefined)
            {
                foreach (var item in allItems)
                {
                    if (!itemToValues.ContainsKey(item))
                    {
                        hits.Add(item);
                    }
                }

                return hits;
            }

            foreach (var kvp in itemToValues)
            {
                ct.ThrowIfCancellationRequested();

                var itemPath = kvp.Key;
                var values = kvp.Value;

                if (Matches(values, rule, effectiveOperator))
                {
                    hits.Add(itemPath);
                }
            }

            return hits;
        }

        private Dictionary<string, List<string>> BuildIndexForProperty(
            IDataScrapeSessionView session,
            string category,
            string property,
            CancellationToken ct)
        {
            var index = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var e in session.RawEntries)
            {
                ct.ThrowIfCancellationRequested();

                if (!string.Equals(e.Category ?? "", category ?? "", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.Equals(e.Property ?? "", property ?? "", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var item = e.ItemPath ?? "";
                if (string.IsNullOrWhiteSpace(item))
                {
                    continue;
                }

                if (!index.TryGetValue(item, out var list))
                {
                    list = new List<string>();
                    index[item] = list;
                }

                list.Add(e.Value ?? "");
            }

            return index;
        }

        private static bool Matches(List<string> values, SmartSetRule rule, SmartSetOperator effectiveOperator)
        {
            values = values ?? new List<string>();
            var needle = rule.Value ?? "";

            switch (effectiveOperator)
            {
                case SmartSetOperator.Equals:
                    return values.Any(v => string.Equals(v ?? "", needle, StringComparison.OrdinalIgnoreCase));

                case SmartSetOperator.NotEquals:
                    return values.Any(v => !string.Equals(v ?? "", needle, StringComparison.OrdinalIgnoreCase));

                case SmartSetOperator.Contains:
                    return values.Any(v => (v ?? "").IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0);

                case SmartSetOperator.Wildcard:
                    return values.Any(v => WildcardMatch(v ?? "", needle));
            }

            return false;
        }

        private static SmartSetOperator GetEffectiveOperator(SmartSetRule rule)
        {
            if (rule == null)
            {
                return SmartSetOperator.Defined;
            }

            var op = rule.Operator;
            if ((op == SmartSetOperator.Equals || op == SmartSetOperator.NotEquals
                || op == SmartSetOperator.Contains || op == SmartSetOperator.Wildcard)
                && string.IsNullOrWhiteSpace(rule.Value))
            {
                return SmartSetOperator.Defined;
            }

            return op;
        }

        private static HashSet<string> BuildAllItemSet(IDataScrapeSessionView session)
        {
            var all = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (session?.RawEntries == null)
            {
                return all;
            }

            foreach (var entry in session.RawEntries)
            {
                var item = entry.ItemPath ?? "";
                if (!string.IsNullOrWhiteSpace(item))
                {
                    all.Add(item);
                }
            }

            return all;
        }

        private static bool WildcardMatch(string input, string pattern)
        {
            if (string.IsNullOrEmpty(pattern))
            {
                return true;
            }

            var regexPattern = "^" + System.Text.RegularExpressions.Regex.Escape(pattern)
                .Replace("\\*", ".*")
                .Replace("\\?", ".") + "$";

            return System.Text.RegularExpressions.Regex.IsMatch(
                input ?? "",
                regexPattern,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }
    }
}
