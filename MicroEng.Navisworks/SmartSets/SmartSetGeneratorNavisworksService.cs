using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;

namespace MicroEng.Navisworks.SmartSets
{
    public sealed class SmartSetGeneratorNavisworksService
    {
        public (int Count, ModelItemCollection Results, bool UsedPostFilter) Preview(Document doc, IReadOnlyList<SmartSetRule> rules)
        {
            var results = GetGroupedResults(doc, rules, out var usedPostFilter);
            return (results?.Count ?? 0, results, usedPostFilter);
        }

        public void Generate(Document doc, SmartSetRecipe recipe, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (recipe == null) throw new ArgumentNullException(nameof(recipe));

            var rules = recipe.Rules?.Where(r => r != null && r.Enabled).ToList() ?? new List<SmartSetRule>();
            var groups = GetGroupIds(rules);

            var supportsSearchSet = groups.Count <= 1 && rules.All(IsSearchSetSupportedOperator);

            if (!supportsSearchSet && (recipe.OutputType == SmartSetOutputType.SearchSet || recipe.OutputType == SmartSetOutputType.Both))
            {
                log?.Invoke("Search Set output disabled: multi-group or unsupported operators present.");
            }

            if (recipe.OutputType == SmartSetOutputType.SearchSet || recipe.OutputType == SmartSetOutputType.Both)
            {
                if (supportsSearchSet)
                {
                    var search = BuildSearchForGroup(rules, out _);
                    var folder = EnsureFolder(doc, recipe.FolderPath);

                    var ss = new SelectionSet(search)
                    {
                        DisplayName = MakeUniqueName(folder, recipe.Name)
                    };

                    doc.SelectionSets.AddCopy(folder, ss);
                }
            }

            if (recipe.OutputType == SmartSetOutputType.SelectionSet || recipe.OutputType == SmartSetOutputType.Both)
            {
                var folder = EnsureFolder(doc, recipe.FolderPath);
                var results = GetGroupedResults(doc, rules, out _);
                var selSet = new SelectionSet(results)
                {
                    DisplayName = MakeUniqueName(folder, recipe.Name + " (Snapshot)")
                };

                doc.SelectionSets.AddCopy(folder, selSet);
            }
        }

        public void GenerateGroupedSearchSets(
            Document doc,
            string folderPath,
            string baseName,
            string cat1,
            string prop1,
            bool use2,
            string cat2,
            string prop2,
            IEnumerable<SmartGroupRow> groups)
        {
            var folder = EnsureFolder(doc, folderPath);

            foreach (var group in groups)
            {
                var recipeName = string.IsNullOrWhiteSpace(baseName)
                    ? group.DisplayKey
                    : $"{baseName} - {group.DisplayKey}";

                var rules = new List<SmartSetRule>
                {
                    new SmartSetRule
                    {
                        Category = cat1,
                        Property = prop1,
                        Operator = SmartSetOperator.Equals,
                        Value = group.Value1
                    }
                };

                if (use2)
                {
                    rules.Add(new SmartSetRule
                    {
                        Category = cat2,
                        Property = prop2,
                        Operator = SmartSetOperator.Equals,
                        Value = group.Value2
                    });
                }

                var search = BuildSearchForGroup(rules, out _);

                var set = new SelectionSet(search)
                {
                    DisplayName = MakeUniqueName(folder, recipeName)
                };

                doc.SelectionSets.AddCopy(folder, set);
            }
        }

        private ModelItemCollection GetGroupedResults(Document doc, IReadOnlyList<SmartSetRule> rules, out bool usedPostFilter)
        {
            usedPostFilter = false;

            if (doc == null)
            {
                return new ModelItemCollection();
            }

            var enabledRules = rules?.Where(r => r != null && r.Enabled).ToList() ?? new List<SmartSetRule>();
            if (enabledRules.Count == 0)
            {
                return new ModelItemCollection();
            }

            var groupIds = GetGroupIds(enabledRules);

            var combined = new Dictionary<Guid, ModelItem>();

            foreach (var gid in groupIds)
            {
                var groupRules = enabledRules
                    .Where(r => string.Equals((r.GroupId ?? "A").Trim(), gid, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (groupRules.Count == 0)
                {
                    continue;
                }

                var search = BuildSearchForGroup(groupRules, out var postFilterNeeded);
                var groupResults = search.FindAll(doc, false);

                if (postFilterNeeded)
                {
                    usedPostFilter = true;
                    groupResults = ApplyPostFilter(groupResults, groupRules);
                }

                foreach (var item in groupResults)
                {
                    var id = item?.InstanceGuid ?? Guid.Empty;
                    if (id == Guid.Empty)
                    {
                        continue;
                    }

                    if (!combined.ContainsKey(id))
                    {
                        combined[id] = item;
                    }
                }
            }

            var collection = new ModelItemCollection();
            foreach (var item in combined.Values)
            {
                collection.Add(item);
            }

            return collection;
        }

        private static List<string> GetGroupIds(IEnumerable<SmartSetRule> rules)
        {
            return rules
                .Select(r => string.IsNullOrWhiteSpace(r.GroupId) ? "A" : r.GroupId.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private Search BuildSearchForGroup(IEnumerable<SmartSetRule> rules, out bool requiresPostFilter)
        {
            var search = new Search();
            search.Selection.SelectAll();

            requiresPostFilter = false;

            foreach (var rule in rules)
            {
                if (rule == null || !rule.Enabled)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(rule.Category) || string.IsNullOrWhiteSpace(rule.Property))
                {
                    continue;
                }

                var cond = SearchCondition.HasPropertyByDisplayName(rule.Category, rule.Property);

                var effectiveOperator = GetEffectiveOperator(rule);

                switch (effectiveOperator)
                {
                    case SmartSetOperator.Defined:
                        search.SearchConditions.Add(cond);
                        break;
                    case SmartSetOperator.Undefined:
                        search.SearchConditions.Add(cond.Negate());
                        break;
                    case SmartSetOperator.Equals:
                        search.SearchConditions.Add(cond.EqualValue(new VariantData(rule.Value ?? "")));
                        break;
                    case SmartSetOperator.NotEquals:
                        search.SearchConditions.Add(cond.EqualValue(new VariantData(rule.Value ?? "")).Negate());
                        break;
                    case SmartSetOperator.Contains:
                        search.SearchConditions.Add(cond.DisplayStringContains(rule.Value ?? ""));
                        break;
                    case SmartSetOperator.Wildcard:
                        search.SearchConditions.Add(cond.DisplayStringWildcard(rule.Value ?? ""));
                        break;
                    default:
                        requiresPostFilter = true;
                        search.SearchConditions.Add(cond);
                        break;
                }
            }

            return search;
        }

        private static bool IsSearchSetSupportedOperator(SmartSetRule rule)
        {
            if (rule == null)
            {
                return true;
            }

            switch (GetEffectiveOperator(rule))
            {
                case SmartSetOperator.Equals:
                case SmartSetOperator.Contains:
                case SmartSetOperator.NotEquals:
                case SmartSetOperator.Wildcard:
                case SmartSetOperator.Defined:
                case SmartSetOperator.Undefined:
                    return true;
                default:
                    return false;
            }
        }

        private static ModelItemCollection ApplyPostFilter(ModelItemCollection items, IReadOnlyList<SmartSetRule> rules)
        {
            var results = new ModelItemCollection();
            foreach (var item in items)
            {
                if (item == null)
                {
                    continue;
                }

                if (MatchesAllRules(item, rules))
                {
                    results.Add(item);
                }
            }

            return results;
        }

        private static bool MatchesAllRules(ModelItem item, IReadOnlyList<SmartSetRule> rules)
        {
            foreach (var rule in rules)
            {
                if (rule == null || !rule.Enabled)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(rule.Category) || string.IsNullOrWhiteSpace(rule.Property))
                {
                    continue;
                }

                if (!MatchesRule(item, rule))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool MatchesRule(ModelItem item, SmartSetRule rule)
        {
            var hasProperty = TryGetPropertyValue(item, rule.Category, rule.Property, out var value);

            var effectiveOperator = GetEffectiveOperator(rule);
            if (!hasProperty && effectiveOperator != SmartSetOperator.Undefined)
            {
                return false;
            }

            switch (effectiveOperator)
            {
                case SmartSetOperator.Defined:
                    return hasProperty;
                case SmartSetOperator.Undefined:
                    return !hasProperty;
                case SmartSetOperator.Equals:
                    return string.Equals(value, rule.Value ?? "", StringComparison.OrdinalIgnoreCase);
                case SmartSetOperator.NotEquals:
                    return !string.Equals(value, rule.Value ?? "", StringComparison.OrdinalIgnoreCase);
                case SmartSetOperator.Contains:
                    return (value ?? "").IndexOf(rule.Value ?? "", StringComparison.OrdinalIgnoreCase) >= 0;
                case SmartSetOperator.Wildcard:
                    return WildcardMatch(value ?? "", rule.Value ?? "");
            }

            return false;
        }

        private static bool TryGetPropertyValue(ModelItem item, string categoryName, string propertyName, out string value)
        {
            value = "";
            try
            {
                if (item?.PropertyCategories == null)
                {
                    return false;
                }

                foreach (PropertyCategory category in item.PropertyCategories)
                {
                    if (category == null)
                    {
                        continue;
                    }

                    var catName = category.DisplayName ?? category.Name ?? "";
                    if (!string.Equals(catName, categoryName ?? "", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    foreach (DataProperty prop in category.Properties)
                    {
                        if (prop == null)
                        {
                            continue;
                        }

                        var propName = prop.DisplayName ?? prop.Name ?? "";
                        if (!string.Equals(propName, propertyName ?? "", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        if (prop.Value == null)
                        {
                            value = "";
                            return true;
                        }

                        if (prop.Value.IsDisplayString)
                        {
                            value = prop.Value.ToDisplayString() ?? "";
                            return true;
                        }

                        value = prop.Value.ToString() ?? "";
                        return true;
                    }
                }
            }
            catch
            {
                // ignore and treat as blank
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

        private static GroupItem EnsureFolder(Document doc, string folderPath)
        {
            var parts = (folderPath ?? "").Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);

            GroupItem current = doc.SelectionSets.RootItem;
            foreach (var part in parts)
            {
                var existing = current.Children
                    .OfType<GroupItem>()
                    .FirstOrDefault(x => string.Equals(x.DisplayName, part, StringComparison.OrdinalIgnoreCase));

                if (existing != null)
                {
                    current = existing;
                    continue;
                }

                var folder = new FolderItem { DisplayName = part };
                doc.SelectionSets.AddCopy(current, folder);
                current = current.Children
                    .OfType<GroupItem>()
                    .FirstOrDefault(x => string.Equals(x.DisplayName, part, StringComparison.OrdinalIgnoreCase))
                          ?? current;
            }

            return current;
        }

        private static string MakeUniqueName(GroupItem folder, string desired)
        {
            desired = SanitizeName(desired);
            if (string.IsNullOrWhiteSpace(desired))
            {
                desired = "Set";
            }

            var names = new HashSet<string>(
                folder.Children.Select(x => x.DisplayName),
                StringComparer.OrdinalIgnoreCase);

            if (!names.Contains(desired))
            {
                return desired;
            }

            for (var i = 2; i < 9999; i++)
            {
                var candidate = $"{desired} ({i})";
                if (!names.Contains(candidate))
                {
                    return candidate;
                }
            }

            return desired + " (9999)";
        }

        private static string SanitizeName(string value)
        {
            if (value == null) return "Set";

            foreach (var c in System.IO.Path.GetInvalidFileNameChars())
            {
                value = value.Replace(c, '_');
            }

            return value.Trim();
        }
    }
}
