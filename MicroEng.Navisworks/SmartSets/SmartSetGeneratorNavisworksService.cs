using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using MicroEng.Navisworks;

namespace MicroEng.Navisworks.SmartSets
{
    public sealed class SmartSetGeneratorNavisworksService
    {
        public (int Count, ModelItemCollection Results, bool UsedPostFilter) Preview(
            Document doc,
            SmartSetRecipe recipe,
            IReadOnlyList<SmartSetRule> rules)
        {
            var results = GetGroupedResults(doc, recipe, rules, out var usedPostFilter);
            return (results?.Count ?? 0, results, usedPostFilter);
        }

        public void Generate(Document doc, SmartSetRecipe recipe, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (recipe == null) throw new ArgumentNullException(nameof(recipe));

            var rules = recipe.Rules?.Where(r => r != null && r.Enabled).ToList() ?? new List<SmartSetRule>();
            var groups = GetGroupIds(rules);

            var supportsSearchSet = groups.Count <= 1 && rules.All(IsSearchSetSupportedOperator);
            var wantsSearchSet = recipe.OutputType == SmartSetOutputType.SearchSet
                || recipe.OutputType == SmartSetOutputType.Both;
            var wantsSelectionSet = recipe.OutputType == SmartSetOutputType.SelectionSet
                || recipe.OutputType == SmartSetOutputType.Both;
            var fallbackToSelection = recipe.OutputType == SmartSetOutputType.SearchSet && !supportsSearchSet;
            var includeBlanks = recipe.IncludeBlanks;
            ModelItemCollection results = null;
            var resultCount = 0;

            if (!supportsSearchSet && wantsSearchSet)
            {
                log?.Invoke("Search Set output disabled: multi-group or unsupported operators present.");
                if (fallbackToSelection)
                {
                    log?.Invoke("Creating Selection Set snapshot instead.");
                }
            }

            if (wantsSelectionSet || fallbackToSelection || !includeBlanks)
            {
                results = GetGroupedResults(doc, recipe, rules, out _, log);
                resultCount = results?.Count ?? 0;

                if (!includeBlanks && resultCount == 0)
                {
                    log?.Invoke($"Skipped empty set: {recipe.Name} (Include Blanks disabled).");
                    return;
                }
            }

            if (wantsSearchSet)
            {
                if (supportsSearchSet)
                {
                    var search = BuildSearchForGroup(doc, recipe, rules, forSearchSet: true, out _, log);
                    var folder = EnsureFolder(doc, recipe.FolderPath);

                    var ss = new SelectionSet(search)
                    {
                        DisplayName = MakeUniqueName(folder, recipe.Name)
                    };

                    AddCopySafe(doc, folder, ss, recipe.FolderPath, log);
                }
            }

            if (wantsSelectionSet || fallbackToSelection)
            {
                var folder = EnsureFolder(doc, recipe.FolderPath);
                var selectionName = recipe.Name + " (Snapshot)";
                if (fallbackToSelection)
                {
                    selectionName = recipe.Name;
                }
                var selectionResults = results ?? new ModelItemCollection();
                var selSet = new SelectionSet(selectionResults)
                {
                    DisplayName = MakeUniqueName(folder, selectionName)
                };

                AddCopySafe(doc, folder, selSet, recipe.FolderPath, log);
            }
        }

        internal int GenerateSplitSearchSets(Document doc, SmartSetRecipe recipe, ScrapeSession session, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (recipe == null) throw new ArgumentNullException(nameof(recipe));
            if (session == null) throw new ArgumentNullException(nameof(session));

            var rules = recipe.Rules?.Where(r => r != null && r.Enabled).ToList() ?? new List<SmartSetRule>();
            if (rules.Count == 0)
            {
                return 0;
            }

            var groups = GetGroupIds(rules);
            if (groups.Count > 1)
            {
                log?.Invoke("Multiple Search Sets supports a single group only.");
                return 0;
            }

            var splitRules = rules
                .Where(r => r.Operator == SmartSetOperator.Defined
                    && !string.IsNullOrWhiteSpace(r.Category)
                    && !string.IsNullOrWhiteSpace(r.Property))
                .ToList();
            if (splitRules.Count != 1)
            {
                log?.Invoke("Multiple Search Sets requires exactly one Defined rule with Category and Property.");
                return 0;
            }

            var splitRule = splitRules[0];
            var values = GetDistinctValues(session, splitRule.Category, splitRule.Property);
            if (values.Count == 0)
            {
                log?.Invoke("No distinct values found for the Defined rule.");
                return 0;
            }

            var wantsSearchSet = recipe.OutputType == SmartSetOutputType.SearchSet
                || recipe.OutputType == SmartSetOutputType.Both;
            var wantsSelectionSet = recipe.OutputType == SmartSetOutputType.SelectionSet
                || recipe.OutputType == SmartSetOutputType.Both;
            var includeBlanks = recipe.IncludeBlanks;

            var baseRules = rules.Where(r => !ReferenceEquals(r, splitRule)).ToList();
            var supportsSearchSet = baseRules.All(IsSearchSetSupportedOperator);

            if (wantsSearchSet && !supportsSearchSet)
            {
                log?.Invoke("Search Set output disabled: unsupported operators present.");
            }

            GroupItem folder = null;
            var created = 0;

            foreach (var value in values)
            {
                var rulesForSet = new List<SmartSetRule>(baseRules)
                {
                    new SmartSetRule
                    {
                        GroupId = splitRule.GroupId,
                        Category = splitRule.Category,
                        Property = splitRule.Property,
                        Operator = SmartSetOperator.Equals,
                        Value = value,
                        Enabled = true
                    }
                };

                var setName = BuildMultiSetName(recipe.Name, value);

                ModelItemCollection results = null;
                if (!includeBlanks || wantsSelectionSet)
                {
                    results = GetGroupedResults(doc, recipe, rulesForSet, out _, log);
                    if (!includeBlanks && (results?.Count ?? 0) == 0)
                    {
                        log?.Invoke($"Skipped empty set: {setName} (Include Blanks disabled).");
                        continue;
                    }
                }

                if (wantsSearchSet && supportsSearchSet)
                {
                    folder ??= EnsureFolder(doc, recipe.FolderPath);
                    var search = BuildSearchForGroup(doc, recipe, rulesForSet, forSearchSet: true, out _, log);
                    var ss = new SelectionSet(search)
                    {
                        DisplayName = MakeUniqueName(folder, setName)
                    };

                    AddCopySafe(doc, folder, ss, recipe.FolderPath, log);
                    created++;
                }

                if (wantsSelectionSet)
                {
                    folder ??= EnsureFolder(doc, recipe.FolderPath);
                    var selSet = new SelectionSet(results)
                    {
                        DisplayName = MakeUniqueName(folder, setName + " (Snapshot)")
                    };

                    AddCopySafe(doc, folder, selSet, recipe.FolderPath, log);
                    created++;
                }
            }

            return created;
        }

        internal int GenerateExpandedSearchSet(Document doc, SmartSetRecipe recipe, ScrapeSession session, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (recipe == null) throw new ArgumentNullException(nameof(recipe));
            if (session == null) throw new ArgumentNullException(nameof(session));

            var rules = recipe.Rules?.Where(r => r != null && r.Enabled).ToList() ?? new List<SmartSetRule>();
            if (rules.Count == 0)
            {
                return 0;
            }

            var groups = GetGroupIds(rules);
            if (groups.Count > 1)
            {
                log?.Invoke("Expand Values supports a single group only.");
                return 0;
            }

            var splitRules = rules
                .Where(r => r.Operator == SmartSetOperator.Defined
                    && !string.IsNullOrWhiteSpace(r.Category)
                    && !string.IsNullOrWhiteSpace(r.Property))
                .ToList();
            if (splitRules.Count != 1)
            {
                log?.Invoke("Expand Values requires exactly one Defined rule with Category and Property.");
                return 0;
            }

            var splitRule = splitRules[0];
            var values = GetDistinctValues(session, splitRule.Category, splitRule.Property);
            if (values.Count == 0)
            {
                log?.Invoke("No distinct values found for the Defined rule.");
                return 0;
            }

            var wantsSearchSet = recipe.OutputType == SmartSetOutputType.SearchSet
                || recipe.OutputType == SmartSetOutputType.Both;
            var wantsSelectionSet = recipe.OutputType == SmartSetOutputType.SelectionSet
                || recipe.OutputType == SmartSetOutputType.Both;
            var includeBlanks = recipe.IncludeBlanks;

            var baseRules = rules.Where(r => !ReferenceEquals(r, splitRule)).ToList();
            var supportsSearchSet = baseRules.All(IsSearchSetSupportedOperator);

            if (wantsSearchSet && !supportsSearchSet)
            {
                log?.Invoke("Search Set output disabled: unsupported operators present.");
            }

            var search = BuildSearchForValueGroups(doc, recipe, baseRules, splitRule, values, out var supported, log);
            if (!supported)
            {
                log?.Invoke("Expand Values failed: unsupported operators present.");
                return 0;
            }

            ModelItemCollection results = null;
            if (!includeBlanks || wantsSelectionSet)
            {
                results = search.FindAll(doc, false);
                if (!includeBlanks && (results?.Count ?? 0) == 0)
                {
                    log?.Invoke($"Skipped empty set: {recipe.Name} (Include Blanks disabled).");
                    return 0;
                }
            }

            GroupItem folder = null;
            var created = 0;

            if (wantsSearchSet && supportsSearchSet)
            {
                folder ??= EnsureFolder(doc, recipe.FolderPath);
                var ss = new SelectionSet(search)
                {
                    DisplayName = MakeUniqueName(folder, recipe.Name)
                };

                AddCopySafe(doc, folder, ss, recipe.FolderPath, log);
                created++;
            }

            if (wantsSelectionSet)
            {
                folder ??= EnsureFolder(doc, recipe.FolderPath);
                var selSet = new SelectionSet(results)
                {
                    DisplayName = MakeUniqueName(folder, recipe.Name + " (Snapshot)")
                };

                AddCopySafe(doc, folder, selSet, recipe.FolderPath, log);
                created++;
            }

            return created;
        }

        public void GenerateGroupedSearchSets(
            Document doc,
            SmartSetGroupingSpec grouping,
            string cat1,
            string prop1,
            bool use2,
            string cat2,
            string prop2,
            IEnumerable<SmartGroupRow> groups,
            Action<string> log = null)
        {
            var folderPath = grouping?.OutputFolderPath ?? "MicroEng/Smart Sets";
            var baseName = grouping?.OutputName ?? "";
            var includeBlanks = grouping?.IncludeBlanks ?? false;
            GroupItem folder = null;

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

                var search = BuildSearchForGroup(doc, grouping, rules, forSearchSet: true, out _, null);

                if (!includeBlanks)
                {
                    var results = search.FindAll(doc, false);
                    if (results == null || results.Count == 0)
                    {
                        log?.Invoke($"Skipped empty set: {recipeName} (Include Blanks disabled).");
                        continue;
                    }
                }

                folder ??= EnsureFolder(doc, folderPath);
                var set = new SelectionSet(search)
                {
                    DisplayName = MakeUniqueName(folder, recipeName)
                };

                AddCopySafe(doc, folder, set, folderPath);
            }
        }

        private ModelItemCollection GetGroupedResults(
            Document doc,
            SmartSetRecipe recipe,
            IReadOnlyList<SmartSetRule> rules,
            out bool usedPostFilter,
            Action<string> log = null)
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

                var search = BuildSearchForGroup(doc, recipe, groupRules, forSearchSet: false, out var postFilterNeeded, log);
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

        private Search BuildSearchForGroup(
            Document doc,
            SmartSetRecipe recipe,
            IEnumerable<SmartSetRule> rules,
            bool forSearchSet,
            out bool requiresPostFilter,
            Action<string> log = null)
        {
            var search = new Search();
            ApplyScopeToSearch(doc, recipe, search, log);

            requiresPostFilter = false;

            var scopeFilter = BuildScopeFilterCondition(recipe, log);
            if (scopeFilter != null)
            {
                search.SearchConditions.Add(scopeFilter);
            }

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

        private Search BuildSearchForGroup(
            Document doc,
            SmartSetGroupingSpec grouping,
            IEnumerable<SmartSetRule> rules,
            bool forSearchSet,
            out bool requiresPostFilter,
            Action<string> log = null)
        {
            var search = new Search();
            ApplyScopeToSearch(doc, grouping, search, log);

            requiresPostFilter = false;

            var scopeFilter = BuildScopeFilterCondition(grouping, log);
            if (scopeFilter != null)
            {
                search.SearchConditions.Add(scopeFilter);
            }

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

        private Search BuildSearchForValueGroups(
            Document doc,
            SmartSetRecipe recipe,
            IReadOnlyList<SmartSetRule> baseRules,
            SmartSetRule splitRule,
            IReadOnlyList<string> values,
            out bool supported,
            Action<string> log = null)
        {
            supported = true;

            var search = new Search();
            ApplyScopeToSearch(doc, recipe, search, log);
            var scopeFilter = BuildScopeFilterCondition(recipe, log);

            foreach (var value in values)
            {
                var group = new List<SearchCondition>();

                if (scopeFilter != null)
                {
                    group.Add(scopeFilter);
                }

                foreach (var rule in baseRules)
                {
                    if (!TryBuildSearchCondition(rule, out var condition))
                    {
                        supported = false;
                        return search;
                    }

                    if (condition != null)
                    {
                        group.Add(condition);
                    }
                }

                var valueCond = SearchCondition.HasPropertyByDisplayName(splitRule.Category, splitRule.Property)
                    .EqualValue(new VariantData(value ?? ""));
                group.Add(valueCond);

                search.SearchConditions.AddGroup(group);
            }

            return search;
        }

        public IReadOnlyList<string> GetSavedSelectionSetNames(Document doc)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var root = doc?.SelectionSets?.RootItem;
            CollectSelectionSetNames(root, names);

            return names.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static void CollectSelectionSetNames(GroupItem folder, HashSet<string> names)
        {
            if (folder?.Children == null)
            {
                return;
            }

            foreach (SavedItem item in folder.Children)
            {
                if (item is SelectionSet set)
                {
                    var label = set.DisplayName ?? "";
                    if (!string.IsNullOrWhiteSpace(label))
                    {
                        names.Add(label);
                    }
                }
                else if (item is GroupItem group)
                {
                    CollectSelectionSetNames(group, names);
                }
            }
        }

        private static void ApplyScopeToSearch(Document doc, SmartSetRecipe recipe, Search search, Action<string> log)
        {
            if (search == null)
            {
                return;
            }

            if (doc == null || recipe == null || recipe.ScopeMode == SmartSetScopeMode.AllModel)
            {
                search.Selection.SelectAll();
                return;
            }

            if (recipe.ScopeMode == SmartSetScopeMode.PropertyFilter)
            {
                search.Selection.SelectAll();
                return;
            }

            if (recipe.ScopeMode == SmartSetScopeMode.CurrentSelection)
            {
                var selection = doc.CurrentSelection?.SelectedItems;
                if (selection == null || selection.Count == 0)
                {
                    log?.Invoke("Scope uses current selection but it is empty. Using entire model.");
                    search.Selection.SelectAll();
                    return;
                }

                search.Selection.Clear();
                search.Selection.CopyFrom(selection);
                return;
            }

            if (recipe.ScopeMode == SmartSetScopeMode.SavedSelectionSet)
            {
                var name = recipe.ScopeSelectionSetName;
                if (string.IsNullOrWhiteSpace(name))
                {
                    log?.Invoke("Scope selection set not specified. Using entire model.");
                    search.Selection.SelectAll();
                    return;
                }

                var items = GetSelectionSetItems(doc, name);
                if (items == null || items.Count == 0)
                {
                    log?.Invoke($"Scope selection set not found or empty: {name}. Using entire model.");
                    search.Selection.SelectAll();
                    return;
                }

                search.Selection.Clear();
                search.Selection.CopyFrom(items);
                return;
            }

            if (recipe.ScopeMode == SmartSetScopeMode.ModelTree)
            {
                var items = ResolveScopeModelItems(doc, recipe, log);
                if (items == null || items.Count == 0)
                {
                    log?.Invoke("Scope tree selection could not be resolved. Using entire model.");
                    search.Selection.SelectAll();
                    return;
                }

                search.Selection.Clear();
                search.Selection.CopyFrom(items);
                return;
            }

            search.Selection.SelectAll();
        }

        private static void ApplyScopeToSearch(Document doc, SmartSetGroupingSpec grouping, Search search, Action<string> log)
        {
            if (search == null)
            {
                return;
            }

            if (doc == null || grouping == null || grouping.ScopeMode == SmartSetScopeMode.AllModel)
            {
                search.Selection.SelectAll();
                return;
            }

            if (grouping.ScopeMode == SmartSetScopeMode.PropertyFilter)
            {
                search.Selection.SelectAll();
                return;
            }

            if (grouping.ScopeMode == SmartSetScopeMode.CurrentSelection)
            {
                var selection = doc.CurrentSelection?.SelectedItems;
                if (selection == null || selection.Count == 0)
                {
                    log?.Invoke("Scope uses current selection but it is empty. Using entire model.");
                    search.Selection.SelectAll();
                    return;
                }

                search.Selection.Clear();
                search.Selection.CopyFrom(selection);
                return;
            }

            if (grouping.ScopeMode == SmartSetScopeMode.SavedSelectionSet)
            {
                var name = grouping.ScopeSelectionSetName;
                if (string.IsNullOrWhiteSpace(name))
                {
                    log?.Invoke("Scope selection set not specified. Using entire model.");
                    search.Selection.SelectAll();
                    return;
                }

                var items = GetSelectionSetItems(doc, name);
                if (items == null || items.Count == 0)
                {
                    log?.Invoke($"Scope selection set not found or empty: {name}. Using entire model.");
                    search.Selection.SelectAll();
                    return;
                }

                search.Selection.Clear();
                search.Selection.CopyFrom(items);
                return;
            }

            if (grouping.ScopeMode == SmartSetScopeMode.ModelTree)
            {
                var items = ResolveScopeModelItems(doc, grouping, log);
                if (items == null || items.Count == 0)
                {
                    log?.Invoke("Scope tree selection could not be resolved. Using entire model.");
                    search.Selection.SelectAll();
                    return;
                }

                search.Selection.Clear();
                search.Selection.CopyFrom(items);
                return;
            }

            search.Selection.SelectAll();
        }

        private static ModelItemCollection GetSelectionSetItems(Document doc, string name)
        {
            try
            {
                var root = doc?.SelectionSets?.RootItem;
                var selectionSet = FindSelectionSetByName(root, name);
                return selectionSet?.GetSelectedItems(doc);
            }
            catch
            {
                return null;
            }
        }

        private static SelectionSet FindSelectionSetByName(GroupItem folder, string name)
        {
            if (folder == null || string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            var children = folder.Children;
            if (children == null)
            {
                return null;
            }

            foreach (SavedItem item in children)
            {
                if (item is SelectionSet set)
                {
                    if (string.Equals(set.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return set;
                    }
                }
                else if (item is GroupItem group)
                {
                    var found = FindSelectionSetByName(group, name);
                    if (found != null)
                    {
                        return found;
                    }
                }
            }

            return null;
        }

        private static ModelItemCollection ResolveScopeModelItems(Document doc, SmartSetRecipe recipe, Action<string> log)
        {
            if (doc?.Models?.RootItems == null || recipe?.ScopeModelPaths == null || recipe.ScopeModelPaths.Count == 0)
            {
                return null;
            }

            var items = new ModelItemCollection();
            var added = new HashSet<Guid>();
            foreach (var path in recipe.ScopeModelPaths)
            {
                var resolved = ResolveModelPath(doc.Models.RootItems, path);
                if (resolved == null)
                {
                    log?.Invoke($"Scope path not found: {string.Join(" > ", path ?? new List<string>())}");
                    continue;
                }

                var id = resolved.InstanceGuid;
                if (id == Guid.Empty || added.Add(id))
                {
                    items.Add(resolved);
                }
            }

            return items;
        }

        private static ModelItemCollection ResolveScopeModelItems(Document doc, SmartSetGroupingSpec grouping, Action<string> log)
        {
            if (doc?.Models?.RootItems == null || grouping?.ScopeModelPaths == null || grouping.ScopeModelPaths.Count == 0)
            {
                return null;
            }

            var items = new ModelItemCollection();
            var added = new HashSet<Guid>();
            foreach (var path in grouping.ScopeModelPaths)
            {
                var resolved = ResolveModelPath(doc.Models.RootItems, path);
                if (resolved == null)
                {
                    log?.Invoke($"Scope path not found: {string.Join(" > ", path ?? new List<string>())}");
                    continue;
                }

                var id = resolved.InstanceGuid;
                if (id == Guid.Empty || added.Add(id))
                {
                    items.Add(resolved);
                }
            }

            return items;
        }

        private static ModelItem ResolveModelPath(IEnumerable<ModelItem> roots, IReadOnlyList<string> path)
        {
            if (roots == null || path == null || path.Count == 0)
            {
                return null;
            }

            var current = FindRootByLabel(roots, path[0]);
            if (current == null)
            {
                return null;
            }

            for (var i = 1; i < path.Count; i++)
            {
                var segment = path[i];
                if (string.IsNullOrWhiteSpace(segment))
                {
                    return null;
                }

                var next = FindChildByLabel(current, segment);
                if (next == null)
                {
                    return null;
                }

                current = next;
            }

            return current;
        }

        private static ModelItem FindRootByLabel(IEnumerable<ModelItem> roots, string label)
        {
            if (string.IsNullOrWhiteSpace(label))
            {
                return null;
            }

            foreach (var root in roots)
            {
                if (root == null)
                {
                    continue;
                }

                if (string.Equals(GetItemLabel(root), label, StringComparison.OrdinalIgnoreCase))
                {
                    return root;
                }
            }

            return null;
        }

        private static ModelItem FindChildByLabel(ModelItem parent, string label)
        {
            try
            {
                if (parent?.Children == null)
                {
                    return null;
                }

                foreach (var child in parent.Children)
                {
                    if (child == null)
                    {
                        continue;
                    }

                    if (string.Equals(GetItemLabel(child), label, StringComparison.OrdinalIgnoreCase))
                    {
                        return child;
                    }
                }
            }
            catch
            {
                return null;
            }

            return null;
        }

        private static string GetItemLabel(ModelItem item)
        {
            if (item == null)
            {
                return "";
            }

            var label = item.DisplayName ?? "";
            if (!string.IsNullOrWhiteSpace(label))
            {
                return label;
            }

            label = item.ClassDisplayName ?? "";
            if (!string.IsNullOrWhiteSpace(label))
            {
                return label;
            }

            label = item.ClassName ?? "";
            if (!string.IsNullOrWhiteSpace(label))
            {
                return label;
            }

            return item.InstanceGuid.ToString();
        }

        private static SearchCondition BuildScopeFilterCondition(SmartSetRecipe recipe, Action<string> log)
        {
            if (recipe == null || recipe.ScopeMode != SmartSetScopeMode.PropertyFilter)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(recipe.ScopeFilterCategory)
                || string.IsNullOrWhiteSpace(recipe.ScopeFilterProperty)
                || string.IsNullOrWhiteSpace(recipe.ScopeFilterValue))
            {
                log?.Invoke("Scope property filter is incomplete. Using entire model.");
                return null;
            }

            return SearchCondition.HasPropertyByDisplayName(recipe.ScopeFilterCategory, recipe.ScopeFilterProperty)
                .EqualValue(new VariantData(recipe.ScopeFilterValue ?? ""));
        }

        private static SearchCondition BuildScopeFilterCondition(SmartSetGroupingSpec grouping, Action<string> log)
        {
            if (grouping == null || grouping.ScopeMode != SmartSetScopeMode.PropertyFilter)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(grouping.ScopeFilterCategory)
                || string.IsNullOrWhiteSpace(grouping.ScopeFilterProperty)
                || string.IsNullOrWhiteSpace(grouping.ScopeFilterValue))
            {
                log?.Invoke("Scope property filter is incomplete. Using entire model.");
                return null;
            }

            return SearchCondition.HasPropertyByDisplayName(grouping.ScopeFilterCategory, grouping.ScopeFilterProperty)
                .EqualValue(new VariantData(grouping.ScopeFilterValue ?? ""));
        }

        private static bool TryBuildSearchCondition(SmartSetRule rule, out SearchCondition condition)
        {
            condition = null;

            if (rule == null || !rule.Enabled)
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(rule.Category) || string.IsNullOrWhiteSpace(rule.Property))
            {
                return true;
            }

            var cond = SearchCondition.HasPropertyByDisplayName(rule.Category, rule.Property);
            var effectiveOperator = GetEffectiveOperator(rule);

            switch (effectiveOperator)
            {
                case SmartSetOperator.Defined:
                    condition = cond;
                    return true;
                case SmartSetOperator.Undefined:
                    condition = cond.Negate();
                    return true;
                case SmartSetOperator.Equals:
                    condition = cond.EqualValue(new VariantData(rule.Value ?? ""));
                    return true;
                case SmartSetOperator.NotEquals:
                    condition = cond.EqualValue(new VariantData(rule.Value ?? "")).Negate();
                    return true;
                case SmartSetOperator.Contains:
                    condition = cond.DisplayStringContains(rule.Value ?? "");
                    return true;
                case SmartSetOperator.Wildcard:
                    condition = cond.DisplayStringWildcard(rule.Value ?? "");
                    return true;
                default:
                    return false;
            }
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
            var hasValue = hasProperty && !string.IsNullOrWhiteSpace(value);

            var effectiveOperator = GetEffectiveOperator(rule);

            switch (effectiveOperator)
            {
                case SmartSetOperator.Defined:
                    return hasProperty;
                case SmartSetOperator.Undefined:
                    return !hasProperty;
                case SmartSetOperator.Equals:
                    return hasValue && string.Equals(value, rule.Value ?? "", StringComparison.OrdinalIgnoreCase);
                case SmartSetOperator.NotEquals:
                    return hasValue && !string.Equals(value, rule.Value ?? "", StringComparison.OrdinalIgnoreCase);
                case SmartSetOperator.Contains:
                    return hasValue && (value ?? "").IndexOf(rule.Value ?? "", StringComparison.OrdinalIgnoreCase) >= 0;
                case SmartSetOperator.Wildcard:
                    return hasValue && WildcardMatch(value ?? "", rule.Value ?? "");
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

        private static void AddCopySafe(Document doc, GroupItem parent, SavedItem item, string folderPath, Action<string> log = null)
        {
            try
            {
                doc.SelectionSets.AddCopy(parent, item);
                return;
            }
            catch (ArgumentException ex) when (string.Equals(ex.ParamName, "parent", StringComparison.OrdinalIgnoreCase))
            {
                log?.Invoke("Selection set parent not in SelectionSets; refreshing folder.");
            }

            try
            {
                var refreshed = EnsureFolder(doc, folderPath);
                doc.SelectionSets.AddCopy(refreshed, item);
                return;
            }
            catch (ArgumentException ex) when (string.Equals(ex.ParamName, "parent", StringComparison.OrdinalIgnoreCase))
            {
                log?.Invoke("Selection set parent still invalid; adding at root.");
            }

            doc.SelectionSets.AddCopy(doc.SelectionSets.RootItem, item);
        }

        private static GroupItem EnsureFolder(Document doc, string folderPath)
        {
            var parts = (folderPath ?? "").Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0)
            {
                return doc.SelectionSets.RootItem;
            }

            var path = new List<string>(parts.Length);

            foreach (var part in parts)
            {
                path.Add(part);

                var existing = ResolveFolderPath(doc, path);
                if (existing != null)
                {
                    continue;
                }

                var parent = ResolveFolderPath(doc, path.Count > 1 ? path.GetRange(0, path.Count - 1) : Array.Empty<string>());
                parent ??= doc.SelectionSets.RootItem;

                var folder = new FolderItem { DisplayName = part };
                if (!TryAddFolder(doc, parent, folder))
                {
                    parent = ResolveFolderPath(doc, path.Count > 1 ? path.GetRange(0, path.Count - 1) : Array.Empty<string>());
                    if (parent == null || !TryAddFolder(doc, parent, folder))
                    {
                        return doc.SelectionSets.RootItem;
                    }
                }

                existing = ResolveFolderPath(doc, path);
                if (existing == null)
                {
                    return doc.SelectionSets.RootItem;
                }
            }

            return ResolveFolderPath(doc, path) ?? doc.SelectionSets.RootItem;
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

        private static GroupItem FindChildGroup(GroupItem parent, string name)
        {
            return parent.Children
                .OfType<GroupItem>()
                .FirstOrDefault(x => string.Equals(x.DisplayName, name, StringComparison.OrdinalIgnoreCase));
        }

        private static bool TryAddFolder(Document doc, GroupItem parent, FolderItem folder)
        {
            try
            {
                doc.SelectionSets.AddCopy(parent, folder);
                return true;
            }
            catch (ArgumentException ex) when (string.Equals(ex.ParamName, "parent", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        private static GroupItem ResolveFolderPath(Document doc, IReadOnlyList<string> parts)
        {
            if (parts == null || parts.Count == 0)
            {
                return doc.SelectionSets.RootItem;
            }

            for (var attempt = 0; attempt < 3; attempt++)
            {
                GroupItem current = doc.SelectionSets.RootItem;
                var found = true;

                foreach (var part in parts)
                {
                    var next = FindChildGroup(current, part);
                    if (next == null)
                    {
                        found = false;
                        break;
                    }

                    current = next;
                }

                if (found)
                {
                    return current;
                }

                System.Windows.Forms.Application.DoEvents();
            }

            return null;
        }

        private static List<string> GetDistinctValues(ScrapeSession session, string category, string property)
        {
            var groups = SmartSetGroupingEngine.BuildGroups(
                session,
                category,
                property,
                useThenBy: false,
                thenByCategory: "",
                thenByProperty: "",
                minCount: 1,
                includeBlanks: false);

            return groups
                .Select(g => g.Value1 ?? "")
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string BuildMultiSetName(string baseName, string value)
        {
            if (string.IsNullOrWhiteSpace(baseName))
            {
                return value ?? "Set";
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                return baseName;
            }

            return $"{baseName} - {value}";
        }
    }
}
