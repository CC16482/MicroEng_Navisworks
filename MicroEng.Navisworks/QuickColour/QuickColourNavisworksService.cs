using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using Autodesk.Navisworks.Api;

namespace MicroEng.Navisworks.QuickColour
{
    public sealed class QuickColourNavisworksService
    {
        public int ApplyBySingleProperty(
            Document doc,
            string category,
            string property,
            IEnumerable<QuickColourValueRow> values,
            QuickColourScope scope,
            string scopeSelectionSetName,
            IReadOnlyList<List<string>> scopeModelPaths,
            string scopeFilterCategory,
            string scopeFilterProperty,
            string scopeFilterValue,
            bool permanent,
            bool createSearchSets,
            bool createSnapshots,
            string folderPath,
            string profileName,
            Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            var enabledValues = (values ?? Enumerable.Empty<QuickColourValueRow>())
                .Where(v => v != null && v.Enabled)
                .ToList();

            if (enabledValues.Count == 0)
            {
                log?.Invoke("QuickColour: no enabled values.");
                return 0;
            }

            log?.Invoke($"QuickColour: apply started. Scope={scope}, Values={enabledValues.Count}.");
            var verboseLogging = IsVerboseLoggingEnabled();

            GroupItem folder = null;
            if ((createSearchSets || createSnapshots) && !string.IsNullOrWhiteSpace(folderPath))
            {
                folder = EnsureFolder(doc, CombinePath(folderPath, profileName));
            }

            var totalColored = 0;

            var zeroHits = 0;
            var rowsWithHits = 0;
            var debugLogged = 0;
            const int maxDebug = 12;

            foreach (var row in enabledValues)
            {
                var search = new Search();
                ApplyScopeToSearch(doc, search, scope, scopeSelectionSetName, scopeModelPaths, log);
                var scopeFilter = BuildScopeFilterCondition(scope, scopeFilterCategory, scopeFilterProperty, scopeFilterValue, log);
                if (scopeFilter != null)
                {
                    search.SearchConditions.Add(scopeFilter);
                }

                var normalized = NormalizeSearchValue(row.Value);
                var variant = BuildVariantData(normalized, out var variantLabel);
                search.SearchConditions.Add(
                    SearchCondition.HasPropertyByDisplayName(category, property)
                        .EqualValue(variant));

                var results = search.FindAll(doc, false);

                if (results == null || results.Count == 0)
                {
                    zeroHits++;
                    if (verboseLogging && debugLogged < maxDebug)
                    {
                        log?.Invoke($"QuickColour: value='{row.Value}' normalized='{normalized}' variant={variantLabel} -> 0 items.");
                        debugLogged++;
                    }
                    continue;
                }

                rowsWithHits++;
                if (verboseLogging && debugLogged < maxDebug)
                {
                    log?.Invoke($"QuickColour: value='{row.Value}' normalized='{normalized}' variant={variantLabel} -> {results.Count} items.");
                    debugLogged++;
                }

                var navisColor = QuickColourPalette.ToNavisworksColor(row.Color);
                if (permanent)
                {
                    doc.Models.OverridePermanentColor(results, navisColor);
                }
                else
                {
                    doc.Models.OverrideTemporaryColor(results, navisColor);
                }

                totalColored += results.Count;

                if (folder != null)
                {
                    var baseName = SanitizeName($"{property} = {TrimForName(row.Value)}");

                    if (createSearchSets)
                    {
                        var ss = new SelectionSet(search)
                        {
                            DisplayName = MakeUniqueName(folder, baseName)
                        };
                        AddCopySafe(doc, folder, ss, folderPath, log);
                    }

                    if (createSnapshots)
                    {
                        var snap = new SelectionSet(results)
                        {
                            DisplayName = MakeUniqueName(folder, baseName + " (Snapshot)")
                        };
                        AddCopySafe(doc, folder, snap, folderPath, log);
                    }
                }
            }

            log?.Invoke($"QuickColour: apply summary. RowsWithHits={rowsWithHits}, ZeroHits={zeroHits}, TotalColored={totalColored}.");
            return totalColored;
        }

        public int ApplyByHierarchy(
            Document doc,
            string l1Category,
            string l1Property,
            string l2Category,
            string l2Property,
            IEnumerable<QuickColourHierarchyGroup> groups,
            QuickColourScope scope,
            string scopeSelectionSetName,
            IReadOnlyList<List<string>> scopeModelPaths,
            string scopeFilterCategory,
            string scopeFilterProperty,
            string scopeFilterValue,
            bool permanent,
            bool createSearchSets,
            bool createSnapshots,
            string folderPath,
            string profileName,
            bool createFoldersByHueGroup,
            Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            var enabledGroups = (groups ?? Enumerable.Empty<QuickColourHierarchyGroup>())
                .Where(g => g != null && g.Enabled)
                .ToList();

            if (enabledGroups.Count == 0)
            {
                log?.Invoke("QuickColour(Hierarchy): no enabled groups.");
                return 0;
            }

            log?.Invoke($"QuickColour(Hierarchy): apply started. Scope={scope}, Groups={enabledGroups.Count}.");
            var verboseLogging = IsVerboseLoggingEnabled();

            var totalColored = 0;
            var zeroHits = 0;
            var rowsWithHits = 0;
            var debugLogged = 0;
            const int maxDebug = 12;

            foreach (var group in enabledGroups)
            {
                var enabledTypes = group.Types.Where(t => t != null && t.Enabled).ToList();
                if (enabledTypes.Count == 0)
                {
                    continue;
                }

                var folder = EnsureOutputFolder(doc, folderPath, profileName, group, createFoldersByHueGroup, createSearchSets || createSnapshots);

                foreach (var typeRow in enabledTypes)
                {
                    var search = new Search();
                    ApplyScopeToSearch(doc, search, scope, scopeSelectionSetName, scopeModelPaths, log);
                    var scopeFilter = BuildScopeFilterCondition(scope, scopeFilterCategory, scopeFilterProperty, scopeFilterValue, log);
                    if (scopeFilter != null)
                    {
                        search.SearchConditions.Add(scopeFilter);
                    }
                    var l1Normalized = NormalizeSearchValue(group.Value);
                    var l2Normalized = NormalizeSearchValue(typeRow.Value);
                    var l1Variant = BuildVariantData(l1Normalized, out var l1VariantLabel);
                    var l2Variant = BuildVariantData(l2Normalized, out var l2VariantLabel);
                    search.SearchConditions.Add(
                        SearchCondition.HasPropertyByDisplayName(l1Category, l1Property)
                            .EqualValue(l1Variant));
                    search.SearchConditions.Add(
                        SearchCondition.HasPropertyByDisplayName(l2Category, l2Property)
                            .EqualValue(l2Variant));

                    var results = search.FindAll(doc, false);

                    if (results == null || results.Count == 0)
                    {
                        zeroHits++;
                        if (verboseLogging && debugLogged < maxDebug)
                        {
                            log?.Invoke($"QuickColour(Hierarchy): L1='{group.Value}' ({l1VariantLabel}) L2='{typeRow.Value}' ({l2VariantLabel}) -> 0 items.");
                            debugLogged++;
                        }
                        continue;
                    }

                    rowsWithHits++;
                    if (verboseLogging && debugLogged < maxDebug)
                    {
                        log?.Invoke($"QuickColour(Hierarchy): L1='{group.Value}' ({l1VariantLabel}) L2='{typeRow.Value}' ({l2VariantLabel}) -> {results.Count} items.");
                        debugLogged++;
                    }

                    var navisColor = QuickColourPalette.ToNavisworksColor(typeRow.Color);
                    if (permanent)
                    {
                        doc.Models.OverridePermanentColor(results, navisColor);
                    }
                    else
                    {
                        doc.Models.OverrideTemporaryColor(results, navisColor);
                    }

                    totalColored += results.Count;

                    if (folder != null)
                    {
                        var baseName = SanitizeName($"{l2Property} = {TrimForName(typeRow.Value)}");

                        if (createSearchSets)
                        {
                            var ss = new SelectionSet(search)
                            {
                                DisplayName = MakeUniqueName(folder, baseName)
                            };
                            AddCopySafe(doc, folder, ss, folderPath, log);
                        }

                        if (createSnapshots)
                        {
                            var snap = new SelectionSet(results)
                            {
                                DisplayName = MakeUniqueName(folder, baseName + " (Snapshot)")
                            };
                            AddCopySafe(doc, folder, snap, folderPath, log);
                        }
                    }
                }
            }

            log?.Invoke($"QuickColour(Hierarchy): apply summary. RowsWithHits={rowsWithHits}, ZeroHits={zeroHits}, TotalColored={totalColored}.");
            return totalColored;
        }

        private static GroupItem EnsureOutputFolder(
            Document doc,
            string folderPath,
            string profileName,
            QuickColourHierarchyGroup group,
            bool createFoldersByHueGroup,
            bool createSets)
        {
            if (!createSets || string.IsNullOrWhiteSpace(folderPath))
            {
                return null;
            }

            var groupPathParts = new List<string>
            {
                folderPath,
                profileName
            };

            if (createFoldersByHueGroup && !string.IsNullOrWhiteSpace(group?.HueGroupName))
            {
                groupPathParts.Add(group.HueGroupName);
            }

            if (!string.IsNullOrWhiteSpace(group?.Value))
            {
                groupPathParts.Add(SanitizeName(group.Value));
            }

            var full = CombinePath(groupPathParts.ToArray());
            return EnsureFolder(doc, full);
        }

        private static string NormalizeSearchValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var trimmed = value.Trim();
            var idx = trimmed.IndexOf(':');
            if (idx > 0)
            {
                var prefix = trimmed.Substring(0, idx).Trim();
                if (IsKnownValuePrefix(prefix))
                {
                    return trimmed.Substring(idx + 1).Trim();
                }
            }

            return trimmed;
        }

        private static bool IsVerboseLoggingEnabled()
        {
            return string.Equals(
                Environment.GetEnvironmentVariable("MICROENG_QUICKCOLOUR_TRACE"),
                "1",
                StringComparison.OrdinalIgnoreCase);
        }

        private static VariantData BuildVariantData(string value, out string debugLabel)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                debugLabel = "string(empty)";
                return new VariantData("");
            }

            var trimmed = value.Trim();

            if (bool.TryParse(trimmed, out var boolValue))
            {
                debugLabel = "bool";
                return new VariantData(boolValue);
            }

            if (int.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
            {
                debugLabel = "int";
                return new VariantData(intValue);
            }

            if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var doubleValue))
            {
                debugLabel = "double";
                return new VariantData(doubleValue);
            }

            debugLabel = "string";
            return new VariantData(trimmed);
        }

        private static bool IsKnownValuePrefix(string prefix)
        {
            if (string.IsNullOrWhiteSpace(prefix))
            {
                return false;
            }

            return prefix.Equals("Int32", StringComparison.OrdinalIgnoreCase)
                   || prefix.Equals("Int64", StringComparison.OrdinalIgnoreCase)
                   || prefix.Equals("UInt32", StringComparison.OrdinalIgnoreCase)
                   || prefix.Equals("UInt64", StringComparison.OrdinalIgnoreCase)
                   || prefix.Equals("Double", StringComparison.OrdinalIgnoreCase)
                   || prefix.Equals("Single", StringComparison.OrdinalIgnoreCase)
                   || prefix.Equals("Float", StringComparison.OrdinalIgnoreCase)
                   || prefix.Equals("Decimal", StringComparison.OrdinalIgnoreCase)
                   || prefix.Equals("Boolean", StringComparison.OrdinalIgnoreCase)
                   || prefix.Equals("Bool", StringComparison.OrdinalIgnoreCase)
                   || prefix.Equals("DateTime", StringComparison.OrdinalIgnoreCase)
                   || prefix.Equals("Date", StringComparison.OrdinalIgnoreCase);
        }

        private static void ApplyScopeToSearch(
            Document doc,
            Search search,
            QuickColourScope scope,
            string selectionSetName,
            IReadOnlyList<List<string>> modelPaths,
            Action<string> log)
        {
            if (search == null)
            {
                return;
            }

            if (doc == null || scope == QuickColourScope.EntireModel)
            {
                search.Selection.SelectAll();
                return;
            }

            if (scope == QuickColourScope.PropertyFilter)
            {
                search.Selection.SelectAll();
                return;
            }

            if (scope == QuickColourScope.CurrentSelection)
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

            if (scope == QuickColourScope.SavedSelectionSet)
            {
                var name = selectionSetName;
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

            if (scope == QuickColourScope.ModelTree)
            {
                var items = ResolveScopeModelItems(doc, modelPaths, log);
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

        private static SearchCondition BuildScopeFilterCondition(
            QuickColourScope scope,
            string category,
            string property,
            string value,
            Action<string> log)
        {
            if (scope != QuickColourScope.PropertyFilter)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(category)
                || string.IsNullOrWhiteSpace(property)
                || string.IsNullOrWhiteSpace(value))
            {
                log?.Invoke("Scope property filter is incomplete. Using entire model.");
                return null;
            }

            return SearchCondition.HasPropertyByDisplayName(category, property)
                .EqualValue(new VariantData(value ?? ""));
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

        private static ModelItemCollection ResolveScopeModelItems(
            Document doc,
            IReadOnlyList<List<string>> paths,
            Action<string> log)
        {
            if (doc?.Models?.RootItems == null || paths == null || paths.Count == 0)
            {
                return null;
            }

            var items = new ModelItemCollection();
            var added = new HashSet<Guid>();
            foreach (var path in paths)
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
            if (roots == null || string.IsNullOrWhiteSpace(label))
            {
                return null;
            }

            foreach (var root in roots)
            {
                var name = root?.DisplayName ?? "";
                if (string.Equals(name, label, StringComparison.OrdinalIgnoreCase))
                {
                    return root;
                }
            }

            return null;
        }

        private static ModelItem FindChildByLabel(ModelItem parent, string label)
        {
            if (parent?.Children == null || string.IsNullOrWhiteSpace(label))
            {
                return null;
            }

            foreach (var child in parent.Children)
            {
                var name = child?.DisplayName ?? "";
                if (string.Equals(name, label, StringComparison.OrdinalIgnoreCase))
                {
                    return child;
                }
            }

            return null;
        }

        private static string CombinePath(params string[] parts)
        {
            if (parts == null || parts.Length == 0)
            {
                return string.Empty;
            }

            var list = new List<string>();
            foreach (var p in parts)
            {
                if (string.IsNullOrWhiteSpace(p)) continue;
                list.Add(p.Trim().Trim('/').Trim('\\'));
            }

            return string.Join("/", list);
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

        private static string TrimForName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "";
            }

            const int max = 80;
            value = value.Trim();
            if (value.Length <= max)
            {
                return value;
            }

            return value.Substring(0, max - 3) + "...";
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
    }
}
