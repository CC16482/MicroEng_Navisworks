using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;

namespace MicroEng.Navisworks.QuickColour
{
    public sealed class QuickColourNavisworksService
    {
        public int ApplyByHierarchy(
            Document doc,
            string l1Category,
            string l1Property,
            string l2Category,
            string l2Property,
            IEnumerable<QuickColourHierarchyGroup> groups,
            QuickColourScope scope,
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

            HashSet<Guid> scopeIds = null;
            if (scope == QuickColourScope.CurrentSelection)
            {
                var sel = doc.CurrentSelection?.SelectedItems;
                if (sel != null && sel.Count > 0)
                {
                    scopeIds = new HashSet<Guid>(
                        sel.Where(i => i != null)
                           .Select(i => i.InstanceGuid)
                           .Where(id => id != Guid.Empty));
                }
            }

            var totalColored = 0;

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
                    search.Selection.SelectAll();
                    search.SearchConditions.Add(
                        SearchCondition.HasPropertyByDisplayName(l1Category, l1Property)
                            .EqualValue(new VariantData(group.Value ?? "")));
                    search.SearchConditions.Add(
                        SearchCondition.HasPropertyByDisplayName(l2Category, l2Property)
                            .EqualValue(new VariantData(typeRow.Value ?? "")));

                    var results = search.FindAll(doc, false);
                    if (scopeIds != null)
                    {
                        results = FilterByIds(results, scopeIds);
                    }

                    if (results == null || results.Count == 0)
                    {
                        continue;
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

        private static ModelItemCollection FilterByIds(ModelItemCollection items, HashSet<Guid> allowed)
        {
            if (items == null || allowed == null || allowed.Count == 0)
            {
                return items;
            }

            var filtered = new ModelItemCollection();
            foreach (var item in items)
            {
                var id = item?.InstanceGuid ?? Guid.Empty;
                if (id != Guid.Empty && allowed.Contains(id))
                {
                    filtered.Add(item);
                }
            }
            return filtered;
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
