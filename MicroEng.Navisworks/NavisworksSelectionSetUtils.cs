using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;

namespace MicroEng.Navisworks
{
    internal static class NavisworksSelectionSetUtils
    {
        public static IReadOnlyList<string> GetSelectionSetNames(Document doc)
        {
            return EnumerateSelectionSets(doc)
                .Where(ss => !IsSearchSet(ss))
                .Select(ss => SafeDisplayName(ss))
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n)
                .ToList();
        }

        public static IReadOnlyList<string> GetSearchSetNames(Document doc)
        {
            return EnumerateSelectionSets(doc)
                .Where(ss => IsSearchSet(ss))
                .Select(ss => SafeDisplayName(ss))
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n)
                .ToList();
        }

        public static IEnumerable<ModelItem> GetItemsFromSet(
            Document doc,
            string setName,
            bool expectSearchSet,
            out string description)
        {
            description = string.Empty;

            if (doc == null)
            {
                description = "No active document.";
                return Enumerable.Empty<ModelItem>();
            }

            if (string.IsNullOrWhiteSpace(setName))
            {
                description = expectSearchSet ? "Search Set: (none selected)" : "Selection Set: (none selected)";
                return Enumerable.Empty<ModelItem>();
            }

            var match = EnumerateSelectionSets(doc)
                .FirstOrDefault(ss =>
                    string.Equals(SafeDisplayName(ss), setName, StringComparison.OrdinalIgnoreCase) &&
                    IsSearchSet(ss) == expectSearchSet);

            if (match == null)
            {
                description = (expectSearchSet ? "Search Set" : "Selection Set") + $": {setName} (not found)";
                return Enumerable.Empty<ModelItem>();
            }

            try
            {
                var items = match.GetSelectedItems();
                var count = items?.Count ?? 0;
                description = (expectSearchSet ? "Search Set" : "Selection Set") + $": {SafeDisplayName(match)} ({count} items)";
                return items?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
            }
            catch (Exception ex)
            {
                description = (expectSearchSet ? "Search Set" : "Selection Set") + $": {SafeDisplayName(match)} (failed: {ex.Message})";
                return Enumerable.Empty<ModelItem>();
            }
        }

        public static bool IsSearchSet(SelectionSet set)
        {
            try
            {
                return set?.HasSearch == true;
            }
            catch
            {
                return false;
            }
        }

        private static IEnumerable<SelectionSet> EnumerateSelectionSets(Document doc)
        {
            if (doc?.SelectionSets?.RootItem == null)
            {
                yield break;
            }

            foreach (var ss in EnumerateSelectionSets(doc.SelectionSets.RootItem))
            {
                yield return ss;
            }
        }

        private static IEnumerable<SelectionSet> EnumerateSelectionSets(GroupItem group)
        {
            if (group?.Children == null)
            {
                yield break;
            }

            foreach (SavedItem child in group.Children)
            {
                if (child is GroupItem g)
                {
                    foreach (var ss in EnumerateSelectionSets(g))
                    {
                        yield return ss;
                    }
                }
                else if (child is SelectionSet ss)
                {
                    yield return ss;
                }
            }
        }

        private static string SafeDisplayName(SavedItem item)
        {
            try
            {
                return item?.DisplayName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
