using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Navisworks.Api;
using NavisApp = Autodesk.Navisworks.Api.Application;

namespace MicroEng.Navisworks
{
    internal enum ScrapeScopeType
    {
        SingleItem,
        CurrentSelection,
        SelectionSet,
        SearchSet,
        EntireModel
    }

    internal class DataScraperService
    {
        public ScrapeSession Scrape(string profileName, ScrapeScopeType scopeType, string scopeDescription, IEnumerable<ModelItem> items)
        {
            var sourceItems = items as ICollection<ModelItem>;
            var estimatedItemCount = sourceItems?.Count ?? 0;
            var estimatedRawCapacity = estimatedItemCount > 0 ? estimatedItemCount * 8 : 0;

            var session = new ScrapeSession
            {
                ProfileName = profileName,
                ScopeType = scopeType.ToString(),
                ScopeDescription = scopeDescription,
                Timestamp = DateTime.Now
            };

            var doc = NavisApp.ActiveDocument;
            var documentFile = doc?.FileName ?? string.Empty;
            session.DocumentFile = documentFile;
            session.DocumentFileKey = BuildDocumentFileKey(documentFile);
            session.RawEntries = estimatedRawCapacity > 0
                ? new List<RawEntry>(estimatedRawCapacity)
                : new List<RawEntry>();

            var propertyMap = new Dictionary<string, PropertyAccumulator>(
                estimatedItemCount > 0 ? estimatedItemCount : 0,
                StringComparer.OrdinalIgnoreCase);

            foreach (var item in items ?? Enumerable.Empty<ModelItem>())
            {
                if (item == null)
                {
                    continue;
                }

                session.ItemsScanned++;

                var itemPath = ItemToPath(item);
                var itemKey = item.InstanceGuid.ToString("D");
                var seenPropsThisItem = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var category in item.PropertyCategories)
                {
                    if (category == null)
                    {
                        continue;
                    }

                    var catName = category.DisplayName ?? category.Name ?? string.Empty;

                    foreach (var prop in category.Properties)
                    {
                        if (prop == null)
                        {
                            continue;
                        }

                        var name = prop.DisplayName ?? prop.Name ?? string.Empty;
                        var dtype = prop.Value?.DataType.ToString() ?? "Unknown";
                        var values = GetPropertyValueStrings(prop);
                        if (values.Count == 0)
                        {
                            continue;
                        }

                        values = NormalizeValues(values);

                        if (values.Count == 0)
                        {
                            continue;
                        }

                        var key = $"{catName}\u001F{name}\u001F{dtype}";
                        var firstThisItem = seenPropsThisItem.Add(key);
                        TouchProperty(propertyMap, key, catName, name, dtype, firstThisItem, values);

                        foreach (var v in values)
                        {
                            session.RawEntries.Add(new RawEntry
                            {
                                Profile = profileName,
                                Scope = scopeDescription,
                                ItemKey = itemKey,
                                ItemPath = itemPath,
                                Category = catName,
                                Name = name,
                                DataType = dtype,
                                Value = v
                            });
                        }
                    }
                }
            }

            session.Properties = propertyMap.Values
                .Select(a => new ScrapedProperty
                {
                    Category = a.Category,
                    Name = a.Name,
                    DataType = a.DataType,
                    ItemCount = a.ItemCount,
                    DistinctValueCount = a.Distinct.Count,
                    SampleValues = a.Sample
                })
                .OrderBy(p => p.Category)
                .ThenBy(p => p.Name)
                .ToList();

            DataScraperCache.AddSession(session);
            return session;
        }

        public IEnumerable<ModelItem> ResolveScope(ScrapeScopeType scopeType, string selectionSetName, string searchSetName, out string description)
        {
            description = string.Empty;
            var doc = NavisApp.ActiveDocument;
            if (doc == null)
            {
                return Enumerable.Empty<ModelItem>();
            }

            switch (scopeType)
            {
                case ScrapeScopeType.SingleItem:
                    var single = doc.CurrentSelection?.SelectedItems?.FirstOrDefault();
                    if (single != null)
                    {
                        description = $"Single Item: {single.DisplayName}";
                        return new[] { single };
                    }
                    return Enumerable.Empty<ModelItem>();
                case ScrapeScopeType.CurrentSelection:
                    var selection = doc.CurrentSelection?.SelectedItems?.Cast<ModelItem>()?.ToList()
                                   ?? new List<ModelItem>();
                    description = $"Current Selection ({selection.Count} items)";
                    return selection;
                case ScrapeScopeType.SelectionSet:
                {
                    var items = NavisworksSelectionSetUtils.GetItemsFromSet(doc, selectionSetName, expectSearchSet: false, out var desc);
                    description = desc;
                    return items;
                }
                case ScrapeScopeType.SearchSet:
                {
                    var items = NavisworksSelectionSetUtils.GetItemsFromSet(doc, searchSetName, expectSearchSet: true, out var desc);
                    description = desc;
                    return items;
                }
                case ScrapeScopeType.EntireModel:
                    description = "Entire Model";
                    return Traverse(doc.Models.RootItems);
                default:
                    return Enumerable.Empty<ModelItem>();
            }
        }

        private static IEnumerable<ModelItem> Traverse(IEnumerable<ModelItem> items)
        {
            foreach (ModelItem item in items)
            {
                yield return item;
                if (item.Children != null)
                {
                    foreach (var child in Traverse(item.Children))
                    {
                        yield return child;
                    }
                }
            }
        }

        private static string BuildDocumentFileKey(string documentFile)
        {
            if (string.IsNullOrWhiteSpace(documentFile))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFileName(documentFile) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ItemToPath(ModelItem item)
        {
            if (item == null)
            {
                return string.Empty;
            }

            try
            {
                var names = new List<string>();
                var current = item;
                var guard = 0;
                while (current != null && guard++ < 128)
                {
                    var name = current.TryGetDisplayName();
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        names.Add(name);
                    }

                    current = current.Parent;
                }

                names.Reverse();
                if (names.Count == 0)
                {
                    return item.InstanceGuid.ToString("D");
                }

                return string.Join(" / ", names);
            }
            catch
            {
                return item.TryGetDisplayName();
            }
        }

        private static List<string> GetPropertyValueStrings(DataProperty prop)
        {
            var values = new List<string>();
            if (prop?.Value == null)
            {
                return values;
            }

            try
            {
                if (prop.Value.IsDisplayString)
                {
                    var display = prop.Value.ToDisplayString() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(display))
                    {
                        values.Add(display);
                    }

                    return values;
                }
            }
            catch
            {
                // ignore display string failures
            }

            try
            {
                var text = prop.Value.ToString() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    values.Add(text);
                }
            }
            catch
            {
                // ignore
            }

            return values;
        }

        private static List<string> NormalizeValues(List<string> values)
        {
            if (values == null || values.Count == 0)
            {
                return new List<string>();
            }

            var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var normalized = new List<string>(values.Count);
            foreach (var value in values)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                if (unique.Add(value))
                {
                    normalized.Add(value);
                }
            }

            return normalized;
        }

        private static void TouchProperty(
            Dictionary<string, PropertyAccumulator> map,
            string propertyKey,
            string category,
            string name,
            string dataType,
            bool firstThisItem,
            List<string> values)
        {
            if (!map.TryGetValue(propertyKey, out var acc))
            {
                acc = new PropertyAccumulator
                {
                    Category = category,
                    Name = name,
                    DataType = dataType
                };
                map[propertyKey] = acc;
            }

            if (firstThisItem)
            {
                acc.ItemCount++;
            }

            foreach (var v in values)
            {
                if (acc.Distinct.Count < 5000)
                {
                    acc.Distinct.Add(v);
                }

                if (acc.Sample.Count < 12 && acc.SampleSet.Add(v))
                {
                    acc.Sample.Add(v);
                }
            }
        }
    }

    internal sealed class PropertyAccumulator
    {
        public string Category { get; set; }
        public string Name { get; set; }
        public string DataType { get; set; }
        public int ItemCount { get; set; }
        public HashSet<string> Distinct { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<string> Sample { get; } = new();
        public HashSet<string> SampleSet { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    internal static class ModelItemExtensions
    {
        public static string TryGetDisplayName(this ModelItem item)
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
