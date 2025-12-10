using System;
using System.Collections.Generic;
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
            var session = new ScrapeSession
            {
                ProfileName = profileName,
                ScopeType = scopeType.ToString(),
                ScopeDescription = scopeDescription
            };

            var propertyMap = new Dictionary<string, ScrapedProperty>(StringComparer.OrdinalIgnoreCase);
            var raw = new List<RawEntry>();

            foreach (var item in items)
            {
                session.ItemsScanned++;
                var itemPath = item.TryGetDisplayName();
                foreach (var category in item.PropertyCategories)
                {
                    if (category == null) continue;
                    foreach (var prop in category.Properties)
                    {
                        if (prop == null) continue;
                        var cat = category.DisplayName ?? category.Name ?? string.Empty;
                        var name = prop.DisplayName ?? prop.Name ?? string.Empty;
                        var key = $"{cat}|{name}";
                        if (!propertyMap.TryGetValue(key, out var sp))
                        {
                            sp = new ScrapedProperty
                            {
                                Category = cat,
                                Name = name,
                                DataType = prop.Value?.DataType.ToString() ?? "Unknown",
                                ItemCount = 0
                            };
                            propertyMap[key] = sp;
                        }
                        sp.ItemCount++;

                        string display = string.Empty;
                        try
                        {
                            display = prop.Value?.IsDisplayString == true
                                ? prop.Value.ToDisplayString()
                                : prop.Value?.ToString();
                            if (!string.IsNullOrWhiteSpace(display) && !sp.SampleValues.Contains(display) && sp.SampleValues.Count < 10)
                            {
                                sp.SampleValues.Add(display);
                            }
                        }
                        catch
                        {
                            // ignore sample capture errors
                        }

                        raw.Add(new RawEntry
                        {
                            Profile = profileName,
                            Scope = scopeDescription,
                            ItemKey = item.InstanceGuid.ToString(),
                            ItemPath = itemPath,
                            Category = cat,
                            Name = name,
                            DataType = prop.Value?.DataType.ToString() ?? "Unknown",
                            Value = display
                        });
                    }
                }
            }

            foreach (var sp in propertyMap.Values)
            {
                sp.DistinctValueCount = sp.SampleValues.Count;
            }

            session.Properties = propertyMap.Values.OrderBy(p => p.Category).ThenBy(p => p.Name).ToList();
            session.RawEntries = raw;
            DataScraperCache.AddSession(session);
            return session;
        }

        public IEnumerable<ModelItem> ResolveScope(ScrapeScopeType scopeType, string selectionSetName, string searchSetName, out string description)
        {
            description = string.Empty;
            var doc = NavisApp.ActiveDocument;
            if (doc == null) return Enumerable.Empty<ModelItem>();

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
                    var sel = doc.CurrentSelection?.SelectedItems?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
                    description = $"Current Selection ({sel.Count()} items)";
                    return sel;
                case ScrapeScopeType.SelectionSet:
                    description = $"Selection Set: {selectionSetName} (not yet implemented)";
                    return Enumerable.Empty<ModelItem>();
                case ScrapeScopeType.SearchSet:
                    description = $"Search Set: {searchSetName} (not yet implemented)";
                    return Enumerable.Empty<ModelItem>();
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
                if (item.Children != null && item.Children.Any())
                {
                    foreach (var child in Traverse(item.Children))
                        yield return child;
                }
            }
        }

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
