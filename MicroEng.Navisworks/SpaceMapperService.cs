using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Diagnostics;
using Autodesk.Navisworks.Api;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

namespace MicroEng.Navisworks
{
    internal class SpaceMapperRequest
    {
        public string ProfileName { get; set; }
        public SpaceMapperScope Scope { get; set; } = SpaceMapperScope.EntireModel;
        public ZoneSourceType ZoneSource { get; set; } = ZoneSourceType.DataScraperZones;
        public string ZoneSetName { get; set; }
        public TargetSourceType TargetSource { get; set; } = TargetSourceType.EntireModel;
        public List<SpaceMapperTargetRule> TargetRules { get; set; } = new();
        public List<SpaceMapperMappingDefinition> Mappings { get; set; } = new();
        public SpaceMapperProcessingSettings ProcessingSettings { get; set; } = new();
    }

    internal class SpaceMapperRunResult
    {
        public SpaceMapperRunStats Stats { get; set; } = new();
        public List<ZoneTargetIntersection> Intersections { get; set; } = new();
        public string Message { get; set; }
    }

    internal class SpaceMapperService
    {
        private readonly Action<string> _log;

        public SpaceMapperService(Action<string> log)
        {
            _log = log;
        }

        public SpaceMapperRunResult Run(SpaceMapperRequest request, CancellationToken token = default)
        {
            var sw = Stopwatch.StartNew();
            var result = new SpaceMapperRunResult();
            var doc = Application.ActiveDocument;
            if (doc == null)
            {
                result.Message = "No active document.";
                return result;
            }

            request.ProcessingSettings.ProcessingMode = SpaceMapperProcessingMode.CpuNormal;

            var session = GetSession(request.ProfileName);
            if (session == null)
            {
                result.Message = $"No Data Scraper sessions found for profile '{request.ProfileName}'.";
                return result;
            }

            DataScraperCache.LastSession = session;

            var zoneModels = ResolveZones(session, request.ZoneSource, request.ZoneSetName, doc).ToList();
            var targetsByRule = new Dictionary<string, List<SpaceMapperTargetRule>>(StringComparer.OrdinalIgnoreCase);
            var targetGeometries = ResolveTargets(doc, request.TargetSource, request.TargetRules, targetsByRule).ToList();

            if (!zoneModels.Any())
            {
                result.Message = "No zones found.";
                return result;
            }

            if (!targetGeometries.Any())
            {
                result.Message = "No targets found for the selected rules.";
                return result;
            }

            var zones = zoneModels
                .Select(z => GeometryExtractor.ExtractZoneGeometry(z.ModelItem, z.ItemKey, z.DisplayName, request.ProcessingSettings))
                .Where(z => z?.BoundingBox != null && z.Vertices.Any())
                .ToList();

            var targets = targetGeometries
                .Select(t => GeometryExtractor.ExtractTargetGeometry(t.ModelItem, t.ItemKey, t.DisplayName))
                .Where(t => t?.BoundingBox != null && t.Vertices.Any())
                .ToList();

            var engine = SpaceMapperEngineFactory.Create(request.ProcessingSettings.ProcessingMode);
            var intersections = engine.ComputeIntersections(zones, targets, request.ProcessingSettings, null, token) ?? new List<ZoneTargetIntersection>();
            result.Intersections = intersections.ToList();

            var stats = new SpaceMapperRunStats { ZonesProcessed = zones.Count, TargetsProcessed = targets.Count, ModeUsed = engine.Mode.ToString() };
            var zoneLookup = zones.ToDictionary(z => z.ZoneId, z => z);
            var targetLookup = targets.ToDictionary(t => t.ItemKey, t => t);

            var ruleMembership = BuildRuleMembership(targetsByRule);
            var zoneValueLookup = new ZoneValueLookup(session);

            foreach (var group in intersections.GroupBy(i => i.TargetItemKey, StringComparer.OrdinalIgnoreCase))
            {
                if (!targetLookup.TryGetValue(group.Key, out var tgt)) continue;
                var rulesForTarget = ruleMembership.TryGetValue(group.Key, out var list) ? list : new List<SpaceMapperTargetRule>();
                if (!rulesForTarget.Any()) continue;

                var relevant = FilterByMembership(group, rulesForTarget, request.ProcessingSettings);
                if (!relevant.Any())
                {
                    stats.Skipped++;
                    continue;
                }

                var isMultiZone = relevant.Count() > 1;
                if (isMultiZone) stats.MultiZoneTagged++;

                foreach (var mapping in request.Mappings)
                {
                    var values = new List<string>();
                    foreach (var inter in relevant)
                    {
                        if (!zoneLookup.TryGetValue(inter.ZoneId, out var zone)) continue;
                        var val = zoneValueLookup.GetValue(inter.ZoneId, mapping.ZoneCategory, mapping.ZonePropertyName);
                        if (inter.IsPartial && request.ProcessingSettings.TagPartialSeparately && !string.IsNullOrWhiteSpace(mapping.PartialFlagValue))
                        {
                            val = string.IsNullOrWhiteSpace(val)
                                ? mapping.PartialFlagValue
                                : $"{val}{mapping.AppendSeparator}{mapping.PartialFlagValue}";
                        }
                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            values.Add(val);
                        }
                    }

                    var combined = CombineValues(values, mapping.MultiZoneCombineMode, mapping.AppendSeparator);
                    if (string.IsNullOrWhiteSpace(combined)) continue;
                    PropertyWriter.WriteProperty(tgt.ModelItem, mapping.TargetCategory, mapping.TargetPropertyName, combined, mapping.WriteMode, mapping.AppendSeparator);
                }

                foreach (var inter in relevant)
                {
                    if (inter.IsContained) stats.ContainedTagged++;
                    if (inter.IsPartial) stats.PartialTagged++;
                    if (zoneLookup.TryGetValue(inter.ZoneId, out var z))
                    {
                        var summary = stats.ZoneSummaries.FirstOrDefault(s => string.Equals(s.ZoneId, z.ZoneId, StringComparison.OrdinalIgnoreCase));
                        if (summary == null)
                        {
                            summary = new ZoneSummary { ZoneId = z.ZoneId, ZoneName = z.DisplayName };
                            stats.ZoneSummaries.Add(summary);
                        }
                        if (inter.IsContained) summary.ContainedCount++; else summary.PartialCount++;
                    }
                }
            }

            result.Stats = stats;
            sw.Stop();
            result.Stats.Elapsed = sw.Elapsed;
            result.Message = $"Space Mapper: {stats.ZonesProcessed} zones, {stats.TargetsProcessed} targets, {stats.ContainedTagged} contained, {stats.PartialTagged} partial. Mode={stats.ModeUsed}.";
            _log?.Invoke(result.Message);
            return result;
        }

        private static Dictionary<string, List<SpaceMapperTargetRule>> BuildRuleMembership(Dictionary<string, List<SpaceMapperTargetRule>> targetsByRule)
        {
            var map = new Dictionary<string, List<SpaceMapperTargetRule>>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in targetsByRule)
            {
                if (!map.TryGetValue(kvp.Key, out var list))
                {
                    list = new List<SpaceMapperTargetRule>();
                    map[kvp.Key] = list;
                }
                list.AddRange(kvp.Value);
            }
            return map;
        }

        private static IEnumerable<ZoneTargetIntersection> FilterByMembership(IEnumerable<ZoneTargetIntersection> intersections,
            List<SpaceMapperTargetRule> rules, SpaceMapperProcessingSettings settings)
        {
            var allowContained = rules.Any(r => r.MembershipMode != SpaceMembershipMode.PartialOnly);
            var allowPartial = rules.Any(r => r.MembershipMode != SpaceMembershipMode.ContainedOnly);

            foreach (var inter in intersections)
            {
                var contained = inter.IsContained || (settings.TreatPartialAsContained && inter.IsPartial);
                var partial = inter.IsPartial && !settings.TreatPartialAsContained;

                if (contained && allowContained)
                {
                    yield return new ZoneTargetIntersection
                    {
                        ZoneId = inter.ZoneId,
                        TargetItemKey = inter.TargetItemKey,
                        IsContained = true,
                        IsPartial = inter.IsPartial,
                        OverlapVolume = inter.OverlapVolume
                    };
                }
                else if (partial && allowPartial)
                {
                    yield return inter;
                }
            }
        }

        private static string CombineValues(List<string> values, MultiZoneCombineMode mode, string sep)
        {
            if (values == null || values.Count == 0) return string.Empty;
            switch (mode)
            {
                case MultiZoneCombineMode.First:
                    return values.First();
                case MultiZoneCombineMode.Concatenate:
                    return string.Join(string.IsNullOrWhiteSpace(sep) ? ", " : sep, values);
                case MultiZoneCombineMode.Min:
                    return values.Select(TryDouble).Where(v => v.HasValue).DefaultIfEmpty(null).Min()?.ToString() ?? values.First();
                case MultiZoneCombineMode.Max:
                    return values.Select(TryDouble).Where(v => v.HasValue).DefaultIfEmpty(null).Max()?.ToString() ?? values.First();
                case MultiZoneCombineMode.Average:
                    var nums = values.Select(TryDouble).Where(v => v.HasValue).Select(v => v.Value).ToList();
                    return nums.Any() ? nums.Average().ToString("0.###") : values.First();
                default:
                    return values.First();
            }
        }

        private static double? TryDouble(string value)
        {
            return double.TryParse(value, out var d) ? d : (double?)null;
        }

        private static ScrapeSession GetSession(string profile)
        {
            var name = string.IsNullOrWhiteSpace(profile) ? "Default" : profile;
            return DataScraperCache.AllSessions
                .Where(s => string.Equals(s.ProfileName ?? "Default", name, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(s => s.Timestamp)
                .FirstOrDefault();
        }

        private static string GetItemKey(ModelItem item)
        {
            try
            {
                return item.InstanceGuid.ToString();
            }
            catch
            {
                return Guid.NewGuid().ToString();
            }
        }

        private static IEnumerable<ResolvedItem> ResolveZones(ScrapeSession session, ZoneSourceType source, string setName, Document doc)
        {
            switch (source)
            {
                case ZoneSourceType.ZoneSelectionSet:
                case ZoneSourceType.ZoneSearchSet:
                    // Placeholder: fall back to selection for now
                    foreach (ModelItem item in doc.CurrentSelection?.SelectedItems ?? new ModelItemCollection())
                        yield return new ResolvedItem { ItemKey = GetItemKey(item), ModelItem = item, DisplayName = item.DisplayName };
                    break;
                case ZoneSourceType.DataScraperZones:
                default:
                    if (session == null) yield break;
                    var zoneKeys = session.RawEntries
                        .Where(r => IsZoneLike(r.Category) || IsZoneLike(r.Name))
                        .Select(r => r.ItemKey)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    foreach (var key in zoneKeys)
                    {
                        if (!Guid.TryParse(key, out var guid)) continue;
                        var item = FindByGuid(doc, guid);
                        if (item == null) continue;
                        yield return new ResolvedItem
                        {
                            ItemKey = key,
                            ModelItem = item,
                            DisplayName = item.DisplayName
                        };
                    }
                    break;
            }
        }

        private static bool IsZoneLike(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            var val = value.ToLowerInvariant();
            return val.Contains("zone") || val.Contains("room") || val.Contains("space");
        }

        private static IEnumerable<TargetGeometry> ResolveTargets(Document doc, TargetSourceType targetSource,
            IEnumerable<SpaceMapperTargetRule> rules,
            Dictionary<string, List<SpaceMapperTargetRule>> targetsByRule)
        {
            var added = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var rule in rules.Where(r => r.Enabled))
            {
                var items = ResolveTargetsForRule(doc, targetSource, rule);
                foreach (var item in items)
                {
                    var key = GetItemKey(item);
                    if (!targetsByRule.TryGetValue(key, out var list))
                    {
                        list = new List<SpaceMapperTargetRule>();
                        targetsByRule[key] = list;
                    }
                    list.Add(rule);

                    if (added.Add(key))
                    {
                        yield return new TargetGeometry
                        {
                            ItemKey = key,
                            ModelItem = item,
                            DisplayName = item.DisplayName
                        };
                    }
                }
            }
        }

        private static IEnumerable<ModelItem> ResolveTargetsForRule(Document doc, TargetSourceType targetSource, SpaceMapperTargetRule rule)
        {
            switch (rule.TargetType)
            {
                case SpaceMapperTargetType.CurrentSelection:
                    return doc.CurrentSelection?.SelectedItems?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
                case SpaceMapperTargetType.SelectionTreeLevel:
                    return TraverseWithDepth(doc.Models.RootItems, rule.MinTreeLevel ?? 0, rule.MaxTreeLevel ?? int.MaxValue)
                        .Where(mi => CategoryMatches(mi, rule.CategoryFilter));
                case SpaceMapperTargetType.VisibleInView:
                    // Fallback to current selection for visible items
                    return doc.CurrentSelection?.SelectedItems?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
                case SpaceMapperTargetType.SelectionSet:
                case SpaceMapperTargetType.SearchSet:
                    // Selection/search set resolution placeholder: fall back to current selection
                    return doc.CurrentSelection?.SelectedItems?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
                default:
                    return ResolveByTargetSource(doc, targetSource);
            }
        }

        private static IEnumerable<ModelItem> ResolveByTargetSource(Document doc, TargetSourceType targetSource)
        {
            switch (targetSource)
            {
                case TargetSourceType.SelectionSet:
                    return doc.CurrentSelection?.SelectedItems?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
                case TargetSourceType.SearchSet:
                    // Placeholder: search set resolution not implemented, fallback to current selection
                    return doc.CurrentSelection?.SelectedItems?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
                case TargetSourceType.ViewpointVisible:
                case TargetSourceType.Visible:
                    return TraverseAll(doc.Models.RootItems).Where(mi => !IsHidden(mi));
                case TargetSourceType.Hidden:
                    return TraverseAll(doc.Models.RootItems).Where(IsHidden);
                case TargetSourceType.EntireModel:
                default:
                    return TraverseAll(doc.Models.RootItems);
            }
        }

        private static bool IsHidden(ModelItem item)
        {
            try
            {
                var prop = item.GetType().GetProperty("IsHidden");
                if (prop != null && prop.PropertyType == typeof(bool))
                {
                    return (bool)prop.GetValue(item);
                }
            }
            catch
            {
            }
            return false;
        }

        private static IEnumerable<ModelItem> TraverseAll(IEnumerable<ModelItem> items)
        {
            foreach (ModelItem item in items)
            {
                yield return item;
                if (item.Children != null && item.Children.Any())
                {
                    foreach (var child in TraverseAll(item.Children))
                        yield return child;
                }
            }
        }

        private static ModelItem FindByGuid(Document doc, Guid guid)
        {
            foreach (var item in TraverseAll(doc.Models.RootItems))
            {
                try
                {
                    if (item.InstanceGuid == guid)
                        return item;
                }
                catch
                {
                    // ignore
                }
            }
            return null;
        }

        private static IEnumerable<ModelItem> TraverseWithDepth(IEnumerable<ModelItem> items, int minDepth, int maxDepth, int depth = 0)
        {
            foreach (ModelItem item in items)
            {
                if (depth >= minDepth && depth <= maxDepth)
                {
                    yield return item;
                }
                if (item.Children != null && item.Children.Any())
                {
                    foreach (var child in TraverseWithDepth(item.Children, minDepth, maxDepth, depth + 1))
                    {
                        yield return child;
                    }
                }
            }
        }

        private static bool CategoryMatches(ModelItem item, string filter)
        {
            if (string.IsNullOrWhiteSpace(filter)) return true;
            var f = filter.ToLowerInvariant();
            try
            {
                foreach (var cat in item.PropertyCategories)
                {
                    if (cat == null) continue;
                    if ((cat.DisplayName ?? cat.Name ?? string.Empty).ToLowerInvariant().Contains(f))
                        return true;
                }
                return (item.DisplayName ?? string.Empty).ToLowerInvariant().Contains(f);
            }
            catch
            {
                return true;
            }
        }

        private class ResolvedItem
        {
            public string ItemKey { get; set; }
            public string DisplayName { get; set; }
            public ModelItem ModelItem { get; set; }
        }
    }

    internal class ZoneValueLookup
    {
        private readonly Dictionary<string, Dictionary<string, Dictionary<string, string>>> _values;

        public ZoneValueLookup(ScrapeSession session)
        {
            _values = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>(StringComparer.OrdinalIgnoreCase);
            if (session?.RawEntries == null) return;

            foreach (var entry in session.RawEntries)
            {
                if (!_values.TryGetValue(entry.ItemKey, out var catDict))
                {
                    catDict = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                    _values[entry.ItemKey] = catDict;
                }

                var category = entry.Category ?? string.Empty;
                if (!catDict.TryGetValue(category, out var propDict))
                {
                    propDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    catDict[category] = propDict;
                }

                var propName = entry.Name ?? string.Empty;
                if (!propDict.ContainsKey(propName))
                {
                    propDict[propName] = entry.Value ?? string.Empty;
                }
            }
        }

        public string GetValue(string itemKey, string category, string property)
        {
            if (itemKey == null) return string.Empty;
            if (!_values.TryGetValue(itemKey, out var catDict)) return string.Empty;

            if (!string.IsNullOrWhiteSpace(category))
            {
                if (catDict.TryGetValue(category, out var propDict))
                {
                    if (!string.IsNullOrWhiteSpace(property) && propDict.TryGetValue(property, out var val))
                        return val;
                }
            }

            if (!string.IsNullOrWhiteSpace(property))
            {
                foreach (var props in catDict.Values)
                {
                    if (props.TryGetValue(property, out var val))
                        return val;
                }
            }

            return string.Empty;
        }
    }

    internal static class PropertyWriter
    {
        public static void WriteProperty(ModelItem item, string categoryName, string propertyName, string value, WriteMode mode, string appendSeparator)
        {
            if (item == null || string.IsNullOrWhiteSpace(categoryName) || string.IsNullOrWhiteSpace(propertyName)) return;

            var existing = ReadProperty(item, categoryName, propertyName);
            if (mode == WriteMode.OnlyIfBlank && !string.IsNullOrWhiteSpace(existing))
            {
                return;
            }

            var finalValue = value;
            if (mode == WriteMode.Append && !string.IsNullOrWhiteSpace(existing))
            {
                finalValue = string.IsNullOrWhiteSpace(appendSeparator)
                    ? $"{existing},{value}"
                    : $"{existing}{appendSeparator}{value}";
            }

            try
            {
                var state = ComBridge.State;
                var path = ComBridge.ToInwOaPath(item);
                var propertyNode = (ComApi.InwGUIPropertyNode2)state.GetGUIPropertyNode(path, true);

                var propertyVector = (ComApi.InwOaPropertyVec)state.ObjectFactory(
                    ComApi.nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

                var newProp = (ComApi.InwOaProperty)state.ObjectFactory(
                    ComApi.nwEObjectType.eObjectType_nwOaProperty, null, null);
                newProp.name = propertyName;
                newProp.UserName = propertyName;
                newProp.value = finalValue;

                propertyVector.Properties().Add(newProp);
                propertyNode.SetUserDefined(0, categoryName, categoryName, propertyVector);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[SpaceMapper] Failed to write {categoryName}.{propertyName} on {item.DisplayName}: {ex.Message}");
            }
        }

        private static string ReadProperty(ModelItem item, string categoryName, string propertyName)
        {
            try
            {
                foreach (var cat in item.PropertyCategories)
                {
                    if (!string.Equals(cat.DisplayName ?? cat.Name, categoryName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    foreach (var prop in cat.Properties)
                    {
                        if (string.Equals(prop.DisplayName ?? prop.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                        {
                            return prop.Value?.ToDisplayString() ?? prop.Value?.ToString() ?? string.Empty;
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }
            return string.Empty;
        }
    }
}
