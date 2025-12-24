using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Diagnostics;
using Autodesk.Navisworks.Api;
using MicroEng.Navisworks.SpaceMapper.Estimation;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

namespace MicroEng.Navisworks
{
    internal class SpaceMapperRequest
    {
        public string TemplateName { get; set; }
        public string ScraperProfileName { get; set; }
        public SpaceMapperScope Scope { get; set; } = SpaceMapperScope.EntireModel;
        public ZoneSourceType ZoneSource { get; set; } = ZoneSourceType.DataScraperZones;
        public string ZoneSetName { get; set; }
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

        public SpaceMapperRunResult Run(SpaceMapperRequest request, SpaceMapperPreflightCache preflightCache = null, CancellationToken token = default)
        {
            var sw = Stopwatch.StartNew();
            var result = new SpaceMapperRunResult();
            var stats = new SpaceMapperRunStats();
            var doc = Application.ActiveDocument;
            if (doc == null)
            {
                result.Message = "No active document.";
                return result;
            }

            var session = GetSession(request.ScraperProfileName);
            var requiresSession = request.ZoneSource == ZoneSourceType.DataScraperZones;
            if (requiresSession && session == null)
            {
                result.Message = $"No Data Scraper sessions found for profile '{request.ScraperProfileName ?? "Default"}'.";
                return result;
            }

            if (session != null)
            {
                DataScraperCache.LastSession = session;
            }

            var resolveSw = Stopwatch.StartNew();
            var zoneModels = ResolveZones(session, request.ZoneSource, request.ZoneSetName, doc).ToList();
            var targetsByRule = new Dictionary<string, List<SpaceMapperTargetRule>>(StringComparer.OrdinalIgnoreCase);
            var targetModels = ResolveTargets(doc, request.TargetRules, targetsByRule).ToList();
            resolveSw.Stop();
            stats.ResolveTime = resolveSw.Elapsed;

            if (!zoneModels.Any())
            {
                result.Message = "No zones found.";
                return result;
            }

            if (!targetModels.Any())
            {
                result.Message = "No targets found for the selected rules.";
                return result;
            }

            var buildSw = Stopwatch.StartNew();
            var zones = zoneModels
                .Select(z => GeometryExtractor.ExtractZoneGeometry(z.ModelItem, z.ItemKey, z.DisplayName, request.ProcessingSettings))
                .Where(z => z?.BoundingBox != null && z.Vertices.Any())
                .ToList();

            var cacheToUse = TryGetReusablePreflightCache(request, preflightCache, targetModels);
            List<TargetGeometry> targetsForEngine;
            if (cacheToUse != null)
            {
                targetsForEngine = targetModels;
            }
            else
            {
                targetsForEngine = targetModels
                    .Select(t =>
                    {
                        var bbox = t.ModelItem?.BoundingBox();
                        if (bbox == null) return null;
                        return new TargetGeometry
                        {
                            ItemKey = t.ItemKey,
                            DisplayName = t.DisplayName,
                            ModelItem = t.ModelItem,
                            BoundingBox = bbox
                        };
                    })
                    .Where(t => t != null)
                    .ToList();
            }
            buildSw.Stop();
            stats.BuildGeometryTime = buildSw.Elapsed;

            if (!zones.Any())
            {
                result.Message = "No zones found.";
                return result;
            }

            if (cacheToUse == null && !targetsForEngine.Any())
            {
                result.Message = "No targets with bounding boxes found.";
                return result;
            }

            var engine = SpaceMapperEngineFactory.Create(request.ProcessingSettings.ProcessingMode);
            var diagnostics = new SpaceMapperEngineDiagnostics();
            var intersections = engine.ComputeIntersections(zones, targetsForEngine, request.ProcessingSettings, cacheToUse, diagnostics, null, token)
                ?? new List<ZoneTargetIntersection>();
            result.Intersections = intersections.ToList();

            stats.ZonesProcessed = zones.Count;
            stats.TargetsProcessed = cacheToUse?.TargetKeys?.Length ?? targetsForEngine.Count;
            stats.ModeUsed = engine.Mode.ToString();
            stats.PresetUsed = diagnostics.PresetUsed;
            stats.CandidatePairs = diagnostics.CandidatePairs;
            stats.AvgCandidatesPerZone = diagnostics.AvgCandidatesPerZone;
            stats.MaxCandidatesPerZone = diagnostics.MaxCandidatesPerZone;
            stats.UsedPreflightIndex = diagnostics.UsedPreflightIndex;
            stats.BuildIndexTime = diagnostics.BuildIndexTime;
            stats.CandidateQueryTime = diagnostics.CandidateQueryTime;
            stats.NarrowPhaseTime = diagnostics.NarrowPhaseTime;

            var zoneLookup = zones.ToDictionary(z => z.ZoneId, z => z);
            var targetLookup = targetModels.ToDictionary(t => t.ItemKey, t => t);

            var ruleMembership = BuildRuleMembership(targetsByRule);
            var zoneValueLookup = new ZoneValueLookup(session);

            var writeSw = Stopwatch.StartNew();
            foreach (var group in intersections.GroupBy(i => i.TargetItemKey, StringComparer.OrdinalIgnoreCase))
            {
                if (!targetLookup.TryGetValue(group.Key, out var tgt)) continue;
                var rulesForTarget = ruleMembership.TryGetValue(group.Key, out var list) ? list : new List<SpaceMapperTargetRule>();
                if (!rulesForTarget.Any()) continue;

                var relevant = FilterByMembership(group, rulesForTarget, request.ProcessingSettings).ToList();
                if (relevant.Count == 0)
                {
                    stats.Skipped++;
                    continue;
                }

                if (!request.ProcessingSettings.EnableMultipleZones)
                {
                    var best = SelectBestIntersection(relevant);
                    relevant = best != null ? new List<ZoneTargetIntersection> { best } : new List<ZoneTargetIntersection>();
                }

                var isMultiZone = relevant.Count > 1;
                if (isMultiZone) stats.MultiZoneTagged++;

                if (request.ProcessingSettings.TagPartialSeparately)
                {
                    var behaviorCategory = request.ProcessingSettings.ZoneBehaviorCategory;
                    var behaviorProperty = request.ProcessingSettings.ZoneBehaviorPropertyName;
                    var containedValue = request.ProcessingSettings.ZoneBehaviorContainedValue;
                    var partialValue = request.ProcessingSettings.ZoneBehaviorPartialValue;

                    if (!string.IsNullOrWhiteSpace(behaviorCategory)
                        && !string.IsNullOrWhiteSpace(behaviorProperty))
                    {
                        var hasPartial = relevant.Any(r => r.IsPartial);
                        var behaviorValue = hasPartial ? partialValue : containedValue;
                        if (!string.IsNullOrWhiteSpace(behaviorValue))
                        {
                            PropertyWriter.WriteProperty(tgt.ModelItem, behaviorCategory, behaviorProperty, behaviorValue, WriteMode.Overwrite, string.Empty);
                            stats.WritesPerformed++;
                        }
                    }
                }

                foreach (var mapping in request.Mappings)
                {
                    var values = new List<string>();
                    foreach (var inter in relevant)
                    {
                        if (!zoneLookup.TryGetValue(inter.ZoneId, out var zone)) continue;
                        var val = zoneValueLookup.GetValue(inter.ZoneId, mapping.ZoneCategory, mapping.ZonePropertyName);
                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            values.Add(val);
                        }
                    }

                    var combined = CombineValues(values, mapping.MultiZoneCombineMode, mapping.AppendSeparator);
                    if (string.IsNullOrWhiteSpace(combined)) continue;
                    PropertyWriter.WriteProperty(tgt.ModelItem, mapping.TargetCategory, mapping.TargetPropertyName, combined, mapping.WriteMode, mapping.AppendSeparator);
                    stats.WritesPerformed++;
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
                        if (inter.IsContained) summary.ContainedCount++;
                        if (inter.IsPartial) summary.PartialCount++;
                    }
                }
            }
            writeSw.Stop();
            stats.WriteBackTime = writeSw.Elapsed;

            result.Stats = stats;
            sw.Stop();
            result.Stats.Elapsed = sw.Elapsed;
            result.Message = $"Space Mapper: {stats.ZonesProcessed} zones, {stats.TargetsProcessed} targets, {stats.ContainedTagged} contained, {stats.PartialTagged} partial. Mode={stats.ModeUsed}.";
            _log?.Invoke(result.Message);
            _log?.Invoke($"SpaceMapper Timings: resolve {stats.ResolveTime.TotalMilliseconds:0}ms, build {stats.BuildGeometryTime.TotalMilliseconds:0}ms, index {stats.BuildIndexTime.TotalMilliseconds:0}ms, candidates {stats.CandidateQueryTime.TotalMilliseconds:0}ms, narrow {stats.NarrowPhaseTime.TotalMilliseconds:0}ms, write {stats.WriteBackTime.TotalMilliseconds:0}ms");
            return result;
        }

        private static SpaceMapperPreflightCache TryGetReusablePreflightCache(SpaceMapperRequest request, SpaceMapperPreflightCache cache, IReadOnlyList<TargetGeometry> targets)
        {
            if (request == null || cache == null || cache.Grid == null || cache.TargetBounds == null || cache.TargetKeys == null)
            {
                return null;
            }

            var signature = SpaceMapperPreflightService.BuildSignature(request);
            if (!string.Equals(signature, cache.Signature, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            if (cache.TargetKeys.Length != targets.Count)
            {
                return null;
            }

            var keySet = new HashSet<string>(cache.TargetKeys, StringComparer.OrdinalIgnoreCase);
            foreach (var target in targets)
            {
                if (!keySet.Contains(target.ItemKey))
                {
                    return null;
                }
            }

            return cache;
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

        private static ZoneTargetIntersection SelectBestIntersection(IReadOnlyList<ZoneTargetIntersection> intersections)
        {
            if (intersections == null || intersections.Count == 0)
            {
                return null;
            }

            var bestContained = intersections
                .Where(i => i.IsContained)
                .OrderByDescending(i => i.OverlapVolume)
                .FirstOrDefault();

            if (bestContained != null)
            {
                return bestContained;
            }

            return intersections
                .OrderByDescending(i => i.OverlapVolume)
                .First();
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

        internal static ScrapeSession GetSession(string profile)
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

        internal static IEnumerable<SpaceMapperResolvedItem> ResolveZones(ScrapeSession session, ZoneSourceType source, string setName, Document doc)
        {
            switch (source)
            {
                case ZoneSourceType.ZoneSelectionSet:
                case ZoneSourceType.ZoneSearchSet:
                    // Placeholder: fall back to selection for now
                    foreach (ModelItem item in doc.CurrentSelection?.SelectedItems ?? new ModelItemCollection())
                        yield return new SpaceMapperResolvedItem { ItemKey = GetItemKey(item), ModelItem = item, DisplayName = item.DisplayName };
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
                        yield return new SpaceMapperResolvedItem
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

        internal static IEnumerable<TargetGeometry> ResolveTargets(Document doc,
            IEnumerable<SpaceMapperTargetRule> rules,
            Dictionary<string, List<SpaceMapperTargetRule>> targetsByRule)
        {
            var added = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var rule in rules.Where(r => r.Enabled))
            {
                var items = ResolveTargetsForRule(doc, rule);
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

        private static IEnumerable<ModelItem> ResolveTargetsForRule(Document doc, SpaceMapperTargetRule rule)
        {
            IEnumerable<ModelItem> items = Enumerable.Empty<ModelItem>();

            switch (rule.TargetDefinition)
            {
                case SpaceMapperTargetDefinition.CurrentSelection:
                    items = doc.CurrentSelection?.SelectedItems?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
                    break;
                case SpaceMapperTargetDefinition.SelectionTreeLevel:
                    items = TraverseWithDepth(doc.Models.RootItems, rule.MinLevel ?? 0, rule.MaxLevel ?? int.MaxValue);
                    break;
                case SpaceMapperTargetDefinition.SelectionSet:
                    items = ResolveSelectionSet(rule.SetSearchName);
                    break;
                case SpaceMapperTargetDefinition.SearchSet:
                    items = ResolveSearchSet(rule.SetSearchName);
                    break;
                case SpaceMapperTargetDefinition.EntireModel:
                default:
                    items = TraverseAll(doc.Models.RootItems);
                    break;
            }

            if (!string.IsNullOrWhiteSpace(rule.CategoryFilter))
            {
                items = items.Where(mi => CategoryMatches(mi, rule.CategoryFilter));
            }

            return items;
        }

        private static IEnumerable<ModelItem> ResolveSelectionSet(string setName)
        {
            return ResolveSelectionSetInternal(setName);
        }

        private static IEnumerable<ModelItem> ResolveSearchSet(string setName)
        {
            return ResolveSelectionSetInternal(setName);
        }

        private static IEnumerable<ModelItem> ResolveSelectionSetInternal(string setName)
        {
            if (string.IsNullOrWhiteSpace(setName))
            {
                return Enumerable.Empty<ModelItem>();
            }

            try
            {
                var doc = Application.ActiveDocument;
                var root = doc?.SelectionSets?.RootItem;
                var selectionSet = FindSelectionSetByName(root, setName);
                if (selectionSet != null)
                {
                    var items = selectionSet.GetSelectedItems(doc);
                    return items?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
                }
            }
            catch
            {
                // ignore managed API selection set failures
            }

            try
            {
                var state = ComBridge.State;
                var sets = state.SelectionSetsEx();
                var selectionSet = FindSelectionSetByName(sets, setName);
                if (selectionSet == null)
                {
                    return Enumerable.Empty<ModelItem>();
                }

                var items = ComBridge.ToModelItemCollection(selectionSet.selection);
                return items?.Cast<ModelItem>() ?? Enumerable.Empty<ModelItem>();
            }
            catch
            {
                return Enumerable.Empty<ModelItem>();
            }
        }

        private static SelectionSet FindSelectionSetByName(FolderItem folder, string name)
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
                else if (item is FolderItem subFolder)
                {
                    var found = FindSelectionSetByName(subFolder, name);
                    if (found != null)
                    {
                        return found;
                    }
                }
            }

            return null;
        }

        private static ComApi.InwOpSelectionSet FindSelectionSetByName(ComApi.InwSelectionSetExColl collection, string name)
        {
            if (collection == null || string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            for (int i = 1; i <= collection.Count; i++)
            {
                var item = collection[i];
                if (item is ComApi.InwOpSelectionSet set)
                {
                    if (string.Equals(set.name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return set;
                    }
                }
                else if (item is ComApi.InwSelectionSetFolder folder)
                {
                    var found = FindSelectionSetByName(folder.SelectionSets(), name);
                    if (found != null)
                    {
                        return found;
                    }
                }
            }

            return null;
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

        // Resolved item model moved to SpaceMapperModels.cs
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
            if (item == null) return;

            categoryName = string.IsNullOrWhiteSpace(categoryName) ? null : categoryName.Trim();
            propertyName = string.IsNullOrWhiteSpace(propertyName) ? null : propertyName.Trim();
            if (string.IsNullOrWhiteSpace(categoryName) || string.IsNullOrWhiteSpace(propertyName)) return;

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

                ComApi.InwOaPropertyVec propertyVector = null;
                try
                {
                    var getMethod = propertyNode.GetType().GetMethod("GetUserDefined");
                    propertyVector = getMethod?.Invoke(propertyNode, new object[] { 0, categoryName }) as ComApi.InwOaPropertyVec;
                }
                catch
                {
                    // ignore read failures
                }

                if (propertyVector == null)
                {
                    propertyVector = (ComApi.InwOaPropertyVec)state.ObjectFactory(
                        ComApi.nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
                }

                var existingProp = FindProperty(propertyVector, propertyName);
                if (existingProp == null)
                {
                    existingProp = (ComApi.InwOaProperty)state.ObjectFactory(
                        ComApi.nwEObjectType.eObjectType_nwOaProperty, null, null);
                    existingProp.name = propertyName;
                    existingProp.UserName = propertyName;
                    propertyVector.Properties().Add(existingProp);
                }

                existingProp.value = finalValue;
                propertyNode.SetUserDefined(0, categoryName, categoryName, propertyVector);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[SpaceMapper] Failed to write {categoryName}.{propertyName} on {item.DisplayName}: {ex.Message}");
            }
        }

        private static ComApi.InwOaProperty FindProperty(ComApi.InwOaPropertyVec vec, string propertyName)
        {
            foreach (ComApi.InwOaProperty prop in vec.Properties())
            {
                if (prop == null) continue;
                if (string.Equals(prop.name, propertyName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(prop.UserName, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    return prop;
                }
            }

            return null;
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
