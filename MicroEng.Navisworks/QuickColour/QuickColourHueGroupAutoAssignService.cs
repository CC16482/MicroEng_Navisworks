using System;
using System.Collections.Generic;
using System.Linq;

namespace MicroEng.Navisworks.QuickColour
{
    public sealed class HueGroupAutoAssignOptions
    {
        public bool OnlyAssignWhenCurrentlyFallback { get; set; } = true;
        public bool SkipLockedCategories { get; set; }
        public bool AutoCreateMissingHueGroups { get; set; }
    }

    public sealed class HueGroupAutoAssignResult
    {
        public int TotalCategories;
        public int Assigned;
        public int Unmatched;
        public int SkippedLocked;
        public List<string> MissingHueGroups = new List<string>();
        public Dictionary<string, int> AssignedByGroup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class QuickColourHueGroupAutoAssignService
    {
        private sealed class ResolveDetails
        {
            public string Group;
            public string MatcherType;
            public string MatcherValue;
        }

        public HueGroupAutoAssignResult Apply(
            DisciplineMapFile map,
            IList<QuickColourHierarchyGroup> categories,
            IList<QuickColourHueGroup> hueGroups,
            HueGroupAutoAssignOptions options,
            Action<string> log = null)
        {
            if (map == null) throw new ArgumentNullException(nameof(map));
            if (categories == null) throw new ArgumentNullException(nameof(categories));
            if (hueGroups == null) throw new ArgumentNullException(nameof(hueGroups));
            options = options ?? new HueGroupAutoAssignOptions();

            var enabledHueGroupNames = new HashSet<string>(
                hueGroups.Where(h => h != null && h.Enabled && !string.IsNullOrWhiteSpace(h.Name))
                         .Select(h => h.Name.Trim()),
                StringComparer.OrdinalIgnoreCase);

            var fallback = string.IsNullOrWhiteSpace(map.FallbackGroup) ? "Other" : map.FallbackGroup.Trim();
            if (!enabledHueGroupNames.Contains(fallback))
            {
                fallback = enabledHueGroupNames.Contains("Other")
                    ? "Other"
                    : enabledHueGroupNames.FirstOrDefault() ?? "Other";
            }

            var result = new HueGroupAutoAssignResult();
            result.TotalCategories = categories.Count;

            var missing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var cat in categories.Where(c => c != null))
            {
                if (!cat.Enabled)
                {
                    continue;
                }

                if (options.SkipLockedCategories && cat.UseCustomBaseColor)
                {
                    result.SkippedLocked++;
                    continue;
                }

                if (options.OnlyAssignWhenCurrentlyFallback)
                {
                    var current = (cat.HueGroupName ?? "").Trim();
                    if (!string.Equals(current, fallback, StringComparison.OrdinalIgnoreCase)
                        && !string.IsNullOrWhiteSpace(current))
                    {
                        continue;
                    }
                }

                var details = ResolveGroupWithDetails(map, cat.Value, fallback);
                var assignedGroup = (details?.Group ?? fallback).Trim();
                if (string.IsNullOrWhiteSpace(assignedGroup))
                {
                    assignedGroup = fallback;
                }

                if (!enabledHueGroupNames.Contains(assignedGroup))
                {
                    if (options.AutoCreateMissingHueGroups)
                    {
                        hueGroups.Add(new QuickColourHueGroup { Name = assignedGroup, HueHex = "#6B7280" });
                        enabledHueGroupNames.Add(assignedGroup);
                    }
                    else
                    {
                        missing.Add(assignedGroup);
                        assignedGroup = fallback;
                    }
                }

                cat.HueGroupName = assignedGroup;

                result.Assigned++;
                if (!result.AssignedByGroup.ContainsKey(assignedGroup))
                {
                    result.AssignedByGroup[assignedGroup] = 0;
                }
                result.AssignedByGroup[assignedGroup]++;
            }

            result.MissingHueGroups = missing.OrderBy(s => s).ToList();
            result.Unmatched = categories.Count(c => c != null && c.Enabled
                && string.Equals((c.HueGroupName ?? "").Trim(), fallback, StringComparison.OrdinalIgnoreCase));

            if (result.MissingHueGroups.Count > 0)
            {
                log?.Invoke("HueGroup AutoAssign: mapping referenced unknown groups: " + string.Join(", ", result.MissingHueGroups));
            }

            log?.Invoke($"HueGroup AutoAssign: Assigned={result.Assigned}, Fallback='{fallback}', SkippedLocked={result.SkippedLocked}.");

            return result;
        }

        public HueGroupAutoAssignPreviewResult BuildPreview(
            DisciplineMapFile map,
            IList<QuickColourHierarchyGroup> categories,
            IList<QuickColourHueGroup> hueGroups,
            HueGroupAutoAssignOptions options,
            IList<HueGroupAutoAssignPreviewRow> outputRows,
            Action<string> log = null)
        {
            if (map == null) throw new ArgumentNullException(nameof(map));
            if (categories == null) throw new ArgumentNullException(nameof(categories));
            if (hueGroups == null) throw new ArgumentNullException(nameof(hueGroups));
            if (outputRows == null) throw new ArgumentNullException(nameof(outputRows));
            options = options ?? new HueGroupAutoAssignOptions();

            outputRows.Clear();

            var enabledHueGroupNames = new HashSet<string>(
                hueGroups.Where(h => h != null && h.Enabled && !string.IsNullOrWhiteSpace(h.Name))
                         .Select(h => h.Name.Trim()),
                StringComparer.OrdinalIgnoreCase);

            var fallback = string.IsNullOrWhiteSpace(map.FallbackGroup) ? "Other" : map.FallbackGroup.Trim();
            if (!enabledHueGroupNames.Contains(fallback))
            {
                fallback = enabledHueGroupNames.Contains("Other")
                    ? "Other"
                    : enabledHueGroupNames.FirstOrDefault() ?? "Other";
            }

            var summary = new HueGroupAutoAssignPreviewResult();
            summary.Total = categories.Count;

            foreach (var cat in categories.Where(c => c != null))
            {
                if (!cat.Enabled)
                {
                    continue;
                }

                var current = (cat.HueGroupName ?? "").Trim();
                if (string.IsNullOrWhiteSpace(current))
                {
                    current = fallback;
                }

                var row = new HueGroupAutoAssignPreviewRow
                {
                    Category = cat.Value ?? "",
                    Count = cat.Count,
                    TypeCount = cat.Types?.Count ?? 0,
                    Locked = cat.UseCustomBaseColor,
                    CurrentGroup = current
                };

                if (options.SkipLockedCategories && cat.UseCustomBaseColor)
                {
                    row.ProposedGroup = current;
                    row.Status = "Skipped (Locked)";
                    row.Reason = "Category is locked.";
                    row.WillChange = false;
                    summary.SkippedLocked++;
                    outputRows.Add(row);
                    continue;
                }

                if (options.OnlyAssignWhenCurrentlyFallback)
                {
                    if (!string.Equals(current, fallback, StringComparison.OrdinalIgnoreCase))
                    {
                        row.ProposedGroup = current;
                        row.Status = "Skipped (Already assigned)";
                        row.Reason = $"Only assigns categories currently set to '{fallback}'.";
                        row.WillChange = false;
                        summary.SkippedAlreadyAssigned++;
                        outputRows.Add(row);
                        continue;
                    }
                }

                var details = ResolveGroupWithDetails(map, cat.Value, fallback);
                var resolved = (details?.Group ?? fallback).Trim();
                if (string.IsNullOrWhiteSpace(resolved))
                {
                    resolved = fallback;
                }

                row.MatchedRuleGroup = resolved;
                row.MatchedMatcherType = string.IsNullOrWhiteSpace(details?.MatcherType) ? "" : details.MatcherType;
                row.MatchedMatcherValue = string.IsNullOrWhiteSpace(details?.MatcherValue) ? "" : details.MatcherValue;

                if (!enabledHueGroupNames.Contains(resolved))
                {
                    if (options.AutoCreateMissingHueGroups)
                    {
                        row.ProposedGroup = resolved;
                        row.Status = "Will create group";
                        row.Reason = "Group not found in UI; will be created (Auto-create enabled).";
                        row.WillChange = !string.Equals(current, resolved, StringComparison.OrdinalIgnoreCase);
                        if (row.WillChange) summary.WillChange++; else summary.NoChange++;
                        summary.WillCreateMissingGroup++;
                        outputRows.Add(row);
                        continue;
                    }
                    else
                    {
                        row.ProposedGroup = fallback;
                        row.Status = "Missing group + fallback";
                        row.Reason = $"Mapping resolved to '{resolved}' which does not exist; will use fallback '{fallback}'.";
                        row.WillChange = !string.Equals(current, fallback, StringComparison.OrdinalIgnoreCase);
                        if (row.WillChange) summary.WillChange++; else summary.NoChange++;
                        summary.MissingGroupFallback++;
                        outputRows.Add(row);
                        continue;
                    }
                }

                row.ProposedGroup = resolved;

                if (string.Equals(resolved, fallback, StringComparison.OrdinalIgnoreCase)
                    && string.IsNullOrWhiteSpace(row.MatchedMatcherType))
                {
                    row.Status = "Unmatched + fallback";
                    row.Reason = $"No rule matched; fallback '{fallback}'.";
                    summary.UnmatchedFallback++;
                }
                else
                {
                    row.Status = string.Equals(current, resolved, StringComparison.OrdinalIgnoreCase)
                        ? "No change"
                        : "Will change";
                    row.Reason = string.Equals(current, resolved, StringComparison.OrdinalIgnoreCase)
                        ? "Already assigned to this group."
                        : "Will be updated by mapping.";
                }

                row.WillChange = !string.Equals(current, resolved, StringComparison.OrdinalIgnoreCase);
                if (row.WillChange) summary.WillChange++; else summary.NoChange++;

                outputRows.Add(row);
            }

            log?.Invoke($"HueGroup Preview: total={summary.Total}, willChange={summary.WillChange}, skippedLocked={summary.SkippedLocked}, unmatched={summary.UnmatchedFallback}.");
            return summary;
        }

        private static ResolveDetails ResolveGroupWithDetails(DisciplineMapFile map, string categoryValue, string fallback)
        {
            categoryValue = (categoryValue ?? "").Trim();

            foreach (var rule in map.Rules ?? new List<DisciplineMapRule>())
            {
                var group = (rule?.Group ?? "").Trim();
                if (string.IsNullOrWhiteSpace(group))
                {
                    continue;
                }

                var matchers = rule.Match ?? new List<DisciplineMapMatcher>();
                foreach (var m in matchers)
                {
                    var type = (m?.Type ?? "exact").Trim();
                    var value = (m?.Value ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        continue;
                    }

                    if (DisciplineMapMatcherEngine.IsMatch(categoryValue, type, value))
                    {
                        return new ResolveDetails
                        {
                            Group = group,
                            MatcherType = type,
                            MatcherValue = value
                        };
                    }
                }
            }

            return new ResolveDetails
            {
                Group = fallback,
                MatcherType = "",
                MatcherValue = ""
            };
        }
    }
}
