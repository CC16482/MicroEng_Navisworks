
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using Autodesk.Navisworks.Api;
using MicroEng.Navisworks.SpaceMapper.Estimation;
using MicroEng.Navisworks.SpaceMapper.Geometry;

namespace MicroEng.Navisworks
{
    internal sealed class SpaceMapperComparisonOptions
    {
        public int MaxZones { get; set; } = 5000;
        public int MaxTargets { get; set; } = 100000;
        public int MismatchSampleCount { get; set; } = 10;
        public bool IncludeLegacyFastAabb { get; set; }
        public bool ReusePreflightForRun { get; set; } = true;
        public SpaceMapperBenchmarkMode BenchmarkMode { get; set; } = SpaceMapperBenchmarkMode.ComputeOnly;
        public int? SampleSeed { get; set; }
    }

    internal sealed class SpaceMapperComparisonInput
    {
        public SpaceMapperRequest Request { get; set; }
        public SpaceMapperComputeDataset Dataset { get; set; }
        public SpaceMapperComparisonOptions Options { get; set; } = new();
        public string ModelName { get; set; }
    }

    internal sealed class SpaceMapperComparisonOutput
    {
        public SpaceMapperComparisonReport Report { get; set; }
        public string Markdown { get; set; }
        public string Json { get; set; }
    }

    internal sealed class SpaceMapperComparisonProgress
    {
        public string Stage { get; set; }
        public int Percentage { get; set; } = -1;
    }

    internal sealed class SpaceMapperComparisonRunner
    {
        private readonly Action<string> _log;

        public SpaceMapperComparisonRunner(Action<string> log)
        {
            _log = log;
        }

        public SpaceMapperComparisonOutput Run(
            SpaceMapperComparisonInput input,
            IProgress<SpaceMapperComparisonProgress> progress,
            CancellationToken token)
        {
            if (input == null) throw new ArgumentNullException(nameof(input));
            if (input.Dataset == null) throw new ArgumentNullException(nameof(input.Dataset));

            var options = NormalizeOptions(input.Options);
            if (options.BenchmarkMode == SpaceMapperBenchmarkMode.FullWriteback)
            {
                options.MaxZones = int.MaxValue;
                options.MaxTargets = int.MaxValue;
            }
            var report = new SpaceMapperComparisonReport
            {
                SchemaVersion = "1.0",
                GeneratedUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
                Environment = BuildEnvironmentInfo(input),
                Warnings = new List<string>()
            };

            var baseZones = input.Dataset.Zones?.Where(z => z?.BoundingBox != null).ToList() ?? new List<ZoneGeometry>();
            var baseTargets = input.Dataset.TargetsForEngine?.Where(t => t?.BoundingBox != null).ToList() ?? new List<TargetGeometry>();
            var originalZoneCount = baseZones.Count;
            var originalTargetCount = baseTargets.Count;

            var targetsByRule = FilterTargetsByRule(input.Dataset.TargetsByRule, baseTargets);

            var seed = options.SampleSeed ?? unchecked(Environment.TickCount * 397);
            var rng = new Random(seed);

            var zones = SampleList(baseZones, options.MaxZones, rng, out var zonesSampled, token);
            var targets = SampleList(baseTargets, options.MaxTargets, rng, out var targetsSampled, token);

            if (zonesSampled || targetsSampled)
            {
                targetsByRule = FilterTargetsByRule(targetsByRule, targets);
            }

            _log?.Invoke($"SpaceMapper Comparison: zones {zones.Count}, targets {targets.Count}, sampled {(zonesSampled || targetsSampled ? "yes" : "no")}.");

            report.Dataset = new SpaceMapperComparisonDataset
            {
                ModelName = input.ModelName ?? string.Empty,
                ZoneCount = originalZoneCount,
                TargetCount = originalTargetCount,
                RawZoneCount = input.Dataset.ZoneModels?.Count ?? 0,
                RawTargetCount = input.Dataset.TargetModels?.Count ?? 0,
                Sampled = zonesSampled || targetsSampled,
                SampleSeed = seed,
                SampledZones = zones.Count,
                SampledTargets = targets.Count,
                ResolveMs = input.Dataset.ResolveTime.TotalMilliseconds,
                BuildGeometryMs = input.Dataset.BuildGeometryTime.TotalMilliseconds,
                ZoneSource = input.Request?.ZoneSource.ToString() ?? string.Empty,
                TargetRuleCount = input.Request?.TargetRules?.Count ?? 0
            };

            if (report.Dataset.RawZoneCount > 0 && report.Dataset.RawZoneCount != report.Dataset.ZoneCount)
            {
                report.Warnings.Add($"Zones with geometry: {report.Dataset.ZoneCount:N0}/{report.Dataset.RawZoneCount:N0}.");
            }

            if (report.Dataset.RawTargetCount > 0 && report.Dataset.RawTargetCount != report.Dataset.TargetCount)
            {
                report.Warnings.Add($"Targets with geometry: {report.Dataset.TargetCount:N0}/{report.Dataset.RawTargetCount:N0}.");
            }

            if (zones.Count == 0 || targets.Count == 0)
            {
                report.Warnings.Add("No zones or targets available for comparison.");
                return BuildOutput(report);
            }

            var comparisonDataset = new SpaceMapperComputeDataset
            {
                Zones = zones,
                TargetsForEngine = targets,
                TargetModels = targets,
                TargetsByRule = targetsByRule,
                ResolveTime = input.Dataset.ResolveTime,
                BuildGeometryTime = input.Dataset.BuildGeometryTime
            };

            var targetBoundsData = BuildTargetBounds(targets);
            var zoneBoundsData = BuildZoneBounds(zones, input.Request?.ProcessingSettings);
            var zoneNameById = BuildZoneNameLookup(zones);

            var variantDefinitions = BuildVariantDefinitions(options.IncludeLegacyFastAabb);
            var variantResults = new List<VariantRunResult>(variantDefinitions.Count);

            var zonesWithoutPlanes = zones.Count(z => z?.Planes == null || z.Planes.Count == 0);
            var baseSettings = input.Request?.ProcessingSettings;

            for (int i = 0; i < variantDefinitions.Count; i++)
            {
                token.ThrowIfCancellationRequested();

                var def = variantDefinitions[i];
                var settings = CloneSettings(baseSettings);
                settings.PerformancePreset = def.Preset;
                settings.FastTraversalMode = def.Traversal;
                settings.UseOriginPointOnly = def.UseOriginPointOnly;
                var forcedPartialsOff = false;

                var stagePrefix = $"[{i + 1}/{variantDefinitions.Count}] {def.Name}";
                ReportProgress(progress, $"{stagePrefix}: preflight", ComputePercent(i, variantDefinitions.Count, 0.05));

                if (def.RequiresNoPartial && (settings.TagPartialSeparately || settings.TreatPartialAsContained))
                {
                    forcedPartialsOff = true;
                    settings.TagPartialSeparately = false;
                    settings.TreatPartialAsContained = false;
                }

                if (def.LegacyFastAabbOnly)
                {
                    variantResults.Add(VariantRunResult.Skipped(def, settings, "Legacy AABB-only fast comparison not implemented.", options.BenchmarkMode));
                    continue;
                }

                var preflightSw = Stopwatch.StartNew();
                var preflight = SpaceMapperPreflightService.RunForDataset(
                    zoneBoundsData.Bounds,
                    targetBoundsData.Bounds,
                    targetBoundsData.Keys,
                    settings,
                    token,
                    out var preflightCache);
                preflightSw.Stop();

                ReportProgress(progress, $"{stagePrefix}: compute", ComputePercent(i, variantDefinitions.Count, 0.45));

                var engine = SpaceMapperEngineFactory.Create(settings.ProcessingMode);
                var diagnostics = new SpaceMapperEngineDiagnostics();
                var computeSw = Stopwatch.StartNew();
                var intersections = engine.ComputeIntersections(
                    zones,
                    targets,
                    settings,
                    options.ReusePreflightForRun ? preflightCache : null,
                    diagnostics,
                    null,
                    token) ?? new List<ZoneTargetIntersection>();
                computeSw.Stop();

                ReportProgress(progress, $"{stagePrefix}: postprocess", ComputePercent(i, variantDefinitions.Count, 0.75));

                var postSw = Stopwatch.StartNew();
                var assignmentSummary = BuildAssignments(targets, targetsByRule, intersections, settings);
                postSw.Stop();

                SpaceMapperWritebackResult writeback = null;
                if (options.BenchmarkMode != SpaceMapperBenchmarkMode.ComputeOnly)
                {
                    var session = SpaceMapperService.GetSession(input.Request?.ScraperProfileName);
                    var writeToModel = options.BenchmarkMode == SpaceMapperBenchmarkMode.FullWriteback;
                    writeback = SpaceMapperService.ExecuteWriteback(
                        comparisonDataset,
                        settings,
                        input.Request?.Mappings ?? new List<SpaceMapperMappingDefinition>(),
                        session,
                        intersections,
                        writeToModel,
                        token);
                }

                var variantReport = BuildVariantReport(
                    def,
                    settings,
                    engine,
                    diagnostics,
                    preflight,
                    preflightSw.Elapsed,
                    computeSw.Elapsed,
                    postSw.Elapsed,
                    zones.Count,
                    targets.Count,
                    assignmentSummary,
                    writeback,
                    options.BenchmarkMode,
                    input.Dataset.ResolveTime.TotalMilliseconds,
                    input.Dataset.BuildGeometryTime.TotalMilliseconds,
                    zonesWithoutPlanes);

                if (forcedPartialsOff)
                {
                    report.Warnings.Add($"{def.Name}: partial options forced OFF to allow target-major benchmark.");
                }

                variantResults.Add(new VariantRunResult
                {
                    Definition = def,
                    Report = variantReport,
                    Assignments = assignmentSummary.Assignments,
                    ZoneBoundsById = zoneBoundsData.ById,
                    Settings = settings,
                    TargetsProcessed = targets.Count,
                    ZonesProcessed = zones.Count,
                    MaxIntersectionsPerTarget = assignmentSummary.MaxIntersectionsPerTarget,
                    NeedsPartial = settings.TagPartialSeparately || settings.TreatPartialAsContained
                });

                ReportProgress(progress, $"{stagePrefix}: done", ComputePercent(i + 1, variantDefinitions.Count, 0.0));
            }

            report.Variants = variantResults.Select(r => r.Report).ToList();

            var comparisons = new List<SpaceMapperComparisonAgreement>();
            var mismatchSamples = new List<SpaceMapperComparisonMismatch>();
            var baseline = variantResults.FirstOrDefault(r => r.Definition.IsBaseline && !r.Report.Skipped);

            if (baseline == null)
            {
                report.Warnings.Add("Baseline variant (Normal / Zone-major) did not run.");
            }
            else
            {
                foreach (var variant in variantResults)
                {
                    if (variant == baseline || variant.Report.Skipped)
                    {
                        continue;
                    }

                    var comparison = CompareAssignments(
                        baseline,
                        variant,
                        targets,
                        targetBoundsData.ByKey,
                        zoneNameById,
                        options.MismatchSampleCount,
                        mismatchSamples);

                    comparisons.Add(comparison);
                }

                report.Comparisons = comparisons;
                report.MismatchSamples = mismatchSamples;
            }

            report.Warnings.AddRange(BuildWarnings(variantResults, baseline, zonesWithoutPlanes));

            return BuildOutput(report);
        }

        private static SpaceMapperComparisonOptions NormalizeOptions(SpaceMapperComparisonOptions options)
        {
            options ??= new SpaceMapperComparisonOptions();
            if (options.MaxZones <= 0) options.MaxZones = 5000;
            if (options.MaxTargets <= 0) options.MaxTargets = 100000;
            if (options.MismatchSampleCount <= 0) options.MismatchSampleCount = 10;
            return options;
        }

        private static SpaceMapperComparisonEnvironment BuildEnvironmentInfo(SpaceMapperComparisonInput input)
        {
            var process = Process.GetCurrentProcess();
            return new SpaceMapperComparisonEnvironment
            {
                NavisworksApiVersion = typeof(Application).Assembly.GetName().Version?.ToString() ?? string.Empty,
                PluginVersion = typeof(SpaceMapperControl).Assembly.GetName().Version?.ToString() ?? string.Empty,
                OSVersion = Environment.OSVersion.VersionString,
                CpuCores = Environment.ProcessorCount,
                ProcessWorkingSetMb = Math.Round(process.WorkingSet64 / 1024d / 1024d, 1),
                MachineName = Environment.MachineName,
                ModelName = input.ModelName ?? string.Empty
            };
        }

        private static SpaceMapperProcessingSettings CloneSettings(SpaceMapperProcessingSettings settings)
        {
            if (settings == null)
            {
                return new SpaceMapperProcessingSettings();
            }

            return new SpaceMapperProcessingSettings
            {
                ProcessingMode = settings.ProcessingMode,
                TreatPartialAsContained = settings.TreatPartialAsContained,
                TagPartialSeparately = settings.TagPartialSeparately,
                EnableMultipleZones = settings.EnableMultipleZones,
                Offset3D = settings.Offset3D,
                OffsetTop = settings.OffsetTop,
                OffsetBottom = settings.OffsetBottom,
                OffsetSides = settings.OffsetSides,
                Units = settings.Units,
                OffsetMode = settings.OffsetMode,
                MaxThreads = settings.MaxThreads,
                BatchSize = settings.BatchSize,
                IndexGranularity = settings.IndexGranularity,
                PerformancePreset = settings.PerformancePreset,
                ZoneBehaviorCategory = settings.ZoneBehaviorCategory,
                ZoneBehaviorPropertyName = settings.ZoneBehaviorPropertyName,
                ZoneBehaviorContainedValue = settings.ZoneBehaviorContainedValue,
                ZoneBehaviorPartialValue = settings.ZoneBehaviorPartialValue,
                UseOriginPointOnly = settings.UseOriginPointOnly,
                FastTraversalMode = settings.FastTraversalMode,
                WritebackStrategy = settings.WritebackStrategy,
                ShowInternalPropertiesDuringWriteback = settings.ShowInternalPropertiesDuringWriteback,
                CloseDockPanesDuringRun = settings.CloseDockPanesDuringRun,
                SkipUnchangedWriteback = settings.SkipUnchangedWriteback,
                PackWritebackProperties = settings.PackWritebackProperties
            };
        }

        private static List<VariantDefinition> BuildVariantDefinitions(bool includeLegacyFastAabb)
        {
            var list = new List<VariantDefinition>
            {
                new VariantDefinition("Normal / Zone-major", SpaceMapperPerformancePreset.Normal, SpaceMapperFastTraversalMode.ZoneMajor, useOriginPointOnly: false, isBaseline: true),
                new VariantDefinition("Fast / Zone-major", SpaceMapperPerformancePreset.Fast, SpaceMapperFastTraversalMode.ZoneMajor, useOriginPointOnly: true),
                new VariantDefinition("Accurate / Zone-major", SpaceMapperPerformancePreset.Accurate, SpaceMapperFastTraversalMode.ZoneMajor, useOriginPointOnly: false),
                new VariantDefinition("Fast / Target-major", SpaceMapperPerformancePreset.Fast, SpaceMapperFastTraversalMode.TargetMajor, useOriginPointOnly: true, requiresNoPartial: true)
            };

            if (includeLegacyFastAabb)
            {
                list.Add(new VariantDefinition("Fast (Legacy AABB) / Zone-major", SpaceMapperPerformancePreset.Fast, SpaceMapperFastTraversalMode.ZoneMajor, useOriginPointOnly: false, legacyFastAabbOnly: true));
            }

            return list;
        }

        private static SpaceMapperComparisonVariant BuildVariantReport(
            VariantDefinition def,
            SpaceMapperProcessingSettings settings,
            ISpatialIntersectionEngine engine,
            SpaceMapperEngineDiagnostics diagnostics,
            SpaceMapperPreflightResult preflight,
            TimeSpan preflightElapsed,
            TimeSpan computeElapsed,
            TimeSpan postElapsed,
            int zoneCount,
            int targetCount,
            VariantAssignmentSummary assignmentSummary,
            SpaceMapperWritebackResult writeback,
            SpaceMapperBenchmarkMode benchmarkMode,
            double resolveMs,
            double buildGeometryMs,
            int zonesWithoutPlanes)
        {
            var preflightMs = preflightElapsed.TotalMilliseconds;
            var postprocessMs = postElapsed.TotalMilliseconds;
            if (benchmarkMode != SpaceMapperBenchmarkMode.ComputeOnly && writeback != null)
            {
                postprocessMs += writeback.Elapsed.TotalMilliseconds;
            }

            var totalMs = preflightMs + computeElapsed.TotalMilliseconds + postprocessMs;

            return new SpaceMapperComparisonVariant
            {
                Name = def.Name,
                Skipped = false,
                SkipReason = string.Empty,
                Settings = new SpaceMapperComparisonSettings
                {
                    Preset = def.Preset.ToString(),
                    PresetResolved = diagnostics.PresetUsed.ToString(),
                    TraversalRequested = def.Traversal.ToString(),
                    TraversalResolved = diagnostics.TraversalUsed ?? string.Empty,
                    ProcessingMode = engine.Mode.ToString(),
                    IndexGranularity = settings.IndexGranularity,
                    MaxThreads = settings.MaxThreads ?? 0,
                    BatchSize = settings.BatchSize ?? 0,
                    UseOriginPointOnly = settings.UseOriginPointOnly,
                    TreatPartialAsContained = settings.TreatPartialAsContained,
                    TagPartialSeparately = settings.TagPartialSeparately,
                    EnableMultipleZones = settings.EnableMultipleZones,
                    Offset3D = settings.Offset3D,
                    OffsetTop = settings.OffsetTop,
                    OffsetBottom = settings.OffsetBottom,
                    OffsetSides = settings.OffsetSides,
                    OffsetMode = settings.OffsetMode,
                    Units = settings.Units,
                    BestZoneBehavior = settings.EnableMultipleZones ? "Multiple" : "Single",
                    ZonesWithoutPlanes = zonesWithoutPlanes,
                    BenchmarkMode = benchmarkMode.ToString(),
                    WritebackStrategy = settings.WritebackStrategy.ToString(),
                    ShowInternalPropertiesDuringWriteback = settings.ShowInternalPropertiesDuringWriteback,
                    SkipUnchangedWriteback = settings.SkipUnchangedWriteback,
                    PackWritebackProperties = settings.PackWritebackProperties
                },
                Timings = new SpaceMapperComparisonTimings
                {
                    ResolveMs = resolveMs,
                    BuildGeometryMs = buildGeometryMs,
                    PreflightMs = preflightMs,
                    IndexBuildMs = diagnostics.BuildIndexTime.TotalMilliseconds,
                    CandidateQueryMs = diagnostics.CandidateQueryTime.TotalMilliseconds,
                    NarrowPhaseMs = diagnostics.NarrowPhaseTime.TotalMilliseconds,
                    PostprocessMs = postprocessMs,
                    TotalMs = totalMs,
                    PreflightBuildMs = preflight?.BuildIndexTime.TotalMilliseconds ?? 0,
                    PreflightQueryMs = preflight?.QueryTime.TotalMilliseconds ?? 0
                },
                Workload = new SpaceMapperComparisonWorkload
                {
                    ZonesProcessed = zoneCount,
                    TargetsProcessed = targetCount,
                    CandidatePairs = diagnostics.CandidatePairs,
                    AvgCandidatesPerZone = diagnostics.AvgCandidatesPerZone,
                    MaxCandidatesPerZone = diagnostics.MaxCandidatesPerZone,
                    AvgCandidatesPerTarget = diagnostics.AvgCandidatesPerTarget,
                    MaxCandidatesPerTarget = diagnostics.MaxCandidatesPerTarget,
                    ContainedHits = assignmentSummary.ContainedHits,
                    PartialHits = assignmentSummary.PartialHits,
                    UnmatchedTargets = assignmentSummary.UnmatchedTargets,
                    MultiZoneTargets = assignmentSummary.MultiZoneTargets,
                    MaxIntersectionsPerTarget = assignmentSummary.MaxIntersectionsPerTarget,
                    WritesPerformed = writeback?.WritesPerformed ?? 0,
                    WritebackTargetsWritten = writeback?.TargetsWritten ?? 0,
                    WritebackCategoriesWritten = writeback?.CategoriesWritten ?? 0,
                    WritebackPropertiesWritten = writeback?.PropertiesWritten ?? 0,
                    AvgMsPerCategoryWrite = writeback?.AvgMsPerCategoryWrite ?? 0,
                    AvgMsPerTargetWrite = writeback?.AvgMsPerTargetWrite ?? 0,
                    SkippedUnchanged = writeback?.SkippedUnchanged ?? 0
                },
                Diagnostics = new SpaceMapperComparisonDiagnostics
                {
                    UsedPreflightIndex = diagnostics.UsedPreflightIndex,
                    PreflightCandidatePairs = preflight?.CandidatePairs ?? 0,
                    PreflightTraversalUsed = preflight?.TraversalUsed.ToString() ?? string.Empty
                }
            };
        }
        private static VariantAssignmentSummary BuildAssignments(
            IReadOnlyList<TargetGeometry> targets,
            Dictionary<string, List<SpaceMapperTargetRule>> targetsByRule,
            IEnumerable<ZoneTargetIntersection> intersections,
            SpaceMapperProcessingSettings settings)
        {
            var summary = new VariantAssignmentSummary();
            var byTarget = new Dictionary<string, List<ZoneTargetIntersection>>(StringComparer.OrdinalIgnoreCase);

            if (intersections != null)
            {
                foreach (var hit in intersections)
                {
                    if (hit == null || string.IsNullOrWhiteSpace(hit.TargetItemKey))
                    {
                        continue;
                    }

                    if (!byTarget.TryGetValue(hit.TargetItemKey, out var list))
                    {
                        list = new List<ZoneTargetIntersection>();
                        byTarget[hit.TargetItemKey] = list;
                    }
                    list.Add(hit);
                }
            }

            foreach (var target in targets)
            {
                if (target == null || string.IsNullOrWhiteSpace(target.ItemKey))
                {
                    continue;
                }

                if (!targetsByRule.TryGetValue(target.ItemKey, out var rules) || rules.Count == 0)
                {
                    continue;
                }

                byTarget.TryGetValue(target.ItemKey, out var list);
                var relevant = list != null
                    ? FilterByMembership(list, rules, settings).ToList()
                    : new List<ZoneTargetIntersection>();

                summary.MaxIntersectionsPerTarget = Math.Max(summary.MaxIntersectionsPerTarget, relevant.Count);

                if (relevant.Count == 0)
                {
                    summary.UnmatchedTargets++;
                    continue;
                }

                if (settings.EnableMultipleZones && relevant.Count > 1)
                {
                    summary.MultiZoneTargets++;
                }

                if (settings.EnableMultipleZones)
                {
                    foreach (var inter in relevant)
                    {
                        if (inter.IsContained) summary.ContainedHits++;
                        if (inter.IsPartial) summary.PartialHits++;
                    }
                }
                else
                {
                    var best = SelectBestIntersection(relevant);
                    if (best != null)
                    {
                        if (best.IsContained) summary.ContainedHits++;
                        if (best.IsPartial) summary.PartialHits++;
                        relevant = new List<ZoneTargetIntersection> { best };
                    }
                }

                var bestForCompare = SelectBestIntersection(relevant);
                if (bestForCompare != null)
                {
                    summary.Assignments[target.ItemKey] = new Assignment
                    {
                        ZoneId = bestForCompare.ZoneId,
                        IsContained = bestForCompare.IsContained,
                        IsPartial = bestForCompare.IsPartial
                    };
                }
            }

            return summary;
        }

        private static IEnumerable<ZoneTargetIntersection> FilterByMembership(
            IEnumerable<ZoneTargetIntersection> intersections,
            List<SpaceMapperTargetRule> rules,
            SpaceMapperProcessingSettings settings)
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

        private static SpaceMapperComparisonAgreement CompareAssignments(
            VariantRunResult baseline,
            VariantRunResult variant,
            IReadOnlyList<TargetGeometry> targets,
            Dictionary<string, Aabb> targetBoundsByKey,
            Dictionary<string, string> zoneNameById,
            int mismatchLimit,
            List<SpaceMapperComparisonMismatch> mismatchSamples)
        {
            var total = targets.Count;
            var same = 0;
            var baselineAssignedVariantUnassigned = 0;
            var baselineUnassignedVariantAssigned = 0;
            var different = 0;
            var perVariantSamples = 0;

            foreach (var target in targets)
            {
                var key = target?.ItemKey;
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                baseline.Assignments.TryGetValue(key, out var baseAssignment);
                variant.Assignments.TryGetValue(key, out var variantAssignment);

                var baseZone = baseAssignment?.ZoneId;
                var variantZone = variantAssignment?.ZoneId;

                if (string.IsNullOrWhiteSpace(baseZone) && string.IsNullOrWhiteSpace(variantZone))
                {
                    same++;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(baseZone))
                {
                    baselineUnassignedVariantAssigned++;
                    if (perVariantSamples < mismatchLimit)
                    {
                        mismatchSamples.Add(BuildMismatch(variant, key, target, targetBoundsByKey, zoneNameById, baseAssignment, variantAssignment, baseline.ZoneBoundsById, variant.ZoneBoundsById));
                        perVariantSamples++;
                    }
                    continue;
                }

                if (string.IsNullOrWhiteSpace(variantZone))
                {
                    baselineAssignedVariantUnassigned++;
                    if (perVariantSamples < mismatchLimit)
                    {
                        mismatchSamples.Add(BuildMismatch(variant, key, target, targetBoundsByKey, zoneNameById, baseAssignment, variantAssignment, baseline.ZoneBoundsById, variant.ZoneBoundsById));
                        perVariantSamples++;
                    }
                    continue;
                }

                if (string.Equals(baseZone, variantZone, StringComparison.OrdinalIgnoreCase))
                {
                    same++;
                }
                else
                {
                    different++;
                    if (perVariantSamples < mismatchLimit)
                    {
                        mismatchSamples.Add(BuildMismatch(variant, key, target, targetBoundsByKey, zoneNameById, baseAssignment, variantAssignment, baseline.ZoneBoundsById, variant.ZoneBoundsById));
                        perVariantSamples++;
                    }
                }
            }

            var samePercent = total == 0 ? 0 : (same * 100.0 / total);

            return new SpaceMapperComparisonAgreement
            {
                VariantName = variant.Report.Name,
                TotalTargets = total,
                SameZone = same,
                SameZonePercent = samePercent,
                BaselineAssignedVariantUnassigned = baselineAssignedVariantUnassigned,
                BaselineUnassignedVariantAssigned = baselineUnassignedVariantAssigned,
                DifferentZone = different
            };
        }

        private static SpaceMapperComparisonMismatch BuildMismatch(
            VariantRunResult variant,
            string targetKey,
            TargetGeometry target,
            Dictionary<string, Aabb> targetBoundsByKey,
            Dictionary<string, string> zoneNameById,
            Assignment baselineAssignment,
            Assignment variantAssignment,
            Dictionary<string, Aabb> baselineZoneBounds,
            Dictionary<string, Aabb> variantZoneBounds)
        {
            var targetBounds = targetBoundsByKey.TryGetValue(targetKey, out var bounds)
                ? bounds
                : new Aabb(0, 0, 0, 0, 0, 0);

            var center = new SpaceMapperComparisonPoint
            {
                X = (targetBounds.MinX + targetBounds.MaxX) * 0.5,
                Y = (targetBounds.MinY + targetBounds.MaxY) * 0.5,
                Z = (targetBounds.MinZ + targetBounds.MaxZ) * 0.5
            };

            var baselineZone = BuildZoneMismatchInfo(baselineAssignment?.ZoneId, zoneNameById, baselineZoneBounds, center);
            var variantZone = BuildZoneMismatchInfo(variantAssignment?.ZoneId, zoneNameById, variantZoneBounds, center);

            return new SpaceMapperComparisonMismatch
            {
                VariantName = variant.Report.Name,
                TargetKey = targetKey,
                TargetName = target?.DisplayName ?? string.Empty,
                TargetCenter = center,
                BaselineZone = baselineZone,
                VariantZone = variantZone
            };
        }

        private static SpaceMapperComparisonZone BuildZoneMismatchInfo(
            string zoneId,
            Dictionary<string, string> zoneNameById,
            Dictionary<string, Aabb> boundsById,
            SpaceMapperComparisonPoint center)
        {
            if (string.IsNullOrWhiteSpace(zoneId))
            {
                return null;
            }

            var zoneName = zoneNameById != null && zoneNameById.TryGetValue(zoneId, out var name)
                ? name
                : string.Empty;

            boundsById ??= new Dictionary<string, Aabb>(StringComparer.OrdinalIgnoreCase);
            boundsById.TryGetValue(zoneId, out var bounds);
            var contains = ContainsPoint(bounds, center);

            return new SpaceMapperComparisonZone
            {
                ZoneId = zoneId,
                ZoneName = zoneName,
                Bounds = ToBounds(bounds),
                ContainsPoint = contains
            };
        }

        private static bool ContainsPoint(in Aabb bounds, SpaceMapperComparisonPoint point)
        {
            return point.X >= bounds.MinX && point.X <= bounds.MaxX
                && point.Y >= bounds.MinY && point.Y <= bounds.MaxY
                && point.Z >= bounds.MinZ && point.Z <= bounds.MaxZ;
        }

        private static SpaceMapperComparisonBounds ToBounds(in Aabb bounds)
        {
            return new SpaceMapperComparisonBounds
            {
                MinX = bounds.MinX,
                MinY = bounds.MinY,
                MinZ = bounds.MinZ,
                MaxX = bounds.MaxX,
                MaxY = bounds.MaxY,
                MaxZ = bounds.MaxZ
            };
        }

        private static List<string> BuildWarnings(List<VariantRunResult> variants, VariantRunResult baseline, int zonesWithoutPlanes)
        {
            var warnings = new List<string>();

            if (zonesWithoutPlanes > 0)
            {
                var anyNormalOrAccurate = variants.Any(v => !v.Report.Skipped &&
                    (v.Definition.Preset == SpaceMapperPerformancePreset.Normal || v.Definition.Preset == SpaceMapperPerformancePreset.Accurate));

                if (anyNormalOrAccurate)
                {
                    warnings.Add("Normal/Accurate ran but some zones had 0 planes. Plane extraction may be incomplete.");
                }
            }

            if (baseline != null)
            {
                var baselineTargets = baseline.Report?.Workload?.TargetsProcessed ?? 0;
                foreach (var variant in variants)
                {
                    if (variant.Report.Skipped)
                    {
                        continue;
                    }

                    if (!variant.Settings.EnableMultipleZones && variant.MaxIntersectionsPerTarget > 1)
                    {
                        warnings.Add($"{variant.Report.Name}: EnableMultipleZones=false but variant returned multiple zones for a target.");
                    }

                    if (variant.Report.Workload.CandidatePairs == 0
                        && variant.Report.Workload.ZonesProcessed > 0
                        && variant.Report.Workload.TargetsProcessed > 0)
                    {
                        warnings.Add($"{variant.Report.Name}: CandidatePairs=0 but zones/targets > 0.");
                    }

                    if (variant.Report.Workload.TargetsProcessed != baselineTargets)
                    {
                        warnings.Add($"{variant.Report.Name}: TargetsProcessed differs from baseline.");
                    }

                    if (variant.Definition.Traversal == SpaceMapperFastTraversalMode.TargetMajor && variant.NeedsPartial)
                    {
                        warnings.Add($"{variant.Report.Name}: Target-major executed with partial options enabled.");
                    }
                }
            }

            return warnings;
        }

        private static SpaceMapperComparisonOutput BuildOutput(SpaceMapperComparisonReport report)
        {
            return new SpaceMapperComparisonOutput
            {
                Report = report,
                Markdown = BuildMarkdown(report),
                Json = SerializeJson(report)
            };
        }
        private static string BuildMarkdown(SpaceMapperComparisonReport report)
        {
            var sb = new StringBuilder();
            sb.AppendLine("# Space Mapper Benchmark Report");
            sb.AppendLine($"Generated: {report.GeneratedUtc}");
            sb.AppendLine();

            if (report.Environment != null)
            {
                sb.AppendLine("## Environment");
                sb.AppendLine($"- Model: {report.Environment.ModelName}");
                sb.AppendLine($"- Navisworks API: {report.Environment.NavisworksApiVersion}");
                sb.AppendLine($"- Plugin version: {report.Environment.PluginVersion}");
                sb.AppendLine($"- OS: {report.Environment.OSVersion}");
                sb.AppendLine($"- CPU cores: {report.Environment.CpuCores}");
                sb.AppendLine($"- Working set (MB): {report.Environment.ProcessWorkingSetMb:0.0}");
                sb.AppendLine($"- Machine: {report.Environment.MachineName}");
                sb.AppendLine();
            }

            if (report.Dataset != null)
            {
                sb.AppendLine("## Dataset");
                var rawZones = report.Dataset.RawZoneCount > 0 ? report.Dataset.RawZoneCount : report.Dataset.ZoneCount;
                var rawTargets = report.Dataset.RawTargetCount > 0 ? report.Dataset.RawTargetCount : report.Dataset.TargetCount;
                sb.AppendLine($"- Zones (raw): {rawZones:N0}");
                sb.AppendLine($"- Zones (with bbox): {report.Dataset.ZoneCount:N0}");
                sb.AppendLine($"- Targets (raw): {rawTargets:N0}");
                sb.AppendLine($"- Targets (with bbox): {report.Dataset.TargetCount:N0}");
                if (report.Dataset.Sampled)
                {
                    sb.AppendLine($"- Sampled zones: {report.Dataset.SampledZones:N0}");
                    sb.AppendLine($"- Sampled targets: {report.Dataset.SampledTargets:N0}");
                    sb.AppendLine($"- Sample seed: {report.Dataset.SampleSeed}");
                }
                sb.AppendLine($"- Resolve time (ms): {report.Dataset.ResolveMs:0}");
                sb.AppendLine($"- Build geometry time (ms): {report.Dataset.BuildGeometryMs:0}");
                sb.AppendLine($"- Zone source: {report.Dataset.ZoneSource}");
                sb.AppendLine($"- Target rules: {report.Dataset.TargetRuleCount}");
                sb.AppendLine();
            }

            if (report.Variants != null && report.Variants.Count > 0)
            {
                sb.AppendLine("## Variants");
                foreach (var variant in report.Variants)
                {
                    sb.AppendLine($"### {variant.Name}");
                    if (variant.Skipped)
                    {
                        sb.AppendLine($"- Status: Skipped ({variant.SkipReason})");
                        sb.AppendLine();
                        continue;
                    }

                    if (variant.Settings != null)
                    {
                        sb.AppendLine("- Settings:");
                        sb.AppendLine($"  - Preset: {variant.Settings.Preset} (resolved: {variant.Settings.PresetResolved})");
                        sb.AppendLine($"  - Traversal: {variant.Settings.TraversalRequested} (resolved: {variant.Settings.TraversalResolved})");
                        sb.AppendLine($"  - Processing mode: {variant.Settings.ProcessingMode}");
                        sb.AppendLine($"  - Index granularity: {variant.Settings.IndexGranularity}");
                        sb.AppendLine($"  - Max threads: {variant.Settings.MaxThreads}");
                        sb.AppendLine($"  - Batch size: {variant.Settings.BatchSize}");
                        sb.AppendLine($"  - Use origin point only: {variant.Settings.UseOriginPointOnly}");
                        sb.AppendLine($"  - Treat partial as contained: {variant.Settings.TreatPartialAsContained}");
                        sb.AppendLine($"  - Tag partial separately: {variant.Settings.TagPartialSeparately}");
                        sb.AppendLine($"  - Enable multiple zones: {variant.Settings.EnableMultipleZones}");
                        sb.AppendLine($"  - Best zone behavior: {variant.Settings.BestZoneBehavior}");
                        sb.AppendLine($"  - Benchmark mode: {variant.Settings.BenchmarkMode}");
                        sb.AppendLine($"  - Writeback strategy: {variant.Settings.WritebackStrategy}");
                        sb.AppendLine($"  - Show internal properties during writeback: {variant.Settings.ShowInternalPropertiesDuringWriteback}");
                        sb.AppendLine($"  - Skip unchanged writeback: {variant.Settings.SkipUnchangedWriteback}");
                        sb.AppendLine($"  - Pack writeback outputs: {variant.Settings.PackWritebackProperties}");
                    }

                    if (variant.Timings != null)
                    {
                        sb.AppendLine("- Timings (ms):");
                        sb.AppendLine($"  - Preflight: {variant.Timings.PreflightMs:0} (build {variant.Timings.PreflightBuildMs:0}, query {variant.Timings.PreflightQueryMs:0})");
                        sb.AppendLine($"  - Index build: {variant.Timings.IndexBuildMs:0}");
                        sb.AppendLine($"  - Candidate query: {variant.Timings.CandidateQueryMs:0}");
                        sb.AppendLine($"  - Narrow phase: {variant.Timings.NarrowPhaseMs:0}");
                        sb.AppendLine($"  - Postprocess/writeback: {variant.Timings.PostprocessMs:0}");
                        sb.AppendLine($"  - Total: {variant.Timings.TotalMs:0}");
                    }

                    if (variant.Workload != null)
                    {
                        sb.AppendLine("- Workload:");
                        sb.AppendLine($"  - Zones processed: {variant.Workload.ZonesProcessed:N0}");
                        sb.AppendLine($"  - Targets processed: {variant.Workload.TargetsProcessed:N0}");
                        sb.AppendLine($"  - Candidate pairs: {variant.Workload.CandidatePairs:N0}");
                        sb.AppendLine($"  - Avg candidates/zone: {variant.Workload.AvgCandidatesPerZone:N0}");
                        sb.AppendLine($"  - Avg candidates/target: {variant.Workload.AvgCandidatesPerTarget:N0}");
                        sb.AppendLine($"  - Max candidates/zone: {variant.Workload.MaxCandidatesPerZone:N0}");
                        sb.AppendLine($"  - Max candidates/target: {variant.Workload.MaxCandidatesPerTarget:N0}");
                        sb.AppendLine($"  - Contained hits: {variant.Workload.ContainedHits:N0}");
                        sb.AppendLine($"  - Partial hits: {variant.Workload.PartialHits:N0}");
                        sb.AppendLine($"  - Unmatched targets: {variant.Workload.UnmatchedTargets:N0}");
                        sb.AppendLine($"  - Multi-zone targets: {variant.Workload.MultiZoneTargets:N0}");
                        if (variant.Workload.WritesPerformed > 0)
                        {
                            sb.AppendLine($"  - Writes performed: {variant.Workload.WritesPerformed:N0}");
                            if (variant.Workload.WritebackTargetsWritten > 0)
                            {
                                sb.AppendLine($"  - Targets written: {variant.Workload.WritebackTargetsWritten:N0}");
                            }
                            if (variant.Workload.WritebackCategoriesWritten > 0)
                            {
                                sb.AppendLine($"  - Categories written: {variant.Workload.WritebackCategoriesWritten:N0}");
                            }
                            if (variant.Workload.WritebackPropertiesWritten > 0)
                            {
                                sb.AppendLine($"  - Properties written: {variant.Workload.WritebackPropertiesWritten:N0}");
                            }
                            if (variant.Workload.AvgMsPerCategoryWrite > 0)
                            {
                                sb.AppendLine($"  - Avg ms/category write: {variant.Workload.AvgMsPerCategoryWrite:0.##}");
                            }
                            if (variant.Workload.AvgMsPerTargetWrite > 0)
                            {
                                sb.AppendLine($"  - Avg ms/target write: {variant.Workload.AvgMsPerTargetWrite:0.##}");
                            }
                        }
                        if (variant.Workload.SkippedUnchanged > 0)
                        {
                            sb.AppendLine($"  - Skipped unchanged: {variant.Workload.SkippedUnchanged:N0}");
                        }
                    }

                    if (variant.Diagnostics != null)
                    {
                        sb.AppendLine("- Diagnostics:");
                        sb.AppendLine($"  - Used preflight index: {variant.Diagnostics.UsedPreflightIndex}");
                        sb.AppendLine($"  - Preflight candidate pairs: {variant.Diagnostics.PreflightCandidatePairs:N0}");
                        sb.AppendLine($"  - Preflight traversal: {variant.Diagnostics.PreflightTraversalUsed}");
                    }

                    sb.AppendLine();
                }
            }

            if (report.Comparisons != null && report.Comparisons.Count > 0)
            {
                sb.AppendLine("## Agreement vs Baseline");
                foreach (var cmp in report.Comparisons)
                {
                    sb.AppendLine($"- {cmp.VariantName}: same zone {cmp.SameZonePercent:0.0}% ({cmp.SameZone}/{cmp.TotalTargets}), baseline assigned vs variant unassigned {cmp.BaselineAssignedVariantUnassigned}, baseline unassigned vs variant assigned {cmp.BaselineUnassignedVariantAssigned}, different zones {cmp.DifferentZone}");
                }
                sb.AppendLine();
            }

            if (report.MismatchSamples != null && report.MismatchSamples.Count > 0)
            {
                sb.AppendLine("## Mismatch Samples");
                foreach (var sample in report.MismatchSamples)
                {
                    sb.AppendLine($"### {sample.VariantName}");
                    sb.AppendLine($"- Target: {sample.TargetKey} ({sample.TargetName})");
                    if (sample.TargetCenter != null)
                    {
                        sb.AppendLine($"- Target center: ({sample.TargetCenter.X:0.###}, {sample.TargetCenter.Y:0.###}, {sample.TargetCenter.Z:0.###})");
                    }

                    AppendZoneSample(sb, "Baseline", sample.BaselineZone);
                    AppendZoneSample(sb, "Variant", sample.VariantZone);
                    sb.AppendLine();
                }
            }

            if (report.Warnings != null && report.Warnings.Count > 0)
            {
                sb.AppendLine("## Sanity Checks");
                foreach (var warning in report.Warnings)
                {
                    sb.AppendLine($"- {warning}");
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private static void AppendZoneSample(StringBuilder sb, string label, SpaceMapperComparisonZone zone)
        {
            if (zone == null)
            {
                sb.AppendLine($"- {label} zone: (none)");
                return;
            }

            sb.AppendLine($"- {label} zone: {zone.ZoneId} ({zone.ZoneName})");
            if (zone.Bounds != null)
            {
                sb.AppendLine($"  - Bounds: min({zone.Bounds.MinX:0.###}, {zone.Bounds.MinY:0.###}, {zone.Bounds.MinZ:0.###}) max({zone.Bounds.MaxX:0.###}, {zone.Bounds.MaxY:0.###}, {zone.Bounds.MaxZ:0.###})");
                sb.AppendLine($"  - Target inside: {zone.ContainsPoint}");
            }
        }

        private static string SerializeJson(SpaceMapperComparisonReport report)
        {
            var serializer = new DataContractJsonSerializer(typeof(SpaceMapperComparisonReport));
            using var stream = new MemoryStream();
            serializer.WriteObject(stream, report);
            return Encoding.UTF8.GetString(stream.ToArray());
        }

        private static void ReportProgress(IProgress<SpaceMapperComparisonProgress> progress, string stage, int percent)
        {
            progress?.Report(new SpaceMapperComparisonProgress
            {
                Stage = stage,
                Percentage = percent
            });
        }

        private static int ComputePercent(int index, int total, double offset)
        {
            if (total <= 0) return 0;
            var pct = ((index + offset) / total) * 100.0;
            return (int)Math.Max(0, Math.Min(100, pct));
        }

        private static List<T> SampleList<T>(
            IReadOnlyList<T> source,
            int maxCount,
            Random rng,
            out bool sampled,
            CancellationToken token)
        {
            var list = source.ToList();
            if (list.Count <= maxCount)
            {
                sampled = false;
                return list;
            }

            sampled = true;
            for (int i = list.Count - 1; i > 0; i--)
            {
                token.ThrowIfCancellationRequested();
                var j = rng.Next(i + 1);
                (list[i], list[j]) = (list[j], list[i]);
            }

            return list.Take(maxCount).ToList();
        }

        private static Dictionary<string, List<SpaceMapperTargetRule>> FilterTargetsByRule(
            Dictionary<string, List<SpaceMapperTargetRule>> targetsByRule,
            IReadOnlyList<TargetGeometry> targets)
        {
            var filtered = new Dictionary<string, List<SpaceMapperTargetRule>>(StringComparer.OrdinalIgnoreCase);
            if (targetsByRule == null || targets == null)
            {
                return filtered;
            }

            var keySet = new HashSet<string>(targets.Select(t => t.ItemKey), StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in targetsByRule)
            {
                if (keySet.Contains(kvp.Key))
                {
                    filtered[kvp.Key] = kvp.Value;
                }
            }

            return filtered;
        }

        private static TargetBoundsData BuildTargetBounds(IReadOnlyList<TargetGeometry> targets)
        {
            var bounds = new Aabb[targets.Count];
            var keys = new string[targets.Count];
            var byKey = new Dictionary<string, Aabb>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < targets.Count; i++)
            {
                var target = targets[i];
                var bbox = target?.BoundingBox;
                if (bbox == null)
                {
                    bounds[i] = new Aabb(0, 0, 0, 0, 0, 0);
                    keys[i] = target?.ItemKey ?? string.Empty;
                    continue;
                }

                var aabb = ToAabb(bbox);
                bounds[i] = aabb;
                keys[i] = target.ItemKey;
                if (!string.IsNullOrWhiteSpace(target.ItemKey) && !byKey.ContainsKey(target.ItemKey))
                {
                    byKey[target.ItemKey] = aabb;
                }
            }

            return new TargetBoundsData
            {
                Bounds = bounds,
                Keys = keys,
                ByKey = byKey
            };
        }

        private static ZoneBoundsData BuildZoneBounds(IReadOnlyList<ZoneGeometry> zones, SpaceMapperProcessingSettings settings)
        {
            var bounds = new Aabb[zones.Count];
            var byId = new Dictionary<string, Aabb>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < zones.Count; i++)
            {
                var zone = zones[i];
                var aabb = GetZoneQueryBounds(zone, settings);
                bounds[i] = aabb;

                if (!string.IsNullOrWhiteSpace(zone?.ZoneId) && !byId.ContainsKey(zone.ZoneId))
                {
                    byId[zone.ZoneId] = aabb;
                }
            }

            return new ZoneBoundsData
            {
                Bounds = bounds,
                ById = byId
            };
        }

        private static Dictionary<string, string> BuildZoneNameLookup(IReadOnlyList<ZoneGeometry> zones)
        {
            var byId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var zone in zones)
            {
                if (zone == null || string.IsNullOrWhiteSpace(zone.ZoneId))
                {
                    continue;
                }

                if (!byId.ContainsKey(zone.ZoneId))
                {
                    byId[zone.ZoneId] = zone.DisplayName ?? string.Empty;
                }
            }

            return byId;
        }

        private static Aabb GetZoneQueryBounds(ZoneGeometry zone, SpaceMapperProcessingSettings settings)
        {
            if (zone == null)
            {
                return new Aabb(0, 0, 0, 0, 0, 0);
            }

            var baseBox = zone.RawBoundingBox ?? zone.BoundingBox;
            if (baseBox == null)
            {
                return new Aabb(0, 0, 0, 0, 0, 0);
            }

            var aabb = ToAabb(baseBox);
            if (zone.RawBoundingBox == null)
            {
                return aabb;
            }

            return Inflate(aabb, settings);
        }

        private static Aabb Inflate(Aabb bbox, SpaceMapperProcessingSettings settings)
        {
            if (settings == null)
            {
                return bbox;
            }

            var offset = settings.Offset3D;
            var offsetSides = settings.OffsetSides;
            var top = settings.OffsetTop;
            var bottom = settings.OffsetBottom;

            return new Aabb(
                bbox.MinX - offset - offsetSides,
                bbox.MinY - offset - offsetSides,
                bbox.MinZ - offset - bottom,
                bbox.MaxX + offset + offsetSides,
                bbox.MaxY + offset + offsetSides,
                bbox.MaxZ + offset + top);
        }

        private static Aabb ToAabb(BoundingBox3D bbox)
        {
            var min = bbox.Min;
            var max = bbox.Max;
            return new Aabb(min.X, min.Y, min.Z, max.X, max.Y, max.Z);
        }
        private sealed class VariantDefinition
        {
            public VariantDefinition(
                string name,
                SpaceMapperPerformancePreset preset,
                SpaceMapperFastTraversalMode traversal,
                bool useOriginPointOnly,
                bool requiresNoPartial = false,
                bool isBaseline = false,
                bool legacyFastAabbOnly = false)
            {
                Name = name;
                Preset = preset;
                Traversal = traversal;
                UseOriginPointOnly = useOriginPointOnly;
                RequiresNoPartial = requiresNoPartial;
                IsBaseline = isBaseline;
                LegacyFastAabbOnly = legacyFastAabbOnly;
            }

            public string Name { get; }
            public SpaceMapperPerformancePreset Preset { get; }
            public SpaceMapperFastTraversalMode Traversal { get; }
            public bool UseOriginPointOnly { get; }
            public bool RequiresNoPartial { get; }
            public bool IsBaseline { get; }
            public bool LegacyFastAabbOnly { get; }
        }

        private sealed class Assignment
        {
            public string ZoneId { get; set; }
            public bool IsContained { get; set; }
            public bool IsPartial { get; set; }
        }

        private sealed class VariantAssignmentSummary
        {
            public Dictionary<string, Assignment> Assignments { get; } = new(StringComparer.OrdinalIgnoreCase);
            public int ContainedHits { get; set; }
            public int PartialHits { get; set; }
            public int UnmatchedTargets { get; set; }
            public int MultiZoneTargets { get; set; }
            public int MaxIntersectionsPerTarget { get; set; }
        }

        private sealed class VariantRunResult
        {
            public VariantDefinition Definition { get; set; }
            public SpaceMapperComparisonVariant Report { get; set; }
            public Dictionary<string, Assignment> Assignments { get; set; }
            public Dictionary<string, Aabb> ZoneBoundsById { get; set; }
            public SpaceMapperProcessingSettings Settings { get; set; }
            public int TargetsProcessed { get; set; }
            public int ZonesProcessed { get; set; }
            public int MaxIntersectionsPerTarget { get; set; }
            public bool NeedsPartial { get; set; }

            public static VariantRunResult Skipped(
                VariantDefinition definition,
                SpaceMapperProcessingSettings settings,
                string reason,
                SpaceMapperBenchmarkMode benchmarkMode)
            {
                return new VariantRunResult
                {
                    Definition = definition,
                    Settings = settings,
                    Assignments = new Dictionary<string, Assignment>(StringComparer.OrdinalIgnoreCase),
                    ZoneBoundsById = new Dictionary<string, Aabb>(StringComparer.OrdinalIgnoreCase),
                    Report = new SpaceMapperComparisonVariant
                    {
                        Name = definition.Name,
                        Skipped = true,
                        SkipReason = reason,
                        Settings = new SpaceMapperComparisonSettings
                        {
                            Preset = definition.Preset.ToString(),
                            PresetResolved = definition.Preset.ToString(),
                            TraversalRequested = definition.Traversal.ToString(),
                            TraversalResolved = string.Empty,
                            ProcessingMode = settings?.ProcessingMode.ToString() ?? string.Empty,
                            IndexGranularity = settings?.IndexGranularity ?? 0,
                            MaxThreads = settings?.MaxThreads ?? 0,
                            BatchSize = settings?.BatchSize ?? 0,
                            UseOriginPointOnly = settings?.UseOriginPointOnly ?? false,
                            TreatPartialAsContained = settings?.TreatPartialAsContained ?? false,
                            TagPartialSeparately = settings?.TagPartialSeparately ?? false,
                            EnableMultipleZones = settings?.EnableMultipleZones ?? false,
                            BestZoneBehavior = settings?.EnableMultipleZones == true ? "Multiple" : "Single",
                            BenchmarkMode = benchmarkMode.ToString(),
                            WritebackStrategy = settings?.WritebackStrategy.ToString() ?? SpaceMapperWritebackStrategy.OptimizedSingleCategory.ToString(),
                            ShowInternalPropertiesDuringWriteback = settings?.ShowInternalPropertiesDuringWriteback ?? false,
                            SkipUnchangedWriteback = settings?.SkipUnchangedWriteback ?? false,
                            PackWritebackProperties = settings?.PackWritebackProperties ?? false
                        },
                        Timings = new SpaceMapperComparisonTimings(),
                        Workload = new SpaceMapperComparisonWorkload(),
                        Diagnostics = new SpaceMapperComparisonDiagnostics()
                    }
                };
            }
        }

        private sealed class TargetBoundsData
        {
            public Aabb[] Bounds { get; set; }
            public string[] Keys { get; set; }
            public Dictionary<string, Aabb> ByKey { get; set; }
        }

        private sealed class ZoneBoundsData
        {
            public Aabb[] Bounds { get; set; }
            public Dictionary<string, Aabb> ById { get; set; }
        }
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonReport
    {
        [DataMember(Order = 0)]
        public string SchemaVersion { get; set; }

        [DataMember(Order = 1)]
        public string GeneratedUtc { get; set; }

        [DataMember(Order = 2)]
        public SpaceMapperComparisonEnvironment Environment { get; set; }

        [DataMember(Order = 3)]
        public SpaceMapperComparisonDataset Dataset { get; set; }

        [DataMember(Order = 4)]
        public List<SpaceMapperComparisonVariant> Variants { get; set; } = new();

        [DataMember(Order = 5)]
        public List<SpaceMapperComparisonAgreement> Comparisons { get; set; } = new();

        [DataMember(Order = 6)]
        public List<SpaceMapperComparisonMismatch> MismatchSamples { get; set; } = new();

        [DataMember(Order = 7)]
        public List<string> Warnings { get; set; } = new();
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonEnvironment
    {
        [DataMember(Order = 0)]
        public string NavisworksApiVersion { get; set; }

        [DataMember(Order = 1)]
        public string PluginVersion { get; set; }

        [DataMember(Order = 2)]
        public string OSVersion { get; set; }

        [DataMember(Order = 3)]
        public int CpuCores { get; set; }

        [DataMember(Order = 4)]
        public double ProcessWorkingSetMb { get; set; }

        [DataMember(Order = 5)]
        public string MachineName { get; set; }

        [DataMember(Order = 6)]
        public string ModelName { get; set; }
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonDataset
    {
        [DataMember(Order = 0)]
        public string ModelName { get; set; }

        [DataMember(Order = 1)]
        public int ZoneCount { get; set; }

        [DataMember(Order = 2)]
        public int TargetCount { get; set; }

        [DataMember(Order = 3)]
        public int RawZoneCount { get; set; }

        [DataMember(Order = 4)]
        public int RawTargetCount { get; set; }

        [DataMember(Order = 5)]
        public bool Sampled { get; set; }

        [DataMember(Order = 6)]
        public int SampleSeed { get; set; }

        [DataMember(Order = 7)]
        public int SampledZones { get; set; }

        [DataMember(Order = 8)]
        public int SampledTargets { get; set; }

        [DataMember(Order = 9)]
        public double ResolveMs { get; set; }

        [DataMember(Order = 10)]
        public double BuildGeometryMs { get; set; }

        [DataMember(Order = 11)]
        public string ZoneSource { get; set; }

        [DataMember(Order = 12)]
        public int TargetRuleCount { get; set; }
    }
    [DataContract]
    internal sealed class SpaceMapperComparisonVariant
    {
        [DataMember(Order = 0)]
        public string Name { get; set; }

        [DataMember(Order = 1)]
        public bool Skipped { get; set; }

        [DataMember(Order = 2)]
        public string SkipReason { get; set; }

        [DataMember(Order = 3)]
        public SpaceMapperComparisonSettings Settings { get; set; }

        [DataMember(Order = 4)]
        public SpaceMapperComparisonTimings Timings { get; set; }

        [DataMember(Order = 5)]
        public SpaceMapperComparisonWorkload Workload { get; set; }

        [DataMember(Order = 6)]
        public SpaceMapperComparisonDiagnostics Diagnostics { get; set; }
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonSettings
    {
        [DataMember(Order = 0)]
        public string Preset { get; set; }

        [DataMember(Order = 1)]
        public string PresetResolved { get; set; }

        [DataMember(Order = 2)]
        public string TraversalRequested { get; set; }

        [DataMember(Order = 3)]
        public string TraversalResolved { get; set; }

        [DataMember(Order = 4)]
        public string ProcessingMode { get; set; }

        [DataMember(Order = 5)]
        public int IndexGranularity { get; set; }

        [DataMember(Order = 6)]
        public int MaxThreads { get; set; }

        [DataMember(Order = 7)]
        public int BatchSize { get; set; }

        [DataMember(Order = 8)]
        public bool UseOriginPointOnly { get; set; }

        [DataMember(Order = 9)]
        public bool TreatPartialAsContained { get; set; }

        [DataMember(Order = 10)]
        public bool TagPartialSeparately { get; set; }

        [DataMember(Order = 11)]
        public bool EnableMultipleZones { get; set; }

        [DataMember(Order = 12)]
        public double Offset3D { get; set; }

        [DataMember(Order = 13)]
        public double OffsetTop { get; set; }

        [DataMember(Order = 14)]
        public double OffsetBottom { get; set; }

        [DataMember(Order = 15)]
        public double OffsetSides { get; set; }

        [DataMember(Order = 16)]
        public string OffsetMode { get; set; }

        [DataMember(Order = 17)]
        public string Units { get; set; }

        [DataMember(Order = 18)]
        public string BestZoneBehavior { get; set; }

        [DataMember(Order = 19)]
        public int ZonesWithoutPlanes { get; set; }

        [DataMember(Order = 20)]
        public string BenchmarkMode { get; set; }

        [DataMember(Order = 21)]
        public string WritebackStrategy { get; set; }

        [DataMember(Order = 22)]
        public bool ShowInternalPropertiesDuringWriteback { get; set; }

        [DataMember(Order = 23)]
        public bool SkipUnchangedWriteback { get; set; }

        [DataMember(Order = 24)]
        public bool PackWritebackProperties { get; set; }
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonTimings
    {
        [DataMember(Order = 0)]
        public double ResolveMs { get; set; }

        [DataMember(Order = 1)]
        public double BuildGeometryMs { get; set; }

        [DataMember(Order = 2)]
        public double PreflightMs { get; set; }

        [DataMember(Order = 3)]
        public double PreflightBuildMs { get; set; }

        [DataMember(Order = 4)]
        public double PreflightQueryMs { get; set; }

        [DataMember(Order = 5)]
        public double IndexBuildMs { get; set; }

        [DataMember(Order = 6)]
        public double CandidateQueryMs { get; set; }

        [DataMember(Order = 7)]
        public double NarrowPhaseMs { get; set; }

        [DataMember(Order = 8)]
        public double PostprocessMs { get; set; }

        [DataMember(Order = 9)]
        public double TotalMs { get; set; }
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonWorkload
    {
        [DataMember(Order = 0)]
        public int ZonesProcessed { get; set; }

        [DataMember(Order = 1)]
        public int TargetsProcessed { get; set; }

        [DataMember(Order = 2)]
        public long CandidatePairs { get; set; }

        [DataMember(Order = 3)]
        public double AvgCandidatesPerZone { get; set; }

        [DataMember(Order = 4)]
        public int MaxCandidatesPerZone { get; set; }

        [DataMember(Order = 5)]
        public double AvgCandidatesPerTarget { get; set; }

        [DataMember(Order = 6)]
        public int MaxCandidatesPerTarget { get; set; }

        [DataMember(Order = 7)]
        public int ContainedHits { get; set; }

        [DataMember(Order = 8)]
        public int PartialHits { get; set; }

        [DataMember(Order = 9)]
        public int UnmatchedTargets { get; set; }

        [DataMember(Order = 10)]
        public int MultiZoneTargets { get; set; }

        [DataMember(Order = 11)]
        public int MaxIntersectionsPerTarget { get; set; }

        [DataMember(Order = 12)]
        public long WritesPerformed { get; set; }

        [DataMember(Order = 13)]
        public int WritebackTargetsWritten { get; set; }

        [DataMember(Order = 14)]
        public int WritebackCategoriesWritten { get; set; }

        [DataMember(Order = 15)]
        public int WritebackPropertiesWritten { get; set; }

        [DataMember(Order = 16)]
        public double AvgMsPerCategoryWrite { get; set; }

        [DataMember(Order = 17)]
        public double AvgMsPerTargetWrite { get; set; }

        [DataMember(Order = 18)]
        public int SkippedUnchanged { get; set; }
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonDiagnostics
    {
        [DataMember(Order = 0)]
        public bool UsedPreflightIndex { get; set; }

        [DataMember(Order = 1)]
        public long PreflightCandidatePairs { get; set; }

        [DataMember(Order = 2)]
        public string PreflightTraversalUsed { get; set; }
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonAgreement
    {
        [DataMember(Order = 0)]
        public string VariantName { get; set; }

        [DataMember(Order = 1)]
        public int TotalTargets { get; set; }

        [DataMember(Order = 2)]
        public int SameZone { get; set; }

        [DataMember(Order = 3)]
        public double SameZonePercent { get; set; }

        [DataMember(Order = 4)]
        public int BaselineAssignedVariantUnassigned { get; set; }

        [DataMember(Order = 5)]
        public int BaselineUnassignedVariantAssigned { get; set; }

        [DataMember(Order = 6)]
        public int DifferentZone { get; set; }
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonMismatch
    {
        [DataMember(Order = 0)]
        public string VariantName { get; set; }

        [DataMember(Order = 1)]
        public string TargetKey { get; set; }

        [DataMember(Order = 2)]
        public string TargetName { get; set; }

        [DataMember(Order = 3)]
        public SpaceMapperComparisonPoint TargetCenter { get; set; }

        [DataMember(Order = 4)]
        public SpaceMapperComparisonZone BaselineZone { get; set; }

        [DataMember(Order = 5)]
        public SpaceMapperComparisonZone VariantZone { get; set; }
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonZone
    {
        [DataMember(Order = 0)]
        public string ZoneId { get; set; }

        [DataMember(Order = 1)]
        public string ZoneName { get; set; }

        [DataMember(Order = 2)]
        public SpaceMapperComparisonBounds Bounds { get; set; }

        [DataMember(Order = 3)]
        public bool ContainsPoint { get; set; }
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonBounds
    {
        [DataMember(Order = 0)]
        public double MinX { get; set; }

        [DataMember(Order = 1)]
        public double MinY { get; set; }

        [DataMember(Order = 2)]
        public double MinZ { get; set; }

        [DataMember(Order = 3)]
        public double MaxX { get; set; }

        [DataMember(Order = 4)]
        public double MaxY { get; set; }

        [DataMember(Order = 5)]
        public double MaxZ { get; set; }
    }

    [DataContract]
    internal sealed class SpaceMapperComparisonPoint
    {
        [DataMember(Order = 0)]
        public double X { get; set; }

        [DataMember(Order = 1)]
        public double Y { get; set; }

        [DataMember(Order = 2)]
        public double Z { get; set; }
    }
}
