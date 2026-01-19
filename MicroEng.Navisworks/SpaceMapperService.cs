using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.Threading;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Navisworks.Api;
using MicroEng.Navisworks.SpaceMapper.Estimation;
using MicroEng.Navisworks.SpaceMapper.Geometry;
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
        public string ReportPath { get; set; }
        public List<ModelItem> TargetsWithoutBounds { get; set; } = new();
        public List<ModelItem> TargetsUnmatched { get; set; } = new();
    }

    internal sealed class SpaceMapperWritebackResult
    {
        public int ContainedTagged { get; set; }
        public int PartialTagged { get; set; }
        public int MultiZoneTagged { get; set; }
        public int Skipped { get; set; }
        public int SkippedUnchanged { get; set; }
        public long WritesPerformed { get; set; }
        public int TargetsWritten { get; set; }
        public int CategoriesWritten { get; set; }
        public int PropertiesWritten { get; set; }
        public double AvgMsPerCategoryWrite { get; set; }
        public double AvgMsPerTargetWrite { get; set; }
        public SpaceMapperWritebackStrategy StrategyUsed { get; set; } = SpaceMapperWritebackStrategy.OptimizedSingleCategory;
        public TimeSpan Elapsed { get; set; }
        public List<ZoneSummary> ZoneSummaries { get; set; } = new();
    }

    internal sealed class SpaceMapperResolvedData
    {
        public List<SpaceMapperResolvedItem> ZoneModels { get; set; } = new();
        public List<TargetGeometry> TargetModels { get; set; } = new();
        public Dictionary<string, List<SpaceMapperTargetRule>> TargetsByRule { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public TimeSpan ResolveTime { get; set; }
    }

    internal sealed class SpaceMapperComputeDataset
    {
        public List<ZoneGeometry> Zones { get; set; } = new();
        public List<TargetGeometry> TargetsForEngine { get; set; } = new();
        public List<SpaceMapperResolvedItem> ZoneModels { get; set; } = new();
        public List<TargetGeometry> TargetModels { get; set; } = new();
        public Dictionary<string, List<SpaceMapperTargetRule>> TargetsByRule { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public TimeSpan ResolveTime { get; set; }
        public TimeSpan BuildGeometryTime { get; set; }
        public int ZonesWithMesh { get; set; }
        public int ZonesMeshFallback { get; set; }
        public int MeshExtractionErrors { get; set; }
    }

    internal class SpaceMapperService
    {
        private readonly Action<string> _log;
        private const string PackedWritebackPropertyName = "SpaceMapperPacked";
        private const string WritebackSignaturePropertyName = "SpaceMapperSignature";
        private const int MaxSequenceOutputsPerMapping = 8;
        private const string ZoneContainmentPercentPropertyName = "Zone Containment %";
        private const string ZoneOffsetMatchPropertyName = "Zone Offset Match";
        private const double ContainmentFullTolerance = 1e-6;

        public SpaceMapperService(Action<string> log)
        {
            _log = log;
        }

        public SpaceMapperRunResult RunWithProgress(
            SpaceMapperRequest request,
            SpaceMapperPreflightCache preflightCache,
            SpaceMapperRunProgressState runProgress,
            CancellationToken token = default)
        {
            return RunInternal(request, preflightCache, token, runProgress);
        }

        public SpaceMapperRunResult Run(
            SpaceMapperRequest request,
            SpaceMapperPreflightCache preflightCache = null,
            CancellationToken token = default)
        {
            return RunInternal(request, preflightCache, token, null);
        }

        private SpaceMapperRunResult RunInternal(
            SpaceMapperRequest request,
            SpaceMapperPreflightCache preflightCache,
            CancellationToken token,
            SpaceMapperRunProgressState runProgress)
        {
            runProgress?.Start();
            runProgress?.SetStage(SpaceMapperRunStage.ResolvingInputs, "Resolving zones and targets...");

            var sw = Stopwatch.StartNew();
            var result = new SpaceMapperRunResult();
            var stats = new SpaceMapperRunStats();
            result.Stats = stats;
            SpaceMapperResolvedData resolved = null;
            SpaceMapperComputeDataset dataset = null;
            SpaceMapperRunReportWriter.SpaceMapperLiveReportCache liveReport = null;
            Timer liveTimer = null;
            var liveInterval = TimeSpan.FromSeconds(10);
            var liveWriteGate = 0;

            void WriteLiveReportSnapshot()
            {
                if (liveReport == null)
                {
                    return;
                }

                if (Interlocked.Exchange(ref liveWriteGate, 1) == 1)
                {
                    return;
                }

                try
                {
                    SpaceMapperRunReportWriter.TryWriteLiveReport(
                        liveReport,
                        request,
                        resolved,
                        dataset,
                        result,
                        runProgress);
                }
                finally
                {
                    Interlocked.Exchange(ref liveWriteGate, 0);
                }
            }

            try
            {
                if (runProgress != null)
                {
                    liveReport = SpaceMapperRunReportWriter.CreateLiveReportCache(request);
                    if (!string.IsNullOrWhiteSpace(liveReport?.ReportPath))
                    {
                        _log?.Invoke($"SpaceMapper live report: {liveReport.ReportPath}");
                    }
                    WriteLiveReportSnapshot();
                    liveTimer = new Timer(_ => WriteLiveReportSnapshot(), null, liveInterval, liveInterval);
                }

                var doc = Application.ActiveDocument;
                if (doc == null)
                {
                    result.Message = "No active document.";
                    runProgress?.MarkFailed(new InvalidOperationException(result.Message));
                    return result;
                }

                var session = GetSession(request.ScraperProfileName);
                var requiresSession = request.ZoneSource == ZoneSourceType.DataScraperZones;
                if (requiresSession && session == null)
                {
                    result.Message = $"No Data Scraper sessions found for profile '{request.ScraperProfileName ?? "Default"}'.";
                    runProgress?.MarkFailed(new InvalidOperationException(result.Message));
                    return result;
                }

                if (session != null)
                {
                    DataScraperCache.LastSession = session;
                }

                resolved = ResolveData(request, doc, session);
                stats.ResolveTime = resolved.ResolveTime;
                runProgress?.SetTotals(resolved.ZoneModels.Count, resolved.TargetModels.Count);

                if (!resolved.ZoneModels.Any())
                {
                    result.Message = "No zones found.";
                    runProgress?.MarkFailed(new InvalidOperationException(result.Message));
                    return result;
                }

                if (!resolved.TargetModels.Any())
                {
                    result.Message = "No targets found for the selected rules.";
                    runProgress?.MarkFailed(new InvalidOperationException(result.Message));
                    return result;
                }

                runProgress?.SetStage(SpaceMapperRunStage.ExtractingGeometry, "Preparing geometry...");

                var cacheToUse = TryGetReusablePreflightCache(request, preflightCache, resolved.TargetModels);
                dataset = BuildGeometryData(
                    resolved,
                    request.ProcessingSettings,
                    buildTargetGeometry: cacheToUse == null,
                    token,
                    runProgress);
                stats.BuildGeometryTime = dataset.BuildGeometryTime;

                if (!dataset.Zones.Any())
                {
                    result.Message = "No zones found.";
                    runProgress?.MarkFailed(new InvalidOperationException(result.Message));
                    return result;
                }

                if (cacheToUse == null && !dataset.TargetsForEngine.Any())
                {
                    result.Message = "No targets with bounding boxes found.";
                    runProgress?.MarkFailed(new InvalidOperationException(result.Message));
                    return result;
                }

                runProgress?.SetStage(SpaceMapperRunStage.BuildingIndex, "Building spatial index...");

                var engine = SpaceMapperEngineFactory.Create(request.ProcessingSettings.ProcessingMode);
                var diagnostics = new SpaceMapperEngineDiagnostics();
                diagnostics.TargetsTotal = dataset.TargetModels.Count;
                diagnostics.TargetsWithBounds = dataset.TargetsForEngine.Count;
                diagnostics.TargetsWithoutBounds = Math.Max(0, diagnostics.TargetsTotal - diagnostics.TargetsWithBounds);
                IList<ZoneTargetIntersection> baselineIntersections = null;
                if (request.ProcessingSettings.EnableZoneOffsets
                    && request.ProcessingSettings.EnableOffsetAreaPass)
                {
                    var baselineSettings = CloneProcessingSettings(request.ProcessingSettings);
                    baselineSettings.EnableZoneOffsets = false;
                    baselineSettings.EnableOffsetAreaPass = false;
                    baselineSettings.WriteZoneOffsetMatchProperty = false;
                    baselineSettings.Offset3D = 0;
                    baselineSettings.OffsetTop = 0;
                    baselineSettings.OffsetBottom = 0;
                    baselineSettings.OffsetSides = 0;
                    baselineSettings.OffsetMode = "None";

                    baselineIntersections = engine.ComputeIntersections(
                            dataset.Zones,
                            dataset.TargetsForEngine,
                            baselineSettings,
                            null,
                            new SpaceMapperEngineDiagnostics(),
                            null,
                            token,
                            runProgress)
                        ?? new List<ZoneTargetIntersection>();
                }
                var intersections = engine.ComputeIntersections(
                        dataset.Zones,
                        dataset.TargetsForEngine,
                        request.ProcessingSettings,
                        cacheToUse,
                        diagnostics,
                        null,
                        token,
                        runProgress)
                    ?? new List<ZoneTargetIntersection>();

                if (baselineIntersections != null && baselineIntersections.Count > 0)
                {
                    var baselineSet = new HashSet<string>(
                        baselineIntersections
                            .Where(i => i != null && !string.IsNullOrWhiteSpace(i.ZoneId) && !string.IsNullOrWhiteSpace(i.TargetItemKey))
                            .Select(i => $"{i.TargetItemKey}|{i.ZoneId}"),
                        StringComparer.OrdinalIgnoreCase);

                    for (var i = 0; i < intersections.Count; i++)
                    {
                        var inter = intersections[i];
                        if (inter == null || string.IsNullOrWhiteSpace(inter.ZoneId) || string.IsNullOrWhiteSpace(inter.TargetItemKey))
                        {
                            continue;
                        }

                        var key = $"{inter.TargetItemKey}|{inter.ZoneId}";
                        if (!baselineSet.Contains(key))
                        {
                            inter.IsOffsetOnly = true;
                        }
                    }
                }
                result.Intersections = intersections.ToList();
                PopulateRunDiagnostics(result, dataset, cacheToUse);

                stats.ZonesProcessed = dataset.Zones.Count;
                stats.TargetsProcessed = cacheToUse?.TargetKeys?.Length ?? dataset.TargetsForEngine.Count;
                stats.ZonesWithMesh = dataset.ZonesWithMesh;
                stats.ZonesMeshFallback = dataset.ZonesMeshFallback;
                stats.MeshExtractionErrors = dataset.MeshExtractionErrors;
                stats.ModeUsed = engine.Mode.ToString();
                stats.PresetUsed = diagnostics.PresetUsed;
                stats.TraversalUsed = diagnostics.TraversalUsed;
                stats.CandidatePairs = diagnostics.CandidatePairs;
                stats.AvgCandidatesPerZone = diagnostics.AvgCandidatesPerZone;
                stats.MaxCandidatesPerZone = diagnostics.MaxCandidatesPerZone;
                stats.AvgCandidatesPerTarget = diagnostics.AvgCandidatesPerTarget;
                stats.MaxCandidatesPerTarget = diagnostics.MaxCandidatesPerTarget;
                stats.CandidateTargetStatsAvailable = diagnostics.CandidateTargetStatsAvailable;
                stats.TargetsTotal = diagnostics.TargetsTotal;
                stats.TargetsWithBounds = diagnostics.TargetsWithBounds;
                stats.TargetsWithoutBounds = diagnostics.TargetsWithoutBounds;
                stats.TargetsSampled = diagnostics.TargetsSampled;
                stats.TargetsSampleSkippedNoBounds = diagnostics.TargetsSampleSkippedNoBounds;
                stats.TargetsSampleSkippedNoGeometry = diagnostics.TargetsSampleSkippedNoGeometry;
                stats.MeshPointTests = diagnostics.MeshPointTests;
                stats.BoundsPointTests = diagnostics.BoundsPointTests;
                stats.MeshFallbackPointTests = diagnostics.MeshFallbackPointTests;
                stats.GpuBackend = diagnostics.GpuBackend;
                stats.GpuInitFailureReason = diagnostics.GpuInitFailureReason;
                stats.GpuZonesProcessed = diagnostics.GpuZonesProcessed;
                stats.GpuPointsTested = diagnostics.GpuPointsTested;
                stats.GpuTrianglesTested = diagnostics.GpuTrianglesTested;
                stats.GpuDispatchTime = diagnostics.GpuDispatchTime;
                stats.GpuReadbackTime = diagnostics.GpuReadbackTime;
                stats.GpuAdapterName = diagnostics.GpuAdapterName;
                stats.GpuAdapterLuid = diagnostics.GpuAdapterLuid;
                stats.GpuVendorId = diagnostics.GpuVendorId;
                stats.GpuDeviceId = diagnostics.GpuDeviceId;
                stats.GpuSubSysId = diagnostics.GpuSubSysId;
                stats.GpuRevision = diagnostics.GpuRevision;
                stats.GpuDedicatedVideoMemory = diagnostics.GpuDedicatedVideoMemory;
                stats.GpuDedicatedSystemMemory = diagnostics.GpuDedicatedSystemMemory;
                stats.GpuSharedSystemMemory = diagnostics.GpuSharedSystemMemory;
                stats.GpuFeatureLevel = diagnostics.GpuFeatureLevel;
                stats.GpuPointThreshold = diagnostics.GpuPointThreshold;
                stats.GpuSamplePointsPerTarget = diagnostics.GpuSamplePointsPerTarget;
                stats.GpuZonesEligible = diagnostics.GpuZonesEligible;
                stats.GpuZonesSkippedNoMesh = diagnostics.GpuZonesSkippedNoMesh;
                stats.GpuZonesSkippedMissingTriangles = diagnostics.GpuZonesSkippedMissingTriangles;
                stats.GpuZonesSkippedOpenMesh = diagnostics.GpuZonesSkippedOpenMesh;
                stats.GpuZonesSkippedLowPoints = diagnostics.GpuZonesSkippedLowPoints;
                stats.GpuUncertainPoints = diagnostics.GpuUncertainPoints;
                stats.GpuMaxTrianglesPerZone = diagnostics.GpuMaxTrianglesPerZone;
                stats.GpuMaxPointsPerZone = diagnostics.GpuMaxPointsPerZone;
                stats.GpuOpenMeshZonesEligible = diagnostics.GpuOpenMeshZonesEligible;
                stats.GpuOpenMeshZonesProcessed = diagnostics.GpuOpenMeshZonesProcessed;
                stats.GpuOpenMeshBoundaryEdgeLimit = diagnostics.GpuOpenMeshBoundaryEdgeLimit;
                stats.GpuOpenMeshNonManifoldEdgeLimit = diagnostics.GpuOpenMeshNonManifoldEdgeLimit;
                stats.GpuOpenMeshOutsideTolerance = diagnostics.GpuOpenMeshOutsideTolerance;
                stats.GpuOpenMeshNudge = diagnostics.GpuOpenMeshNudge;
                stats.GpuBatchDispatchCount = diagnostics.GpuBatchDispatchCount;
                stats.GpuBatchMaxZones = diagnostics.GpuBatchMaxZones;
                stats.GpuBatchMaxPoints = diagnostics.GpuBatchMaxPoints;
                stats.GpuBatchMaxTriangles = diagnostics.GpuBatchMaxTriangles;
                stats.GpuBatchAvgZonesPerDispatch = diagnostics.GpuBatchAvgZonesPerDispatch;
                stats.GpuZoneDiagnostics = diagnostics.GpuZoneDiagnostics ?? new List<SpaceMapperGpuZoneDiagnostic>();
                stats.SlowZoneThresholdSeconds = diagnostics.SlowZoneThresholdSeconds;
                stats.SlowZones = diagnostics.SlowZones ?? new List<SpaceMapperSlowZoneInfo>();
                stats.UsedPreflightIndex = diagnostics.UsedPreflightIndex;
                stats.BuildIndexTime = diagnostics.BuildIndexTime;
                stats.CandidateQueryTime = diagnostics.CandidateQueryTime;
                stats.NarrowPhaseTime = diagnostics.NarrowPhaseTime;

                runProgress?.SetStage(SpaceMapperRunStage.ResolvingAssignments, "Resolving assignments...");

                var writeback = ExecuteWriteback(
                    dataset,
                    request.ProcessingSettings,
                    request.Mappings,
                    session,
                    intersections,
                    writeToModel: true,
                    token,
                    runProgress);

                stats.ContainedTagged = writeback.ContainedTagged;
                stats.PartialTagged = writeback.PartialTagged;
                stats.MultiZoneTagged = writeback.MultiZoneTagged;
                stats.Skipped = writeback.Skipped;
                stats.SkippedUnchanged = writeback.SkippedUnchanged;
                stats.WritesPerformed = writeback.WritesPerformed;
                stats.WritebackTargetsWritten = writeback.TargetsWritten;
                stats.WritebackCategoriesWritten = writeback.CategoriesWritten;
                stats.WritebackPropertiesWritten = writeback.PropertiesWritten;
                stats.AvgMsPerCategoryWrite = writeback.AvgMsPerCategoryWrite;
                stats.AvgMsPerTargetWrite = writeback.AvgMsPerTargetWrite;
                stats.WritebackStrategy = writeback.StrategyUsed;
                stats.WriteBackTime = writeback.Elapsed;
                stats.ZoneSummaries = writeback.ZoneSummaries;
                if (writeback.StrategyUsed == SpaceMapperWritebackStrategy.LegacyPerMapping)
                {
                    _log?.Invoke("SpaceMapper warning: Legacy per-mapping writeback enabled (slower).");
                }

                result.Stats = stats;
                sw.Stop();
                result.Stats.Elapsed = sw.Elapsed;
                result.Message = $"Space Mapper: {stats.ZonesProcessed} zones, {stats.TargetsProcessed} targets, {stats.ContainedTagged} contained, {stats.PartialTagged} partial. Mode={stats.ModeUsed}.";
                _log?.Invoke(result.Message);
                _log?.Invoke($"SpaceMapper Timings: resolve {stats.ResolveTime.TotalMilliseconds:0}ms, build {stats.BuildGeometryTime.TotalMilliseconds:0}ms, index {stats.BuildIndexTime.TotalMilliseconds:0}ms, candidates {stats.CandidateQueryTime.TotalMilliseconds:0}ms, narrow {stats.NarrowPhaseTime.TotalMilliseconds:0}ms, write {stats.WriteBackTime.TotalMilliseconds:0}ms");

                result.ReportPath = SpaceMapperRunReportWriter.TryWriteReport(request, resolved, dataset, result, runProgress);
                if (!string.IsNullOrWhiteSpace(result.ReportPath))
                {
                    _log?.Invoke($"SpaceMapper report saved: {result.ReportPath}");
                }

                runProgress?.SetStage(SpaceMapperRunStage.Finalizing, "Finalizing...");
                runProgress?.MarkCompleted();

                return result;
            }
            catch (OperationCanceledException)
            {
                runProgress?.MarkCancelled();
                if (runProgress != null)
                {
                    stats.ZonesProcessed = Math.Max(stats.ZonesProcessed, runProgress.ZonesProcessed);
                    stats.TargetsProcessed = Math.Max(stats.TargetsProcessed, runProgress.TargetsProcessed);
                    stats.WritebackTargetsWritten = Math.Max(stats.WritebackTargetsWritten, runProgress.WriteTargetsProcessed);
                    stats.CandidatePairs = Math.Max(stats.CandidatePairs, runProgress.CandidatePairs);
                }
                result.Stats = stats;
                sw.Stop();
                result.Stats.Elapsed = sw.Elapsed;
                result.Message = "Space Mapper cancelled.";
                result.ReportPath = SpaceMapperRunReportWriter.TryWriteReport(request, resolved, dataset, result, runProgress);
                if (!string.IsNullOrWhiteSpace(result.ReportPath))
                {
                    _log?.Invoke($"SpaceMapper report saved: {result.ReportPath}");
                }
                throw;
            }
            catch (Exception ex)
            {
                runProgress?.MarkFailed(ex);
                if (runProgress != null)
                {
                    stats.ZonesProcessed = Math.Max(stats.ZonesProcessed, runProgress.ZonesProcessed);
                    stats.TargetsProcessed = Math.Max(stats.TargetsProcessed, runProgress.TargetsProcessed);
                    stats.WritebackTargetsWritten = Math.Max(stats.WritebackTargetsWritten, runProgress.WriteTargetsProcessed);
                    stats.CandidatePairs = Math.Max(stats.CandidatePairs, runProgress.CandidatePairs);
                }
                result.Stats = stats;
                sw.Stop();
                result.Stats.Elapsed = sw.Elapsed;
                result.Message = $"Space Mapper failed: {ex.Message}";
                result.ReportPath = SpaceMapperRunReportWriter.TryWriteReport(request, resolved, dataset, result, runProgress);
                if (!string.IsNullOrWhiteSpace(result.ReportPath))
                {
                    _log?.Invoke($"SpaceMapper report saved: {result.ReportPath}");
                }
                throw;
            }
            finally
            {
                if (liveTimer != null)
                {
                    liveTimer.Dispose();
                    liveTimer = null;
                }

                WriteLiveReportSnapshot();
            }
        }

        internal static SpaceMapperWritebackResult ExecuteWriteback(
            SpaceMapperComputeDataset dataset,
            SpaceMapperProcessingSettings settings,
            IReadOnlyList<SpaceMapperMappingDefinition> mappings,
            ScrapeSession session,
            IEnumerable<ZoneTargetIntersection> intersections,
            bool writeToModel,
            CancellationToken token,
            SpaceMapperRunProgressState runProgress = null)
        {
            var result = new SpaceMapperWritebackResult();
            if (dataset == null || intersections == null)
            {
                return result;
            }

            settings ??= new SpaceMapperProcessingSettings();
            mappings ??= Array.Empty<SpaceMapperMappingDefinition>();

            var strategy = settings.WritebackStrategy;
            var showInternalProperties = settings.ShowInternalPropertiesDuringWriteback;
            var writeToModelEffective = writeToModel && strategy != SpaceMapperWritebackStrategy.VirtualNoBake;
            result.StrategyUsed = strategy;

            var zoneLookup = dataset.Zones.ToDictionary(z => z.ZoneId, z => z);
            var zoneOrderLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < dataset.Zones.Count; i++)
            {
                var zoneId = dataset.Zones[i]?.ZoneId;
                if (!string.IsNullOrWhiteSpace(zoneId) && !zoneOrderLookup.ContainsKey(zoneId))
                {
                    zoneOrderLookup[zoneId] = i;
                }
            }
            var targetLookup = dataset.TargetModels.ToDictionary(t => t.ItemKey, t => t);
            var ruleMembership = BuildRuleMembership(dataset.TargetsByRule);
            var zoneValueLookup = new ZoneValueLookup(session);
            var summaryLookup = new Dictionary<string, ZoneSummary>(StringComparer.OrdinalIgnoreCase);

            var intersectionList = intersections as IList<ZoneTargetIntersection> ?? intersections.ToList();
            var groupedTargets = intersectionList
                .GroupBy(i => i.TargetItemKey, StringComparer.OrdinalIgnoreCase)
                .ToList();

            runProgress?.SetStage(SpaceMapperRunStage.WritingProperties, "Writing properties...");
            runProgress?.SetWriteTotals(groupedTargets.Count);

            var sw = Stopwatch.StartNew();
            foreach (var group in groupedTargets)
            {
                token.ThrowIfCancellationRequested();

                try
                {
                if (!targetLookup.TryGetValue(group.Key, out var tgt)) continue;
                var rulesForTarget = ruleMembership.TryGetValue(group.Key, out var list) ? list : new List<SpaceMapperTargetRule>();
                if (!rulesForTarget.Any()) continue;

                var relevant = FilterByMembership(group, rulesForTarget, settings).ToList();
                if (relevant.Count == 0)
                {
                    result.Skipped++;
                    continue;
                }

                if (!settings.EnableMultipleZones)
                {
                    var best = SelectBestIntersection(relevant, settings, zoneLookup, zoneOrderLookup, tgt);
                    relevant = best != null ? new List<ZoneTargetIntersection> { best } : new List<ZoneTargetIntersection>();
                }

                var isLegacy = strategy == SpaceMapperWritebackStrategy.LegacyPerMapping;
                var targetCategories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var targetPropertiesWritten = 0;
                var useSignature = settings.SkipUnchangedWriteback;
                var usePacking = settings.PackWritebackProperties;

                PropertyWriter.PropertyWriterSession writer = null;
                if (writeToModelEffective && !isLegacy)
                {
                    writer = new PropertyWriter.PropertyWriterSession(tgt.ModelItem, showInternalProperties);
                }

                bool RegisterWrite(string categoryName, bool wrote)
                {
                    if (!wrote)
                    {
                        return false;
                    }

                    result.WritesPerformed++;
                    result.PropertiesWritten++;
                    targetPropertiesWritten++;

                    if (isLegacy)
                    {
                        result.CategoriesWritten++;
                    }
                    else if (!string.IsNullOrWhiteSpace(categoryName))
                    {
                        targetCategories.Add(categoryName);
                    }

                    return true;
                }

                var isMultiZone = relevant.Count > 1;
                if (isMultiZone) result.MultiZoneTagged++;

                UpdateAssignmentStats(relevant, zoneLookup, summaryLookup, result);

                var writeEntries = BuildWritebackEntries(relevant, mappings, settings, zoneLookup, zoneValueLookup);
                if (writeEntries.Count == 0)
                {
                    result.Skipped++;
                    continue;
                }

                List<ResolvedWritebackEntry> resolvedEntries = null;
                if (useSignature || usePacking)
                {
                    resolvedEntries = ResolveWritebackEntries(tgt.ModelItem, writeEntries);
                }

                string signatureValue = null;
                string signatureCategory = null;
                if (useSignature)
                {
                    var signatureSource = resolvedEntries ?? ResolveWritebackEntries(tgt.ModelItem, writeEntries);
                    signatureCategory = ResolveSignatureCategory(signatureSource, settings);
                    signatureValue = BuildWritebackSignature(signatureSource, usePacking);
                    var existingSignature = ReadPropertyValue(tgt.ModelItem, signatureCategory, WritebackSignaturePropertyName);
                    if (string.Equals(existingSignature, signatureValue, StringComparison.Ordinal))
                    {
                        result.SkippedUnchanged++;
                        continue;
                    }
                }

                if (resolvedEntries != null && resolvedEntries.Count == 0)
                {
                    result.Skipped++;
                    continue;
                }

                void WriteResolvedEntries(IReadOnlyList<ResolvedWritebackEntry> entriesToWrite)
                {
                    foreach (var entry in entriesToWrite)
                    {
                        if (entry == null)
                        {
                            continue;
                        }

                        var wrote = false;
                        if (writeToModelEffective)
                        {
                            wrote = isLegacy
                                ? PropertyWriter.WriteProperty(tgt.ModelItem, entry.CategoryName, entry.PropertyName, entry.Value, WriteMode.Overwrite, string.Empty, showInternalProperties)
                                : writer?.WriteProperty(entry.CategoryName, entry.PropertyName, entry.Value, WriteMode.Overwrite, string.Empty) == true;
                        }
                        else
                        {
                            wrote = true;
                        }

                        RegisterWrite(entry.CategoryName, wrote);
                    }
                }

                if (usePacking)
                {
                    var sourceEntries = resolvedEntries ?? ResolveWritebackEntries(tgt.ModelItem, writeEntries);
                    var packedValue = BuildPackedValue(sourceEntries);
                    var packedEntries = new List<ResolvedWritebackEntry>();
                    if (!string.IsNullOrWhiteSpace(packedValue))
                    {
                        var packedCategory = ResolvePackedCategory(sourceEntries, settings);
                        packedEntries.Add(new ResolvedWritebackEntry(packedCategory, PackedWritebackPropertyName, packedValue));
                    }

                    if (useSignature && !string.IsNullOrWhiteSpace(signatureValue))
                    {
                        var sigCategory = signatureCategory ?? ResolveSignatureCategory(sourceEntries, settings);
                        packedEntries.Add(new ResolvedWritebackEntry(sigCategory, WritebackSignaturePropertyName, signatureValue));
                    }

                    if (packedEntries.Count == 0)
                    {
                        result.Skipped++;
                        continue;
                    }

                    WriteResolvedEntries(packedEntries);
                }
                else if (resolvedEntries != null)
                {
                    var entriesToWrite = resolvedEntries;
                    if (useSignature && !string.IsNullOrWhiteSpace(signatureValue))
                    {
                        var sigCategory = signatureCategory ?? ResolveSignatureCategory(resolvedEntries, settings);
                        entriesToWrite = new List<ResolvedWritebackEntry>(resolvedEntries)
                        {
                            new ResolvedWritebackEntry(sigCategory, WritebackSignaturePropertyName, signatureValue)
                        };
                    }

                    WriteResolvedEntries(entriesToWrite);
                }
                else
                {
                    foreach (var entry in writeEntries)
                    {
                        var wrote = false;
                        if (writeToModelEffective)
                        {
                            wrote = isLegacy
                                ? PropertyWriter.WriteProperty(tgt.ModelItem, entry.CategoryName, entry.PropertyName, entry.Value, entry.Mode, entry.AppendSeparator, showInternalProperties)
                                : writer?.WriteProperty(entry.CategoryName, entry.PropertyName, entry.Value, entry.Mode, entry.AppendSeparator) == true;
                        }
                        else
                        {
                            wrote = true;
                        }

                        RegisterWrite(entry.CategoryName, wrote);
                    }
                }

                writer?.Commit();

                if (!isLegacy && targetCategories.Count > 0)
                {
                    result.CategoriesWritten += targetCategories.Count;
                }

                if (targetPropertiesWritten > 0)
                {
                    result.TargetsWritten++;
                }
                }
                finally
                {
                    if (runProgress != null)
                    {
                        runProgress.IncrementWriteProcessed();
                    }
                }
            }

            sw.Stop();
            result.Elapsed = sw.Elapsed;
            if (result.CategoriesWritten > 0)
            {
                result.AvgMsPerCategoryWrite = result.Elapsed.TotalMilliseconds / result.CategoriesWritten;
            }
            if (result.TargetsWritten > 0)
            {
                result.AvgMsPerTargetWrite = result.Elapsed.TotalMilliseconds / result.TargetsWritten;
            }
            result.ZoneSummaries = summaryLookup.Values.ToList();
            return result;
        }

        private sealed class WritebackEntry
        {
            public string CategoryName { get; set; }
            public string PropertyName { get; set; }
            public string Value { get; set; }
            public WriteMode Mode { get; set; }
            public string AppendSeparator { get; set; }
        }

        private sealed class ResolvedWritebackEntry
        {
            public ResolvedWritebackEntry(string categoryName, string propertyName, string value)
            {
                CategoryName = categoryName;
                PropertyName = propertyName;
                Value = value;
            }

            public string CategoryName { get; }
            public string PropertyName { get; }
            public string Value { get; }
        }

        private static string SequencedPropertyName(string baseName, int index)
        {
            return index <= 0 ? baseName : $"{baseName}({index})";
        }

        private static bool ShouldSequenceBehaviour(
            IReadOnlyList<ZoneTargetIntersection> relevant,
            IEnumerable<SpaceMapperMappingDefinition> mappings,
            SpaceMapperProcessingSettings settings)
        {
            if (settings?.EnableMultipleZones != true) return false;
            if (relevant == null || relevant.Count <= 1) return false;

            return mappings != null && mappings.Any(m => m != null && m.MultiZoneCombineMode == MultiZoneCombineMode.Sequence);
        }

        private static string FormatPercent(double? fraction)
        {
            if (!fraction.HasValue)
            {
                return null;
            }

            var value = fraction.Value;
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return null;
            }

            if (value < 0) value = 0;
            if (value > 1) value = 1;

            return (value * 100.0).ToString("0.##", CultureInfo.InvariantCulture);
        }

        private static List<WritebackEntry> BuildWritebackEntries(
            IReadOnlyList<ZoneTargetIntersection> relevant,
            IReadOnlyList<SpaceMapperMappingDefinition> mappings,
            SpaceMapperProcessingSettings settings,
            IReadOnlyDictionary<string, ZoneGeometry> zoneLookup,
            ZoneValueLookup zoneValueLookup)
        {
            var entries = new List<WritebackEntry>();

            if (settings?.WriteZoneBehaviorProperty == true)
            {
                var behaviorCategory = settings.ZoneBehaviorCategory;
                var behaviorProperty = settings.ZoneBehaviorPropertyName;
                var containedValue = settings.ZoneBehaviorContainedValue;
                var partialValue = settings.ZoneBehaviorPartialValue;

                if (!string.IsNullOrWhiteSpace(behaviorCategory)
                    && !string.IsNullOrWhiteSpace(behaviorProperty))
                {
                    if (ShouldSequenceBehaviour(relevant, mappings, settings))
                    {
                        var count = Math.Min(relevant.Count, MaxSequenceOutputsPerMapping);
                        for (var i = 0; i < count; i++)
                        {
                            var inter = relevant[i];
                            var behaviorValue = IsPartialForBehavior(inter, settings) ? partialValue : containedValue;
                            if (string.IsNullOrWhiteSpace(behaviorValue))
                            {
                                continue;
                            }

                            entries.Add(new WritebackEntry
                            {
                                CategoryName = behaviorCategory,
                                PropertyName = SequencedPropertyName(behaviorProperty, i),
                                Value = behaviorValue,
                                Mode = WriteMode.Overwrite,
                                AppendSeparator = string.Empty
                            });
                        }
                    }
                    else
                    {
                        var hasPartial = relevant.Any(r => IsPartialForBehavior(r, settings));
                        var behaviorValue = hasPartial ? partialValue : containedValue;
                        if (!string.IsNullOrWhiteSpace(behaviorValue))
                        {
                            entries.Add(new WritebackEntry
                            {
                                CategoryName = behaviorCategory,
                                PropertyName = behaviorProperty,
                                Value = behaviorValue,
                                Mode = WriteMode.Overwrite,
                                AppendSeparator = string.Empty
                            });
                        }
                    }
                }
            }

            if (settings?.WriteZoneContainmentPercentProperty == true && relevant != null && relevant.Count > 0)
            {
                var category = string.IsNullOrWhiteSpace(settings.ZoneBehaviorCategory)
                    ? "ME_SpaceInfo"
                    : settings.ZoneBehaviorCategory;

                if (ShouldSequenceBehaviour(relevant, mappings, settings))
                {
                    var count = Math.Min(relevant.Count, MaxSequenceOutputsPerMapping);
                    for (var i = 0; i < count; i++)
                    {
                        var value = FormatPercent(relevant[i]?.ContainmentFraction);
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            continue;
                        }

                        entries.Add(new WritebackEntry
                        {
                            CategoryName = category,
                            PropertyName = SequencedPropertyName(ZoneContainmentPercentPropertyName, i),
                            Value = value,
                            Mode = WriteMode.Overwrite,
                            AppendSeparator = string.Empty
                        });
                    }
                }
                else
                {
                    var values = relevant
                        .Select(r => FormatPercent(r?.ContainmentFraction))
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .ToList();

                    if (values.Count > 0)
                    {
                        var combined = values.Count == 1 ? values[0] : string.Join(", ", values);
                        entries.Add(new WritebackEntry
                        {
                            CategoryName = category,
                            PropertyName = ZoneContainmentPercentPropertyName,
                            Value = combined,
                            Mode = WriteMode.Overwrite,
                            AppendSeparator = string.Empty
                        });
                    }
                }
            }

            if (settings?.WriteZoneOffsetMatchProperty == true && relevant != null && relevant.Count > 0)
            {
                var category = string.IsNullOrWhiteSpace(settings.ZoneBehaviorCategory)
                    ? "ME_SpaceInfo"
                    : settings.ZoneBehaviorCategory;

                if (ShouldSequenceBehaviour(relevant, mappings, settings))
                {
                    var count = Math.Min(relevant.Count, MaxSequenceOutputsPerMapping);
                    for (var i = 0; i < count; i++)
                    {
                        var value = relevant[i]?.IsOffsetOnly == true ? "OffsetOnly" : "Core";
                        entries.Add(new WritebackEntry
                        {
                            CategoryName = category,
                            PropertyName = SequencedPropertyName(ZoneOffsetMatchPropertyName, i),
                            Value = value,
                            Mode = WriteMode.Overwrite,
                            AppendSeparator = string.Empty
                        });
                    }
                }
                else
                {
                    var hasOffsetOnly = relevant.Any(r => r?.IsOffsetOnly == true);
                    var value = hasOffsetOnly ? "OffsetOnly" : "Core";
                    entries.Add(new WritebackEntry
                    {
                        CategoryName = category,
                        PropertyName = ZoneOffsetMatchPropertyName,
                        Value = value,
                        Mode = WriteMode.Overwrite,
                        AppendSeparator = string.Empty
                    });
                }
            }

            if (mappings == null || mappings.Count == 0)
            {
                return entries;
            }

            foreach (var mapping in mappings)
            {
                if (mapping == null
                    || string.IsNullOrWhiteSpace(mapping.TargetCategory)
                    || string.IsNullOrWhiteSpace(mapping.TargetPropertyName))
                {
                    continue;
                }

                var values = new List<string>();
                foreach (var inter in relevant)
                {
                    if (inter == null || string.IsNullOrWhiteSpace(inter.ZoneId))
                    {
                        continue;
                    }

                    if (zoneLookup != null && !zoneLookup.ContainsKey(inter.ZoneId))
                    {
                        continue;
                    }

                    var val = zoneValueLookup?.GetValue(inter.ZoneId, mapping.ZoneCategory, mapping.ZonePropertyName);
                    if (!string.IsNullOrWhiteSpace(val))
                    {
                        values.Add(val);
                    }
                }

                if (values.Count == 0)
                {
                    continue;
                }

                if (mapping.MultiZoneCombineMode == MultiZoneCombineMode.Sequence
                    && settings?.EnableMultipleZones == true
                    && relevant.Count > 1)
                {
                    var count = Math.Min(values.Count, MaxSequenceOutputsPerMapping);
                    for (var i = 0; i < count; i++)
                    {
                        var value = values[i] ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            continue;
                        }

                        entries.Add(new WritebackEntry
                        {
                            CategoryName = mapping.TargetCategory,
                            PropertyName = SequencedPropertyName(mapping.TargetPropertyName, i),
                            Value = value,
                            Mode = mapping.WriteMode,
                            AppendSeparator = mapping.AppendSeparator
                        });
                    }

                    continue;
                }

                var combined = CombineValues(values, mapping.MultiZoneCombineMode, mapping.AppendSeparator);
                if (string.IsNullOrWhiteSpace(combined))
                {
                    continue;
                }

                entries.Add(new WritebackEntry
                {
                    CategoryName = mapping.TargetCategory,
                    PropertyName = mapping.TargetPropertyName,
                    Value = combined,
                    Mode = mapping.WriteMode,
                    AppendSeparator = mapping.AppendSeparator
                });
            }

            return entries;
        }

        private static bool IsPartialForBehavior(ZoneTargetIntersection intersection, SpaceMapperProcessingSettings settings)
        {
            if (intersection == null)
            {
                return false;
            }

            var mode = settings?.ContainmentCalculationMode ?? SpaceMapperContainmentCalculationMode.Auto;
            if (mode == SpaceMapperContainmentCalculationMode.Auto)
            {
                return intersection.IsPartial;
            }

            if (intersection.ContainmentFraction.HasValue)
            {
                return intersection.ContainmentFraction.Value < 1.0 - ContainmentFullTolerance;
            }

            return intersection.IsPartial;
        }

        private static List<ResolvedWritebackEntry> ResolveWritebackEntries(ModelItem item, IReadOnlyList<WritebackEntry> entries)
        {
            var resolved = new Dictionary<string, ResolvedWritebackEntry>(StringComparer.OrdinalIgnoreCase);
            if (item == null || entries == null || entries.Count == 0)
            {
                return new List<ResolvedWritebackEntry>();
            }

            foreach (var entry in entries)
            {
                if (entry == null)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(entry.CategoryName)
                    || string.IsNullOrWhiteSpace(entry.PropertyName))
                {
                    continue;
                }

                if (!TryResolveWriteValue(item, entry, out var finalValue))
                {
                    continue;
                }

                var key = $"{entry.CategoryName}|{entry.PropertyName}";
                resolved[key] = new ResolvedWritebackEntry(entry.CategoryName, entry.PropertyName, finalValue);
            }

            return resolved.Values.ToList();
        }

        private static bool TryResolveWriteValue(ModelItem item, WritebackEntry entry, out string finalValue)
        {
            finalValue = entry?.Value ?? string.Empty;
            if (entry == null || string.IsNullOrWhiteSpace(finalValue))
            {
                return false;
            }

            var category = entry.CategoryName;
            var property = entry.PropertyName;
            switch (entry.Mode)
            {
                case WriteMode.Append:
                {
                    var existing = ReadPropertyValue(item, category, property);
                    if (!string.IsNullOrWhiteSpace(existing))
                    {
                        var sep = entry.AppendSeparator ?? string.Empty;
                        finalValue = string.IsNullOrWhiteSpace(sep)
                            ? $"{existing},{finalValue}"
                            : $"{existing}{sep}{finalValue}";
                    }
                    return !string.IsNullOrWhiteSpace(finalValue);
                }
                case WriteMode.OnlyIfBlank:
                {
                    var existing = ReadPropertyValue(item, category, property);
                    return string.IsNullOrWhiteSpace(existing);
                }
                default:
                    return true;
            }
        }

        private static string ResolvePackedCategory(IReadOnlyList<ResolvedWritebackEntry> entries, SpaceMapperProcessingSettings settings)
        {
            var category = GetDominantCategory(entries);
            if (!string.IsNullOrWhiteSpace(category))
            {
                return category;
            }

            if (!string.IsNullOrWhiteSpace(settings?.ZoneBehaviorCategory))
            {
                return settings.ZoneBehaviorCategory;
            }

            return "ME_SpaceInfo";
        }

        private static string ResolveSignatureCategory(IReadOnlyList<ResolvedWritebackEntry> entries, SpaceMapperProcessingSettings settings)
        {
            var category = GetDominantCategory(entries);
            if (!string.IsNullOrWhiteSpace(category))
            {
                return category;
            }

            if (!string.IsNullOrWhiteSpace(settings?.ZoneBehaviorCategory))
            {
                return settings.ZoneBehaviorCategory;
            }

            return "ME_SpaceInfo";
        }

        private static string GetDominantCategory(IReadOnlyList<ResolvedWritebackEntry> entries)
        {
            if (entries == null || entries.Count == 0)
            {
                return null;
            }

            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in entries)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.CategoryName))
                {
                    continue;
                }

                if (!counts.ContainsKey(entry.CategoryName))
                {
                    counts[entry.CategoryName] = 0;
                }
                counts[entry.CategoryName]++;
            }

            if (counts.Count == 0)
            {
                return null;
            }

            return counts
                .OrderByDescending(kvp => kvp.Value)
                .ThenBy(kvp => kvp.Key, StringComparer.OrdinalIgnoreCase)
                .Select(kvp => kvp.Key)
                .FirstOrDefault();
        }

        private static string BuildPackedValue(IReadOnlyList<ResolvedWritebackEntry> entries)
        {
            if (entries == null || entries.Count == 0)
            {
                return string.Empty;
            }

            var ordered = entries
                .Where(e => e != null
                    && !string.IsNullOrWhiteSpace(e.CategoryName)
                    && !string.IsNullOrWhiteSpace(e.PropertyName)
                    && !string.IsNullOrWhiteSpace(e.Value))
                .OrderBy(e => e.CategoryName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(e => e.PropertyName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (ordered.Count == 0)
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            foreach (var entry in ordered)
            {
                if (sb.Length > 0)
                {
                    sb.Append(" | ");
                }

                sb.Append(entry.CategoryName)
                    .Append('.')
                    .Append(entry.PropertyName)
                    .Append('=')
                    .Append(EscapePackedValue(entry.Value));
            }

            return sb.ToString();
        }

        private static string BuildWritebackSignature(IReadOnlyList<ResolvedWritebackEntry> entries, bool packed)
        {
            var sb = new StringBuilder();
            sb.Append("v1|pack=").Append(packed ? '1' : '0').Append('\n');

            if (entries != null && entries.Count > 0)
            {
                var ordered = entries
                    .Where(e => e != null
                        && !string.IsNullOrWhiteSpace(e.CategoryName)
                        && !string.IsNullOrWhiteSpace(e.PropertyName))
                    .OrderBy(e => e.CategoryName, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(e => e.PropertyName, StringComparer.OrdinalIgnoreCase);

                foreach (var entry in ordered)
                {
                    sb.Append(entry.CategoryName)
                        .Append('|')
                        .Append(entry.PropertyName)
                        .Append('|')
                        .Append(entry.Value ?? string.Empty)
                        .Append('\n');
                }
            }

            using var sha = SHA256.Create();
            var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));
            return ToHexString(hash);
        }

        private static string EscapePackedValue(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value.Replace("\r", string.Empty).Replace("\n", "\\n");
        }

        private static string ToHexString(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
            {
                return string.Empty;
            }

            var sb = new StringBuilder(bytes.Length * 2);
            foreach (var b in bytes)
            {
                sb.Append(b.ToString("x2"));
            }
            return sb.ToString();
        }

        private static string ReadPropertyValue(ModelItem item, string categoryName, string propertyName)
        {
            if (item == null || string.IsNullOrWhiteSpace(categoryName) || string.IsNullOrWhiteSpace(propertyName))
            {
                return string.Empty;
            }

            try
            {
                foreach (var cat in item.PropertyCategories)
                {
                    if (!string.Equals(cat.DisplayName ?? cat.Name, categoryName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

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
                // ignore read failures
            }

            return string.Empty;
        }

        private static void UpdateAssignmentStats(
            IEnumerable<ZoneTargetIntersection> relevant,
            IReadOnlyDictionary<string, ZoneGeometry> zoneLookup,
            IDictionary<string, ZoneSummary> summaryLookup,
            SpaceMapperWritebackResult result)
        {
            if (relevant == null || result == null || summaryLookup == null)
            {
                return;
            }

            foreach (var inter in relevant)
            {
                if (inter == null)
                {
                    continue;
                }

                if (inter.IsContained) result.ContainedTagged++;
                if (inter.IsPartial) result.PartialTagged++;
                if (zoneLookup != null && zoneLookup.TryGetValue(inter.ZoneId, out var z))
                {
                    if (!summaryLookup.TryGetValue(z.ZoneId, out var summary))
                    {
                        summary = new ZoneSummary { ZoneId = z.ZoneId, ZoneName = z.DisplayName };
                        summaryLookup[z.ZoneId] = summary;
                    }
                    if (inter.IsContained) summary.ContainedCount++;
                    if (inter.IsPartial) summary.PartialCount++;
                }
            }
        }

        private static SpaceMapperProcessingSettings CloneProcessingSettings(SpaceMapperProcessingSettings settings)
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
                PackWritebackProperties = settings.PackWritebackProperties,
                ZoneBoundsMode = settings.ZoneBoundsMode,
                ZoneKDopVariant = settings.ZoneKDopVariant,
                TargetBoundsMode = settings.TargetBoundsMode,
                TargetKDopVariant = settings.TargetKDopVariant,
                TargetMidpointMode = settings.TargetMidpointMode,
                ZoneContainmentEngine = settings.ZoneContainmentEngine,
                ZoneResolutionStrategy = settings.ZoneResolutionStrategy,
                ExcludeZonesFromTargets = settings.ExcludeZonesFromTargets,
                WriteZoneBehaviorProperty = settings.WriteZoneBehaviorProperty,
                WriteZoneContainmentPercentProperty = settings.WriteZoneContainmentPercentProperty,
                ContainmentCalculationMode = settings.ContainmentCalculationMode,
                GpuRayCount = settings.GpuRayCount,
                DockPaneCloseDelaySeconds = settings.DockPaneCloseDelaySeconds,
                EnableZoneOffsets = settings.EnableZoneOffsets,
                EnableOffsetAreaPass = settings.EnableOffsetAreaPass,
                WriteZoneOffsetMatchProperty = settings.WriteZoneOffsetMatchProperty
            };
        }

        internal static SpaceMapperResolvedData ResolveData(SpaceMapperRequest request, Document doc, ScrapeSession session)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            var sw = Stopwatch.StartNew();
            var zoneModels = ResolveZones(session, request.ZoneSource, request.ZoneSetName, doc).ToList();
            var targetsByRule = new Dictionary<string, List<SpaceMapperTargetRule>>(StringComparer.OrdinalIgnoreCase);
            var targetModels = ResolveTargets(doc, request.TargetRules, targetsByRule).ToList();
            var settings = request.ProcessingSettings;
            if (settings?.ExcludeZonesFromTargets == true && zoneModels.Count > 0 && targetModels.Count > 0)
            {
                var zoneKeys = new HashSet<string>(zoneModels.Select(z => z.ItemKey), StringComparer.OrdinalIgnoreCase);
                if (zoneKeys.Count > 0)
                {
                    targetModels = targetModels.Where(t => !zoneKeys.Contains(t.ItemKey)).ToList();
                    foreach (var key in zoneKeys)
                    {
                        targetsByRule.Remove(key);
                    }
                }
            }
            sw.Stop();

            return new SpaceMapperResolvedData
            {
                ZoneModels = zoneModels,
                TargetModels = targetModels,
                TargetsByRule = targetsByRule,
                ResolveTime = sw.Elapsed
            };
        }

        internal static SpaceMapperComputeDataset BuildGeometryData(
            SpaceMapperResolvedData resolved,
            SpaceMapperProcessingSettings settings,
            bool buildTargetGeometry,
            CancellationToken token,
            SpaceMapperRunProgressState runProgress = null,
            bool? forceRequirePlanes = null)
        {
            if (resolved == null) throw new ArgumentNullException(nameof(resolved));

            var sw = Stopwatch.StartNew();
            var zones = new List<ZoneGeometry>(resolved.ZoneModels.Count);
            var requirePlanes = forceRequirePlanes ?? false;
            var useMesh = settings?.ZoneContainmentEngine == SpaceMapperZoneContainmentEngine.MeshAccurate;
            var zonesWithMesh = 0;
            var zonesMeshFallback = 0;
            var meshExtractionErrors = 0;
            var requireTargetTriangles = settings != null
                && (settings.ContainmentCalculationMode == SpaceMapperContainmentCalculationMode.TargetGeometry
                    || settings.ContainmentCalculationMode == SpaceMapperContainmentCalculationMode.TargetGeometryGpu)
                && (settings.WriteZoneContainmentPercentProperty || settings.WriteZoneBehaviorProperty);
            var zoneTotal = resolved.ZoneModels.Count;
            var targetTotal = resolved.TargetModels.Count;

            runProgress?.UpdateDetail($"Extracting zones (0/{zoneTotal})...");
            for (int i = 0; i < resolved.ZoneModels.Count; i++)
            {
                token.ThrowIfCancellationRequested();
                var zoneModel = resolved.ZoneModels[i];
                if (runProgress != null)
                {
                    var zoneName = string.IsNullOrWhiteSpace(zoneModel?.DisplayName) ? zoneModel?.ItemKey : zoneModel.DisplayName;
                    runProgress.UpdateDetail($"Extracting zone {i + 1}/{zoneTotal}: {zoneName}");
                }
                var zone = GeometryExtractor.ExtractZoneGeometry(zoneModel.ModelItem, zoneModel.ItemKey, zoneModel.DisplayName, settings);
                var hasBounds = zone?.BoundingBox != null;
                var hasPlanes = !requirePlanes || (zone?.Planes != null && zone.Planes.Count > 0);

                if (hasBounds && hasPlanes)
                {
                    zones.Add(zone);
                    if (zone.HasTriangleMesh)
                    {
                        zonesWithMesh++;
                    }
                    else if (useMesh)
                    {
                        zonesMeshFallback++;
                    }

                    if (zone.MeshExtractionFailed)
                    {
                        meshExtractionErrors++;
                    }
                }

                if (runProgress != null && ((i & 15) == 0 || i == zoneTotal - 1))
                {
                    runProgress.UpdateZonesProcessed(i + 1);
                    runProgress.UpdateDetail($"Extracting zones ({i + 1}/{zoneTotal})...");
                }
            }

            var targetsForEngine = new List<TargetGeometry>(resolved.TargetModels.Count);
            if (buildTargetGeometry || requireTargetTriangles == true)
            {
                for (int i = 0; i < resolved.TargetModels.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    var target = resolved.TargetModels[i];
                    if (runProgress != null)
                    {
                        var targetName = string.IsNullOrWhiteSpace(target?.DisplayName) ? target?.ItemKey : target.DisplayName;
                        runProgress.UpdateDetail($"Extracting target {i + 1}/{targetTotal}: {targetName}");
                    }
                    var geom = GeometryExtractor.ExtractTargetGeometry(
                        target.ModelItem,
                        target.ItemKey,
                        target.DisplayName,
                        extractTriangles: requireTargetTriangles == true);
                    if (geom?.BoundingBox != null)
                    {
                        targetsForEngine.Add(geom);
                    }

                    if (runProgress != null && ((i & 31) == 0 || i == targetTotal - 1))
                    {
                        runProgress.UpdateTargetsProcessed(i + 1);
                        runProgress.UpdateDetail($"Extracting targets ({i + 1}/{targetTotal})...");
                    }
                }
            }
            else
            {
                targetsForEngine.AddRange(resolved.TargetModels);
                if (runProgress != null && targetTotal > 0)
                {
                    runProgress.UpdateTargetsProcessed(targetTotal);
                    runProgress.UpdateDetail($"Using cached targets ({targetTotal}/{targetTotal})...");
                }
            }

            sw.Stop();

            return new SpaceMapperComputeDataset
            {
                Zones = zones,
                TargetsForEngine = targetsForEngine,
                ZoneModels = resolved.ZoneModels,
                TargetModels = resolved.TargetModels,
                TargetsByRule = resolved.TargetsByRule,
                ResolveTime = resolved.ResolveTime,
                BuildGeometryTime = sw.Elapsed,
                ZonesWithMesh = zonesWithMesh,
                ZonesMeshFallback = zonesMeshFallback,
                MeshExtractionErrors = meshExtractionErrors
            };
        }

        private static SpaceMapperPreflightCache TryGetReusablePreflightCache(SpaceMapperRequest request, SpaceMapperPreflightCache cache, IReadOnlyList<TargetGeometry> targets)
        {
            if (request == null || cache == null)
            {
                return null;
            }

            var hasTargetGrid = cache.Grid != null
                && cache.TargetBounds != null
                && cache.TargetKeys != null
                && cache.TargetIndices != null;

            var hasZoneGrid = cache.ZoneGrid != null
                && cache.TargetBounds != null
                && cache.TargetKeys != null
                && cache.TargetIndices != null
                && cache.ZoneBoundsInflated != null
                && cache.ZoneIndexMap != null;

            if (!hasTargetGrid && !hasZoneGrid)
            {
                return null;
            }

            var signature = SpaceMapperPreflightService.BuildSignature(request);
            if (!string.Equals(signature, cache.Signature, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var targetKeysWithBounds = targets
                .Where(t => t?.BoundingBox != null)
                .Select(t => t.ItemKey)
                .ToList();

            if (cache.TargetKeys.Length != targetKeysWithBounds.Count)
            {
                return null;
            }

            var keySet = new HashSet<string>(cache.TargetKeys, StringComparer.OrdinalIgnoreCase);
            foreach (var key in targetKeysWithBounds)
            {
                if (!keySet.Contains(key))
                {
                    return null;
                }
            }

            return cache;
        }

        private static void PopulateRunDiagnostics(
            SpaceMapperRunResult result,
            SpaceMapperComputeDataset dataset,
            SpaceMapperPreflightCache cacheToUse)
        {
            if (result == null || dataset == null)
            {
                return;
            }

            var targetLookup = new Dictionary<string, ModelItem>(StringComparer.OrdinalIgnoreCase);
            var missingBounds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var target in dataset.TargetModels)
            {
                if (target?.ModelItem == null || string.IsNullOrWhiteSpace(target.ItemKey))
                {
                    continue;
                }

                if (!targetLookup.ContainsKey(target.ItemKey))
                {
                    targetLookup[target.ItemKey] = target.ModelItem;
                }

                if (!HasTargetBounds(target))
                {
                    missingBounds.Add(target.ItemKey);
                }
            }

            foreach (var key in missingBounds)
            {
                if (targetLookup.TryGetValue(key, out var item))
                {
                    result.TargetsWithoutBounds.Add(item);
                }
            }

            var matchedKeys = new HashSet<string>(
                result.Intersections
                    .Where(i => i != null && !string.IsNullOrWhiteSpace(i.TargetItemKey))
                    .Select(i => i.TargetItemKey),
                StringComparer.OrdinalIgnoreCase);

            var processedKeys = cacheToUse?.TargetKeys
                ?? dataset.TargetsForEngine
                    .Where(t => t != null && !string.IsNullOrWhiteSpace(t.ItemKey))
                    .Select(t => t.ItemKey)
                    .ToArray();

            foreach (var key in processedKeys)
            {
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                if (matchedKeys.Contains(key))
                {
                    continue;
                }

                if (targetLookup.TryGetValue(key, out var item))
                {
                    result.TargetsUnmatched.Add(item);
                }
            }
        }

        private static bool HasTargetBounds(TargetGeometry target)
        {
            if (target?.BoundingBox != null)
            {
                return true;
            }

            return target?.ModelItem?.BoundingBox() != null;
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
                        OverlapVolume = inter.OverlapVolume,
                        ContainmentFraction = inter.ContainmentFraction
                    };
                }
                else if (partial && allowPartial)
                {
                    yield return inter;
                }
            }
        }

        private struct BestZoneCandidate
        {
            public ZoneTargetIntersection Intersection;
            public int ZoneOrder;
            public double Volume;
            public double DistanceSq;
        }

        private static ZoneTargetIntersection SelectBestIntersection(
            IReadOnlyList<ZoneTargetIntersection> intersections,
            SpaceMapperProcessingSettings settings,
            IReadOnlyDictionary<string, ZoneGeometry> zoneLookup,
            IReadOnlyDictionary<string, int> zoneOrderLookup,
            TargetGeometry target)
        {
            if (intersections == null || intersections.Count == 0)
            {
                return null;
            }

            var strategy = settings?.ZoneResolutionStrategy ?? SpaceMapperZoneResolutionStrategy.MostSpecific;
            var containmentEngine = settings?.ZoneContainmentEngine ?? SpaceMapperZoneContainmentEngine.BoundsFast;

            var hasTargetPoint = TryGetTargetPoint(target, settings, containmentEngine, out var tx, out var ty, out var tz);
            BestZoneCandidate best = default;
            var hasBest = false;

            foreach (var inter in intersections)
            {
                if (inter == null || string.IsNullOrWhiteSpace(inter.ZoneId))
                {
                    continue;
                }

                var candidate = new BestZoneCandidate
                {
                    Intersection = inter,
                    ZoneOrder = zoneOrderLookup != null && zoneOrderLookup.TryGetValue(inter.ZoneId, out var order)
                        ? order
                        : int.MaxValue,
                    Volume = double.PositiveInfinity,
                    DistanceSq = double.PositiveInfinity
                };

                if (zoneLookup != null && zoneLookup.TryGetValue(inter.ZoneId, out var zone))
                {
                    var bounds = GetZoneQueryBounds(zone, settings);
                    candidate.Volume = bounds.SizeX * bounds.SizeY * bounds.SizeZ;

                    if (hasTargetPoint)
                    {
                        var cx = (bounds.MinX + bounds.MaxX) * 0.5;
                        var cy = (bounds.MinY + bounds.MaxY) * 0.5;
                        var cz = (bounds.MinZ + bounds.MaxZ) * 0.5;
                        var dx = tx - cx;
                        var dy = ty - cy;
                        var dz = tz - cz;
                        candidate.DistanceSq = (dx * dx) + (dy * dy) + (dz * dz);
                    }
                }

                if (!hasBest || IsBetterCandidate(candidate, best, strategy))
                {
                    best = candidate;
                    hasBest = true;
                }
            }

            return hasBest ? best.Intersection : null;
        }

        private static bool IsBetterCandidate(
            in BestZoneCandidate candidate,
            in BestZoneCandidate current,
            SpaceMapperZoneResolutionStrategy strategy)
        {
            if (candidate.Intersection.IsContained != current.Intersection.IsContained)
            {
                return candidate.Intersection.IsContained;
            }

            switch (strategy)
            {
                case SpaceMapperZoneResolutionStrategy.FirstMatch:
                    return candidate.ZoneOrder < current.ZoneOrder;
                case SpaceMapperZoneResolutionStrategy.LargestOverlap:
                    if (candidate.Intersection.OverlapVolume > current.Intersection.OverlapVolume)
                    {
                        return true;
                    }
                    if (candidate.Intersection.OverlapVolume < current.Intersection.OverlapVolume)
                    {
                        return false;
                    }
                    break;
            }

            if (candidate.Volume < current.Volume)
            {
                return true;
            }

            if (candidate.Volume > current.Volume)
            {
                return false;
            }

            if (candidate.DistanceSq < current.DistanceSq)
            {
                return true;
            }

            if (candidate.DistanceSq > current.DistanceSq)
            {
                return false;
            }

            return candidate.ZoneOrder < current.ZoneOrder;
        }

        private static bool TryGetTargetPoint(
            TargetGeometry target,
            SpaceMapperProcessingSettings settings,
            SpaceMapperZoneContainmentEngine containmentEngine,
            out double x,
            out double y,
            out double z)
        {
            x = 0;
            y = 0;
            z = 0;

            var bbox = target?.BoundingBox;
            if (bbox == null)
            {
                return false;
            }

            _ = containmentEngine;

            var min = bbox.Min;
            var max = bbox.Max;
            var useBottom = settings?.TargetMidpointMode == SpaceMapperMidpointMode.BoundingBoxBottomCenter;
            var useMidpoint = (settings?.TargetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
                || (settings?.UseOriginPointOnly == true);

            x = (min.X + max.X) * 0.5;
            y = (min.Y + max.Y) * 0.5;
            z = useMidpoint && useBottom
                ? min.Z
                : (min.Z + max.Z) * 0.5;

            return true;
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
            if (settings == null || !settings.EnableZoneOffsets)
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
                case MultiZoneCombineMode.Sequence:
                    return values.First();
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
                    foreach (var item in ResolveSelectionSetInternal(setName))
                    {
                        if (item == null)
                        {
                            continue;
                        }

                        yield return new SpaceMapperResolvedItem
                        {
                            ItemKey = GetItemKey(item),
                            ModelItem = item,
                            DisplayName = item.DisplayName
                        };
                    }
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
        public static bool WriteProperty(
            ModelItem item,
            string categoryName,
            string propertyName,
            string value,
            WriteMode mode,
            string appendSeparator,
            bool showInternalProperties)
        {
            if (TryWriteProperty(item, categoryName, propertyName, value, mode, appendSeparator, showInternalProperties))
            {
                return true;
            }

            if (!showInternalProperties)
            {
                return TryWriteProperty(item, categoryName, propertyName, value, mode, appendSeparator, true);
            }

            return false;
        }

        private static bool TryWriteProperty(
            ModelItem item,
            string categoryName,
            string propertyName,
            string value,
            WriteMode mode,
            string appendSeparator,
            bool showInternalProperties)
        {
            if (item == null) return false;

            categoryName = string.IsNullOrWhiteSpace(categoryName) ? null : categoryName.Trim();
            propertyName = string.IsNullOrWhiteSpace(propertyName) ? null : propertyName.Trim();
            if (string.IsNullOrWhiteSpace(categoryName) || string.IsNullOrWhiteSpace(propertyName)) return false;

            var existing = string.Empty;
            if (mode == WriteMode.Append || mode == WriteMode.OnlyIfBlank)
            {
                existing = ReadProperty(item, categoryName, propertyName);
                if (mode == WriteMode.OnlyIfBlank && !string.IsNullOrWhiteSpace(existing))
                {
                    return false;
                }
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
                var propertyNode = (ComApi.InwGUIPropertyNode2)state.GetGUIPropertyNode(path, showInternalProperties);
                if (propertyNode == null)
                {
                    return false;
                }

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
            if (existingProp != null && IsSameValue(existingProp, finalValue))
            {
                return false;
            }

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
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[SpaceMapper] Failed to write {categoryName}.{propertyName} on {item.DisplayName}: {ex.Message}");
                return false;
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

        private static ComApi.InwOaPropertyVec GetOrCreatePropertyVector(
            ComApi.InwGUIPropertyNode2 propertyNode,
            string categoryName)
        {
            if (propertyNode == null || string.IsNullOrWhiteSpace(categoryName))
            {
                return null;
            }

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
                var state = ComBridge.State;
                propertyVector = (ComApi.InwOaPropertyVec)state.ObjectFactory(
                    ComApi.nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
            }

            return propertyVector;
        }

        private static bool IsSameValue(ComApi.InwOaProperty prop, string value)
        {
            if (prop == null)
            {
                return false;
            }

            try
            {
                var existing = prop.value?.ToString() ?? string.Empty;
                return string.Equals(existing, value ?? string.Empty, StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
        }

        private static string GetPropertyValue(ComApi.InwOaPropertyVec vec, string propertyName)
        {
            if (vec == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return string.Empty;
            }

            var prop = FindProperty(vec, propertyName);
            if (prop == null)
            {
                return string.Empty;
            }

            try
            {
                return prop.value?.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
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

        internal sealed class PropertyWriterSession
        {
            private readonly ModelItem _item;
            private ComApi.InwGUIPropertyNode2 _propertyNode;
            private readonly bool _preferShowInternal;
            private bool _fallbackTried;
            private readonly Dictionary<string, ComApi.InwOaPropertyVec> _categoryVectors =
                new(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _dirtyCategories = new(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, HashSet<string>> _dirtyProperties =
                new(StringComparer.OrdinalIgnoreCase);

            public PropertyWriterSession(ModelItem item, bool showInternalProperties)
            {
                _item = item;
                _preferShowInternal = showInternalProperties;
                if (item == null)
                {
                    return;
                }

                _propertyNode = TryGetPropertyNode(showInternalProperties);
                if (_propertyNode == null && !showInternalProperties)
                {
                    _fallbackTried = true;
                    _propertyNode = TryGetPropertyNode(true);
                }
            }

            public bool WriteProperty(string categoryName, string propertyName, string value, WriteMode mode, string appendSeparator)
            {
                if (_propertyNode == null || _item == null)
                {
                    return false;
                }

                categoryName = string.IsNullOrWhiteSpace(categoryName) ? null : categoryName.Trim();
                propertyName = string.IsNullOrWhiteSpace(propertyName) ? null : propertyName.Trim();
                if (string.IsNullOrWhiteSpace(categoryName) || string.IsNullOrWhiteSpace(propertyName))
                {
                    return false;
                }

                if (!_categoryVectors.TryGetValue(categoryName, out var propertyVector))
                {
                    propertyVector = GetOrCreatePropertyVector(_propertyNode, categoryName);
                    if (propertyVector == null)
                    {
                        return false;
                    }
                    _categoryVectors[categoryName] = propertyVector;
                }

                var existing = string.Empty;
                if (mode == WriteMode.Append || mode == WriteMode.OnlyIfBlank)
                {
                    existing = GetPropertyValue(propertyVector, propertyName);
                    if (string.IsNullOrWhiteSpace(existing))
                    {
                        existing = ReadProperty(_item, categoryName, propertyName);
                    }

                    if (mode == WriteMode.OnlyIfBlank && !string.IsNullOrWhiteSpace(existing))
                    {
                        return false;
                    }
                }

                var finalValue = value;
                if (mode == WriteMode.Append && !string.IsNullOrWhiteSpace(existing))
                {
                    finalValue = string.IsNullOrWhiteSpace(appendSeparator)
                        ? $"{existing},{value}"
                        : $"{existing}{appendSeparator}{value}";
                }

                var existingProp = FindProperty(propertyVector, propertyName);
                if (existingProp != null && IsSameValue(existingProp, finalValue))
                {
                    return false;
                }

                if (existingProp == null)
                {
                    var state = ComBridge.State;
                    existingProp = (ComApi.InwOaProperty)state.ObjectFactory(
                        ComApi.nwEObjectType.eObjectType_nwOaProperty, null, null);
                    existingProp.name = propertyName;
                    existingProp.UserName = propertyName;
                    propertyVector.Properties().Add(existingProp);
                }

                existingProp.value = finalValue;
                _dirtyCategories.Add(categoryName);
                TrackDirtyProperty(categoryName, propertyName);
                return true;
            }

            public void Commit()
            {
                if (_propertyNode == null || _dirtyCategories.Count == 0)
                {
                    return;
                }

                foreach (var kvp in _categoryVectors)
                {
                    if (!_dirtyCategories.Contains(kvp.Key))
                    {
                        continue;
                    }

                    try
                    {
                        _propertyNode.SetUserDefined(0, kvp.Key, kvp.Key, kvp.Value);
                    }
                    catch (Exception ex)
                    {
                        if (!_preferShowInternal && !_fallbackTried && TryFallbackPropertyNode())
                        {
                            try
                            {
                                _propertyNode.SetUserDefined(0, kvp.Key, kvp.Key, kvp.Value);
                                continue;
                            }
                            catch (Exception retryEx)
                            {
                                System.Diagnostics.Trace.WriteLine($"[SpaceMapper] Failed to write category {kvp.Key} on {_item?.DisplayName}: {retryEx.Message}");
                                continue;
                            }
                        }

                        System.Diagnostics.Trace.WriteLine($"[SpaceMapper] Failed to write category {kvp.Key} on {_item?.DisplayName}: {ex.Message}");
                    }
                }
            }

            public int DirtyCategoryCount => _dirtyCategories.Count;

            public int DirtyPropertyCount
            {
                get
                {
                    var count = 0;
                    foreach (var set in _dirtyProperties.Values)
                    {
                        count += set.Count;
                    }
                    return count;
                }
            }

            private ComApi.InwGUIPropertyNode2 TryGetPropertyNode(bool showInternalProperties)
            {
                try
                {
                    var state = ComBridge.State;
                    var path = ComBridge.ToInwOaPath(_item);
                    return (ComApi.InwGUIPropertyNode2)state.GetGUIPropertyNode(path, showInternalProperties);
                }
                catch
                {
                    return null;
                }
            }

            private bool TryFallbackPropertyNode()
            {
                _fallbackTried = true;
                _propertyNode = TryGetPropertyNode(true);
                return _propertyNode != null;
            }

            private void TrackDirtyProperty(string categoryName, string propertyName)
            {
                if (!_dirtyProperties.TryGetValue(categoryName, out var set))
                {
                    set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    _dirtyProperties[categoryName] = set;
                }
                set.Add(propertyName);
            }
        }
    }
}
