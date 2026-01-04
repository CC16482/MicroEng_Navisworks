using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;
using MicroEng.Navisworks.SpaceMapper.Estimation;
using MicroEng.Navisworks.SpaceMapper.Gpu;
using MicroEng.Navisworks.SpaceMapper.Geometry;

namespace MicroEng.Navisworks
{
    internal interface ISpatialIntersectionEngine
    {
        SpaceMapperProcessingMode Mode { get; }

        IList<ZoneTargetIntersection> ComputeIntersections(
            IReadOnlyList<ZoneGeometry> zones,
            IReadOnlyList<TargetGeometry> targets,
            SpaceMapperProcessingSettings settings,
            SpaceMapperPreflightCache preflightCache,
            SpaceMapperEngineDiagnostics diagnostics,
            IProgress<SpaceMapperProgress> progress = null,
            CancellationToken cancellationToken = default,
            SpaceMapperRunProgressState runProgress = null);
    }

    internal class CpuIntersectionEngine : ISpatialIntersectionEngine
    {
        public SpaceMapperProcessingMode Mode { get; }

        public CpuIntersectionEngine(SpaceMapperProcessingMode mode = SpaceMapperProcessingMode.CpuNormal)
        {
            Mode = mode;
        }

        public IList<ZoneTargetIntersection> ComputeIntersections(
            IReadOnlyList<ZoneGeometry> zones,
            IReadOnlyList<TargetGeometry> targets,
            SpaceMapperProcessingSettings settings,
            SpaceMapperPreflightCache preflightCache,
            SpaceMapperEngineDiagnostics diagnostics,
            IProgress<SpaceMapperProgress> progress = null,
            CancellationToken cancellationToken = default,
            SpaceMapperRunProgressState runProgress = null)
        {
            var results = new List<ZoneTargetIntersection>();
            if (zones == null || targets == null || zones.Count == 0 || targets.Count == 0)
            {
                return results;
            }

            diagnostics ??= new SpaceMapperEngineDiagnostics();
            var containmentEngine = ResolveZoneContainmentEngine(settings);
            var resolutionStrategy = ResolveZoneResolutionStrategy(settings);
            var zoneBoundsMode = ResolveZoneBoundsMode(settings);
            var targetBoundsMode = SpaceMapperBoundsResolver.ResolveTargetBoundsMode(settings, containmentEngine);
            var targetMidpointMode = settings?.TargetMidpointMode ?? SpaceMapperMidpointMode.BoundingBoxCenter;
            var containmentCalculationMode = ResolveContainmentCalculationMode(settings);
            var needsFraction = settings != null
                && (settings.WriteZoneContainmentPercentProperty
                    || (settings.WriteZoneBehaviorProperty
                        && containmentCalculationMode != SpaceMapperContainmentCalculationMode.Auto));
            var computeContainmentFraction = needsFraction
                && containmentCalculationMode == SpaceMapperContainmentCalculationMode.Auto;
            var needsPartial = settings != null
                && (settings.TagPartialSeparately
                    || settings.TreatPartialAsContained
                    || settings.WriteZoneBehaviorProperty
                    || settings.WriteZoneContainmentPercentProperty
                    || needsFraction);
            if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                needsPartial = false;
                computeContainmentFraction = false;
                needsFraction = false;
            }
            var usePointIndex = targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint;
            var treatPartialAsContained = settings != null && settings.TreatPartialAsContained;
            if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                treatPartialAsContained = false;
            }
            var enableMultiZone = settings == null || settings.EnableMultipleZones;
            diagnostics.PresetUsed = InferPresetFromBounds(zoneBoundsMode, targetBoundsMode, containmentEngine);
            var traversal = ResolveFastTraversal(
                settings?.FastTraversalMode ?? SpaceMapperFastTraversalMode.Auto,
                usePointIndex,
                needsPartial,
                targets.Count,
                zones.Count);
            diagnostics.TraversalUsed = traversal.ToString();

            if (traversal == SpaceMapperFastTraversalMode.TargetMajor)
            {
                return ComputeTargetMajorIntersections(
                    zones,
                    targets,
                    settings,
                    preflightCache,
                    diagnostics,
                    progress,
                    cancellationToken,
                    zoneBoundsMode,
                    containmentEngine,
                    targetBoundsMode,
                    targetMidpointMode,
                    enableMultiZone,
                    resolutionStrategy,
                    runProgress);
            }

            SpatialHashGrid grid;
            Aabb[] targetBounds;
            string[] targetKeys;

            var buildSw = Stopwatch.StartNew();
            if (preflightCache?.Grid != null
                && preflightCache.TargetBounds != null
                && preflightCache.TargetKeys != null
                && preflightCache.PointIndex == usePointIndex)
            {
                grid = preflightCache.Grid;
                targetBounds = preflightCache.TargetBounds;
                targetKeys = preflightCache.TargetKeys;
                diagnostics.UsedPreflightIndex = true;
            }
            else
            {
                diagnostics.UsedPreflightIndex = false;
                var prepared = BuildTargetBounds(targets, targetBoundsMode, targetMidpointMode);
                targetBounds = prepared.Bounds;
                targetKeys = prepared.Keys;
                if (targetBounds.Length == 0)
                {
                    buildSw.Stop();
                    diagnostics.BuildIndexTime = buildSw.Elapsed;
                    return results;
                }

                var worldBounds = ComputeWorldBounds(targetBounds);
                var cellSize = SpatialGridSizing.ComputeCellSize(worldBounds, settings?.IndexGranularity ?? 0);
                grid = new SpatialHashGrid(worldBounds, cellSize, targetBounds);
            }
            buildSw.Stop();
            diagnostics.BuildIndexTime = buildSw.Elapsed;

            if (grid == null || targetBounds == null || targetBounds.Length == 0)
            {
                return results;
            }

            IReadOnlyList<Vector3D>[] targetSamplePoints = null;
            if (needsFraction && containmentCalculationMode == SpaceMapperContainmentCalculationMode.TargetGeometry)
            {
                targetSamplePoints = BuildTargetGeometrySamples(targets);
                if (targetSamplePoints == null || targetSamplePoints.Length != targetBounds.Length)
                {
                    targetSamplePoints = null;
                }
            }

            runProgress?.SetStage(
                SpaceMapperRunStage.ComputingIntersections,
                "Computing intersections...",
                "Classifying targets into zones...");

            var useBestSelection = usePointIndex && !enableMultiZone;
            var collectAllHits = !useBestSelection;
            var zoneCount = zones.Count;
            var zoneBounds = new Aabb[zoneCount];
            var zoneValid = new bool[zoneCount];
            double[] zoneVolumes = null;
            double[] zoneCenterX = null;
            double[] zoneCenterY = null;
            double[] zoneCenterZ = null;

            for (int i = 0; i < zoneCount; i++)
            {
                var zone = zones[i];
                if (zone?.BoundingBox == null)
                {
                    zoneBounds[i] = new Aabb(0, 0, 0, 0, 0, 0);
                    continue;
                }

                zoneBounds[i] = GetZoneQueryBounds(zone, settings);
                zoneValid[i] = zone?.BoundingBox != null || zone?.RawBoundingBox != null;
            }

            if (useBestSelection)
            {
                zoneVolumes = new double[zoneCount];
                zoneCenterX = new double[zoneCount];
                zoneCenterY = new double[zoneCount];
                zoneCenterZ = new double[zoneCount];

                for (int i = 0; i < zoneCount; i++)
                {
                    if (!zoneValid[i])
                    {
                        zoneVolumes[i] = double.PositiveInfinity;
                        continue;
                    }

                    var bounds = zoneBounds[i];
                    zoneVolumes[i] = bounds.SizeX * bounds.SizeY * bounds.SizeZ;
                    zoneCenterX[i] = (bounds.MinX + bounds.MaxX) * 0.5;
                    zoneCenterY[i] = (bounds.MinY + bounds.MaxY) * 0.5;
                    zoneCenterZ[i] = (bounds.MinZ + bounds.MaxZ) * 0.5;
                }
            }

            ThreadLocal<List<ZoneTargetIntersection>> resultsLocal = null;
            ThreadLocal<Dictionary<int, BestZoneCandidate>> bestLocal = null;

            if (collectAllHits)
            {
                resultsLocal = new ThreadLocal<List<ZoneTargetIntersection>>(
                    () => new List<ZoneTargetIntersection>(),
                    trackAllValues: true);
            }
            else
            {
                bestLocal = new ThreadLocal<Dictionary<int, BestZoneCandidate>>(
                    () => new Dictionary<int, BestZoneCandidate>(),
                    trackAllValues: true);
            }
            var totalZones = zones.Count;
            var processedZones = 0;
            long candidatePairs = 0;
            int maxCandidates = 0;
            long totalTicks = 0;
            long narrowTicks = 0;
            const double SlowZoneThresholdSeconds = 10;
            var slowZoneTicksThreshold = SlowZoneThresholdSeconds * Stopwatch.Frequency;
            var slowZones = new ConcurrentQueue<SpaceMapperSlowZoneInfo>();
            long lastDetailTicks = 0;
            var detailIntervalTicks = Stopwatch.Frequency / 2;

            var visitedLocal = new ThreadLocal<int[]>(() => new int[targetBounds.Length]);
            var stampLocal = new ThreadLocal<int>(() => 0);
            var trackNarrow = Mode == SpaceMapperProcessingMode.Debug;
            var parallelOptions = new ParallelOptions
            {
                CancellationToken = cancellationToken,
                MaxDegreeOfParallelism = settings?.MaxThreads ?? Environment.ProcessorCount
            };

            var loopSw = Stopwatch.StartNew();
            Parallel.For(0, zoneCount, parallelOptions, zoneIndex =>
            {
                var zone = zones[zoneIndex];
                if (!zoneValid[zoneIndex])
                {
                    var zonesDoneLocal = Interlocked.Increment(ref processedZones);
                    if (runProgress != null && (((zonesDoneLocal & 63) == 0) || zonesDoneLocal == totalZones))
                    {
                        runProgress.SetTotals(totalZones, targets.Count);
                        runProgress.UpdateZonesProcessed(zonesDoneLocal);
                        runProgress.UpdateCandidatePairs(Interlocked.Read(ref candidatePairs));
                    }
                    return;
                }

                var zoneName = string.IsNullOrWhiteSpace(zone.DisplayName) ? zone.ZoneId : zone.DisplayName;
                if (runProgress != null)
                {
                    var nowTicks = Stopwatch.GetTimestamp();
                    var prevTicks = Interlocked.Read(ref lastDetailTicks);
                    if (nowTicks - prevTicks > detailIntervalTicks
                        && Interlocked.CompareExchange(ref lastDetailTicks, nowTicks, prevTicks) == prevTicks)
                    {
                        runProgress.UpdateDetail($"Zone {zoneIndex + 1}/{totalZones}: {zoneName}");
                    }
                }

                var zoneBoundsLocal = zoneBounds[zoneIndex];
                var visited = visitedLocal.Value;
                var stamp = stampLocal.Value;
                var zoneCandidates = 0;
                var zoneNarrowTicks = 0L;
                var zoneStart = Stopwatch.GetTimestamp();
                var localResults = resultsLocal?.Value;
                var localBest = bestLocal?.Value;

                grid.VisitCandidates(zoneBoundsLocal, visited, ref stamp, idx =>
                {
                    zoneCandidates++;
                    if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
                    {
                        var targetBoundsLocal = targetBounds[idx];
                        var cx = (targetBoundsLocal.MinX + targetBoundsLocal.MaxX) * 0.5;
                        var cy = (targetBoundsLocal.MinY + targetBoundsLocal.MaxY) * 0.5;
                        var cz = (targetBoundsLocal.MinZ + targetBoundsLocal.MaxZ) * 0.5;

                        var inside = ZoneContainsPoint(zone, zoneBoundsLocal, zoneBoundsMode, containmentEngine, diagnostics, cx, cy, cz);
                        if (!inside)
                        {
                            return;
                        }

                        var overlapVolume = EstimateOverlap(zoneBoundsLocal, targetBoundsLocal);

                        if (useBestSelection)
                        {
                            var dx = cx - zoneCenterX[zoneIndex];
                            var dy = cy - zoneCenterY[zoneIndex];
                            var dz = cz - zoneCenterZ[zoneIndex];
                            var candidate = new BestZoneCandidate
                            {
                                ZoneIndex = zoneIndex,
                                IsContained = true,
                                IsPartial = false,
                                Volume = zoneVolumes[zoneIndex],
                                DistanceSq = (dx * dx) + (dy * dy) + (dz * dz),
                                OverlapVolume = overlapVolume
                            };

                            if (!localBest.TryGetValue(idx, out var best) || IsBetterCandidate(candidate, best, resolutionStrategy))
                            {
                                localBest[idx] = candidate;
                            }
                        }
                        else
                        {
                            localResults.Add(new ZoneTargetIntersection
                            {
                                ZoneId = zone.ZoneId,
                                TargetItemKey = targetKeys[idx],
                                IsContained = true,
                                IsPartial = false,
                                OverlapVolume = overlapVolume
                            });
                        }

                        return;
                    }

                    ZoneTargetIntersection hit;
                    if (trackNarrow)
                    {
                        var t0 = Stopwatch.GetTimestamp();
                        hit = ClassifyIntersection(
                            zone,
                            zoneBoundsLocal,
                            targetBounds[idx],
                            targetKeys[idx],
                            zoneBoundsMode,
                            containmentEngine,
                            diagnostics,
                            targetBoundsMode,
                            computeContainmentFraction,
                            needsPartial,
                            treatPartialAsContained);
                        var t1 = Stopwatch.GetTimestamp();
                        zoneNarrowTicks += t1 - t0;
                    }
                    else
                    {
                        hit = ClassifyIntersection(
                            zone,
                            zoneBoundsLocal,
                            targetBounds[idx],
                            targetKeys[idx],
                            zoneBoundsMode,
                            containmentEngine,
                            diagnostics,
                            targetBoundsMode,
                            computeContainmentFraction,
                            needsPartial,
                            treatPartialAsContained);
                    }

                    if (hit != null)
                    {
                        if (needsFraction && containmentCalculationMode != SpaceMapperContainmentCalculationMode.Auto)
                        {
                            hit.ContainmentFraction = ComputeContainmentFraction(
                                zone,
                                zoneBoundsLocal,
                                targetBounds[idx],
                                zoneBoundsMode,
                                containmentEngine,
                                containmentCalculationMode,
                                targetSamplePoints?[idx],
                                diagnostics);
                        }

                        localResults.Add(hit);
                    }
                });

                var zoneEnd = Stopwatch.GetTimestamp();
                var zoneElapsedTicks = zoneEnd - zoneStart;
                if (!trackNarrow || targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
                {
                    zoneNarrowTicks = zoneElapsedTicks;
                }
                stampLocal.Value = stamp;

                Interlocked.Add(ref candidatePairs, zoneCandidates);
                UpdateMax(ref maxCandidates, zoneCandidates);
                Interlocked.Add(ref totalTicks, zoneElapsedTicks);
                Interlocked.Add(ref narrowTicks, zoneNarrowTicks);
                if (zoneElapsedTicks > slowZoneTicksThreshold)
                {
                    slowZones.Enqueue(new SpaceMapperSlowZoneInfo
                    {
                        ZoneIndex = zoneIndex,
                        ZoneId = zone.ZoneId,
                        ZoneName = zoneName,
                        CandidateCount = zoneCandidates,
                        Elapsed = TimeSpan.FromSeconds(zoneElapsedTicks / (double)Stopwatch.Frequency)
                    });
                }

                var zonesDone = Interlocked.Increment(ref processedZones);
                if (runProgress != null && (((zonesDone & 63) == 0) || zonesDone == totalZones))
                {
                    runProgress.SetTotals(totalZones, targets.Count);
                    runProgress.UpdateZonesProcessed(zonesDone);
                    runProgress.UpdateCandidatePairs(Interlocked.Read(ref candidatePairs));
                }
                if (progress != null && (zonesDone % 4 == 0 || zonesDone == totalZones))
                {
                    var pairsDone = Interlocked.Read(ref candidatePairs);
                    var totalPairs = preflightCache?.LastResult?.CandidatePairs ?? 0;
                    progress.Report(new SpaceMapperProgress
                    {
                        ProcessedPairs = pairsDone > int.MaxValue ? int.MaxValue : (int)pairsDone,
                        TotalPairs = totalPairs > int.MaxValue ? int.MaxValue : (int)totalPairs,
                        ZonesProcessed = zonesDone,
                        TargetsProcessed = targetBounds.Length
                    });
                }
            });
            loopSw.Stop();

            diagnostics.CandidatePairs = candidatePairs;
            diagnostics.MaxCandidatesPerZone = maxCandidates;
            diagnostics.AvgCandidatesPerZone = totalZones == 0 ? 0 : candidatePairs / (double)totalZones;
            diagnostics.MaxCandidatesPerTarget = 0;
            diagnostics.AvgCandidatesPerTarget = 0;
            diagnostics.SlowZoneThresholdSeconds = SlowZoneThresholdSeconds;
            if (!slowZones.IsEmpty)
            {
                diagnostics.SlowZones = slowZones
                    .OrderByDescending(z => z.Elapsed)
                    .Take(50)
                    .ToList();
            }
            if (trackNarrow)
            {
                diagnostics.NarrowPhaseTime = ToTimeSpan(narrowTicks);
                diagnostics.CandidateQueryTime = ToTimeSpan(Math.Max(0, totalTicks - narrowTicks));
            }
            else
            {
                diagnostics.NarrowPhaseTime = loopSw.Elapsed;
                diagnostics.CandidateQueryTime = TimeSpan.Zero;
            }

            if (collectAllHits)
            {
                foreach (var local in resultsLocal.Values)
                {
                    if (local != null && local.Count > 0)
                    {
                        results.AddRange(local);
                    }
                }

                resultsLocal.Dispose();
            }
            else if (bestLocal != null)
            {
                var merged = new Dictionary<int, BestZoneCandidate>();
                foreach (var local in bestLocal.Values)
                {
                    foreach (var kvp in local)
                    {
                        if (!merged.TryGetValue(kvp.Key, out var best) || IsBetterCandidate(kvp.Value, best, resolutionStrategy))
                        {
                            merged[kvp.Key] = kvp.Value;
                        }
                    }
                }

                foreach (var kvp in merged)
                {
                    var targetIndex = kvp.Key;
                    var best = kvp.Value;
                    if (best.ZoneIndex < 0 || best.ZoneIndex >= zones.Count)
                    {
                        continue;
                    }

                    results.Add(new ZoneTargetIntersection
                    {
                        ZoneId = zones[best.ZoneIndex].ZoneId,
                        TargetItemKey = targetKeys[targetIndex],
                        IsContained = best.IsContained,
                        IsPartial = best.IsPartial,
                        OverlapVolume = best.OverlapVolume
                    });
                }

                bestLocal.Dispose();
            }

            visitedLocal.Dispose();
            stampLocal.Dispose();

            return results;
        }

        private IList<ZoneTargetIntersection> ComputeTargetMajorIntersections(
            IReadOnlyList<ZoneGeometry> zones,
            IReadOnlyList<TargetGeometry> targets,
            SpaceMapperProcessingSettings settings,
            SpaceMapperPreflightCache preflightCache,
            SpaceMapperEngineDiagnostics diagnostics,
            IProgress<SpaceMapperProgress> progress,
            CancellationToken cancellationToken,
            SpaceMapperZoneBoundsMode zoneBoundsMode,
            SpaceMapperZoneContainmentEngine containmentEngine,
            SpaceMapperTargetBoundsMode targetBoundsMode,
            SpaceMapperMidpointMode targetMidpointMode,
            bool enableMultiZone,
            SpaceMapperZoneResolutionStrategy resolutionStrategy,
            SpaceMapperRunProgressState runProgress)
        {
            var results = new List<ZoneTargetIntersection>();
            var zoneCount = zones.Count;
            var zoneBounds = new Aabb[zoneCount];
            var zoneValid = new bool[zoneCount];
            var usePointIndex = targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint;
            var containmentCalculationMode = ResolveContainmentCalculationMode(settings);
            var needsFraction = settings != null
                && (settings.WriteZoneContainmentPercentProperty
                    || (settings.WriteZoneBehaviorProperty
                        && containmentCalculationMode != SpaceMapperContainmentCalculationMode.Auto));
            if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                needsFraction = false;
            }

            for (int i = 0; i < zoneCount; i++)
            {
                var zone = zones[i];
                if (zone?.BoundingBox == null)
                {
                    zoneBounds[i] = new Aabb(0, 0, 0, 0, 0, 0);
                    continue;
                }

                zoneBounds[i] = GetZoneQueryBounds(zone, settings);
                zoneValid[i] = true;
            }

            double[] zoneVolumes = null;
            double[] zoneCenterX = null;
            double[] zoneCenterY = null;
            double[] zoneCenterZ = null;
            if (!enableMultiZone)
            {
                zoneVolumes = new double[zoneCount];
                zoneCenterX = new double[zoneCount];
                zoneCenterY = new double[zoneCount];
                zoneCenterZ = new double[zoneCount];

                for (int i = 0; i < zoneCount; i++)
                {
                    if (!zoneValid[i])
                    {
                        zoneVolumes[i] = double.PositiveInfinity;
                        continue;
                    }

                    var bounds = zoneBounds[i];
                    zoneVolumes[i] = bounds.SizeX * bounds.SizeY * bounds.SizeZ;
                    zoneCenterX[i] = (bounds.MinX + bounds.MaxX) * 0.5;
                    zoneCenterY[i] = (bounds.MinY + bounds.MaxY) * 0.5;
                    zoneCenterZ[i] = (bounds.MinZ + bounds.MaxZ) * 0.5;
                }
            }

            SpatialHashGrid zoneGrid = null;
            Aabb[] targetBounds;
            string[] targetKeys;
            int[] zoneIndexMap = null;

            var buildSw = Stopwatch.StartNew();
            if (preflightCache?.ZoneGrid != null
                && preflightCache.TargetBounds != null
                && preflightCache.TargetKeys != null
                && preflightCache.ZoneBoundsInflated != null
                && preflightCache.ZoneIndexMap != null
                && preflightCache.PointIndex == usePointIndex)
            {
                zoneGrid = preflightCache.ZoneGrid;
                targetBounds = preflightCache.TargetBounds;
                targetKeys = preflightCache.TargetKeys;
                zoneIndexMap = preflightCache.ZoneIndexMap;
                diagnostics.UsedPreflightIndex = true;
            }
            else
            {
                diagnostics.UsedPreflightIndex = false;
                var prepared = BuildTargetBounds(targets, targetBoundsMode, targetMidpointMode);
                targetBounds = prepared.Bounds;
                targetKeys = prepared.Keys;
                if (targetBounds.Length == 0)
                {
                    buildSw.Stop();
                    diagnostics.BuildIndexTime = buildSw.Elapsed;
                    return results;
                }

                var boundsForGrid = new List<Aabb>(zoneCount);
                var indexMap = new List<int>(zoneCount);
                for (int i = 0; i < zoneCount; i++)
                {
                    if (!zoneValid[i])
                    {
                        continue;
                    }
                    boundsForGrid.Add(zoneBounds[i]);
                    indexMap.Add(i);
                }

                var zoneBoundsInflated = boundsForGrid.ToArray();
                zoneIndexMap = indexMap.ToArray();
                if (zoneBoundsInflated.Length == 0)
                {
                    buildSw.Stop();
                    diagnostics.BuildIndexTime = buildSw.Elapsed;
                    return results;
                }

                var worldBounds = ComputeWorldBounds(zoneBoundsInflated);
                var cellSize = SpatialGridSizing.ComputeCellSize(worldBounds, settings?.IndexGranularity ?? 0);
                zoneGrid = new SpatialHashGrid(worldBounds, cellSize, zoneBoundsInflated);
            }
            buildSw.Stop();
            diagnostics.BuildIndexTime = buildSw.Elapsed;

            if (zoneGrid == null || targetBounds == null || targetBounds.Length == 0 || zoneIndexMap == null || zoneIndexMap.Length == 0)
            {
                return results;
            }

            IReadOnlyList<Vector3D>[] targetSamplePoints = null;
            if (needsFraction && containmentCalculationMode == SpaceMapperContainmentCalculationMode.TargetGeometry)
            {
                targetSamplePoints = BuildTargetGeometrySamples(targets);
                if (targetSamplePoints == null || targetSamplePoints.Length != targetBounds.Length)
                {
                    targetSamplePoints = null;
                }
            }

            var trackNarrow = Mode == SpaceMapperProcessingMode.Debug;
            var parallelOptions = new ParallelOptions
            {
                CancellationToken = cancellationToken,
                MaxDegreeOfParallelism = settings?.MaxThreads ?? Environment.ProcessorCount
            };

            ThreadLocal<List<ZoneTargetIntersection>> resultsLocal = null;
            ZoneTargetIntersection[] bestHits = null;

            if (enableMultiZone)
            {
                resultsLocal = new ThreadLocal<List<ZoneTargetIntersection>>(
                    () => new List<ZoneTargetIntersection>(),
                    trackAllValues: true);
            }
            else
            {
                bestHits = new ZoneTargetIntersection[targetBounds.Length];
            }

            long candidatePairs = 0;
            int maxCandidates = 0;
            long totalTicks = 0;
            long narrowTicks = 0;
            var processedTargets = 0;
            long lastDetailTicks = 0;
            var detailIntervalTicks = Stopwatch.Frequency / 2;

            runProgress?.SetStage(
                SpaceMapperRunStage.ComputingIntersections,
                "Computing intersections...",
                "Classifying targets into zones...");

            var loopSw = Stopwatch.StartNew();
            Parallel.For(0, targetBounds.Length, parallelOptions, targetIndex =>
            {
                var bounds = targetBounds[targetIndex];
                var cx = (bounds.MinX + bounds.MaxX) * 0.5;
                var cy = (bounds.MinY + bounds.MaxY) * 0.5;
                var cz = (bounds.MinZ + bounds.MaxZ) * 0.5;
                var candidateCount = 0;
                var localResults = resultsLocal?.Value;

                BestZoneCandidate best = default;
                var hasBest = false;
                var targetStart = Stopwatch.GetTimestamp();

                if (runProgress != null)
                {
                    var nowTicks = Stopwatch.GetTimestamp();
                    var prevTicks = Interlocked.Read(ref lastDetailTicks);
                    if (nowTicks - prevTicks > detailIntervalTicks
                        && Interlocked.CompareExchange(ref lastDetailTicks, nowTicks, prevTicks) == prevTicks)
                    {
                        string targetName = null;
                        if (targetIndex < targets.Count)
                        {
                            targetName = targets[targetIndex]?.DisplayName;
                        }
                        if (string.IsNullOrWhiteSpace(targetName) && targetKeys != null && targetIndex < targetKeys.Length)
                        {
                            targetName = targetKeys[targetIndex];
                        }
                        runProgress.UpdateDetail($"Target {targetIndex + 1}/{targetBounds.Length}: {targetName}");
                    }
                }

                zoneGrid.VisitPointCandidates(cx, cy, cz, idx =>
                {
                    candidateCount++;
                    if ((uint)idx >= (uint)zoneIndexMap.Length)
                    {
                        return;
                    }

                    var zoneIndex = zoneIndexMap[idx];
                    if ((uint)zoneIndex >= (uint)zoneCount)
                    {
                        return;
                    }

                    if (!zoneValid[zoneIndex])
                    {
                        return;
                    }

                    var zoneBoundsLocal = zoneBounds[zoneIndex];
                    var inside = ZoneContainsPoint(zones[zoneIndex], zoneBoundsLocal, zoneBoundsMode, containmentEngine, diagnostics, cx, cy, cz);

                    if (!inside)
                    {
                        return;
                    }

                    var overlapVolume = EstimateOverlap(zoneBoundsLocal, bounds);

                    if (enableMultiZone)
                    {
                        localResults.Add(new ZoneTargetIntersection
                        {
                            ZoneId = zones[zoneIndex].ZoneId,
                            TargetItemKey = targetKeys[targetIndex],
                            IsContained = true,
                            IsPartial = false,
                            OverlapVolume = overlapVolume
                        });
                    }
                    else
                    {
                        var dx = cx - zoneCenterX[zoneIndex];
                        var dy = cy - zoneCenterY[zoneIndex];
                        var dz = cz - zoneCenterZ[zoneIndex];
                        var candidate = new BestZoneCandidate
                        {
                            ZoneIndex = zoneIndex,
                            IsContained = true,
                            IsPartial = false,
                            Volume = zoneVolumes[zoneIndex],
                            DistanceSq = (dx * dx) + (dy * dy) + (dz * dz),
                            OverlapVolume = overlapVolume
                        };

                        if (!hasBest || IsBetterCandidate(candidate, best, resolutionStrategy))
                        {
                            best = candidate;
                            hasBest = true;
                        }
                    }
                });

                if (!enableMultiZone && hasBest)
                {
                    bestHits[targetIndex] = new ZoneTargetIntersection
                    {
                        ZoneId = zones[best.ZoneIndex].ZoneId,
                        TargetItemKey = targetKeys[targetIndex],
                        IsContained = best.IsContained,
                        IsPartial = best.IsPartial,
                        OverlapVolume = best.OverlapVolume
                    };
                }

                var targetEnd = Stopwatch.GetTimestamp();
                if (!trackNarrow)
                {
                    Interlocked.Add(ref narrowTicks, targetEnd - targetStart);
                }
                else
                {
                    Interlocked.Add(ref narrowTicks, targetEnd - targetStart);
                }
                Interlocked.Add(ref totalTicks, targetEnd - targetStart);

                Interlocked.Add(ref candidatePairs, candidateCount);
                UpdateMax(ref maxCandidates, candidateCount);

                var targetsDone = Interlocked.Increment(ref processedTargets);
                if (runProgress != null && (((targetsDone & 255) == 0) || targetsDone == targetBounds.Length))
                {
                    runProgress.SetTotals(zones.Count, targets.Count);
                    runProgress.UpdateTargetsProcessed(targetsDone);
                    runProgress.UpdateCandidatePairs(Interlocked.Read(ref candidatePairs));
                }
                if (progress != null && (targetsDone % 256 == 0 || targetsDone == targetBounds.Length))
                {
                    var pairsDone = Interlocked.Read(ref candidatePairs);
                    var totalPairs = preflightCache?.LastResult?.CandidatePairs ?? 0;
                    progress.Report(new SpaceMapperProgress
                    {
                        ProcessedPairs = pairsDone > int.MaxValue ? int.MaxValue : (int)pairsDone,
                        TotalPairs = totalPairs > int.MaxValue ? int.MaxValue : (int)totalPairs,
                        ZonesProcessed = zones.Count,
                        TargetsProcessed = targetsDone
                    });
                }
            });
            loopSw.Stop();

            diagnostics.CandidatePairs = candidatePairs;
            diagnostics.MaxCandidatesPerTarget = maxCandidates;
            diagnostics.AvgCandidatesPerTarget = targetBounds.Length == 0 ? 0 : candidatePairs / (double)targetBounds.Length;
            diagnostics.MaxCandidatesPerZone = 0;
            diagnostics.AvgCandidatesPerZone = 0;
            if (trackNarrow)
            {
                diagnostics.NarrowPhaseTime = ToTimeSpan(narrowTicks);
                diagnostics.CandidateQueryTime = ToTimeSpan(Math.Max(0, totalTicks - narrowTicks));
            }
            else
            {
                diagnostics.NarrowPhaseTime = loopSw.Elapsed;
                diagnostics.CandidateQueryTime = TimeSpan.Zero;
            }

            if (enableMultiZone)
            {
                foreach (var local in resultsLocal.Values)
                {
                    if (local != null && local.Count > 0)
                    {
                        results.AddRange(local);
                    }
                }

                resultsLocal.Dispose();
            }
            else
            {
                for (int i = 0; i < bestHits.Length; i++)
                {
                    if (bestHits[i] != null)
                    {
                        results.Add(bestHits[i]);
                    }
                }
            }

            return results;
        }

        internal struct BestZoneCandidate
        {
            public int ZoneIndex;
            public bool IsContained;
            public bool IsPartial;
            public double Volume;
            public double DistanceSq;
            public double OverlapVolume;
        }

        internal static ZoneTargetIntersection ClassifyIntersection(
            ZoneGeometry zone,
            in Aabb zoneBounds,
            in Aabb targetBounds,
            string targetKey,
            SpaceMapperZoneBoundsMode zoneBoundsMode,
            SpaceMapperZoneContainmentEngine containmentEngine,
            SpaceMapperEngineDiagnostics diagnostics,
            SpaceMapperTargetBoundsMode targetBoundsMode,
            bool computeContainmentFraction,
            bool needsPartial,
            bool treatPartialAsContained)
        {
            // Tier A: mesh accurate point-in-mesh (existing behavior)
            if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                return ClassifyPointTarget(
                    zone,
                    zoneBounds,
                    targetBounds,
                    targetKey,
                    zoneBoundsMode,
                    containmentEngine,
                    diagnostics,
                    computeContainmentFraction);
            }

            // Tier B: mesh accurate multi-sample classification
            if (containmentEngine == SpaceMapperZoneContainmentEngine.MeshAccurate)
            {
                return ClassifyByMeshAccurateSamples(
                    zone,
                    zoneBounds,
                    targetBounds,
                    targetKey,
                    zoneBoundsMode,
                    containmentEngine,
                    diagnostics,
                    targetBoundsMode,
                    computeContainmentFraction,
                    needsPartial,
                    treatPartialAsContained);
            }

            var usePlanes = zoneBoundsMode != SpaceMapperZoneBoundsMode.Aabb
                && zone?.Planes != null
                && zone.Planes.Count > 0;

            if (!usePlanes)
            {
                return ClassifyAabbOnly(
                    zone,
                    zoneBounds,
                    targetBounds,
                    targetKey,
                    computeContainmentFraction,
                    needsPartial,
                    treatPartialAsContained);
            }

            var extraSamples = targetBoundsMode != SpaceMapperTargetBoundsMode.Aabb;
            return ClassifyByPlanes(
                zone,
                zoneBounds,
                targetBounds,
                targetKey,
                computeContainmentFraction,
                needsPartial,
                treatPartialAsContained,
                extraSamples);
        }

        internal static ZoneTargetIntersection ClassifyPointTarget(
            ZoneGeometry zone,
            in Aabb zoneBounds,
            in Aabb targetBounds,
            string targetKey,
            SpaceMapperZoneBoundsMode zoneBoundsMode,
            SpaceMapperZoneContainmentEngine containmentEngine,
            SpaceMapperEngineDiagnostics diagnostics,
            bool computeContainmentFraction)
        {
            var cx = (targetBounds.MinX + targetBounds.MaxX) * 0.5;
            var cy = (targetBounds.MinY + targetBounds.MaxY) * 0.5;
            var cz = (targetBounds.MinZ + targetBounds.MaxZ) * 0.5;

            var inside = ZoneContainsPoint(zone, zoneBounds, zoneBoundsMode, containmentEngine, diagnostics, cx, cy, cz);

            if (!inside)
            {
                return null;
            }

            return new ZoneTargetIntersection
            {
                ZoneId = zone.ZoneId,
                TargetItemKey = targetKey,
                IsContained = true,
                IsPartial = false,
                OverlapVolume = EstimateOverlap(zoneBounds, targetBounds),
                ContainmentFraction = computeContainmentFraction ? 1.0 : null
            };
        }

        internal static ZoneTargetIntersection ClassifyByMeshAccurateSamples(
            ZoneGeometry zone,
            in Aabb zoneBounds,
            in Aabb targetBounds,
            string targetKey,
            SpaceMapperZoneBoundsMode zoneBoundsMode,
            SpaceMapperZoneContainmentEngine containmentEngine,
            SpaceMapperEngineDiagnostics diagnostics,
            SpaceMapperTargetBoundsMode targetBoundsMode,
            bool computeContainmentFraction,
            bool needsPartial,
            bool treatPartialAsContained)
        {
            if (zone == null)
            {
                return null;
            }

            var zoneBoundsLocal = zoneBounds;
            var targetBoundsLocal = targetBounds;

            if (!zoneBoundsLocal.Intersects(targetBoundsLocal))
            {
                return null;
            }

            var anyInside = false;
            var anyOutside = false;
            var sampleCount = 0;
            var insideCount = 0;

            ZoneTargetIntersection BuildResult()
            {
                if (!anyInside)
                {
                    return null;
                }

                var isPartial = anyOutside;
                if (isPartial && !needsPartial)
                {
                    return null;
                }

                double? fraction = null;
                if (computeContainmentFraction && sampleCount > 0)
                {
                    fraction = insideCount / (double)sampleCount;
                    if (fraction < 0) fraction = 0;
                    else if (fraction > 1) fraction = 1;
                }

                return new ZoneTargetIntersection
                {
                    ZoneId = zone.ZoneId,
                    TargetItemKey = targetKey,
                    IsContained = !isPartial || treatPartialAsContained,
                    IsPartial = isPartial,
                    OverlapVolume = EstimateOverlap(zoneBoundsLocal, targetBoundsLocal),
                    ContainmentFraction = fraction
                };
            }

            bool Sample(double x, double y, double z)
            {
                sampleCount++;
                var inside = ZoneContainsPoint(zone, zoneBoundsLocal, zoneBoundsMode, containmentEngine, diagnostics, x, y, z);

                if (inside)
                {
                    anyInside = true;
                    insideCount++;
                }
                else
                {
                    anyOutside = true;
                }

                if (!computeContainmentFraction)
                {
                    if (!needsPartial && anyOutside)
                    {
                        return false;
                    }

                    if (anyInside && anyOutside)
                    {
                        return false;
                    }
                }

                return true;
            }

            var minX = targetBoundsLocal.MinX;
            var minY = targetBoundsLocal.MinY;
            var minZ = targetBoundsLocal.MinZ;
            var maxX = targetBoundsLocal.MaxX;
            var maxY = targetBoundsLocal.MaxY;
            var maxZ = targetBoundsLocal.MaxZ;

            if (!Sample(minX, minY, minZ)) return BuildResult();
            if (!Sample(maxX, minY, minZ)) return BuildResult();
            if (!Sample(minX, maxY, minZ)) return BuildResult();
            if (!Sample(maxX, maxY, minZ)) return BuildResult();
            if (!Sample(minX, minY, maxZ)) return BuildResult();
            if (!Sample(maxX, minY, maxZ)) return BuildResult();
            if (!Sample(minX, maxY, maxZ)) return BuildResult();
            if (!Sample(maxX, maxY, maxZ)) return BuildResult();

            if (targetBoundsMode != SpaceMapperTargetBoundsMode.Aabb && !(anyInside && anyOutside))
            {
                var cx = (minX + maxX) * 0.5;
                var cy = (minY + maxY) * 0.5;
                var cz = (minZ + maxZ) * 0.5;

                if (!Sample(cx, cy, cz)) return BuildResult();

                if ((targetBoundsMode == SpaceMapperTargetBoundsMode.KDop
                    || targetBoundsMode == SpaceMapperTargetBoundsMode.Hull)
                    && !(anyInside && anyOutside))
                {
                    if (!Sample(minX, cy, cz)) return BuildResult();
                    if (!Sample(maxX, cy, cz)) return BuildResult();
                    if (!Sample(cx, minY, cz)) return BuildResult();
                    if (!Sample(cx, maxY, cz)) return BuildResult();
                    if (!Sample(cx, cy, minZ)) return BuildResult();
                    if (!Sample(cx, cy, maxZ)) return BuildResult();
                }

                if (targetBoundsMode == SpaceMapperTargetBoundsMode.Hull && !(anyInside && anyOutside))
                {
                    if (!Sample(cx, minY, minZ)) return BuildResult();
                    if (!Sample(cx, maxY, minZ)) return BuildResult();
                    if (!Sample(cx, minY, maxZ)) return BuildResult();
                    if (!Sample(cx, maxY, maxZ)) return BuildResult();

                    if (!Sample(minX, cy, minZ)) return BuildResult();
                    if (!Sample(maxX, cy, minZ)) return BuildResult();
                    if (!Sample(minX, cy, maxZ)) return BuildResult();
                    if (!Sample(maxX, cy, maxZ)) return BuildResult();

                    if (!Sample(minX, minY, cz)) return BuildResult();
                    if (!Sample(maxX, minY, cz)) return BuildResult();
                    if (!Sample(minX, maxY, cz)) return BuildResult();
                    if (!Sample(maxX, maxY, cz)) return BuildResult();
                }
            }

            return BuildResult();
        }

        internal static ZoneTargetIntersection ClassifyAabbOnly(
            ZoneGeometry zone,
            in Aabb zoneBounds,
            in Aabb targetBounds,
            string targetKey,
            bool computeContainmentFraction,
            bool needsPartial,
            bool treatPartialAsContained)
        {
            if (!zoneBounds.Intersects(targetBounds))
            {
                return null;
            }

            var contained = targetBounds.MinX >= zoneBounds.MinX
                && targetBounds.MinY >= zoneBounds.MinY
                && targetBounds.MinZ >= zoneBounds.MinZ
                && targetBounds.MaxX <= zoneBounds.MaxX
                && targetBounds.MaxY <= zoneBounds.MaxY
                && targetBounds.MaxZ <= zoneBounds.MaxZ;

            if (!contained && !needsPartial)
            {
                return null;
            }

            double? fraction = null;
            if (computeContainmentFraction)
            {
                fraction = contained
                    ? 1.0
                    : ComputeFractionFromOverlap(EstimateOverlap(zoneBounds, targetBounds), targetBounds);
            }

            return new ZoneTargetIntersection
            {
                ZoneId = zone.ZoneId,
                TargetItemKey = targetKey,
                IsContained = contained || treatPartialAsContained,
                IsPartial = !contained,
                OverlapVolume = EstimateOverlap(zoneBounds, targetBounds),
                ContainmentFraction = fraction
            };
        }

        internal static ZoneTargetIntersection ClassifyByPlanes(
            ZoneGeometry zone,
            in Aabb zoneBounds,
            in Aabb targetBounds,
            string targetKey,
            bool computeContainmentFraction,
            bool needsPartial,
            bool treatPartialAsContained,
            bool extraSamples)
        {
            if (zone?.Planes == null || zone.Planes.Count == 0)
            {
                return null;
            }

            var anyInside = false;
            var anyOutside = false;

            TestCorner(zone.Planes, targetBounds.MinX, targetBounds.MinY, targetBounds.MinZ, ref anyInside, ref anyOutside);
            TestCorner(zone.Planes, targetBounds.MaxX, targetBounds.MinY, targetBounds.MinZ, ref anyInside, ref anyOutside);
            TestCorner(zone.Planes, targetBounds.MinX, targetBounds.MaxY, targetBounds.MinZ, ref anyInside, ref anyOutside);
            TestCorner(zone.Planes, targetBounds.MaxX, targetBounds.MaxY, targetBounds.MinZ, ref anyInside, ref anyOutside);
            TestCorner(zone.Planes, targetBounds.MinX, targetBounds.MinY, targetBounds.MaxZ, ref anyInside, ref anyOutside);
            TestCorner(zone.Planes, targetBounds.MaxX, targetBounds.MinY, targetBounds.MaxZ, ref anyInside, ref anyOutside);
            TestCorner(zone.Planes, targetBounds.MinX, targetBounds.MaxY, targetBounds.MaxZ, ref anyInside, ref anyOutside);
            TestCorner(zone.Planes, targetBounds.MaxX, targetBounds.MaxY, targetBounds.MaxZ, ref anyInside, ref anyOutside);

            if (extraSamples && !(anyInside && anyOutside))
            {
                var cx = (targetBounds.MinX + targetBounds.MaxX) * 0.5;
                var cy = (targetBounds.MinY + targetBounds.MaxY) * 0.5;
                var cz = (targetBounds.MinZ + targetBounds.MaxZ) * 0.5;
                TestCorner(zone.Planes, cx, cy, cz, ref anyInside, ref anyOutside);

                TestCorner(zone.Planes, targetBounds.MinX, cy, cz, ref anyInside, ref anyOutside);
                TestCorner(zone.Planes, targetBounds.MaxX, cy, cz, ref anyInside, ref anyOutside);
                TestCorner(zone.Planes, cx, targetBounds.MinY, cz, ref anyInside, ref anyOutside);
                TestCorner(zone.Planes, cx, targetBounds.MaxY, cz, ref anyInside, ref anyOutside);
                TestCorner(zone.Planes, cx, cy, targetBounds.MinZ, ref anyInside, ref anyOutside);
                TestCorner(zone.Planes, cx, cy, targetBounds.MaxZ, ref anyInside, ref anyOutside);
            }

            if (!anyInside)
            {
                return null;
            }

            var isPartial = anyInside && anyOutside;
            if (isPartial && !needsPartial)
            {
                return null;
            }

            double? fraction = null;
            if (computeContainmentFraction)
            {
                fraction = isPartial
                    ? ComputeFractionFromOverlap(EstimateOverlap(zoneBounds, targetBounds), targetBounds)
                    : 1.0;
            }

            return new ZoneTargetIntersection
            {
                ZoneId = zone.ZoneId,
                TargetItemKey = targetKey,
                IsContained = !isPartial || treatPartialAsContained,
                IsPartial = isPartial,
                OverlapVolume = EstimateOverlap(zoneBounds, targetBounds),
                ContainmentFraction = fraction
            };
        }

        internal static void TestCorner(IReadOnlyList<PlaneEquation> planes, double x, double y, double z, ref bool anyInside, ref bool anyOutside)
        {
            if (anyInside && anyOutside)
            {
                return;
            }

            var inside = GeometryMath.IsInside(planes, new Vector3D(x, y, z));
            if (inside) anyInside = true; else anyOutside = true;
        }

        internal static bool ZoneContainsPoint(
            ZoneGeometry zone,
            in Aabb zoneBounds,
            SpaceMapperZoneBoundsMode zoneBoundsMode,
            SpaceMapperZoneContainmentEngine containmentEngine,
            SpaceMapperEngineDiagnostics diagnostics,
            double x,
            double y,
            double z)
        {
            if (zone == null)
            {
                return false;
            }

            if (containmentEngine == SpaceMapperZoneContainmentEngine.MeshAccurate
                && zone.HasTriangleMesh
                && zone.TriangleVertices != null)
            {
                var point = new Vector3D(x, y, z);
                bool inside;
                var ok = zone.MeshIsClosed
                    ? GeometryMath.TryIsInsideMesh(zone.TriangleVertices, point, out inside)
                    : GeometryMath.TryIsInsideMeshRobust(zone.TriangleVertices, point, out inside);

                if (ok)
                {
                    if (diagnostics != null)
                    {
                        Interlocked.Increment(ref diagnostics.MeshPointTests);
                    }
                    return inside;
                }

                if (diagnostics != null)
                {
                    Interlocked.Increment(ref diagnostics.MeshFallbackPointTests);
                }
            }
            else if (containmentEngine == SpaceMapperZoneContainmentEngine.MeshAccurate)
            {
                if (diagnostics != null)
                {
                    Interlocked.Increment(ref diagnostics.MeshFallbackPointTests);
                }
            }

            if (zoneBoundsMode != SpaceMapperZoneBoundsMode.Aabb
                && zone?.Planes != null
                && zone.Planes.Count > 0)
            {
                if (diagnostics != null)
                {
                    Interlocked.Increment(ref diagnostics.BoundsPointTests);
                }
                return GeometryMath.IsInside(zone.Planes, new Vector3D(x, y, z));
            }

            if (diagnostics != null)
            {
                Interlocked.Increment(ref diagnostics.BoundsPointTests);
            }
            return x >= zoneBounds.MinX && x <= zoneBounds.MaxX
                && y >= zoneBounds.MinY && y <= zoneBounds.MaxY
                && z >= zoneBounds.MinZ && z <= zoneBounds.MaxZ;
        }

        internal static (Aabb[] Bounds, string[] Keys) BuildTargetBounds(
            IReadOnlyList<TargetGeometry> targets,
            SpaceMapperTargetBoundsMode targetBoundsMode,
            SpaceMapperMidpointMode midpointMode)
        {
            var bounds = new List<Aabb>(targets.Count);
            var keys = new List<string>(targets.Count);
            var useMidpoint = targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint;

            for (int i = 0; i < targets.Count; i++)
            {
                var target = targets[i];
                var bbox = target?.BoundingBox;
                if (bbox == null) continue;
                bounds.Add(useMidpoint ? ToMidpointAabb(bbox, midpointMode) : ToAabb(bbox));
                keys.Add(target.ItemKey);
            }

            return (bounds.ToArray(), keys.ToArray());
        }

        internal static Aabb ToAabb(BoundingBox3D bbox)
        {
            var min = bbox.Min;
            var max = bbox.Max;
            return new Aabb(min.X, min.Y, min.Z, max.X, max.Y, max.Z);
        }

        internal static Aabb ToMidpointAabb(BoundingBox3D bbox, SpaceMapperMidpointMode mode)
        {
            var min = bbox.Min;
            var max = bbox.Max;
            var cx = (min.X + max.X) * 0.5;
            var cy = (min.Y + max.Y) * 0.5;
            var cz = mode == SpaceMapperMidpointMode.BoundingBoxBottomCenter
                ? min.Z
                : (min.Z + max.Z) * 0.5;
            return new Aabb(cx, cy, cz, cx, cy, cz);
        }

        internal static Aabb ComputeWorldBounds(Aabb[] targets)
        {
            var hasAny = false;
            double minX = 0, minY = 0, minZ = 0, maxX = 0, maxY = 0, maxZ = 0;

            foreach (var t in targets)
            {
                if (!hasAny)
                {
                    minX = t.MinX; minY = t.MinY; minZ = t.MinZ;
                    maxX = t.MaxX; maxY = t.MaxY; maxZ = t.MaxZ;
                    hasAny = true;
                    continue;
                }

                minX = Math.Min(minX, t.MinX);
                minY = Math.Min(minY, t.MinY);
                minZ = Math.Min(minZ, t.MinZ);
                maxX = Math.Max(maxX, t.MaxX);
                maxY = Math.Max(maxY, t.MaxY);
                maxZ = Math.Max(maxZ, t.MaxZ);
            }

            return hasAny ? new Aabb(minX, minY, minZ, maxX, maxY, maxZ) : new Aabb(0, 0, 0, 0, 0, 0);
        }

        internal static TimeSpan ToTimeSpan(long ticks)
        {
            if (ticks <= 0)
            {
                return TimeSpan.Zero;
            }

            var seconds = ticks / (double)Stopwatch.Frequency;
            return TimeSpan.FromSeconds(seconds);
        }

        internal static bool IsBetterCandidate(
            in BestZoneCandidate candidate,
            in BestZoneCandidate current,
            SpaceMapperZoneResolutionStrategy strategy)
        {
            if (candidate.IsContained != current.IsContained)
            {
                return candidate.IsContained;
            }

            switch (strategy)
            {
                case SpaceMapperZoneResolutionStrategy.FirstMatch:
                    return candidate.ZoneIndex < current.ZoneIndex;
                case SpaceMapperZoneResolutionStrategy.LargestOverlap:
                    if (candidate.OverlapVolume > current.OverlapVolume)
                    {
                        return true;
                    }
                    if (candidate.OverlapVolume < current.OverlapVolume)
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

            return candidate.ZoneIndex < current.ZoneIndex;
        }

        internal static SpaceMapperZoneBoundsMode ResolveZoneBoundsMode(SpaceMapperProcessingSettings settings)
        {
            if (settings == null)
            {
                return SpaceMapperZoneBoundsMode.Aabb;
            }

            return settings.ZoneBoundsMode;
        }

        internal static SpaceMapperZoneContainmentEngine ResolveZoneContainmentEngine(SpaceMapperProcessingSettings settings)
        {
            if (settings == null)
            {
                return SpaceMapperZoneContainmentEngine.BoundsFast;
            }

            return settings.ZoneContainmentEngine;
        }

        internal static SpaceMapperContainmentCalculationMode ResolveContainmentCalculationMode(SpaceMapperProcessingSettings settings)
        {
            if (settings == null)
            {
                return SpaceMapperContainmentCalculationMode.Auto;
            }

            return settings.ContainmentCalculationMode;
        }

        internal static SpaceMapperZoneResolutionStrategy ResolveZoneResolutionStrategy(SpaceMapperProcessingSettings settings)
        {
            if (settings == null)
            {
                return SpaceMapperZoneResolutionStrategy.MostSpecific;
            }

            return settings.ZoneResolutionStrategy;
        }

        internal static SpaceMapperPerformancePreset InferPresetFromBounds(
            SpaceMapperZoneBoundsMode zoneMode,
            SpaceMapperTargetBoundsMode targetMode,
            SpaceMapperZoneContainmentEngine containmentEngine)
        {
            if (containmentEngine == SpaceMapperZoneContainmentEngine.MeshAccurate)
            {
                return SpaceMapperPerformancePreset.Accurate;
            }

            if (targetMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                return SpaceMapperPerformancePreset.Fast;
            }

            if (zoneMode == SpaceMapperZoneBoundsMode.Hull
                || targetMode == SpaceMapperTargetBoundsMode.Hull
                || zoneMode == SpaceMapperZoneBoundsMode.KDop
                || targetMode == SpaceMapperTargetBoundsMode.KDop
                || zoneMode == SpaceMapperZoneBoundsMode.Obb
                || targetMode == SpaceMapperTargetBoundsMode.Obb)
            {
                return SpaceMapperPerformancePreset.Accurate;
            }

            return SpaceMapperPerformancePreset.Normal;
        }

        internal static SpaceMapperFastTraversalMode ResolveFastTraversal(
            SpaceMapperFastTraversalMode requested,
            bool targetIsPoint,
            bool needsPartial,
            int targetCount,
            int zoneCount)
        {
            var allowTargetMajor = targetIsPoint && !needsPartial;
            if (!allowTargetMajor)
            {
                return SpaceMapperFastTraversalMode.ZoneMajor;
            }

            if (requested == SpaceMapperFastTraversalMode.TargetMajor)
            {
                return SpaceMapperFastTraversalMode.TargetMajor;
            }

            if (requested == SpaceMapperFastTraversalMode.ZoneMajor)
            {
                return SpaceMapperFastTraversalMode.ZoneMajor;
            }

            return targetCount > zoneCount * 2
                ? SpaceMapperFastTraversalMode.TargetMajor
                : SpaceMapperFastTraversalMode.ZoneMajor;
        }

        internal static void UpdateMax(ref int current, int candidate)
        {
            int initial;
            do
            {
                initial = current;
                if (candidate <= initial)
                {
                    return;
                }
            }
            while (Interlocked.CompareExchange(ref current, candidate, initial) != initial);
        }

        internal static double EstimateOverlap(in Aabb zoneBounds, in Aabb targetBounds)
        {
            var dx = Math.Max(0, Math.Min(zoneBounds.MaxX, targetBounds.MaxX) - Math.Max(zoneBounds.MinX, targetBounds.MinX));
            var dy = Math.Max(0, Math.Min(zoneBounds.MaxY, targetBounds.MaxY) - Math.Max(zoneBounds.MinY, targetBounds.MinY));
            var dz = Math.Max(0, Math.Min(zoneBounds.MaxZ, targetBounds.MaxZ) - Math.Max(zoneBounds.MinZ, targetBounds.MinZ));
            return dx * dy * dz;
        }

        internal static double? ComputeFractionFromOverlap(double overlapVolume, in Aabb targetBounds)
        {
            var targetVolume = targetBounds.SizeX * targetBounds.SizeY * targetBounds.SizeZ;
            if (targetVolume <= 0)
            {
                return null;
            }

            var fraction = overlapVolume / targetVolume;
            if (double.IsNaN(fraction) || double.IsInfinity(fraction))
            {
                return null;
            }

            if (fraction < 0) return 0;
            if (fraction > 1) return 1;
            return fraction;
        }

        private const int SamplePointsFast = 3;
        private const int SamplePointsDense = 5;
        private const int TargetGeometrySampleLimit = 200;

        internal static double? ComputeContainmentFraction(
            ZoneGeometry zone,
            in Aabb zoneBounds,
            in Aabb targetBounds,
            SpaceMapperZoneBoundsMode zoneBoundsMode,
            SpaceMapperZoneContainmentEngine containmentEngine,
            SpaceMapperContainmentCalculationMode calculationMode,
            IReadOnlyList<Vector3D> targetSamplePoints,
            SpaceMapperEngineDiagnostics diagnostics)
        {
            if (zone == null)
            {
                return null;
            }

            switch (calculationMode)
            {
                case SpaceMapperContainmentCalculationMode.BoundsOverlap:
                    return ComputeFractionFromOverlap(EstimateOverlap(zoneBounds, targetBounds), targetBounds);
                case SpaceMapperContainmentCalculationMode.SamplePoints:
                    return ComputeSampleFractionFromBounds(
                        zone,
                        zoneBounds,
                        targetBounds,
                        zoneBoundsMode,
                        containmentEngine,
                        diagnostics,
                        SamplePointsFast);
                case SpaceMapperContainmentCalculationMode.SamplePointsDense:
                    return ComputeSampleFractionFromBounds(
                        zone,
                        zoneBounds,
                        targetBounds,
                        zoneBoundsMode,
                        containmentEngine,
                        diagnostics,
                        SamplePointsDense);
                case SpaceMapperContainmentCalculationMode.TargetGeometry:
                {
                    var fraction = ComputeSampleFractionFromPoints(
                        zone,
                        zoneBounds,
                        zoneBoundsMode,
                        containmentEngine,
                        diagnostics,
                        targetSamplePoints);
                    if (fraction.HasValue)
                    {
                        return fraction;
                    }

                    return ComputeSampleFractionFromBounds(
                        zone,
                        zoneBounds,
                        targetBounds,
                        zoneBoundsMode,
                        containmentEngine,
                        diagnostics,
                        SamplePointsFast);
                }
                default:
                    return null;
            }
        }

        internal static double? ComputeSampleFractionFromBounds(
            ZoneGeometry zone,
            in Aabb zoneBounds,
            in Aabb targetBounds,
            SpaceMapperZoneBoundsMode zoneBoundsMode,
            SpaceMapperZoneContainmentEngine containmentEngine,
            SpaceMapperEngineDiagnostics diagnostics,
            int samplesPerAxis)
        {
            if (samplesPerAxis <= 0)
            {
                return null;
            }

            var steps = samplesPerAxis - 1;
            if (steps <= 0)
            {
                var cx = (targetBounds.MinX + targetBounds.MaxX) * 0.5;
                var cy = (targetBounds.MinY + targetBounds.MaxY) * 0.5;
                var cz = (targetBounds.MinZ + targetBounds.MaxZ) * 0.5;
                var inside = ZoneContainsPoint(zone, zoneBounds, zoneBoundsMode, containmentEngine, diagnostics, cx, cy, cz);
                return inside ? 1.0 : 0.0;
            }

            var insideCount = 0;
            var sampleCount = 0;
            for (int xi = 0; xi <= steps; xi++)
            {
                var tx = targetBounds.MinX + (targetBounds.MaxX - targetBounds.MinX) * (xi / (double)steps);
                for (int yi = 0; yi <= steps; yi++)
                {
                    var ty = targetBounds.MinY + (targetBounds.MaxY - targetBounds.MinY) * (yi / (double)steps);
                    for (int zi = 0; zi <= steps; zi++)
                    {
                        var tz = targetBounds.MinZ + (targetBounds.MaxZ - targetBounds.MinZ) * (zi / (double)steps);
                        sampleCount++;
                        if (ZoneContainsPoint(zone, zoneBounds, zoneBoundsMode, containmentEngine, diagnostics, tx, ty, tz))
                        {
                            insideCount++;
                        }
                    }
                }
            }

            if (sampleCount == 0)
            {
                return null;
            }

            var fraction = insideCount / (double)sampleCount;
            if (fraction < 0) return 0;
            if (fraction > 1) return 1;
            return fraction;
        }

        internal static double? ComputeSampleFractionFromPoints(
            ZoneGeometry zone,
            in Aabb zoneBounds,
            SpaceMapperZoneBoundsMode zoneBoundsMode,
            SpaceMapperZoneContainmentEngine containmentEngine,
            SpaceMapperEngineDiagnostics diagnostics,
            IReadOnlyList<Vector3D> samplePoints)
        {
            if (samplePoints == null || samplePoints.Count == 0)
            {
                return null;
            }

            var insideCount = 0;
            for (int i = 0; i < samplePoints.Count; i++)
            {
                var p = samplePoints[i];
                if (ZoneContainsPoint(zone, zoneBounds, zoneBoundsMode, containmentEngine, diagnostics, p.X, p.Y, p.Z))
                {
                    insideCount++;
                }
            }

            var fraction = insideCount / (double)samplePoints.Count;
            if (fraction < 0) return 0;
            if (fraction > 1) return 1;
            return fraction;
        }

        internal static IReadOnlyList<Vector3D>[] BuildTargetGeometrySamples(IReadOnlyList<TargetGeometry> targets)
        {
            if (targets == null || targets.Count == 0)
            {
                return null;
            }

            var samples = new IReadOnlyList<Vector3D>[targets.Count];
            for (int i = 0; i < targets.Count; i++)
            {
                var target = targets[i];
                var vertices = target?.TriangleVertices;
                if (vertices == null || vertices.Count < 3)
                {
                    continue;
                }

                var triangleCount = vertices.Count / 3;
                if (triangleCount == 0)
                {
                    continue;
                }

                var stride = Math.Max(1, triangleCount / TargetGeometrySampleLimit);
                var list = new List<Vector3D>(Math.Min(triangleCount, TargetGeometrySampleLimit));
                for (int tri = 0; tri < triangleCount; tri += stride)
                {
                    var idx = tri * 3;
                    if (idx + 2 >= vertices.Count)
                    {
                        break;
                    }

                    var p1 = vertices[idx];
                    var p2 = vertices[idx + 1];
                    var p3 = vertices[idx + 2];
                    list.Add(new Vector3D(
                        (p1.X + p2.X + p3.X) / 3.0,
                        (p1.Y + p2.Y + p3.Y) / 3.0,
                        (p1.Z + p2.Z + p3.Z) / 3.0));

                    if (list.Count >= TargetGeometrySampleLimit)
                    {
                        break;
                    }
                }

                if (list.Count > 0)
                {
                    samples[i] = list;
                }
            }

            return samples;
        }

        internal static Aabb GetZoneQueryBounds(ZoneGeometry zone, SpaceMapperProcessingSettings settings)
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

        internal static Aabb Inflate(Aabb bbox, SpaceMapperProcessingSettings settings)
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
    }

    internal class CudaIntersectionEngine : ISpatialIntersectionEngine
    {
        public SpaceMapperProcessingMode Mode { get; }
        private readonly CpuIntersectionEngine _cpuFallback;
        private const int GpuMinPointsQuick = 64;
        private const int GpuMinPointsIntensive = 32;
        private const int GpuMinPointsPack = 16;
        private const long GpuMinWorkQuick = 20000;
        private const long GpuMinWorkIntensive = 10000;
        private const long GpuMinWorkPack = 5000;
        private const int DefaultMaxBatchZones = 32;
        private const int DefaultMaxBatchPoints = 200000;
        private const int DefaultMaxBatchTriangles = 250000;
        private const int CudaBvhLeafSize = 8;
        private const int OpenMeshBoundaryEdgeLimit = 32;
        private const int OpenMeshNonManifoldEdgeLimit = 8;
        private const int OpenMeshOutsideSampleTolerance = 1;
        private const double OpenMeshPointNudgeRatio = 1e-6;
        private const double OpenMeshPointNudgeMin = 1e-4;
        private const double OpenMeshPointNudgeMax = 1.0;

        public CudaIntersectionEngine(SpaceMapperProcessingMode mode)
        {
            Mode = mode;
            _cpuFallback = new CpuIntersectionEngine(SpaceMapperProcessingMode.CpuNormal);
        }

        public IList<ZoneTargetIntersection> ComputeIntersections(
            IReadOnlyList<ZoneGeometry> zones,
            IReadOnlyList<TargetGeometry> targets,
            SpaceMapperProcessingSettings settings,
            SpaceMapperPreflightCache preflightCache,
            SpaceMapperEngineDiagnostics diagnostics,
            IProgress<SpaceMapperProgress> progress = null,
            CancellationToken cancellationToken = default,
            SpaceMapperRunProgressState runProgress = null)
        {
            var results = new List<ZoneTargetIntersection>();
            if (zones == null || targets == null || zones.Count == 0 || targets.Count == 0)
            {
                return results;
            }

            diagnostics ??= new SpaceMapperEngineDiagnostics();
            var containmentEngine = CpuIntersectionEngine.ResolveZoneContainmentEngine(settings);
            if (containmentEngine != SpaceMapperZoneContainmentEngine.MeshAccurate)
            {
                return _cpuFallback.ComputeIntersections(zones, targets, settings, preflightCache, diagnostics, progress, cancellationToken, runProgress);
            }

            string cudaReason = null;
            string d3dReason = null;
            CudaPointInMeshGpu cudaBackend = null;
            D3D11PointInMeshGpu d3dBackend = null;

            if (CudaPointInMeshGpu.TryCreate(out var cudaCreated, out cudaReason))
            {
                cudaBackend = cudaCreated;
            }

            if (D3D11PointInMeshGpu.TryCreate(out var d3dCreated, out d3dReason))
            {
                d3dBackend = d3dCreated;
            }

            if (cudaBackend == null && d3dBackend == null)
            {
                diagnostics.GpuInitFailureReason = BuildGpuInitFailureReason(cudaReason, d3dReason);
                return _cpuFallback.ComputeIntersections(zones, targets, settings, preflightCache, diagnostics, progress, cancellationToken, runProgress);
            }

            var primaryGpu = (IPointInMeshGpuBackend)d3dBackend ?? cudaBackend;
            diagnostics.GpuBackend = primaryGpu.BackendName;
            diagnostics.GpuAdapterName = primaryGpu.DeviceName;
            if (d3dBackend != null)
            {
                var adapter = d3dBackend.GpuAdapter;
                if (adapter != null)
                {
                    diagnostics.GpuAdapterName = adapter.Description;
                    diagnostics.GpuAdapterLuid = adapter.Luid;
                    diagnostics.GpuVendorId = adapter.VendorId;
                    diagnostics.GpuDeviceId = adapter.DeviceId;
                    diagnostics.GpuSubSysId = adapter.SubsystemId;
                    diagnostics.GpuRevision = adapter.Revision;
                    diagnostics.GpuDedicatedVideoMemory = adapter.DedicatedVideoMemory;
                    diagnostics.GpuDedicatedSystemMemory = adapter.DedicatedSystemMemory;
                    diagnostics.GpuSharedSystemMemory = adapter.SharedSystemMemory;
                    diagnostics.GpuFeatureLevel = adapter.FeatureLevel;
                }
            }

            try
            {
                return ComputeZoneMajorGpu(
                    zones,
                    targets,
                    settings,
                    preflightCache,
                    diagnostics,
                    progress,
                    cancellationToken,
                    runProgress,
                    d3dBackend,
                    cudaBackend);
            }
            finally
            {
                d3dBackend?.Dispose();
                cudaBackend?.Dispose();
            }
        }

        private IList<ZoneTargetIntersection> ComputeZoneMajorGpu(
            IReadOnlyList<ZoneGeometry> zones,
            IReadOnlyList<TargetGeometry> targets,
            SpaceMapperProcessingSettings settings,
            SpaceMapperPreflightCache preflightCache,
            SpaceMapperEngineDiagnostics diagnostics,
            IProgress<SpaceMapperProgress> progress,
            CancellationToken cancellationToken,
            SpaceMapperRunProgressState runProgress,
            D3D11PointInMeshGpu d3dGpu,
            CudaPointInMeshGpu cudaGpu)
        {
            var results = new List<ZoneTargetIntersection>();
            diagnostics ??= new SpaceMapperEngineDiagnostics();
            diagnostics.GpuZoneDiagnostics ??= new List<SpaceMapperGpuZoneDiagnostic>(zones?.Count ?? 0);
            var gpuZoneDiagnostics = diagnostics.GpuZoneDiagnostics;
            var primaryGpu = (IPointInMeshGpuBackend)d3dGpu ?? cudaGpu;
            if (primaryGpu == null)
            {
                return results;
            }

            var containmentEngine = CpuIntersectionEngine.ResolveZoneContainmentEngine(settings);
            var resolutionStrategy = CpuIntersectionEngine.ResolveZoneResolutionStrategy(settings);
            var zoneBoundsMode = CpuIntersectionEngine.ResolveZoneBoundsMode(settings);
            var targetBoundsMode = SpaceMapperBoundsResolver.ResolveTargetBoundsMode(settings, containmentEngine);
            var targetMidpointMode = settings?.TargetMidpointMode ?? SpaceMapperMidpointMode.BoundingBoxCenter;
            var containmentCalculationMode = CpuIntersectionEngine.ResolveContainmentCalculationMode(settings);
            var needsFraction = settings != null
                && (settings.WriteZoneContainmentPercentProperty
                    || (settings.WriteZoneBehaviorProperty
                        && containmentCalculationMode != SpaceMapperContainmentCalculationMode.Auto));
            var computeContainmentFraction = needsFraction
                && containmentCalculationMode == SpaceMapperContainmentCalculationMode.Auto;
            var needsPartial = settings != null
                && (settings.TagPartialSeparately
                    || settings.TreatPartialAsContained
                    || settings.WriteZoneBehaviorProperty
                    || settings.WriteZoneContainmentPercentProperty
                    || needsFraction);
            if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                needsPartial = false;
                computeContainmentFraction = false;
                needsFraction = false;
            }
            var usePointIndex = targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint;
            var treatPartialAsContained = settings != null && settings.TreatPartialAsContained;
            if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                treatPartialAsContained = false;
            }
            var enableMultiZone = settings == null || settings.EnableMultipleZones;
            diagnostics.PresetUsed = CpuIntersectionEngine.InferPresetFromBounds(zoneBoundsMode, targetBoundsMode, containmentEngine);
            diagnostics.TraversalUsed = SpaceMapperFastTraversalMode.ZoneMajor.ToString();

            SpatialHashGrid grid;
            Aabb[] targetBounds;
            string[] targetKeys;

            var buildSw = Stopwatch.StartNew();
            if (preflightCache?.Grid != null
                && preflightCache.TargetBounds != null
                && preflightCache.TargetKeys != null
                && preflightCache.PointIndex == usePointIndex)
            {
                grid = preflightCache.Grid;
                targetBounds = preflightCache.TargetBounds;
                targetKeys = preflightCache.TargetKeys;
                diagnostics.UsedPreflightIndex = true;
            }
            else
            {
                diagnostics.UsedPreflightIndex = false;
                var prepared = CpuIntersectionEngine.BuildTargetBounds(targets, targetBoundsMode, targetMidpointMode);
                targetBounds = prepared.Bounds;
                targetKeys = prepared.Keys;
                if (targetBounds.Length == 0)
                {
                    buildSw.Stop();
                    diagnostics.BuildIndexTime = buildSw.Elapsed;
                    return results;
                }

                var worldBounds = CpuIntersectionEngine.ComputeWorldBounds(targetBounds);
                var cellSize = SpatialGridSizing.ComputeCellSize(worldBounds, settings?.IndexGranularity ?? 0);
                grid = new SpatialHashGrid(worldBounds, cellSize, targetBounds);
            }
            buildSw.Stop();
            diagnostics.BuildIndexTime = buildSw.Elapsed;

            if (grid == null || targetBounds == null || targetBounds.Length == 0)
            {
                return results;
            }

            IReadOnlyList<Vector3D>[] targetSamplePoints = null;
            if (needsFraction && containmentCalculationMode == SpaceMapperContainmentCalculationMode.TargetGeometry)
            {
                targetSamplePoints = CpuIntersectionEngine.BuildTargetGeometrySamples(targets);
                if (targetSamplePoints == null || targetSamplePoints.Length != targetBounds.Length)
                {
                    targetSamplePoints = null;
                }
            }

            runProgress?.SetStage(
                SpaceMapperRunStage.ComputingIntersections,
                "Computing intersections (GPU)...",
                "Classifying targets into zones...");

            var useBestSelection = usePointIndex && !enableMultiZone;
            Dictionary<int, GpuBestZoneCandidate> bestHits = null;
            if (useBestSelection)
            {
                bestHits = new Dictionary<int, GpuBestZoneCandidate>();
            }

            var zoneCount = zones.Count;
            var zoneBounds = new Aabb[zoneCount];
            var zoneValid = new bool[zoneCount];
            double[] zoneVolumes = null;
            double[] zoneCenterX = null;
            double[] zoneCenterY = null;
            double[] zoneCenterZ = null;

            for (int i = 0; i < zoneCount; i++)
            {
                var zone = zones[i];
                if (zone?.BoundingBox == null)
                {
                    zoneBounds[i] = new Aabb(0, 0, 0, 0, 0, 0);
                    continue;
                }

                zoneBounds[i] = CpuIntersectionEngine.GetZoneQueryBounds(zone, settings);
                zoneValid[i] = zone?.BoundingBox != null || zone?.RawBoundingBox != null;
            }

            if (useBestSelection)
            {
                zoneVolumes = new double[zoneCount];
                zoneCenterX = new double[zoneCount];
                zoneCenterY = new double[zoneCount];
                zoneCenterZ = new double[zoneCount];

                for (int i = 0; i < zoneCount; i++)
                {
                    if (!zoneValid[i])
                    {
                        zoneVolumes[i] = double.PositiveInfinity;
                        continue;
                    }

                    var bounds = zoneBounds[i];
                    zoneVolumes[i] = bounds.SizeX * bounds.SizeY * bounds.SizeZ;
                    zoneCenterX[i] = (bounds.MinX + bounds.MaxX) * 0.5;
                    zoneCenterY[i] = (bounds.MinY + bounds.MaxY) * 0.5;
                    zoneCenterZ[i] = (bounds.MinZ + bounds.MaxZ) * 0.5;
                }
            }

            var visited = new int[targetBounds.Length];
            var stamp = 0;
            var processedZones = 0;
            long candidatePairs = 0;
            int maxCandidates = 0;
            long candidateTicks = 0;
            long narrowTicks = 0;
            long gpuDispatchTicks = 0;
            long gpuReadbackTicks = 0;
            long gpuPointsTested = 0;
            long gpuTrianglesTested = 0;
            int gpuZonesProcessed = 0;
            int gpuZonesEligible = 0;
            int gpuZonesSkippedNoMesh = 0;
            int gpuZonesSkippedMissingTriangles = 0;
            int gpuZonesSkippedOpenMesh = 0;
            int gpuZonesSkippedLowPoints = 0;
            long gpuUncertainPoints = 0;
            int gpuMaxTrianglesPerZone = 0;
            int gpuMaxPointsPerZone = 0;
            int gpuOpenMeshZonesEligible = 0;
            int gpuOpenMeshZonesProcessed = 0;
            double maxOpenMeshNudge = 0;
            var rayCount = ResolveGpuRayCount(settings, Mode);
            var intensive = rayCount >= 2;
            var sampleCountPerTarget = GetSampleCount(targetBoundsMode);
            var minPoints = Mode == SpaceMapperProcessingMode.GpuQuick ? GpuMinPointsQuick : GpuMinPointsIntensive;
            var minWork = Mode == SpaceMapperProcessingMode.GpuQuick ? GpuMinWorkQuick : GpuMinWorkIntensive;
            var minPointsPack = GpuMinPointsPack;
            var minWorkPack = GpuMinWorkPack;
            var maxBatchZones = settings?.BatchSize > 0 ? settings.BatchSize : DefaultMaxBatchZones;
            var maxBatchPoints = DefaultMaxBatchPoints;
            var maxBatchTriangles = DefaultMaxBatchTriangles;

            var supportsBatching = d3dGpu != null || cudaGpu != null;

            var batchTriangles = supportsBatching ? new List<Triangle>() : null;
            var batchPointsLocal = supportsBatching ? new List<Float4>() : null;
            var batchPointsWorld = supportsBatching ? new List<Vector3D>() : null;
            var batchPointZoneIds = supportsBatching ? new List<uint>() : null;
            var batchZoneRanges = supportsBatching ? new List<ZoneRange>() : null;
            var batchJobs = supportsBatching ? new List<GpuZoneJob>() : null;
            var batchIntensive = false;
            var gpuBatchDispatchCount = 0;
            var gpuBatchMaxZones = 0;
            var gpuBatchMaxPoints = 0;
            var gpuBatchMaxTriangles = 0;
            var gpuBatchZonesTotal = 0L;
            var cudaBvhAvailable = true;
            var usedD3d = false;
            var usedCuda = false;
            var usedCudaBvh = false;

            void ProcessZoneCpu(ZoneGeometry zone, in Aabb zoneBoundsLocal, List<int> candidates, int zoneIndex)
            {
                for (int ci = 0; ci < candidates.Count; ci++)
                {
                    var targetIndex = candidates[ci];
                    var hit = CpuIntersectionEngine.ClassifyIntersection(
                        zone,
                        zoneBoundsLocal,
                        targetBounds[targetIndex],
                        targetKeys[targetIndex],
                        zoneBoundsMode,
                        containmentEngine,
                        diagnostics,
                        targetBoundsMode,
                        computeContainmentFraction,
                        needsPartial,
                        treatPartialAsContained);
                    if (hit == null)
                    {
                        continue;
                    }

                    if (needsFraction && containmentCalculationMode != SpaceMapperContainmentCalculationMode.Auto)
                    {
                        hit.ContainmentFraction = CpuIntersectionEngine.ComputeContainmentFraction(
                            zone,
                            zoneBoundsLocal,
                            targetBounds[targetIndex],
                            zoneBoundsMode,
                            containmentEngine,
                            containmentCalculationMode,
                            targetSamplePoints?[targetIndex],
                            diagnostics);
                    }

                    AddHit(results, bestHits, hit, targetBounds[targetIndex], targetIndex, zoneIndex, zoneVolumes, zoneCenterX, zoneCenterY, zoneCenterZ, resolutionStrategy);
                }
            }

            void ProcessZoneGpuSingle(
                int zoneIndex,
                ZoneGeometry zone,
                in Aabb zoneBoundsLocal,
                List<int> candidates,
                bool useOpenMeshRetry,
                double openMeshNudge,
                bool intensiveForZone)
            {
                var originX = (zoneBoundsLocal.MinX + zoneBoundsLocal.MaxX) * 0.5;
                var originY = (zoneBoundsLocal.MinY + zoneBoundsLocal.MaxY) * 0.5;
                var originZ = (zoneBoundsLocal.MinZ + zoneBoundsLocal.MaxZ) * 0.5;

                var trianglesLocal = BuildTriangles(zone, originX, originY, originZ);
                if (trianglesLocal == null || trianglesLocal.Length == 0)
                {
                    ProcessZoneCpu(zone, zoneBoundsLocal, candidates, zoneIndex);
                    return;
                }

                var estimatedPoints = candidates.Count * sampleCountPerTarget;
                var pointsLocal = new Float4[estimatedPoints];
                var pointsWorld = new Vector3D[estimatedPoints];
                var targetStart = new int[candidates.Count];
                var targetCount = new int[candidates.Count];
                var cursor = 0;

                for (int ci = 0; ci < candidates.Count; ci++)
                {
                    var targetIndex = candidates[ci];
                    targetStart[ci] = cursor;

                    if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
                    {
                        var bounds = targetBounds[targetIndex];
                        var cx = (bounds.MinX + bounds.MaxX) * 0.5;
                        var cy = (bounds.MinY + bounds.MaxY) * 0.5;
                        var cz = (bounds.MinZ + bounds.MaxZ) * 0.5;
                        var worldPoint = new Vector3D(cx, cy, cz);
                        pointsWorld[cursor] = worldPoint;
                        var gpuPoint = useOpenMeshRetry
                            ? ApplyPointNudge(worldPoint, openMeshNudge, BuildNudgeSeed(zoneIndex, targetIndex, 0))
                            : worldPoint;
                        pointsLocal[cursor] = new Float4(
                            (float)(gpuPoint.X - originX),
                            (float)(gpuPoint.Y - originY),
                            (float)(gpuPoint.Z - originZ));
                        targetCount[ci] = 1;
                        cursor++;
                    }
                    else
                    {
                        var samples = BuildBoundsSamplePoints(targetBounds[targetIndex], targetBoundsMode);
                        targetCount[ci] = samples.Count;
                        for (int si = 0; si < samples.Count; si++)
                        {
                            var p = samples[si];
                            pointsWorld[cursor] = p;
                            var gpuPoint = useOpenMeshRetry
                                ? ApplyPointNudge(p, openMeshNudge, BuildNudgeSeed(zoneIndex, targetIndex, si))
                                : p;
                            pointsLocal[cursor] = new Float4(
                                (float)(gpuPoint.X - originX),
                                (float)(gpuPoint.Y - originY),
                                (float)(gpuPoint.Z - originZ));
                            cursor++;
                        }
                    }
                }

                var trimPoints = cursor;
                if (trimPoints <= 0)
                {
                    return;
                }

                if (trimPoints != pointsLocal.Length)
                {
                    Array.Resize(ref pointsLocal, trimPoints);
                    Array.Resize(ref pointsWorld, trimPoints);
                }

                var insideFlags = primaryGpu.TestPoints(
                    trianglesLocal,
                    pointsLocal,
                    intensiveForZone,
                    cancellationToken);

                if (primaryGpu is D3D11PointInMeshGpu d3dTiming)
                {
                    usedD3d = true;
                    gpuDispatchTicks += d3dTiming.LastDispatchTime.Ticks;
                    gpuReadbackTicks += d3dTiming.LastReadbackTime.Ticks;
                }
                else if (primaryGpu is CudaPointInMeshGpu)
                {
                    usedCuda = true;
                }
                gpuPointsTested += trimPoints;
                gpuTrianglesTested += trianglesLocal.Length;
                gpuZonesProcessed++;
                if (useOpenMeshRetry)
                {
                    gpuOpenMeshZonesProcessed++;
                }
                CpuIntersectionEngine.UpdateMax(ref gpuMaxTrianglesPerZone, trianglesLocal.Length);
                CpuIntersectionEngine.UpdateMax(ref gpuMaxPointsPerZone, trimPoints);

                for (int ci = 0; ci < candidates.Count; ci++)
                {
                    var targetIndex = candidates[ci];
                    var start = targetStart[ci];
                    var count = targetCount[ci];
                    if (count <= 0)
                    {
                        continue;
                    }

                    var insideCount = 0;
                    var anyInside = false;
                    var anyOutside = false;
                    for (int si = 0; si < count; si++)
                    {
                        var idx = start + si;
                        var flag = insideFlags[idx];
                        if (flag == D3D11PointInMeshGpu.Uncertain)
                        {
                            gpuUncertainPoints++;
                            var p = pointsWorld[idx];
                            var inside = CpuIntersectionEngine.ZoneContainsPoint(
                                zone,
                                zoneBoundsLocal,
                                zoneBoundsMode,
                                containmentEngine,
                                diagnostics,
                                p.X,
                                p.Y,
                                p.Z);
                            flag = inside ? D3D11PointInMeshGpu.Inside : D3D11PointInMeshGpu.Outside;
                        }

                        if (flag == D3D11PointInMeshGpu.Inside)
                        {
                            insideCount++;
                            anyInside = true;
                        }
                        else
                        {
                            anyOutside = true;
                        }
                    }

                    if (!anyInside)
                    {
                        continue;
                    }

                    var outsideCount = count - insideCount;
                    if (useOpenMeshRetry && outsideCount > 0 && outsideCount <= OpenMeshOutsideSampleTolerance)
                    {
                        anyOutside = false;
                    }
                    var isPartial = anyOutside;
                    if (isPartial && !needsPartial)
                    {
                        continue;
                    }

                    double? fraction = null;
                    if (computeContainmentFraction && count > 0)
                    {
                        fraction = insideCount / (double)count;
                        if (fraction < 0) fraction = 0;
                        else if (fraction > 1) fraction = 1;
                    }

                    var hit = new ZoneTargetIntersection
                    {
                        ZoneId = zone.ZoneId,
                        TargetItemKey = targetKeys[targetIndex],
                        IsContained = !isPartial || treatPartialAsContained,
                        IsPartial = isPartial,
                        OverlapVolume = CpuIntersectionEngine.EstimateOverlap(zoneBoundsLocal, targetBounds[targetIndex]),
                        ContainmentFraction = fraction
                    };

                    if (needsFraction && containmentCalculationMode != SpaceMapperContainmentCalculationMode.Auto)
                    {
                        hit.ContainmentFraction = CpuIntersectionEngine.ComputeContainmentFraction(
                            zone,
                            zoneBoundsLocal,
                            targetBounds[targetIndex],
                            zoneBoundsMode,
                            containmentEngine,
                            containmentCalculationMode,
                            targetSamplePoints?[targetIndex],
                            diagnostics);
                    }

                    AddHit(results, bestHits, hit, targetBounds[targetIndex], targetIndex, zoneIndex, zoneVolumes, zoneCenterX, zoneCenterY, zoneCenterZ, resolutionStrategy);
                }
            }

            void FlushBatch()
            {
                if (!supportsBatching || batchJobs.Count == 0)
                {
                    return;
                }

                var batchStart = Stopwatch.GetTimestamp();
                var trianglesAll = batchTriangles.ToArray();
                var pointsAll = batchPointsLocal.ToArray();
                var pointZoneIds = batchPointZoneIds.ToArray();
                var zoneRanges = batchZoneRanges.ToArray();
                uint[] insideFlags = null;
                TimeSpan dispatchTime = TimeSpan.Zero;
                TimeSpan readbackTime = TimeSpan.Zero;
                if (cudaBvhAvailable)
                {
                    if (CudaBvhPointInMeshGpu.TryCreateScene(
                        trianglesAll,
                        zoneRanges,
                        CudaBvhLeafSize,
                        out var scene,
                        out _))
                    {
                        using (scene)
                        {
                            var callStart = Stopwatch.GetTimestamp();
                            insideFlags = scene.TestPoints(pointsAll, pointZoneIds, batchIntensive, cancellationToken);
                            var elapsed = TimeSpan.FromSeconds((Stopwatch.GetTimestamp() - callStart) / (double)Stopwatch.Frequency);
                            readbackTime = elapsed;
                        }

                        usedCudaBvh = true;
                    }
                    else
                    {
                        cudaBvhAvailable = false;
                    }
                }

                if (insideFlags == null && d3dGpu != null)
                {
                    insideFlags = d3dGpu.TestPointsBatched(
                        trianglesAll,
                        pointsAll,
                        pointZoneIds,
                        zoneRanges,
                        batchIntensive,
                        cancellationToken,
                        out dispatchTime,
                        out readbackTime);
                    usedD3d = true;
                }
                else if (insideFlags == null && cudaGpu != null)
                {
                    var callStart = Stopwatch.GetTimestamp();
                    insideFlags = cudaGpu.TestPointsBatched(
                        trianglesAll,
                        pointsAll,
                        pointZoneIds,
                        zoneRanges,
                        batchIntensive,
                        cancellationToken);
                    var elapsed = TimeSpan.FromSeconds((Stopwatch.GetTimestamp() - callStart) / (double)Stopwatch.Frequency);
                    readbackTime = elapsed;
                    usedCuda = true;
                }
                else if (insideFlags == null)
                {
                    return;
                }

                gpuDispatchTicks += dispatchTime.Ticks;
                gpuReadbackTicks += readbackTime.Ticks;
                gpuPointsTested += batchPointsLocal.Count;
                gpuTrianglesTested += batchTriangles.Count;
                gpuZonesProcessed += batchJobs.Count;
                gpuBatchDispatchCount++;
                gpuBatchZonesTotal += batchJobs.Count;
                CpuIntersectionEngine.UpdateMax(ref gpuBatchMaxZones, batchJobs.Count);
                CpuIntersectionEngine.UpdateMax(ref gpuBatchMaxPoints, batchPointsLocal.Count);
                CpuIntersectionEngine.UpdateMax(ref gpuBatchMaxTriangles, batchTriangles.Count);

                for (int jobIndex = 0; jobIndex < batchJobs.Count; jobIndex++)
                {
                    var job = batchJobs[jobIndex];
                    var zone = job.Zone;
                    var zoneBoundsLocal = job.ZoneBounds;
                    var candidates = job.CandidateTargets;

                    if (job.UseOpenMeshRetry)
                    {
                        gpuOpenMeshZonesProcessed++;
                    }
                    CpuIntersectionEngine.UpdateMax(ref gpuMaxTrianglesPerZone, job.TrianglesAdded);
                    CpuIntersectionEngine.UpdateMax(ref gpuMaxPointsPerZone, job.PointsAdded);

                    for (int ci = 0; ci < candidates.Length; ci++)
                    {
                        var targetIndex = candidates[ci];
                        var start = job.TargetStartAbs[ci];
                        var count = job.TargetCount[ci];
                        if (count <= 0)
                        {
                            continue;
                        }

                        var insideCount = 0;
                        var anyInside = false;
                        var anyOutside = false;
                        for (int si = 0; si < count; si++)
                        {
                            var idx = start + si;
                            var flag = insideFlags[idx];
                            if (flag == D3D11PointInMeshGpu.Uncertain)
                            {
                                gpuUncertainPoints++;
                                var p = batchPointsWorld[idx];
                                var inside = CpuIntersectionEngine.ZoneContainsPoint(
                                    zone,
                                    zoneBoundsLocal,
                                    zoneBoundsMode,
                                    containmentEngine,
                                    diagnostics,
                                    p.X,
                                    p.Y,
                                    p.Z);
                                flag = inside ? D3D11PointInMeshGpu.Inside : D3D11PointInMeshGpu.Outside;
                            }

                            if (flag == D3D11PointInMeshGpu.Inside)
                            {
                                insideCount++;
                                anyInside = true;
                            }
                            else
                            {
                                anyOutside = true;
                            }
                        }

                        if (!anyInside)
                        {
                            continue;
                        }

                        var outsideCount = count - insideCount;
                        if (job.UseOpenMeshRetry && outsideCount > 0 && outsideCount <= OpenMeshOutsideSampleTolerance)
                        {
                            anyOutside = false;
                        }
                        var isPartial = anyOutside;
                        if (isPartial && !needsPartial)
                        {
                            continue;
                        }

                        double? fraction = null;
                        if (computeContainmentFraction && count > 0)
                        {
                            fraction = insideCount / (double)count;
                            if (fraction < 0) fraction = 0;
                            else if (fraction > 1) fraction = 1;
                        }

                        var hit = new ZoneTargetIntersection
                        {
                            ZoneId = zone.ZoneId,
                            TargetItemKey = targetKeys[targetIndex],
                            IsContained = !isPartial || treatPartialAsContained,
                            IsPartial = isPartial,
                            OverlapVolume = CpuIntersectionEngine.EstimateOverlap(zoneBoundsLocal, targetBounds[targetIndex]),
                            ContainmentFraction = fraction
                        };

                        if (needsFraction && containmentCalculationMode != SpaceMapperContainmentCalculationMode.Auto)
                        {
                            hit.ContainmentFraction = CpuIntersectionEngine.ComputeContainmentFraction(
                                zone,
                                zoneBoundsLocal,
                                targetBounds[targetIndex],
                                zoneBoundsMode,
                                containmentEngine,
                                containmentCalculationMode,
                                targetSamplePoints?[targetIndex],
                                diagnostics);
                        }

                        AddHit(results, bestHits, hit, targetBounds[targetIndex], targetIndex, job.ZoneIndex, zoneVolumes, zoneCenterX, zoneCenterY, zoneCenterZ, resolutionStrategy);
                    }
                }

                narrowTicks += Stopwatch.GetTimestamp() - batchStart;

                batchTriangles.Clear();
                batchPointsLocal.Clear();
                batchPointsWorld.Clear();
                batchPointZoneIds.Clear();
                batchZoneRanges.Clear();
                batchJobs.Clear();
                batchIntensive = false;
            }

            for (int zoneIndex = 0; zoneIndex < zoneCount; zoneIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var zone = zones[zoneIndex];
                if (!zoneValid[zoneIndex])
                {
                    gpuZoneDiagnostics?.Add(new SpaceMapperGpuZoneDiagnostic
                    {
                        ZoneIndex = zoneIndex,
                        ZoneId = zone?.ZoneId,
                        ZoneName = zone?.DisplayName,
                        CandidateCount = 0,
                        EstimatedPoints = 0,
                        TriangleCount = 0,
                        WorkEstimate = 0,
                        PointThreshold = minPoints,
                        WorkThreshold = minWork,
                        HasMesh = false,
                        IsOpenMesh = false,
                        AllowOpenMeshGpu = false,
                        EligibleForGpu = false,
                        UsedGpu = false,
                        UsedOpenMeshRetry = false,
                        PackedThresholds = false,
                        SkipReason = "InvalidZone"
                    });
                    processedZones++;
                    if (runProgress != null && (((processedZones & 63) == 0) || processedZones == zoneCount))
                    {
                        runProgress.SetTotals(zoneCount, targets.Count);
                        runProgress.UpdateZonesProcessed(processedZones);
                        runProgress.UpdateCandidatePairs(candidatePairs);
                    }
                    continue;
                }

                var zoneName = string.IsNullOrWhiteSpace(zone.DisplayName) ? zone.ZoneId : zone.DisplayName;
                runProgress?.UpdateDetail($"Zone {zoneIndex + 1}/{zoneCount}: {zoneName}");

                var zoneBoundsLocal = zoneBounds[zoneIndex];
                var candidates = new List<int>();
                var candidateStart = Stopwatch.GetTimestamp();
                grid.VisitCandidates(zoneBoundsLocal, visited, ref stamp, idx => candidates.Add(idx));
                candidateTicks += Stopwatch.GetTimestamp() - candidateStart;

                var candidateCount = candidates.Count;
                candidatePairs += candidateCount;
                CpuIntersectionEngine.UpdateMax(ref maxCandidates, candidateCount);

                var hasTriangleMesh = zone.HasTriangleMesh;
                var hasTriangleVertices = zone.TriangleVertices != null && zone.TriangleVertices.Count >= 3;
                var isClosed = zone.MeshIsClosed;
                var isOpenMesh = hasTriangleMesh && hasTriangleVertices && !isClosed;
                var allowOpenMeshGpu = isOpenMesh
                    && zone.MeshBoundaryEdgeCount <= OpenMeshBoundaryEdgeLimit
                    && zone.MeshNonManifoldEdgeCount <= OpenMeshNonManifoldEdgeLimit;
                var triCountEstimate = hasTriangleVertices ? zone.TriangleVertices.Count / 3 : 0;

                if (candidateCount == 0)
                {
                    gpuZoneDiagnostics?.Add(new SpaceMapperGpuZoneDiagnostic
                    {
                        ZoneIndex = zoneIndex,
                        ZoneId = zone.ZoneId,
                        ZoneName = zone.DisplayName,
                        CandidateCount = 0,
                        EstimatedPoints = 0,
                        TriangleCount = triCountEstimate,
                        WorkEstimate = 0,
                        PointThreshold = minPoints,
                        WorkThreshold = minWork,
                        HasMesh = hasTriangleMesh,
                        IsOpenMesh = isOpenMesh,
                        AllowOpenMeshGpu = allowOpenMeshGpu,
                        EligibleForGpu = hasTriangleMesh && hasTriangleVertices && (isClosed || allowOpenMeshGpu),
                        UsedGpu = false,
                        UsedOpenMeshRetry = false,
                        PackedThresholds = false,
                        SkipReason = "NoCandidates"
                    });
                    processedZones++;
                    if (runProgress != null && (((processedZones & 63) == 0) || processedZones == zoneCount))
                    {
                        runProgress.SetTotals(zoneCount, targets.Count);
                        runProgress.UpdateZonesProcessed(processedZones);
                        runProgress.UpdateCandidatePairs(candidatePairs);
                    }
                    continue;
                }

                if (!hasTriangleMesh)
                {
                    gpuZonesSkippedNoMesh++;
                }
                else if (!hasTriangleVertices)
                {
                    gpuZonesSkippedMissingTriangles++;
                }
                else if (isOpenMesh && !allowOpenMeshGpu)
                {
                    gpuZonesSkippedOpenMesh++;
                }

                var hasMesh = hasTriangleMesh && hasTriangleVertices && (isClosed || allowOpenMeshGpu);
                if (hasMesh)
                {
                    gpuZonesEligible++;
                    if (allowOpenMeshGpu)
                    {
                        gpuOpenMeshZonesEligible++;
                    }
                }
                var estimatedPoints = candidateCount * sampleCountPerTarget;
                var workEstimate = (long)estimatedPoints * triCountEstimate;
                var intensiveForZone = intensive || allowOpenMeshGpu;
                var canPack = supportsBatching
                    && batchJobs.Count > 0
                    && intensiveForZone == batchIntensive
                    && batchJobs.Count + 1 <= maxBatchZones
                    && batchPointsLocal.Count + estimatedPoints <= maxBatchPoints
                    && batchTriangles.Count + triCountEstimate <= maxBatchTriangles;
                var pointThreshold = canPack ? minPointsPack : minPoints;
                var workThreshold = canPack ? minWorkPack : minWork;
                var belowThreshold = hasMesh && (estimatedPoints < pointThreshold || workEstimate < workThreshold);
                if (belowThreshold)
                {
                    gpuZonesSkippedLowPoints++;
                }
                var useGpuForZone = hasMesh && !belowThreshold;
                var useOpenMeshRetry = useGpuForZone && allowOpenMeshGpu;
                var openMeshNudge = useOpenMeshRetry ? ComputeOpenMeshNudge(zoneBoundsLocal) : 0.0;
                if (openMeshNudge > maxOpenMeshNudge)
                {
                    maxOpenMeshNudge = openMeshNudge;
                }
                intensiveForZone = intensiveForZone || useOpenMeshRetry;
                if (gpuZoneDiagnostics != null)
                {
                    var skipReason = string.Empty;
                    if (!hasTriangleMesh)
                    {
                        skipReason = "NoMesh";
                    }
                    else if (!hasTriangleVertices)
                    {
                        skipReason = "MissingTriangles";
                    }
                    else if (isOpenMesh && !allowOpenMeshGpu)
                    {
                        skipReason = "OpenMesh";
                    }
                    else if (belowThreshold)
                    {
                        skipReason = "BelowThreshold";
                    }

                    gpuZoneDiagnostics.Add(new SpaceMapperGpuZoneDiagnostic
                    {
                        ZoneIndex = zoneIndex,
                        ZoneId = zone.ZoneId,
                        ZoneName = zone.DisplayName,
                        CandidateCount = candidateCount,
                        EstimatedPoints = estimatedPoints,
                        TriangleCount = triCountEstimate,
                        WorkEstimate = workEstimate,
                        PointThreshold = pointThreshold,
                        WorkThreshold = workThreshold,
                        HasMesh = hasTriangleMesh,
                        IsOpenMesh = isOpenMesh,
                        AllowOpenMeshGpu = allowOpenMeshGpu,
                        EligibleForGpu = hasMesh,
                        UsedGpu = useGpuForZone,
                        UsedOpenMeshRetry = useOpenMeshRetry,
                        PackedThresholds = canPack,
                        SkipReason = skipReason
                    });
                }

                var zoneNarrowStart = Stopwatch.GetTimestamp();
                if (!useGpuForZone)
                {
                    ProcessZoneCpu(zone, zoneBoundsLocal, candidates, zoneIndex);
                }
                else if (!supportsBatching)
                {
                    ProcessZoneGpuSingle(zoneIndex, zone, zoneBoundsLocal, candidates, useOpenMeshRetry, openMeshNudge, intensiveForZone);
                }
                else
                {
                    if (batchJobs.Count > 0)
                    {
                        if (intensiveForZone != batchIntensive
                            || batchJobs.Count + 1 > maxBatchZones
                            || batchPointsLocal.Count + estimatedPoints > maxBatchPoints
                            || batchTriangles.Count + triCountEstimate > maxBatchTriangles)
                        {
                            FlushBatch();
                        }
                    }

                    var originX = (zoneBoundsLocal.MinX + zoneBoundsLocal.MaxX) * 0.5;
                    var originY = (zoneBoundsLocal.MinY + zoneBoundsLocal.MaxY) * 0.5;
                    var originZ = (zoneBoundsLocal.MinZ + zoneBoundsLocal.MaxZ) * 0.5;

                    var trianglesLocal = BuildTriangles(zone, originX, originY, originZ);
                    if (trianglesLocal == null || trianglesLocal.Length == 0)
                    {
                        ProcessZoneCpu(zone, zoneBoundsLocal, candidates, zoneIndex);
                    }
                    else
                    {
                        if (batchJobs.Count == 0)
                        {
                            batchIntensive = intensiveForZone;
                        }

                        var zoneBatchId = batchZoneRanges.Count;
                        var triStart = batchTriangles.Count;
                        batchTriangles.AddRange(trianglesLocal);
                        batchZoneRanges.Add(new ZoneRange
                        {
                            TriStart = (uint)triStart,
                            TriCount = (uint)trianglesLocal.Length
                        });

                        var candidateTargets = candidates.ToArray();
                        var targetStartAbs = new int[candidateTargets.Length];
                        var targetCount = new int[candidateTargets.Length];
                        var pointsAdded = 0;

                        for (int ci = 0; ci < candidateTargets.Length; ci++)
                        {
                            var targetIndex = candidateTargets[ci];
                            targetStartAbs[ci] = batchPointsLocal.Count;

                            if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
                            {
                                var bounds = targetBounds[targetIndex];
                                var cx = (bounds.MinX + bounds.MaxX) * 0.5;
                                var cy = (bounds.MinY + bounds.MaxY) * 0.5;
                                var cz = (bounds.MinZ + bounds.MaxZ) * 0.5;
                                var worldPoint = new Vector3D(cx, cy, cz);
                                batchPointsWorld.Add(worldPoint);
                                var gpuPoint = useOpenMeshRetry
                                    ? ApplyPointNudge(worldPoint, openMeshNudge, BuildNudgeSeed(zoneIndex, targetIndex, 0))
                                    : worldPoint;
                                batchPointsLocal.Add(new Float4(
                                    (float)(gpuPoint.X - originX),
                                    (float)(gpuPoint.Y - originY),
                                    (float)(gpuPoint.Z - originZ)));
                                batchPointZoneIds.Add((uint)zoneBatchId);
                                targetCount[ci] = 1;
                                pointsAdded++;
                            }
                            else
                            {
                                var samples = BuildBoundsSamplePoints(targetBounds[targetIndex], targetBoundsMode);
                                targetCount[ci] = samples.Count;
                                for (int si = 0; si < samples.Count; si++)
                                {
                                    var p = samples[si];
                                    batchPointsWorld.Add(p);
                                    var gpuPoint = useOpenMeshRetry
                                        ? ApplyPointNudge(p, openMeshNudge, BuildNudgeSeed(zoneIndex, targetIndex, si))
                                        : p;
                                    batchPointsLocal.Add(new Float4(
                                        (float)(gpuPoint.X - originX),
                                        (float)(gpuPoint.Y - originY),
                                        (float)(gpuPoint.Z - originZ)));
                                    batchPointZoneIds.Add((uint)zoneBatchId);
                                    pointsAdded++;
                                }
                            }
                        }

                        if (pointsAdded > 0)
                        {
                            var job = new GpuZoneJob
                            {
                                ZoneIndex = zoneIndex,
                                Zone = zone,
                                ZoneBounds = zoneBoundsLocal,
                                CandidateTargets = candidateTargets,
                                TargetStartAbs = targetStartAbs,
                                TargetCount = targetCount,
                                UseOpenMeshRetry = useOpenMeshRetry,
                                OpenMeshNudge = openMeshNudge,
                                Intensive = intensiveForZone,
                                PointsAdded = pointsAdded,
                                TrianglesAdded = trianglesLocal.Length
                            };
                            batchJobs.Add(job);

                            CpuIntersectionEngine.UpdateMax(ref gpuMaxTrianglesPerZone, trianglesLocal.Length);
                            CpuIntersectionEngine.UpdateMax(ref gpuMaxPointsPerZone, pointsAdded);
                        }
                    }
                }

                narrowTicks += Stopwatch.GetTimestamp() - zoneNarrowStart;
                processedZones++;
                if (runProgress != null && (((processedZones & 63) == 0) || processedZones == zoneCount))
                {
                    runProgress.SetTotals(zoneCount, targets.Count);
                    runProgress.UpdateZonesProcessed(processedZones);
                    runProgress.UpdateCandidatePairs(candidatePairs);
                }
                if (progress != null && (processedZones % 4 == 0 || processedZones == zoneCount))
                {
                    var pairsDone = candidatePairs;
                    var totalPairs = preflightCache?.LastResult?.CandidatePairs ?? 0;
                    progress.Report(new SpaceMapperProgress
                    {
                        ProcessedPairs = pairsDone > int.MaxValue ? int.MaxValue : (int)pairsDone,
                        TotalPairs = totalPairs > int.MaxValue ? int.MaxValue : (int)totalPairs,
                        ZonesProcessed = processedZones,
                        TargetsProcessed = targetBounds.Length
                    });
                }
            }

            FlushBatch();

            if (useBestSelection && bestHits != null)
            {
                foreach (var hit in bestHits.Values)
                {
                    if (hit.Intersection != null)
                    {
                        results.Add(hit.Intersection);
                    }
                }
            }

            diagnostics.CandidatePairs = candidatePairs;
            diagnostics.MaxCandidatesPerZone = maxCandidates;
            diagnostics.AvgCandidatesPerZone = zoneCount == 0 ? 0 : candidatePairs / (double)zoneCount;
            diagnostics.MaxCandidatesPerTarget = 0;
            diagnostics.AvgCandidatesPerTarget = 0;
            diagnostics.CandidateQueryTime = TimeSpan.FromSeconds(candidateTicks / (double)Stopwatch.Frequency);
            diagnostics.NarrowPhaseTime = TimeSpan.FromSeconds(narrowTicks / (double)Stopwatch.Frequency);
            diagnostics.GpuZonesProcessed = gpuZonesProcessed;
            diagnostics.GpuPointsTested = gpuPointsTested;
            diagnostics.GpuTrianglesTested = gpuTrianglesTested;
            diagnostics.GpuDispatchTime = TimeSpan.FromTicks(gpuDispatchTicks);
            diagnostics.GpuReadbackTime = TimeSpan.FromTicks(gpuReadbackTicks);
            diagnostics.GpuPointThreshold = minPoints;
            diagnostics.GpuSamplePointsPerTarget = sampleCountPerTarget;
            diagnostics.GpuZonesEligible = gpuZonesEligible;
            diagnostics.GpuZonesSkippedNoMesh = gpuZonesSkippedNoMesh;
            diagnostics.GpuZonesSkippedMissingTriangles = gpuZonesSkippedMissingTriangles;
            diagnostics.GpuZonesSkippedOpenMesh = gpuZonesSkippedOpenMesh;
            diagnostics.GpuZonesSkippedLowPoints = gpuZonesSkippedLowPoints;
            diagnostics.GpuUncertainPoints = gpuUncertainPoints;
            diagnostics.GpuMaxTrianglesPerZone = gpuMaxTrianglesPerZone;
            diagnostics.GpuMaxPointsPerZone = gpuMaxPointsPerZone;
            diagnostics.GpuOpenMeshZonesEligible = gpuOpenMeshZonesEligible;
            diagnostics.GpuOpenMeshZonesProcessed = gpuOpenMeshZonesProcessed;
            diagnostics.GpuOpenMeshBoundaryEdgeLimit = OpenMeshBoundaryEdgeLimit;
            diagnostics.GpuOpenMeshNonManifoldEdgeLimit = OpenMeshNonManifoldEdgeLimit;
            diagnostics.GpuOpenMeshOutsideTolerance = OpenMeshOutsideSampleTolerance;
            diagnostics.GpuOpenMeshNudge = maxOpenMeshNudge;
            diagnostics.GpuBatchDispatchCount = gpuBatchDispatchCount;
            diagnostics.GpuBatchMaxZones = gpuBatchMaxZones;
            diagnostics.GpuBatchMaxPoints = gpuBatchMaxPoints;
            diagnostics.GpuBatchMaxTriangles = gpuBatchMaxTriangles;
            diagnostics.GpuBatchAvgZonesPerDispatch = gpuBatchDispatchCount > 0
                ? gpuBatchZonesTotal / (double)gpuBatchDispatchCount
                : 0;
            if (usedCudaBvh)
            {
                diagnostics.GpuBackend = "CUDA BVH";
            }
            else if (usedD3d)
            {
                diagnostics.GpuBackend = "D3D11 batched";
            }
            else if (usedCuda)
            {
                diagnostics.GpuBackend = "CUDA";
            }

            return results;
        }

        private static double ComputeOpenMeshNudge(in Aabb bounds)
        {
            var dx = bounds.SizeX;
            var dy = bounds.SizeY;
            var dz = bounds.SizeZ;
            var diag = Math.Sqrt((dx * dx) + (dy * dy) + (dz * dz));
            var nudge = diag * OpenMeshPointNudgeRatio;
            if (nudge < OpenMeshPointNudgeMin)
            {
                nudge = OpenMeshPointNudgeMin;
            }
            if (nudge > OpenMeshPointNudgeMax)
            {
                nudge = OpenMeshPointNudgeMax;
            }
            return nudge;
        }

        private static Vector3D ApplyPointNudge(Vector3D point, double nudge, int seed)
        {
            if (nudge <= 0)
            {
                return point;
            }

            var jx = HashToUnit(seed * 73856093) * nudge;
            var jy = HashToUnit(seed * 19349663) * nudge;
            var jz = HashToUnit(seed * 83492791) * nudge;
            return new Vector3D(point.X + jx, point.Y + jy, point.Z + jz);
        }

        private static int BuildNudgeSeed(int zoneIndex, int targetIndex, int sampleIndex)
        {
            unchecked
            {
                var seed = (zoneIndex + 1) * 73856093;
                seed ^= (targetIndex + 1) * 19349663;
                seed ^= (sampleIndex + 1) * 83492791;
                return seed;
            }
        }

        private static double HashToUnit(int seed)
        {
            unchecked
            {
                uint x = (uint)seed;
                x ^= x >> 16;
                x *= 0x7feb352d;
                x ^= x >> 15;
                x *= 0x846ca68b;
                x ^= x >> 16;
                return (x / (double)uint.MaxValue) * 2.0 - 1.0;
            }
        }

        private static void AddHit(
            List<ZoneTargetIntersection> results,
            Dictionary<int, GpuBestZoneCandidate> bestHits,
            ZoneTargetIntersection hit,
            in Aabb targetBounds,
            int targetIndex,
            int zoneIndex,
            double[] zoneVolumes,
            double[] zoneCenterX,
            double[] zoneCenterY,
            double[] zoneCenterZ,
            SpaceMapperZoneResolutionStrategy resolutionStrategy)
        {
            if (hit == null)
            {
                return;
            }

            if (bestHits == null)
            {
                results.Add(hit);
                return;
            }

            var cx = (targetBounds.MinX + targetBounds.MaxX) * 0.5;
            var cy = (targetBounds.MinY + targetBounds.MaxY) * 0.5;
            var cz = (targetBounds.MinZ + targetBounds.MaxZ) * 0.5;
            var dx = cx - zoneCenterX[zoneIndex];
            var dy = cy - zoneCenterY[zoneIndex];
            var dz = cz - zoneCenterZ[zoneIndex];

            var candidate = new GpuBestZoneCandidate
            {
                Intersection = hit,
                ZoneIndex = zoneIndex,
                Volume = zoneVolumes[zoneIndex],
                DistanceSq = (dx * dx) + (dy * dy) + (dz * dz),
                OverlapVolume = hit.OverlapVolume
            };

            if (!bestHits.TryGetValue(targetIndex, out var best) || IsBetterCandidate(candidate, best, resolutionStrategy))
            {
                bestHits[targetIndex] = candidate;
            }
        }

        private static Triangle[] BuildTriangles(ZoneGeometry zone, double originX, double originY, double originZ)
        {
            var verts = zone?.TriangleVertices;
            if (verts == null || verts.Count < 3)
            {
                return null;
            }

            var triCount = verts.Count / 3;
            if (triCount == 0)
            {
                return null;
            }

            var triangles = new Triangle[triCount];
            for (int i = 0; i < triCount; i++)
            {
                var idx = i * 3;
                if (idx + 2 >= verts.Count)
                {
                    break;
                }

                var v0 = verts[idx];
                var v1 = verts[idx + 1];
                var v2 = verts[idx + 2];

                triangles[i] = new Triangle
                {
                    V0 = new Float4((float)(v0.X - originX), (float)(v0.Y - originY), (float)(v0.Z - originZ)),
                    V1 = new Float4((float)(v1.X - originX), (float)(v1.Y - originY), (float)(v1.Z - originZ)),
                    V2 = new Float4((float)(v2.X - originX), (float)(v2.Y - originY), (float)(v2.Z - originZ))
                };
            }

            return triangles;
        }

        private static List<Vector3D> BuildBoundsSamplePoints(in Aabb bounds, SpaceMapperTargetBoundsMode mode)
        {
            var points = new List<Vector3D>(GetSampleCount(mode));

            var minX = bounds.MinX;
            var minY = bounds.MinY;
            var minZ = bounds.MinZ;
            var maxX = bounds.MaxX;
            var maxY = bounds.MaxY;
            var maxZ = bounds.MaxZ;

            points.Add(new Vector3D(minX, minY, minZ));
            points.Add(new Vector3D(maxX, minY, minZ));
            points.Add(new Vector3D(minX, maxY, minZ));
            points.Add(new Vector3D(maxX, maxY, minZ));
            points.Add(new Vector3D(minX, minY, maxZ));
            points.Add(new Vector3D(maxX, minY, maxZ));
            points.Add(new Vector3D(minX, maxY, maxZ));
            points.Add(new Vector3D(maxX, maxY, maxZ));

            if (mode != SpaceMapperTargetBoundsMode.Aabb)
            {
                var cx = (minX + maxX) * 0.5;
                var cy = (minY + maxY) * 0.5;
                var cz = (minZ + maxZ) * 0.5;

                points.Add(new Vector3D(cx, cy, cz));

                if (mode == SpaceMapperTargetBoundsMode.KDop || mode == SpaceMapperTargetBoundsMode.Hull)
                {
                    points.Add(new Vector3D(minX, cy, cz));
                    points.Add(new Vector3D(maxX, cy, cz));
                    points.Add(new Vector3D(cx, minY, cz));
                    points.Add(new Vector3D(cx, maxY, cz));
                    points.Add(new Vector3D(cx, cy, minZ));
                    points.Add(new Vector3D(cx, cy, maxZ));
                }

                if (mode == SpaceMapperTargetBoundsMode.Hull)
                {
                    points.Add(new Vector3D(cx, minY, minZ));
                    points.Add(new Vector3D(cx, maxY, minZ));
                    points.Add(new Vector3D(cx, minY, maxZ));
                    points.Add(new Vector3D(cx, maxY, maxZ));

                    points.Add(new Vector3D(minX, cy, minZ));
                    points.Add(new Vector3D(maxX, cy, minZ));
                    points.Add(new Vector3D(minX, cy, maxZ));
                    points.Add(new Vector3D(maxX, cy, maxZ));

                    points.Add(new Vector3D(minX, minY, cz));
                    points.Add(new Vector3D(maxX, minY, cz));
                    points.Add(new Vector3D(minX, maxY, cz));
                    points.Add(new Vector3D(maxX, maxY, cz));
                }
            }

            return points;
        }

        private static int GetSampleCount(SpaceMapperTargetBoundsMode mode)
        {
            return mode switch
            {
                SpaceMapperTargetBoundsMode.Midpoint => 1,
                SpaceMapperTargetBoundsMode.Aabb => 8,
                SpaceMapperTargetBoundsMode.Obb => 9,
                SpaceMapperTargetBoundsMode.KDop => 15,
                SpaceMapperTargetBoundsMode.Hull => 27,
                _ => 8
            };
        }

        private static int ResolveGpuRayCount(SpaceMapperProcessingSettings settings, SpaceMapperProcessingMode mode)
        {
            if (settings != null)
            {
                if (settings.GpuRayCount >= 2)
                {
                    return 2;
                }

                if (settings.GpuRayCount <= 1)
                {
                    return 1;
                }
            }

            return mode == SpaceMapperProcessingMode.GpuQuick ? 1 : 2;
        }

        private static string BuildGpuInitFailureReason(string cudaReason, string d3dReason)
        {
            var hasCuda = !string.IsNullOrWhiteSpace(cudaReason);
            var hasD3d = !string.IsNullOrWhiteSpace(d3dReason);
            if (hasCuda && hasD3d)
            {
                return $"CUDA: {cudaReason}; D3D11: {d3dReason}";
            }

            if (hasCuda)
            {
                return $"CUDA: {cudaReason}";
            }

            if (hasD3d)
            {
                return $"D3D11: {d3dReason}";
            }

            return "Unable to initialize GPU backend.";
        }

        private static bool IsBetterCandidate(
            in GpuBestZoneCandidate candidate,
            in GpuBestZoneCandidate current,
            SpaceMapperZoneResolutionStrategy strategy)
        {
            if (candidate.Intersection.IsContained != current.Intersection.IsContained)
            {
                return candidate.Intersection.IsContained;
            }

            switch (strategy)
            {
                case SpaceMapperZoneResolutionStrategy.FirstMatch:
                    return candidate.ZoneIndex < current.ZoneIndex;
                case SpaceMapperZoneResolutionStrategy.LargestOverlap:
                    if (candidate.OverlapVolume > current.OverlapVolume)
                    {
                        return true;
                    }
                    if (candidate.OverlapVolume < current.OverlapVolume)
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

            return candidate.ZoneIndex < current.ZoneIndex;
        }

        private struct GpuZoneJob
        {
            public int ZoneIndex;
            public ZoneGeometry Zone;
            public Aabb ZoneBounds;
            public int[] CandidateTargets;
            public int[] TargetStartAbs;
            public int[] TargetCount;
            public bool UseOpenMeshRetry;
            public double OpenMeshNudge;
            public bool Intensive;
            public int PointsAdded;
            public int TrianglesAdded;
        }

        private struct GpuBestZoneCandidate
        {
            public ZoneTargetIntersection Intersection;
            public int ZoneIndex;
            public double Volume;
            public double DistanceSq;
            public double OverlapVolume;
        }
    }

    internal static class SpaceMapperEngineFactory
    {
        public static ISpatialIntersectionEngine Create(SpaceMapperProcessingMode requested)
        {
            return requested switch
            {
                SpaceMapperProcessingMode.Auto => new CudaIntersectionEngine(requested),
                SpaceMapperProcessingMode.CpuNormal => new CpuIntersectionEngine(requested),
                SpaceMapperProcessingMode.GpuQuick => new CudaIntersectionEngine(requested),
                SpaceMapperProcessingMode.GpuIntensive => new CudaIntersectionEngine(requested),
                SpaceMapperProcessingMode.Debug => new CpuIntersectionEngine(requested),
                _ => new CpuIntersectionEngine(SpaceMapperProcessingMode.CpuNormal)
            };
        }
    }
}
