using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;
using MicroEng.Navisworks.SpaceMapper.Estimation;
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
            CancellationToken cancellationToken = default);
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
            CancellationToken cancellationToken = default)
        {
            var results = new List<ZoneTargetIntersection>();
            if (zones == null || targets == null || zones.Count == 0 || targets.Count == 0)
            {
                return results;
            }

            diagnostics ??= new SpaceMapperEngineDiagnostics();
            var effectivePreset = SpaceMapperPresetLogic.ResolvePreset(settings, preflightCache?.LastResult, out _);
            diagnostics.PresetUsed = effectivePreset;
            var needsPartial = settings != null && (settings.TagPartialSeparately || settings.TreatPartialAsContained);
            var useOriginOnly = settings != null && (settings.UseOriginPointOnly || effectivePreset == SpaceMapperPerformancePreset.Fast);
            var usePointIndex = useOriginOnly && !needsPartial;
            var treatPartialAsContained = settings != null && settings.TreatPartialAsContained;
            var enableMultiZone = settings == null || settings.EnableMultipleZones;
            var traversal = ResolveFastTraversal(
                settings?.FastTraversalMode ?? SpaceMapperFastTraversalMode.Auto,
                useOriginOnly,
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
                    usePointIndex,
                    enableMultiZone);
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
                var prepared = BuildTargetBounds(targets, usePointIndex);
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

            var useBestSelection = useOriginOnly && !enableMultiZone;
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
                zoneValid[i] = useOriginOnly || (zone.Planes != null && zone.Planes.Count > 0);
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
                    Interlocked.Increment(ref processedZones);
                    return;
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
                    if (useOriginOnly)
                    {
                        var targetBoundsLocal = targetBounds[idx];
                        var cx = (targetBoundsLocal.MinX + targetBoundsLocal.MaxX) * 0.5;
                        var cy = (targetBoundsLocal.MinY + targetBoundsLocal.MaxY) * 0.5;
                        var cz = (targetBoundsLocal.MinZ + targetBoundsLocal.MaxZ) * 0.5;

                        var inside = cx >= zoneBoundsLocal.MinX && cx <= zoneBoundsLocal.MaxX
                            && cy >= zoneBoundsLocal.MinY && cy <= zoneBoundsLocal.MaxY
                            && cz >= zoneBoundsLocal.MinZ && cz <= zoneBoundsLocal.MaxZ;

                        bool isContained;
                        bool isPartial;
                        double overlapVolume;

                        if (inside)
                        {
                            isContained = true;
                            isPartial = false;
                            overlapVolume = EstimateOverlap(zoneBoundsLocal, targetBoundsLocal);
                        }
                        else if (needsPartial && zoneBoundsLocal.Intersects(targetBoundsLocal))
                        {
                            isContained = treatPartialAsContained;
                            isPartial = true;
                            overlapVolume = EstimateOverlap(zoneBoundsLocal, targetBoundsLocal);
                        }
                        else
                        {
                            return;
                        }

                        if (useBestSelection)
                        {
                            var dx = cx - zoneCenterX[zoneIndex];
                            var dy = cy - zoneCenterY[zoneIndex];
                            var dz = cz - zoneCenterZ[zoneIndex];
                            var candidate = new BestZoneCandidate
                            {
                                ZoneIndex = zoneIndex,
                                IsContained = isContained,
                                IsPartial = isPartial,
                                Volume = zoneVolumes[zoneIndex],
                                DistanceSq = (dx * dx) + (dy * dy) + (dz * dz),
                                OverlapVolume = overlapVolume
                            };

                            if (!localBest.TryGetValue(idx, out var best) || IsBetterCandidate(candidate, best))
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
                                IsContained = isContained,
                                IsPartial = isPartial,
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
                            effectivePreset,
                            useOriginOnly,
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
                            effectivePreset,
                            useOriginOnly,
                            needsPartial,
                            treatPartialAsContained);
                    }

                    if (hit != null)
                    {
                        localResults.Add(hit);
                    }
                });

                var zoneEnd = Stopwatch.GetTimestamp();
                if (!trackNarrow || useOriginOnly)
                {
                    zoneNarrowTicks = zoneEnd - zoneStart;
                }
                stampLocal.Value = stamp;

                Interlocked.Add(ref candidatePairs, zoneCandidates);
                UpdateMax(ref maxCandidates, zoneCandidates);
                Interlocked.Add(ref totalTicks, zoneEnd - zoneStart);
                Interlocked.Add(ref narrowTicks, zoneNarrowTicks);

                var zonesDone = Interlocked.Increment(ref processedZones);
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
                        if (!merged.TryGetValue(kvp.Key, out var best) || IsBetterCandidate(kvp.Value, best))
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
            bool usePointIndex,
            bool enableMultiZone)
        {
            var results = new List<ZoneTargetIntersection>();
            var zoneCount = zones.Count;
            var zoneBounds = new Aabb[zoneCount];
            var zoneValid = new bool[zoneCount];

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
                var prepared = BuildTargetBounds(targets, usePointIndex);
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
                    var inside = cx >= zoneBoundsLocal.MinX && cx <= zoneBoundsLocal.MaxX
                        && cy >= zoneBoundsLocal.MinY && cy <= zoneBoundsLocal.MaxY
                        && cz >= zoneBoundsLocal.MinZ && cz <= zoneBoundsLocal.MaxZ;

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

                        if (!hasBest || IsBetterCandidate(candidate, best))
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

        private struct BestZoneCandidate
        {
            public int ZoneIndex;
            public bool IsContained;
            public bool IsPartial;
            public double Volume;
            public double DistanceSq;
            public double OverlapVolume;
        }

        private static ZoneTargetIntersection ClassifyIntersection(
            ZoneGeometry zone,
            in Aabb zoneBounds,
            in Aabb targetBounds,
            string targetKey,
            SpaceMapperPerformancePreset preset,
            bool useOriginOnly,
            bool needsPartial,
            bool treatPartialAsContained)
        {
            if (useOriginOnly)
            {
                return ClassifyByOriginPointInZoneAabb(
                    zone,
                    zoneBounds,
                    targetBounds,
                    targetKey,
                    needsPartial,
                    treatPartialAsContained);
            }

            return ClassifyByPlanes(zone, zoneBounds, targetBounds, targetKey, preset);
        }

        private static ZoneTargetIntersection ClassifyByOriginPointInZoneAabb(
            ZoneGeometry zone,
            in Aabb zoneBounds,
            in Aabb targetBounds,
            string targetKey,
            bool needsPartial,
            bool treatPartialAsContained)
        {
            var cx = (targetBounds.MinX + targetBounds.MaxX) * 0.5;
            var cy = (targetBounds.MinY + targetBounds.MaxY) * 0.5;
            var cz = (targetBounds.MinZ + targetBounds.MaxZ) * 0.5;

            var inside = cx >= zoneBounds.MinX && cx <= zoneBounds.MaxX
                && cy >= zoneBounds.MinY && cy <= zoneBounds.MaxY
                && cz >= zoneBounds.MinZ && cz <= zoneBounds.MaxZ;

            if (inside)
            {
                return new ZoneTargetIntersection
                {
                    ZoneId = zone.ZoneId,
                    TargetItemKey = targetKey,
                    IsContained = true,
                    IsPartial = false,
                    OverlapVolume = EstimateOverlap(zoneBounds, targetBounds)
                };
            }

            if (!needsPartial || !zoneBounds.Intersects(targetBounds))
            {
                return null;
            }

            return new ZoneTargetIntersection
            {
                ZoneId = zone.ZoneId,
                TargetItemKey = targetKey,
                IsContained = treatPartialAsContained,
                IsPartial = true,
                OverlapVolume = EstimateOverlap(zoneBounds, targetBounds)
            };
        }

        private static ZoneTargetIntersection ClassifyAabbOnly(ZoneGeometry zone, in Aabb zoneBounds, in Aabb targetBounds, string targetKey)
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

            return new ZoneTargetIntersection
            {
                ZoneId = zone.ZoneId,
                TargetItemKey = targetKey,
                IsContained = contained,
                IsPartial = !contained,
                OverlapVolume = EstimateOverlap(zoneBounds, targetBounds)
            };
        }

        private static ZoneTargetIntersection ClassifyByPlanes(ZoneGeometry zone, in Aabb zoneBounds, in Aabb targetBounds, string targetKey, SpaceMapperPerformancePreset preset)
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

            if (preset == SpaceMapperPerformancePreset.Accurate && !(anyInside && anyOutside))
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

            return new ZoneTargetIntersection
            {
                ZoneId = zone.ZoneId,
                TargetItemKey = targetKey,
                IsContained = anyInside && !anyOutside,
                IsPartial = anyInside && anyOutside,
                OverlapVolume = EstimateOverlap(zoneBounds, targetBounds)
            };
        }

        private static void TestCorner(IReadOnlyList<PlaneEquation> planes, double x, double y, double z, ref bool anyInside, ref bool anyOutside)
        {
            if (anyInside && anyOutside)
            {
                return;
            }

            var inside = GeometryMath.IsInside(planes, new Vector3D(x, y, z));
            if (inside) anyInside = true; else anyOutside = true;
        }

        private static (Aabb[] Bounds, string[] Keys) BuildTargetBounds(IReadOnlyList<TargetGeometry> targets, bool usePointIndex)
        {
            var bounds = new List<Aabb>(targets.Count);
            var keys = new List<string>(targets.Count);

            for (int i = 0; i < targets.Count; i++)
            {
                var target = targets[i];
                var bbox = target?.BoundingBox;
                if (bbox == null) continue;
                bounds.Add(usePointIndex ? ToPointAabb(bbox) : ToAabb(bbox));
                keys.Add(target.ItemKey);
            }

            return (bounds.ToArray(), keys.ToArray());
        }

        private static Aabb ToAabb(BoundingBox3D bbox)
        {
            var min = bbox.Min;
            var max = bbox.Max;
            return new Aabb(min.X, min.Y, min.Z, max.X, max.Y, max.Z);
        }

        private static Aabb ToPointAabb(BoundingBox3D bbox)
        {
            var min = bbox.Min;
            var max = bbox.Max;
            var cx = (min.X + max.X) * 0.5;
            var cy = (min.Y + max.Y) * 0.5;
            var cz = (min.Z + max.Z) * 0.5;
            return new Aabb(cx, cy, cz, cx, cy, cz);
        }

        private static Aabb ComputeWorldBounds(Aabb[] targets)
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

        private static TimeSpan ToTimeSpan(long ticks)
        {
            if (ticks <= 0)
            {
                return TimeSpan.Zero;
            }

            var seconds = ticks / (double)Stopwatch.Frequency;
            return TimeSpan.FromSeconds(seconds);
        }

        private static bool IsBetterCandidate(in BestZoneCandidate candidate, in BestZoneCandidate current)
        {
            if (candidate.IsContained != current.IsContained)
            {
                return candidate.IsContained;
            }

            if (candidate.Volume < current.Volume)
            {
                return true;
            }

            if (candidate.Volume > current.Volume)
            {
                return false;
            }

            return candidate.DistanceSq < current.DistanceSq;
        }

        private static SpaceMapperFastTraversalMode ResolveFastTraversal(
            SpaceMapperFastTraversalMode requested,
            bool useOriginOnly,
            bool needsPartial,
            int targetCount,
            int zoneCount)
        {
            var allowTargetMajor = useOriginOnly && !needsPartial;
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

        private static void UpdateMax(ref int current, int candidate)
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

        private static double EstimateOverlap(in Aabb zoneBounds, in Aabb targetBounds)
        {
            var dx = Math.Max(0, Math.Min(zoneBounds.MaxX, targetBounds.MaxX) - Math.Max(zoneBounds.MinX, targetBounds.MinX));
            var dy = Math.Max(0, Math.Min(zoneBounds.MaxY, targetBounds.MaxY) - Math.Max(zoneBounds.MinY, targetBounds.MinY));
            var dz = Math.Max(0, Math.Min(zoneBounds.MaxZ, targetBounds.MaxZ) - Math.Max(zoneBounds.MinZ, targetBounds.MinZ));
            return dx * dy * dz;
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
    }

    internal class CudaIntersectionEngine : ISpatialIntersectionEngine
    {
        public SpaceMapperProcessingMode Mode { get; }

        public CudaIntersectionEngine(SpaceMapperProcessingMode mode)
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
            CancellationToken cancellationToken = default)
        {
            throw new NotImplementedException("GPU engine not implemented yet.");
        }
    }

    internal static class SpaceMapperEngineFactory
    {
        public static ISpatialIntersectionEngine Create(SpaceMapperProcessingMode requested)
        {
            if (requested == SpaceMapperProcessingMode.Auto)
            {
                requested = SpaceMapperProcessingMode.CpuNormal;
            }

            return requested switch
            {
                SpaceMapperProcessingMode.CpuNormal => new CpuIntersectionEngine(requested),
                SpaceMapperProcessingMode.GpuQuick => new CudaIntersectionEngine(requested),
                SpaceMapperProcessingMode.GpuIntensive => new CudaIntersectionEngine(requested),
                SpaceMapperProcessingMode.Debug => new CpuIntersectionEngine(requested),
                _ => new CpuIntersectionEngine(SpaceMapperProcessingMode.CpuNormal)
            };
        }
    }
}
