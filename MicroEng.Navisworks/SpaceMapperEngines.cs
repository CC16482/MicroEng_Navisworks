using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
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

            SpatialHashGrid grid;
            Aabb[] targetBounds;
            string[] targetKeys;

            var buildSw = Stopwatch.StartNew();
            if (preflightCache?.Grid != null && preflightCache.TargetBounds != null && preflightCache.TargetKeys != null)
            {
                grid = preflightCache.Grid;
                targetBounds = preflightCache.TargetBounds;
                targetKeys = preflightCache.TargetKeys;
                diagnostics.UsedPreflightIndex = true;
            }
            else
            {
                diagnostics.UsedPreflightIndex = false;
                var prepared = BuildTargetBounds(targets);
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

            var bag = new ConcurrentBag<ZoneTargetIntersection>();
            var totalZones = zones.Count;
            var processedZones = 0;
            long candidatePairs = 0;
            int maxCandidates = 0;
            long totalTicks = 0;
            long narrowTicks = 0;

            var visitedLocal = new ThreadLocal<int[]>(() => new int[targetBounds.Length]);
            var stampLocal = new ThreadLocal<int>(() => 0);
            var parallelOptions = new ParallelOptions
            {
                CancellationToken = cancellationToken,
                MaxDegreeOfParallelism = settings?.MaxThreads ?? Environment.ProcessorCount
            };

            Parallel.ForEach(zones, parallelOptions, zone =>
            {
                if (zone?.BoundingBox == null)
                {
                    Interlocked.Increment(ref processedZones);
                    return;
                }

                if (effectivePreset != SpaceMapperPerformancePreset.Fast && (zone.Planes == null || zone.Planes.Count == 0))
                {
                    Interlocked.Increment(ref processedZones);
                    return;
                }

                var zoneBounds = GetZoneQueryBounds(zone, settings);
                var visited = visitedLocal.Value;
                var stamp = stampLocal.Value;
                var zoneCandidates = 0;
                var zoneNarrowTicks = 0L;
                var zoneStart = Stopwatch.GetTimestamp();

                grid.VisitCandidates(zoneBounds, visited, ref stamp, idx =>
                {
                    zoneCandidates++;
                    var t0 = Stopwatch.GetTimestamp();
                    var hit = ClassifyIntersection(zone, zoneBounds, targetBounds[idx], targetKeys[idx], effectivePreset);
                    var t1 = Stopwatch.GetTimestamp();
                    zoneNarrowTicks += t1 - t0;
                    if (hit != null)
                    {
                        bag.Add(hit);
                    }
                });

                var zoneEnd = Stopwatch.GetTimestamp();
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

            diagnostics.CandidatePairs = candidatePairs;
            diagnostics.MaxCandidatesPerZone = maxCandidates;
            diagnostics.AvgCandidatesPerZone = totalZones == 0 ? 0 : candidatePairs / (double)totalZones;
            diagnostics.NarrowPhaseTime = ToTimeSpan(narrowTicks);
            diagnostics.CandidateQueryTime = ToTimeSpan(Math.Max(0, totalTicks - narrowTicks));

            return bag.ToList();
        }

        private static ZoneTargetIntersection ClassifyIntersection(ZoneGeometry zone, in Aabb zoneBounds, in Aabb targetBounds, string targetKey, SpaceMapperPerformancePreset preset)
        {
            if (preset == SpaceMapperPerformancePreset.Fast)
            {
                return ClassifyAabbOnly(zone, zoneBounds, targetBounds, targetKey);
            }

            return ClassifyByPlanes(zone, zoneBounds, targetBounds, targetKey, preset);
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

        private static (Aabb[] Bounds, string[] Keys) BuildTargetBounds(IReadOnlyList<TargetGeometry> targets)
        {
            var bounds = new List<Aabb>(targets.Count);
            var keys = new List<string>(targets.Count);

            for (int i = 0; i < targets.Count; i++)
            {
                var target = targets[i];
                var bbox = target?.BoundingBox;
                if (bbox == null) continue;
                bounds.Add(ToAabb(bbox));
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
