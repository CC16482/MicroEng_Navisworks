using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;

namespace MicroEng.Navisworks
{
    internal interface ISpatialIntersectionEngine
    {
        SpaceMapperProcessingMode Mode { get; }

        IList<ZoneTargetIntersection> ComputeIntersections(
            IReadOnlyList<ZoneGeometry> zones,
            IReadOnlyList<TargetGeometry> targets,
            SpaceMapperProcessingSettings settings,
            IProgress<SpaceMapperProgress> progress = null,
            CancellationToken cancellationToken = default);
    }

    internal class CpuIntersectionEngine : ISpatialIntersectionEngine
    {
        public SpaceMapperProcessingMode Mode { get; }

        public CpuIntersectionEngine(SpaceMapperProcessingMode mode = SpaceMapperProcessingMode.CpuNormal)
        {
            Mode = SpaceMapperProcessingMode.CpuNormal;
        }

        public IList<ZoneTargetIntersection> ComputeIntersections(
            IReadOnlyList<ZoneGeometry> zones,
            IReadOnlyList<TargetGeometry> targets,
            SpaceMapperProcessingSettings settings,
            IProgress<SpaceMapperProgress> progress = null,
            CancellationToken cancellationToken = default)
        {
            var results = new List<ZoneTargetIntersection>();
            if (zones == null || targets == null || zones.Count == 0 || targets.Count == 0) return results;

            var candidates = new List<(ZoneGeometry zone, TargetGeometry target)>();
            foreach (var zone in zones)
            {
                foreach (var target in targets)
                {
                    if (zone?.BoundingBox == null || target?.BoundingBox == null) continue;
                    if (!GeometryMath.BoundingBoxesIntersect(zone.BoundingBox, target.BoundingBox)) continue;
                    candidates.Add((zone, target));
                }
            }

            var total = candidates.Count;
            var processed = 0;
            var bag = new ConcurrentBag<ZoneTargetIntersection>();
            var parallelOptions = new ParallelOptions
            {
                CancellationToken = cancellationToken,
                MaxDegreeOfParallelism = settings?.MaxThreads ?? Environment.ProcessorCount
            };

            Parallel.ForEach(candidates, parallelOptions, pair =>
            {
                var hit = ClassifyIntersection(pair.zone, pair.target);
                if (hit != null)
                {
                    bag.Add(hit);
                }

                var done = Interlocked.Increment(ref processed);
                progress?.Report(new SpaceMapperProgress
                {
                    ProcessedPairs = done,
                    TotalPairs = total,
                    ZonesProcessed = zones.Count,
                    TargetsProcessed = targets.Count
                });
            });

            return bag.ToList();
        }

        private static ZoneTargetIntersection ClassifyIntersection(ZoneGeometry zone, TargetGeometry target)
        {
            if (zone?.Planes == null || zone.Planes.Count == 0 || target?.Vertices == null || target.Vertices.Count == 0)
            {
                return null;
            }

            var anyInside = false;
            var anyOutside = false;
            foreach (var v in target.Vertices)
            {
                var inside = GeometryMath.IsInside(zone.Planes, v);
                if (inside) anyInside = true; else anyOutside = true;
                if (anyInside && anyOutside) break;
            }

            if (!anyInside) return null;

            return new ZoneTargetIntersection
            {
                ZoneId = zone.ZoneId,
                TargetItemKey = target.ItemKey,
                IsContained = anyInside && !anyOutside,
                IsPartial = anyInside && anyOutside,
                OverlapVolume = GeometryMath.EstimateOverlap(zone.BoundingBox, target.BoundingBox)
            };
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
            // For now only CPU Normal is supported.
            return new CpuIntersectionEngine(SpaceMapperProcessingMode.CpuNormal);
        }
    }
}
