using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;
using MicroEng.Navisworks;
using MicroEng.Navisworks.SpaceMapper.Geometry;

namespace MicroEng.Navisworks.SpaceMapper.Estimation
{
    internal sealed class SpaceMapperPreflightCache
    {
        public string Signature { get; set; }
        public SpatialHashGrid Grid { get; set; }
        public SpatialHashGrid ZoneGrid { get; set; }
        public string[] TargetKeys { get; set; }
        public Aabb[] TargetBounds { get; set; }
        public Aabb[] ZoneBoundsInflated { get; set; }
        public int[] ZoneIndexMap { get; set; }
        public Aabb WorldBounds { get; set; }
        public double CellSizeUsed { get; set; }
        public bool PointIndex { get; set; }
        public SpaceMapperFastTraversalMode TraversalUsed { get; set; } = SpaceMapperFastTraversalMode.ZoneMajor;
        public SpaceMapperPreflightResult LastResult { get; set; }
    }

    internal sealed class SpaceMapperPreflightProgress
    {
        public int ZonesProcessed { get; set; }
        public int TotalZones { get; set; }
        public string Stage { get; set; }
        public int Percentage => TotalZones <= 0 ? 0 : (int)(ZonesProcessed * 100.0 / TotalZones);
    }

    internal sealed class SpaceMapperPreflightService
    {
        private SpaceMapperPreflightCache _cache;

        public SpaceMapperPreflightCache Cache => _cache;

        public Task<SpaceMapperPreflightResult> RunAsync(
            SpaceMapperRequest request,
            CancellationToken token,
            IProgress<SpaceMapperPreflightProgress> progress = null)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));

            var signature = BuildSignature(request);
            if (_cache != null
                && string.Equals(_cache.Signature, signature, StringComparison.OrdinalIgnoreCase)
                && _cache.LastResult != null)
            {
                return Task.FromResult(_cache.LastResult);
            }

            var doc = Application.ActiveDocument;
            if (doc == null)
            {
                return Task.FromResult<SpaceMapperPreflightResult>(null);
            }

            var session = SpaceMapperService.GetSession(request.ScraperProfileName);
            var requiresSession = request.ZoneSource == ZoneSourceType.DataScraperZones;
            if (requiresSession && session == null)
            {
                return Task.FromResult<SpaceMapperPreflightResult>(null);
            }

            var zones = SpaceMapperService.ResolveZones(session, request.ZoneSource, request.ZoneSetName, doc).ToList();
            var targetsByRule = new Dictionary<string, List<SpaceMapperTargetRule>>(StringComparer.OrdinalIgnoreCase);
            var targets = SpaceMapperService.ResolveTargets(doc, request.TargetRules, targetsByRule).ToList();
            var settings = request.ProcessingSettings ?? new SpaceMapperProcessingSettings();
            var containmentEngine = ResolveZoneContainmentEngine(settings);
            var targetBoundsMode = SpaceMapperBoundsResolver.ResolveTargetBoundsMode(settings, containmentEngine);
            var needsPartial = settings.TagPartialSeparately
                || settings.TreatPartialAsContained
                || settings.WriteZoneBehaviorProperty
                || settings.WriteZoneContainmentPercentProperty;
            if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                needsPartial = false;
            }
            var usePointIndex = targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint;
            var traversal = ResolveFastTraversal(settings.FastTraversalMode, usePointIndex, needsPartial, targets.Count, zones.Count);
            var usePointIndexForBounds = usePointIndex;

            if (zones.Count == 0 || targets.Count == 0)
            {
                return Task.FromResult<SpaceMapperPreflightResult>(new SpaceMapperPreflightResult
                {
                    ZoneCount = zones.Count,
                    TargetCount = targets.Count,
                    Signature = signature,
                    TraversalUsed = traversal
                });
            }

            var zoneBounds = new Aabb[zones.Count];
            var validZoneBounds = new List<Aabb>(zones.Count);
            var zoneIndexMap = traversal == SpaceMapperFastTraversalMode.TargetMajor
                ? new List<int>(zones.Count)
                : null;

            for (int i = 0; i < zones.Count; i++)
            {
                token.ThrowIfCancellationRequested();
                var bbox = zones[i].ModelItem?.BoundingBox();
                if (bbox == null)
                {
                    zoneBounds[i] = new Aabb(0, 0, 0, 0, 0, 0);
                    continue;
                }

                var inflated = Inflate(ToAabb(bbox), settings);
                zoneBounds[i] = inflated;
                validZoneBounds.Add(inflated);
                zoneIndexMap?.Add(i);
            }

            var validZoneBoundsArray = validZoneBounds.ToArray();

            var canReuseTargetGrid = _cache != null
                && string.Equals(_cache.Signature, signature, StringComparison.OrdinalIgnoreCase)
                && _cache.Grid != null
                && _cache.TargetBounds != null
                && _cache.TargetKeys != null
                && _cache.PointIndex == usePointIndex
                && _cache.TraversalUsed == SpaceMapperFastTraversalMode.ZoneMajor;

            var canReuseZoneGrid = _cache != null
                && string.Equals(_cache.Signature, signature, StringComparison.OrdinalIgnoreCase)
                && _cache.ZoneGrid != null
                && _cache.ZoneBoundsInflated != null
                && _cache.ZoneIndexMap != null
                && _cache.TargetBounds != null
                && _cache.TargetKeys != null
                && _cache.PointIndex == usePointIndex
                && _cache.TraversalUsed == SpaceMapperFastTraversalMode.TargetMajor;

            var canReuseCache = traversal == SpaceMapperFastTraversalMode.TargetMajor
                ? canReuseZoneGrid
                : canReuseTargetGrid;

            Aabb[] targetBounds;
            string[] targetKeys;
            Aabb worldBounds;
            double cellSize;

            if (canReuseCache)
            {
                targetBounds = _cache.TargetBounds;
                targetKeys = _cache.TargetKeys;
                worldBounds = _cache.WorldBounds;
                cellSize = _cache.CellSizeUsed;
            }
            else
            {
                targetBounds = new Aabb[targets.Count];
                targetKeys = new string[targets.Count];
                for (int i = 0; i < targets.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    var bbox = targets[i].ModelItem?.BoundingBox();
                    if (bbox == null)
                    {
                        targetBounds[i] = new Aabb(0, 0, 0, 0, 0, 0);
                    }
                    else
                    {
                        targetBounds[i] = usePointIndexForBounds
                            ? ToMidpointAabb(bbox, settings.TargetMidpointMode)
                            : ToAabb(bbox);
                    }
                    targetKeys[i] = targets[i].ItemKey;
                }

                worldBounds = ComputeWorldBounds(validZoneBoundsArray, targetBounds);
                cellSize = SpatialGridSizing.ComputeCellSize(worldBounds, request.ProcessingSettings?.IndexGranularity ?? 0);
            }

            return Task.Run(() =>
            {
                token.ThrowIfCancellationRequested();

                var totalWork = traversal == SpaceMapperFastTraversalMode.TargetMajor ? targetBounds.Length : zoneBounds.Length;
                progress?.Report(new SpaceMapperPreflightProgress { Stage = "Building index", ZonesProcessed = 0, TotalZones = totalWork });

                SpatialHashGrid grid = null;
                SpatialHashGrid zoneGrid = null;
                var buildSw = Stopwatch.StartNew();

                if (canReuseCache)
                {
                    grid = _cache.Grid;
                    zoneGrid = _cache.ZoneGrid;
                    buildSw.Stop();
                }
                else
                {
                    if (traversal == SpaceMapperFastTraversalMode.TargetMajor)
                    {
                        zoneGrid = new SpatialHashGrid(worldBounds, cellSize, validZoneBoundsArray);
                    }
                    else
                    {
                        grid = new SpatialHashGrid(worldBounds, cellSize, targetBounds);
                    }
                    buildSw.Stop();
                }

                var querySw = Stopwatch.StartNew();
                long candidatePairs = 0;
                int maxCandidates = 0;
                double avgCandidatesPerZone = 0;
                double avgCandidatesPerTarget = 0;

                if (traversal == SpaceMapperFastTraversalMode.TargetMajor)
                {
                    for (int i = 0; i < targetBounds.Length; i++)
                    {
                        token.ThrowIfCancellationRequested();

                        var bounds = targetBounds[i];
                        var cx = (bounds.MinX + bounds.MaxX) * 0.5;
                        var cy = (bounds.MinY + bounds.MaxY) * 0.5;
                        var cz = (bounds.MinZ + bounds.MaxZ) * 0.5;
                        var count = zoneGrid.CountPointCandidates(cx, cy, cz);

                        candidatePairs += count;
                        if (count > maxCandidates) maxCandidates = count;

                        if (progress != null && (i % 16 == 0 || i == targetBounds.Length - 1))
                        {
                            progress.Report(new SpaceMapperPreflightProgress
                            {
                                Stage = "Querying",
                                ZonesProcessed = i + 1,
                                TotalZones = targetBounds.Length
                            });
                        }
                    }

                    avgCandidatesPerTarget = targetBounds.Length == 0 ? 0 : candidatePairs / (double)targetBounds.Length;
                }
                else
                {
                    var visited = new int[targetBounds.Length];
                    var stamp = 0;

                    for (int i = 0; i < zoneBounds.Length; i++)
                    {
                        token.ThrowIfCancellationRequested();
                        var count = grid.CountCandidates(zoneBounds[i], visited, ref stamp);
                        candidatePairs += count;
                        if (count > maxCandidates) maxCandidates = count;

                        if (progress != null && (i % 8 == 0 || i == zoneBounds.Length - 1))
                        {
                            progress.Report(new SpaceMapperPreflightProgress
                            {
                                Stage = "Querying",
                                ZonesProcessed = i + 1,
                                TotalZones = zoneBounds.Length
                            });
                        }
                    }

                    avgCandidatesPerZone = zoneBounds.Length == 0 ? 0 : candidatePairs / (double)zoneBounds.Length;
                }

                querySw.Stop();

                var result = new SpaceMapperPreflightResult
                {
                    ZoneCount = zoneBounds.Length,
                    TargetCount = targetBounds.Length,
                    CandidatePairs = candidatePairs,
                    MaxCandidatesPerZone = traversal == SpaceMapperFastTraversalMode.ZoneMajor ? maxCandidates : 0,
                    AvgCandidatesPerZone = avgCandidatesPerZone,
                    MaxCandidatesPerTarget = traversal == SpaceMapperFastTraversalMode.TargetMajor ? maxCandidates : 0,
                    AvgCandidatesPerTarget = avgCandidatesPerTarget,
                    CellSizeUsed = (grid ?? zoneGrid).CellSize,
                    BuildIndexTime = buildSw.Elapsed,
                    QueryTime = querySw.Elapsed,
                    Signature = signature,
                    TraversalUsed = traversal
                };

                if (canReuseCache)
                {
                    _cache.LastResult = result;
                }
                else
                {
                    _cache = new SpaceMapperPreflightCache
                    {
                        Signature = signature,
                        Grid = grid,
                        ZoneGrid = zoneGrid,
                        TargetBounds = targetBounds,
                        TargetKeys = targetKeys,
                        ZoneBoundsInflated = traversal == SpaceMapperFastTraversalMode.TargetMajor ? validZoneBoundsArray : null,
                        ZoneIndexMap = traversal == SpaceMapperFastTraversalMode.TargetMajor ? zoneIndexMap?.ToArray() : null,
                        WorldBounds = worldBounds,
                        CellSizeUsed = (grid ?? zoneGrid).CellSize,
                        PointIndex = usePointIndex,
                        TraversalUsed = traversal,
                        LastResult = result
                    };
                }

                return result;
            }, token);
        }

        private static Aabb ToAabb(BoundingBox3D bbox)
        {
            var min = bbox.Min;
            var max = bbox.Max;
            return new Aabb(min.X, min.Y, min.Z, max.X, max.Y, max.Z);
        }

        private static Aabb ToMidpointAabb(BoundingBox3D bbox, SpaceMapperMidpointMode mode)
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

        private static Aabb ToMidpointAabb(in Aabb bbox, SpaceMapperMidpointMode mode)
        {
            var cx = (bbox.MinX + bbox.MaxX) * 0.5;
            var cy = (bbox.MinY + bbox.MaxY) * 0.5;
            var cz = mode == SpaceMapperMidpointMode.BoundingBoxBottomCenter
                ? bbox.MinZ
                : (bbox.MinZ + bbox.MaxZ) * 0.5;
            return new Aabb(cx, cy, cz, cx, cy, cz);
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

        private static SpaceMapperZoneContainmentEngine ResolveZoneContainmentEngine(SpaceMapperProcessingSettings settings)
        {
            if (settings == null)
            {
                return SpaceMapperZoneContainmentEngine.BoundsFast;
            }

            return settings.ZoneContainmentEngine;
        }

        private static Aabb ComputeWorldBounds(Aabb[] zones, Aabb[] targets)
        {
            var hasAny = false;
            double minX = 0, minY = 0, minZ = 0, maxX = 0, maxY = 0, maxZ = 0;

            void Add(Aabb box)
            {
                if (!hasAny)
                {
                    minX = box.MinX; minY = box.MinY; minZ = box.MinZ;
                    maxX = box.MaxX; maxY = box.MaxY; maxZ = box.MaxZ;
                    hasAny = true;
                    return;
                }

                minX = Math.Min(minX, box.MinX);
                minY = Math.Min(minY, box.MinY);
                minZ = Math.Min(minZ, box.MinZ);
                maxX = Math.Max(maxX, box.MaxX);
                maxY = Math.Max(maxY, box.MaxY);
                maxZ = Math.Max(maxZ, box.MaxZ);
            }

            foreach (var z in zones) Add(z);
            foreach (var t in targets) Add(t);

            return hasAny ? new Aabb(minX, minY, minZ, maxX, maxY, maxZ) : new Aabb(0, 0, 0, 0, 0, 0);
        }

        internal static string BuildSignature(SpaceMapperRequest request)
        {
            var sb = new StringBuilder();
            sb.Append(request.ScraperProfileName ?? string.Empty).Append('|');
            sb.Append(request.Scope).Append('|');
            sb.Append(request.ZoneSource).Append('|');
            sb.Append(request.ZoneSetName ?? string.Empty).Append('|');

            var settings = request.ProcessingSettings ?? new SpaceMapperProcessingSettings();
            sb.Append(settings.ProcessingMode).Append('|');
            sb.Append(settings.GpuRayCount).Append('|');
            sb.Append(settings.Offset3D.ToString("0.###")).Append('|');
            sb.Append(settings.OffsetTop.ToString("0.###")).Append('|');
            sb.Append(settings.OffsetBottom.ToString("0.###")).Append('|');
            sb.Append(settings.OffsetSides.ToString("0.###")).Append('|');
            sb.Append(settings.TreatPartialAsContained ? '1' : '0').Append('|');
            sb.Append(settings.TagPartialSeparately ? '1' : '0').Append('|');
            sb.Append(settings.WriteZoneBehaviorProperty ? '1' : '0').Append('|');
            sb.Append(settings.WriteZoneContainmentPercentProperty ? '1' : '0').Append('|');
            sb.Append(settings.ContainmentCalculationMode).Append('|');
            sb.Append(settings.EnableMultipleZones ? '1' : '0').Append('|');
            sb.Append(settings.IndexGranularity).Append('|');
            sb.Append(settings.PerformancePreset).Append('|');
            sb.Append(settings.FastTraversalMode).Append('|');
            sb.Append(settings.UseOriginPointOnly ? '1' : '0').Append('|');
            sb.Append(settings.ZoneBoundsMode).Append('|');
            sb.Append(settings.ZoneKDopVariant).Append('|');
            sb.Append(settings.TargetBoundsMode).Append('|');
            sb.Append(settings.TargetKDopVariant).Append('|');
            sb.Append(settings.TargetMidpointMode).Append('|');
            sb.Append(settings.ZoneContainmentEngine).Append('|');
            sb.Append(settings.ZoneResolutionStrategy).Append('|');
            sb.Append(settings.ExcludeZonesFromTargets ? '1' : '0').Append('|');

            foreach (var rule in request.TargetRules.OrderBy(r => r.Name ?? string.Empty))
            {
                sb.Append(rule.Name ?? string.Empty).Append('|');
                sb.Append(rule.TargetDefinition).Append('|');
                sb.Append(rule.MinLevel ?? -1).Append('|');
                sb.Append(rule.MaxLevel ?? -1).Append('|');
                sb.Append(rule.SetSearchName ?? string.Empty).Append('|');
                sb.Append(rule.CategoryFilter ?? string.Empty).Append('|');
                sb.Append(rule.MembershipMode).Append('|');
                sb.Append(rule.Enabled ? '1' : '0').Append('|');
            }

            using (var sha1 = SHA1.Create())
            {
                var bytes = Encoding.UTF8.GetBytes(sb.ToString());
                var hash = sha1.ComputeHash(bytes);
                var hex = new StringBuilder(hash.Length * 2);
                foreach (var b in hash)
                {
                    hex.Append(b.ToString("x2"));
                }
                return hex.ToString();
            }
        }

        internal static SpaceMapperPreflightResult RunForDataset(
            IReadOnlyList<Aabb> zoneBounds,
            IReadOnlyList<Aabb> targetBounds,
            string[] targetKeys,
            SpaceMapperProcessingSettings settings,
            CancellationToken token,
            out SpaceMapperPreflightCache cache,
            IProgress<SpaceMapperPreflightProgress> progress = null)
        {
            cache = null;
            settings ??= new SpaceMapperProcessingSettings();

            var zoneBoundsArray = zoneBounds?.ToArray() ?? Array.Empty<Aabb>();
            var targetBoundsArray = targetBounds?.ToArray() ?? Array.Empty<Aabb>();

            if (zoneBoundsArray.Length == 0 || targetBoundsArray.Length == 0)
            {
                return new SpaceMapperPreflightResult
                {
                    ZoneCount = zoneBoundsArray.Length,
                    TargetCount = targetBoundsArray.Length,
                    Signature = string.Empty
                };
            }

            var containmentEngine = ResolveZoneContainmentEngine(settings);
            var targetBoundsMode = SpaceMapperBoundsResolver.ResolveTargetBoundsMode(settings, containmentEngine);
            var needsPartial = settings.TagPartialSeparately
                || settings.TreatPartialAsContained
                || settings.WriteZoneBehaviorProperty
                || settings.WriteZoneContainmentPercentProperty;
            if (targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint)
            {
                needsPartial = false;
            }
            var usePointIndex = targetBoundsMode == SpaceMapperTargetBoundsMode.Midpoint;
            var traversal = ResolveFastTraversal(settings.FastTraversalMode, usePointIndex, needsPartial, targetBoundsArray.Length, zoneBoundsArray.Length);
            var usePointIndexForBounds = usePointIndex;

            if (usePointIndexForBounds)
            {
                for (int i = 0; i < targetBoundsArray.Length; i++)
                {
                    targetBoundsArray[i] = ToMidpointAabb(targetBoundsArray[i], settings.TargetMidpointMode);
                }
            }

            var totalWork = traversal == SpaceMapperFastTraversalMode.TargetMajor ? targetBoundsArray.Length : zoneBoundsArray.Length;
            progress?.Report(new SpaceMapperPreflightProgress { Stage = "Building index", ZonesProcessed = 0, TotalZones = totalWork });

            SpatialHashGrid grid = null;
            SpatialHashGrid zoneGrid = null;
            var buildSw = Stopwatch.StartNew();

            var worldBounds = ComputeWorldBounds(zoneBoundsArray, targetBoundsArray);
            var cellSize = SpatialGridSizing.ComputeCellSize(worldBounds, settings.IndexGranularity);

            if (traversal == SpaceMapperFastTraversalMode.TargetMajor)
            {
                zoneGrid = new SpatialHashGrid(worldBounds, cellSize, zoneBoundsArray);
            }
            else
            {
                grid = new SpatialHashGrid(worldBounds, cellSize, targetBoundsArray);
            }

            buildSw.Stop();

            var querySw = Stopwatch.StartNew();
            long candidatePairs = 0;
            int maxCandidates = 0;
            double avgCandidatesPerZone = 0;
            double avgCandidatesPerTarget = 0;

            if (traversal == SpaceMapperFastTraversalMode.TargetMajor)
            {
                for (int i = 0; i < targetBoundsArray.Length; i++)
                {
                    token.ThrowIfCancellationRequested();
                    var bounds = targetBoundsArray[i];
                    var cx = (bounds.MinX + bounds.MaxX) * 0.5;
                    var cy = (bounds.MinY + bounds.MaxY) * 0.5;
                    var cz = (bounds.MinZ + bounds.MaxZ) * 0.5;
                    var count = zoneGrid.CountPointCandidates(cx, cy, cz);
                    candidatePairs += count;
                    if (count > maxCandidates) maxCandidates = count;

                    if (progress != null && (i % 16 == 0 || i == targetBoundsArray.Length - 1))
                    {
                        progress.Report(new SpaceMapperPreflightProgress
                        {
                            Stage = "Querying",
                            ZonesProcessed = i + 1,
                            TotalZones = targetBoundsArray.Length
                        });
                    }
                }

                avgCandidatesPerTarget = targetBoundsArray.Length == 0 ? 0 : candidatePairs / (double)targetBoundsArray.Length;
            }
            else
            {
                var visited = new int[targetBoundsArray.Length];
                var stamp = 0;

                for (int i = 0; i < zoneBoundsArray.Length; i++)
                {
                    token.ThrowIfCancellationRequested();
                    var count = grid.CountCandidates(zoneBoundsArray[i], visited, ref stamp);
                    candidatePairs += count;
                    if (count > maxCandidates) maxCandidates = count;

                    if (progress != null && (i % 8 == 0 || i == zoneBoundsArray.Length - 1))
                    {
                        progress.Report(new SpaceMapperPreflightProgress
                        {
                            Stage = "Querying",
                            ZonesProcessed = i + 1,
                            TotalZones = zoneBoundsArray.Length
                        });
                    }
                }

                avgCandidatesPerZone = zoneBoundsArray.Length == 0 ? 0 : candidatePairs / (double)zoneBoundsArray.Length;
            }

            querySw.Stop();

            var result = new SpaceMapperPreflightResult
            {
                ZoneCount = zoneBoundsArray.Length,
                TargetCount = targetBoundsArray.Length,
                CandidatePairs = candidatePairs,
                MaxCandidatesPerZone = traversal == SpaceMapperFastTraversalMode.ZoneMajor ? maxCandidates : 0,
                AvgCandidatesPerZone = avgCandidatesPerZone,
                MaxCandidatesPerTarget = traversal == SpaceMapperFastTraversalMode.TargetMajor ? maxCandidates : 0,
                AvgCandidatesPerTarget = avgCandidatesPerTarget,
                CellSizeUsed = (grid ?? zoneGrid).CellSize,
                BuildIndexTime = buildSw.Elapsed,
                QueryTime = querySw.Elapsed,
                Signature = string.Empty,
                TraversalUsed = traversal
            };

            cache = new SpaceMapperPreflightCache
            {
                Signature = string.Empty,
                Grid = grid,
                ZoneGrid = zoneGrid,
                TargetBounds = targetBoundsArray,
                TargetKeys = targetKeys ?? Array.Empty<string>(),
                ZoneBoundsInflated = traversal == SpaceMapperFastTraversalMode.TargetMajor ? zoneBoundsArray : null,
                ZoneIndexMap = traversal == SpaceMapperFastTraversalMode.TargetMajor ? Enumerable.Range(0, zoneBoundsArray.Length).ToArray() : null,
                WorldBounds = worldBounds,
                CellSizeUsed = (grid ?? zoneGrid).CellSize,
                PointIndex = usePointIndex,
                TraversalUsed = traversal,
                LastResult = result
            };

            return result;
        }

        private static SpaceMapperFastTraversalMode ResolveFastTraversal(
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
    }
}
