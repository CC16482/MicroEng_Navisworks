using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;
using MicroEng.Navisworks.SpaceMapper.Geometry;

namespace MicroEng.Navisworks.SpaceMapper.Estimation
{
    internal sealed class SpaceMapperPreflightCache
    {
        public string Signature { get; set; }
        public SpatialHashGrid Grid { get; set; }
        public string[] TargetKeys { get; set; }
        public Aabb[] TargetBounds { get; set; }
        public Aabb WorldBounds { get; set; }
        public double CellSizeUsed { get; set; }
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

            if (zones.Count == 0 || targets.Count == 0)
            {
                return Task.FromResult<SpaceMapperPreflightResult>(new SpaceMapperPreflightResult
                {
                    ZoneCount = zones.Count,
                    TargetCount = targets.Count,
                    Signature = signature
                });
            }

            var zoneBounds = new Aabb[zones.Count];
            for (int i = 0; i < zones.Count; i++)
            {
                token.ThrowIfCancellationRequested();
                var bbox = zones[i].ModelItem?.BoundingBox();
                if (bbox == null)
                {
                    zoneBounds[i] = new Aabb(0, 0, 0, 0, 0, 0);
                    continue;
                }
                zoneBounds[i] = Inflate(ToAabb(bbox), request.ProcessingSettings);
            }

            var canReuseCache = _cache != null
                && string.Equals(_cache.Signature, signature, StringComparison.OrdinalIgnoreCase)
                && _cache.Grid != null
                && _cache.TargetBounds != null
                && _cache.TargetKeys != null;

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
                    targetBounds[i] = bbox == null ? new Aabb(0, 0, 0, 0, 0, 0) : ToAabb(bbox);
                    targetKeys[i] = targets[i].ItemKey;
                }

                worldBounds = ComputeWorldBounds(zoneBounds, targetBounds);
                cellSize = SpatialGridSizing.ComputeCellSize(worldBounds, request.ProcessingSettings?.IndexGranularity ?? 0);
            }

            return Task.Run(() =>
            {
                token.ThrowIfCancellationRequested();

                progress?.Report(new SpaceMapperPreflightProgress { Stage = "Building index", ZonesProcessed = 0, TotalZones = zoneBounds.Length });

                SpatialHashGrid grid;
                var buildSw = Stopwatch.StartNew();

                if (canReuseCache)
                {
                    grid = _cache.Grid;
                    buildSw.Stop();
                }
                else
                {
                    grid = new SpatialHashGrid(worldBounds, cellSize, targetBounds);
                    buildSw.Stop();
                }

                var querySw = Stopwatch.StartNew();
                long candidatePairs = 0;
                int maxCandidates = 0;
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

                querySw.Stop();

                var result = new SpaceMapperPreflightResult
                {
                    ZoneCount = zoneBounds.Length,
                    TargetCount = targetBounds.Length,
                    CandidatePairs = candidatePairs,
                    MaxCandidatesPerZone = maxCandidates,
                    AvgCandidatesPerZone = zoneBounds.Length == 0 ? 0 : candidatePairs / (double)zoneBounds.Length,
                    CellSizeUsed = grid.CellSize,
                    BuildIndexTime = buildSw.Elapsed,
                    QueryTime = querySw.Elapsed,
                    Signature = signature
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
                        TargetBounds = targetBounds,
                        TargetKeys = targetKeys,
                        WorldBounds = worldBounds,
                        CellSizeUsed = grid.CellSize,
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
            sb.Append(settings.Offset3D.ToString("0.###")).Append('|');
            sb.Append(settings.OffsetTop.ToString("0.###")).Append('|');
            sb.Append(settings.OffsetBottom.ToString("0.###")).Append('|');
            sb.Append(settings.OffsetSides.ToString("0.###")).Append('|');
            sb.Append(settings.TreatPartialAsContained ? '1' : '0').Append('|');
            sb.Append(settings.TagPartialSeparately ? '1' : '0').Append('|');
            sb.Append(settings.EnableMultipleZones ? '1' : '0').Append('|');
            sb.Append(settings.IndexGranularity).Append('|');

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
    }
}
