using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;

namespace MicroEng.Navisworks
{
    internal static class GeometryExtractor
    {
        private static readonly Dictionary<string, CachedGeometry> Cache = new(StringComparer.OrdinalIgnoreCase);

        public static ZoneGeometry ExtractZoneGeometry(ModelItem item, string zoneId, string displayName, SpaceMapperProcessingSettings settings)
        {
            if (item == null) return null;
            var key = GetKey(item, zoneId);

            if (!Cache.TryGetValue(key, out var cached))
            {
                cached = BuildGeometry(item);
                Cache[key] = cached;
            }

            var bbox = ApplyOffsets(cached.BoundingBox, settings);
            return new ZoneGeometry
            {
                ZoneId = zoneId,
                DisplayName = displayName,
                ModelItem = item,
                BoundingBox = bbox,
                Vertices = cached.Vertices,
                Planes = cached.Planes
            };
        }

        public static TargetGeometry ExtractTargetGeometry(ModelItem item, string itemKey, string displayName)
        {
            if (item == null) return null;
            var key = GetKey(item, itemKey);

            if (!Cache.TryGetValue(key, out var cached))
            {
                cached = BuildGeometry(item);
                Cache[key] = cached;
            }

            return new TargetGeometry
            {
                ItemKey = itemKey,
                DisplayName = displayName,
                ModelItem = item,
                BoundingBox = cached.BoundingBox,
                Vertices = cached.Vertices
            };
        }

        private static string GetKey(ModelItem item, string fallback)
        {
            try
            {
                return item.InstanceGuid.ToString();
            }
            catch
            {
                return fallback ?? Guid.NewGuid().ToString();
            }
        }

        private static CachedGeometry BuildGeometry(ModelItem item)
        {
            var bbox = item.BoundingBox();
            var vertices = ExtractVertices(item);
            if (!vertices.Any() && bbox != null)
            {
                vertices = BoundingBoxToVertices(bbox);
            }

            var planes = GeometryMath.BuildPlanes(vertices, bbox);
            return new CachedGeometry
            {
                BoundingBox = bbox,
                Vertices = vertices,
                Planes = planes
            };
        }

        private static List<Vector3D> ExtractVertices(ModelItem item)
        {
            // TODO: Extract triangle vertices via ComApi on the Navisworks main thread.
            // For now, we return an empty list so the caller can fall back to bbox vertices.
            return new List<Vector3D>();
        }

        private static BoundingBox3D ApplyOffsets(BoundingBox3D bbox, SpaceMapperProcessingSettings settings)
        {
            if (bbox == null || settings == null) return bbox;
            var offset = settings.Offset3D;
            var offsetSides = settings.OffsetSides;
            var top = settings.OffsetTop;
            var bottom = settings.OffsetBottom;

            var min = bbox.Min;
            var max = bbox.Max;

            min = new Point3D(min.X - offset - offsetSides, min.Y - offset - offsetSides, min.Z - offset - bottom);
            max = new Point3D(max.X + offset + offsetSides, max.Y + offset + offsetSides, max.Z + offset + top);

            return new BoundingBox3D(min, max);
        }

        private static List<Vector3D> BoundingBoxToVertices(BoundingBox3D bbox)
        {
            var min = bbox.Min;
            var max = bbox.Max;
            return new List<Vector3D>
            {
                new Vector3D(min.X, min.Y, min.Z),
                new Vector3D(max.X, min.Y, min.Z),
                new Vector3D(min.X, max.Y, min.Z),
                new Vector3D(max.X, max.Y, min.Z),
                new Vector3D(min.X, min.Y, max.Z),
                new Vector3D(max.X, min.Y, max.Z),
                new Vector3D(min.X, max.Y, max.Z),
                new Vector3D(max.X, max.Y, max.Z),
            };
        }

        private class CachedGeometry
        {
            public BoundingBox3D BoundingBox { get; set; }
            public List<Vector3D> Vertices { get; set; }
            public List<PlaneEquation> Planes { get; set; }
        }
    }

    internal static class GeometryMath
    {
        private const double Epsilon = 1e-6;

        public static bool BoundingBoxesIntersect(BoundingBox3D a, BoundingBox3D b)
        {
            return a.Min.X <= b.Max.X && a.Max.X >= b.Min.X &&
                   a.Min.Y <= b.Max.Y && a.Max.Y >= b.Min.Y &&
                   a.Min.Z <= b.Max.Z && a.Max.Z >= b.Min.Z;
        }

        public static List<PlaneEquation> BuildPlanes(IReadOnlyList<Vector3D> vertices, BoundingBox3D bbox)
        {
            var planes = new List<PlaneEquation>();
            if (vertices != null && vertices.Count >= 3)
            {
                // Use triangle-based planes (one per tri from triplets)
                for (int i = 0; i + 2 < vertices.Count; i += 3)
                {
                    var p0 = vertices[i];
                    var p1 = vertices[i + 1];
                    var p2 = vertices[i + 2];
                    var plane = FromTriangle(p0, p1, p2);
                    if (plane.HasValue)
                    {
                        planes.Add(plane.Value);
                    }
                }
            }

            if (planes.Count == 0 && bbox != null)
            {
                // Fallback to bounding-box planes
                planes.Add(new PlaneEquation { Normal = new Vector3D(1, 0, 0), D = -bbox.Max.X });
                planes.Add(new PlaneEquation { Normal = new Vector3D(-1, 0, 0), D = bbox.Min.X });
                planes.Add(new PlaneEquation { Normal = new Vector3D(0, 1, 0), D = -bbox.Max.Y });
                planes.Add(new PlaneEquation { Normal = new Vector3D(0, -1, 0), D = bbox.Min.Y });
                planes.Add(new PlaneEquation { Normal = new Vector3D(0, 0, 1), D = -bbox.Max.Z });
                planes.Add(new PlaneEquation { Normal = new Vector3D(0, 0, -1), D = bbox.Min.Z });
            }

            return planes;
        }

        public static bool IsInside(IReadOnlyList<PlaneEquation> planes, Vector3D point)
        {
            if (planes == null || planes.Count == 0) return false;
            foreach (var plane in planes)
            {
                var v = plane.Normal;
                var val = v.X * point.X + v.Y * point.Y + v.Z * point.Z + plane.D;
                if (val > Epsilon) return false;
            }
            return true;
        }

        public static double EstimateOverlap(BoundingBox3D a, BoundingBox3D b)
        {
            var dx = Math.Max(0, Math.Min(a.Max.X, b.Max.X) - Math.Max(a.Min.X, b.Min.X));
            var dy = Math.Max(0, Math.Min(a.Max.Y, b.Max.Y) - Math.Max(a.Min.Y, b.Min.Y));
            var dz = Math.Max(0, Math.Min(a.Max.Z, b.Max.Z) - Math.Max(a.Min.Z, b.Min.Z));
            return dx * dy * dz;
        }

        private static PlaneEquation? FromTriangle(Vector3D p0, Vector3D p1, Vector3D p2)
        {
            var u = new Vector3D(p1.X - p0.X, p1.Y - p0.Y, p1.Z - p0.Z);
            var v = new Vector3D(p2.X - p0.X, p2.Y - p0.Y, p2.Z - p0.Z);
            var n = new Vector3D(
                u.Y * v.Z - u.Z * v.Y,
                u.Z * v.X - u.X * v.Z,
                u.X * v.Y - u.Y * v.X);
            var len = Math.Sqrt(n.X * n.X + n.Y * n.Y + n.Z * n.Z);
            if (len < Epsilon) return null;
            n = new Vector3D(n.X / len, n.Y / len, n.Z / len);
            var d = -(n.X * p0.X + n.Y * p0.Y + n.Z * p0.Z);
            return new PlaneEquation { Normal = n, D = d };
        }
    }
}
