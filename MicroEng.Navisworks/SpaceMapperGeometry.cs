using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using Autodesk.Navisworks.Api;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

namespace MicroEng.Navisworks
{
    internal static class GeometryExtractor
    {
        private static readonly Dictionary<string, CachedGeometry> Cache = new(StringComparer.OrdinalIgnoreCase);
        private static readonly object CacheLock = new();
        private static readonly HashSet<string> MeshFailureLogged = new(StringComparer.OrdinalIgnoreCase);
        private static Dispatcher _uiDispatcher;

        public static void SetUiDispatcher(Dispatcher dispatcher)
        {
            _uiDispatcher = dispatcher;
        }

        public static void ClearCache()
        {
            lock (CacheLock)
            {
                Cache.Clear();
                MeshFailureLogged.Clear();
            }
        }

        public static ZoneGeometry ExtractZoneGeometry(ModelItem item, string zoneId, string displayName, SpaceMapperProcessingSettings settings)
        {
            if (item == null) return null;
            var extractTriangles = settings?.ZoneContainmentEngine == SpaceMapperZoneContainmentEngine.MeshAccurate;
            var key = GetCacheKey(item, zoneId, extractTriangles);
            CachedGeometry cached;
            lock (CacheLock)
            {
                if (!Cache.TryGetValue(key, out cached))
                {
                    cached = BuildGeometry(item, extractTriangles);
                    Cache[key] = cached;
                }
            }

            var rawBox = cached.BoundingBox;
            var bbox = ApplyOffsets(rawBox, settings);
            return new ZoneGeometry
            {
                ZoneId = zoneId,
                DisplayName = displayName,
                ModelItem = item,
                RawBoundingBox = rawBox,
                BoundingBox = bbox,
                Vertices = cached.Vertices,
                Planes = cached.Planes,
                TriangleVertices = cached.TriangleVertices,
                TriangleCount = cached.TriangleCount,
                HasTriangleMesh = cached.HasTriangleMesh,
                MeshExtractionFailed = cached.MeshExtractionFailed,
                MeshExtractionError = cached.MeshExtractionError,
                MeshIsClosed = cached.MeshIsClosed,
                MeshBoundaryEdgeCount = cached.MeshBoundaryEdgeCount,
                MeshNonManifoldEdgeCount = cached.MeshNonManifoldEdgeCount,
                MeshFallbackReason = cached.MeshFallbackReason,
                MeshFallbackDetail = cached.MeshFallbackDetail
            };
        }

        public static TargetGeometry ExtractTargetGeometry(ModelItem item, string itemKey, string displayName, bool extractTriangles = false)
        {
            if (item == null) return null;
            var key = GetCacheKey(item, itemKey, includeTriangles: extractTriangles);
            CachedGeometry cached;
            lock (CacheLock)
            {
                if (!Cache.TryGetValue(key, out cached))
                {
                    cached = BuildGeometry(item, extractTriangles: extractTriangles);
                    Cache[key] = cached;
                }
            }

            return new TargetGeometry
            {
                ItemKey = itemKey,
                DisplayName = displayName,
                ModelItem = item,
                BoundingBox = cached.BoundingBox,
                Vertices = cached.Vertices,
                TriangleVertices = cached.TriangleVertices,
                TriangleCount = cached.TriangleCount
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

        private static string GetCacheKey(ModelItem item, string fallback, bool includeTriangles)
        {
            var baseKey = GetKey(item, fallback);
            return $"{baseKey}|{(includeTriangles ? "mesh" : "bounds")}";
        }

        private static CachedGeometry BuildGeometry(ModelItem item, bool extractTriangles)
        {
            var bbox = item.BoundingBox();
            List<Vector3D> triangleVertices = null;
            var extractionFailed = false;
            string extractionError = null;
            if (extractTriangles)
            {
                triangleVertices = ExtractVertices(item, out extractionFailed, out extractionError);
            }

            var verticesForPlanes = bbox != null ? BoundingBoxToVertices(bbox) : new List<Vector3D>();
            var planes = GeometryMath.BuildPlanes(verticesForPlanes, bbox);
            var hasTriangleMesh = triangleVertices != null
                && triangleVertices.Count >= 3
                && (triangleVertices.Count % 3 == 0);
            var meshIsClosed = false;
            var boundaryEdges = 0;
            var nonManifoldEdges = 0;
            if (hasTriangleMesh && GeometryMath.TryEvaluateMeshClosure(triangleVertices, out boundaryEdges, out nonManifoldEdges))
            {
                meshIsClosed = boundaryEdges == 0 && nonManifoldEdges == 0;
            }

            string fallbackReason = null;
            string fallbackDetail = null;
            if (extractTriangles)
            {
                if (extractionFailed)
                {
                    fallbackReason = "ExtractionFailed";
                    fallbackDetail = extractionError;
                    hasTriangleMesh = false;
                }
                else if (triangleVertices == null || triangleVertices.Count == 0)
                {
                    fallbackReason = "NoTriangles";
                    fallbackDetail = "No triangle vertices extracted.";
                    hasTriangleMesh = false;
                }
                else if (!hasTriangleMesh)
                {
                    fallbackReason = "InvalidTriangleList";
                    fallbackDetail = $"Triangle vertex count={triangleVertices.Count}.";
                    hasTriangleMesh = false;
                }
            }
            return new CachedGeometry
            {
                BoundingBox = bbox,
                Vertices = verticesForPlanes,
                Planes = planes,
                TriangleVertices = triangleVertices,
                TriangleCount = hasTriangleMesh ? triangleVertices.Count / 3 : 0,
                HasTriangleMesh = hasTriangleMesh,
                MeshExtractionFailed = extractionFailed,
                MeshExtractionError = extractionError,
                MeshIsClosed = meshIsClosed,
                MeshBoundaryEdgeCount = boundaryEdges,
                MeshNonManifoldEdgeCount = nonManifoldEdges,
                MeshFallbackReason = fallbackReason,
                MeshFallbackDetail = fallbackDetail
            };
        }

        private static List<Vector3D> ExtractVertices(ModelItem item, out bool failed, out string errorMessage)
        {
            failed = false;
            errorMessage = null;
            if (item == null)
            {
                return new List<Vector3D>();
            }

            var dispatcher = _uiDispatcher ?? System.Windows.Application.Current?.Dispatcher;
            if (dispatcher != null && !dispatcher.CheckAccess())
            {
                var localFailed = false;
                string localError = null;
                var vertices = dispatcher.Invoke(new Func<List<Vector3D>>(() => ExtractVerticesOnUiThread(item, out localFailed, out localError)));
                failed = localFailed;
                errorMessage = localError;
                return vertices;
            }

            return ExtractVerticesOnUiThread(item, out failed, out errorMessage);
        }

        private static List<Vector3D> ExtractVerticesOnUiThread(ModelItem item, out bool failed, out string errorMessage)
        {
            failed = false;
            errorMessage = null;
            var vertices = new List<Vector3D>();
            try
            {
                var selection = ComBridge.ToInwOpSelection(new ModelItemCollection { item });
                var paths = selection?.Paths();
                if (paths == null)
                {
                    return vertices;
                }

                foreach (ComApi.InwOaPath3 path in paths)
                {
                    if (path == null)
                    {
                        continue;
                    }

                    var fragments = path.Fragments();
                    if (fragments == null)
                    {
                        continue;
                    }

                    foreach (ComApi.InwOaFragment3 fragment in fragments)
                    {
                        if (fragment == null)
                        {
                            continue;
                        }

                        object matrixObj = null;
                        try
                        {
                            matrixObj = fragment.GetLocalToWorldMatrix()?.Matrix;
                        }
                        catch
                        {
                            matrixObj = null;
                        }

                        var callback = new SimplePrimitivesCallback(vertices, matrixObj);
                        fragment.GenerateSimplePrimitives(ComApi.nwEVertexProperty.eNORMAL, callback);
                    }
                }
            }
            catch (Exception ex)
            {
                failed = true;
                errorMessage = ex.Message;
                LogMeshFailure(item, ex);
            }

            return vertices;
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
            public List<Vector3D> TriangleVertices { get; set; }
            public int TriangleCount { get; set; }
            public bool HasTriangleMesh { get; set; }
            public bool MeshExtractionFailed { get; set; }
            public string MeshExtractionError { get; set; }
            public bool MeshIsClosed { get; set; }
            public int MeshBoundaryEdgeCount { get; set; }
            public int MeshNonManifoldEdgeCount { get; set; }
            public string MeshFallbackReason { get; set; }
            public string MeshFallbackDetail { get; set; }
        }

        private static void LogMeshFailure(ModelItem item, Exception ex)
        {
            if (item == null)
            {
                return;
            }

            var key = GetKey(item, null);
            lock (CacheLock)
            {
                if (MeshFailureLogged.Contains(key))
                {
                    return;
                }

                MeshFailureLogged.Add(key);
            }

            var name = item.DisplayName ?? key;
            MicroEngActions.Log($"Mesh extraction failed for '{name}': {ex.Message}");
        }

        [ComVisible(true)]
        [ClassInterface(ClassInterfaceType.None)]
        private sealed class SimplePrimitivesCallback : ComApi.InwSimplePrimitivesCB
        {
            private readonly List<Vector3D> _vertices;
            private readonly double[] _matrix;

            public SimplePrimitivesCallback(List<Vector3D> vertices, object matrixObj)
            {
                _vertices = vertices;
                _matrix = ExtractMatrix(matrixObj);
            }

            public void Triangle(ComApi.InwSimpleVertex v1, ComApi.InwSimpleVertex v2, ComApi.InwSimpleVertex v3)
            {
                if (!TryGetCoord(v1, out var p1) || !TryGetCoord(v2, out var p2) || !TryGetCoord(v3, out var p3))
                {
                    return;
                }

                p1 = TransformPoint(p1, _matrix);
                p2 = TransformPoint(p2, _matrix);
                p3 = TransformPoint(p3, _matrix);

                _vertices.Add(p1);
                _vertices.Add(p2);
                _vertices.Add(p3);
            }

            public void Line(ComApi.InwSimpleVertex v1, ComApi.InwSimpleVertex v2)
            {
            }

            public void Point(ComApi.InwSimpleVertex v1)
            {
            }

            public void SnapPoint(ComApi.InwSimpleVertex v1)
            {
            }
        }

        private static bool TryGetCoord(ComApi.InwSimpleVertex vertex, out Vector3D point)
        {
            point = new Vector3D();
            if (vertex == null)
            {
                return false;
            }

            object coordObj;
            try
            {
                coordObj = vertex.coord;
            }
            catch
            {
                return false;
            }

            if (coordObj == null)
            {
                return false;
            }

            if (coordObj is ComApi.InwLPos3f pos)
            {
                point = new Vector3D(pos.data1, pos.data2, pos.data3);
                return true;
            }

            if (coordObj is Array arr && arr.Length >= 3)
            {
                var start = arr.GetLowerBound(0);
                point = new Vector3D(
                    ToDouble(arr.GetValue(start)),
                    ToDouble(arr.GetValue(start + 1)),
                    ToDouble(arr.GetValue(start + 2)));
                return true;
            }

            return false;
        }

        private static double[] ExtractMatrix(object matrixObj)
        {
            if (matrixObj is Array arr && arr.Length >= 12)
            {
                var start = arr.GetLowerBound(0);
                var result = new double[arr.Length];
                for (int i = 0; i < arr.Length; i++)
                {
                    result[i] = ToDouble(arr.GetValue(start + i));
                }
                return result;
            }

            return null;
        }

        private static Vector3D TransformPoint(Vector3D point, double[] matrix)
        {
            if (matrix == null || matrix.Length < 12)
            {
                return point;
            }

            var x = point.X;
            var y = point.Y;
            var z = point.Z;

            if (matrix.Length >= 16)
            {
                var rowTx = matrix[3];
                var rowTy = matrix[7];
                var rowTz = matrix[11];
                var colTx = matrix[12];
                var colTy = matrix[13];
                var colTz = matrix[14];
                var rowHasTranslation = !IsNearZero(rowTx) || !IsNearZero(rowTy) || !IsNearZero(rowTz);
                var colHasTranslation = !IsNearZero(colTx) || !IsNearZero(colTy) || !IsNearZero(colTz);

                if (colHasTranslation && !rowHasTranslation)
                {
                    return new Vector3D(
                        matrix[0] * x + matrix[4] * y + matrix[8] * z + matrix[12],
                        matrix[1] * x + matrix[5] * y + matrix[9] * z + matrix[13],
                        matrix[2] * x + matrix[6] * y + matrix[10] * z + matrix[14]);
                }

                return new Vector3D(
                    matrix[0] * x + matrix[1] * y + matrix[2] * z + matrix[3],
                    matrix[4] * x + matrix[5] * y + matrix[6] * z + matrix[7],
                    matrix[8] * x + matrix[9] * y + matrix[10] * z + matrix[11]);
            }

            var rowTx3 = matrix[3];
            var rowTy3 = matrix[7];
            var colTx3 = matrix[9];
            var colTy3 = matrix[10];
            var rowHasTranslation3 = !IsNearZero(rowTx3) || !IsNearZero(rowTy3);
            var colHasTranslation3 = !IsNearZero(colTx3) || !IsNearZero(colTy3);
            if (colHasTranslation3 && !rowHasTranslation3)
            {
                return new Vector3D(
                    matrix[0] * x + matrix[3] * y + matrix[6] * z + matrix[9],
                    matrix[1] * x + matrix[4] * y + matrix[7] * z + matrix[10],
                    matrix[2] * x + matrix[5] * y + matrix[8] * z + matrix[11]);
            }

            return new Vector3D(
                matrix[0] * x + matrix[1] * y + matrix[2] * z + matrix[3],
                matrix[4] * x + matrix[5] * y + matrix[6] * z + matrix[7],
                matrix[8] * x + matrix[9] * y + matrix[10] * z + matrix[11]);
        }

        private static bool IsNearZero(double value)
        {
            return Math.Abs(value) <= 1e-9;
        }

        private static double ToDouble(object value)
        {
            return value == null ? 0.0 : Convert.ToDouble(value);
        }
    }

    internal static class GeometryMath
    {
        private const double Epsilon = 1e-6;
        private const double MeshQuantizeStep = 1e-3;

        public static bool BoundingBoxesIntersect(BoundingBox3D a, BoundingBox3D b)
        {
            return a.Min.X <= b.Max.X && a.Max.X >= b.Min.X &&
                   a.Min.Y <= b.Max.Y && a.Max.Y >= b.Min.Y &&
                   a.Min.Z <= b.Max.Z && a.Max.Z >= b.Min.Z;
        }

        public static List<PlaneEquation> BuildPlanes(IReadOnlyList<Vector3D> vertices, BoundingBox3D bbox)
        {
            var planes = new List<PlaneEquation>();
            if (vertices != null && vertices.Count >= 3 && (vertices.Count % 3 == 0))
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

        public static bool TryIsInsideMesh(IReadOnlyList<Vector3D> triVerts, Vector3D point, out bool inside)
        {
            inside = false;
            if (!TryGetMeshBounds(triVerts, out var min, out var max))
            {
                return false;
            }

            var rayDir = new Vector3D(1.0, 0.237, 0.103);
            return TryIsInsideMeshWithDirection(triVerts, point, rayDir, min, max, out inside);
        }

        public static bool TryIsInsideMeshRobust(IReadOnlyList<Vector3D> triVerts, Vector3D point, out bool inside)
        {
            inside = false;
            if (!TryGetMeshBounds(triVerts, out var min, out var max))
            {
                return false;
            }

            var dirs = new[]
            {
                new Vector3D(1.0, 0.237, 0.103),
                new Vector3D(0.103, 1.0, 0.237),
                new Vector3D(0.237, 0.103, 1.0)
            };

            var insideVotes = 0;
            var outsideVotes = 0;
            var any = false;

            foreach (var dir in dirs)
            {
                if (TryIsInsideMeshWithDirection(triVerts, point, dir, min, max, out var voteInside))
                {
                    any = true;
                    if (voteInside)
                    {
                        insideVotes++;
                    }
                    else
                    {
                        outsideVotes++;
                    }
                }
            }

            if (!any)
            {
                return false;
            }

            inside = insideVotes >= outsideVotes;
            return true;
        }

        private static bool TryIsInsideMeshWithDirection(
            IReadOnlyList<Vector3D> triVerts,
            Vector3D point,
            Vector3D rayDir,
            Vector3D min,
            Vector3D max,
            out bool inside)
        {
            inside = false;
            if (triVerts == null || triVerts.Count < 3 || (triVerts.Count % 3 != 0))
            {
                return false;
            }

            if (point.X < min.X - Epsilon || point.X > max.X + Epsilon
                || point.Y < min.Y - Epsilon || point.Y > max.Y + Epsilon
                || point.Z < min.Z - Epsilon || point.Z > max.Z + Epsilon)
            {
                inside = false;
                return true;
            }

            var len = Math.Sqrt(rayDir.X * rayDir.X + rayDir.Y * rayDir.Y + rayDir.Z * rayDir.Z);
            if (len < Epsilon)
            {
                return false;
            }

            var dir = new Vector3D(rayDir.X / len, rayDir.Y / len, rayDir.Z / len);
            var hits = 0;

            for (int i = 0; i + 2 < triVerts.Count; i += 3)
            {
                var v0 = triVerts[i];
                var v1 = triVerts[i + 1];
                var v2 = triVerts[i + 2];

                if (IsPointOnTriangle(point, v0, v1, v2))
                {
                    inside = true;
                    return true;
                }

                if (RayIntersectsTriangle(point, dir, v0, v1, v2))
                {
                    hits++;
                }
            }

            inside = (hits & 1) == 1;
            return true;
        }

        private static bool TryGetMeshBounds(IReadOnlyList<Vector3D> triVerts, out Vector3D min, out Vector3D max)
        {
            min = new Vector3D();
            max = new Vector3D();

            if (triVerts == null || triVerts.Count < 3 || (triVerts.Count % 3 != 0))
            {
                return false;
            }

            double minX = double.PositiveInfinity;
            double minY = double.PositiveInfinity;
            double minZ = double.PositiveInfinity;
            double maxX = double.NegativeInfinity;
            double maxY = double.NegativeInfinity;
            double maxZ = double.NegativeInfinity;

            for (int i = 0; i < triVerts.Count; i++)
            {
                var v = triVerts[i];
                if (v.X < minX) minX = v.X;
                if (v.Y < minY) minY = v.Y;
                if (v.Z < minZ) minZ = v.Z;
                if (v.X > maxX) maxX = v.X;
                if (v.Y > maxY) maxY = v.Y;
                if (v.Z > maxZ) maxZ = v.Z;
            }

            min = new Vector3D(minX, minY, minZ);
            max = new Vector3D(maxX, maxY, maxZ);
            return true;
        }

        public static bool TryEvaluateMeshClosure(IReadOnlyList<Vector3D> triVerts, out int boundaryEdges, out int nonManifoldEdges)
        {
            boundaryEdges = 0;
            nonManifoldEdges = 0;
            if (triVerts == null || triVerts.Count < 3 || (triVerts.Count % 3 != 0))
            {
                return false;
            }

            var edgeCounts = new Dictionary<EdgeKey, int>(triVerts.Count);
            for (int i = 0; i + 2 < triVerts.Count; i += 3)
            {
                var v0 = Quantize(triVerts[i]);
                var v1 = Quantize(triVerts[i + 1]);
                var v2 = Quantize(triVerts[i + 2]);

                AddEdge(edgeCounts, v0, v1);
                AddEdge(edgeCounts, v1, v2);
                AddEdge(edgeCounts, v2, v0);
            }

            foreach (var count in edgeCounts.Values)
            {
                if (count == 1)
                {
                    boundaryEdges++;
                }
                else if (count != 2)
                {
                    nonManifoldEdges++;
                }
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

        private static void AddEdge(Dictionary<EdgeKey, int> counts, QuantizedVertex a, QuantizedVertex b)
        {
            if (a.Equals(b))
            {
                return;
            }

            var key = new EdgeKey(a, b);
            if (counts.TryGetValue(key, out var current))
            {
                counts[key] = current + 1;
            }
            else
            {
                counts[key] = 1;
            }
        }

        private static QuantizedVertex Quantize(Vector3D v)
        {
            return new QuantizedVertex(
                Quantize(v.X),
                Quantize(v.Y),
                Quantize(v.Z));
        }

        private static long Quantize(double value)
        {
            return (long)Math.Round(value / MeshQuantizeStep);
        }

        private static bool RayIntersectsTriangle(Vector3D origin, Vector3D dir, Vector3D v0, Vector3D v1, Vector3D v2)
        {
            var edge1 = new Vector3D(v1.X - v0.X, v1.Y - v0.Y, v1.Z - v0.Z);
            var edge2 = new Vector3D(v2.X - v0.X, v2.Y - v0.Y, v2.Z - v0.Z);

            var pvec = Cross(dir, edge2);
            var det = Dot(edge1, pvec);
            if (Math.Abs(det) < Epsilon)
            {
                return false;
            }

            var invDet = 1.0 / det;
            var tvec = new Vector3D(origin.X - v0.X, origin.Y - v0.Y, origin.Z - v0.Z);
            var u = Dot(tvec, pvec) * invDet;
            if (u < -Epsilon || u > 1.0 + Epsilon)
            {
                return false;
            }

            var qvec = Cross(tvec, edge1);
            var v = Dot(dir, qvec) * invDet;
            if (v < -Epsilon || u + v > 1.0 + Epsilon)
            {
                return false;
            }

            var t = Dot(edge2, qvec) * invDet;
            return t > Epsilon;
        }

        private static bool IsPointOnTriangle(Vector3D p, Vector3D a, Vector3D b, Vector3D c)
        {
            var ab = new Vector3D(b.X - a.X, b.Y - a.Y, b.Z - a.Z);
            var ac = new Vector3D(c.X - a.X, c.Y - a.Y, c.Z - a.Z);
            var ap = new Vector3D(p.X - a.X, p.Y - a.Y, p.Z - a.Z);

            var normal = Cross(ab, ac);
            var normLenSq = Dot(normal, normal);
            if (normLenSq < Epsilon * Epsilon)
            {
                return false;
            }

            var dist = Dot(normal, ap);
            if (Math.Abs(dist) > Epsilon * Math.Sqrt(normLenSq))
            {
                return false;
            }

            var dot00 = Dot(ac, ac);
            var dot01 = Dot(ac, ab);
            var dot02 = Dot(ac, ap);
            var dot11 = Dot(ab, ab);
            var dot12 = Dot(ab, ap);

            var denom = dot00 * dot11 - dot01 * dot01;
            if (Math.Abs(denom) < Epsilon)
            {
                return false;
            }

            var invDenom = 1.0 / denom;
            var u = (dot11 * dot02 - dot01 * dot12) * invDenom;
            var v = (dot00 * dot12 - dot01 * dot02) * invDenom;

            return u >= -Epsilon && v >= -Epsilon && (u + v) <= 1.0 + Epsilon;
        }

        private static Vector3D Cross(Vector3D a, Vector3D b)
        {
            return new Vector3D(
                a.Y * b.Z - a.Z * b.Y,
                a.Z * b.X - a.X * b.Z,
                a.X * b.Y - a.Y * b.X);
        }

        private static double Dot(Vector3D a, Vector3D b)
        {
            return a.X * b.X + a.Y * b.Y + a.Z * b.Z;
        }

        private readonly struct QuantizedVertex : IEquatable<QuantizedVertex>
        {
            public readonly long X;
            public readonly long Y;
            public readonly long Z;

            public QuantizedVertex(long x, long y, long z)
            {
                X = x;
                Y = y;
                Z = z;
            }

            public bool Equals(QuantizedVertex other)
            {
                return X == other.X && Y == other.Y && Z == other.Z;
            }

            public override bool Equals(object obj)
            {
                return obj is QuantizedVertex other && Equals(other);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    var hash = (int)2166136261;
                    hash = (hash * 16777619) ^ X.GetHashCode();
                    hash = (hash * 16777619) ^ Y.GetHashCode();
                    hash = (hash * 16777619) ^ Z.GetHashCode();
                    return hash;
                }
            }
        }

        private readonly struct EdgeKey : IEquatable<EdgeKey>
        {
            public readonly QuantizedVertex A;
            public readonly QuantizedVertex B;

            public EdgeKey(QuantizedVertex a, QuantizedVertex b)
            {
                if (Compare(a, b) <= 0)
                {
                    A = a;
                    B = b;
                }
                else
                {
                    A = b;
                    B = a;
                }
            }

            public bool Equals(EdgeKey other)
            {
                return A.Equals(other.A) && B.Equals(other.B);
            }

            public override bool Equals(object obj)
            {
                return obj is EdgeKey other && Equals(other);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    return (A.GetHashCode() * 397) ^ B.GetHashCode();
                }
            }

            private static int Compare(QuantizedVertex left, QuantizedVertex right)
            {
                var cmp = left.X.CompareTo(right.X);
                if (cmp != 0) return cmp;
                cmp = left.Y.CompareTo(right.Y);
                if (cmp != 0) return cmp;
                return left.Z.CompareTo(right.Z);
            }
        }
    }
}
