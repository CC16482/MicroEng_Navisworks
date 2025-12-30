using System;
using System.Collections.Generic;

namespace MicroEng.Navisworks.SpaceMapper.Geometry
{
    internal readonly struct CellKey : IEquatable<CellKey>
    {
        public readonly int X, Y, Z;
        public CellKey(int x, int y, int z) { X = x; Y = y; Z = z; }

        public bool Equals(CellKey other) => X == other.X && Y == other.Y && Z == other.Z;
        public override bool Equals(object obj) => obj is CellKey other && Equals(other);

        public override int GetHashCode()
        {
            unchecked
            {
                int h = 17;
                h = (h * 31) + X;
                h = (h * 31) + Y;
                h = (h * 31) + Z;
                return h;
            }
        }
    }

    /// <summary>
    /// Uniform-grid spatial hash for AABB candidates. Designed for fast "zone AABB -> candidate target indices" queries.
    /// </summary>
    public sealed class SpatialHashGrid
    {
        private readonly double _cellSize;
        private readonly double _invCellSize;
        private readonly Aabb _worldBounds;
        private readonly Dictionary<CellKey, List<int>> _cells = new Dictionary<CellKey, List<int>>(1024);

        public Aabb[] TargetBounds { get; }

        public SpatialHashGrid(Aabb worldBounds, double cellSize, Aabb[] targetBounds)
        {
            _worldBounds = worldBounds;
            _cellSize = Math.Max(1e-6, cellSize);
            _invCellSize = 1.0 / _cellSize;
            TargetBounds = targetBounds ?? throw new ArgumentNullException(nameof(targetBounds));

            Build();
        }

        public double CellSize => _cellSize;

        private void Build()
        {
            for (int i = 0; i < TargetBounds.Length; i++)
            {
                Add(i, TargetBounds[i]);
            }
        }

        private int ToCell(double v, double origin) => (int)Math.Floor((v - origin) * _invCellSize);

        private void Add(int index, in Aabb bounds)
        {
            int minX = ToCell(bounds.MinX, _worldBounds.MinX);
            int minY = ToCell(bounds.MinY, _worldBounds.MinY);
            int minZ = ToCell(bounds.MinZ, _worldBounds.MinZ);
            int maxX = ToCell(bounds.MaxX, _worldBounds.MinX);
            int maxY = ToCell(bounds.MaxY, _worldBounds.MinY);
            int maxZ = ToCell(bounds.MaxZ, _worldBounds.MinZ);

            for (int x = minX; x <= maxX; x++)
            for (int y = minY; y <= maxY; y++)
            for (int z = minZ; z <= maxZ; z++)
            {
                var key = new CellKey(x, y, z);
                if (!_cells.TryGetValue(key, out var list))
                {
                    list = new List<int>(16);
                    _cells[key] = list;
                }
                list.Add(index);
            }
        }

        /// <summary>
        /// Counts unique candidate targets overlapping the query AABB (AABB check included).
        /// visited[] is an int stamp array: visited[i] == stamp means "already seen".
        /// </summary>
        public int CountCandidates(in Aabb query, int[] visited, ref int stamp)
        {
            if (visited == null) throw new ArgumentNullException(nameof(visited));
            if (visited.Length < TargetBounds.Length) throw new ArgumentException("visited[] must be >= target count");

            stamp++;
            if (stamp == int.MaxValue)
            {
                Array.Clear(visited, 0, visited.Length);
                stamp = 1;
            }

            int minX = ToCell(query.MinX, _worldBounds.MinX);
            int minY = ToCell(query.MinY, _worldBounds.MinY);
            int minZ = ToCell(query.MinZ, _worldBounds.MinZ);
            int maxX = ToCell(query.MaxX, _worldBounds.MinX);
            int maxY = ToCell(query.MaxY, _worldBounds.MinY);
            int maxZ = ToCell(query.MaxZ, _worldBounds.MinZ);

            int count = 0;

            for (int x = minX; x <= maxX; x++)
            for (int y = minY; y <= maxY; y++)
            for (int z = minZ; z <= maxZ; z++)
            {
                if (!_cells.TryGetValue(new CellKey(x, y, z), out var list))
                    continue;

                for (int k = 0; k < list.Count; k++)
                {
                    int idx = list[k];
                    if (visited[idx] == stamp) continue;
                    visited[idx] = stamp;

                    if (query.Intersects(TargetBounds[idx]))
                        count++;
                }
            }

            return count;
        }

        /// <summary>
        /// Visits unique candidate indices overlapping query AABB (AABB check included).
        /// </summary>
        public void VisitCandidates(in Aabb query, int[] visited, ref int stamp, Action<int> onCandidate)
        {
            if (onCandidate == null) throw new ArgumentNullException(nameof(onCandidate));
            if (visited == null) throw new ArgumentNullException(nameof(visited));
            if (visited.Length < TargetBounds.Length) throw new ArgumentException("visited[] must be >= target count");

            stamp++;
            if (stamp == int.MaxValue)
            {
                Array.Clear(visited, 0, visited.Length);
                stamp = 1;
            }

            int minX = ToCell(query.MinX, _worldBounds.MinX);
            int minY = ToCell(query.MinY, _worldBounds.MinY);
            int minZ = ToCell(query.MinZ, _worldBounds.MinZ);
            int maxX = ToCell(query.MaxX, _worldBounds.MinX);
            int maxY = ToCell(query.MaxY, _worldBounds.MinY);
            int maxZ = ToCell(query.MaxZ, _worldBounds.MinZ);

            for (int x = minX; x <= maxX; x++)
            for (int y = minY; y <= maxY; y++)
            for (int z = minZ; z <= maxZ; z++)
            {
                if (!_cells.TryGetValue(new CellKey(x, y, z), out var list))
                    continue;

                for (int k = 0; k < list.Count; k++)
                {
                    int idx = list[k];
                    if (visited[idx] == stamp) continue;
                    visited[idx] = stamp;

                    if (query.Intersects(TargetBounds[idx]))
                        onCandidate(idx);
                }
            }
        }

        public int CountPointCandidates(double x, double y, double z)
        {
            var key = new CellKey(
                ToCell(x, _worldBounds.MinX),
                ToCell(y, _worldBounds.MinY),
                ToCell(z, _worldBounds.MinZ));

            return _cells.TryGetValue(key, out var list) ? list.Count : 0;
        }

        public void VisitPointCandidates(double x, double y, double z, Action<int> onCandidate)
        {
            if (onCandidate == null) throw new ArgumentNullException(nameof(onCandidate));

            var key = new CellKey(
                ToCell(x, _worldBounds.MinX),
                ToCell(y, _worldBounds.MinY),
                ToCell(z, _worldBounds.MinZ));

            if (!_cells.TryGetValue(key, out var list))
            {
                return;
            }

            for (int k = 0; k < list.Count; k++)
            {
                onCandidate(list[k]);
            }
        }
    }
}
