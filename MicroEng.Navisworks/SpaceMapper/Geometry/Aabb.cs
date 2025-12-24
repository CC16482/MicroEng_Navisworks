using System;

namespace MicroEng.Navisworks.SpaceMapper.Geometry
{
    /// <summary>Axis-aligned bounding box (AABB) in model/world coordinates.</summary>
    public readonly struct Aabb
    {
        public readonly double MinX, MinY, MinZ;
        public readonly double MaxX, MaxY, MaxZ;

        public Aabb(double minX, double minY, double minZ, double maxX, double maxY, double maxZ)
        {
            MinX = Math.Min(minX, maxX);
            MinY = Math.Min(minY, maxY);
            MinZ = Math.Min(minZ, maxZ);
            MaxX = Math.Max(minX, maxX);
            MaxY = Math.Max(minY, maxY);
            MaxZ = Math.Max(minZ, maxZ);
        }

        public double SizeX => MaxX - MinX;
        public double SizeY => MaxY - MinY;
        public double SizeZ => MaxZ - MinZ;

        public bool Intersects(in Aabb other)
        {
            return !(other.MinX > MaxX || other.MaxX < MinX
                  || other.MinY > MaxY || other.MaxY < MinY
                  || other.MinZ > MaxZ || other.MaxZ < MinZ);
        }

        public Aabb Inflate(double uniform)
            => Inflate(uniform, uniform, uniform);

        public Aabb Inflate(double dx, double dy, double dz)
            => new Aabb(MinX - dx, MinY - dy, MinZ - dz, MaxX + dx, MaxY + dy, MaxZ + dz);

        public Aabb OffsetZ(double bottom, double top)
            => new Aabb(MinX, MinY, MinZ - bottom, MaxX, MaxY, MaxZ + top);
    }
}
