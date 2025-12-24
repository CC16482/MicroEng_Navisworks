using System;

namespace MicroEng.Navisworks.SpaceMapper.Geometry
{
    internal static class SpatialGridSizing
    {
        public static double ComputeAutoCellSize(Aabb worldBounds)
        {
            var extent = Math.Max(worldBounds.SizeX, Math.Max(worldBounds.SizeY, worldBounds.SizeZ));
            return Math.Max(1000.0, extent / 128.0);
        }

        public static double ComputeCellSize(Aabb worldBounds, int indexGranularity)
        {
            var autoCell = ComputeAutoCellSize(worldBounds);
            var multiplier = GetGranularityMultiplier(indexGranularity);
            return autoCell * multiplier;
        }

        public static double GetGranularityMultiplier(int indexGranularity)
        {
            switch (indexGranularity)
            {
                case 1:
                    return 2.0;
                case 2:
                    return 1.0;
                case 3:
                    return 0.5;
                case 4:
                    return 0.25;
                default:
                    return 1.0;
            }
        }
    }
}
