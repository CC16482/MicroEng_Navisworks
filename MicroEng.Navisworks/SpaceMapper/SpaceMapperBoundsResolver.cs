using System;

namespace MicroEng.Navisworks
{
    internal static class SpaceMapperBoundsResolver
    {
        internal static SpaceMapperTargetBoundsMode ResolveTargetBoundsMode(
            SpaceMapperProcessingSettings settings,
            SpaceMapperZoneContainmentEngine containmentEngine)
        {
            if (settings == null)
            {
                return SpaceMapperTargetBoundsMode.Aabb;
            }

            // Tier A/B: do not force Midpoint when MeshAccurate is selected.
            _ = containmentEngine;

            if (settings.TargetBoundsMode == SpaceMapperTargetBoundsMode.Aabb && settings.UseOriginPointOnly)
            {
                return SpaceMapperTargetBoundsMode.Midpoint;
            }

            return settings.TargetBoundsMode;
        }
    }
}
