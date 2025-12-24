using System;

namespace MicroEng.Navisworks.SpaceMapper.Estimation
{
    internal static class SpaceMapperPresetLogic
    {
        public static SpaceMapperPerformancePreset ResolvePreset(SpaceMapperPerformancePreset preset, SpaceMapperPreflightResult preflight, out string reason)
        {
            reason = string.Empty;
            if (preset != SpaceMapperPerformancePreset.Auto)
            {
                return preset;
            }

            if (preflight == null)
            {
                reason = "Auto: run preflight to resolve.";
                return SpaceMapperPerformancePreset.Normal;
            }

            if (preflight.CandidatePairs >= 20000000L || preflight.TargetCount >= 250000)
            {
                reason = $"Auto resolved: Fast (pairs: {preflight.CandidatePairs:N0})";
                return SpaceMapperPerformancePreset.Fast;
            }

            if (preflight.CandidatePairs <= 2000000L && preflight.ZoneCount <= 5000)
            {
                reason = $"Auto resolved: Accurate (pairs: {preflight.CandidatePairs:N0})";
                return SpaceMapperPerformancePreset.Accurate;
            }

            reason = $"Auto resolved: Normal (pairs: {preflight.CandidatePairs:N0})";
            return SpaceMapperPerformancePreset.Normal;
        }

        public static SpaceMapperPerformancePreset ResolvePreset(SpaceMapperProcessingSettings settings, SpaceMapperPreflightResult preflight, out string reason)
        {
            var preset = settings?.PerformancePreset ?? SpaceMapperPerformancePreset.Auto;
            return ResolvePreset(preset, preflight, out reason);
        }

        public static string DescribePreset(SpaceMapperPerformancePreset preset)
        {
            switch (preset)
            {
                case SpaceMapperPerformancePreset.Fast:
                    return "Fast: spatial grid + AABB-only classification. Best for quick tagging.";
                case SpaceMapperPerformancePreset.Normal:
                    return "Normal: spatial grid + bbox corner sampling. Balanced default.";
                case SpaceMapperPerformancePreset.Accurate:
                    return "Accurate: spatial grid + extra sampling points. Slowest, best quality.";
                case SpaceMapperPerformancePreset.Auto:
                    return "Auto: chooses Fast/Normal/Accurate using preflight density and model size.";
                default:
                    return string.Empty;
            }
        }
    }
}
