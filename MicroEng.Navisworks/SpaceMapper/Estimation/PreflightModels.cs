using System;

using MicroEng.Navisworks;

namespace MicroEng.Navisworks.SpaceMapper.Estimation
{
    public sealed class SpaceMapperPreflightResult
    {
        public DateTime TimestampUtc { get; set; } = DateTime.UtcNow;

        public int ZoneCount { get; set; }
        public int TargetCount { get; set; }

        public long CandidatePairs { get; set; }
        public int MaxCandidatesPerZone { get; set; }
        public double AvgCandidatesPerZone { get; set; }
        public int MaxCandidatesPerTarget { get; set; }
        public double AvgCandidatesPerTarget { get; set; }

        public double CellSizeUsed { get; set; }

        public TimeSpan BuildIndexTime { get; set; }
        public TimeSpan QueryTime { get; set; }

        public string Signature { get; set; }
        public SpaceMapperFastTraversalMode TraversalUsed { get; set; } = SpaceMapperFastTraversalMode.ZoneMajor;
    }

    public sealed class SpaceMapperRuntimeEstimate
    {
        public DateTime TimestampUtc { get; set; } = DateTime.UtcNow;

        public TimeSpan EstimatedTotal { get; set; }
        public TimeSpan EstimatedCompute { get; set; }
        public TimeSpan EstimatedWrite { get; set; }

        public double Confidence01 { get; set; }
        public string ConfidenceLabel { get; set; }

        public string Notes { get; set; }
    }
}
