using System;
using System.Collections.Generic;

namespace MicroEng.Navisworks.SpaceMapper.Estimation
{
    public sealed class SpaceMapperCalibrationEntry
    {
        public int Samples { get; set; }

        public double SecondsPerCandidatePair { get; set; } = 0.00000008;

        public double SecondsPerWrite { get; set; } = 0.00002;

        public double FixedSeconds { get; set; } = 1.0;
    }

    public sealed class SpaceMapperCalibrationStore
    {
        private readonly Dictionary<string, SpaceMapperCalibrationEntry> _entries = new Dictionary<string, SpaceMapperCalibrationEntry>();

        public SpaceMapperCalibrationEntry GetOrCreate(string key)
        {
            if (!_entries.TryGetValue(key, out var e))
            {
                e = new SpaceMapperCalibrationEntry();
                _entries[key] = e;
            }
            return e;
        }

        public void Update(string key, double elapsedSeconds, long candidatePairs, long writes)
        {
            var e = GetOrCreate(key);
            e.Samples++;

            candidatePairs = Math.Max(1, candidatePairs);
            writes = Math.Max(1, writes);

            double fixedS = e.FixedSeconds;
            double writePart = writes * e.SecondsPerWrite;
            double remaining = Math.Max(0.0, elapsedSeconds - fixedS - writePart);
            double kPairsNew = remaining / candidatePairs;

            double alpha = e.Samples < 5 ? 0.35 : 0.15;
            e.SecondsPerCandidatePair = (1 - alpha) * e.SecondsPerCandidatePair + alpha * kPairsNew;
        }
    }
}
