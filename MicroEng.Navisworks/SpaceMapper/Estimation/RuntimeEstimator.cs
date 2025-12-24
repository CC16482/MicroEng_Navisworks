using System;

namespace MicroEng.Navisworks.SpaceMapper.Estimation
{
    public static class SpaceMapperRuntimeEstimator
    {
        public static SpaceMapperRuntimeEstimate Estimate(
            SpaceMapperPreflightResult preflight,
            SpaceMapperCalibrationEntry calibration,
            int mappingColumnsCount,
            double expectedAssignmentRate = 0.10)
        {
            if (preflight == null) throw new ArgumentNullException(nameof(preflight));
            if (calibration == null) throw new ArgumentNullException(nameof(calibration));

            double computeSeconds = calibration.FixedSeconds + (preflight.CandidatePairs * calibration.SecondsPerCandidatePair);

            double estimatedAssignments = Math.Max(1, preflight.TargetCount * expectedAssignmentRate);
            double writes = estimatedAssignments * Math.Max(1, mappingColumnsCount);
            double writeSeconds = writes * calibration.SecondsPerWrite;

            double totalSeconds = computeSeconds + writeSeconds;

            double conf = Math.Min(1.0, 0.2 + 0.15 * calibration.Samples);
            var label = conf < 0.45 ? "Low" : (conf < 0.75 ? "Medium" : "High");

            return new SpaceMapperRuntimeEstimate
            {
                EstimatedCompute = TimeSpan.FromSeconds(computeSeconds),
                EstimatedWrite = TimeSpan.FromSeconds(writeSeconds),
                EstimatedTotal = TimeSpan.FromSeconds(totalSeconds),
                Confidence01 = conf,
                ConfidenceLabel = label,
                Notes = calibration.Samples < 2 ? "Using default heuristics + preflight candidate pairs" : $"Calibrated from {calibration.Samples} run(s)"
            };
        }
    }
}
