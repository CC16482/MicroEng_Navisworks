using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace MicroEng.Navisworks
{
    internal enum SpaceMapperRunStage
    {
        Starting,
        ResolvingInputs,
        ExtractingGeometry,
        BuildingIndex,
        ComputingIntersections,
        ResolvingAssignments,
        WritingProperties,
        Finalizing,
        Completed,
        Cancelled,
        Failed
    }

    internal sealed class SpaceMapperStageTiming
    {
        public SpaceMapperRunStage Stage { get; set; }
        public string StageText { get; set; }
        public string DetailText { get; set; }
        public DateTimeOffset StartUtc { get; set; }
        public DateTimeOffset? EndUtc { get; set; }
    }

    /// <summary>
    /// Thread-safe progress state. Writers update via Interlocked, UI polls periodically.
    /// </summary>
    internal sealed class SpaceMapperRunProgressState
    {
        public volatile SpaceMapperRunStage Stage = SpaceMapperRunStage.Starting;
        public volatile string StageText = "Starting...";
        public volatile string DetailText = string.Empty;

        public int ZonesTotal;
        public int TargetsTotal;

        public int ZonesProcessed;
        public int TargetsProcessed;

        public int WriteTargetsTotal;
        public int WriteTargetsProcessed;

        public long CandidatePairs;

        public volatile bool IsFinished;
        public volatile bool IsFailed;
        public volatile bool IsCancelled;
        public volatile string ErrorText = string.Empty;

        private readonly Stopwatch _sw = new Stopwatch();
        private readonly object _timelineLock = new object();
        private readonly List<SpaceMapperStageTiming> _timeline = new List<SpaceMapperStageTiming>();
        private DateTimeOffset _stageStartUtc;
        private DateTimeOffset _lastProgressUtc;
        private DateTimeOffset _lastZoneProgressUtc;
        private DateTimeOffset _lastTargetProgressUtc;
        private DateTimeOffset _lastWriteProgressUtc;

        public void Start()
        {
            ZonesTotal = 0;
            TargetsTotal = 0;
            CandidatePairs = 0;

            Interlocked.Exchange(ref ZonesProcessed, 0);
            Interlocked.Exchange(ref TargetsProcessed, 0);
            Interlocked.Exchange(ref WriteTargetsProcessed, 0);
            WriteTargetsTotal = 0;

            IsFinished = false;
            IsFailed = false;
            IsCancelled = false;
            ErrorText = string.Empty;

            Stage = SpaceMapperRunStage.Starting;
            StageText = "Starting...";
            DetailText = string.Empty;

            _sw.Restart();
            var now = DateTimeOffset.UtcNow;
            _stageStartUtc = now;
            _lastProgressUtc = now;
            _lastZoneProgressUtc = now;
            _lastTargetProgressUtc = now;
            _lastWriteProgressUtc = now;

            lock (_timelineLock)
            {
                _timeline.Clear();
                _timeline.Add(new SpaceMapperStageTiming
                {
                    Stage = Stage,
                    StageText = StageText,
                    DetailText = DetailText,
                    StartUtc = now
                });
            }
        }

        public TimeSpan Elapsed => _sw.Elapsed;

        public void SetStage(SpaceMapperRunStage stage, string stageText, string detailText = null)
        {
            Stage = stage;
            StageText = stageText ?? stage.ToString();
            if (detailText != null)
            {
                DetailText = detailText;
            }

            RecordStageChange();
        }

        public void SetTotals(int zonesTotal, int targetsTotal)
        {
            ZonesTotal = zonesTotal;
            TargetsTotal = targetsTotal;
            TouchProgress();
        }

        public void SetWriteTotals(int writeTargetsTotal)
        {
            WriteTargetsTotal = writeTargetsTotal;
            TouchProgress();
        }

        public void UpdateDetail(string detailText)
        {
            if (detailText == null)
            {
                return;
            }

            DetailText = detailText;
            TouchProgress();
            UpdateCurrentStageDetail(detailText);
        }

        public void UpdateZonesProcessed(int processed)
        {
            Interlocked.Exchange(ref ZonesProcessed, processed);
            TouchProgress();
            _lastZoneProgressUtc = _lastProgressUtc;
        }

        public void UpdateTargetsProcessed(int processed)
        {
            Interlocked.Exchange(ref TargetsProcessed, processed);
            TouchProgress();
            _lastTargetProgressUtc = _lastProgressUtc;
        }

        public void UpdateWriteProcessed(int processed)
        {
            Interlocked.Exchange(ref WriteTargetsProcessed, processed);
            TouchProgress();
            _lastWriteProgressUtc = _lastProgressUtc;
        }

        public int IncrementWriteProcessed()
        {
            var value = Interlocked.Increment(ref WriteTargetsProcessed);
            TouchProgress();
            _lastWriteProgressUtc = _lastProgressUtc;
            return value;
        }

        public void UpdateCandidatePairs(long candidatePairs)
        {
            Interlocked.Exchange(ref CandidatePairs, candidatePairs);
            TouchProgress();
        }

        public void MarkCompleted()
        {
            if (Stage != SpaceMapperRunStage.Completed)
            {
                Stage = SpaceMapperRunStage.Completed;
                StageText = "Completed";
                RecordStageChange();
            }
            else
            {
                TouchProgress();
                CloseCurrentStage();
            }
            IsFinished = true;
        }

        public void MarkCancelled()
        {
            if (Stage != SpaceMapperRunStage.Cancelled)
            {
                Stage = SpaceMapperRunStage.Cancelled;
                StageText = "Cancelled";
                RecordStageChange();
            }
            else
            {
                TouchProgress();
                CloseCurrentStage();
            }
            IsFinished = true;
            IsCancelled = true;
        }

        public void MarkFailed(Exception ex)
        {
            if (Stage != SpaceMapperRunStage.Failed)
            {
                Stage = SpaceMapperRunStage.Failed;
                StageText = "Failed";
                RecordStageChange();
            }
            else
            {
                TouchProgress();
                CloseCurrentStage();
            }
            IsFinished = true;
            IsFailed = true;
            ErrorText = ex?.Message ?? "Unknown error";
        }

        public IReadOnlyList<SpaceMapperStageTiming> GetTimelineSnapshot()
        {
            lock (_timelineLock)
            {
                return _timeline
                    .Select(t => new SpaceMapperStageTiming
                    {
                        Stage = t.Stage,
                        StageText = t.StageText,
                        DetailText = t.DetailText,
                        StartUtc = t.StartUtc,
                        EndUtc = t.EndUtc
                    })
                    .ToList();
            }
        }

        public DateTimeOffset StageStartUtc => _stageStartUtc;
        public DateTimeOffset LastProgressUtc => _lastProgressUtc;
        public DateTimeOffset LastZoneProgressUtc => _lastZoneProgressUtc;
        public DateTimeOffset LastTargetProgressUtc => _lastTargetProgressUtc;
        public DateTimeOffset LastWriteProgressUtc => _lastWriteProgressUtc;

        private void TouchProgress()
        {
            _lastProgressUtc = DateTimeOffset.UtcNow;
        }

        private void RecordStageChange()
        {
            var now = DateTimeOffset.UtcNow;
            _stageStartUtc = now;
            _lastProgressUtc = now;

            lock (_timelineLock)
            {
                if (_timeline.Count > 0 && !_timeline[_timeline.Count - 1].EndUtc.HasValue)
                {
                    _timeline[_timeline.Count - 1].EndUtc = now;
                }

                _timeline.Add(new SpaceMapperStageTiming
                {
                    Stage = Stage,
                    StageText = StageText,
                    DetailText = DetailText,
                    StartUtc = now
                });
            }
        }

        private void UpdateCurrentStageDetail(string detailText)
        {
            lock (_timelineLock)
            {
                if (_timeline.Count == 0)
                {
                    return;
                }

                _timeline[_timeline.Count - 1].DetailText = detailText;
            }
        }

        private void CloseCurrentStage()
        {
            var now = DateTimeOffset.UtcNow;
            lock (_timelineLock)
            {
                if (_timeline.Count == 0)
                {
                    return;
                }

                if (!_timeline[_timeline.Count - 1].EndUtc.HasValue)
                {
                    _timeline[_timeline.Count - 1].EndUtc = now;
                }
            }
        }
    }
}
