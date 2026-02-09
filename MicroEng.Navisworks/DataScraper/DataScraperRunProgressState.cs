using System;

namespace MicroEng.Navisworks
{
    internal sealed class DataScraperRunProgressState
    {
        private DateTimeOffset _startUtc = DateTimeOffset.UtcNow;
        private DateTimeOffset _stageStartUtc = DateTimeOffset.UtcNow;
        private DateTimeOffset _lastProgressUtc = DateTimeOffset.UtcNow;

        public string StageText { get; private set; } = "Starting...";
        public string DetailText { get; private set; } = "-";
        public bool IsFinished { get; private set; }

        public TimeSpan Elapsed => DateTimeOffset.UtcNow - _startUtc;
        public DateTimeOffset StageStartUtc => _stageStartUtc;
        public DateTimeOffset LastProgressUtc => _lastProgressUtc;

        public void SetStage(string stageText, string detailText = null)
        {
            StageText = string.IsNullOrWhiteSpace(stageText) ? StageText : stageText;
            DetailText = string.IsNullOrWhiteSpace(detailText) ? "-" : detailText;
            _stageStartUtc = DateTimeOffset.UtcNow;
            _lastProgressUtc = _stageStartUtc;
        }

        public void TouchProgress()
        {
            _lastProgressUtc = DateTimeOffset.UtcNow;
        }

        public void MarkFinished()
        {
            IsFinished = true;
            _lastProgressUtc = DateTimeOffset.UtcNow;
        }
    }
}
