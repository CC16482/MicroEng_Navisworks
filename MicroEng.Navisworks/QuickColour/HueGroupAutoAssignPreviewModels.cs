namespace MicroEng.Navisworks.QuickColour
{
    public sealed class HueGroupAutoAssignPreviewRow : NotifyBase
    {
        private string _category = "";
        private int _count;
        private int _typeCount;

        private bool _locked;
        private string _currentGroup = "";
        private string _proposedGroup = "";

        private string _status = "";
        private string _reason = "";

        private string _matchedRuleGroup = "";
        private string _matchedMatcherType = "";
        private string _matchedMatcherValue = "";

        private bool _willChange;

        public string Category { get => _category; set => SetField(ref _category, value ?? ""); }
        public int Count { get => _count; set => SetField(ref _count, value); }
        public int TypeCount { get => _typeCount; set => SetField(ref _typeCount, value); }

        public bool Locked { get => _locked; set => SetField(ref _locked, value); }
        public string CurrentGroup { get => _currentGroup; set => SetField(ref _currentGroup, value ?? ""); }
        public string ProposedGroup { get => _proposedGroup; set => SetField(ref _proposedGroup, value ?? ""); }

        public string Status { get => _status; set => SetField(ref _status, value ?? ""); }
        public string Reason { get => _reason; set => SetField(ref _reason, value ?? ""); }

        public string MatchedRuleGroup { get => _matchedRuleGroup; set => SetField(ref _matchedRuleGroup, value ?? ""); }
        public string MatchedMatcherType { get => _matchedMatcherType; set => SetField(ref _matchedMatcherType, value ?? ""); }
        public string MatchedMatcherValue { get => _matchedMatcherValue; set => SetField(ref _matchedMatcherValue, value ?? ""); }

        public bool WillChange { get => _willChange; set => SetField(ref _willChange, value); }
    }

    public sealed class HueGroupAutoAssignPreviewResult
    {
        public int Total;
        public int WillChange;
        public int NoChange;
        public int SkippedLocked;
        public int SkippedAlreadyAssigned;
        public int UnmatchedFallback;
        public int MissingGroupFallback;
        public int WillCreateMissingGroup;
    }
}
