using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Threading;
using Wpf.Ui.Controls;

namespace MicroEng.Navisworks
{
    internal partial class SpaceMapperRunProgressWindow : FluentWindow
    {
        private readonly SpaceMapperRunProgressState _state;
        private readonly Action _cancelAction;
        private readonly DispatcherTimer _timer;
        private bool _allowClose;

        internal SpaceMapperRunProgressWindow(SpaceMapperRunProgressState state, Action cancelAction)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(this);

            _state = state ?? throw new ArgumentNullException(nameof(state));
            _cancelAction = cancelAction;

            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
            _timer.Tick += (_, __) => RefreshUi();
            _timer.Start();

            RefreshUi();
        }

        public void AllowClose()
        {
            _allowClose = true;
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (!_allowClose && !_state.IsFinished)
            {
                e.Cancel = true;
                TriggerCancel();
                return;
            }

            base.OnClosing(e);
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            TriggerCancel();
        }

        private void TriggerCancel()
        {
            if (_state.IsFinished)
            {
                return;
            }

            CancelButton.IsEnabled = false;
            _state.SetStage(
                SpaceMapperRunStage.Cancelled,
                "Cancelling...",
                "Waiting for the current operation to stop...");
            _cancelAction?.Invoke();
        }

        private void RefreshUi()
        {
            StageTextBlock.Text = _state.StageText ?? _state.Stage.ToString();

            var detail = _state.DetailText;
            if (string.IsNullOrWhiteSpace(detail))
            {
                detail = "-";
            }
            DetailTextBlock.Text = detail;

            var elapsed = _state.Elapsed;
            var elapsedText = $"{elapsed:hh\\:mm\\:ss}";

            var zonesLine = _state.ZonesTotal > 0
                ? $"Zones: {_state.ZonesProcessed}/{_state.ZonesTotal}"
                : "Zones: -";

            var targetsLine = _state.TargetsTotal > 0
                ? $"Targets: {_state.TargetsProcessed}/{_state.TargetsTotal}"
                : "Targets: -";

            var writeLine = _state.WriteTargetsTotal > 0
                ? $"Writeback: {_state.WriteTargetsProcessed}/{_state.WriteTargetsTotal}"
                : "Writeback: -";

            var stageElapsed = DateTimeOffset.UtcNow - _state.StageStartUtc;
            var lastProgressAge = DateTimeOffset.UtcNow - _state.LastProgressUtc;
            var stageLine = $"Stage: {FormatDuration(stageElapsed)}";
            var progressLine = $"Last progress: {FormatDuration(lastProgressAge)} ago";

            StatsTextBlock.Text = $"{zonesLine}   |   {targetsLine}\n{writeLine}   |   Elapsed: {elapsedText}\n{stageLine}   |   {progressLine}";

            if (_state.IsFinished)
            {
                CancelButton.IsEnabled = false;
                _timer.Stop();
            }
        }

        private static string FormatDuration(TimeSpan span)
        {
            if (span < TimeSpan.Zero)
            {
                span = TimeSpan.Zero;
            }

            return span.ToString("hh\\:mm\\:ss");
        }
    }
}
