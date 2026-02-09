using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Threading;
using Wpf.Ui.Controls;

namespace MicroEng.Navisworks
{
    internal partial class DataScraperRunProgressWindow : FluentWindow
    {
        private readonly DataScraperRunProgressState _state;
        private readonly DispatcherTimer _timer;
        private bool _allowClose;

        internal DataScraperRunProgressWindow(DataScraperRunProgressState state)
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(this);

            _state = state ?? throw new ArgumentNullException(nameof(state));
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
                return;
            }

            base.OnClosing(e);
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            // Cancellation not supported for Data Scraper; keep button disabled.
        }

        private void RefreshUi()
        {
            StageTextBlock.Text = string.IsNullOrWhiteSpace(_state.StageText) ? "Working..." : _state.StageText;
            DetailTextBlock.Text = string.IsNullOrWhiteSpace(_state.DetailText) ? "-" : _state.DetailText;

            var elapsed = _state.Elapsed;
            var stageElapsed = DateTimeOffset.UtcNow - _state.StageStartUtc;
            var lastProgressAge = DateTimeOffset.UtcNow - _state.LastProgressUtc;

            StatsTextBlock.Text = $"Elapsed: {elapsed:hh\\:mm\\:ss}\nStage: {stageElapsed:hh\\:mm\\:ss}   |   Last progress: {lastProgressAge:hh\\:mm\\:ss} ago";

            if (_state.IsFinished)
            {
                _timer.Stop();
            }
        }
    }
}
