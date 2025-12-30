using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using WpfFlyout = Wpf.Ui.Controls.Flyout;

namespace MicroEng.Navisworks
{
    public partial class SpaceMapperStepProcessingPage : Page
    {
        private const double VideoCornerRadius = 10;
        private static readonly string MediaRoot = Path.Combine(
            Path.GetDirectoryName(typeof(SpaceMapperStepProcessingPage).Assembly.Location)
                ?? AppDomain.CurrentDomain.BaseDirectory,
            "ReferenceMedia",
            "SpaceMapper");
        private const string Offset3DVideoFile = "Zone_3D_Offset_720p.mp4";
        private const string OffsetTopVideoFile = "Zone_Top_720p.mp4";
        private const string OffsetBottomVideoFile = "Zone_Bottom_720p.mp4";
        private const string OffsetSidesVideoFile = "Zone_Sides_720p.mp4";

        private bool _offset3dVideoLoaded;
        private bool _offsetTopVideoLoaded;
        private bool _offsetBottomVideoLoaded;
        private bool _offsetSidesVideoLoaded;

        public SpaceMapperStepProcessingPage(SpaceMapperControl host)
        {
            var sw = Stopwatch.StartNew();
            InitializeComponent();
            sw.Stop();
            MicroEngActions.Log($"SpaceMapper Processing: InitializeComponent {sw.ElapsedMilliseconds}ms");
            MicroEngWpfUiTheme.ApplyTo(this);
            DataContext = host;

            _ = new HoverFlyoutController(Offset3DFlyoutButton, Offset3DFlyoutContent, Offset3DFlyout);
            _ = new HoverFlyoutController(OffsetTopFlyoutButton, OffsetTopFlyoutContent, OffsetTopFlyout);
            _ = new HoverFlyoutController(OffsetBottomFlyoutButton, OffsetBottomFlyoutContent, OffsetBottomFlyout);
            _ = new HoverFlyoutController(OffsetSidesFlyoutButton, OffsetSidesFlyoutContent, OffsetSidesFlyout);
            _ = new HoverFlyoutController(PresetHelpFlyoutButton, PresetHelpFlyoutContent, PresetHelpFlyout);
            _ = new HoverFlyoutController(UnitsFlyoutButton, UnitsFlyoutContent, UnitsFlyout);
            _ = new HoverFlyoutController(OffsetModeFlyoutButton, OffsetModeFlyoutContent, OffsetModeFlyout);
        }

        internal CheckBox TreatPartialCheck => TreatPartialCheckControl;
        internal CheckBox TagPartialCheck => TagPartialCheckControl;
        internal CheckBox EnableMultiZoneCheck => EnableMultiZoneCheckControl;
        internal TextBox ZoneBehaviorCategoryBox => ZoneBehaviorCategoryBoxControl;
        internal TextBox ZoneBehaviorPropertyBox => ZoneBehaviorPropertyBoxControl;
        internal TextBox ZoneBehaviorContainedBox => ZoneBehaviorContainedBoxControl;
        internal TextBox ZoneBehaviorPartialBox => ZoneBehaviorPartialBoxControl;
        internal TextBox Offset3DBox => Offset3DBoxControl;
        internal TextBox OffsetTopBox => OffsetTopBoxControl;
        internal TextBox OffsetBottomBox => OffsetBottomBoxControl;
        internal TextBox OffsetSidesBox => OffsetSidesBoxControl;
        internal TextBox UnitsBox => UnitsBoxControl;
        internal TextBox OffsetModeBox => OffsetModeBoxControl;
        internal TextBox MaxThreadsBox => MaxThreadsBoxControl;
        internal TextBox BatchSizeBox => BatchSizeBoxControl;
        internal CheckBox AutoPresetToggle => AutoPresetToggleControl;
        internal Slider PresetSlider => PresetSliderControl;
        internal TextBlock PresetDescriptionText => PresetDescriptionTextControl;
        internal TextBlock PresetResolvedText => PresetResolvedTextControl;
        internal Slider IndexGranularitySlider => IndexGranularitySliderControl;
        internal TextBlock IndexGranularityHintText => IndexGranularityHintTextControl;
        internal TextBlock EstimateRuntimeText => EstimateRuntimeTextControl;
        internal TextBlock EstimateCountsText => EstimateCountsTextControl;
        internal TextBlock EstimatePairsText => EstimatePairsTextControl;
        internal TextBlock EstimateAvgLabelText => EstimateAvgLabelTextControl;
        internal TextBlock EstimateAvgText => EstimateAvgTextControl;
        internal TextBlock EstimateConfidenceText => EstimateConfidenceTextControl;
        internal TextBlock EstimateCellSizeText => EstimateCellSizeTextControl;
        internal TextBlock EstimatePreflightTimeText => EstimatePreflightTimeTextControl;
        internal TextBlock PreflightStatusText => PreflightStatusTextControl;
        internal ProgressBar PreflightProgressBar => PreflightProgressBarControl;
        internal Wpf.Ui.Controls.Button RunPreflightButton => RunPreflightButtonControl;
        internal CheckBox LiveEstimateToggle => LiveEstimateToggleControl;
        internal CheckBox ReusePreflightToggle => ReusePreflightToggleControl;
        internal Wpf.Ui.Controls.Button ComparisonReportButton => ComparisonReportButtonControl;
        internal TextBlock ComparisonReportStatusText => ComparisonReportStatusTextControl;
        internal ComboBox BenchmarkModeCombo => BenchmarkModeComboControl;
        internal TextBlock BenchmarkModeHintText => BenchmarkModeHintTextControl;
        internal ComboBox WritebackStrategyCombo => WritebackStrategyComboControl;
        internal CheckBox SkipUnchangedWritebackCheck => SkipUnchangedWritebackCheckControl;
        internal CheckBox PackWritebackCheck => PackWritebackCheckControl;
        internal CheckBox ShowInternalWritebackCheck => ShowInternalWritebackCheckControl;
        internal CheckBox CloseDockPanesCheck => CloseDockPanesCheckControl;
        internal TextBlock BenchmarkSummaryText => BenchmarkSummaryTextControl;
        internal TextBlock BenchmarkSummaryDetailText => BenchmarkSummaryDetailTextControl;
        internal Expander AdvancedPerformanceExpander => AdvancedPerformanceExpanderControl;
        internal ComboBox FastTraversalCombo => FastTraversalComboControl;
        internal TextBlock FastTraversalHintText => FastTraversalHintTextControl;
        internal TextBlock FastTraversalResolvedText => FastTraversalResolvedTextControl;

        private void OnOffset3DFlyoutOpened(WpfFlyout sender, RoutedEventArgs args)
        {
            EnsureVideoLoaded(Offset3DVideo, Offset3DVideoFile, ref _offset3dVideoLoaded, "Offset3D");
        }

        private void OnOffset3DFlyoutClosed(WpfFlyout sender, RoutedEventArgs args)
        {
            StopVideo(Offset3DVideo);
        }

        private void OnOffsetTopFlyoutOpened(WpfFlyout sender, RoutedEventArgs args)
        {
            EnsureVideoLoaded(OffsetTopVideo, OffsetTopVideoFile, ref _offsetTopVideoLoaded, "OffsetTop");
        }

        private void OnOffsetTopFlyoutClosed(WpfFlyout sender, RoutedEventArgs args)
        {
            StopVideo(OffsetTopVideo);
        }

        private void OnOffsetBottomFlyoutOpened(WpfFlyout sender, RoutedEventArgs args)
        {
            EnsureVideoLoaded(OffsetBottomVideo, OffsetBottomVideoFile, ref _offsetBottomVideoLoaded, "OffsetBottom");
        }

        private void OnOffsetBottomFlyoutClosed(WpfFlyout sender, RoutedEventArgs args)
        {
            StopVideo(OffsetBottomVideo);
        }

        private void OnOffsetSidesFlyoutOpened(WpfFlyout sender, RoutedEventArgs args)
        {
            EnsureVideoLoaded(OffsetSidesVideo, OffsetSidesVideoFile, ref _offsetSidesVideoLoaded, "OffsetSides");
        }

        private void OnOffsetSidesFlyoutClosed(WpfFlyout sender, RoutedEventArgs args)
        {
            StopVideo(OffsetSidesVideo);
        }

        private static void EnsureVideoLoaded(MediaElement target, string fileName, ref bool loaded, string label)
        {
            if (target == null)
            {
                return;
            }

            var path = Path.Combine(MediaRoot, fileName);
            if (!File.Exists(path))
            {
                MicroEngActions.Log($"SpaceMapper Processing: {label} video missing at {path}");
                return;
            }

            if (!loaded)
            {
                target.Source = new Uri(path, UriKind.Absolute);
                target.Tag = new VideoLoadState(label, Stopwatch.StartNew());
                loaded = true;
            }

            target.Position = TimeSpan.Zero;
            target.Play();
        }

        private static void StopVideo(MediaElement target)
        {
            if (target == null)
            {
                return;
            }

            target.Stop();
        }

        private void OnHelpVideoOpened(object sender, RoutedEventArgs args)
        {
            if (sender is not MediaElement media)
            {
                return;
            }

            if (media.Tag is not VideoLoadState state)
            {
                return;
            }

            state.Stopwatch.Stop();
            MicroEngActions.Log($"SpaceMapper Processing: {state.Label} video opened in {state.Stopwatch.ElapsedMilliseconds}ms");
            media.Tag = null;
        }

        private void OnHelpVideoLoaded(object sender, RoutedEventArgs args)
        {
            ApplyRoundedClip(sender as MediaElement);
        }

        private void OnHelpVideoSizeChanged(object sender, SizeChangedEventArgs args)
        {
            ApplyRoundedClip(sender as MediaElement);
        }

        private void OnHelpVideoEnded(object sender, RoutedEventArgs args)
        {
            if (sender is not MediaElement media)
            {
                return;
            }

            media.Position = TimeSpan.Zero;
            media.Play();
        }

        private void OnHelpVideoFailed(object sender, ExceptionRoutedEventArgs args)
        {
            MicroEngActions.Log($"SpaceMapper Processing: video failed - {args.ErrorException?.Message}");
        }

        private static void ApplyRoundedClip(MediaElement media)
        {
            if (media == null || media.ActualWidth <= 0 || media.ActualHeight <= 0)
            {
                return;
            }

            var rect = new Rect(0, 0, media.ActualWidth, media.ActualHeight);
            media.Clip = new RectangleGeometry(rect, VideoCornerRadius, VideoCornerRadius);
        }

        private sealed class VideoLoadState
        {
            public VideoLoadState(string label, Stopwatch stopwatch)
            {
                Label = label;
                Stopwatch = stopwatch;
            }

            public string Label { get; }
            public Stopwatch Stopwatch { get; }
        }
    }
}
