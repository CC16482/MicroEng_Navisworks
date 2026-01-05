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

            WireHover(Offset3DFlyoutButton, Offset3DFlyout);
            WireHover(OffsetTopFlyoutButton, OffsetTopFlyout);
            WireHover(OffsetBottomFlyoutButton, OffsetBottomFlyout);
            WireHover(OffsetSidesFlyoutButton, OffsetSidesFlyout);
            WireHover(UnitsFlyoutButton, UnitsFlyout);
            WireHover(OffsetModeFlyoutButton, OffsetModeFlyout);
            WireHover(EnableMultiZoneHelpToggle, EnableMultiZoneHelpFlyout);
            WireHover(ExcludeZonesFromTargetsHelpToggle, ExcludeZonesFromTargetsHelpFlyout);
            WireHover(TreatPartialHelpToggle, TreatPartialHelpFlyout);
            WireHover(TagPartialHelpToggle, TagPartialHelpFlyout);
            WireHover(SkipUnchangedHelpToggle, SkipUnchangedHelpFlyout);
            WireHover(PackWritebackHelpToggle, PackWritebackHelpFlyout);
            WireHover(ShowInternalHelpToggle, ShowInternalHelpFlyout);
            WireHover(WriteZoneBehaviorHelpToggle, WriteZoneBehaviorHelpFlyout);
            WireHover(WriteContainmentPercentHelpToggle, WriteContainmentPercentHelpFlyout);
            WireHover(ContainmentCalculationHelpToggle, ContainmentCalculationHelpFlyout);
            WireHover(CloseDockPanesHelpToggle, CloseDockPanesHelpFlyout);
            WireHover(DockPaneDelayHelpToggle, DockPaneDelayHelpFlyout);
            WireHover(GpuRayAccuracyHelpToggle, GpuRayAccuracyHelpFlyout);
            HookUiStateUpdates();
            Loaded += (_, __) => UpdateProcessingUiState();
        }

        private void HookUiStateUpdates()
        {
            if (ZoneContainmentEngineComboControl != null)
            {
                ZoneContainmentEngineComboControl.SelectionChanged += (_, __) => UpdateProcessingUiState();
            }

            if (ZoneBoundsSliderControl != null)
            {
                ZoneBoundsSliderControl.ValueChanged += (_, __) => UpdateProcessingUiState();
            }

            if (TargetBoundsSliderControl != null)
            {
                TargetBoundsSliderControl.ValueChanged += (_, __) => UpdateProcessingUiState();
            }

            if (ZoneResolutionStrategyComboControl != null)
            {
                ZoneResolutionStrategyComboControl.SelectionChanged += (_, __) => UpdateProcessingUiState();
            }

            if (EnableMultiZoneCheckControl != null)
            {
                EnableMultiZoneCheckControl.Checked += (_, __) => UpdateProcessingUiState();
                EnableMultiZoneCheckControl.Unchecked += (_, __) => UpdateProcessingUiState();
            }

            if (TreatPartialCheckControl != null)
            {
                TreatPartialCheckControl.Checked += (_, __) => UpdateProcessingUiState();
                TreatPartialCheckControl.Unchecked += (_, __) => UpdateProcessingUiState();
            }

            if (TagPartialCheckControl != null)
            {
                TagPartialCheckControl.Checked += (_, __) => UpdateProcessingUiState();
                TagPartialCheckControl.Unchecked += (_, __) => UpdateProcessingUiState();
            }

            if (WriteZoneBehaviorCheckControl != null)
            {
                WriteZoneBehaviorCheckControl.Checked += (_, __) => UpdateProcessingUiState();
                WriteZoneBehaviorCheckControl.Unchecked += (_, __) => UpdateProcessingUiState();
            }

            if (WriteContainmentPercentCheckControl != null)
            {
                WriteContainmentPercentCheckControl.Checked += (_, __) => UpdateProcessingUiState();
                WriteContainmentPercentCheckControl.Unchecked += (_, __) => UpdateProcessingUiState();
            }

            if (CloseDockPanesCheckControl != null)
            {
                CloseDockPanesCheckControl.Checked += (_, __) => UpdateProcessingUiState();
                CloseDockPanesCheckControl.Unchecked += (_, __) => UpdateProcessingUiState();
            }
        }

        private static void WireHover(FrameworkElement trigger, WpfFlyout flyout)
        {
            if (trigger == null || flyout == null)
            {
                return;
            }

            if (flyout.Content is not FrameworkElement content)
            {
                return;
            }

            _ = new HoverFlyoutController(trigger, content, flyout);
        }

        internal CheckBox TreatPartialCheck => TreatPartialCheckControl;
        internal CheckBox TagPartialCheck => TagPartialCheckControl;
        internal CheckBox WriteZoneBehaviorCheck => WriteZoneBehaviorCheckControl;
        internal CheckBox EnableMultiZoneCheck => EnableMultiZoneCheckControl;
        internal CheckBox ExcludeZonesFromTargetsCheck => ExcludeZonesFromTargetsCheckControl;
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
        internal Slider ZoneBoundsSlider => ZoneBoundsSliderControl;
        internal Slider TargetBoundsSlider => TargetBoundsSliderControl;
        internal FrameworkElement ZoneKDopVariantRow => ZoneKDopVariantRowControl;
        internal ComboBox ZoneKDopVariantCombo => ZoneKDopVariantComboControl;
        internal FrameworkElement TargetKDopVariantRow => TargetKDopVariantRowControl;
        internal ComboBox TargetKDopVariantCombo => TargetKDopVariantComboControl;
        internal FrameworkElement TargetMidpointModeRow => TargetMidpointModeRowControl;
        internal ComboBox TargetMidpointModeCombo => TargetMidpointModeComboControl;
        internal ComboBox ZoneContainmentEngineCombo => ZoneContainmentEngineComboControl;
        internal ComboBox ZoneResolutionStrategyCombo => ZoneResolutionStrategyComboControl;
        internal TextBlock ZoneContainmentHintText => ZoneContainmentHintTextControl;
        internal CheckBox SkipUnchangedWritebackCheck => SkipUnchangedWritebackCheckControl;
        internal CheckBox PackWritebackCheck => PackWritebackCheckControl;
        internal CheckBox ShowInternalWritebackCheck => ShowInternalWritebackCheckControl;
        internal CheckBox WriteContainmentPercentCheck => WriteContainmentPercentCheckControl;
        internal ComboBox ContainmentCalculationCombo => ContainmentCalculationComboControl;
        internal FrameworkElement ContainmentCalculationPanel => ContainmentCalculationPanelControl;
        internal CheckBox CloseDockPanesCheck => CloseDockPanesCheckControl;
        internal TextBox DockPaneDelayBox => DockPaneDelayBoxControl;
        internal FrameworkElement DockPaneDelayRow => DockPaneDelayRowControl;
        internal Button VariationCheckButton => VariationCheckButtonControl;
        internal ComboBox GpuRayAccuracyCombo => GpuRayAccuracyComboControl;

        internal void UpdateProcessingUiState()
        {
            var isMeshAccurate = ZoneContainmentEngineComboControl?.SelectedIndex == 1;
            var isMidpoint = TargetBoundsSliderControl != null
                && (int)Math.Round(TargetBoundsSliderControl.Value) == 0;
            var isMultiZone = EnableMultiZoneCheckControl?.IsChecked == true;

            if (PartialSectionControl != null)
            {
                PartialSectionControl.Visibility = isMidpoint ? Visibility.Collapsed : Visibility.Visible;
            }

            if (MidpointNoPartialNoteTextControl != null)
            {
                MidpointNoPartialNoteTextControl.Visibility = isMidpoint ? Visibility.Visible : Visibility.Collapsed;
            }

            if (TargetMidpointModeRowControl != null)
            {
                TargetMidpointModeRowControl.Visibility = isMidpoint ? Visibility.Visible : Visibility.Collapsed;
            }

            if (TreatPartialCheckControl != null && TagPartialCheckControl != null)
            {
                if (isMidpoint)
                {
                    if (TreatPartialCheckControl.IsChecked == true)
                    {
                        TreatPartialCheckControl.IsChecked = false;
                    }

                    if (TagPartialCheckControl.IsChecked == true)
                    {
                        TagPartialCheckControl.IsChecked = false;
                    }

                    TreatPartialCheckControl.IsEnabled = false;
                    TagPartialCheckControl.IsEnabled = false;
                }
                else
                {
                    TreatPartialCheckControl.IsEnabled = true;
                    TagPartialCheckControl.IsEnabled = true;
                }
            }

            if (WriteContainmentPercentCheckControl != null)
            {
                if (isMidpoint)
                {
                    if (WriteContainmentPercentCheckControl.IsChecked == true)
                    {
                        WriteContainmentPercentCheckControl.IsChecked = false;
                    }

                    WriteContainmentPercentCheckControl.IsEnabled = false;
                }
                else
                {
                    WriteContainmentPercentCheckControl.IsEnabled = true;
                }
            }

            if (DockPaneDelayRowControl != null)
            {
                DockPaneDelayRowControl.Visibility = CloseDockPanesCheckControl?.IsChecked == true
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            }

            if (MeshAccurateNoteTextControl != null)
            {
                MeshAccurateNoteTextControl.Visibility = isMeshAccurate ? Visibility.Visible : Visibility.Collapsed;
            }

            if (GpuRayAccuracyRowControl != null)
            {
                GpuRayAccuracyRowControl.Visibility = isMeshAccurate ? Visibility.Visible : Visibility.Collapsed;
            }

            if (ZoneBoundsHeaderTextControl != null)
            {
                ZoneBoundsHeaderTextControl.Text = isMeshAccurate
                    ? "Fallback Bounds"
                    : "Zone Representation (for candidate filtering / fallback)";
            }

            UpdateResolutionUi(isMidpoint, isMultiZone);
            UpdateEffectiveMethodSummary(isMeshAccurate, isMidpoint, isMultiZone);

            var writeBehavior = WriteZoneBehaviorCheckControl?.IsChecked == true;
            var writePercent = WriteContainmentPercentCheckControl?.IsChecked == true;
            if (ZoneBehaviorFieldsPanelControl != null)
            {
                ZoneBehaviorFieldsPanelControl.IsEnabled = writeBehavior;
                ZoneBehaviorFieldsPanelControl.Visibility = writeBehavior ? Visibility.Visible : Visibility.Collapsed;
            }

            if (ContainmentCalculationPanelControl != null)
            {
                var showCalc = writeBehavior || writePercent;
                ContainmentCalculationPanelControl.IsEnabled = showCalc;
                ContainmentCalculationPanelControl.Visibility = showCalc ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void UpdateResolutionUi(bool isMidpoint, bool isMultiZone)
        {
            if (ZoneResolutionStrategyComboControl == null)
            {
                return;
            }

            ZoneResolutionStrategyComboControl.IsEnabled = !isMultiZone;

            if (ZoneResolutionHintTextControl != null)
            {
                if (isMultiZone)
                {
                    ZoneResolutionHintTextControl.Text = "Resolution is only used when assigning a single best zone.";
                    ZoneResolutionHintTextControl.Visibility = Visibility.Visible;
                }
                else
                {
                    ZoneResolutionHintTextControl.Text = string.Empty;
                    ZoneResolutionHintTextControl.Visibility = Visibility.Collapsed;
                }
            }

            if (ZoneResolutionStrategyComboControl.Items.Count > 1
                && ZoneResolutionStrategyComboControl.Items[1] is ComboBoxItem overlapItem)
            {
                overlapItem.IsEnabled = !isMidpoint;
            }

            if (isMidpoint && ZoneResolutionStrategyComboControl.SelectedIndex == 1)
            {
                ZoneResolutionStrategyComboControl.SelectedIndex = 0;
            }
        }

        private void UpdateEffectiveMethodSummary(bool isMeshAccurate, bool isMidpoint, bool isMultiZone)
        {
            if (EffectiveMethodTextControl == null)
            {
                return;
            }

            var engineLabel = GetComboLabel(ZoneContainmentEngineComboControl) ?? "Bounds (fast)";
            var zoneRep = GetZoneRepresentationLabel();
            var targetRep = GetTargetRepresentationLabel();
            var resolutionLabel = GetResolutionLabel(isMultiZone);
            var multiZoneLabel = isMultiZone ? "On" : "Off";
            var partialLabel = GetPartialsLabel(isMidpoint);
            var zoneLabel = isMeshAccurate ? $"Fallback = {zoneRep}" : $"Zone = {zoneRep}";

            EffectiveMethodTextControl.Text =
                $"Effective: Engine = {engineLabel}; {zoneLabel}; Target = {targetRep}; Resolution = {resolutionLabel}; Multi-zone = {multiZoneLabel}; Partials = {partialLabel}.";
        }

        private string GetZoneRepresentationLabel()
        {
            if (ZoneBoundsSliderControl == null)
            {
                return "AABB";
            }

            return (int)Math.Round(ZoneBoundsSliderControl.Value) switch
            {
                1 => "OBB",
                2 => "k-DOP",
                3 => "Hull",
                _ => "AABB"
            };
        }

        private string GetTargetRepresentationLabel()
        {
            if (TargetBoundsSliderControl == null)
            {
                return "AABB";
            }

            return (int)Math.Round(TargetBoundsSliderControl.Value) switch
            {
                0 => "Midpoint",
                1 => "AABB",
                2 => "OBB",
                3 => "k-DOP",
                4 => "Hull",
                _ => "AABB"
            };
        }

        private string GetResolutionLabel(bool isMultiZone)
        {
            if (isMultiZone)
            {
                return "n/a (multi-zone)";
            }

            var label = GetComboLabel(ZoneResolutionStrategyComboControl);
            if (string.IsNullOrWhiteSpace(label))
            {
                return "Most specific";
            }

            if (label.IndexOf("Most", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Most specific";
            }

            if (label.IndexOf("Largest", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Largest overlap";
            }

            if (label.IndexOf("First", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "First match";
            }

            return label;
        }

        private string GetPartialsLabel(bool isMidpoint)
        {
            if (isMidpoint)
            {
                return "Off (Midpoint)";
            }

            if (TreatPartialCheckControl?.IsChecked == true)
            {
                return "Treated as contained";
            }

            if (TagPartialCheckControl?.IsChecked == true)
            {
                return "Tagged separately";
            }

            return "Off";
        }

        private static string GetComboLabel(ComboBox comboBox)
        {
            if (comboBox?.SelectedItem is ComboBoxItem item && item.Content != null)
            {
                return item.Content.ToString();
            }

            return comboBox?.Text;
        }

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
