using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using Wpf.Ui;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;
using Wpf.Ui.Extensions;
using Wpf.Ui.Markup;
using System.Windows.Media;

namespace MicroEng.Navisworks
{
    public enum MicroEngThemeMode
    {
        Dark = 0,
        Light = 1
    }

    public enum MicroEngAccentMode
    {
        System = 0,
        Custom = 1,
        BlackWhite = 2
    }

    public static class MicroEngWpfUiTheme
    {
        private static readonly Color DefaultDataGridGridLineColor = Color.FromRgb(0xC0, 0xC0, 0xC0);
        private static readonly object Sync = new object();
        private static readonly List<WeakReference<FrameworkElement>> Roots = new List<WeakReference<FrameworkElement>>();

        private sealed class AppliedState
        {
            public MicroEngThemeMode Theme { get; }
            public MicroEngAccentMode AccentMode { get; }
            public Color? CustomAccentColor { get; }
            public Color DataGridGridLineColor { get; }

            public AppliedState(MicroEngThemeMode theme, MicroEngAccentMode accentMode, Color? customAccentColor, Color dataGridGridLineColor)
            {
                Theme = theme;
                AccentMode = accentMode;
                CustomAccentColor = customAccentColor;
                DataGridGridLineColor = dataGridGridLineColor;
            }
        }

        private static bool _isBroadcasting;
        private static MicroEngThemeMode _currentTheme = MicroEngThemeMode.Dark;
        private static MicroEngAccentMode _accentMode = MicroEngAccentMode.BlackWhite;
        private static Color? _customAccentColor;
        private static Color _dataGridGridLineColor = DefaultDataGridGridLineColor;

        private static readonly DependencyProperty AppliedThemeProperty =
            DependencyProperty.RegisterAttached(
                "AppliedTheme",
                typeof(object),
                typeof(MicroEngWpfUiTheme),
                new PropertyMetadata(null));

        public static event Action<MicroEngThemeMode> ThemeChanged;

        public static MicroEngThemeMode CurrentTheme => _currentTheme;
        public static MicroEngAccentMode CurrentAccentMode => _accentMode;
        public static Color? CustomAccentColor => _customAccentColor;
        public static Color DataGridGridLineColor => _dataGridGridLineColor;

        public static void SetTheme(MicroEngThemeMode theme)
        {
            bool changed;
            lock (Sync)
            {
                changed = _currentTheme != theme;
                _currentTheme = theme;
            }

            if (!changed)
            {
                return;
            }

            try
            {
                MicroEngActions.Log($"Theme: mode={theme}");
            }
            catch
            {
                // ignore
            }

            ThemeChanged?.Invoke(theme);
            BroadcastThemeChange();
        }

        public static void ToggleTheme()
        {
            SetTheme(_currentTheme == MicroEngThemeMode.Dark ? MicroEngThemeMode.Light : MicroEngThemeMode.Dark);
        }

        public static void SetAccentMode(MicroEngAccentMode mode, Color? customColor = null)
        {
            bool changed;
            lock (Sync)
            {
                changed = _accentMode != mode || (_customAccentColor != customColor);
                _accentMode = mode;
                _customAccentColor = mode == MicroEngAccentMode.Custom ? customColor : null;
            }

            if (!changed)
            {
                return;
            }

            try
            {
                var colorText = _customAccentColor.HasValue ? _customAccentColor.Value.ToString() : "(none)";
                MicroEngActions.Log($"Theme: accentMode={_accentMode}, custom={colorText}");
            }
            catch
            {
                // ignore
            }

            BroadcastThemeChange();
        }

        public static void SetCustomAccentColor(Color customColor)
        {
            SetAccentMode(MicroEngAccentMode.Custom, customColor);
        }

        public static void SetDataGridGridLineColor(Color color)
        {
            bool changed;
            lock (Sync)
            {
                changed = _dataGridGridLineColor != color;
                _dataGridGridLineColor = color;
            }

            if (!changed)
            {
                return;
            }

            try
            {
                MicroEngActions.Log($"Theme: datagridGridLineColor={color}");
            }
            catch
            {
                // ignore
            }

            BroadcastThemeChange();
        }

        public static void ResetDataGridGridLineColor()
        {
            SetDataGridGridLineColor(DefaultDataGridGridLineColor);
        }

        public static void ApplyTo(FrameworkElement root, bool forceDark = false)
        {
            if (root == null)
            {
                return;
            }

            RegisterRoot(root);

            var targetTheme = forceDark ? MicroEngThemeMode.Dark : _currentTheme;
            ApplyToInternal(root, targetTheme);
        }

        private static void BroadcastThemeChange()
        {
            if (_isBroadcasting)
            {
                return;
            }

            try
            {
                _isBroadcasting = true;

                var aliveRoots = new List<FrameworkElement>();
                lock (Sync)
                {
                    for (var i = Roots.Count - 1; i >= 0; i--)
                    {
                        FrameworkElement target;
                        if (!Roots[i].TryGetTarget(out target) || target == null)
                        {
                            Roots.RemoveAt(i);
                            continue;
                        }

                        aliveRoots.Add(target);
                    }
                }

                try
                {
                    MicroEngActions.Log($"Theme: broadcasting to {aliveRoots.Count} root(s)");
                }
                catch
                {
                    // ignore
                }

                foreach (var root in aliveRoots)
                {
                    if (root.Dispatcher == null)
                    {
                        continue;
                    }

                    void Apply()
                    {
                        try
                        {
                            ApplyToInternal(root, _currentTheme);
                        }
                        catch (Exception ex)
                        {
                            // Stability-first: never allow theme application to crash Navisworks.
                            try
                            {
                                MicroEngActions.Log($"Theme: apply failed ({root.GetType().Name}): {ex.GetType().Name}: {ex.Message}");
                            }
                            catch
                            {
                                // ignore
                            }
                        }
                    }

                    if (root.Dispatcher.CheckAccess())
                    {
                        Apply();
                    }
                    else
                    {
                        root.Dispatcher.BeginInvoke((Action)Apply, DispatcherPriority.Background);
                    }
                }
            }
            finally
            {
                _isBroadcasting = false;
            }
        }

        private static void ApplyToInternal(FrameworkElement root, MicroEngThemeMode theme)
        {
            MicroEngAccentMode accentMode;
            Color? customAccent;
            Color dataGridGridLineColor;
            lock (Sync)
            {
                accentMode = _accentMode;
                customAccent = _customAccentColor;
                dataGridGridLineColor = _dataGridGridLineColor;
            }

            var alreadyAppliedObj = root.GetValue(AppliedThemeProperty);
            if (alreadyAppliedObj is AppliedState alreadyApplied
                && alreadyApplied.Theme == theme
                && alreadyApplied.AccentMode == accentMode
                && alreadyApplied.CustomAccentColor == customAccent
                && alreadyApplied.DataGridGridLineColor == dataGridGridLineColor
                && HasAccentResources(root))
            {
                return;
            }

            EnsureWpfUiRootDictionaryMerged(root);

            var appTheme = theme == MicroEngThemeMode.Dark ? ApplicationTheme.Dark : ApplicationTheme.Light;
            ApplyThemeDictionaries(root, appTheme);
            ApplyTextResources(root, appTheme);
            ApplyAccentResources(root, appTheme);
            ApplyDataGridResources(root, dataGridGridLineColor);
            ApplyPrimaryActionButtonOverrides(root, appTheme);

            root.SetValue(AppliedThemeProperty, new AppliedState(theme, accentMode, customAccent, dataGridGridLineColor));
        }

        private static void ApplyDataGridResources(FrameworkElement root, Color gridLineColor)
        {
            try
            {
                var themeDictionary = TryGetThemesDictionary(root) ?? root.Resources;
                // Resource used by styles in MicroEngUiKit.xaml.
                SetColorResource(root, themeDictionary, "MicroEngDataGridGridLineColor", gridLineColor);
                SetBrushResource(root, themeDictionary, "MicroEngDataGridGridLineBrush", gridLineColor, opacity: 0.55);
            }
            catch
            {
                // ignore
            }
        }

        private static void ApplyTextResources(FrameworkElement root, ApplicationTheme theme)
        {
            try
            {
                var themeDictionary = TryGetThemesDictionary(root) ?? root.Resources;

                // Values sourced from Wpf.Ui theme dictionaries (Dark.xaml / Light.xaml).
                Color primary;
                Color secondary;
                Color tertiary;
                Color disabled;
                Color placeholder;
                Color inverse;

                if (theme == ApplicationTheme.Dark)
                {
                    primary = Color.FromArgb(0xFF, 0xFF, 0xFF, 0xFF);
                    secondary = Color.FromArgb(0xC5, 0xFF, 0xFF, 0xFF);
                    tertiary = Color.FromArgb(0x87, 0xFF, 0xFF, 0xFF);
                    disabled = Color.FromArgb(0x5D, 0xFF, 0xFF, 0xFF);
                    placeholder = Color.FromArgb(0x87, 0xFF, 0xFF, 0xFF);
                    inverse = Color.FromArgb(0xE4, 0x00, 0x00, 0x00);
                }
                else
                {
                    primary = Color.FromArgb(0xE4, 0x00, 0x00, 0x00);
                    secondary = Color.FromArgb(0x9E, 0x00, 0x00, 0x00);
                    tertiary = Color.FromArgb(0x72, 0x00, 0x00, 0x00);
                    disabled = Color.FromArgb(0x5C, 0x00, 0x00, 0x00);
                    placeholder = Color.FromArgb(0x9E, 0x00, 0x00, 0x00);
                    inverse = Color.FromArgb(0xFF, 0xFF, 0xFF, 0xFF);
                }

                SetColorResource(root, themeDictionary, "TextFillColorPrimary", primary);
                SetColorResource(root, themeDictionary, "TextFillColorSecondary", secondary);
                SetColorResource(root, themeDictionary, "TextFillColorTertiary", tertiary);
                SetColorResource(root, themeDictionary, "TextFillColorDisabled", disabled);
                SetColorResource(root, themeDictionary, "TextPlaceholderColor", placeholder);
                SetColorResource(root, themeDictionary, "TextFillColorInverse", inverse);

                SetBrushResource(root, themeDictionary, "TextFillColorPrimaryBrush", primary);
                SetBrushResource(root, themeDictionary, "TextFillColorSecondaryBrush", secondary);
                SetBrushResource(root, themeDictionary, "TextFillColorTertiaryBrush", tertiary);
                SetBrushResource(root, themeDictionary, "TextFillColorDisabledBrush", disabled);
                SetBrushResource(root, themeDictionary, "TextPlaceholderColorBrush", placeholder);
                SetBrushResource(root, themeDictionary, "TextFillColorInverseBrush", inverse);
            }
            catch
            {
                // ignore
            }
        }

        private static void RegisterRoot(FrameworkElement root)
        {
            lock (Sync)
            {
                for (var i = Roots.Count - 1; i >= 0; i--)
                {
                    FrameworkElement existing;
                    if (!Roots[i].TryGetTarget(out existing) || existing == null)
                    {
                        Roots.RemoveAt(i);
                        continue;
                    }

                    if (ReferenceEquals(existing, root))
                    {
                        return;
                    }
                }

                Roots.Add(new WeakReference<FrameworkElement>(root));
            }
        }

        private static void EnsureWpfUiRootDictionaryMerged(FrameworkElement root)
        {
            try
            {
                var merged = root.Resources.MergedDictionaries;
                if (merged.Any(d => d.Source != null && d.Source.Equals(MicroEngResourceUris.WpfUiRoot)))
                {
                    return;
                }

                merged.Insert(0, new ResourceDictionary { Source = MicroEngResourceUris.WpfUiRoot });
            }
            catch
            {
                // Ignore: failure to merge should not crash host; XAML should already merge it.
            }
        }

        private static void ApplyThemeDictionaries(FrameworkElement root, ApplicationTheme appTheme)
        {
            try
            {
                var updated = 0;
                foreach (var dict in EnumerateAllDictionaries(root.Resources))
                {
                    if (dict is ThemesDictionary themes)
                    {
                        themes.Theme = appTheme;
                        updated++;
                    }
                }

                if (updated == 0)
                {
                    try
                    {
                        MicroEngActions.Log($"Theme: no ThemesDictionary found to set {appTheme}.");
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    MicroEngActions.Log($"Theme: ApplyThemeDictionaries failed ({ex.GetType().Name}): {ex.Message}");
                }
                catch
                {
                    // ignore
                }
            }
        }

        private static IEnumerable<ResourceDictionary> EnumerateAllDictionaries(ResourceDictionary resourceDictionary)
        {
            var visited = new HashSet<ResourceDictionary>();
            foreach (var dict in EnumerateAllDictionariesCore(resourceDictionary, visited))
            {
                yield return dict;
            }
        }

        private static IEnumerable<ResourceDictionary> EnumerateAllDictionariesCore(
            ResourceDictionary resourceDictionary,
            HashSet<ResourceDictionary> visited)
        {
            if (resourceDictionary == null || visited == null)
            {
                yield break;
            }

            if (!visited.Add(resourceDictionary))
            {
                yield break;
            }

            yield return resourceDictionary;

            foreach (var merged in resourceDictionary.MergedDictionaries)
            {
                foreach (var nested in EnumerateAllDictionariesCore(merged, visited))
                {
                    yield return nested;
                }
            }
        }

        private static bool HasAccentResources(FrameworkElement root)
        {
            try
            {
                return root.Resources.Contains("SystemAccentColor")
                    && root.Resources.Contains("AccentFillColorDefault")
                    && root.Resources.Contains("AccentFillColorDefaultBrush");
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Applies WPF-UI accent resources to a specific root (no global App.xaml / UiApplication.Current).
        /// Mirrors <see cref="Wpf.Ui.Appearance.ApplicationAccentColorManager.Apply"/> behavior, but scoped per-root.
        /// </summary>
        private static void ApplyAccentResources(FrameworkElement root, ApplicationTheme applicationTheme)
        {
            try
            {
                // Use the official WPF-UI accent derivation logic (Gallery-like),
                // but write results into the current root's resources (NOT the whole Navisworks process).
                //
                // Important: ApplicationAccentColorManager.Apply writes into UiApplication.Current.Resources,
                // which is safe in plugin-hosted scenarios because it falls back to an internal dictionary
                // when there is no global App.xaml/Application resources for WPF-UI.
                var systemAccent = GetBaseAccentColor(applicationTheme);
                try
                {
                    ApplicationAccentColorManager.Apply(systemAccent, applicationTheme, systemGlassColor: false);

                    ApplyAccentResourcesCore(
                        root,
                        applicationTheme,
                        ApplicationAccentColorManager.SystemAccent,
                        ApplicationAccentColorManager.PrimaryAccent,
                        ApplicationAccentColorManager.SecondaryAccent,
                        ApplicationAccentColorManager.TertiaryAccent);

                    return;
                }
                catch (Exception ex)
                {
                    // Fall back to the documented/default HSV adjustments if the WPF-UI manager fails.
                    try
                    {
                        MicroEngActions.Log($"Theme: accent manager failed ({ex.GetType().Name}): {ex.Message}");
                    }
                    catch
                    {
                        // ignore
                    }
                }

                Color primaryAccent;
                Color secondaryAccent;
                Color tertiaryAccent;

                if (applicationTheme == ApplicationTheme.Dark)
                {
                    primaryAccent = systemAccent.Update(17f, -30f);
                    secondaryAccent = systemAccent.Update(17f, -45f);
                    tertiaryAccent = systemAccent.Update(17f, -65f);
                }
                else
                {
                    primaryAccent = systemAccent.Update(-10f);
                    secondaryAccent = systemAccent.Update(-25f);
                    tertiaryAccent = systemAccent.Update(-40f);
                }

                ApplyAccentResourcesCore(root, applicationTheme, systemAccent, primaryAccent, secondaryAccent, tertiaryAccent);
            }
            catch (Exception ex)
            {
                // Stability-first: never allow accent application to crash Navisworks.
                try
                {
                    MicroEngActions.Log($"Theme: ApplyAccentResources failed ({ex.GetType().Name}): {ex.Message}");
                }
                catch
                {
                    // ignore
                }
            }
        }

        private static Color GetBaseAccentColor(ApplicationTheme applicationTheme)
        {
            lock (Sync)
            {
                if (_accentMode == MicroEngAccentMode.Custom && _customAccentColor.HasValue)
                {
                    return _customAccentColor.Value;
                }

                if (_accentMode == MicroEngAccentMode.BlackWhite)
                {
                    return applicationTheme == ApplicationTheme.Dark
                        ? Color.FromRgb(0xFF, 0xFF, 0xFF)
                        : Color.FromRgb(0x00, 0x00, 0x00);
                }
            }

            return ApplicationAccentColorManager.GetColorizationColor();
        }

        private static void ApplyAccentResourcesCore(
            FrameworkElement root,
            ApplicationTheme applicationTheme,
            Color systemAccent,
            Color primaryAccent,
            Color secondaryAccent,
            Color tertiaryAccent)
        {
            // Mirrors Wpf.Ui.Appearance.ApplicationAccentColorManager.UpdateColorResources(...)
            // but writes into a specific root's resource dictionary (no global App.xaml / Navisworks theming).

            // NOTE:
            // WPF-UI defines many resources inside its own merged dictionaries (ThemesDictionary -> Theme/*.xaml).
            // When those dictionaries use {DynamicResource ...} internally, the lookup often resolves within the
            // ThemesDictionary tree rather than the root FrameworkElement.Resources. To ensure consistent behavior
            // (and reliable hot-updates), write accent keys into BOTH:
            // - the root.Resources (so app code can resolve them), and
            // - the ThemesDictionary (so WPF-UI templates resolve them).
            var themeDictionary = TryGetThemesDictionary(root) ?? root.Resources;

            // Text contrast on accent (matches WPF-UI threshold logic).
            const double backgroundBrightnessThresholdValue = 80d;
            // MicroEng: prefer the *base* accent for control fills (matches user expectation and avoids lighter button fills in dark mode).
            var themeAccent = systemAccent;
            var isDarkTextOnAccent = themeAccent.GetBrightness() > backgroundBrightnessThresholdValue;

            if (isDarkTextOnAccent)
            {
                root.Resources["TextOnAccentFillColorPrimary"] = Color.FromArgb(0xFF, 0x00, 0x00, 0x00);
                root.Resources["TextOnAccentFillColorSecondary"] = Color.FromArgb(0x80, 0x00, 0x00, 0x00);
                root.Resources["TextOnAccentFillColorDisabled"] = Color.FromArgb(0x77, 0x00, 0x00, 0x00);
                root.Resources["TextOnAccentFillColorSelectedText"] = Color.FromArgb(0x00, 0x00, 0x00, 0x00);
                root.Resources["AccentTextFillColorDisabled"] = Color.FromArgb(0x5D, 0x00, 0x00, 0x00);
            }
            else
            {
                root.Resources["TextOnAccentFillColorPrimary"] = Color.FromArgb(0xFF, 0xFF, 0xFF, 0xFF);
                root.Resources["TextOnAccentFillColorSecondary"] = Color.FromArgb(0x80, 0xFF, 0xFF, 0xFF);
                root.Resources["TextOnAccentFillColorDisabled"] = Color.FromArgb(0x87, 0xFF, 0xFF, 0xFF);
                root.Resources["TextOnAccentFillColorSelectedText"] = Color.FromArgb(0xFF, 0xFF, 0xFF, 0xFF);
                root.Resources["AccentTextFillColorDisabled"] = Color.FromArgb(0x5D, 0xFF, 0xFF, 0xFF);
            }

            // In the official WPF-UI Gallery, Dark theme uses "SystemAccentColorSecondary" for many control fills.
            // In Navisworks plugin hosting, we want a single, consistent accent (no unexpected lighter orange).
            // Force all derived accent variants to the base accent; state variations are handled via
            // AccentFillColorSecondary/Tertiary opacity keys below.
            primaryAccent = systemAccent;
            secondaryAccent = systemAccent;
            tertiaryAccent = systemAccent;

            SetColorResource(root, themeDictionary, "SystemAccentColor", systemAccent);
            SetColorResource(root, themeDictionary, "SystemAccentColorPrimary", primaryAccent);
            SetColorResource(root, themeDictionary, "SystemAccentColorSecondary", secondaryAccent);
            SetColorResource(root, themeDictionary, "SystemAccentColorTertiary", tertiaryAccent);

            // Match WPF-UI Accent.xaml keys so control templates update via DynamicResource.
            SetBrushResource(root, themeDictionary, "SystemAccentColorBrush", systemAccent);
            SetBrushResource(root, themeDictionary, "SystemAccentColorPrimaryBrush", primaryAccent);
            SetBrushResource(root, themeDictionary, "SystemAccentColorSecondaryBrush", secondaryAccent);
            SetBrushResource(root, themeDictionary, "SystemAccentColorTertiaryBrush", tertiaryAccent);

            // Back-compat key (older MicroEng XAML).
            SetBrushResource(root, themeDictionary, "SystemAccentBrush", systemAccent);

            SetBrushResource(root, themeDictionary, "SystemFillColorAttentionBrush", secondaryAccent);

            SetBrushResource(root, themeDictionary, "AccentTextFillColorPrimaryBrush", secondaryAccent);
            SetBrushResource(root, themeDictionary, "AccentTextFillColorSecondaryBrush", tertiaryAccent);
            SetBrushResource(root, themeDictionary, "AccentTextFillColorTertiaryBrush", primaryAccent);

            SetBrushResource(root, themeDictionary, "AccentFillColorSelectedTextBackgroundBrush", systemAccent);

            SetColorResource(root, themeDictionary, "AccentFillColorDefault", themeAccent);
            SetBrushResource(root, themeDictionary, "AccentFillColorDefaultBrush", themeAccent);

            SetColorResource(root, themeDictionary, "AccentFillColorSecondary", Color.FromArgb(229, themeAccent.R, themeAccent.G, themeAccent.B)); // 229 = 0.9 * 255
            SetBrushResource(root, themeDictionary, "AccentFillColorSecondaryBrush", themeAccent, 0.9);

            SetColorResource(root, themeDictionary, "AccentFillColorTertiary", Color.FromArgb(204, themeAccent.R, themeAccent.G, themeAccent.B)); // 204 = 0.8 * 255
            SetBrushResource(root, themeDictionary, "AccentFillColorTertiaryBrush", themeAccent, 0.8);

            // Important: many WPF-UI theme dictionaries create these brushes using StaticResource -> Color,
            // so changing only the Color key doesn't update already-created SolidColorBrush instances.
            // Override brush keys too, so custom accent changes apply immediately to live UI.
            var textPrimary = (Color)root.Resources["TextOnAccentFillColorPrimary"];
            var textSecondary = (Color)root.Resources["TextOnAccentFillColorSecondary"];
            var textDisabled = (Color)root.Resources["TextOnAccentFillColorDisabled"];
            var textSelected = (Color)root.Resources["TextOnAccentFillColorSelectedText"];
            var accentTextDisabled = (Color)root.Resources["AccentTextFillColorDisabled"];

            SetBrushResource(root, themeDictionary, "TextOnAccentFillColorPrimaryBrush", textPrimary);
            SetBrushResource(root, themeDictionary, "TextOnAccentFillColorSecondaryBrush", textSecondary);
            SetBrushResource(root, themeDictionary, "TextOnAccentFillColorDisabledBrush", textDisabled);
            SetBrushResource(root, themeDictionary, "TextOnAccentFillColorSelectedTextBrush", textSelected);
            SetBrushResource(root, themeDictionary, "AccentTextFillColorDisabledBrush", accentTextDisabled);

            // Primary action buttons should be high-contrast and consistent across tools:
            // Dark theme  -> white button with black text.
            // Light theme -> black button with white text.
            var actionButtonBackground = applicationTheme == ApplicationTheme.Dark
                ? Color.FromRgb(0xFF, 0xFF, 0xFF)
                : Color.FromRgb(0x00, 0x00, 0x00);
            var actionButtonBackgroundPointerOver = applicationTheme == ApplicationTheme.Dark
                ? Color.FromRgb(0xE6, 0xE6, 0xE6)
                : Color.FromRgb(0x1A, 0x1A, 0x1A);
            var actionButtonBackgroundPressed = applicationTheme == ApplicationTheme.Dark
                ? Color.FromRgb(0xCC, 0xCC, 0xCC)
                : Color.FromRgb(0x33, 0x33, 0x33);

            var actionButtonForeground = applicationTheme == ApplicationTheme.Dark
                ? Color.FromRgb(0x00, 0x00, 0x00)
                : Color.FromRgb(0xFF, 0xFF, 0xFF);
            var actionButtonForegroundPressed = applicationTheme == ApplicationTheme.Dark
                ? Color.FromRgb(0x20, 0x20, 0x20)
                : Color.FromRgb(0xE6, 0xE6, 0xE6);

            // WPF-UI Buttons (Appearance=Primary) consume these AccentButton* brushes.
            SetBrushResource(root, themeDictionary, "AccentButtonBackground", actionButtonBackground);
            SetBrushResource(root, themeDictionary, "AccentButtonBackgroundPointerOver", actionButtonBackgroundPointerOver);
            SetBrushResource(root, themeDictionary, "AccentButtonBackgroundPressed", actionButtonBackgroundPressed);

            SetBrushResource(root, themeDictionary, "AccentButtonForeground", actionButtonForeground);
            SetBrushResource(root, themeDictionary, "AccentButtonForegroundPointerOver", actionButtonForeground);
            SetBrushResource(root, themeDictionary, "AccentButtonForegroundPressed", actionButtonForegroundPressed);

            // TabControl/TabItem styling (WPF-UI templates use these keys).
            // Make the selected tab reflect the current accent (e.g. white in BlackWhite dark mode),
            // while keeping the rest of the TabControl template intact for stability.
            SetBrushResource(root, themeDictionary, "TabViewItemHeaderBackgroundSelected", themeAccent);
            SetBrushResource(root, themeDictionary, "TabViewItemForegroundSelected", textPrimary);
            SetBrushResource(root, themeDictionary, "TabViewSelectedItemBorderBrush", themeAccent, 0.9);
        }

        private static void ApplyPrimaryActionButtonOverrides(FrameworkElement root, ApplicationTheme applicationTheme)
        {
            try
            {
                var backgroundColor = applicationTheme == ApplicationTheme.Dark
                    ? Color.FromRgb(0xFF, 0xFF, 0xFF)
                    : Color.FromRgb(0x00, 0x00, 0x00);
                var foregroundColor = applicationTheme == ApplicationTheme.Dark
                    ? Color.FromRgb(0x00, 0x00, 0x00)
                    : Color.FromRgb(0xFF, 0xFF, 0xFF);

                var backgroundBrush = new SolidColorBrush(backgroundColor);
                backgroundBrush.Freeze();
                var foregroundBrush = new SolidColorBrush(foregroundColor);
                foregroundBrush.Freeze();

                foreach (var button in FindVisualChildren<Button>(root))
                {
                    if (button.Appearance != ControlAppearance.Primary)
                    {
                        continue;
                    }

                    button.Background = backgroundBrush;
                    button.BorderBrush = backgroundBrush;
                    button.Foreground = foregroundBrush;
                }
            }
            catch
            {
                // ignore
            }
        }

        private static IEnumerable<T> FindVisualChildren<T>(DependencyObject parent)
            where T : DependencyObject
        {
            if (parent == null)
            {
                yield break;
            }

            var childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (var i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T match)
                {
                    yield return match;
                }

                foreach (var nested in FindVisualChildren<T>(child))
                {
                    yield return nested;
                }
            }
        }

        private static ResourceDictionary TryGetThemesDictionary(FrameworkElement root)
        {
            try
            {
                foreach (var dict in EnumerateAllDictionaries(root.Resources))
                {
                    if (dict is ThemesDictionary)
                    {
                        return dict;
                    }
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private static void SetColorResource(FrameworkElement root, ResourceDictionary themeDictionary, string key, Color color)
        {
            root.Resources[key] = color;

            // Ensure the key is also available in the ThemesDictionary root (WPF-UI templates often resolve from there).
            try
            {
                if (!ReferenceEquals(themeDictionary, root.Resources))
                {
                    themeDictionary[key] = color;
                }
            }
            catch
            {
                // ignore
            }

            // Hot-update: if the key exists deeper in merged dictionaries (Accent.xaml / Theme/*.xaml),
            // update it there too so any templates using StaticResource-resolved brushes can be refreshed
            // by re-binding to the updated resource instances.
            try
            {
                foreach (var dict in EnumerateAllDictionaries(root.Resources))
                {
                    if (ReferenceEquals(dict, root.Resources) || ReferenceEquals(dict, themeDictionary))
                    {
                        continue;
                    }

                    if (!dict.Contains(key))
                    {
                        continue;
                    }

                    dict[key] = color;
                }
            }
            catch
            {
                // ignore
            }
        }

        private static void SetBrushResource(FrameworkElement root, ResourceDictionary themeDictionary, string key, Color color, double opacity = 1.0)
        {
            var targetColor = opacity >= 0.999 ? color : Color.FromArgb((byte)Math.Round(opacity * 255d), color.R, color.G, color.B);

            try
            {
                // Prefer updating the brush in the theme dictionary (WPF-UI templates tend to resolve from there).
                if (!ReferenceEquals(themeDictionary, root.Resources))
                {
                    if (themeDictionary[key] is SolidColorBrush themeBrush && !themeBrush.IsFrozen)
                    {
                        themeBrush.Color = targetColor;
                        root.Resources[key] = themeBrush;
                        // continue to update other dictionaries too
                    }
                }

                var existing = root.TryFindResource(key) as SolidColorBrush;
                if (existing != null && !existing.IsFrozen)
                {
                    existing.Color = targetColor;
                    // continue to update other dictionaries too
                }
            }
            catch
            {
                // ignore
            }

            SolidColorBrush brushToSet = null;
            try
            {
                // Prefer reusing an existing brush instance if present (helps DynamicResource & bindings).
                var existingBrush = root.TryFindResource(key) as SolidColorBrush;
                if (existingBrush != null && !existingBrush.IsFrozen)
                {
                    existingBrush.Color = targetColor;
                    brushToSet = existingBrush;
                }
            }
            catch
            {
                // ignore
            }

            if (brushToSet == null)
            {
                brushToSet = new SolidColorBrush(targetColor);
            }

            root.Resources[key] = brushToSet;

            try
            {
                if (!ReferenceEquals(themeDictionary, root.Resources))
                {
                    themeDictionary[key] = brushToSet;
                }
            }
            catch
            {
                // ignore
            }

            // Hot-update any existing brush resources deeper in merged dictionaries (Accent.xaml / Theme/*.xaml).
            try
            {
                foreach (var dict in EnumerateAllDictionaries(root.Resources))
                {
                    if (ReferenceEquals(dict, root.Resources) || ReferenceEquals(dict, themeDictionary))
                    {
                        continue;
                    }

                    if (!dict.Contains(key))
                    {
                        continue;
                    }

                    if (dict[key] is SolidColorBrush b && !b.IsFrozen)
                    {
                        b.Color = targetColor;
                    }
                    else
                    {
                        dict[key] = new SolidColorBrush(targetColor);
                    }
                }
            }
            catch
            {
                // ignore
            }
        }
    }
}
