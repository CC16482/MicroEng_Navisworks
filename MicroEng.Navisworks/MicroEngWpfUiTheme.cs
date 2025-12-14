using System;
using System.Windows;
using System.Diagnostics;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

namespace MicroEng.Navisworks
{
    internal static class MicroEngWpfUiTheme
    {
        private static readonly DependencyProperty AppliedProperty =
            DependencyProperty.RegisterAttached("AppliedFlag", typeof(bool), typeof(MicroEngWpfUiTheme), new PropertyMetadata(false));

        public static void ApplyTo(FrameworkElement root, bool forceDark = false)
        {
            if (root == null) return;
            var already = (bool)root.GetValue(AppliedProperty);
            if (already) return;

            try
            {
                var system = ApplicationThemeManager.GetSystemTheme();
                var theme = forceDark || system == SystemTheme.Dark ? ApplicationTheme.Dark : ApplicationTheme.Light;
                ApplicationThemeManager.Apply(theme, WindowBackdropType.None);
                root.SetValue(AppliedProperty, true);
            }
            catch (System.IO.FileNotFoundException ex)
            {
                // Log and continue without theming if a dependency cannot be loaded.
                MicroEngActions.Log($"WPF-UI theme load failed: {ex.FileName ?? ex.Message}");
                Trace.WriteLine(ex.ToString());
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"WPF-UI theme apply failed: {ex.Message}");
                Trace.WriteLine(ex.ToString());
            }

        }
    }
}
