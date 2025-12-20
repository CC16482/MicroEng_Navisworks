using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;

namespace MicroEng.Navisworks
{
    public partial class MicroEngSettingsWindow : Window
    {
        private bool _isInitializing;
        // MicroEng brand accent (matches MicroEng theme guide / screenshots).
        private Color _currentCustomAccent = Color.FromRgb(0xF8, 0x74, 0x1D);
        private Color _currentGridLineColor = MicroEngWpfUiTheme.DataGridGridLineColor;

        public MicroEngSettingsWindow()
        {
            InitializeComponent();
            MicroEngWpfUiTheme.ApplyTo(this);
            MicroEngWindowPositioning.ApplyTopMostTopCenter(this);
            InitFromCurrentTheme();
        }

        private void InitFromCurrentTheme()
        {
            _isInitializing = true;
            try
            {
                var mode = MicroEngWpfUiTheme.CurrentAccentMode;
                var custom = MicroEngWpfUiTheme.CustomAccentColor;
                if (custom.HasValue)
                {
                    _currentCustomAccent = custom.Value;
                }

                UseSystemAccentRadio.IsChecked = mode == MicroEngAccentMode.System;
                UseBlackWhiteAccentRadio.IsChecked = mode == MicroEngAccentMode.BlackWhite;
                UseCustomAccentRadio.IsChecked = mode == MicroEngAccentMode.Custom;
                PickAccentButton.IsEnabled = mode == MicroEngAccentMode.Custom;

                UpdateAccentPreview();
                _currentGridLineColor = MicroEngWpfUiTheme.DataGridGridLineColor;
                UpdateGridLinePreview();
            }
            finally
            {
                _isInitializing = false;
            }
        }

        private void AccentModeChanged(object sender, RoutedEventArgs e)
        {
            if (_isInitializing)
            {
                return;
            }

            var useCustom = UseCustomAccentRadio.IsChecked == true;
            var useBlackWhite = UseBlackWhiteAccentRadio.IsChecked == true;
            PickAccentButton.IsEnabled = useCustom;

            if (useCustom)
            {
                MicroEngWpfUiTheme.SetCustomAccentColor(_currentCustomAccent);
            }
            else if (useBlackWhite)
            {
                MicroEngWpfUiTheme.SetAccentMode(MicroEngAccentMode.BlackWhite);
            }
            else
            {
                MicroEngWpfUiTheme.SetAccentMode(MicroEngAccentMode.System);
            }

            UpdateAccentPreview();
        }

        private void PickAccent_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var dialog = new ColorDialog())
                {
                    dialog.FullOpen = true;
                    dialog.Color = System.Drawing.Color.FromArgb(_currentCustomAccent.A, _currentCustomAccent.R, _currentCustomAccent.G, _currentCustomAccent.B);

                    if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    {
                        return;
                    }

                    _currentCustomAccent = Color.FromArgb(dialog.Color.A, dialog.Color.R, dialog.Color.G, dialog.Color.B);
                }

                MicroEngWpfUiTheme.SetCustomAccentColor(_currentCustomAccent);
                UpdateAccentPreview();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"Settings: pick accent failed: {ex}");
            }
        }

        private void PickGridLine_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var dialog = new ColorDialog())
                {
                    dialog.FullOpen = true;
                    dialog.Color = System.Drawing.Color.FromArgb(_currentGridLineColor.A, _currentGridLineColor.R, _currentGridLineColor.G, _currentGridLineColor.B);

                    if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    {
                        return;
                    }

                    _currentGridLineColor = Color.FromArgb(dialog.Color.A, dialog.Color.R, dialog.Color.G, dialog.Color.B);
                }

                MicroEngWpfUiTheme.SetDataGridGridLineColor(_currentGridLineColor);
                UpdateGridLinePreview();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"Settings: pick grid line color failed: {ex}");
            }
        }

        private void ResetGridLine_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MicroEngWpfUiTheme.ResetDataGridGridLineColor();
                _currentGridLineColor = MicroEngWpfUiTheme.DataGridGridLineColor;
                UpdateGridLinePreview();
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"Settings: reset grid line color failed: {ex}");
            }
        }

        private void UpdateAccentPreview()
        {
            try
            {
                Color color;

                if (UseCustomAccentRadio.IsChecked == true)
                {
                    color = _currentCustomAccent;
                }
                else if (UseBlackWhiteAccentRadio.IsChecked == true)
                {
                    color = MicroEngWpfUiTheme.CurrentTheme == MicroEngThemeMode.Dark
                        ? Color.FromRgb(0xFF, 0xFF, 0xFF)
                        : Color.FromRgb(0x00, 0x00, 0x00);
                }
                else
                {
                    color = Wpf.Ui.Appearance.ApplicationAccentColorManager.GetColorizationColor();
                }

                AccentPreview.Background = new SolidColorBrush(color);
            }
            catch
            {
                // ignore
            }
        }

        private void UpdateGridLinePreview()
        {
            try
            {
                GridLinePreview.Background = new SolidColorBrush(_currentGridLineColor);
            }
            catch
            {
                // ignore
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
