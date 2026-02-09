using System;
using System.Diagnostics;
using System.IO;
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
            InitStorageSettings();
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

        private void InitStorageSettings()
        {
            try
            {
                DataCachePathBox.Text = MicroEngStorageSettings.DataStorageDirectory;
                UpdateStorageStatus("Ready.");
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"Settings: init storage failed: {ex}");
                UpdateStorageStatus("Storage settings failed to load.");
            }
        }

        private void BrowseDataCache_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "Select Data Scraper cache folder";
                    dialog.SelectedPath = DataCachePathBox.Text?.Trim();
                    dialog.ShowNewFolderButton = true;

                    if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    {
                        return;
                    }

                    DataCachePathBox.Text = dialog.SelectedPath ?? string.Empty;
                }
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"Settings: browse data cache failed: {ex}");
                UpdateStorageStatus("Browse failed.");
            }
        }

        private void ApplyDataCache_Click(object sender, RoutedEventArgs e)
        {
            ApplyDataCacheDirectory(DataCachePathBox.Text);
        }

        private void ResetDataCache_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MicroEngStorageSettings.ResetDataStorageDirectory(out var resolved);
                DataScraperCache.ReloadFromStorage();
                DataCachePathBox.Text = resolved;
                UpdateStorageStatus("Location reset to default.");
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"Settings: reset data cache failed: {ex}");
                UpdateStorageStatus("Reset failed.");
            }
        }

        private void OpenDataCacheFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var path = ResolveDirectoryPath(DataCachePathBox.Text);
                Directory.CreateDirectory(path);
                Process.Start("explorer.exe", path);
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"Settings: open data cache folder failed: {ex}");
                UpdateStorageStatus("Open folder failed.");
            }
        }

        private void DeleteDataCache_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var result = System.Windows.MessageBox.Show(
                    "Delete all Data Scraper cached sessions from disk and memory?",
                    "MicroEng",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result != MessageBoxResult.Yes)
                {
                    return;
                }

                DataScraperCache.ClearSessions(deletePersistedFile: true);
                UpdateStorageStatus("Data Scraper cache deleted.");
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"Settings: delete data cache failed: {ex}");
                UpdateStorageStatus("Delete cache failed.");
            }
        }

        private void ApplyDataCacheDirectory(string rawPath)
        {
            try
            {
                var resolved = ResolveDirectoryPath(rawPath);
                Directory.CreateDirectory(resolved);

                var changed = MicroEngStorageSettings.SetDataStorageDirectory(resolved, out var finalPath);
                DataScraperCache.ReloadFromStorage();
                DataCachePathBox.Text = finalPath;
                UpdateStorageStatus(changed ? "Location updated." : "Location unchanged.");
            }
            catch (Exception ex)
            {
                MicroEngActions.Log($"Settings: apply data cache failed: {ex}");
                UpdateStorageStatus("Invalid location.");
            }
        }

        private static string ResolveDirectoryPath(string rawPath)
        {
            var path = string.IsNullOrWhiteSpace(rawPath)
                ? MicroEngStorageSettings.DefaultDataStorageDirectory
                : Environment.ExpandEnvironmentVariables(rawPath.Trim());

            if (!Path.IsPathRooted(path))
            {
                path = Path.GetFullPath(Path.Combine(MicroEngStorageSettings.DefaultDataStorageDirectory, path));
            }
            else
            {
                path = Path.GetFullPath(path);
            }

            return path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }

        private void UpdateStorageStatus(string prefix)
        {
            try
            {
                var filePath = DataScraperCache.GetStoreFilePath();
                var count = DataScraperCache.AllSessions.Count;
                StorageStatusText.Text = $"{prefix} Sessions: {count}. File: {filePath}";
            }
            catch
            {
                StorageStatusText.Text = prefix;
            }
        }
    }
}
