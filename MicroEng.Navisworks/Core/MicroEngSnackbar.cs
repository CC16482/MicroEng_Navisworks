using System;
using System.Windows.Media;
using WpfUiControls = Wpf.Ui.Controls;

namespace MicroEng.Navisworks
{
    internal static class MicroEngSnackbar
    {
        private static readonly Brush ForegroundBrush = Brushes.Black;
        private const int IconFontSize = 25;
        private const int TimeoutSeconds = 4;

        public static void Show(
            WpfUiControls.SnackbarPresenter presenter,
            string title,
            string message,
            WpfUiControls.ControlAppearance appearance,
            WpfUiControls.SymbolRegular icon = WpfUiControls.SymbolRegular.PresenceAvailable24)
        {
            if (presenter == null)
            {
                return;
            }

            var snackbar = new WpfUiControls.Snackbar(presenter)
            {
                Title = title,
                Content = message,
                Appearance = appearance,
                Icon = new WpfUiControls.SymbolIcon(icon)
                {
                    Filled = true,
                    FontSize = IconFontSize
                },
                Foreground = ForegroundBrush,
                ContentForeground = ForegroundBrush,
                Timeout = TimeSpan.FromSeconds(TimeoutSeconds),
                IsCloseButtonEnabled = false
            };

            snackbar.Show();
        }
    }
}
