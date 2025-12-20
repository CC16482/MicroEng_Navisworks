using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Threading;

namespace MicroEng.Navisworks
{
    internal static class MicroEngWindowPositioning
    {
        /// <summary>
        /// Keeps MicroEng popup windows on top of Navisworks and positions them near the top of the screen.
        /// Dock panes are not affected.
        /// </summary>
        public static void ApplyTopMostTopCenter(Window window)
        {
            if (window == null)
            {
                return;
            }

            if (DesignerProperties.GetIsInDesignMode(window))
            {
                return;
            }

            window.Topmost = true;

            void ApplyPosition()
            {
                try
                {
                    window.WindowStartupLocation = WindowStartupLocation.Manual;

                    var workArea = SystemParameters.WorkArea;
                    var width = window.ActualWidth > 0 ? window.ActualWidth : window.Width;

                    if (double.IsNaN(width) || width <= 0)
                    {
                        width = Math.Min(900, workArea.Width);
                    }

                    window.Left = workArea.Left + Math.Max(0, (workArea.Width - width) / 2);
                    window.Top = workArea.Top + 24;
                }
                catch
                {
                    // best-effort only
                }
            }

            if (window.IsLoaded)
            {
                window.Dispatcher.BeginInvoke((Action)ApplyPosition, DispatcherPriority.Loaded);
            }
            else
            {
                window.Loaded += (_, __) =>
                    window.Dispatcher.BeginInvoke((Action)ApplyPosition, DispatcherPriority.Loaded);
            }
        }
    }
}

