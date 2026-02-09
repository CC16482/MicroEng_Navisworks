using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using WinFormsScreen = System.Windows.Forms.Screen;

namespace MicroEng.Navisworks
{
    internal static class MicroEngWindowPositioning
    {
        /// <summary>
        /// Keeps MicroEng popup windows on top of Navisworks and centers them on the same monitor as Navisworks.
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
            TryAttachOwner(window);

            void ApplyPosition()
            {
                try
                {
                    window.WindowStartupLocation = WindowStartupLocation.Manual;

                    var workArea = GetNavisworksWorkArea(window);
                    var width = window.ActualWidth > 0 ? window.ActualWidth : window.Width;
                    var height = window.ActualHeight > 0 ? window.ActualHeight : window.Height;

                    if (double.IsNaN(width) || width <= 0)
                    {
                        width = Math.Min(900, workArea.Width);
                    }

                    if (double.IsNaN(height) || height <= 0)
                    {
                        height = Math.Min(700, workArea.Height);
                    }

                    window.Left = workArea.Left + Math.Max(0, (workArea.Width - width) / 2);
                    window.Top = workArea.Top + Math.Max(0, (workArea.Height - height) / 2);
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

        private static void TryAttachOwner(Window window)
        {
            try
            {
                var navisworksHandle = TryGetNavisworksMainWindowHandle();
                if (navisworksHandle == IntPtr.Zero)
                {
                    return;
                }

                var helper = new WindowInteropHelper(window);
                if (helper.Owner == IntPtr.Zero)
                {
                    helper.Owner = navisworksHandle;
                }
            }
            catch
            {
                // best-effort only
            }
        }

        private static Rect GetNavisworksWorkArea(Window window)
        {
            try
            {
                var navisworksHandle = TryGetNavisworksMainWindowHandle();
                if (navisworksHandle == IntPtr.Zero)
                {
                    return SystemParameters.WorkArea;
                }

                var screen = WinFormsScreen.FromHandle(navisworksHandle) ?? WinFormsScreen.PrimaryScreen;
                var workingArea = screen.WorkingArea;
                return DevicePixelsToDip(window, workingArea);
            }
            catch
            {
                return SystemParameters.WorkArea;
            }
        }

        private static Rect DevicePixelsToDip(Window window, System.Drawing.Rectangle rectPixels)
        {
            try
            {
                var source = PresentationSource.FromVisual(window);
                var compositionTarget = source?.CompositionTarget;
                if (compositionTarget != null)
                {
                    var fromDevice = compositionTarget.TransformFromDevice;
                    var topLeft = fromDevice.Transform(new System.Windows.Point(rectPixels.Left, rectPixels.Top));
                    var bottomRight = fromDevice.Transform(new System.Windows.Point(rectPixels.Right, rectPixels.Bottom));
                    return new Rect(topLeft, bottomRight);
                }
            }
            catch
            {
                // best-effort only
            }

            return new Rect(rectPixels.Left, rectPixels.Top, rectPixels.Width, rectPixels.Height);
        }

        private static IntPtr TryGetNavisworksMainWindowHandle()
        {
            try
            {
                using (var process = Process.GetCurrentProcess())
                {
                    if (process.MainWindowHandle != IntPtr.Zero)
                    {
                        return process.MainWindowHandle;
                    }
                }
            }
            catch
            {
                // best-effort only
            }

            return IntPtr.Zero;
        }
    }
}

