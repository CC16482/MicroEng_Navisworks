using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using Forms = System.Windows.Forms;
using WpfFlyout = Wpf.Ui.Controls.Flyout;

namespace MicroEng.Navisworks
{
    internal sealed class HoverFlyoutController
    {
        private readonly FrameworkElement _trigger;
        private readonly FrameworkElement _flyoutContent;
        private readonly WpfFlyout _flyout;
        private readonly DispatcherTimer _pollTimer;
        private int _missedTicks;
        private DateTime _graceUntil;
        public HoverFlyoutController(FrameworkElement trigger, FrameworkElement flyoutContent, WpfFlyout flyout)
        {
            _trigger = trigger ?? throw new ArgumentNullException(nameof(trigger));
            _flyoutContent = flyoutContent ?? throw new ArgumentNullException(nameof(flyoutContent));
            _flyout = flyout ?? throw new ArgumentNullException(nameof(flyout));

            _trigger.MouseEnter += OnTriggerEnter;
            _trigger.MouseLeave += OnTriggerLeave;
            _flyoutContent.MouseEnter += OnFlyoutEnter;
            _flyoutContent.MouseLeave += OnFlyoutLeave;

            _pollTimer = new DispatcherTimer(DispatcherPriority.Normal, _trigger.Dispatcher)
            {
                Interval = TimeSpan.FromMilliseconds(120)
            };
            _pollTimer.Tick += OnPollTick;
        }

        private void OnTriggerEnter(object sender, MouseEventArgs e)
        {
            _pollTimer.Stop();
            _missedTicks = 0;
            _graceUntil = DateTime.UtcNow.AddMilliseconds(250);
            _flyout.Show();
            _pollTimer.Start();
        }

        private void OnTriggerLeave(object sender, MouseEventArgs e)
        {
            ScheduleClose();
        }

        private void OnFlyoutEnter(object sender, MouseEventArgs e)
        {
            _pollTimer.Stop();
            _missedTicks = 0;
            _pollTimer.Start();
        }

        private void OnFlyoutLeave(object sender, MouseEventArgs e)
        {
            ScheduleClose();
        }

        private void ScheduleClose()
        {
            _pollTimer.Stop();
            _pollTimer.Start();
        }

        private void OnPollTick(object sender, EventArgs e)
        {
            if (DateTime.UtcNow < _graceUntil)
            {
                return;
            }

            if (IsPointerOver(_trigger) || IsPointerOver(_flyoutContent))
            {
                _missedTicks = 0;
                return;
            }

            _missedTicks++;
            if (_missedTicks < 2)
            {
                return;
            }

            _pollTimer.Stop();
            _flyout.Hide();
        }

        private static bool IsPointerOver(FrameworkElement element)
        {
            if (element == null || !element.IsVisible || element.ActualWidth <= 0 || element.ActualHeight <= 0)
            {
                return false;
            }

            if (PresentationSource.FromVisual(element) == null)
            {
                return false;
            }

            var screenPoint = Forms.Control.MousePosition;
            var point = element.PointFromScreen(new Point(screenPoint.X, screenPoint.Y));
            return point.X >= 0 && point.X <= element.ActualWidth
                   && point.Y >= 0 && point.Y <= element.ActualHeight;
        }
    }
}
