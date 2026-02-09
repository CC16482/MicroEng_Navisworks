using System;
using System.Threading;
using System.Windows.Threading;

namespace MicroEng.Navisworks
{
    internal sealed class SpaceMapperRunProgressHost : IDisposable
    {
        private Thread _thread;
        private Dispatcher _dispatcher;
        private SpaceMapperRunProgressWindow _window;
        private readonly ManualResetEventSlim _ready = new ManualResetEventSlim(false);

        public static SpaceMapperRunProgressHost Show(SpaceMapperRunProgressState state, Action cancelAction)
        {
            var host = new SpaceMapperRunProgressHost();
            host.Start(state, cancelAction);
            return host;
        }

        private void Start(SpaceMapperRunProgressState state, Action cancelAction)
        {
            _thread = new Thread(() =>
            {
                _dispatcher = Dispatcher.CurrentDispatcher;

                _window = new SpaceMapperRunProgressWindow(state, cancelAction)
                {
                    Topmost = true,
                    ShowInTaskbar = false
                };

                _window.Closed += (_, __) =>
                {
                    try
                    {
                        Dispatcher.CurrentDispatcher.BeginInvokeShutdown(DispatcherPriority.Background);
                    }
                    catch
                    {
                        // ignore
                    }
                };

                _ready.Set();
                _window.Show();

                Dispatcher.Run();
            });

            _thread.IsBackground = true;
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();

            _ready.Wait();
        }

        public void Close()
        {
            if (_dispatcher == null)
            {
                return;
            }

            try
            {
                _dispatcher.BeginInvoke(new Action(() =>
                {
                    if (_window != null)
                    {
                        _window.AllowClose();
                        _window.Close();
                    }
                }));
            }
            catch
            {
                // ignore teardown issues
            }
        }

        public void Dispose()
        {
            Close();
        }
    }
}
