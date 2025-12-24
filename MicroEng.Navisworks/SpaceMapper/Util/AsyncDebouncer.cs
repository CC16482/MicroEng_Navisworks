using System;
using System.Threading;
using System.Threading.Tasks;

namespace MicroEng.Navisworks.SpaceMapper.Util
{
    public sealed class AsyncDebouncer
    {
        private readonly TimeSpan _delay;
        private CancellationTokenSource _cts = new CancellationTokenSource();

        public AsyncDebouncer(TimeSpan delay) { _delay = delay; }

        public void Debounce(Func<CancellationToken, Task> action)
        {
            _cts.Cancel();
            _cts.Dispose();
            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            Task.Run(async () =>
            {
                try
                {
                    await Task.Delay(_delay, token).ConfigureAwait(false);
                    if (!token.IsCancellationRequested)
                        await action(token).ConfigureAwait(false);
                }
                catch (TaskCanceledException) { }
            }, token);
        }

        public void Cancel()
        {
            _cts.Cancel();
        }
    }
}
