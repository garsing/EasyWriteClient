using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAddIn1
{
    /// <summary>
    /// Desktop 专用：Office COM 的 STA 工作线程 + 消息泵。
    /// Plugin 不调用 <see cref="Start"/>；RCW 仍挂在 Word 主 STA。
    /// </summary>
    public static class OfficeStaScheduler
    {
        private static readonly object StartGate = new object();
        private static Thread _thread;
        private static Form _pump;
        private static SynchronizationContext _ctx;
        private static ManualResetEventSlim _ready;
        private static int _threadId;

        /// <summary>主窗，供浏览/WebView2 从 STA 跳回 UI。</summary>
        public static Control UiControl { get; set; }

        public static bool IsEnabled
        {
            get
            {
                Form pump = _pump;
                return pump != null && !pump.IsDisposed;
            }
        }

        public static bool IsCurrentThread
        {
            get { return _threadId != 0 && Environment.CurrentManagedThreadId == _threadId; }
        }

        public static bool ShouldHop
        {
            get { return IsEnabled && !IsCurrentThread; }
        }

        public static SynchronizationContext Context
        {
            get { return _ctx; }
        }

        public static void Start()
        {
            lock (StartGate)
            {
                if (IsEnabled)
                {
                    return;
                }

                _ready = new ManualResetEventSlim(false);
                _thread = new Thread(ThreadMain)
                {
                    Name = "EasyWrite-OfficeSTA",
                    IsBackground = true
                };
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
                if (!_ready.Wait(TimeSpan.FromSeconds(15)))
                {
                    throw new TimeoutException("Office STA 线程启动超时");
                }
            }
        }

        public static void Shutdown()
        {
            lock (StartGate)
            {
                Form pump = _pump;
                UiControl = null;
                _ctx = null;
                _pump = null;
                _threadId = 0;
                if (pump != null && !pump.IsDisposed)
                {
                    try
                    {
                        if (pump.IsHandleCreated && pump.InvokeRequired)
                        {
                            pump.Invoke(new Action(() =>
                            {
                                if (!pump.IsDisposed)
                                {
                                    pump.Close();
                                }
                            }));
                        }
                        else if (!pump.IsDisposed)
                        {
                            pump.Close();
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                _thread = null;
                if (_ready != null)
                {
                    _ready.Dispose();
                    _ready = null;
                }
            }
        }

        public static void Invoke(Action action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            if (!ShouldHop)
            {
                action();
                return;
            }

            _pump.Invoke(action);
        }

        public static T Invoke<T>(Func<T> func)
        {
            if (func == null)
            {
                throw new ArgumentNullException(nameof(func));
            }

            if (!ShouldHop)
            {
                return func();
            }

            T result = default(T);
            Exception error = null;
            _pump.Invoke(new Action(() =>
            {
                try
                {
                    result = func();
                }
                catch (Exception ex)
                {
                    error = ex;
                }
            }));
            if (error != null)
            {
                throw error;
            }

            return result;
        }

        public static Task<T> InvokeAsync<T>(Func<Task<T>> action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            if (!ShouldHop)
            {
                return action();
            }

            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            _pump.BeginInvoke(new Action(async () =>
            {
                try
                {
                    tcs.TrySetResult(await action().ConfigureAwait(true));
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            }));
            return tcs.Task;
        }

        /// <summary>从专属 STA 跳回 Desktop 主窗线程（浏览窗 / WebView2）。</summary>
        public static Task<T> InvokeOnUiAsync<T>(Func<Task<T>> action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            Control ui = UiControl;
            if (ui == null || ui.IsDisposed || !ui.InvokeRequired)
            {
                return action();
            }

            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            ui.BeginInvoke(new Action(async () =>
            {
                try
                {
                    tcs.TrySetResult(await action().ConfigureAwait(true));
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            }));
            return tcs.Task;
        }

        private static void ThreadMain()
        {
            _threadId = Environment.CurrentManagedThreadId;
            _pump = new Form
            {
                Text = "EasyWrite-OfficeSTA",
                ShowInTaskbar = false,
                FormBorderStyle = FormBorderStyle.FixedToolWindow,
                Opacity = 0,
                Size = new Size(1, 1),
                StartPosition = FormStartPosition.Manual,
                Location = new Point(-32000, -32000)
            };
            _pump.Shown += (s, e) =>
            {
                _pump.Hide();
                _ctx = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();
                SynchronizationContext.SetSynchronizationContext(_ctx);
                _ready.Set();
            };
            Application.Run(_pump);
        }
    }
}
