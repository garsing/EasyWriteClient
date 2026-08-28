using System;
using System.Windows.Forms;
using WordAddIn1;
using WordAddIn1.HostPlatform;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// Desktop 前台编排：缩 → 目标窗焦点（不 Maximize）→ 易写再焦点。
    /// 计数仅内存；关进程清零。Plugin 不持有本类型。
    /// </summary>
    internal sealed class ForegroundDance : IDisposable
    {
        private const int StepDelayMs = 150;

        private readonly MainForm _form;
        private readonly object _gate = new object();
        private int _count;
        private Timer _step3Timer;

        public ForegroundDance(MainForm form)
        {
            _form = form ?? throw new ArgumentNullException(nameof(form));
        }

        public void Request(ForegroundDanceRequest request)
        {
            if (request == null || _form.IsDisposed)
            {
                return;
            }

            if (_form.InvokeRequired)
            {
                _form.BeginInvoke(new Action(() => Request(request)));
                return;
            }

            TryRun(request);
        }

        private void TryRun(ForegroundDanceRequest request)
        {
            if (!WindowLayoutStore.GetAutoCompactEnabled())
            {
                return;
            }

            lock (_gate)
            {
                if (request.Kind == ForegroundDanceKind.Operate && _count >= 1)
                {
                    return;
                }

                // 过门即占位，避免并发 Operate 都走进三步
                _count = 1;
            }

            if (_form.IsFloatBall)
            {
                _form.LeaveFloatBall();
            }

            if (!_form.IsCompactLayout)
            {
                _form.SetLayoutMode(WindowLayoutMode.Compact);
            }

            int hwnd = request.TargetHwnd;
            if (hwnd == 0 && !string.IsNullOrWhiteSpace(request.ChannelId))
            {
                hwnd = ForegroundDanceHwnd.TryResolve(request.ChannelId);
            }

            if (hwnd != 0)
            {
                NativeWindowActivate.FocusWithoutMaximize(hwnd);
            }

            ScheduleFocusEasyWrite();
        }

        private void ScheduleFocusEasyWrite()
        {
            if (_step3Timer != null)
            {
                _step3Timer.Stop();
                _step3Timer.Dispose();
            }

            _step3Timer = new Timer { Interval = StepDelayMs };
            Timer timer = _step3Timer;
            timer.Tick += (_, __) =>
            {
                timer.Stop();
                if (ReferenceEquals(_step3Timer, timer))
                {
                    _step3Timer.Dispose();
                    _step3Timer = null;
                }
                else
                {
                    timer.Dispose();
                }

                if (!_form.IsDisposed)
                {
                    _form.FocusEasyWriteAfterDance();
                }
            };
            timer.Start();
        }

        public void Dispose()
        {
            if (_step3Timer != null)
            {
                _step3Timer.Stop();
                _step3Timer.Dispose();
                _step3Timer = null;
            }
        }
    }
}
