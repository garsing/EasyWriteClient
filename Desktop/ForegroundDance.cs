using System;
using WordAddIn1;
using WordAddIn1.HostPlatform;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// Desktop 前台编排：缩 → 目标窗焦点（不 Maximize）→ 易写再焦点。
    /// 第 3 步必须同步：WinForms Timer 走消息泵，COM 写页占住 UI 时 Tick 到不了。
    /// 计数仅内存；关进程清零。Plugin 不持有本类型。
    /// </summary>
    internal sealed class ForegroundDance : IDisposable
    {
        private readonly MainForm _form;
        private readonly object _gate = new object();
        private int _count;
        private bool _operateHold;

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

            // 一开始就叠好：文档第二层、易写第一层。Operate 写页期间保持 TopMost，
            // 避免 COM 把文档重新盖上来；handler 结束立刻关掉。
            _operateHold = request.Kind == ForegroundDanceKind.Operate;
            _form.FocusEasyWriteAfterDance(keepTopMost: _operateHold);
        }

        public void CompletePending()
        {
            if (_form.IsDisposed)
            {
                return;
            }

            if (_form.InvokeRequired)
            {
                _form.BeginInvoke(new Action(CompletePending));
                return;
            }

            if (!_operateHold)
            {
                return;
            }

            _form.ReleaseDanceTopMost();
            _operateHold = false;
        }

        public void Dispose()
        {
        }
    }
}
