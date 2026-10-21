using System;
using System.Threading.Tasks;

namespace WordAddIn1
{
    public enum ToolStopRiskKind
    {
        None = 0,
        Low = 1,
        High = 2
    }

    /// <summary>点停止后仍在收尾的前端工具。High 灌回期间与 isProcessing 一起挡住输入；Low 与 isProcessing 正交。</summary>
    public static class ToolDrainState
    {
        private static readonly object Gate = new object();
        private static bool _active;
        private static ToolStopRiskKind _risk;
        private static TaskCompletionSource<bool> _idle = CompletedSource();

        public static event Action Changed;

        public static bool Active
        {
            get { lock (Gate) { return _active; } }
        }

        public static ToolStopRiskKind Risk
        {
            get { lock (Gate) { return _risk; } }
        }

        public static bool IsMutatingActive
        {
            get
            {
                lock (Gate)
                {
                    return _active && (_risk == ToolStopRiskKind.Low || _risk == ToolStopRiskKind.High);
                }
            }
        }

        public static Task WaitIdleAsync()
        {
            lock (Gate)
            {
                return _idle.Task;
            }
        }

        public static void Begin(ToolStopRiskKind risk)
        {
            lock (Gate)
            {
                _active = true;
                _risk = risk;
                if (_idle.Task.IsCompleted)
                {
                    _idle = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                }
            }

            Raise();
        }

        public static void End()
        {
            TaskCompletionSource<bool> done;
            lock (Gate)
            {
                if (!_active)
                {
                    return;
                }

                _active = false;
                done = _idle;
            }

            done.TrySetResult(true);
            Raise();
            lock (Gate)
            {
                if (!_active)
                {
                    _risk = ToolStopRiskKind.None;
                }
            }
        }

        public static bool IsHighRollbackActive
        {
            get
            {
                lock (Gate)
                {
                    return _active && _risk == ToolStopRiskKind.High;
                }
            }
        }

        public static string RiskWireName()
        {
            switch (Risk)
            {
                case ToolStopRiskKind.High:
                    return "high";
                case ToolStopRiskKind.Low:
                    return "low";
                default:
                    return "none";
            }
        }

        private static TaskCompletionSource<bool> CompletedSource()
        {
            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            tcs.TrySetResult(true);
            return tcs;
        }

        private static void Raise()
        {
            Action handler = Changed;
            if (handler != null)
            {
                handler();
            }
        }
    }
}
