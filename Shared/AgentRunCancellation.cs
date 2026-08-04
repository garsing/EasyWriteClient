using System;
using System.Threading;

namespace WordAddIn1
{
    /// <summary>
    /// Agent 单次 run 的协作式取消 token（链到 TaskPane 的 run CTS）。
    /// </summary>
    public static class AgentRunCancellation
    {
        private static CancellationTokenSource _linked;

        public static CancellationToken Token => _linked?.Token ?? CancellationToken.None;

        public static void BeginRun(CancellationTokenSource taskPaneRunCts)
        {
            EndRun();
            if (taskPaneRunCts == null)
            {
                return;
            }
            _linked = CancellationTokenSource.CreateLinkedTokenSource(taskPaneRunCts.Token);
        }

        public static void EndRun()
        {
            _linked?.Dispose();
            _linked = null;
        }

        public static void ThrowIfCancelled()
        {
            Token.ThrowIfCancellationRequested();
        }
    }
}
