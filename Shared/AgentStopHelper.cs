using System.Threading;
using System.Threading.Tasks;

namespace WordAddIn1
{
    internal static class AgentStopHelper
    {
        public const int WatchdogMs = 3000;

        public static async Task RequestStopAsync(
            WsClient wsClient,
            string conversationId,
            CancellationTokenSource runCts)
        {
            if (!string.IsNullOrEmpty(conversationId)
                && conversationId != "-1"
                && wsClient != null)
            {
                await wsClient.SendRunCancelAsync(conversationId).ConfigureAwait(true);
            }

            wsClient?.CancelActiveInvokeLocal();
            AgentRunCancellation.RequestCancel();
            if (runCts != null && !runCts.IsCancellationRequested)
            {
                runCts.Cancel();
            }

            string active = wsClient != null ? wsClient.ActiveInvokeMethod : null;
            if (string.IsNullOrWhiteSpace(active))
            {
                return;
            }

            ToolStopRiskKind risk = ConfigManager.ClassifyToolStopRisk(active);
            ToolDrainState.Begin(risk);
            if (risk == ToolStopRiskKind.None)
            {
                ToolDrainAbort.AbortNoneRisk();
            }
        }
    }
}
