using WordAddIn1.BrowserHost;
using WordAddIn1.Terminal;

namespace WordAddIn1
{
    /// <summary>无风险档：拆等待与本机命令，不写工具名。</summary>
    internal static class ToolDrainAbort
    {
        public static void AbortNoneRisk()
        {
            try
            {
                YiWriteBrowserHost.CancelActiveWaits();
            }
            catch
            {
            }

            try
            {
                TerminalSession session;
                if (TerminalSessionHub.TryGet(ConversationContext.CurrentId, out session) && session != null)
                {
                    session.KillCurrentCommand();
                }
            }
            catch
            {
            }
        }
    }
}
