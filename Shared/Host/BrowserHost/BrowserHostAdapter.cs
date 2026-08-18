using System.Threading.Tasks;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 浏览器适配层入口。本批仅落地易写浏览器 Navigate；扩展附着另册。
    /// </summary>
    public static class BrowserHostAdapter
    {
        public static bool IsAgentPageLive(string tabUuid)
        {
            return YiWriteBrowserHost.IsPageLive(tabUuid);
        }

        public static bool TryShowAgentPage(string tabUuid)
        {
            return YiWriteBrowserHost.TryShowPage(tabUuid);
        }

        public static Task<BrowserNavigateResult> NavigateAgentAsync(
            string url,
            bool visible,
            string existingTabUuid)
        {
            return YiWriteBrowserHost.NavigateAsync(url, visible, existingTabUuid);
        }
    }
}
