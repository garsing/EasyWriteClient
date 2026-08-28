using System;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// Shared → 宿主（Plugin / Desktop）回调，避免 Shared 直接依赖 VSTO 类型。
    /// Plugin 在启动时注册；未注册时相关通知为 no-op。
    /// </summary>
    public static class HostCallbacks
    {
        /// <summary>Desktop：F_open_document 创建/附着 Word 后刷新 WsClient Registry。</summary>
        public static Action<object> WordApplicationResolved { get; set; }

        /// <summary>Desktop：F_open_document 创建/附着 Excel 后写回宿主缓存并刷新打开文件。</summary>
        public static Action<object> ExcelApplicationResolved { get; set; }

        /// <summary>Desktop：F_open_document 创建/附着 PowerPoint 后写回宿主缓存并刷新打开文件。</summary>
        public static Action<object> PowerPointApplicationResolved { get; set; }

        /// <summary>打开成功后刷新侧栏「打开文件」探测器。</summary>
        public static Action OpenFilesRefresh { get; set; }

        /// <summary>
        /// Desktop：易写浏览页建/更新后写入侧栏「打开文件」。
        /// 参数：channelId, displayName, url。
        /// </summary>
        public static Action<string, string, string> BrowserOpenFileUpsert { get; set; }

        /// <summary>Desktop：易写浏览窗关闭后从侧栏移除对应项（参数为 channelId）。</summary>
        public static Action<string> BrowserOpenFileRemove { get; set; }

        /// <summary>用户登录成功后通知各任务窗格/桌面 UI 刷新。</summary>
        public static Func<Task> NotifyUserLoggedInAllAsync { get; set; }

        /// <summary>
        /// Desktop 注册：返回 { showInteraction, autoCompactEnabled, autoFloatEnabled, autoFloatIdleSeconds }。
        /// Plugin 不注册 → 设置页不展示「交互」。
        /// </summary>
        public static Func<object> GetDesktopInteractionSettings { get; set; }

        /// <summary>
        /// Desktop 注册：key = autoCompactEnabled | autoFloatEnabled | autoFloatIdleSeconds。
        /// </summary>
        public static Action<string, object> SetDesktopInteractionSetting { get; set; }

        /// <summary>Desktop：请求进入缩小版窗口形态；Plugin 不注册则为 no-op。编排第 1 步 / 手动缩小复用。</summary>
        public static Action RequestCompact { get; set; }

        /// <summary>
        /// Desktop：前台编排（缩 → 目标窗 → 易写）。Plugin 不注册则为 no-op。
        /// </summary>
        public static Action<ForegroundDanceRequest> RequestForegroundDance { get; set; }

        /// <summary>
        /// 已过时：可见浏览改为 <see cref="RaiseForegroundDance"/>（Kind=Open）。
        /// </summary>
        public static Action BringDesktopToFront { get; set; }

        /// <summary>Desktop：浏览窗关闭/隐藏后取消置顶。</summary>
        public static Action ClearDesktopTopMost { get; set; }

        public static void RaiseWordApplicationResolved(object wordApplication)
        {
            WordApplicationResolved?.Invoke(wordApplication);
        }

        public static void RaiseExcelApplicationResolved(object excelApplication)
        {
            ExcelApplicationResolved?.Invoke(excelApplication);
        }

        public static void RaisePowerPointApplicationResolved(object powerPointApplication)
        {
            PowerPointApplicationResolved?.Invoke(powerPointApplication);
        }

        public static void RaiseRequestCompact()
        {
            RequestCompact?.Invoke();
        }

        public static void RaiseForegroundDance(ForegroundDanceRequest request)
        {
            if (request == null || request.Kind == ForegroundDanceKind.None)
            {
                return;
            }

            RequestForegroundDance?.Invoke(request);
        }

        public static void RaiseBringDesktopToFront()
        {
            BringDesktopToFront?.Invoke();
        }

        public static void RaiseClearDesktopTopMost()
        {
            ClearDesktopTopMost?.Invoke();
        }

        public static void RaiseOpenFilesRefresh()
        {
            OpenFilesRefresh?.Invoke();
        }

        public static void RaiseBrowserOpenFileUpsert(string channelId, string displayName, string url)
        {
            BrowserOpenFileUpsert?.Invoke(channelId, displayName, url);
        }

        public static void RaiseBrowserOpenFileRemove(string channelId)
        {
            BrowserOpenFileRemove?.Invoke(channelId);
        }

        public static Task RaiseNotifyUserLoggedInAllAsync()
        {
            Func<Task> handler = NotifyUserLoggedInAllAsync;
            if (handler == null)
            {
                return Task.CompletedTask;
            }

            return handler() ?? Task.CompletedTask;
        }
    }
}
