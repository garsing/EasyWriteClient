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

        /// <summary>用户登录成功后通知各任务窗格/桌面 UI 刷新。</summary>
        public static Func<Task> NotifyUserLoggedInAllAsync { get; set; }

        /// <summary>Desktop：请求进入缩小版窗口形态；Plugin 不注册则为 no-op。</summary>
        public static Action RequestCompact { get; set; }

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

        public static void RaiseOpenFilesRefresh()
        {
            OpenFilesRefresh?.Invoke();
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
