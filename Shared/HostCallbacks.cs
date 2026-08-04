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
        /// <summary>用户登录成功后通知各任务窗格/桌面 UI 刷新。</summary>
        public static Func<Task> NotifyUserLoggedInAllAsync { get; set; }

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
