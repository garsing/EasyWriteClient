using System;
using WordAddIn1;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// Desktop 侧 Excel.Application 缓存：对齐 <see cref="WordHost"/>。
    /// 不直接引用 Excel Interop（Desktop 工程无该引用）；经 Shared Resolver 读写。
    /// </summary>
    internal static class ExcelHost
    {
        private static readonly object Gate = new object();
        private static object _app;

        /// <param name="createIfMissing">
        /// false（默认）：仅缓存 / GetActiveObject；没有则返回 null。
        /// true：无运行实例时才 new Application。
        /// </param>
        public static object GetOrAttach(bool createIfMissing = false)
        {
            lock (Gate)
            {
                if (ExcelApplicationResolver.TryResolve(
                        out object app,
                        out _,
                        createIfMissing: createIfMissing,
                        makeVisible: false)
                    && app != null)
                {
                    _app = app;
                    return _app;
                }

                _app = null;
                return null;
            }
        }

        /// <summary>F_open 等路径已拿到 Application 时写回缓存。</summary>
        public static void Attach(object excelApplication)
        {
            if (excelApplication == null)
            {
                return;
            }

            lock (Gate)
            {
                ExcelApplicationResolver.Attach(excelApplication);
                _app = excelApplication;
            }
        }

        /// <summary>关闭 Desktop：不 Quit 用户 Excel。</summary>
        public static void Shutdown()
        {
            lock (Gate)
            {
                _app = null;
                ExcelApplicationResolver.Attach(null);
            }
        }
    }
}
