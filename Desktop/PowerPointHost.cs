using System;
using WordAddIn1;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// Desktop 侧 PowerPoint.Application 缓存：对齐 <see cref="ExcelHost"/>。
    /// 不直接引用 PowerPoint Interop；经 Shared Resolver 读写。不 Quit 用户 PPT。
    /// </summary>
    internal static class PowerPointHost
    {
        private static readonly object Gate = new object();
        private static object _app;

        public static object GetOrAttach(bool createIfMissing = false)
        {
            if (OfficeStaScheduler.ShouldHop)
            {
                return OfficeStaScheduler.Invoke(() => GetOrAttach(createIfMissing));
            }

            lock (Gate)
            {
                if (PowerPointApplicationResolver.TryResolve(
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

        public static void Attach(object powerPointApplication)
        {
            if (powerPointApplication == null)
            {
                return;
            }

            lock (Gate)
            {
                PowerPointApplicationResolver.Attach(powerPointApplication);
                _app = powerPointApplication;
            }
        }

        /// <summary>关闭 Desktop：不 Quit 用户 PowerPoint。</summary>
        public static void Shutdown()
        {
            if (OfficeStaScheduler.ShouldHop)
            {
                OfficeStaScheduler.Invoke(Shutdown);
                return;
            }

            lock (Gate)
            {
                _app = null;
                PowerPointApplicationResolver.Attach(null);
            }
        }
    }
}
