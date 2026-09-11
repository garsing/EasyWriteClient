using System;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1
{
    /// <summary>
    /// 解析 PowerPoint.Application：缓存 → GetActiveObject → 可选 new。
    /// 跨程序集公开 API 用 object，避免 Desktop 无 PowerPoint Interop 时 CS1748。
    /// 不 Quit 用户 PowerPoint；关窗后清缓存 / 补文档窗，避免「有稿无窗」空壳。
    /// </summary>
    public static class PowerPointApplicationResolver
    {
        private static readonly object Gate = new object();
        private static PowerPoint.Application _hosted;

        public static void Attach(object application)
        {
            lock (Gate)
            {
                if (application == null)
                {
                    _hosted = null;
                    return;
                }

                PowerPoint.Application typed = application as PowerPoint.Application;
                if (typed == null)
                {
                    try
                    {
                        typed = (PowerPoint.Application)application;
                    }
                    catch (Exception)
                    {
                        return;
                    }
                }

                _hosted = typed;
            }
        }

        /// <summary>
        /// 无打开演示时清掉托管缓存（不 Quit）。对齐 Excel Attach(null) 的空闲释放意图。
        /// </summary>
        public static void ReleaseHostedIfIdle()
        {
            if (OfficeStaScheduler.ShouldHop)
            {
                OfficeStaScheduler.Invoke(ReleaseHostedIfIdle);
                return;
            }

            lock (Gate)
            {
                if (_hosted == null)
                {
                    return;
                }

                try
                {
                    if (_hosted.Presentations.Count > 0)
                    {
                        return;
                    }
                }
                catch (Exception)
                {
                }

                _hosted = null;
            }
        }

        public static bool TryResolve(
            out object application,
            out string error,
            bool createIfMissing = false,
            bool makeVisible = true)
        {
            if (OfficeStaScheduler.ShouldHop)
            {
                object app = null;
                string err = null;
                bool ok = OfficeStaScheduler.Invoke(() =>
                    TryResolve(out app, out err, createIfMissing, makeVisible));
                application = app;
                error = err;
                return ok;
            }

            application = null;
            error = null;
            if (!TryResolveCore(out PowerPoint.Application typed, out error, createIfMissing, makeVisible))
            {
                return false;
            }

            application = typed;
            return true;
        }

        internal static bool TryResolveTyped(
            out PowerPoint.Application application,
            out string error,
            bool createIfMissing = false,
            bool makeVisible = true)
        {
            if (OfficeStaScheduler.ShouldHop)
            {
                PowerPoint.Application app = null;
                string err = null;
                bool ok = OfficeStaScheduler.Invoke(() =>
                    TryResolveTyped(out app, out err, createIfMissing, makeVisible));
                application = app;
                error = err;
                return ok;
            }

            return TryResolveCore(out application, out error, createIfMissing, makeVisible);
        }

        private static bool TryResolveCore(
            out PowerPoint.Application application,
            out string error,
            bool createIfMissing,
            bool makeVisible)
        {
            application = null;
            error = null;

            lock (Gate)
            {
                if (_hosted != null)
                {
                    try
                    {
                        var _ = _hosted.Name;
                        application = _hosted;
                        if (makeVisible)
                        {
                            EnsureVisibleCore(application);
                        }

                        return true;
                    }
                    catch (Exception)
                    {
                        _hosted = null;
                    }
                }
            }

            try
            {
                object raw = Marshal.GetActiveObject("PowerPoint.Application");
                application = raw as PowerPoint.Application ?? (PowerPoint.Application)raw;
                if (application != null)
                {
                    Attach(application);
                    if (makeVisible)
                    {
                        EnsureVisibleCore(application);
                    }

                    return true;
                }
            }
            catch (Exception)
            {
            }

            if (!createIfMissing)
            {
                error = "PowerPoint 应用程序不可用；请先启动 PowerPoint。";
                return false;
            }

            try
            {
                application = new PowerPoint.Application();
                Attach(application);
                if (makeVisible)
                {
                    EnsureVisibleCore(application);
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "无法创建 PowerPoint.Application（是否未安装 PowerPoint？）: " + ex.Message;
                application = null;
                return false;
            }
        }

        public static void EnsureVisible(object application)
        {
            EnsureVisibleCore(application as PowerPoint.Application);
        }

        /// <summary>
        /// COM 持有 Presentation 但 Windows=0 时补文档窗（易写关窗后空壳的兜底）。
        /// </summary>
        public static void EnsurePresentationHasWindow(object presentation)
        {
            EnsurePresentationHasWindowCore(presentation as PowerPoint.Presentation);
        }

        private static void EnsureVisibleCore(PowerPoint.Application application)
        {
            if (application == null)
            {
                return;
            }

            try
            {
                application.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            }
            catch (Exception)
            {
                try
                {
                    // EmbedInteropTypes 下部分环境用 int
                    application.Visible = (Microsoft.Office.Core.MsoTriState)(-1);
                }
                catch (Exception)
                {
                }
            }

            try
            {
                foreach (PowerPoint.Presentation presentation in application.Presentations)
                {
                    EnsurePresentationHasWindowCore(presentation);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void EnsurePresentationHasWindowCore(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return;
            }

            try
            {
                if (presentation.Windows != null && presentation.Windows.Count > 0)
                {
                    return;
                }

                presentation.NewWindow();
                try
                {
                    EasyWriteDiagnostics.Log(
                        DebugCategory.OpenFiles,
                        "[PowerPointApplicationResolver] NewWindow for windowless presentation "
                            + (presentation.Name ?? ""));
                }
                catch (Exception)
                {
                }
            }
            catch (Exception ex)
            {
                try
                {
                    EasyWriteDiagnostics.Log(
                        DebugCategory.OpenFiles,
                        "[PowerPointApplicationResolver] NewWindow failed: " + ex.Message);
                }
                catch (Exception)
                {
                }
            }
        }
    }
}
