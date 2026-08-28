using System;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1
{
    /// <summary>
    /// 解析 PowerPoint.Application：缓存 → GetActiveObject → 可选 new。
    /// 跨程序集公开 API 用 object，避免 Desktop 无 PowerPoint Interop 时 CS1748。
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

        public static bool TryResolve(
            out object application,
            out string error,
            bool createIfMissing = false,
            bool makeVisible = true)
        {
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

        }
    }
}
