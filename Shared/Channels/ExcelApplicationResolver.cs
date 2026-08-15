using System;
using System.Runtime.InteropServices;
using WordAddIn1.HostPlatform;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1
{
    /// <summary>
    /// 解析 Excel.Application：优先宿主注入缓存，其次 GetActiveObject，可选 new（对齐 WordApplicationResolver）。
    /// 跨程序集公开 API 一律 object，避免 Desktop 无 Excel Interop 时 CS1748。
    /// </summary>
    public static class ExcelApplicationResolver
    {
        private static readonly object Gate = new object();
        private static Excel.Application _hosted;

        /// <summary>Desktop / Host 打开成功后写回，供探测与后续工具共用同一实例。</summary>
        public static void Attach(object application)
        {
            lock (Gate)
            {
                _hosted = application as Excel.Application;
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
            if (!TryResolveCore(out Excel.Application typed, out error, createIfMissing, makeVisible))
            {
                return false;
            }

            application = typed;
            return true;
        }

        /// <summary>Shared 内 Excel COM 路径使用。</summary>
        internal static bool TryResolveTyped(
            out Excel.Application application,
            out string error,
            bool createIfMissing = false,
            bool makeVisible = true)
        {
            return TryResolveCore(out application, out error, createIfMissing, makeVisible);
        }

        private static bool TryResolveCore(
            out Excel.Application application,
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
                application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
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
            catch (COMException)
            {
            }
            catch (Exception ex)
            {
                error = "附着已有 Excel 失败: " + ex.Message;
                return false;
            }

            if (!createIfMissing)
            {
                error = "Excel 应用程序不可用；请先启动 Excel。";
                return false;
            }

            try
            {
                application = new Excel.Application();
                Attach(application);
                if (makeVisible)
                {
                    EnsureVisibleCore(application);
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "无法创建 Excel.Application（是否未安装 Excel？）: " + ex.Message;
                application = null;
                return false;
            }
        }

        public static void EnsureVisible(object application)
        {
            EnsureVisibleCore(application as Excel.Application);
        }

        private static void EnsureVisibleCore(Excel.Application application)
        {
            if (application == null)
            {
                return;
            }

            try
            {
                application.Visible = true;
            }
            catch (Exception)
            {
            }

            try
            {
                application.UserControl = true;
            }
            catch (Exception)
            {
            }

            try
            {
                NativeWindowActivate.BringToFront(application.Hwnd);
            }
            catch (Exception)
            {
            }
        }
    }
}
