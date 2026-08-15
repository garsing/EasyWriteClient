using System;
using System.Runtime.InteropServices;
using WordAddIn1.HostPlatform;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1
{
    public static class ExcelApplicationResolver
    {
        public static bool TryResolve(
            out Excel.Application application,
            out string error,
            bool createIfMissing = false,
            bool makeVisible = true)
        {
            application = null;
            error = null;

            try
            {
                application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                if (application != null)
                {
                    if (makeVisible)
                    {
                        EnsureVisible(application);
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
                if (makeVisible)
                {
                    EnsureVisible(application);
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

        public static void EnsureVisible(Excel.Application application)
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
