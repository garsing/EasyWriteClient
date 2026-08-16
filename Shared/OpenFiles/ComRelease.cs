using System;
using System.Runtime.InteropServices;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// 释放 Office/WPS COM RCW。长期持有 Application/Workbook 会阻止 et 进程退出。
    /// </summary>
    internal static class ComRelease
    {
        public static void Safe(object com)
        {
            if (com == null || !Marshal.IsComObject(com))
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(com);
            }
            catch (Exception)
            {
            }
        }

        /// <summary>仅 Release 一次；用于渠道替换 RCW，避免 FinalRelease 拆死仍在使用的 COM。</summary>
        public static void ReleaseOnce(object com)
        {
            if (com == null || !Marshal.IsComObject(com))
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(com);
            }
            catch (Exception)
            {
            }
        }

        public static void CollectPending()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception)
            {
            }
        }
    }
}
