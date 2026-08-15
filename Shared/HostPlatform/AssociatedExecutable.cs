using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace WordAddIn1.HostPlatform
{
    /// <summary>
    /// 查后缀的系统默认可执行文件，再映射到 word/wps/excel/et。禁止裸 ShellExecute。
    /// </summary>
    internal static class AssociatedExecutable
    {
        private const int AssocStrExecutable = 2;
        private const int AssocNoUi = 0x00000001;

        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        private static extern int AssocQueryString(
            int flags,
            int str,
            string pszAssoc,
            string pszExtra,
            StringBuilder pszOut,
            ref uint pcchOut);

        public static bool TryMapDefaultApp(string fullPath, out string app, out string error)
        {
            app = null;
            error = null;
            string exe = TryQuery(fullPath);
            if (string.IsNullOrEmpty(exe))
            {
                error = "无法解析该后缀的系统默认程序，请显式指定 app=word|wps|excel|et";
                return false;
            }

            string name = Path.GetFileName(exe) ?? "";
            if (name.Equals("WINWORD.EXE", StringComparison.OrdinalIgnoreCase))
            {
                app = "word";
                return true;
            }

            if (name.Equals("EXCEL.EXE", StringComparison.OrdinalIgnoreCase))
            {
                app = "excel";
                return true;
            }

            if (name.Equals("et.exe", StringComparison.OrdinalIgnoreCase))
            {
                app = "et";
                return true;
            }

            if (name.Equals("wps.exe", StringComparison.OrdinalIgnoreCase)
                || name.IndexOf("wps", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                app = "wps";
                return true;
            }

            error = "系统默认程序不是 Word/WPS/Excel/et（查到: " + exe + "），请显式指定 app";
            return false;
        }

        private static string TryQuery(string fullPath)
        {
            string ext = Path.GetExtension(fullPath);
            if (string.IsNullOrEmpty(ext))
            {
                return null;
            }

            uint size = 0;
            AssocQueryString(AssocNoUi, AssocStrExecutable, ext, null, null, ref size);
            if (size == 0)
            {
                return null;
            }

            var sb = new StringBuilder((int)size);
            int hr = AssocQueryString(AssocNoUi, AssocStrExecutable, ext, null, sb, ref size);
            if (hr != 0)
            {
                return null;
            }

            return sb.ToString();
        }
    }
}
