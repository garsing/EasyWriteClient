using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// Word 改 PageSetup 时会查询活动打印机；默认可用离线打印机会挂死。
    /// 将 <see cref="Word.Application.ActivePrinter"/> 切到本机 PDF/XPS（不改 Windows 系统默认）。
    /// </summary>
    internal static class WordActivePrinterHelper
    {
        private static readonly string[] PreferredInstalledNameSubstrings =
        {
            "Microsoft Print to PDF",
            "Microsoft XPS Document Writer",
        };

        public static bool PreferOnStartup =>
            ConfigManager.Config?.App?.PreferPdfActivePrinterOnStartup ?? true;

        /// <summary>
        /// 将 Word 活动打印机设为已安装的 PDF/XPS。成功返回 true。
        /// </summary>
        public static bool EnsurePreferredActivePrinter(Word.Application app)
        {
            if (app == null)
            {
                return false;
            }

            string targetInstalledName = FindPreferredInstalledPrinterName();
            if (string.IsNullOrEmpty(targetInstalledName))
            {
                return false;
            }

            string current = null;
            try
            {
                current = app.ActivePrinter;
            }
            catch
            {
                // 当前打印机不可用时也继续尝试切换
            }

            if (!string.IsNullOrEmpty(current)
                && current.IndexOf(targetInstalledName, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return TrySetWordActivePrinter(app, targetInstalledName);
        }

        public static string EnsurePreferredActivePrinterWithRestoreToken(Word.Application app)
        {
            if (app == null)
            {
                return null;
            }

            string previous = null;
            try
            {
                previous = app.ActivePrinter;
            }
            catch
            {
                // ignore
            }

            EnsurePreferredActivePrinter(app);
            return previous;
        }

        public static void RestoreActivePrinter(Word.Application app, string previousActivePrinter)
        {
            if (app == null || string.IsNullOrWhiteSpace(previousActivePrinter))
            {
                return;
            }

            try
            {
                if (!string.Equals(app.ActivePrinter, previousActivePrinter, StringComparison.OrdinalIgnoreCase))
                {
                    app.ActivePrinter = previousActivePrinter;
                }
            }
            catch
            {
                // 恢复失败不影响主流程
            }
        }

        private static string FindPreferredInstalledPrinterName()
        {
            var installed = PrinterSettings.InstalledPrinters.Cast<string>().ToList();
            foreach (string key in PreferredInstalledNameSubstrings)
            {
                string hit = installed.FirstOrDefault(n =>
                    n.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0);
                if (!string.IsNullOrEmpty(hit))
                {
                    return hit;
                }
            }

            return null;
        }

        private static bool TrySetWordActivePrinter(Word.Application app, string installedDisplayName)
        {
            var attempts = new List<string> { installedDisplayName };

            string current = null;
            try
            {
                current = app.ActivePrinter;
            }
            catch
            {
                // ignore
            }

            if (!string.IsNullOrEmpty(current))
            {
                int onIdx = current.IndexOf(" on ", StringComparison.OrdinalIgnoreCase);
                if (onIdx >= 0)
                {
                    string portSuffix = current.Substring(onIdx + 1).Trim();
                    attempts.Add($"{installedDisplayName} on {portSuffix}");
                }
            }

            for (int i = 1; i <= 9; i++)
            {
                attempts.Add($"{installedDisplayName} on Ne0{i}:");
            }

            foreach (string candidate in attempts.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    app.ActivePrinter = candidate;
                    string verify = app.ActivePrinter;
                    if (verify.IndexOf(installedDisplayName, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                }
                catch
                {
                    // try next
                }
            }

            return false;
        }
    }
}
