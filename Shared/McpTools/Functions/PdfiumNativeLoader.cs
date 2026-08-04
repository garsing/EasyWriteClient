using System;
using System.IO;
using PdfiumViewer;

namespace WordAddIn1
{
    /// <summary>
    /// VSTO 宿主下 PdfiumViewer 默认搜索路径可能失效，需在首次调用前显式解析 pdfium.dll。
    /// </summary>
    internal static class PdfiumNativeLoader
    {
        private static bool _initialized;
        private static readonly object SyncRoot = new object();

        public static void EnsureLoaded()
        {
            if (_initialized)
            {
                return;
            }

            lock (SyncRoot)
            {
                if (_initialized)
                {
                    return;
                }

                PdfiumResolver.Resolve += (_, e) =>
                {
                    e.PdfiumFileName = ResolvePdfiumPath();
                };

                _initialized = true;
            }
        }

        private static string ResolvePdfiumPath()
        {
            string archFolder = Environment.Is64BitProcess ? "x64" : "x86";
            string addInDir = GetAddInDirectory();

            string[] candidates =
            {
                Path.Combine(addInDir, archFolder, "pdfium.dll"),
                Path.Combine(addInDir, archFolder, "Pdfium.dll"),
                Path.Combine(addInDir, "pdfium.dll"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? string.Empty, archFolder, "pdfium.dll"),
            };

            foreach (string candidate in candidates)
            {
                if (!string.IsNullOrWhiteSpace(candidate) && File.Exists(candidate))
                {
                    System.Diagnostics.Debug.WriteLine($"[PdfiumNativeLoader] 使用 pdfium: {candidate}");
                    return candidate;
                }
            }

            throw new FileNotFoundException(
                $"未找到 pdfium.dll。进程位数={archFolder}，插件目录={addInDir}。请确认已复制 {archFolder}\\pdfium.dll 到输出目录。");
        }

        private static string GetAddInDirectory()
        {
            var assembly = typeof(PdfiumNativeLoader).Assembly;
            if (!string.IsNullOrEmpty(assembly.Location))
            {
                return Path.GetDirectoryName(assembly.Location);
            }

            if (!string.IsNullOrEmpty(assembly.CodeBase))
            {
                return Path.GetDirectoryName(new Uri(assembly.CodeBase).LocalPath);
            }

            return AppDomain.CurrentDomain.BaseDirectory ?? string.Empty;
        }
    }
}
