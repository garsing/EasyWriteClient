using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentHost
{
    /// <summary>
    /// 文字宿主 / 检查点 / 表后落点诊断。总闸 + DebugCategories 的 host（或 all）。
    /// </summary>
    internal static class HostProbe
    {
        /// <summary>总闸。false 时不算详情、不碰 COM。排障再改 true，且 config 含 host 或 all。</summary>
        internal const bool Enable = false;

        internal static bool Enabled => Enable && EasyWriteDiagnostics.IsEnabled(DebugCategory.Host);

        public static void Log(string site, string detail)
        {
            if (!Enabled)
            {
                return;
            }

            EasyWriteDiagnostics.Log(
                DebugCategory.Host,
                site + " " + (detail ?? ""));
        }

        public static string ArgsBrief(IReadOnlyDictionary<string, object> args)
        {
            if (args == null)
            {
                return "args=null";
            }

            object raw;
            string id = args.TryGetValue("channel_id", out raw) && raw != null
                ? Convert.ToString(raw)
                : null;
            return "ch=" + (string.IsNullOrWhiteSpace(id) ? "(null)" : id.Trim())
                + " hasAction=" + args.ContainsKey("action");
        }

        public static string Ex(Exception ex)
        {
            if (ex == null)
            {
                return "ex=null";
            }

            string hr = "";
            if (ex is COMException com)
            {
                hr = " hr=0x" + com.ErrorCode.ToString("X8");
            }

            string stack = ex.StackTrace ?? "";
            if (stack.Length > 400)
            {
                stack = stack.Substring(0, 400);
            }

            return "type=" + ex.GetType().Name + " msg=" + ex.Message + hr + " stack=" + stack.Replace('\r', '|').Replace('\n', '|');
        }

        public static void Touch(string site, Word.Document document)
        {
            if (!Enabled)
            {
                return;
            }

            if (document == null)
            {
                Log(site, "doc=null");
                return;
            }

            string name = Safe(() => document.Name);
            string full = Safe(() => document.FullName);
            string paras = Safe(() => document.Paragraphs.Count.ToString());
            string appName = Safe(() =>
            {
                Word.Application app = document.Application;
                return app == null ? "(app-null)" : Convert.ToString(app.Name);
            });
            string appDocs = Safe(() =>
            {
                Word.Application app = document.Application;
                return app == null ? "(app-null)" : app.Documents.Count.ToString();
            });
            Log(site, "Name=" + name + " FullName=" + full + " Paragraphs=" + paras
                + " App.Name=" + appName + " App.Documents=" + appDocs);
        }

        public static void TouchLate(string site, object document)
        {
            if (!Enabled)
            {
                return;
            }

            if (document == null)
            {
                Log(site, "late-doc=null");
                return;
            }

            string name = SafeLate(document, "Name");
            string full = SafeLate(document, "FullName");
            Log(site, "late Name=" + name + " FullName=" + full);
        }

        private static string Safe(Func<string> read)
        {
            try
            {
                return read() ?? "(null)";
            }
            catch (Exception ex)
            {
                return "ERR:" + ex.Message;
            }
        }

        private static string SafeLate(object document, string prop)
        {
            try
            {
                object value = WordAddIn1.OpenFiles.WpsCom.GetProperty(document, prop);
                return value == null ? "(null)" : Convert.ToString(value);
            }
            catch (Exception ex)
            {
                return "ERR:" + ex.Message;
            }
        }
    }
}
