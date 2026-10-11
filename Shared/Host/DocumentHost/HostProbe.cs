using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
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

        public static void DumpContentChars(string site, Word.Document document)
        {
            if (!Enabled || document == null)
            {
                return;
            }

            try
            {
                Word.Range content = document.Content;
                string text = content.Text ?? "";
                Log(site, "Content.Start=" + content.Start + " End=" + content.End
                    + " Text.Len=" + text.Length
                    + " hex=" + ToHex(text)
                    + " q=" + Quote(text));

                Word.Characters chars = content.Characters;
                int n = chars.Count;
                Log(site, "Characters.Count=" + n);
                int limit = n < 40 ? n : 40;
                for (int i = 1; i <= limit; i++)
                {
                    string ct = chars[i].Text ?? "";
                    Log(site, "ch[" + i + "] start=" + chars[i].Start
                        + " hex=" + ToHex(ct) + " q=" + Quote(ct));
                }
            }
            catch (Exception ex)
            {
                Log(site, Ex(ex));
            }
        }

        public static string Quote(string text)
        {
            if (text == null)
            {
                return "(null)";
            }

            var sb = new StringBuilder("\"");
            foreach (char ch in text)
            {
                if (ch == '\r')
                {
                    sb.Append("\\r");
                }
                else if (ch == '\n')
                {
                    sb.Append("\\n");
                }
                else if (ch == '\t')
                {
                    sb.Append("\\t");
                }
                else if (ch == '\a')
                {
                    sb.Append("\\a");
                }
                else if (ch == '\f')
                {
                    sb.Append("\\f");
                }
                else if (ch < 32)
                {
                    sb.Append("\\x").Append(((int)ch).ToString("X2"));
                }
                else
                {
                    sb.Append(ch);
                }
            }

            return sb.Append('"').ToString();
        }

        public static string ToHex(string text)
        {
            if (text == null)
            {
                return "(null)";
            }

            var sb = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                if (i > 0)
                {
                    sb.Append(' ');
                }

                sb.Append(((int)text[i]).ToString("X4"));
            }

            return sb.ToString();
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
