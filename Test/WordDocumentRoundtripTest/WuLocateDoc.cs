using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 真稿：F_get_document_content 后，对每个可操作码走 Locator（get_format_context）。
    /// 句级：D22 跳过；真空格 empty_cell_no_format 算已定位。
    /// 段级（--level=paragraph）：paragraph display 的每个 P_ 走 target_paragraph_codes。
    /// </summary>
    internal static class WuLocateDoc
    {
        private static readonly Regex D22Start = new Regex(
            @"\[(S_[0-9A-Za-z]+)-start 歧义无法操作\]",
            RegexOptions.Compiled);

        private static readonly Regex OperableStart = new Regex(
            @"\[((?:S_[0-9A-Za-z]+)|(?:T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+))-start\]",
            RegexOptions.Compiled);

        private static readonly Regex ParagraphStart = new Regex(
            @"\[((P_[0-9a-zA-Z]{5}(?:_[1-9][0-9]*)?))(?:-start| start)\]",
            RegexOptions.Compiled);

        public static int Run(string host, string docPath, bool paragraph)
        {
            Console.WriteLine(paragraph
                ? "真稿逐码定位 P_（paragraph display → get_format_context）"
                : "真稿逐码定位（跳过 D22）");
            Console.WriteLine(paragraph
                ? "读 F_get_document_content code_level=paragraph，定位 target_paragraph_codes"
                : "读 F_get_document_content，定位 F_get_format_context");
            Console.WriteLine(host == "wps" ? "宿主 WPS 文字" : "宿主 Word");
            if (string.IsNullOrWhiteSpace(docPath) || !File.Exists(docPath))
            {
                Console.WriteLine("稿不存在: " + docPath);
                return 1;
            }

            if (!DesktopLauncher.EnsureAvailable())
            {
                Console.WriteLine("连不上易写 Desktop。");
                return 1;
            }

            var s = new WdSession(host);
            ToolResult opened = s.Open(docPath, false, s.App);
            if (!opened.Success)
            {
                Console.WriteLine("F_open_document 失败: " + (opened.Error ?? ""));
                return 1;
            }

            string ch = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(docPath);
            Console.WriteLine("channel_id=" + ch);
            Console.WriteLine("稿=" + docPath);

            return paragraph ? LocateParagraphs(s, ch) : LocateSentences(s, ch);
        }

        private static int LocateSentences(WdSession s, string ch)
        {
            ToolResult got = s.GetContent(ch);
            if (!got.Success)
            {
                Console.WriteLine("F_get_document_content 失败: " + (got.Error ?? ""));
                return 1;
            }

            string body = WdSession.Str(got, "document_content");
            Console.WriteLine("display 字数=" + body.Length);

            var skip = new HashSet<string>(StringComparer.Ordinal);
            foreach (Match m in D22Start.Matches(body))
            {
                skip.Add(m.Groups[1].Value);
            }

            var codes = UniqueCodes(OperableStart, body, skip);
            Console.WriteLine("D22 跳过 " + skip.Count + "  可定位 " + codes.Count);
            if (skip.Count > 0)
            {
                Console.WriteLine("D22: " + string.Join(", ", skip));
            }

            return LocateEach(s, ch, codes, sentence: true, skip.Count);
        }

        private static int LocateParagraphs(WdSession s, string ch)
        {
            ToolResult got = s.GetContent(ch, "paragraph");
            if (!got.Success)
            {
                Console.WriteLine("F_get_document_content(paragraph) 失败: " + (got.Error ?? ""));
                return 1;
            }

            string body = WdSession.Str(got, "document_content");
            Console.WriteLine("paragraph display 字数=" + body.Length);

            var codes = UniqueCodes(ParagraphStart, body, null);
            Console.WriteLine("可定位 P_ " + codes.Count);
            return LocateEach(s, ch, codes, sentence: false, skipD22: 0);
        }

        private static List<string> UniqueCodes(Regex start, string body, HashSet<string> skip)
        {
            var codes = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (Match m in start.Matches(body ?? ""))
            {
                string code = m.Groups[1].Value;
                if ((skip != null && skip.Contains(code)) || !seen.Add(code))
                {
                    continue;
                }

                codes.Add(code);
            }

            return codes;
        }

        private static int LocateEach(
            WdSession s,
            string ch,
            List<string> codes,
            bool sentence,
            int skipD22)
        {
            int ok = 0;
            int vacuum = 0;
            var fails = new List<string>();
            for (int i = 0; i < codes.Count; i++)
            {
                string code = codes[i];
                var args = new Dictionary<string, object> { ["channel_id"] = ch };
                if (sentence)
                {
                    args["target_codes"] = code;
                }
                else
                {
                    args["target_paragraph_codes"] = code;
                }

                ToolResult loc = s.GetFormatContext(args);
                string err = loc == null ? "null" : (loc.Error ?? "");
                if (loc != null && loc.Success)
                {
                    ok++;
                    continue;
                }

                if (sentence && err.IndexOf("empty_cell_no_format", StringComparison.Ordinal) >= 0)
                {
                    vacuum++;
                    continue;
                }

                fails.Add(code + " => " + TrimErr(err));
                Console.WriteLine("[FAIL] " + (i + 1) + "/" + codes.Count + "  " + code);
                Console.WriteLine("       " + TrimErr(err));
            }

            Console.WriteLine(
                "定位成功 " + ok
                + (sentence ? "  真空格(已定位) " + vacuum : "")
                + "  失败 " + fails.Count
                + (sentence ? "  跳过 D22 " + skipD22 : "")
                + "  共 " + codes.Count);
            return fails.Count == 0 ? 0 : 1;
        }

        private static string TrimErr(string err)
        {
            if (string.IsNullOrEmpty(err))
            {
                return "(无 Error)";
            }

            return err.Length <= 400 ? err : err.Substring(0, 400);
        }
    }
}
