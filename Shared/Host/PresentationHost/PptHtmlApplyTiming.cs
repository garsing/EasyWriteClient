using System;
using System.Collections.Generic;
using System.Text;

namespace WordAddIn1.PresentationHost
{
    /// <summary>F_apply_ppt_html 分步耗时（timing / ppt_html 均可看到）。</summary>
    internal static class PptHtmlApplyTiming
    {
        public static void Step(string name, long elapsedMs, string detail = null)
        {
            EasyWriteDiagnostics.LogTiming("apply_ppt_html." + name, elapsedMs, detail);
            if (string.IsNullOrEmpty(detail))
            {
                EasyWriteDiagnostics.Log(DebugCategory.PptHtml, "timing " + name + " elapsed_ms=" + elapsedMs);
            }
            else
            {
                EasyWriteDiagnostics.Log(
                    DebugCategory.PptHtml,
                    "timing " + name + " elapsed_ms=" + elapsedMs + " " + detail);
            }
        }

        public static void NoteTopNodes(List<KeyValuePair<long, string>> timings)
        {
            if (timings == null || timings.Count == 0)
            {
                return;
            }

            timings.Sort((a, b) => b.Key.CompareTo(a.Key));
            int n = Math.Min(8, timings.Count);
            var sb = new StringBuilder();
            sb.Append("count=").Append(timings.Count);
            for (int i = 0; i < n; i++)
            {
                sb.Append(" | ").Append(timings[i].Value).Append('=').Append(timings[i].Key).Append("ms");
            }

            string line = sb.ToString();
            EasyWriteDiagnostics.Log(DebugCategory.Timing, "step=apply_ppt_html.nodes_top " + line);
            EasyWriteDiagnostics.Log(DebugCategory.PptHtml, "timing nodes_top " + line);
        }
    }
}
