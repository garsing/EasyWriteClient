using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using WordAddIn1;

namespace WordAddIn1.PresentationHost
{
    /// <summary>F_read_ppt_html 字色诊断（定位主题色误读/省略）。</summary>
    internal sealed class PptHtmlReadDebug
    {
        private readonly List<string> _lines = new List<string>();

        public IReadOnlyList<string> Lines => _lines;

        public void Line(string message)
        {
            string ts = DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture);
            _lines.Add("[" + ts + "] " + (message ?? ""));
        }

        public void Probe(string message)
        {
            Line(message);
            EasyWriteDiagnostics.Log(DebugCategory.PptHtml, message);
        }

        public string BuildText(string header)
        {
            var sb = new StringBuilder();
            if (!string.IsNullOrEmpty(header))
            {
                sb.AppendLine(header);
                sb.AppendLine(new string('-', 72));
            }

            foreach (string line in _lines)
            {
                sb.AppendLine(line);
            }

            return sb.ToString();
        }

        /// <summary>写入会话目录；返回相对工作区路径。</summary>
        public string TryWriteToSession(string slideId, out string error)
        {
            error = null;
            try
            {
                string dir = WorkspacePathResolver.GetSessionDirectory();
                Directory.CreateDirectory(dir);
                string name = "read_font_color_slide_" + (slideId ?? "x") + ".txt";
                string full = Path.Combine(dir, name);
                File.WriteAllText(
                    full,
                    BuildText("F_read_ppt_html font-color probe slide_id=" + slideId),
                    Encoding.UTF8);
                return name;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return null;
            }
        }
    }
}
