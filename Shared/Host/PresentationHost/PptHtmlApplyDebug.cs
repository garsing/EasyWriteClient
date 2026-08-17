using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using WordAddIn1;

namespace WordAddIn1.PresentationHost
{
    /// <summary>F_apply_ppt_html 诊断轨迹（定位「形状建了又没了」等问题）。</summary>
    internal sealed class PptHtmlApplyDebug
    {
        private readonly List<string> _lines = new List<string>();

        public IReadOnlyList<string> Lines => _lines;

        public void Line(string message)
        {
            string ts = DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture);
            _lines.Add("[" + ts + "] " + (message ?? ""));
        }

        public void Step(
            string action,
            PptHtmlApplyNode node,
            string detail = null)
        {
            if (node == null)
            {
                Line(action + " (node=null)");
                return;
            }

            var sb = new StringBuilder();
            sb.Append(action);
            sb.Append(" id=").Append(node.ShapeId ?? "");
            sb.Append(" com=").Append(node.ShapeComId.HasValue
                ? node.ShapeComId.Value.ToString(CultureInfo.InvariantCulture)
                : "-");
            sb.Append(" type=").Append(node.ShapeType ?? "");
            if (node.HasGeometry)
            {
                sb.Append(" geo=")
                    .Append(Fmt(node.LeftPct)).Append(',')
                    .Append(Fmt(node.TopPct)).Append(',')
                    .Append(Fmt(node.WidthPct)).Append(',')
                    .Append(Fmt(node.HeightPct));
            }

            if (node.Z.HasValue)
            {
                sb.Append(" z=").Append(node.Z.Value.ToString(CultureInfo.InvariantCulture));
            }

            if (!string.IsNullOrEmpty(node.Fill))
            {
                sb.Append(" fill=").Append(node.Fill);
            }

            if (!string.IsNullOrEmpty(node.FontColor))
            {
                sb.Append(" font=").Append(node.FontColor);
            }

            if (!string.IsNullOrEmpty(detail))
            {
                sb.Append(" | ").Append(detail);
            }

            Line(sb.ToString());
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
                string name = "apply_debug_slide_" + (slideId ?? "x") + ".txt";
                string full = Path.Combine(dir, name);
                File.WriteAllText(full, BuildText("F_apply_ppt_html debug slide_id=" + slideId), Encoding.UTF8);
                return name;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return null;
            }
        }

        private static string Fmt(double? v)
        {
            if (!v.HasValue)
            {
                return "?";
            }

            return v.Value.ToString("0.##", CultureInfo.InvariantCulture);
        }
    }
}
