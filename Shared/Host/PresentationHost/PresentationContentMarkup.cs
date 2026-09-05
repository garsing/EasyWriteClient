using System;
using System.Collections.Generic;
using System.Text;

namespace WordAddIn1.PresentationHost
{
    internal static class PresentationContentMarkup
    {
        public static string BuildDisplayContents(PresentationContentResult result)
        {
            if (result == null)
            {
                return "";
            }

            var sb = new StringBuilder();
            sb.AppendLine("演示文稿：" + (result.Name ?? ""));
            sb.AppendLine("路径：" + (string.IsNullOrEmpty(result.Path) ? "（未保存）" : result.Path));
            sb.AppendLine(
                "kind=" + (result.Kind ?? "")
                + " channel_id=" + (result.ChannelId ?? ""));
            sb.AppendLine("页数：" + result.SlideCount);
            if (result.Truncated)
            {
                sb.AppendLine(
                    "truncated=true"
                    + (string.IsNullOrEmpty(result.TruncatedReason)
                        ? ""
                        : " reason=" + result.TruncatedReason));
            }

            sb.AppendLine();
            if (result.Slides == null || result.Slides.Count == 0)
            {
                sb.AppendLine("（无幻灯片）");
                return sb.ToString().TrimEnd();
            }

            sb.AppendLine("| index | slide_id |");
            sb.AppendLine("|------:|---------:|");
            foreach (PresentationSlideInfo slide in result.Slides)
            {
                if (slide == null)
                {
                    continue;
                }

                sb.Append("| ")
                    .Append(slide.Index)
                    .Append(" | ")
                    .Append(EscapeCell(slide.SlideId))
                    .AppendLine(" |");
            }

            sb.AppendLine();
            sb.AppendLine("说明：机器主键用 slide_id（SlideID），不要用「第 N 页」当主键。");
            return sb.ToString().TrimEnd();
        }

        private static string EscapeCell(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text
                .Replace("\r\n", " ")
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('|', '/')
                .Trim();
        }
    }
}
