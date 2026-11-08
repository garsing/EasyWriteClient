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

            foreach (PresentationSlideInfo slide in result.Slides)
            {
                if (slide == null)
                {
                    continue;
                }

                sb.Append("index=").Append(slide.Index);
                sb.Append(" slide_id=").Append(slide.SlideId ?? "");
                sb.AppendLine();
                string body = string.IsNullOrWhiteSpace(slide.PageText)
                    ? "（无文字）"
                    : slide.PageText.Trim();
                sb.AppendLine(body);
                sb.AppendLine();
            }

            sb.Append("说明：机器主键用 slide_id（SlideID），不要用「第 N 页」当主键。");
            sb.Append("要改某一页或取原图，用该 slide_id 调 F_read_ppt_html。");
            return sb.ToString().TrimEnd();
        }
    }
}
