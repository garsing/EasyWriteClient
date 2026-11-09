using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// 导出回包里已落盘的图，压到长边 320 后打 desc。失败不改抽取成功。
    /// </summary>
    internal static class ExportImageCaption
    {
        internal const int BatchSize = 20;
        internal const int MaxCaptioned = 100;
        internal const int MaxDescChars = 20;
        internal const string OverflowDesc = "图片太多，没有打标内容";
        internal const string SuccessMessage = "抽取成功";

        /// <summary>测试注入：入参是这一批的文件路径，返回模型正文。生产为 null。</summary>
        internal static Func<IList<string>, string> RequestTextOverride;

        private static readonly Regex ImageTagRegex = new Regex(
            @"<image\b[^>]*?/?>",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex SrcRegex = new Regex(
            @"\bsrc\s*=\s*""([^""]*)""",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex LineRegex = new Regex(
            @"^\s*(\d+)\s*[\.．、]\s*(.*?)\s*$",
            RegexOptions.CultureInvariant);

        internal static void ApplyDocumentContent(Dictionary<string, object> data)
        {
            if (!IsSuccess(data) || data == null || !data.ContainsKey("document_content"))
            {
                return;
            }

            data["document_content"] = ApplyHtml(data["document_content"] as string);
        }

        internal static void ApplyDisplayContents(Dictionary<string, object> data)
        {
            if (!IsSuccess(data) || data == null || !data.ContainsKey("display_contents"))
            {
                return;
            }

            data["display_contents"] = ApplyHtml(data["display_contents"] as string);
        }

        internal static void ApplyExcelExport(Dictionary<string, object> data)
        {
            if (!IsSuccess(data) || data == null)
            {
                return;
            }

            string action = data.ContainsKey("action") ? data["action"] as string : null;
            if (!string.Equals(action, "export", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (!(data["images"] is List<Dictionary<string, object>> images))
            {
                return;
            }

            var slots = new List<Dictionary<string, object>>();
            var paths = new List<string>();
            foreach (Dictionary<string, object> row in images)
            {
                string path = row != null && row.ContainsKey("path") ? row["path"] as string : null;
                if (!IsLandedFile(path))
                {
                    continue;
                }

                slots.Add(row);
                paths.Add(path);
            }

            IList<string> descs = CaptionPaths(paths);
            for (int i = 0; i < slots.Count; i++)
            {
                if (!string.IsNullOrEmpty(descs[i]))
                {
                    slots[i]["desc"] = descs[i];
                }
            }
        }

        internal static string ApplyHtml(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return html;
            }

            MatchCollection matches = ImageTagRegex.Matches(html);
            var landed = new List<Match>();
            var paths = new List<string>();
            foreach (Match match in matches)
            {
                string path = SrcPath(match.Value);
                if (!IsLandedFile(path))
                {
                    continue;
                }

                landed.Add(match);
                paths.Add(path);
            }

            if (landed.Count == 0)
            {
                return html;
            }

            IList<string> descs = CaptionPaths(paths);
            var builder = new StringBuilder();
            int cursor = 0;
            for (int i = 0; i < landed.Count; i++)
            {
                Match match = landed[i];
                builder.Append(html, cursor, match.Index - cursor);
                builder.Append(WithDesc(match.Value, descs[i]));
                cursor = match.Index + match.Length;
            }

            builder.Append(html, cursor, html.Length - cursor);
            return builder.ToString();
        }

        internal static string TruncateDesc(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string trimmed = text.Trim();
            if (trimmed.Length <= MaxDescChars)
            {
                return trimmed;
            }

            return trimmed.Substring(0, MaxDescChars);
        }

        internal static bool TryParseBatch(string text, int count, out string[] descs)
        {
            descs = new string[count];
            if (count < 1 || string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            int matched = 0;
            string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            foreach (string line in lines)
            {
                Match match = LineRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                if (!int.TryParse(match.Groups[1].Value, out int index))
                {
                    continue;
                }

                if (index < 1 || index > count || descs[index - 1] != null)
                {
                    continue;
                }

                descs[index - 1] = TruncateDesc(match.Groups[2].Value);
                matched++;
            }

            return matched > 0;
        }

        private static IList<string> CaptionPaths(IList<string> paths)
        {
            var descs = new string[paths.Count];
            int cap = Math.Min(paths.Count, MaxCaptioned);
            for (int i = MaxCaptioned; i < paths.Count; i++)
            {
                descs[i] = OverflowDesc;
            }

            if (cap == 0)
            {
                return descs;
            }

            var model = new string[cap];
            try
            {
                for (int start = 0; start < cap; start += BatchSize)
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    int length = Math.Min(BatchSize, cap - start);
                    var batch = new List<string>(length);
                    for (int i = 0; i < length; i++)
                    {
                        batch.Add(paths[start + i]);
                    }

                    string text = RequestBatch(batch);
                    if (!TryParseBatch(text, length, out string[] parsed))
                    {
                        return descs;
                    }

                    for (int i = 0; i < length; i++)
                    {
                        model[start + i] = parsed[i];
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception)
            {
                return descs;
            }

            for (int i = 0; i < cap; i++)
            {
                descs[i] = string.IsNullOrEmpty(model[i]) ? null : model[i];
            }

            return descs;
        }

        private static string RequestBatch(IList<string> paths)
        {
            if (RequestTextOverride != null)
            {
                return RequestTextOverride(paths);
            }

            var images = new List<string>(paths.Count);
            foreach (string path in paths)
            {
                using (var bitmap = new Bitmap(path))
                {
                    ImageCaptureCompressor.CompressedImage thumb = ImageCaptureCompressor.CompressThumb(bitmap, 320);
                    images.Add(Convert.ToBase64String(thumb.Bytes));
                }
            }

            if (!UserService.Instance.IsLoggedIn || string.IsNullOrEmpty(UserService.Instance.AccessToken))
            {
                throw new InvalidOperationException("未登录");
            }

            string baseUrl = (ConfigManager.Config.Api.BaseUrl ?? "").TrimEnd('/');
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(new Dictionary<string, object>
            {
                ["images"] = images
            });
            BackendApiClient.JsonResult result = BackendApiClient.PostJsonAuthenticatedAsync(
                baseUrl + "/internal/export-image-caption",
                json,
                retryOnUnauthorized: true,
                timeout: TimeSpan.FromSeconds(180)).GetAwaiter().GetResult();
            if (result == null || !result.Success || string.IsNullOrEmpty(result.Body))
            {
                throw new InvalidOperationException("调用失败");
            }

            JObject payload = JObject.Parse(result.Body);
            string text = payload.Value<string>("text");
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new InvalidOperationException("调用失败");
            }

            return text;
        }

        private static bool IsSuccess(Dictionary<string, object> data)
        {
            if (data == null || !data.ContainsKey("message"))
            {
                return false;
            }

            return string.Equals(data["message"] as string, SuccessMessage, StringComparison.Ordinal);
        }

        private static bool IsLandedFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            if (path.IndexOf("lightweight/placeholder_", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            try
            {
                return File.Exists(path);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string SrcPath(string tag)
        {
            Match match = SrcRegex.Match(tag ?? "");
            if (!match.Success)
            {
                return "";
            }

            return Unescape(match.Groups[1].Value);
        }

        private static string WithDesc(string tag, string desc)
        {
            if (string.IsNullOrEmpty(desc) || string.IsNullOrEmpty(tag))
            {
                return tag;
            }

            string attribute = " desc=\"" + Escape(desc) + "\"";
            if (tag.EndsWith("/>", StringComparison.Ordinal))
            {
                return tag.Substring(0, tag.Length - 2) + attribute + "/>";
            }

            if (tag.EndsWith(">", StringComparison.Ordinal))
            {
                return tag.Substring(0, tag.Length - 1) + attribute + ">";
            }

            return tag;
        }

        private static string Escape(string value)
        {
            return (value ?? "")
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;");
        }

        private static string Unescape(string value)
        {
            return (value ?? "")
                .Replace("&quot;", "\"")
                .Replace("&lt;", "<")
                .Replace("&gt;", ">")
                .Replace("&amp;", "&");
        }
    }
}
