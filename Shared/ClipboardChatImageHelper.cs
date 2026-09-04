using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

namespace WordAddIn1
{
    /// <summary>
    /// 对话粘贴：Win32 剪贴板读图（QQ 截图）或读资源管理器复制的文件。须在 STA / UI 线程调用。
    /// </summary>
    public static class ClipboardChatImageHelper
    {
        public const int MaxFiles = 10;

        public static ClipboardChatReadResult TryRead()
        {
            try
            {
                var dropped = TryReadFileDropList();
                if (dropped != null)
                {
                    return dropped;
                }

                var png = TryReadPngFormat();
                if (png != null)
                {
                    return png;
                }

                if (Clipboard.ContainsImage())
                {
                    using (Image img = Clipboard.GetImage())
                    {
                        if (img != null)
                        {
                            return EncodePng(img, SuggestName("png"));
                        }
                    }
                }

                return ClipboardChatReadResult.Fail();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ClipboardChatImageHelper] " + ex.Message);
                return ClipboardChatReadResult.Fail(ex.Message);
            }
        }

        private static ClipboardChatReadResult TryReadFileDropList()
        {
            if (!Clipboard.ContainsFileDropList())
            {
                return null;
            }

            var list = Clipboard.GetFileDropList();
            if (list == null || list.Count == 0)
            {
                return null;
            }

            var files = new List<ClipboardChatFileRef>();
            int existingCount = 0;
            foreach (string raw in list)
            {
                if (string.IsNullOrWhiteSpace(raw) || !File.Exists(raw))
                {
                    continue;
                }

                existingCount++;
                if (files.Count >= MaxFiles)
                {
                    continue;
                }

                long size = 0;
                try
                {
                    size = new FileInfo(raw).Length;
                }
                catch
                {
                    /* ignore */
                }

                files.Add(new ClipboardChatFileRef
                {
                    path = raw,
                    fileName = Path.GetFileName(raw),
                    fileSize = size
                });
            }

            if (files.Count == 0)
            {
                return null;
            }

            return new ClipboardChatReadResult
            {
                success = true,
                kind = "files",
                files = files,
                limitHit = existingCount > files.Count
            };
        }

        private static ClipboardChatReadResult TryReadPngFormat()
        {
            if (!Clipboard.ContainsData("PNG"))
            {
                return null;
            }

            object data = Clipboard.GetData("PNG");
            byte[] bytes = null;
            if (data is MemoryStream ms)
            {
                bytes = ms.ToArray();
            }
            else if (data is byte[] raw)
            {
                bytes = raw;
            }

            if (bytes == null || bytes.Length == 0)
            {
                return null;
            }

            return new ClipboardChatReadResult
            {
                success = true,
                kind = "image",
                fileName = SuggestName("png"),
                mimeType = "image/png",
                base64 = Convert.ToBase64String(bytes)
            };
        }

        private static ClipboardChatReadResult EncodePng(Image img, string fileName)
        {
            using (var ms = new MemoryStream())
            {
                img.Save(ms, ImageFormat.Png);
                return new ClipboardChatReadResult
                {
                    success = true,
                    kind = "image",
                    fileName = fileName,
                    mimeType = "image/png",
                    base64 = Convert.ToBase64String(ms.ToArray())
                };
            }
        }

        private static string SuggestName(string ext)
        {
            return "截图-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + "." + ext;
        }
    }

    public sealed class ClipboardChatReadResult
    {
        public bool success { get; set; }
        public string kind { get; set; }
        public string fileName { get; set; }
        public string mimeType { get; set; }
        public string base64 { get; set; }
        public string message { get; set; }
        public bool limitHit { get; set; }
        public List<ClipboardChatFileRef> files { get; set; }

        public static ClipboardChatReadResult Fail(string message = null)
        {
            return new ClipboardChatReadResult { success = false, message = message };
        }
    }

    public sealed class ClipboardChatFileRef
    {
        public string path { get; set; }
        public string fileName { get; set; }
        public long fileSize { get; set; }
    }
}
