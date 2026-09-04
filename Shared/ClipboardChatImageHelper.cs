using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

namespace WordAddIn1
{
    /// <summary>
    /// 对话粘贴：从 Win32 剪贴板读图（QQ 截图等）。须在 STA / UI 线程调用。
    /// </summary>
    public static class ClipboardChatImageHelper
    {
        public static object TryRead()
        {
            try
            {
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

                return new { success = false };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ClipboardChatImageHelper] " + ex.Message);
                return new { success = false, message = ex.Message };
            }
        }

        private static object TryReadPngFormat()
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

            return new
            {
                success = true,
                fileName = SuggestName("png"),
                mimeType = "image/png",
                base64 = Convert.ToBase64String(bytes)
            };
        }

        private static object EncodePng(Image img, string fileName)
        {
            using (var ms = new MemoryStream())
            {
                img.Save(ms, ImageFormat.Png);
                return new
                {
                    success = true,
                    fileName,
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
}
