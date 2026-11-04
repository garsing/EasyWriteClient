using System;
using System.IO;

namespace WordAddIn1
{
    /// <summary>
    /// 抽图目录在工作区根 curr_docs 下，按文稿名加宿主后缀，避免 Word 清空时带走 PPT 的图。
    /// </summary>
    public static class CurrDocsImagePaths
    {
        public const string WordSuffix = "word";
        public const string PptSuffix = "ppt";

        public static bool TryEnsurePptDirectory(string presentationName, out string directory, out string error)
        {
            directory = null;
            error = null;
            string root = UserService.Instance == null ? null : UserService.Instance.WorkspaceRootEffective;
            if (string.IsNullOrWhiteSpace(root))
            {
                error = "工作区根不可用，无法导出图片";
                return false;
            }

            string stem = SanitizeStem(presentationName);
            directory = Path.Combine(root, "curr_docs", stem + "_" + PptSuffix);
            try
            {
                Directory.CreateDirectory(directory);
            }
            catch (Exception ex)
            {
                directory = null;
                error = "无法创建图片目录: " + ex.Message;
                return false;
            }

            return true;
        }

        public static string SanitizeStem(string name)
        {
            string raw = Path.GetFileNameWithoutExtension(name ?? "");
            if (string.IsNullOrWhiteSpace(raw))
            {
                raw = "文档";
            }

            foreach (char c in Path.GetInvalidFileNameChars())
            {
                raw = raw.Replace(c, '_');
            }

            raw = raw.Trim();
            return string.IsNullOrWhiteSpace(raw) ? "文档" : raw;
        }
    }
}
