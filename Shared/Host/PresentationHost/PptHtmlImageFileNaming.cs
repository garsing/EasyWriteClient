using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace WordAddIn1.PresentationHost
{
    /// <summary>
    /// ppt_images 下文件命名：对导出 PNG 内容做 SHA256，取十六进制前 8 位（内容寻址，便于跨页共用）。
    /// </summary>
    internal static class PptHtmlImageFileNaming
    {
        public const int HashPrefixLength = 8;

        public static bool TryFinalizeExportedPng(
            string tempExportPath,
            string assetsFolderName,
            string assetsLocalDir,
            out string relativePath,
            out string error)
        {
            relativePath = null;
            error = null;
            if (string.IsNullOrWhiteSpace(tempExportPath) || !File.Exists(tempExportPath))
            {
                error = "临时导出文件不存在";
                return false;
            }

            string hash8;
            try
            {
                hash8 = ComputeSha256HexPrefix(tempExportPath, HashPrefixLength);
            }
            catch (Exception ex)
            {
                TryDelete(tempExportPath);
                error = "计算图片 hash 失败: " + ex.Message;
                return false;
            }

            string safeFile = hash8 + ".png";
            string finalPath = Path.Combine(assetsLocalDir, safeFile);
            try
            {
                Directory.CreateDirectory(assetsLocalDir);
                if (File.Exists(finalPath))
                {
                    // 同 hash → 视为同一内容，丢弃临时文件
                    TryDelete(tempExportPath);
                }
                else
                {
                    File.Move(tempExportPath, finalPath);
                }
            }
            catch (Exception ex)
            {
                TryDelete(tempExportPath);
                error = "写入 ppt_images 失败: " + ex.Message;
                return false;
            }

            string folder = (assetsFolderName ?? "").Trim().TrimEnd('/', '\\');
            string relative = folder + "/" + safeFile;
            relative = WorkspacePathResolver.SanitizeWorkspaceRelativePath(relative);
            if (relative == null)
            {
                error = "非法 assets 相对路径: " + folder + "/" + safeFile;
                return false;
            }

            relativePath = relative;
            return true;
        }

        public static string ComputeSha256HexPrefix(string filePath, int prefixLen)
        {
            using (var stream = File.OpenRead(filePath))
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(stream);
                var sb = new StringBuilder(prefixLen);
                for (int i = 0; i < hash.Length && sb.Length < prefixLen; i++)
                {
                    sb.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
                }

                if (sb.Length > prefixLen)
                {
                    return sb.ToString(0, prefixLen);
                }

                return sb.ToString();
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
