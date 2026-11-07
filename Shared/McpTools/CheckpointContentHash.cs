using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    public static class CheckpointContentHash
    {
        private static readonly Regex ResaveAttribute = new Regex(
            "\\s(?:[\\w.-]+:)?(?:rsid\\w*|paraId|textId)=\"[^\"]*\"",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static string FromFile(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                return "";
            }

            try
            {
                string ext = Path.GetExtension(path);
                var names = new List<string>();
                using (ZipArchive zip = ZipFile.OpenRead(path))
                {
                    foreach (ZipArchiveEntry entry in zip.Entries)
                    {
                        if (entry == null || string.IsNullOrEmpty(entry.FullName))
                        {
                            continue;
                        }

                        if (Include(ext, entry.FullName))
                        {
                            names.Add(entry.FullName);
                        }
                    }

                    names.Sort(StringComparer.Ordinal);
                    if (names.Count == 0)
                    {
                        return "";
                    }

                    using (SHA256 sha = SHA256.Create())
                    {
                        foreach (string name in names)
                        {
                            ZipArchiveEntry entry = zip.GetEntry(name);
                            if (entry == null)
                            {
                                continue;
                            }

                            string xml;
                            using (Stream stream = entry.Open())
                            using (var reader = new StreamReader(stream, Encoding.UTF8))
                            {
                                xml = reader.ReadToEnd();
                            }

                            xml = ResaveAttribute.Replace(xml, "");
                            byte[] nameBytes = Encoding.UTF8.GetBytes(name);
                            sha.TransformBlock(nameBytes, 0, nameBytes.Length, null, 0);
                            byte[] body = Encoding.UTF8.GetBytes(xml);
                            sha.TransformBlock(body, 0, body.Length, null, 0);
                        }

                        sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                        return ToHex(sha.Hash);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[Checkpoint] content hash failed: " + ex.Message);
                return "";
            }
        }

        private static bool Include(string extension, string name)
        {
            string path = (name ?? "").Replace('\\', '/').ToLowerInvariant();
            string ext = (extension ?? "").TrimStart('.').ToLowerInvariant();
            if (ext == "xlsx" || ext == "xlsm")
            {
                return path == "xl/sharedstrings.xml" || path.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal);
            }

            if (ext == "pptx" || ext == "pptm")
            {
                return path.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
                    || path.StartsWith("ppt/notesslides/notesslide", StringComparison.Ordinal);
            }

            return path == "word/document.xml"
                || path == "word/footnotes.xml"
                || path == "word/endnotes.xml"
                || path == "word/comments.xml"
                || path.StartsWith("word/header", StringComparison.Ordinal)
                || path.StartsWith("word/footer", StringComparison.Ordinal);
        }

        private static string ToHex(byte[] hash)
        {
            if (hash == null || hash.Length == 0)
            {
                return "";
            }

            var text = new StringBuilder(hash.Length * 2);
            foreach (byte b in hash)
            {
                text.Append(b.ToString("x2"));
            }

            return text.ToString();
        }
    }
}
