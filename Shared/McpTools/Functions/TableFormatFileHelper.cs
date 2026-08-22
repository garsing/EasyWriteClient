using System;
using System.IO;
using System.Threading.Tasks;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class TableFormatFileHelper
    {
        public static string GetFormatFilesDirectory()
        {
            return McpToolsHelpers.GetWorkspaceSessionDirectory();
        }

        public static string GenerateFilename(Word.Document document, int tableIndex, string suffix)
        {
            try
            {
                string docName = Path.GetFileNameWithoutExtension(document.Name);
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                return $"{docName}_{tableIndex}_{suffix}_{timestamp}.xml";
            }
            catch
            {
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                return $"table_{tableIndex}_{suffix}_{timestamp}.xml";
            }
        }

        public static async Task<(bool Success, string XmlContent)> SaveXmlAsync(string filename, string xmlContent)
        {
            try
            {
                if (!FilePathResolver.TryResolve(filename, out ResolvedFilePath resolved, out string pathError))
                {
                    System.Diagnostics.Debug.WriteLine($"[TableFormatFileHelper] 路径失败：{pathError}");
                    return (false, xmlContent);
                }

                if (!string.IsNullOrEmpty(xmlContent))
                {
                    XDocument.Parse(xmlContent);
                }

                var written = await FilePathResolver.WriteAsync(resolved, xmlContent).ConfigureAwait(false);
                if (!written.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"[TableFormatFileHelper] 写入失败：{written.Error}");
                    return (false, xmlContent);
                }

                return (true, xmlContent);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableFormatFileHelper] 保存失败：{ex.Message}");
                return (false, null);
            }
        }
    }
}
