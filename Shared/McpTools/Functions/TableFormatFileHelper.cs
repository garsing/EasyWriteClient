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
                string filePath = WorkspacePathResolver.ResolveWritePath(filename);

                if (!string.IsNullOrEmpty(xmlContent))
                {
                    XDocument.Parse(xmlContent);
                }

                File.WriteAllText(filePath, xmlContent);
                bool uploadSuccess = await McpToolsHelpers.UploadWorkspaceFileAsync(filePath, filename);
                System.Diagnostics.Debug.WriteLine(
                    $"[TableFormatFileHelper] 上传 {filename} 结果: {uploadSuccess}");
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
