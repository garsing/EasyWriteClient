using System;
using System.IO;

namespace WordAddIn1
{
    internal enum OpenDocumentFamily
    {
        Document = 0,
        Spreadsheet = 1
    }

    internal static class OpenDocumentHostResolver
    {
        public static bool TryGetFamily(string fullPath, out OpenDocumentFamily family, out string error)
        {
            family = OpenDocumentFamily.Document;
            error = null;
            string ext = Path.GetExtension(fullPath)?.ToLowerInvariant();
            switch (ext)
            {
                case ".doc":
                case ".docx":
                case ".docm":
                case ".wps":
                    family = OpenDocumentFamily.Document;
                    return true;
                case ".xls":
                case ".xlsx":
                case ".xlsm":
                case ".et":
                    family = OpenDocumentFamily.Spreadsheet;
                    return true;
                case ".ppt":
                case ".pptx":
                    error = "unsupported: 尚未支持演示文稿";
                    return false;
                default:
                    error = string.IsNullOrEmpty(ext)
                        ? "路径必须带可识别后缀（.docx / .xlsx / .et 等）"
                        : "不支持的后缀: " + ext;
                    return false;
            }
        }

        public static bool TryValidateAppForFamily(OpenDocumentFamily family, string appRaw, out string error)
        {
            error = null;
            if (string.IsNullOrWhiteSpace(appRaw))
            {
                return true;
            }

            string app = appRaw.Trim().ToLowerInvariant();
            bool documentApp = app == "word" || app == "wps";
            bool spreadsheetApp = app == "excel" || app == "et";
            if (!documentApp && !spreadsheetApp)
            {
                error = "app 须为 word / wps / excel / et";
                return false;
            }

            if (family == OpenDocumentFamily.Document && spreadsheetApp)
            {
                error = "文件是文字文档，不能用 app=" + app + " 打开";
                return false;
            }

            if (family == OpenDocumentFamily.Spreadsheet && documentApp)
            {
                error = "文件是表格，不能用 app=" + app + " 打开";
                return false;
            }

            return true;
        }
    }
}
