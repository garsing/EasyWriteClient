using System;
using System.IO;

namespace WordAddIn1
{
    internal enum OpenDocumentFamily
    {
        Document = 0,
        Spreadsheet = 1,
        Presentation = 2
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
                case ".pptm":
                    family = OpenDocumentFamily.Presentation;
                    return true;
                default:
                    error = string.IsNullOrEmpty(ext)
                        ? "路径必须带可识别后缀（.docx / .xlsx / .pptx / .et 等）"
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
            bool presentationApp = app == "powerpoint" || app == "wpp";
            if (!documentApp && !spreadsheetApp && !presentationApp)
            {
                error = "app 须为 word / wps / excel / et / powerpoint / wpp";
                return false;
            }

            if (family == OpenDocumentFamily.Document && !documentApp)
            {
                error = "文件是文字文档，不能用 app=" + app + " 打开";
                return false;
            }

            if (family == OpenDocumentFamily.Spreadsheet && !spreadsheetApp)
            {
                error = "文件是表格，不能用 app=" + app + " 打开";
                return false;
            }

            if (family == OpenDocumentFamily.Presentation && !presentationApp)
            {
                if (app == "wps")
                {
                    error = "文件是演示文稿，请用 app=wpp（WPS 演示），不能用 app=wps（WPS 文字）";
                    return false;
                }

                error = "文件是演示文稿，不能用 app=" + app + " 打开";
                return false;
            }

            return true;
        }
    }
}
