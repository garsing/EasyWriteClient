using System;
using System.Drawing;
using System.IO;
using PdfiumViewer;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal static class PageCaptureHelper
    {
        private const int RenderDpi = 175;

        internal sealed class PageOutOfRangeException : Exception
        {
            public int RequestedPage { get; }
            public int TotalPages { get; }

            public PageOutOfRangeException(int requestedPage, int totalPages)
                : base($"页码 {requestedPage} 超出范围，文档共 {totalPages} 页")
            {
                RequestedPage = requestedPage;
                TotalPages = totalPages;
            }
        }

        internal sealed class CaptureResult
        {
            public int PageNumber { get; set; }
            public int TotalPages { get; set; }
            public string Format { get; set; }
            public int WidthPx { get; set; }
            public int HeightPx { get; set; }
            public string ImageBase64 { get; set; }
            public string CapturedAt { get; set; }
            public string DocTitle { get; set; }
        }

        public static CaptureResult CaptureDocumentPage(
            Word.Application wordApp,
            int? requestedPageNumber = null,
            Word.Document document = null)
        {
            PdfiumNativeLoader.EnsureLoaded();

            Word.Document doc = document;
            if (doc == null)
            {
                if (wordApp == null)
                {
                    throw new InvalidOperationException("Word应用程序不可用");
                }

                doc = wordApp.ActiveDocument;
            }

            if (doc == null)
            {
                throw new InvalidOperationException("没有活动的 Word 文档");
            }

            if (wordApp == null)
            {
                try
                {
                    wordApp = doc.Application;
                }
                catch (Exception)
                {
                    wordApp = null;
                }
            }

            if (wordApp == null)
            {
                throw new InvalidOperationException("Word应用程序不可用");
            }

            int totalPages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
            if (totalPages < 1)
            {
                totalPages = 1;
            }

            int pageNumber;
            if (requestedPageNumber.HasValue)
            {
                pageNumber = requestedPageNumber.Value;
                if (pageNumber < 1)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(requestedPageNumber),
                        "page_number 必须是 ≥1 的整数");
                }

                if (pageNumber > totalPages)
                {
                    throw new PageOutOfRangeException(pageNumber, totalPages);
                }
            }
            else
            {
                pageNumber = GetCursorPageNumber(wordApp, doc);
                if (pageNumber < 1)
                {
                    pageNumber = 1;
                }

                if (pageNumber > totalPages)
                {
                    pageNumber = totalPages;
                }
            }
            string tempDir = Path.Combine(Path.GetTempPath(), "EasyWrite", "capture", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            string pdfPath = Path.Combine(tempDir, "page.pdf");

            try
            {
                int pageIndex = ExportPageToPdf(wordApp, doc, pdfPath, pageNumber);
                using (var pdfDocument = PdfDocument.Load(pdfPath))
                {
                    if (pdfDocument.PageCount <= 0)
                    {
                        throw new InvalidOperationException("PDF 导出结果为空");
                    }

                    if (pageIndex < 0 || pageIndex >= pdfDocument.PageCount)
                    {
                        pageIndex = Math.Max(0, Math.Min(pageNumber - 1, pdfDocument.PageCount - 1));
                    }

                    using (var bitmap = (Bitmap)pdfDocument.Render(pageIndex, RenderDpi, RenderDpi, PdfRenderFlags.Annotations))
                    {
                        var compressed = ImageCaptureCompressor.Compress(bitmap);
                        return new CaptureResult
                        {
                            PageNumber = pageNumber,
                            TotalPages = totalPages,
                            Format = compressed.Format,
                            WidthPx = compressed.Width,
                            HeightPx = compressed.Height,
                            ImageBase64 = Convert.ToBase64String(compressed.Bytes),
                            CapturedAt = DateTimeOffset.Now.ToString("yyyy-MM-ddTHH:mm:sszzz"),
                            DocTitle = doc.Name ?? string.Empty
                        };
                    }
                }
            }
            finally
            {
                TryDeleteFile(pdfPath);
                TryDeleteDirectory(tempDir);
            }
        }

        internal sealed class DocumentPageInfo
        {
            public int TotalPages { get; set; }
            public int CursorPageNumber { get; set; }
            public string DocTitle { get; set; }
        }

        public static DocumentPageInfo GetDocumentPageInfo(
            Word.Application wordApp,
            Word.Document document = null)
        {
            Word.Document doc = document;
            if (doc == null)
            {
                if (wordApp == null)
                {
                    throw new InvalidOperationException("Word应用程序不可用");
                }

                doc = wordApp.ActiveDocument;
            }

            if (doc == null)
            {
                throw new InvalidOperationException("没有活动的 Word 文档");
            }

            if (wordApp == null)
            {
                try
                {
                    wordApp = doc.Application;
                }
                catch (Exception)
                {
                    wordApp = null;
                }
            }

            int totalPages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
            if (totalPages < 1)
            {
                totalPages = 1;
            }

            int cursorPage = GetCursorPageNumber(wordApp, doc);
            if (cursorPage < 1)
            {
                cursorPage = 1;
            }

            if (cursorPage > totalPages)
            {
                cursorPage = totalPages;
            }

            return new DocumentPageInfo
            {
                TotalPages = totalPages,
                CursorPageNumber = cursorPage,
                DocTitle = doc.Name ?? string.Empty
            };
        }

        private static int GetCursorPageNumber(Word.Application wordApp, Word.Document doc = null)
        {
            try
            {
                if (wordApp?.Selection != null)
                {
                    object pageObj = wordApp.Selection.Information[Word.WdInformation.wdActiveEndPageNumber];
                    if (pageObj != null)
                    {
                        return Convert.ToInt32(pageObj);
                    }
                }
            }
            catch (Exception)
            {
            }

            // 无选区时回退文档起始页，避免只读路径强依赖 ActiveDocument
            try
            {
                if (doc?.Content != null)
                {
                    object pageObj = doc.Content.Information[Word.WdInformation.wdActiveEndPageNumber];
                    if (pageObj != null)
                    {
                        return Convert.ToInt32(pageObj);
                    }
                }
            }
            catch (Exception)
            {
            }

            return 1;
        }

        private static int ExportPageToPdf(Word.Application wordApp, Word.Document doc, string pdfPath, int pageNumber)
        {
            if (TryExportPageRange(doc, pdfPath, pageNumber))
            {
                return 0;
            }

            if (TryExportSinglePageViaTempDocument(wordApp, doc, pdfPath, pageNumber))
            {
                return 0;
            }

            throw new InvalidOperationException(
                $"无法导出第 {pageNumber} 页为 PDF（单页范围导出与临时文档导出均失败）");
        }

        /// <summary>
        /// 单页范围导出失败时，复制该页到临时文档再导出，避免全文 PDF 导出阻塞 UI/WebSocket。
        /// </summary>
        private static bool TryExportSinglePageViaTempDocument(
            Word.Application wordApp,
            Word.Document doc,
            string pdfPath,
            int pageNumber)
        {
            Word.Document tempDoc = null;
            try
            {
                int totalPages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                if (pageNumber < 1)
                {
                    pageNumber = 1;
                }

                if (pageNumber > totalPages)
                {
                    pageNumber = totalPages;
                }

                Word.Range pageStart = doc.GoTo(
                    Word.WdGoToItem.wdGoToPage,
                    Word.WdGoToDirection.wdGoToAbsolute,
                    pageNumber);

                Word.Range pageEnd;
                if (pageNumber >= totalPages)
                {
                    pageEnd = doc.Content;
                }
                else
                {
                    pageEnd = doc.GoTo(
                        Word.WdGoToItem.wdGoToPage,
                        Word.WdGoToDirection.wdGoToAbsolute,
                        pageNumber + 1);
                }

                int endPos = Math.Max(pageStart.Start, pageEnd.Start - 1);
                Word.Range pageRange = doc.Range(pageStart.Start, endPos);
                pageRange.Copy();

                tempDoc = wordApp.Documents.Add(Visible: false);
                tempDoc.Content.Paste();

                tempDoc.ExportAsFixedFormat(
                    pdfPath,
                    Word.WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport: false,
                    OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint);

                if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
                {
                    return false;
                }

                using (var pdfDocument = PdfDocument.Load(pdfPath))
                {
                    return pdfDocument.PageCount > 0;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[PageCaptureHelper] 临时文档单页导出失败: {ex.Message}");
                return false;
            }
            finally
            {
                if (tempDoc != null)
                {
                    try
                    {
                        tempDoc.Close(SaveChanges: false);
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }
        }

        private static bool TryExportPageRange(Word.Document doc, string pdfPath, int pageNumber)
        {
            try
            {
                doc.ExportAsFixedFormat(
                    pdfPath,
                    Word.WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport: false,
                    OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Range: Word.WdExportRange.wdExportFromTo,
                    From: pageNumber,
                    To: pageNumber);

                if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
                {
                    return false;
                }

                using (var pdfDocument = PdfDocument.Load(pdfPath))
                {
                    return pdfDocument.PageCount > 0;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PageCaptureHelper] 单页导出失败，将尝试临时文档导出: {ex.Message}");
                return false;
            }
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // ignore cleanup errors
            }
        }

        private static void TryDeleteDirectory(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                {
                    Directory.Delete(path, recursive: true);
                }
            }
            catch
            {
                // ignore cleanup errors
            }
        }
    }
}
