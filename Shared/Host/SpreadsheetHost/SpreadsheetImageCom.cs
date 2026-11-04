using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.SpreadsheetHost
{
    internal static partial class SpreadsheetImageCom
    {
        private const double PointsPerCm = 28.35;
        private const int MsoPicture = 13;
        private const int MsoLinkedPicture = 11;
        private const int XlMove = 2;

        public static bool TryExecuteOnExcelWorkbook(
            Excel.Workbook book,
            SpreadsheetImageRequest request,
            out SpreadsheetImageResult result,
            out string error)
        {
            result = new SpreadsheetImageResult
            {
                Action = request.Action,
                Warnings = new List<string>()
            };
            error = null;
            try
            {
                if (request.Action == "list")
                {
                    return TryListExcel(book, request, result, out error);
                }

                if (request.Action == "export")
                {
                    return TryExportExcel(book, request, result, out error);
                }

                if (request.Action == "insert")
                {
                    return TryInsertExcel(book, request, result, out error);
                }

                if (request.Action == "update")
                {
                    return TryUpdateExcel(book, request, result, out error);
                }

                return TryDeleteExcel(book, request, result, out error);
            }
            catch (Exception ex)
            {
                error = "图片操作失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryListExcel(
            Excel.Workbook book,
            SpreadsheetImageRequest request,
            SpreadsheetImageResult result,
            out string error)
        {
            error = null;
            List<Excel.Shape> pics = CollectExcelPictures(book, request.Sheet, out error);
            if (error != null)
            {
                return false;
            }

            var occupied = new HashSet<string>(StringComparer.Ordinal);
            var collisions = new HashSet<string>(StringComparer.Ordinal);
            foreach (Excel.Shape pic in pics)
            {
                if (SpreadsheetImageStampHelper.TryParse(ReadAlt(pic), out string id, out _))
                {
                    if (!occupied.Add(id))
                    {
                        collisions.Add(id);
                    }
                }
            }

            if (collisions.Count > 0)
            {
                error = "簿内两张图撞同一个 I_";
                return false;
            }

            var images = new List<SpreadsheetImageInfo>();
            foreach (Excel.Shape pic in pics)
            {
                string alt = ReadAlt(pic);
                if (!SpreadsheetImageStampHelper.TryParse(alt, out string id, out string suffix))
                {
                    id = SpreadsheetImageStampHelper.NewRandomImageId(occupied);
                    if (id == null)
                    {
                        error = "无法分配 image_id";
                        return false;
                    }

                    pic.AlternativeText = SpreadsheetImageStampHelper.Format(id, suffix);
                }

                FillInfo(pic, id, images, result.Warnings);
            }

            result.Images = images;
            result.ImagesCount = images.Count;
            return true;
        }

        private static bool TryExportExcel(
            Excel.Workbook book,
            SpreadsheetImageRequest request,
            SpreadsheetImageResult result,
            out string error)
        {
            error = null;
            string fullBefore = "";
            bool savedBefore = true;
            DateTime mtimeBefore = DateTime.MinValue;
            bool sawMtime = false;
            bool stamped = false;
            try
            {
                fullBefore = book.FullName ?? "";
                savedBefore = book.Saved;
                if (File.Exists(fullBefore))
                {
                    mtimeBefore = File.GetLastWriteTimeUtc(fullBefore);
                    sawMtime = true;
                }
            }
            catch
            {
            }

            try
            {
                if (!string.IsNullOrEmpty(request.Sheet))
                {
                    CollectExcelPictures(book, request.Sheet, out error);
                    if (error != null)
                    {
                        return false;
                    }
                }

                List<Excel.Shape> all = CollectExcelPictures(book, null, out error);
                if (error != null)
                {
                    return false;
                }

                if (!TryFindExcelIdCollision(all, out error))
                {
                    return false;
                }

                var chosen = new List<Excel.Shape>();
                foreach (Excel.Shape pic in all)
                {
                    if (!string.IsNullOrEmpty(request.Sheet) && !SameSheet(SheetNameOf(pic), request.Sheet))
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(request.ImageId))
                    {
                        if (!SpreadsheetImageStampHelper.TryParse(ReadAlt(pic), out string existing, out _)
                            || !string.Equals(existing, request.ImageId, StringComparison.Ordinal))
                        {
                            continue;
                        }
                    }

                    chosen.Add(pic);
                }

                if (!string.IsNullOrEmpty(request.ImageId) && chosen.Count == 0)
                {
                    error = "未找到图片: " + request.ImageId;
                    return false;
                }

                if (!TryAssignExcelIds(chosen, out Dictionary<Excel.Shape, string> ids, out bool didStamp, out error))
                {
                    return false;
                }

                stamped = didStamp;

                var images = new List<SpreadsheetImageInfo>();
                if (chosen.Count > 0)
                {
                    if (!CurrDocsImagePaths.TryEnsureHostDirectory(book.Name, CurrDocsImagePaths.ExcelSuffix, out string directory, out error))
                    {
                        return false;
                    }

                    foreach (Excel.Shape pic in chosen)
                    {
                        string id = ids[pic];
                        string finalPath = Path.Combine(directory, id + ".png");
                        if (!TryExportExcelShapePng(pic, finalPath, out error))
                        {
                            return false;
                        }

                        FillInfo(pic, id, images, result.Warnings);
                        images[images.Count - 1].Path = Path.GetFullPath(finalPath).Replace('\\', '/');
                    }
                }

                string fullAfter = "";
                try
                {
                    fullAfter = book.FullName ?? "";
                }
                catch
                {
                }

                if (!string.IsNullOrEmpty(fullBefore)
                    && !string.Equals(NormalizePath(fullBefore), NormalizePath(fullAfter), StringComparison.OrdinalIgnoreCase))
                {
                    error = "导出图片时活动工作簿路径被改变";
                    return false;
                }

                result.Images = images;
                result.ImagesCount = images.Count;
                return true;
            }
            finally
            {
                RestoreExcelBookState(book, fullBefore, savedBefore, mtimeBefore, sawMtime, stamped);
            }
        }

        private static bool TryFindExcelIdCollision(List<Excel.Shape> pics, out string error)
        {
            error = null;
            var occupied = new HashSet<string>(StringComparer.Ordinal);
            foreach (Excel.Shape pic in pics)
            {
                if (SpreadsheetImageStampHelper.TryParse(ReadAlt(pic), out string id, out _) && !occupied.Add(id))
                {
                    error = "簿内两张图撞同一个 I_";
                    return false;
                }
            }

            return true;
        }

        private static bool TryAssignExcelIds(
            List<Excel.Shape> pics,
            out Dictionary<Excel.Shape, string> ids,
            out bool stamped,
            out string error)
        {
            ids = new Dictionary<Excel.Shape, string>();
            stamped = false;
            error = null;
            var occupied = new HashSet<string>(StringComparer.Ordinal);
            var collisions = new HashSet<string>(StringComparer.Ordinal);
            foreach (Excel.Shape pic in pics)
            {
                if (SpreadsheetImageStampHelper.TryParse(ReadAlt(pic), out string id, out _))
                {
                    if (!occupied.Add(id))
                    {
                        collisions.Add(id);
                    }
                }
            }

            if (collisions.Count > 0)
            {
                error = "簿内两张图撞同一个 I_";
                return false;
            }

            foreach (Excel.Shape pic in pics)
            {
                string alt = ReadAlt(pic);
                if (!SpreadsheetImageStampHelper.TryParse(alt, out string id, out string suffix))
                {
                    id = SpreadsheetImageStampHelper.NewRandomImageId(occupied);
                    if (id == null)
                    {
                        error = "无法分配 image_id";
                        return false;
                    }

                    pic.AlternativeText = SpreadsheetImageStampHelper.Format(id, suffix);
                    stamped = true;
                }

                ids[pic] = id;
            }

            return true;
        }

        private static bool TryExportExcelShapePng(Excel.Shape pic, string finalPath, out string error)
        {
            error = null;
            string tempPath = Path.Combine(
                Path.GetDirectoryName(finalPath) ?? "",
                "_tmp_" + Guid.NewGuid().ToString("N") + ".png");
            Excel.ChartObject chartObj = null;
            try
            {
                Excel.Worksheet sheet = pic.Parent as Excel.Worksheet;
                if (sheet == null)
                {
                    error = "导出图片失败 (" + finalPath + "): 找不到所在工作表";
                    return false;
                }

                try
                {
                    sheet.Activate();
                }
                catch
                {
                }

                double width = Math.Max(10, pic.Width);
                double height = Math.Max(10, pic.Height);
                Excel.XlCopyPictureFormat[] formats =
                {
                    Excel.XlCopyPictureFormat.xlBitmap,
                    Excel.XlCopyPictureFormat.xlPicture
                };
                Exception last = null;
                foreach (Excel.XlCopyPictureFormat format in formats)
                {
                    try
                    {
                        pic.CopyPicture(Excel.XlPictureAppearance.xlScreen, format);
                        chartObj = (Excel.ChartObject)sheet.ChartObjects().Add(pic.Left, pic.Top, width, height);
                        try
                        {
                            chartObj.Activate();
                        }
                        catch
                        {
                        }

                        chartObj.Chart.Paste();
                        chartObj.Chart.Export(tempPath, "PNG");
                        DeleteExcelChartQuiet(ref chartObj);
                        if (File.Exists(tempPath) && new FileInfo(tempPath).Length > 8 && IsPngFile(tempPath))
                        {
                            File.Copy(tempPath, finalPath, true);
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        last = ex;
                        DeleteExcelChartQuiet(ref chartObj);
                    }
                }

                error = last != null
                    ? "导出图片失败: " + last.Message
                    : "导出图片失败: 未生成 PNG";
                return false;
            }
            catch (Exception ex)
            {
                error = "导出图片失败: " + ex.Message;
                return false;
            }
            finally
            {
                DeleteExcelChartQuiet(ref chartObj);
                try
                {
                    if (File.Exists(tempPath))
                    {
                        File.Delete(tempPath);
                    }
                }
                catch
                {
                }
            }
        }

        private static void DeleteExcelChartQuiet(ref Excel.ChartObject chartObj)
        {
            if (chartObj == null)
            {
                return;
            }

            try
            {
                chartObj.Delete();
            }
            catch
            {
            }

            chartObj = null;
        }

        private static void RestoreExcelBookState(
            Excel.Workbook book,
            string fullBefore,
            bool savedBefore,
            DateTime mtimeBefore,
            bool sawMtime,
            bool stamped)
        {
            try
            {
                bool diskWritten = false;
                if (sawMtime && File.Exists(fullBefore))
                {
                    diskWritten = File.GetLastWriteTimeUtc(fullBefore) != mtimeBefore;
                }

                if (diskWritten)
                {
                    return;
                }

                if (stamped)
                {
                    if (book.Saved)
                    {
                        book.Saved = false;
                    }

                    return;
                }

                if (book.Saved != savedBefore)
                {
                    book.Saved = savedBefore;
                }
            }
            catch
            {
            }
        }

        private static string SheetNameOf(Excel.Shape pic)
        {
            try
            {
                return ((Excel.Worksheet)pic.Parent).Name ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static bool SameSheet(string actual, string wanted)
        {
            return string.Equals(actual ?? "", wanted ?? "", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsPngFile(string path)
        {
            try
            {
                using (var stream = File.OpenRead(path))
                {
                    if (stream.Length < 8)
                    {
                        return false;
                    }

                    var magic = new byte[4];
                    stream.Read(magic, 0, 4);
                    return magic[0] == 0x89 && magic[1] == 0x50 && magic[2] == 0x4E && magic[3] == 0x47;
                }
            }
            catch
            {
                return false;
            }
        }

        private static string NormalizePath(string path)
        {
            return (path ?? "").Replace('/', '\\').Trim();
        }

        private static bool TryInsertExcel(
            Excel.Workbook book,
            SpreadsheetImageRequest request,
            SpreadsheetImageResult result,
            out string error)
        {
            if (!TryGetExcelSheet(book, request.Sheet, out Excel.Worksheet sheet, out error))
            {
                return false;
            }

            Excel.Range cell = sheet.Range[request.Anchor];
            float left = Convert.ToSingle(cell.Left);
            float top = Convert.ToSingle(cell.Top);
            double widthCm = request.WidthCm ?? SpreadsheetImageParse.DefaultWidthCm;
            float widthPt = (float)(widthCm * PointsPerCm);
            float heightPt = widthPt;
            double? aspect = TryGetImageAspectRatio(request.LocalPath);
            if (aspect.HasValue && aspect.Value > 0)
            {
                heightPt = (float)(widthPt / aspect.Value);
            }

            Excel.Shape pic = sheet.Shapes.AddPicture(
                request.LocalPath,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue,
                left,
                top,
                widthPt,
                heightPt);
            pic.LockAspectRatio = Office.MsoTriState.msoTrue;
            pic.Placement = Excel.XlPlacement.xlMove;

            var occupied = CollectOccupiedExcelIds(book);
            string id = SpreadsheetImageStampHelper.NewRandomImageId(occupied);
            if (id == null)
            {
                error = "无法分配 image_id";
                return false;
            }

            pic.AlternativeText = SpreadsheetImageStampHelper.Format(id, "");
            FillResultFromShape(pic, id, request, result);
            return true;
        }

        private static bool TryUpdateExcel(
            Excel.Workbook book,
            SpreadsheetImageRequest request,
            SpreadsheetImageResult result,
            out string error)
        {
            if (!TryFindExcelPicture(book, request.ImageId, request.Sheet, out Excel.Shape pic, out error))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(request.LocalPath))
            {
                bool replaced = false;
                try
                {
                    pic.Fill.UserPicture(request.LocalPath);
                    replaced = true;
                }
                catch
                {
                }

                if (!replaced)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)pic.Parent;
                    float left = pic.Left;
                    float top = pic.Top;
                    float w = pic.Width;
                    float h = pic.Height;
                    string alt = pic.AlternativeText;
                    Excel.XlPlacement placement = pic.Placement;
                    pic.Delete();
                    pic = sheet.Shapes.AddPicture(
                        request.LocalPath,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        left,
                        top,
                        w,
                        h);
                    pic.AlternativeText = alt;
                    pic.Placement = placement;
                    pic.LockAspectRatio = Office.MsoTriState.msoTrue;
                }
            }

            if (request.WidthCm.HasValue)
            {
                pic.LockAspectRatio = Office.MsoTriState.msoTrue;
                pic.Width = (float)(request.WidthCm.Value * PointsPerCm);
            }

            if (!string.IsNullOrEmpty(request.Anchor))
            {
                string destSheetName = request.Sheet;
                Excel.Worksheet currentSheet = (Excel.Worksheet)pic.Parent;
                if (string.IsNullOrEmpty(destSheetName))
                {
                    destSheetName = currentSheet.Name;
                }

                if (!TryGetExcelSheet(book, destSheetName, out Excel.Worksheet destSheet, out error))
                {
                    return false;
                }

                Excel.Range cell = destSheet.Range[request.Anchor];
                if (!string.Equals(currentSheet.Name, destSheet.Name, StringComparison.OrdinalIgnoreCase))
                {
                    pic.Cut();
                    destSheet.Paste();
                    pic = destSheet.Shapes.Item(destSheet.Shapes.Count);
                }

                pic.Left = Convert.ToSingle(cell.Left);
                pic.Top = Convert.ToSingle(cell.Top);
            }

            FillResultFromShape(pic, request.ImageId, request, result);
            return true;
        }

        private static bool TryDeleteExcel(
            Excel.Workbook book,
            SpreadsheetImageRequest request,
            SpreadsheetImageResult result,
            out string error)
        {
            if (!TryFindExcelPicture(book, request.ImageId, request.Sheet, out Excel.Shape pic, out error))
            {
                return false;
            }

            result.ImageId = request.ImageId;
            result.Sheet = ((Excel.Worksheet)pic.Parent).Name;
            pic.Delete();
            return true;
        }

        private static bool TryFindExcelPicture(
            Excel.Workbook book,
            string imageId,
            string sheetFilter,
            out Excel.Shape pic,
            out string error)
        {
            pic = null;
            error = null;
            List<Excel.Shape> hits = new List<Excel.Shape>();
            foreach (Excel.Shape item in CollectExcelPictures(book, sheetFilter, out error))
            {
                if (error != null)
                {
                    return false;
                }

                if (SpreadsheetImageStampHelper.TryParse(ReadAlt(item), out string id, out _)
                    && string.Equals(id, imageId, StringComparison.Ordinal))
                {
                    hits.Add(item);
                }
            }

            if (hits.Count == 0)
            {
                error = "未找到图片: " + imageId;
                return false;
            }

            if (hits.Count > 1)
            {
                error = "簿内两张图撞同一个 I_";
                return false;
            }

            pic = hits[0];
            return true;
        }

        private static List<Excel.Shape> CollectExcelPictures(Excel.Workbook book, string sheetFilter, out string error)
        {
            error = null;
            var list = new List<Excel.Shape>();
            foreach (object item in book.Sheets)
            {
                Excel.Worksheet sheet = item as Excel.Worksheet;
                if (sheet == null)
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(sheetFilter)
                    && !string.Equals(sheet.Name, sheetFilter, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                foreach (Excel.Shape shape in sheet.Shapes)
                {
                    if (IsPictureShape(shape))
                    {
                        list.Add(shape);
                    }
                }
            }

            if (!string.IsNullOrEmpty(sheetFilter) && list.Count == 0)
            {
                if (!SheetExistsExcel(book, sheetFilter))
                {
                    error = "未找到工作表: " + sheetFilter;
                }
            }

            return list;
        }

        private static bool SheetExistsExcel(Excel.Workbook book, string name)
        {
            foreach (object item in book.Sheets)
            {
                Excel.Worksheet sheet = item as Excel.Worksheet;
                if (sheet != null && string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetExcelSheet(
            Excel.Workbook book,
            string name,
            out Excel.Worksheet sheet,
            out string error)
        {
            sheet = null;
            error = null;
            foreach (object item in book.Sheets)
            {
                Excel.Worksheet ws = item as Excel.Worksheet;
                if (ws != null && string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    sheet = ws;
                    return true;
                }
            }

            error = "未找到工作表: " + name;
            return false;
        }

        private static HashSet<string> CollectOccupiedExcelIds(Excel.Workbook book)
        {
            var occupied = new HashSet<string>(StringComparer.Ordinal);
            foreach (Excel.Shape pic in CollectExcelPictures(book, null, out _))
            {
                if (SpreadsheetImageStampHelper.TryParse(ReadAlt(pic), out string id, out _))
                {
                    occupied.Add(id);
                }
            }

            return occupied;
        }

        private static bool IsPictureShape(Excel.Shape shape)
        {
            try
            {
                if (shape.Type == Office.MsoShapeType.msoChart)
                {
                    return false;
                }

                int t = (int)shape.Type;
                return t == MsoPicture || t == MsoLinkedPicture;
            }
            catch
            {
                return false;
            }
        }

        private static string ReadAlt(Excel.Shape shape)
        {
            try
            {
                return shape.AlternativeText ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static void FillInfo(Excel.Shape pic, string id, List<SpreadsheetImageInfo> images, List<string> warnings)
        {
            string sheetName = "";
            string anchor = "";
            try
            {
                sheetName = ((Excel.Worksheet)pic.Parent).Name;
            }
            catch
            {
            }

            try
            {
                Excel.Range tl = pic.TopLeftCell;
                if (tl != null)
                {
                    anchor = tl.Address[false, false];
                }
            }
            catch
            {
                warnings?.Add("无法读取锚点格");
            }

            images.Add(new SpreadsheetImageInfo
            {
                ImageId = id,
                Sheet = sheetName,
                Anchor = anchor,
                WidthCm = Math.Round(pic.Width / PointsPerCm, 2),
                HeightCm = Math.Round(pic.Height / PointsPerCm, 2)
            });
        }

        private static void FillResultFromShape(
            Excel.Shape pic,
            string id,
            SpreadsheetImageRequest request,
            SpreadsheetImageResult result)
        {
            result.ImageId = id;
            try
            {
                result.Sheet = ((Excel.Worksheet)pic.Parent).Name;
            }
            catch
            {
                result.Sheet = request.Sheet;
            }

            try
            {
                result.Anchor = pic.TopLeftCell != null ? pic.TopLeftCell.Address[false, false] : (request.Anchor ?? "");
            }
            catch
            {
                result.Anchor = request.Anchor ?? "";
            }

            result.WidthCm = Math.Round(pic.Width / PointsPerCm, 2);
            result.HeightCm = Math.Round(pic.Height / PointsPerCm, 2);
            result.PathDisplay = request.PathDisplay;
        }

        internal static double? TryGetImageAspectRatio(string localPath)
        {
            try
            {
                using (var img = Image.FromFile(localPath))
                {
                    if (img.Height <= 0)
                    {
                        return null;
                    }

                    return (double)img.Width / img.Height;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
