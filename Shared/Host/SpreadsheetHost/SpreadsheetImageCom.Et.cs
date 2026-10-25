using System;
using System.Collections.Generic;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.SpreadsheetHost
{
    internal static partial class SpreadsheetImageCom
    {
        public static bool TryExecuteOnEtWorkbook(
            object book,
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
                    return TryListEt(book, request, result, out error);
                }

                if (request.Action == "insert")
                {
                    return TryInsertEt(book, request, result, out error);
                }

                if (request.Action == "update")
                {
                    return TryUpdateEt(book, request, result, out error);
                }

                return TryDeleteEt(book, request, result, out error);
            }
            catch (Exception ex)
            {
                error = "图片操作失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryListEt(
            object book,
            SpreadsheetImageRequest request,
            SpreadsheetImageResult result,
            out string error)
        {
            error = null;
            List<object> pics = CollectEtPictures(book, request.Sheet, out error);
            if (error != null)
            {
                return false;
            }

            var occupied = new HashSet<string>(StringComparer.Ordinal);
            var collisions = new HashSet<string>(StringComparer.Ordinal);
            foreach (object pic in pics)
            {
                if (SpreadsheetImageStampHelper.TryParse(ReadAltEt(pic), out string id, out _))
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
            foreach (object pic in pics)
            {
                string alt = ReadAltEt(pic);
                if (!SpreadsheetImageStampHelper.TryParse(alt, out string id, out string suffix))
                {
                    id = SpreadsheetImageStampHelper.NewRandomImageId(occupied);
                    if (id == null)
                    {
                        error = "无法分配 image_id";
                        return false;
                    }

                    EtCom.TrySetProperty(pic, "AlternativeText", SpreadsheetImageStampHelper.Format(id, suffix));
                }

                FillInfoEt(pic, id, images, result.Warnings);
            }

            result.Images = images;
            result.ImagesCount = images.Count;
            return true;
        }

        private static bool TryInsertEt(
            object book,
            SpreadsheetImageRequest request,
            SpreadsheetImageResult result,
            out string error)
        {
            if (!TryGetEtSheet(book, request.Sheet, out object sheet, out error))
            {
                return false;
            }

            object cell;
            try
            {
                cell = EtCom.InvokeFlex(sheet, "Range", request.Anchor);
            }
            catch
            {
                error = "无法定位锚点格: " + request.Anchor;
                return false;
            }

            float left = Convert.ToSingle(EtCom.GetProperty(cell, "Left"));
            float top = Convert.ToSingle(EtCom.GetProperty(cell, "Top"));
            double widthCm = request.WidthCm ?? SpreadsheetImageParse.DefaultWidthCm;
            float widthPt = (float)(widthCm * PointsPerCm);
            float heightPt = widthPt;
            double? aspect = TryGetImageAspectRatio(request.LocalPath);
            if (aspect.HasValue && aspect.Value > 0)
            {
                heightPt = (float)(widthPt / aspect.Value);
            }

            object shapes = EtCom.GetProperty(sheet, "Shapes");
            object pic = EtCom.Invoke(
                shapes,
                "AddPicture",
                request.LocalPath,
                0,
                -1,
                left,
                top,
                widthPt,
                heightPt);
            if (pic == null)
            {
                error = "AddPicture 失败";
                return false;
            }

            EtCom.TrySetProperty(pic, "LockAspectRatio", -1);
            EtCom.TrySetProperty(pic, "Placement", XlMove);

            var occupied = CollectOccupiedEtIds(book);
            string id = SpreadsheetImageStampHelper.NewRandomImageId(occupied);
            if (id == null)
            {
                error = "无法分配 image_id";
                return false;
            }

            EtCom.TrySetProperty(pic, "AlternativeText", SpreadsheetImageStampHelper.Format(id, ""));
            FillResultFromShapeEt(pic, id, request, result);
            return true;
        }

        private static bool TryUpdateEt(
            object book,
            SpreadsheetImageRequest request,
            SpreadsheetImageResult result,
            out string error)
        {
            if (!TryFindEtPicture(book, request.ImageId, request.Sheet, out object pic, out error))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(request.LocalPath))
            {
                bool replaced = false;
                try
                {
                    object fill = EtCom.GetProperty(pic, "Fill");
                    EtCom.Invoke(fill, "UserPicture", request.LocalPath);
                    replaced = true;
                }
                catch
                {
                }

                if (!replaced)
                {
                    object sheet = EtCom.GetProperty(pic, "Parent");
                    float left = Convert.ToSingle(EtCom.GetProperty(pic, "Left"));
                    float top = Convert.ToSingle(EtCom.GetProperty(pic, "Top"));
                    float w = Convert.ToSingle(EtCom.GetProperty(pic, "Width"));
                    float h = Convert.ToSingle(EtCom.GetProperty(pic, "Height"));
                    string alt = ReadAltEt(pic);
                    object placement = EtCom.GetProperty(pic, "Placement");
                    EtCom.Invoke(pic, "Delete");
                    object shapes = EtCom.GetProperty(sheet, "Shapes");
                    pic = EtCom.Invoke(
                        shapes,
                        "AddPicture",
                        request.LocalPath,
                        0,
                        -1,
                        left,
                        top,
                        w,
                        h);
                    EtCom.TrySetProperty(pic, "AlternativeText", alt);
                    if (placement != null)
                    {
                        EtCom.TrySetProperty(pic, "Placement", placement);
                    }

                    EtCom.TrySetProperty(pic, "LockAspectRatio", -1);
                }
            }

            if (request.WidthCm.HasValue)
            {
                EtCom.TrySetProperty(pic, "LockAspectRatio", -1);
                EtCom.TrySetProperty(pic, "Width", request.WidthCm.Value * PointsPerCm);
            }

            if (!string.IsNullOrEmpty(request.Anchor))
            {
                object currentSheet = EtCom.GetProperty(pic, "Parent");
                string destName = request.Sheet;
                if (string.IsNullOrEmpty(destName))
                {
                    destName = EtCom.TryReadName(currentSheet);
                }

                if (!TryGetEtSheet(book, destName, out object destSheet, out error))
                {
                    return false;
                }

                object cell;
                try
                {
                    cell = EtCom.InvokeFlex(destSheet, "Range", request.Anchor);
                }
                catch
                {
                    error = "无法定位锚点格: " + request.Anchor;
                    return false;
                }

                string currentName = EtCom.TryReadName(currentSheet) ?? "";
                if (!string.Equals(currentName, destName, StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        EtCom.Invoke(pic, "Cut");
                        EtCom.Invoke(destSheet, "Paste");
                        object shapes = EtCom.GetProperty(destSheet, "Shapes");
                        int count = Convert.ToInt32(EtCom.GetProperty(shapes, "Count"));
                        pic = EtCom.GetIndexed(shapes, count);
                    }
                    catch (Exception ex)
                    {
                        error = "换表失败: " + ex.Message;
                        return false;
                    }
                }

                EtCom.TrySetProperty(pic, "Left", EtCom.GetProperty(cell, "Left"));
                EtCom.TrySetProperty(pic, "Top", EtCom.GetProperty(cell, "Top"));
            }

            FillResultFromShapeEt(pic, request.ImageId, request, result);
            return true;
        }

        private static bool TryDeleteEt(
            object book,
            SpreadsheetImageRequest request,
            SpreadsheetImageResult result,
            out string error)
        {
            if (!TryFindEtPicture(book, request.ImageId, request.Sheet, out object pic, out error))
            {
                return false;
            }

            result.ImageId = request.ImageId;
            object parent = EtCom.GetProperty(pic, "Parent");
            result.Sheet = EtCom.TryReadName(parent);
            EtCom.Invoke(pic, "Delete");
            return true;
        }

        private static bool TryFindEtPicture(
            object book,
            string imageId,
            string sheetFilter,
            out object pic,
            out string error)
        {
            pic = null;
            error = null;
            var hits = new List<object>();
            foreach (object item in CollectEtPictures(book, sheetFilter, out error))
            {
                if (error != null)
                {
                    return false;
                }

                if (SpreadsheetImageStampHelper.TryParse(ReadAltEt(item), out string id, out _)
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

        private static List<object> CollectEtPictures(object book, string sheetFilter, out string error)
        {
            error = null;
            var list = new List<object>();
            bool sheetHit = string.IsNullOrEmpty(sheetFilter);
            foreach (object sheet in EtCom.EnumerateSheets(book))
            {
                if (EtCom.IsChartSheet(sheet))
                {
                    continue;
                }

                string name = EtCom.TryReadName(sheet) ?? "";
                if (!string.IsNullOrEmpty(sheetFilter)
                    && !string.Equals(name, sheetFilter, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(sheetFilter))
                {
                    sheetHit = true;
                }

                object shapes = EtCom.GetProperty(sheet, "Shapes");
                if (shapes == null)
                {
                    continue;
                }

                int count;
                try
                {
                    count = Convert.ToInt32(EtCom.GetProperty(shapes, "Count"));
                }
                catch
                {
                    continue;
                }

                for (int i = 1; i <= count; i++)
                {
                    object shape = EtCom.GetIndexed(shapes, i);
                    if (IsPictureShapeEt(shape))
                    {
                        list.Add(shape);
                    }
                }
            }

            if (!string.IsNullOrEmpty(sheetFilter) && !sheetHit)
            {
                error = "未找到工作表: " + sheetFilter;
            }

            return list;
        }

        private static bool TryGetEtSheet(object book, string name, out object sheet, out string error)
        {
            sheet = null;
            error = null;
            foreach (object item in EtCom.EnumerateSheets(book))
            {
                if (EtCom.IsChartSheet(item))
                {
                    continue;
                }

                if (string.Equals(EtCom.TryReadName(item), name, StringComparison.OrdinalIgnoreCase))
                {
                    sheet = item;
                    return true;
                }
            }

            error = "未找到工作表: " + name;
            return false;
        }

        private static HashSet<string> CollectOccupiedEtIds(object book)
        {
            var occupied = new HashSet<string>(StringComparer.Ordinal);
            foreach (object pic in CollectEtPictures(book, null, out _))
            {
                if (SpreadsheetImageStampHelper.TryParse(ReadAltEt(pic), out string id, out _))
                {
                    occupied.Add(id);
                }
            }

            return occupied;
        }

        private static bool IsPictureShapeEt(object shape)
        {
            try
            {
                object typeObj = EtCom.GetProperty(shape, "Type");
                int t = Convert.ToInt32(typeObj);
                if (t == 3)
                {
                    return false;
                }

                return t == MsoPicture || t == MsoLinkedPicture;
            }
            catch
            {
                return false;
            }
        }

        private static string ReadEtAnchor(object pic)
        {
            object tl = EtCom.GetProperty(pic, "TopLeftCell");
            if (tl == null)
            {
                return "";
            }

            try
            {
                int row = Convert.ToInt32(EtCom.GetProperty(tl, "Row"));
                int col = Convert.ToInt32(EtCom.GetProperty(tl, "Column"));
                string a1 = A1Address.Cell(row, col);
                if (!string.IsNullOrEmpty(a1))
                {
                    return a1;
                }
            }
            catch
            {
            }

            try
            {
                object addr = EtCom.InvokeFlex(tl, "Address", false, false);
                return (Convert.ToString(addr) ?? "").Replace("$", "");
            }
            catch
            {
                return (Convert.ToString(EtCom.GetProperty(tl, "Address")) ?? "").Replace("$", "");
            }
        }

        private static string ReadAltEt(object shape)
        {
            try
            {
                return Convert.ToString(EtCom.GetProperty(shape, "AlternativeText")) ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static void FillInfoEt(object pic, string id, List<SpreadsheetImageInfo> images, List<string> warnings)
        {
            string sheetName = "";
            string anchor = "";
            try
            {
                sheetName = EtCom.TryReadName(EtCom.GetProperty(pic, "Parent")) ?? "";
            }
            catch
            {
            }

            try
            {
                anchor = ReadEtAnchor(pic);
                if (string.IsNullOrEmpty(anchor))
                {
                    warnings?.Add("无法读取锚点格");
                }
            }
            catch
            {
                warnings?.Add("无法读取锚点格");
            }

            double w = 0;
            double h = 0;
            try
            {
                w = Convert.ToDouble(EtCom.GetProperty(pic, "Width"));
                h = Convert.ToDouble(EtCom.GetProperty(pic, "Height"));
            }
            catch
            {
            }

            images.Add(new SpreadsheetImageInfo
            {
                ImageId = id,
                Sheet = sheetName,
                Anchor = anchor,
                WidthCm = Math.Round(w / PointsPerCm, 2),
                HeightCm = Math.Round(h / PointsPerCm, 2)
            });
        }

        private static void FillResultFromShapeEt(
            object pic,
            string id,
            SpreadsheetImageRequest request,
            SpreadsheetImageResult result)
        {
            result.ImageId = id;
            object parent = EtCom.GetProperty(pic, "Parent");
            result.Sheet = EtCom.TryReadName(parent) ?? request.Sheet;
            try
            {
                string a1 = ReadEtAnchor(pic);
                result.Anchor = !string.IsNullOrEmpty(a1) ? a1 : (request.Anchor ?? "");
            }
            catch
            {
                result.Anchor = request.Anchor ?? "";
            }

            try
            {
                result.WidthCm = Math.Round(Convert.ToDouble(EtCom.GetProperty(pic, "Width")) / PointsPerCm, 2);
                result.HeightCm = Math.Round(Convert.ToDouble(EtCom.GetProperty(pic, "Height")) / PointsPerCm, 2);
            }
            catch
            {
            }

            result.PathDisplay = request.PathDisplay;
        }
    }
}
