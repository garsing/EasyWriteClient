using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 图片编号 I_ 与 InlineShape（非 Chart）定位。
    /// </summary>
    public static class ImageResolveHelper
    {
        public static List<Word.InlineShape> GetInlineImagesInOrder(Word.Document document)
        {
            var images = new List<Word.InlineShape>();
            if (document == null)
            {
                return images;
            }

            try
            {
                foreach (Word.InlineShape shape in document.InlineShapes)
                {
                    if (IsPictureInlineShape(shape))
                    {
                        images.Add(shape);
                    }
                }

                images.Sort((a, b) => a.Range.Start.CompareTo(b.Range.Start));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ImageResolve] 收集 InlineShape 图片失败: {ex.Message}");
            }

            return images;
        }

        public static Word.InlineShape ResolveInlineShapeByImageId(Word.Document document, string imageId)
        {
            if (document == null || string.IsNullOrEmpty(imageId))
            {
                return null;
            }

            int imageOrderIndex = DocumentState.GetImageIndex(imageId);
            if (imageOrderIndex < 0)
            {
                return null;
            }

            List<Word.InlineShape> allImages = GetInlineImagesInOrder(document);
            if (imageOrderIndex >= allImages.Count)
            {
                return null;
            }

            return allImages[imageOrderIndex];
        }

        /// <summary>
        /// 插图后：轻量重读 → 按 Range.Start 在有序图片列表中定位 → 取 ImageIdOrder[index]。
        /// 失败时写入 warnings，不抛异常。
        /// </summary>
        public static string TryResolveImageIdAfterInsert(
            Word.Document document,
            int pictureRangeStart,
            System.Collections.Generic.List<string> warnings)
        {
            if (document == null || pictureRangeStart < 0)
            {
                warnings?.Add("未能解析 image_id（无效插入位置）");
                return null;
            }

            try
            {
                WordReader.ReadWord(document, docPath: null, extractImages: false);
                List<Word.InlineShape> allImages = GetInlineImagesInOrder(document);
                int imageIndex = allImages.FindIndex(s => s.Range.Start == pictureRangeStart);
                var idOrder = DocumentState.GetImageIdOrder();
                if (imageIndex >= 0 && imageIndex < idOrder.Count)
                {
                    return idOrder[imageIndex];
                }

                warnings?.Add("未能解析 image_id（图片索引与编号映射不一致）");
            }
            catch (Exception ex)
            {
                warnings?.Add($"未能解析 image_id: {ex.Message}");
            }

            return null;
        }

        private static bool IsPictureInlineShape(Word.InlineShape shape)
        {
            if (shape == null)
            {
                return false;
            }

            try
            {
                if (shape.Type == Word.WdInlineShapeType.wdInlineShapeChart)
                {
                    return false;
                }

                return shape.Type == Word.WdInlineShapeType.wdInlineShapePicture
                    || shape.Type == Word.WdInlineShapeType.wdInlineShapePictureHorizontalLine
                    || shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture
                    || shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPictureHorizontalLine;
            }
            catch
            {
                return false;
            }
        }
    }
}
