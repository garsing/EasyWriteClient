using System;
using System.Collections.Generic;
using System.IO;
using WordAddIn1.OpenFiles;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.ChartOoxmlHost
{
    /// <summary>文字侧：WordOpenXML 读包；Save + File.Copy 改包贴回。按 Kind 分叉，不单开 Wps 目录。</summary>
    internal static class WordChartPackageHost
    {
        public static bool TryReadParts(object document, out List<ChartPartRef> parts, out string error)
        {
            parts = null;
            if (!TryGetWordOpenXml(document, out string xml, out error))
            {
                return false;
            }

            return ChartPartIndex.TryIndexWordFlatOpc(xml, out parts, out error);
        }

        public static bool TryGetWordOpenXml(object document, out string xml, out string error)
        {
            xml = null;
            error = null;
            if (document == null)
            {
                error = "读包：文档空";
                return false;
            }

            try
            {
                Word.Document word = document as Word.Document;
                if (word != null && word.Content != null)
                {
                    xml = word.Content.WordOpenXML;
                }
                else
                {
                    object content = WpsCom.GetProperty(document, "Content");
                    xml = Convert.ToString(WpsCom.GetProperty(content, "WordOpenXML") ?? "");
                }
            }
            catch (Exception ex)
            {
                error = "取 WordOpenXML 失败: " + ex.Message;
                return false;
            }

            if (string.IsNullOrEmpty(xml))
            {
                error = "WordOpenXML 空";
                return false;
            }

            return true;
        }

        public static string TryReadFullName(object document)
        {
            try
            {
                Word.Document word = document as Word.Document;
                if (word != null)
                {
                    return word.FullName ?? "";
                }

                return Convert.ToString(WpsCom.GetProperty(document, "FullName") ?? "");
            }
            catch
            {
                return "";
            }
        }

        public static bool TrySaveUserDoc(object document, out string error)
        {
            error = null;
            try
            {
                Word.Document word = document as Word.Document;
                if (word != null)
                {
                    word.Save();
                    return true;
                }

                WpsCom.Invoke(document, "Save");
                return true;
            }
            catch (Exception ex)
            {
                error = "Save 用户稿失败: " + ex.Message;
                return false;
            }
        }

        public static string NewTempDocx(string tag)
        {
            string safe = string.IsNullOrEmpty(tag) ? "chart" : tag;
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                safe = safe.Replace(c, '_');
            }

            return Path.Combine(
                Path.GetTempPath(),
                "ew-wd-" + safe + "-" + Guid.NewGuid().ToString("N") + ".docx");
        }

        public static void TryDeleteFile(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        public static bool TryOpenCopy(object destDoc, string path, out object copy, out string error)
        {
            copy = null;
            error = null;
            try
            {
                Word.Document word = destDoc as Word.Document;
                if (word != null)
                {
                    copy = word.Application.Documents.Open(path, ReadOnly: true, AddToRecentFiles: false);
                    return copy != null;
                }

                object app = WpsCom.GetProperty(destDoc, "Application");
                object docs = WpsCom.GetProperty(app, "Documents");
                try
                {
                    copy = WpsCom.Invoke(docs, "Open", path, true);
                }
                catch
                {
                    copy = WpsCom.Invoke(docs, "Open", path);
                }

                if (copy == null)
                {
                    error = "打开改包稿失败";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "打开改包稿失败: " + ex.Message;
                return false;
            }
        }

        public static void TryCloseCopy(object copy)
        {
            if (copy == null)
            {
                return;
            }

            try
            {
                Word.Document word = copy as Word.Document;
                if (word != null)
                {
                    word.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    return;
                }

                try
                {
                    WpsCom.Invoke(copy, "Close", false);
                }
                catch
                {
                    WpsCom.Invoke(copy, "Close");
                }
            }
            catch
            {
            }
        }

        public static bool TryCopyChartOnDoc(object copyDoc, int ordinal, out string error)
        {
            error = null;
            if (!TryFindChartShape(copyDoc, ordinal, out object shape, out error))
            {
                return false;
            }

            try
            {
                Word.InlineShape inline = shape as Word.InlineShape;
                if (inline != null)
                {
                    inline.Range.Copy();
                    return true;
                }

                Word.Shape floating = shape as Word.Shape;
                if (floating != null)
                {
                    floating.Select();
                    object app = WpsCom.GetProperty(copyDoc, "Application")
                        ?? (copyDoc as Word.Document)?.Application;
                    object selection = WpsCom.GetProperty(app, "Selection");
                    WpsCom.Invoke(selection, "Copy");
                    return true;
                }

                try
                {
                    object range = WpsCom.GetProperty(shape, "Range");
                    if (range != null)
                    {
                        WpsCom.Invoke(range, "Copy");
                        return true;
                    }
                }
                catch
                {
                }

                WpsCom.Invoke(shape, "Copy");
                return true;
            }
            catch (Exception ex)
            {
                error = "改包稿 Copy 失败: " + ex.Message;
                return false;
            }
        }

        public static bool TryDeleteChartOnDoc(object destDoc, int ordinal, out string error)
        {
            if (!TryFindChartShape(destDoc, ordinal, out object shape, out error))
            {
                return false;
            }

            return TryDeleteShape(shape, out error);
        }

        /// <summary>删掉第 ordinal 张后，把剪贴板里的图贴回同一位置。甲乙丙改甲仍是新甲乙丙。</summary>
        public static bool TryReplaceChartAtOrdinal(object destDoc, int ordinal, out string error)
        {
            if (!TryFindChartShape(destDoc, ordinal, out object shape, out error))
            {
                return false;
            }

            if (!TryReadInsertStart(shape, out int start))
            {
                error = "原图位置读不到";
                return false;
            }

            if (!TryDeleteShape(shape, out error))
            {
                return false;
            }

            return TryPasteAt(destDoc, start, out error);
        }

        private static bool TryDeleteShape(object shape, out string error)
        {
            error = null;
            try
            {
                Word.InlineShape inline = shape as Word.InlineShape;
                if (inline != null)
                {
                    inline.Delete();
                    return true;
                }

                Word.Shape floating = shape as Word.Shape;
                if (floating != null)
                {
                    floating.Delete();
                    return true;
                }

                WpsCom.Invoke(shape, "Delete");
                return true;
            }
            catch (Exception ex)
            {
                error = "删目标图失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryReadInsertStart(object shape, out int start)
        {
            start = 0;
            try
            {
                Word.InlineShape inline = shape as Word.InlineShape;
                if (inline != null)
                {
                    start = inline.Range.Start;
                    return true;
                }

                Word.Shape floating = shape as Word.Shape;
                if (floating != null && floating.Anchor != null)
                {
                    start = floating.Anchor.Start;
                    return true;
                }
            }
            catch
            {
            }

            object range = WpsCom.GetProperty(shape, "Range");
            if (range == null)
            {
                range = WpsCom.GetProperty(shape, "Anchor");
            }

            if (range == null)
            {
                return false;
            }

            object raw = WpsCom.GetProperty(range, "Start");
            if (raw == null)
            {
                return false;
            }

            start = Convert.ToInt32(raw);
            return true;
        }

        private static bool TryPasteAt(object destDoc, int start, out string error)
        {
            error = null;
            try
            {
                Word.Document word = destDoc as Word.Document;
                if (word != null && !WpsChartExtract.IsWpsDocument(destDoc))
                {
                    Word.Range hole = word.Range(start, start);
                    if (HostCallbacks.ShouldStealForeground())
                    {
                        word.Activate();
                        hole.Select();
                        word.Application.Selection.Paste();
                    }
                    else
                    {
                        hole.Paste();
                    }

                    return true;
                }

                if (HostCallbacks.ShouldStealForeground())
                {
                    try
                    {
                        WpsCom.Invoke(destDoc, "Activate");
                    }
                    catch
                    {
                    }
                }

                object wpsHole = WpsCom.Invoke(destDoc, "Range", start, start);
                if (HostCallbacks.ShouldStealForeground())
                {
                    WpsCom.Invoke(wpsHole, "Select");
                    object app = WpsCom.GetProperty(destDoc, "Application");
                    object selection = WpsCom.GetProperty(app, "Selection");
                    WpsCom.Invoke(selection, "Paste");
                }
                else
                {
                    WpsCom.Invoke(wpsHole, "Paste");
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "贴回失败: " + ex.Message;
                return false;
            }
        }

        public static bool TryFindChartShape(object document, int ordinal, out object shape, out string error)
        {
            shape = null;
            error = null;
            List<object> shapes = CollectChartShapes(document);
            if (ordinal < 0 || ordinal >= shapes.Count)
            {
                error = "目标图下标无效 " + ordinal + "/" + shapes.Count;
                return false;
            }

            shape = shapes[ordinal];
            return shape != null;
        }

        public static List<object> CollectChartShapes(object document)
        {
            var list = new List<object>();
            if (document == null)
            {
                return list;
            }

            Word.Document word = document as Word.Document;
            if (word != null && !WpsChartExtract.IsWpsDocument(document))
            {
                foreach (ChartResolveHelper.BodyChart slot in ChartResolveHelper.GetBodyChartsInOrder(word))
                {
                    list.Add((object)slot.Inline ?? slot.Floating ?? slot.Raw);
                }

                return list;
            }

            foreach (var item in WpsChartExtract.CollectShapes(document))
            {
                list.Add(item.Shape);
            }

            return list;
        }

        public static bool TryReadShapeTitle(object shape, out string title)
        {
            title = "";
            if (shape == null)
            {
                return false;
            }

            try
            {
                Word.InlineShape inline = shape as Word.InlineShape;
                if (inline != null)
                {
                    title = inline.Title ?? "";
                    return true;
                }

                Word.Shape floating = shape as Word.Shape;
                if (floating != null)
                {
                    title = floating.Title ?? "";
                    return true;
                }

                title = Convert.ToString(WpsCom.GetProperty(shape, "Title") ?? "");
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool TryWriteShapeTitle(object shape, string title)
        {
            if (shape == null)
            {
                return false;
            }

            string want = title ?? "";
            try
            {
                if (TitleEquals(shape, want))
                {
                    return true;
                }

                AssignTitle(shape, want);
                if (TitleEquals(shape, want))
                {
                    return true;
                }

                TrySelectShape(shape);
                AssignTitle(shape, want);
                return TitleEquals(shape, want);
            }
            catch
            {
                return false;
            }
        }

        private static void AssignTitle(object shape, string title)
        {
            try
            {
                Word.InlineShape inline = shape as Word.InlineShape;
                if (inline != null)
                {
                    inline.Title = title ?? "";
                    return;
                }
            }
            catch
            {
            }

            try
            {
                Word.Shape floating = shape as Word.Shape;
                if (floating != null)
                {
                    floating.Title = title ?? "";
                    return;
                }
            }
            catch
            {
            }

            WpsCom.TrySetProperty(shape, "Title", title ?? "");
        }

        private static bool TitleEquals(object shape, string want)
        {
            return TryReadShapeTitle(shape, out string back)
                && string.Equals(back ?? "", want ?? "", StringComparison.Ordinal);
        }

        private static void TrySelectShape(object shape)
        {
            try
            {
                object range = WpsCom.GetProperty(shape, "Range");
                if (range != null)
                {
                    WpsCom.Invoke(range, "Select");
                    return;
                }
            }
            catch
            {
            }

            try
            {
                WpsCom.Invoke(shape, "Select");
            }
            catch
            {
            }
        }
    }
}
