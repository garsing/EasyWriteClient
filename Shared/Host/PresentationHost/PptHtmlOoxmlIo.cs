using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    /// <summary>
    /// 统一整份 SaveCopyAs：读包不 Open；改包贴回改完 zip 才 Open，只 Copy 目标形状。
    /// </summary>
    internal static partial class PptHtmlOoxmlIo
    {
        public static bool TrySaveCopy(object presentation, string path, out string error)
        {
            error = null;
            if (presentation == null || string.IsNullOrEmpty(path))
            {
                error = "SaveCopy：稿或路径空";
                return false;
            }

            try
            {
                PowerPoint.Presentation ppt = presentation as PowerPoint.Presentation;
                if (ppt != null)
                {
                    ppt.SaveCopyAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                }
                else
                {
                    try
                    {
                        WppCom.Invoke(presentation, "SaveCopyAs", path);
                    }
                    catch (Exception)
                    {
                        WppCom.Invoke(presentation, "SaveCopyAs", path, 24);
                    }
                }
            }
            catch (Exception ex)
            {
                error = "SaveCopy 失败: " + ex.Message;
                return false;
            }

            if (!File.Exists(path))
            {
                error = "SaveCopy 未落盘";
                return false;
            }

            return true;
        }

        public static string TrySaveCopyTemp(object presentation, string tag, out string error)
        {
            string path = NewTempPath(tag);
            if (!TrySaveCopy(presentation, path, out error))
            {
                TryDeleteFile(path);
                return null;
            }

            return path;
        }

        public static bool TryReadSavedPackage(
            object presentation,
            Func<string, bool> read,
            out string error,
            string tag = "读包")
        {
            error = null;
            if (presentation == null || read == null)
            {
                error = tag + "：稿或 read 空";
                return false;
            }

            string path = TrySaveCopyTemp(presentation, tag, out error);
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }

            try
            {
                return read(path);
            }
            catch (Exception ex)
            {
                error = tag + "失败: " + ex.Message;
                return false;
            }
            finally
            {
                TryDeleteFile(path);
            }
        }

        public static bool TryRewriteAndCopyBack(
            object destPres,
            int destSlideIndex,
            Func<ZipArchive, IDictionary<string, byte[]>, bool> rewrite,
            Func<object, int, object> findOnCopy,
            Func<bool> pasteToDest,
            out string error,
            string alreadySavedPath = null,
            string tag = "改包",
            List<string> warnings = null)
        {
            error = null;
            object copyPres = null;
            string path = alreadySavedPath;
            string slimPath = null;
            bool ownPath = string.IsNullOrEmpty(alreadySavedPath);
            try
            {
                if (destPres == null || destSlideIndex < 1 || rewrite == null || findOnCopy == null || pasteToDest == null)
                {
                    error = tag + "：参数空";
                    return false;
                }

                if (ownPath)
                {
                    path = TrySaveCopyTemp(destPres, tag, out error);
                    if (string.IsNullOrEmpty(path))
                    {
                        return false;
                    }

                    Log(warnings, tag + "：副本已落盘 " + path);
                    PingDest(warnings, tag + "/SaveCopy", destPres);
                }

                var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                string openPath = path;
                int openSlideIndex = destSlideIndex;
                string slimErr = null;
                using (ZipArchive zin = ZipFile.OpenRead(path))
                {
                    if (!rewrite(zin, replacements))
                    {
                        error = tag + "：改 zip 失败";
                        return false;
                    }

                    if (TrySlimToSlideAndTheme(path, destSlideIndex, zin, replacements, out slimPath, out slimErr))
                    {
                        openPath = slimPath;
                        openSlideIndex = 1;
                        Log(warnings, tag + "：已瘦成一页+主题 " + slimPath
                            + " fromMB=" + Mb(path) + " toMB=" + Mb(slimPath)
                            + " parts=" + replacements.Count);
                    }
                }

                PingDest(warnings, tag + "/改zip", destPres);

                if (openPath == path)
                {
                    if (!TryApplyReplacements(path, replacements, out error))
                    {
                        if (string.IsNullOrEmpty(error))
                        {
                            error = tag + "：回写整份失败";
                        }

                        return false;
                    }

                    Log(warnings, tag + "：瘦包未用（" + slimErr + "），已回写整份再 Open");
                }

                object app = TryGetApp(destPres);
                PingDest(warnings, tag + "/Open前", destPres);
                copyPres = TryOpenCopy(app, openPath, out error, readOnly: true);
                PingDest(warnings, tag + "/Open瘦包" + (copyPres == null ? "失败" : "成功"), destPres);
                if (copyPres == null && openPath != path)
                {
                    Log(warnings, tag + "：瘦包 Open 失败，回退整份");
                    TryCloseCopy(copyPres);
                    if (!TryApplyReplacements(path, replacements, out error))
                    {
                        if (string.IsNullOrEmpty(error))
                        {
                            error = tag + "：回退整份回写失败";
                        }

                        return false;
                    }

                    copyPres = TryOpenCopy(app, path, out error, readOnly: true);
                    PingDest(warnings, tag + "/Open整份" + (copyPres == null ? "失败" : "成功"), destPres);
                    openSlideIndex = destSlideIndex;
                    openPath = path;
                }

                if (copyPres == null)
                {
                    if (string.IsNullOrEmpty(error))
                    {
                        error = tag + "：打开改包稿失败";
                    }

                    return false;
                }

                Log(warnings, tag + "：已 Open 改包稿 slide=" + openSlideIndex);
                object source = findOnCopy(copyPres, openSlideIndex);
                if (source == null)
                {
                    error = tag + "：改包稿找不到目标形状";
                    return false;
                }

                try
                {
                    PowerPoint.Shape pptShape = source as PowerPoint.Shape;
                    if (pptShape != null)
                    {
                        pptShape.Copy();
                    }
                    else
                    {
                        WppCom.Invoke(source, "Copy");
                    }
                }
                catch (Exception ex)
                {
                    PingDest(warnings, tag + "/Copy失败", destPres);
                    error = tag + "：Copy 失败: " + ex.Message;
                    return false;
                }

                PingDest(warnings, tag + "/Copy后", destPres);
                if (!pasteToDest())
                {
                    PingDest(warnings, tag + "/贴回失败", destPres);
                    error = tag + "：贴回失败";
                    return false;
                }

                PingDest(warnings, tag + "/贴回后", destPres);
                return true;
            }
            catch (Exception ex)
            {
                error = tag + "异常: " + ex.Message;
                return false;
            }
            finally
            {
                TryCloseCopy(copyPres);
                TryDeleteFile(slimPath);
                if (ownPath)
                {
                    TryDeleteFile(path);
                }
            }
        }

        private static string Mb(string path)
        {
            try
            {
                return (new FileInfo(path).Length / 1024.0 / 1024.0).ToString("0.00");
            }
            catch (Exception)
            {
                return "?";
            }
        }

        public static object TryOpenCopy(object app, string path, out string error, bool readOnly = false)
        {
            error = null;
            if (app == null || string.IsNullOrEmpty(path))
            {
                error = "Open 改包稿：app/路径空";
                return null;
            }

            PowerPoint.Application pptApp = app as PowerPoint.Application;
            if (pptApp != null)
            {
                try
                {
                    return pptApp.Presentations.Open(
                        path,
                        readOnly ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue);
                }
                catch (Exception ex)
                {
                    error = "Open 改包稿失败: " + ex.Message;
                    return null;
                }
            }

            object presentations = null;
            try
            {
                presentations = WppCom.GetProperty(app, "Presentations");
            }
            catch (Exception ex)
            {
                error = "Open 改包稿：没有 Presentations: " + ex.Message;
                return null;
            }

            if (readOnly)
            {
                try
                {
                    object opened = WppCom.Invoke(presentations, "Open", path, true);
                    if (opened != null)
                    {
                        return opened;
                    }
                }
                catch (Exception)
                {
                }
            }

            try
            {
                object opened = WppCom.Invoke(presentations, "Open", path);
                if (opened != null)
                {
                    return opened;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object opened = WppCom.Invoke(presentations, "Open", path, false, false, false);
                if (opened != null)
                {
                    return opened;
                }
            }
            catch (Exception ex)
            {
                error = "Open 改包稿失败: " + ex.Message;
            }

            if (error == null)
            {
                error = "Open 改包稿无返回";
            }

            return null;
        }

        public static void TryCloseCopy(object presentation)
        {
            if (presentation == null)
            {
                return;
            }

            try
            {
                PowerPoint.Presentation ppt = presentation as PowerPoint.Presentation;
                if (ppt != null)
                {
                    ppt.Saved = Office.MsoTriState.msoTrue;
                    ppt.Close();
                    return;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                WppCom.TrySetProperty(presentation, "Saved", true);
            }
            catch (Exception)
            {
            }

            try
            {
                WppCom.Invoke(presentation, "Close");
            }
            catch (Exception)
            {
            }
        }

        public static object TryFindShapeById(object presentation, int slideIndex, int shapeId)
        {
            if (presentation == null || slideIndex < 1 || shapeId < 1)
            {
                return null;
            }

            try
            {
                PowerPoint.Presentation ppt = presentation as PowerPoint.Presentation;
                if (ppt != null)
                {
                    PowerPoint.Shapes shapes = ppt.Slides[slideIndex].Shapes;
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        if (shapes[i].Id == shapeId)
                        {
                            return shapes[i];
                        }
                    }

                    return null;
                }

                object slides = WppCom.GetProperty(presentation, "Slides");
                object slide = WppCom.GetIndexed(slides, slideIndex);
                object wppShapes = slide == null ? null : WppCom.GetProperty(slide, "Shapes");
                int count = wppShapes == null ? 0 : Convert.ToInt32(WppCom.GetProperty(wppShapes, "Count") ?? 0);
                for (int i = 1; i <= count; i++)
                {
                    object s = WppCom.GetIndexed(wppShapes, i);
                    if (Convert.ToInt32(WppCom.GetProperty(s, "Id") ?? 0) == shapeId)
                    {
                        return s;
                    }
                }
            }
            catch (Exception)
            {
            }

            return null;
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
            catch (Exception)
            {
            }
        }

        public static string NewTempPath(string tag)
        {
            string safe = string.IsNullOrEmpty(tag) ? "ooxml" : tag;
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                safe = safe.Replace(c, '-');
            }

            return Path.Combine(
                Path.GetTempPath(),
                "ew-" + safe + "-" + Guid.NewGuid().ToString("N") + ".pptx");
        }

        private static object TryGetApp(object presentation)
        {
            if (presentation == null)
            {
                return null;
            }

            try
            {
                PowerPoint.Presentation ppt = presentation as PowerPoint.Presentation;
                if (ppt != null)
                {
                    return ppt.Application;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                return WppCom.GetProperty(presentation, "Application");
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static void PingDest(List<string> warnings, string tag, object destPres)
        {
            string dest = destPres == null ? "pres=null" : "活";
            try
            {
                if (destPres != null)
                {
                    object name = destPres is PowerPoint.Presentation ppt
                        ? ppt.Name
                        : WppCom.GetProperty(destPres, "Name");
                    dest = name == null ? "活(Name空)" : "活(" + name + ")";
                }
            }
            catch (Exception ex)
            {
                dest = "死:" + ex.GetType().Name + ":" + ex.Message;
            }

            object app = TryGetApp(destPres);
            string sess = "app=null";
            int excel = 0;
            try
            {
                excel = Process.GetProcessesByName("EXCEL").Length;
            }
            catch (Exception)
            {
            }

            if (app != null)
            {
                try
                {
                    object presentations = app is PowerPoint.Application pptApp
                        ? pptApp.Presentations
                        : WppCom.GetProperty(app, "Presentations");
                    int n = presentations == null
                        ? -1
                        : Convert.ToInt32(
                            app is PowerPoint.Application
                                ? ((PowerPoint.Application)app).Presentations.Count
                                : WppCom.GetProperty(presentations, "Count") ?? -1);
                    sess = "稿=" + n;
                }
                catch (Exception ex)
                {
                    sess = "稿读失败:" + ex.Message;
                }
            }

            string line = "打点 " + tag + " 原稿=" + dest + " " + sess + " excel=" + excel;
            Log(warnings, line);
            try
            {
                Console.WriteLine("  " + line);
            }
            catch (Exception)
            {
            }
        }

        private static void Log(List<string> warnings, string line)
        {
            if (warnings != null && !string.IsNullOrEmpty(line))
            {
                warnings.Add(line);
            }
        }
    }
}
