using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1
{
    public sealed class PresentationCheckpointKindAdapter : ICheckpointKindAdapter
    {
        public string CopyExtension => "pptx";

        public CheckpointCaret CaptureCaret(CheckpointLiveTarget target)
        {
            var caret = new CheckpointCaret();
            object pres = target?.Live;
            if (pres == null)
            {
                return caret;
            }

            try
            {
                object app = InvokeGet(pres, "Application");
                object view = InvokeGet(InvokeGet(app, "ActiveWindow"), "View");
                object slide = InvokeGet(view, "Slide");
                caret.SlideIndex = Convert.ToInt32(InvokeGet(slide, "SlideIndex"));
                try
                {
                    object sel = InvokeGet(InvokeGet(app, "ActiveWindow"), "Selection");
                    object shapeRange = InvokeGet(sel, "ShapeRange");
                    if (Convert.ToInt32(InvokeGet(shapeRange, "Count")) >= 1)
                    {
                        object shape = InvokeGet(shapeRange, "Item", 1);
                        caret.ShapeId = Convert.ToInt32(InvokeGet(shape, "Id"));
                    }
                }
                catch (Exception)
                {
                }

                caret.Captured = caret.SlideIndex >= 1;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[Checkpoint] Presentation CaptureCaret: " + ex.Message);
            }

            return caret;
        }

        public void PrepareForSave(CheckpointLiveTarget target)
        {
        }

        public void SaveCopy(CheckpointLiveTarget target, string destinationPath)
        {
            object pres = target?.Live ?? throw new InvalidOperationException("没有活动的演示文稿");
            PowerPoint.Presentation typed = pres as PowerPoint.Presentation;
            if (TrySaveViaFileCopy(typed, pres, destinationPath))
            {
                return;
            }

            string originalPath = typed != null
                ? ReadFullName(typed)
                : Convert.ToString(InvokeGet(pres, "FullName") ?? "");
            if (typed != null)
            {
                SaveCopyTyped(typed, destinationPath);
            }
            else
            {
                SaveCopyLate(pres, destinationPath);
            }

            if (!File.Exists(destinationPath) || new FileInfo(destinationPath).Length == 0)
            {
                throw new InvalidOperationException("无法创建演示文稿副本。请先保存文档后重试。");
            }

            if (typed != null)
            {
                CloseOpenedCopy(typed.Application, destinationPath, typed);
                RestoreLivePath(typed, originalPath);
            }
            else
            {
                CloseOpenedCopyLate(InvokeGet(pres, "Application"), destinationPath, pres);
                RestoreLivePathLate(pres, originalPath);
            }

            RebindChannel(target);
        }

        /// <summary>
        /// 已落盘就就地 Save 再拷文件，不走 SaveCopyAs。
        /// PPT 的 SaveCopyAs 会改活稿路径或另开一份，随后加页打在空稿上（slide_id=256、页数仍是 1）。
        /// </summary>
        private static bool TrySaveViaFileCopy(
            PowerPoint.Presentation typed,
            object late,
            string destinationPath)
        {
            string originalPath = typed != null ? ReadFullName(typed) : Convert.ToString(InvokeGet(late, "FullName") ?? "");
            if (string.IsNullOrEmpty(originalPath)
                || originalPath.StartsWith("Unsaved", StringComparison.OrdinalIgnoreCase)
                || !File.Exists(originalPath))
            {
                return false;
            }

            try
            {
                try
                {
                    if (typed != null)
                    {
                        typed.Save();
                    }
                    else
                    {
                        Invoke(late, "Save");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[Checkpoint] PPT in-place Save before copy skipped: " + ex.Message);
                }

                CopyFileShared(originalPath, destinationPath);
                if (!File.Exists(destinationPath) || new FileInfo(destinationPath).Length == 0)
                {
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[Checkpoint] PPT File.Copy failed: " + ex.Message);
                return false;
            }
        }

        private static void CopyFileShared(string sourcePath, string destinationPath)
        {
            string dir = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }

            using (var src = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var dst = new FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                src.CopyTo(dst);
            }
        }

        private static void SaveCopyTyped(PowerPoint.Presentation live, string destinationPath)
        {
            live.SaveCopyAs(destinationPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
        }

        private static void SaveCopyLate(object live, string destinationPath)
        {
            try
            {
                Invoke(live, "SaveCopyAs", new object[] { destinationPath, 24 });
            }
            catch (Exception first)
            {
                try
                {
                    Invoke(live, "SaveCopyAs", destinationPath);
                }
                catch (Exception)
                {
                    throw new InvalidOperationException(
                        "无法创建演示文稿副本。请先保存文档后重试。 " + first.Message);
                }
            }
        }

        private static void CloseOpenedCopy(
            PowerPoint.Application app,
            string copyPath,
            PowerPoint.Presentation keep)
        {
            if (app == null || string.IsNullOrEmpty(copyPath))
            {
                return;
            }

            var toClose = new System.Collections.Generic.List<PowerPoint.Presentation>();
            foreach (PowerPoint.Presentation one in app.Presentations)
            {
                if (ReferenceEquals(one, keep))
                {
                    continue;
                }

                if (SamePath(ReadFullName(one), copyPath))
                {
                    toClose.Add(one);
                }
            }

            foreach (PowerPoint.Presentation one in toClose)
            {
                try
                {
                    one.Close();
                }
                catch (Exception)
                {
                }
            }
        }

        private static void CloseOpenedCopyLate(object app, string copyPath, object keep)
        {
            if (app == null || string.IsNullOrEmpty(copyPath))
            {
                return;
            }

            object col = InvokeGet(app, "Presentations");
            int n = Convert.ToInt32(InvokeGet(col, "Count"));
            for (int i = n; i >= 1; i--)
            {
                object one = InvokeGet(col, "Item", i);
                if (ReferenceEquals(one, keep))
                {
                    continue;
                }

                string full = Convert.ToString(InvokeGet(one, "FullName") ?? "");
                if (SamePath(full, copyPath))
                {
                    try
                    {
                        Invoke(one, "Close");
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static void RestoreLivePath(PowerPoint.Presentation live, string originalPath)
        {
            if (live == null || string.IsNullOrEmpty(originalPath))
            {
                return;
            }

            string now = ReadFullName(live);
            if (string.Equals(now, originalPath, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            live.SaveAs(originalPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
        }

        private static void RestoreLivePathLate(object live, string originalPath)
        {
            if (live == null || string.IsNullOrEmpty(originalPath))
            {
                return;
            }

            string now = Convert.ToString(InvokeGet(live, "FullName") ?? "");
            if (string.Equals(now, originalPath, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            try
            {
                Invoke(live, "SaveAs", new object[] { originalPath, 24 });
            }
            catch (Exception)
            {
                Invoke(live, "SaveAs", originalPath);
            }
        }

        private static void RebindChannel(CheckpointLiveTarget target)
        {
            if (target?.Channel is PptChannel ppt && target.Live is PowerPoint.Presentation live)
            {
                try
                {
                    var _ = live.Name;
                    ppt.UpdatePresentation(live, ReadFullName(live));
                }
                catch (Exception)
                {
                }
            }
        }

        private static string ReadFullName(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return "";
            }

            try
            {
                return presentation.FullName ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool SamePath(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
            {
                return false;
            }

            try
            {
                return string.Equals(
                    System.IO.Path.GetFullPath(a).TrimEnd('\\', '/'),
                    System.IO.Path.GetFullPath(b).TrimEnd('\\', '/'),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }

        public void RestoreCopy(CheckpointLiveTarget target, string checkpointPath)
        {
            object live = target?.Live ?? throw new InvalidOperationException("没有活动的演示文稿");
            PowerPoint.Presentation typed = live as PowerPoint.Presentation;
            if (typed != null)
            {
                RestoreCopyTyped(typed, checkpointPath);
                return;
            }

            RestoreCopyLate(live, checkpointPath);
        }

        private static void RestoreCopyTyped(PowerPoint.Presentation live, string checkpointPath)
        {
            PowerPoint.Application app = live.Application;
            PowerPoint.Presentation source = app.Presentations.Open(
                checkpointPath,
                Office.MsoTriState.msoTrue,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoFalse);
            try
            {
                int oldCount = live.Slides.Count;
                int srcCount = source.Slides.Count;
                if (srcCount > 0)
                {
                    var indexes = new object[srcCount];
                    for (int i = 0; i < srcCount; i++)
                    {
                        indexes[i] = i + 1;
                    }

                    source.Slides.Range(indexes).Copy();
                    live.Slides.Paste();
                }

                for (int i = 0; i < oldCount && live.Slides.Count > 1; i++)
                {
                    live.Slides[1].Delete();
                }
            }
            finally
            {
                source.Close();
            }
        }

        private static void RestoreCopyLate(object live, string checkpointPath)
        {
            object app = InvokeGet(live, "Application");
            object source = Invoke(
                InvokeGet(app, "Presentations"),
                "Open",
                checkpointPath,
                -1,
                0,
                0);
            try
            {
                object liveSlides = InvokeGet(live, "Slides");
                object srcSlides = InvokeGet(source, "Slides");
                int oldCount = Convert.ToInt32(InvokeGet(liveSlides, "Count"));
                int srcCount = Convert.ToInt32(InvokeGet(srcSlides, "Count"));
                if (srcCount > 0)
                {
                    var indexes = new object[srcCount];
                    for (int i = 0; i < srcCount; i++)
                    {
                        indexes[i] = i + 1;
                    }

                    object range = InvokeGet(srcSlides, "Range", indexes);
                    Invoke(range, "Copy");
                    Invoke(liveSlides, "Paste");
                }

                for (int i = 0; i < oldCount && Convert.ToInt32(InvokeGet(liveSlides, "Count")) > 1; i++)
                {
                    Invoke(InvokeGet(liveSlides, "Item", 1), "Delete");
                }
            }
            finally
            {
                Invoke(source, "Close");
            }
        }

        public void RestoreCaret(CheckpointLiveTarget target, CheckpointCaret caret)
        {
            if (caret == null || !caret.Captured || caret.SlideIndex < 1)
            {
                return;
            }

            object pres = target?.Live;
            if (pres == null)
            {
                return;
            }

            try
            {
                object slides = InvokeGet(pres, "Slides");
                int count = Convert.ToInt32(InvokeGet(slides, "Count"));
                int index = Math.Min(caret.SlideIndex, Math.Max(1, count));
                object slide = InvokeGet(slides, "Item", index);
                Invoke(slide, "Select");
                if (caret.ShapeId > 0)
                {
                    object shapes = InvokeGet(slide, "Shapes");
                    int n = Convert.ToInt32(InvokeGet(shapes, "Count"));
                    for (int i = 1; i <= n; i++)
                    {
                        object shape = InvokeGet(shapes, "Item", i);
                        if (Convert.ToInt32(InvokeGet(shape, "Id")) == caret.ShapeId)
                        {
                            Invoke(shape, "Select");
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[Checkpoint] Presentation RestoreCaret: " + ex.Message);
            }
        }

        private static object InvokeGet(object target, string name, params object[] args)
        {
            if (target == null)
            {
                return null;
            }

            BindingFlags flags = BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance;
            return target.GetType().InvokeMember(name, flags, null, target, args != null && args.Length > 0 ? args : null);
        }

        private static object Invoke(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(
                name,
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null,
                target,
                args);
        }
    }
}
