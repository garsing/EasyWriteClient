using System;
using System.Diagnostics;
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
            try
            {
                Invoke(pres, "SaveCopyAs", destinationPath);
            }
            catch (Exception first)
            {
                try
                {
                    Invoke(pres, "SaveCopyAs", new object[] { destinationPath, 24 });
                }
                catch (Exception)
                {
                    throw new InvalidOperationException(
                        "无法创建演示文稿副本。请先保存文档后重试。 " + first.Message);
                }
            }

            if (!System.IO.File.Exists(destinationPath) || new System.IO.FileInfo(destinationPath).Length == 0)
            {
                throw new InvalidOperationException("无法创建演示文稿副本。请先保存文档后重试。");
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
