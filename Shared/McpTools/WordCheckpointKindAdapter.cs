using System;
using System.Diagnostics;
using WordAddIn1.DocumentHost;
using WordAddIn1.OpenFiles;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class WordCheckpointKindAdapter : ICheckpointKindAdapter
    {
        public string CopyExtension => "docx";

        public CheckpointCaret CaptureCaret(CheckpointLiveTarget target)
        {
            var caret = new CheckpointCaret();
            Word.Document doc = target?.Live as Word.Document;
            if (doc == null)
            {
                return caret;
            }

            try
            {
                Word.Window window = null;
                try
                {
                    window = doc.ActiveWindow;
                }
                catch (Exception)
                {
                }

                Word.Selection selection = null;
                if (window != null)
                {
                    try
                    {
                        selection = window.Selection;
                    }
                    catch (Exception)
                    {
                    }
                }

                if (selection == null)
                {
                    try
                    {
                        selection = doc.Application?.Selection;
                    }
                    catch (Exception)
                    {
                    }
                }

                if (selection != null)
                {
                    caret.SelectionStart = selection.Start;
                    caret.SelectionEnd = selection.End;
                    caret.Captured = true;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[Checkpoint] Word CaptureCaret: " + ex.Message);
            }

            return caret;
        }

        public void PrepareForSave(CheckpointLiveTarget target)
        {
            Word.Document doc = target?.Live as Word.Document;
            if (doc == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(target.DocUuid))
            {
                DocumentIdentity.Register(doc, target.DocUuid);
            }

            TryActivateInteropDocument(doc);
            DocumentState.BindAndActivate(doc);
            if (!string.IsNullOrEmpty(target.DocUuid))
            {
                DocumentState.ActivateSessionByUuid(target.DocUuid);
            }
        }

        public void SaveCopy(CheckpointLiveTarget target, string destinationPath)
        {
            Word.Document doc = target?.Live as Word.Document
                ?? throw new InvalidOperationException("没有活动的 Word 文档");
            DocumentCheckpointSaveHelper.SaveDocumentCopy(doc, destinationPath);
        }

        public void RestoreCopy(CheckpointLiveTarget target, string checkpointPath)
        {
            Word.Document doc = target?.Live as Word.Document
                ?? throw new InvalidOperationException("没有活动的 Word 文档");
            Word.Application app = TryApplication(doc, target.ChannelId);
            if (app == null)
            {
                throw new InvalidOperationException("无法取得 Word Application");
            }

            DocumentRestoreHelper.RestoreActiveDocumentFromCheckpoint(checkpointPath, app, doc);
        }

        public void RestoreCaret(CheckpointLiveTarget target, CheckpointCaret caret)
        {
            if (caret == null || !caret.Captured)
            {
                return;
            }

            Word.Document doc = target?.Live as Word.Document;
            if (doc == null)
            {
                return;
            }

            try
            {
                int docEnd = doc.Content?.End ?? 0;
                if (docEnd <= 1)
                {
                    return;
                }

                int maxPos = Math.Max(0, docEnd - 1);
                int start = Math.Max(0, Math.Min(caret.SelectionStart, maxPos));
                int end = Math.Max(start, Math.Min(caret.SelectionEnd, maxPos));

                if (HostCallbacks.ShouldStealForeground())
                {
                    try
                    {
                        doc.Activate();
                    }
                    catch (Exception)
                    {
                    }
                }

                Word.Application app = null;
                try
                {
                    app = doc.Application;
                    if (app?.Selection != null)
                    {
                        app.Selection.SetRange(start, end);
                    }
                }
                catch (Exception)
                {
                    try
                    {
                        doc.Range(start, end).Select();
                    }
                    catch (Exception)
                    {
                    }
                }

                Word.Range scrollRange = doc.Range(start, start);
                WordRangeFinder.ScrollRangeIntoView(app ?? doc.Application, scrollRange, "checkpoint-caret");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[Checkpoint] Word RestoreCaret: " + ex.Message);
            }
        }

        private static Word.Application TryApplication(Word.Document target, string channelId)
        {
            if (!string.IsNullOrEmpty(channelId)
                && ChannelRegistry.TryGet(channelId, out IOperationChannel channel)
                && channel is WpsChannel wps
                && wps.TryGetLiveDocument(out object wpsDoc))
            {
                object appObj = WpsCom.GetProperty(wpsDoc, "Application");
                if (appObj is Word.Application late)
                {
                    return late;
                }

                try
                {
                    return (Word.Application)appObj;
                }
                catch (Exception)
                {
                }
            }

            try
            {
                return target?.Application;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void TryActivateInteropDocument(Word.Document document)
        {
            if (document == null || !HostCallbacks.ShouldStealForeground())
            {
                return;
            }

            try
            {
                document.Activate();
            }
            catch (Exception)
            {
            }
        }
    }
}
