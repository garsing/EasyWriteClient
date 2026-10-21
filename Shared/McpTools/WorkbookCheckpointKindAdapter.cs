using System;
using System.Diagnostics;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1
{
    public sealed class WorkbookCheckpointKindAdapter : ICheckpointKindAdapter
    {
        public string CopyExtension => "xlsx";

        public CheckpointCaret CaptureCaret(CheckpointLiveTarget target)
        {
            var caret = new CheckpointCaret();
            object book = target?.Live;
            if (book == null)
            {
                return caret;
            }

            try
            {
                object sheet = InvokeGet(book, "ActiveSheet");
                caret.SheetName = InvokeGet(sheet, "Name") as string ?? "";
                object app = InvokeGet(book, "Application");
                object cell = InvokeGet(app, "ActiveCell");
                caret.CellAddress = InvokeGet(cell, "Address") as string ?? "A1";
                caret.Captured = !string.IsNullOrEmpty(caret.SheetName);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[Checkpoint] Workbook CaptureCaret: " + ex.Message);
            }

            return caret;
        }

        public void PrepareForSave(CheckpointLiveTarget target)
        {
        }

        public void SaveCopy(CheckpointLiveTarget target, string destinationPath)
        {
            object book = target?.Live ?? throw new InvalidOperationException("没有活动的工作簿");
            try
            {
                Invoke(book, "SaveCopyAs", destinationPath);
            }
            catch (Exception)
            {
                Invoke(book, "SaveCopyAs", new object[] { destinationPath });
            }

            if (!System.IO.File.Exists(destinationPath) || new System.IO.FileInfo(destinationPath).Length == 0)
            {
                throw new InvalidOperationException("无法创建工作簿副本。请先保存文档后重试。");
            }
        }

        public void RestoreCopy(CheckpointLiveTarget target, string checkpointPath)
        {
            object live = target?.Live ?? throw new InvalidOperationException("没有活动的工作簿");
            Excel.Workbook typed = live as Excel.Workbook;
            if (typed != null)
            {
                RestoreCopyTyped(typed, checkpointPath);
                return;
            }

            object app = InvokeGet(live, "Application");
            object alerts = null;
            try
            {
                alerts = InvokeGet(app, "DisplayAlerts");
                InvokeSet(app, "DisplayAlerts", false);
                object books = InvokeGet(app, "Workbooks");
                object source = OpenCheckpoint(books, checkpointPath);
                try
                {
                    ReplaceSheets(source, live);
                }
                finally
                {
                    Invoke(source, "Close", false);
                }
            }
            finally
            {
                if (alerts != null)
                {
                    try
                    {
                        InvokeSet(app, "DisplayAlerts", alerts);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static void RestoreCopyTyped(Excel.Workbook live, string checkpointPath)
        {
            Excel.Application app = live.Application;
            bool alerts = true;
            try
            {
                alerts = app.DisplayAlerts;
            }
            catch (Exception)
            {
            }

            try
            {
                app.DisplayAlerts = false;
                Excel.Workbook source = app.Workbooks.Open(checkpointPath, UpdateLinks: 0, ReadOnly: true);
                try
                {
                    int oldCount = live.Worksheets.Count;
                    foreach (Excel.Worksheet srcSheet in source.Worksheets)
                    {
                        Excel.Worksheet after = (Excel.Worksheet)live.Worksheets[live.Worksheets.Count];
                        srcSheet.Copy(After: after);
                    }

                    for (int i = 0; i < oldCount; i++)
                    {
                        ((Excel.Worksheet)live.Worksheets[1]).Delete();
                    }
                }
                finally
                {
                    source.Close(false);
                }
            }
            finally
            {
                try
                {
                    app.DisplayAlerts = alerts;
                }
                catch (Exception)
                {
                }
            }
        }

        public void RestoreCaret(CheckpointLiveTarget target, CheckpointCaret caret)
        {
            if (caret == null || !caret.Captured || string.IsNullOrEmpty(caret.SheetName))
            {
                return;
            }

            object book = target?.Live;
            if (book == null)
            {
                return;
            }

            try
            {
                object sheets = InvokeGet(book, "Worksheets");
                object sheet = InvokeGet(sheets, "Item", caret.SheetName);
                Invoke(sheet, "Activate");
                string address = string.IsNullOrEmpty(caret.CellAddress) ? "A1" : caret.CellAddress;
                object range = InvokeGet(sheet, "Range", address);
                Invoke(range, "Select");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[Checkpoint] Workbook RestoreCaret: " + ex.Message);
            }
        }

        private static object OpenCheckpoint(object books, string checkpointPath)
        {
            try
            {
                return Invoke(books, "Open", checkpointPath, 0, true);
            }
            catch (Exception)
            {
                return Invoke(books, "Open", checkpointPath);
            }
        }

        private static void ReplaceSheets(object source, object target)
        {
            object srcSheets = InvokeGet(source, "Worksheets");
            object dstSheets = InvokeGet(target, "Worksheets");
            int oldCount = Convert.ToInt32(InvokeGet(dstSheets, "Count"));
            int srcCount = Convert.ToInt32(InvokeGet(srcSheets, "Count"));
            for (int i = 1; i <= srcCount; i++)
            {
                object srcSheet = InvokeGet(srcSheets, "Item", i);
                object after = InvokeGet(dstSheets, "Item", Convert.ToInt32(InvokeGet(dstSheets, "Count")));
                CopySheetAfter(srcSheet, after);
            }

            for (int i = 0; i < oldCount; i++)
            {
                object first = InvokeGet(dstSheets, "Item", 1);
                Invoke(first, "Delete");
            }
        }

        private static void CopySheetAfter(object srcSheet, object after)
        {
            dynamic src = srcSheet;
            src.Copy(Type.Missing, after);
        }

        private static object InvokeGet(object target, string name, params object[] args)
        {
            if (target == null)
            {
                return null;
            }

            if (args == null || args.Length == 0)
            {
                return target.GetType().InvokeMember(
                    name,
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    target,
                    null);
            }

            return target.GetType().InvokeMember(
                name,
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                null,
                target,
                args);
        }

        private static void InvokeSet(object target, string name, object value)
        {
            target.GetType().InvokeMember(
                name,
                BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
                null,
                target,
                new[] { value });
        }

        private static object Invoke(object target, string name, params object[] args)
        {
            try
            {
                return target.GetType().InvokeMember(
                    name,
                    BindingFlags.InvokeMethod
                        | BindingFlags.OptionalParamBinding
                        | BindingFlags.Public
                        | BindingFlags.Instance,
                    null,
                    target,
                    args,
                    null,
                    null,
                    null);
            }
            catch (TargetInvocationException ex)
            {
                throw new InvalidOperationException(
                    ex.InnerException == null ? ex.Message : ex.InnerException.Message,
                    ex);
            }
        }

        private static object Invoke(object target, string collection, string method, params object[] args)
        {
            object col = InvokeGet(target, collection);
            return Invoke(col, method, args);
        }
    }
}
