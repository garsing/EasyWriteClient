using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 解析 Word.Application：优先宿主传入，其次 GetActiveObject，可选 new（I5）。
    /// </summary>
    public static class WordApplicationResolver
    {
        /// <param name="createIfMissing">为 true 时，无运行中 Word 则创建（Desktop 打开文档场景）。</param>
        public static bool TryResolve(
            object wordApplication,
            out Word.Application application,
            out string error,
            bool createIfMissing = false)
        {
            if (OfficeStaScheduler.ShouldHop)
            {
                Word.Application app = null;
                string err = null;
                bool ok = OfficeStaScheduler.Invoke(() =>
                    TryResolve(wordApplication, out app, out err, createIfMissing));
                application = app;
                error = err;
                return ok;
            }

            application = null;
            error = null;

            if (wordApplication is Word.Application typed)
            {
                application = typed;
                return true;
            }

            try
            {
                application = (Word.Application)Marshal.GetActiveObject("Word.Application");
                if (application != null)
                {
                    return true;
                }
            }
            catch (COMException)
            {
                // 无运行中的 Word
            }
            catch (Exception ex)
            {
                error = "附着已有 Word 失败: " + ex.Message;
                return false;
            }

            if (!createIfMissing)
            {
                error = "Word 应用程序不可用；请先启动 Word 或通过宿主注入 Application。";
                return false;
            }

            try
            {
                application = new Word.Application { Visible = true };
                return true;
            }
            catch (Exception ex)
            {
                error = "无法创建 Word.Application（是否未安装 Word？）: " + ex.Message;
                application = null;
                return false;
            }
        }
    }
}
