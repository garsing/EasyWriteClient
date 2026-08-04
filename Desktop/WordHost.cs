using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using WordAddIn1;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// Desktop 侧 Word.Application 附着/创建（I5 / I8）。B4：按需获取；退出时保守不 Quit。
    /// </summary>
    internal static class WordHost
    {
        private static readonly object Gate = new object();
        private static Word.Application _app;

        public static Word.Application GetOrAttach(bool createIfMissing = true)
        {
            lock (Gate)
            {
                if (_app != null)
                {
                    try
                    {
                        var _ = _app.Name;
                        return _app;
                    }
                    catch (Exception)
                    {
                        _app = null;
                    }
                }

                try
                {
                    _app = (Word.Application)Marshal.GetActiveObject("Word.Application");
                }
                catch (COMException)
                {
                    _app = null;
                }

                if (_app == null && createIfMissing)
                {
                    try
                    {
                        _app = new Word.Application { Visible = true };
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("[WordHost] new Application failed: " + ex.Message);
                        return null;
                    }
                }

                if (_app != null)
                {
                    try
                    {
                        DocumentCheckpointService.SetWordApplication(_app);
                    }
                    catch (Exception)
                    {
                    }
                }

                return _app;
            }
        }

        public static object GetWordApplicationObject()
        {
            return GetOrAttach(createIfMissing: true);
        }

        /// <summary>
        /// 关闭 Desktop：清理渠道引用；不 Quit 用户 Word（I8 保守）。
        /// </summary>
        public static void Shutdown()
        {
            lock (Gate)
            {
                ChannelRegistry.ClearAll();
                if (_app != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(_app);
                    }
                    catch (Exception)
                    {
                    }
                }

                _app = null;
            }
        }
    }
}
