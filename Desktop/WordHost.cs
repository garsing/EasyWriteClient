using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using WordAddIn1;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// Desktop 侧 Word.Application：默认只附着已运行的 Word，不主动 new（避免启动就弹 Word）。
    /// 仅在工具确实需要且本机无 Word 时才 createIfMissing。
    /// </summary>
    internal static class WordHost
    {
        private static readonly object Gate = new object();
        private static Word.Application _app;

        /// <param name="createIfMissing">
        /// false（默认）：仅 GetActiveObject；没有运行中的 Word 则返回 null。
        /// true：无运行实例时才 new Application（Visible=true）。
        /// </param>
        public static Word.Application GetOrAttach(bool createIfMissing = false)
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

        /// <summary>启动/聊天：不创建 Word；仅附着已有实例。</summary>
        public static object TryGetExistingWordApplication()
        {
            return GetOrAttach(createIfMissing: false);
        }

        /// <summary>工具调用：必要时才创建 Word。</summary>
        public static object GetWordApplicationForTools()
        {
            return GetOrAttach(createIfMissing: true);
        }

        /// <summary>F_open 等路径已拿到 Application 时写回缓存。</summary>
        public static void Attach(object wordApplication)
        {
            if (!(wordApplication is Word.Application app))
            {
                return;
            }

            lock (Gate)
            {
                _app = app;
                try
                {
                    DocumentCheckpointService.SetWordApplication(_app);
                }
                catch (Exception)
                {
                }
            }
        }

        /// <summary>
        /// 关闭 Desktop：清理渠道；不 Quit 用户 Word。
        /// 仅对本进程 new 出的、且无打开文档时尝试 Quit（仍保守：默认不 Quit）。
        /// </summary>
        public static void Shutdown()
        {
            lock (Gate)
            {
                ChannelRegistry.ClearAll();
                // I8 保守：不 Quit（避免误关用户文档）
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
