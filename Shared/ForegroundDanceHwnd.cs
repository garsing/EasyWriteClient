using System;
using System.Collections.Generic;
using WordAddIn1.BrowserHost;
using WordAddIn1.DocumentHost;
using WordAddIn1.OpenFiles;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从渠道解析目标窗 hwnd，供前台编排第 2 步。失败返回 0（跳过第 2 步）。
    /// </summary>
    internal static class ForegroundDanceHwnd
    {
        public static int TryResolveFromArgs(IDictionary<string, object> args)
        {
            if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out _))
            {
                return 0;
            }

            return TryResolve(channel);
        }

        public static int TryResolve(string channelId)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return 0;
            }

            if (!ChannelRegistry.TryGet(channelId, out IOperationChannel channel) || channel == null)
            {
                return 0;
            }

            return TryResolve(channel);
        }

        public static int TryResolve(IOperationChannel channel)
        {
            if (channel == null)
            {
                return 0;
            }

            try
            {
                switch (channel.Kind)
                {
                    case ChannelKind.Word:
                        return TryWord((WordChannel)channel);
                    case ChannelKind.Wps:
                        return TryWps((WpsChannel)channel);
                    case ChannelKind.Excel:
                        return TryExcel((ExcelChannel)channel);
                    case ChannelKind.Et:
                        return TryEt((EtChannel)channel);
                    case ChannelKind.Ppt:
                        return TryPpt((PptChannel)channel);
                    case ChannelKind.Wpp:
                        return TryWpp((WppChannel)channel);
                    case ChannelKind.Browser:
                        return TryBrowser((BrowserChannel)channel);
                    default:
                        return 0;
                }
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static int TryWord(WordChannel channel)
        {
            if (channel == null || !channel.TryGetLiveDocument(out Word.Document document))
            {
                return 0;
            }

            return TryWordDocument(document);
        }

        private static int TryWps(WpsChannel channel)
        {
            if (channel == null || !channel.TryGetLiveDocument(out object wpsDoc))
            {
                return 0;
            }

            if (DocumentHostCompat.TryAsWordDocument(wpsDoc, out Word.Document wordDoc))
            {
                return TryWordDocument(wordDoc);
            }

            return TryComApplicationHwnd(wpsDoc);
        }

        private static int TryExcel(ExcelChannel channel)
        {
            if (channel == null || !channel.TryGetLiveWorkbook(out Excel.Workbook workbook))
            {
                return 0;
            }

            try
            {
                return workbook.Application.Hwnd;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static int TryEt(EtChannel channel)
        {
            if (channel == null || !channel.TryGetLiveWorkbook(out object workbook))
            {
                return 0;
            }

            return TryComApplicationHwnd(workbook);
        }

        private static int TryPpt(PptChannel channel)
        {
            if (channel == null || !channel.TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                return 0;
            }

            try
            {
                return presentation.Application.HWND;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static int TryWpp(WppChannel channel)
        {
            if (channel == null || !channel.TryGetLivePresentation(out object presentation))
            {
                return 0;
            }

            return TryComApplicationHwnd(presentation);
        }

        private static int TryBrowser(BrowserChannel channel)
        {
            if (channel == null)
            {
                return 0;
            }

            if (string.Equals(channel.Track, "attach", StringComparison.OrdinalIgnoreCase))
            {
                return 0;
            }

            return YiWriteBrowserHost.TryGetVisibleFormHandle();
        }

        private static int TryWordDocument(Word.Document document)
        {
            if (document == null)
            {
                return 0;
            }

            try
            {
                int hwnd = document.ActiveWindow.Hwnd;
                if (hwnd != 0)
                {
                    return hwnd;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                int hwnd = document.Application.ActiveWindow.Hwnd;
                if (hwnd != 0)
                {
                    return hwnd;
                }
            }
            catch (Exception)
            {
            }

            return 0;
        }

        private static int TryComApplicationHwnd(object comObject)
        {
            if (comObject == null)
            {
                return 0;
            }

            object app = null;
            try
            {
                app = EtCom.GetProperty(comObject, "Application")
                    ?? WppCom.GetProperty(comObject, "Application");
            }
            catch (Exception)
            {
            }

            if (EtCom.TryGetHwnd(app, out int hwnd) && hwnd != 0)
            {
                return hwnd;
            }

            if (WppCom.TryGetHwnd(app, out hwnd) && hwnd != 0)
            {
                return hwnd;
            }

            if (EtCom.TryGetHwnd(comObject, out hwnd) && hwnd != 0)
            {
                return hwnd;
            }

            return 0;
        }
    }
}
