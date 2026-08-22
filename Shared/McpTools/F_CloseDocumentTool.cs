using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1
{
    public static class F_CloseDocumentTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_close_document"] = args =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (ChannelHost.Kind != ChannelHostKind.Desktop)
                    {
                        return Task.FromResult(Fail("仅易写桌面应用支持"));
                    }

                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return Task.FromResult(Fail(resolveError ?? "没有可关闭的文档渠道"));
                    }

                    if (!TryCloseChannel(channel, out string path, out string kind, out string error))
                    {
                        return Task.FromResult(Fail(error));
                    }

                    ChannelRegistry.Remove(channel.ChannelId);
                    return Task.FromResult(new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            closed = true,
                            path,
                            channel_id = ChannelRegistry.ToPublicId(channel.ChannelId),
                            kind
                        }
                    });
                }
                catch (OperationCanceledException)
                {
                    return Task.FromResult(Fail("cancelled by user"));
                }
                catch (Exception ex)
                {
                    return Task.FromResult(Fail("关闭文档失败: " + ex.Message));
                }
            };
        }

        private static bool TryCloseChannel(
            IOperationChannel channel,
            out string path,
            out string kind,
            out string error)
        {
            path = null;
            kind = null;
            error = null;

            switch (channel.Kind)
            {
                case ChannelKind.Word:
                    kind = "word";
                    return CloseWord((WordChannel)channel, out path, out error);
                case ChannelKind.Wps:
                    kind = "wps";
                    return CloseLateBound(((WpsChannel)channel).TryGetLiveDocument, ((WpsChannel)channel).FilePath, "Close", out path, out error);
                case ChannelKind.Excel:
                    kind = "excel";
                    return CloseExcel((ExcelChannel)channel, out path, out error);
                case ChannelKind.Et:
                    kind = "et";
                    return CloseLateBound(((EtChannel)channel).TryGetLiveWorkbook, ((EtChannel)channel).FilePath, "Close", out path, out error);
                case ChannelKind.Ppt:
                    kind = "ppt";
                    return ClosePpt((PptChannel)channel, out path, out error);
                case ChannelKind.Wpp:
                    kind = "wpp";
                    return CloseLateBound(((WppChannel)channel).TryGetLivePresentation, ((WppChannel)channel).FilePath, "Close", out path, out error);
                default:
                    error = "只能关闭文字/表格/演示渠道，请使用对应专用工具";
                    return false;
            }
        }

        private static bool CloseWord(WordChannel channel, out string path, out string error)
        {
            path = channel.FilePath;
            error = null;
            if (!channel.TryGetLiveDocument(out Word.Document doc))
            {
                error = "文档已关闭";
                return false;
            }

            string disk = TryDiskPath(doc.Path, doc.FullName);
            if (string.IsNullOrEmpty(disk))
            {
                error = "文档尚未保存到磁盘，请先另存或使用 F_open_document 新建";
                return false;
            }

            path = disk;
            doc.Save();
            object save = Word.WdSaveOptions.wdDoNotSaveChanges;
            doc.Close(ref save);
            return true;
        }

        private static bool CloseExcel(ExcelChannel channel, out string path, out string error)
        {
            path = channel.FilePath;
            error = null;
            if (!channel.TryGetLiveWorkbook(out Excel.Workbook wb))
            {
                error = "工作簿已关闭";
                return false;
            }

            string disk = TryDiskPath(wb.Path, wb.FullName);
            if (string.IsNullOrEmpty(disk))
            {
                error = "工作簿尚未保存到磁盘，请先另存或使用 F_open_document 新建";
                return false;
            }

            path = disk;
            wb.Save();
            wb.Close(false);
            return true;
        }

        private static bool ClosePpt(PptChannel channel, out string path, out string error)
        {
            path = channel.FilePath;
            error = null;
            if (!channel.TryGetLivePresentation(out PowerPoint.Presentation pres))
            {
                error = "演示文稿已关闭";
                return false;
            }

            string disk = TryDiskPath(pres.Path, pres.FullName);
            if (string.IsNullOrEmpty(disk))
            {
                error = "演示文稿尚未保存到磁盘，请先另存或使用 F_open_document 新建";
                return false;
            }

            path = disk;
            pres.Save();
            pres.Close();
            return true;
        }

        private delegate bool TryGetLive(out object live);

        private static bool CloseLateBound(
            TryGetLive tryGet,
            string filePath,
            string closeMethod,
            out string path,
            out string error)
        {
            path = filePath;
            error = null;
            if (!tryGet(out object live) || live == null)
            {
                error = "文档已关闭";
                return false;
            }

            dynamic d = live;
            string diskPath = "";
            string fullName = "";
            try { diskPath = d.Path as string ?? ""; } catch { }
            try { fullName = d.FullName as string ?? ""; } catch { }
            string disk = TryDiskPath(diskPath, fullName);
            if (string.IsNullOrEmpty(disk))
            {
                error = "文档尚未保存到磁盘，请先另存或使用 F_open_document 新建";
                return false;
            }

            path = disk;
            d.Save();
            d.Close();
            return true;
        }

        private static string TryDiskPath(string path, string fullName)
        {
            if (!string.IsNullOrWhiteSpace(path)
                && !string.IsNullOrWhiteSpace(fullName)
                && fullName.IndexOfAny(new[] { '\\', '/' }) >= 0)
            {
                return fullName;
            }

            return null;
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
