using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1
{
    /// <summary>
    /// 切换默认操作渠道（任意已登记 channel_id：word/wps/excel/et/ppt/wpp）。
    /// </summary>
    public static class F_SwitchDefaultChannelTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_switch_default_channel"] = async (args) =>
            {
                try
                {
                    string channelId = args != null && args.ContainsKey("channel_id")
                        ? args["channel_id"]?.ToString()?.Trim()
                        : null;
                    if (string.IsNullOrEmpty(channelId))
                    {
                        return new ToolResult { Success = false, Error = "必须提供 channel_id 参数" };
                    }

                    if (!ChannelRegistry.TryGet(channelId, out IOperationChannel channel))
                    {
                        return new ToolResult { Success = false, Error = "未知 channel_id: " + channelId };
                    }

                    if (!TryEnsureLive(channel, out string liveError, out string kind, out string uuid, out string path, out string name))
                    {
                        return new ToolResult { Success = false, Error = liveError };
                    }

                    if (!ChannelRegistry.SetDefault(channelId))
                    {
                        return new ToolResult { Success = false, Error = "设置默认渠道失败: " + channelId };
                    }

                    TryActivate(channel);

                    await Task.CompletedTask;

                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["channel_id"] = channel.ChannelId,
                            ["kind"] = kind,
                            ["doc_uuid"] = uuid ?? "",
                            ["path"] = path ?? "",
                            ["name"] = name ?? "",
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[ERROR] switch_default_channel: " + ex.Message);
                    return new ToolResult
                    {
                        Success = false,
                        Error = "切换默认渠道失败: " + ex.Message
                    };
                }
            };
        }

        private static bool TryEnsureLive(
            IOperationChannel channel,
            out string error,
            out string kind,
            out string uuid,
            out string path,
            out string name)
        {
            error = null;
            kind = null;
            uuid = null;
            path = null;
            name = null;

            if (channel is WordChannel word)
            {
                if (!word.TryGetLiveDocument(out Word.Document _))
                {
                    error = "渠道对应的 Word 文档已关闭: " + channel.ChannelId;
                    return false;
                }

                kind = "word";
                uuid = word.DocUuid;
                path = word.FilePath;
                name = word.TryGetDisplayName() ?? "";
                return true;
            }

            if (channel is WpsChannel wps)
            {
                if (!wps.TryGetLiveDocument(out object _))
                {
                    error = "渠道对应的 WPS 文档已关闭: " + channel.ChannelId;
                    return false;
                }

                kind = "wps";
                uuid = wps.DocUuid;
                path = wps.FilePath;
                name = wps.TryGetDisplayName() ?? "";
                return true;
            }

            if (channel is ExcelChannel excel)
            {
                if (!excel.TryGetLiveWorkbook(out Excel.Workbook _))
                {
                    error = "渠道对应的 Excel 工作簿已关闭: " + channel.ChannelId;
                    return false;
                }

                kind = "excel";
                uuid = excel.DocUuid;
                path = excel.FilePath;
                name = excel.TryGetDisplayName() ?? "";
                return true;
            }

            if (channel is EtChannel et)
            {
                if (!et.TryGetLiveWorkbook(out object _))
                {
                    error = "渠道对应的 WPS 表格已关闭: " + channel.ChannelId;
                    return false;
                }

                kind = "et";
                uuid = et.DocUuid;
                path = et.FilePath;
                name = et.TryGetDisplayName() ?? "";
                return true;
            }

            if (channel is PptChannel ppt)
            {
                if (!ppt.TryGetLivePresentation(out _))
                {
                    error = "渠道对应的 PowerPoint 演示文稿已关闭: " + channel.ChannelId;
                    return false;
                }

                kind = "ppt";
                uuid = ppt.DocUuid;
                path = ppt.FilePath;
                name = ppt.TryGetDisplayName() ?? "";
                return true;
            }

            if (channel is WppChannel wpp)
            {
                if (!wpp.TryGetLivePresentation(out object _))
                {
                    error = "渠道对应的 WPS 演示已关闭: " + channel.ChannelId;
                    return false;
                }

                kind = "wpp";
                uuid = wpp.DocUuid;
                path = wpp.FilePath;
                name = wpp.TryGetDisplayName() ?? "";
                return true;
            }

            error = "不支持的渠道类型: " + channel.ChannelId;
            return false;
        }

        private static void TryActivate(IOperationChannel channel)
        {
            try
            {
                if (channel is WordChannel word)
                {
                    if (word.TryGetLiveDocument(out Word.Document doc))
                    {
                        try
                        {
                            doc.Activate();
                        }
                        catch (Exception)
                        {
                        }

                        DocumentState.BindAndActivate(doc);
                    }

                    return;
                }

                if (channel is WpsChannel wps)
                {
                    if (wps.TryGetLiveDocument(out object doc))
                    {
                        try
                        {
                            OpenFiles.WpsCom.Invoke(doc, "Activate");
                        }
                        catch (Exception)
                        {
                        }
                    }

                    return;
                }

                if (channel is ExcelChannel excel)
                {
                    if (excel.TryGetLiveWorkbook(out Excel.Workbook book))
                    {
                        try
                        {
                            book.Activate();
                        }
                        catch (Exception)
                        {
                        }
                    }

                    return;
                }

                if (channel is EtChannel et)
                {
                    if (et.TryGetLiveWorkbook(out object book))
                    {
                        try
                        {
                            OpenFiles.EtCom.Invoke(book, "Activate");
                        }
                        catch (Exception)
                        {
                        }
                    }

                    return;
                }

                if (channel is PptChannel ppt)
                {
                    ppt.TryActivate();
                    return;
                }

                if (channel is WppChannel wpp)
                {
                    wpp.TryActivate();
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
