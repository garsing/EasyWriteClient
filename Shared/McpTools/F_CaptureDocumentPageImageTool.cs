using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.BrowserHost;
using WordAddIn1.DocumentHost;
using WordAddIn1.PresentationHost;
using WordAddIn1.SpreadsheetHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 截取渠道文档页/表格区域/幻灯片/浏览器可视区。经各 HostAdapter 分流。
    /// </summary>
    public static class F_CaptureDocumentPageImageTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_capture_document_page_image"] = async (args) =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("[F_capture_document_page_image] 开始执行");

                    string unused = RejectUnknownPrimary(args);
                    if (unused != null)
                    {
                        return Fail(unused);
                    }

                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return Fail(resolveError);
                    }

                    string kindError = ValidateArgsForKind(channel.Kind, args, out PageCaptureRequest pageReq);
                    if (kindError != null)
                    {
                        return Fail(kindError);
                    }

                    ToolResult captured;
                    if (channel.Kind == ChannelKind.Word || channel.Kind == ChannelKind.Wps)
                    {
                        captured = await CaptureWordAsync(args, wordApplication, pageReq.PageNumber).ConfigureAwait(false);
                    }
                    else if (channel.Kind == ChannelKind.Excel || channel.Kind == ChannelKind.Et)
                    {
                        captured = CaptureSpreadsheet(channel, args);
                    }
                    else if (channel.Kind == ChannelKind.Ppt || channel.Kind == ChannelKind.Wpp)
                    {
                        captured = pageReq.IsContactSheet
                            ? CapturePresentationRange(channel, pageReq.PageFrom.Value, pageReq.PageTo.Value)
                            : CapturePresentation(channel, pageReq.PageNumber.Value);
                    }
                    else if (channel.Kind == ChannelKind.Browser)
                    {
                        captured = await OfficeStaScheduler
                            .InvokeOnUiAsync(() => CaptureBrowserAsync(channel))
                            .ConfigureAwait(false);
                    }
                    else
                    {
                        return Fail("unsupported: 当前渠道不支持截图");
                    }

                    if (captured == null || !captured.Success)
                    {
                        return captured ?? Fail("截图失败");
                    }

                    var data = captured.Data as Dictionary<string, object>
                        ?? ToDict(captured.Data);
                    byte[] bytes = TryGetBytes(data);
                    string format = data != null && data.ContainsKey("format")
                        ? Convert.ToString(data["format"])
                        : "";
                    ToolResult pathError = await CapturePathWriter
                        .TryAttachPathAsync(args, bytes, format, data)
                        .ConfigureAwait(false);
                    if (pathError != null)
                    {
                        return pathError;
                    }

                    return new ToolResult { Success = true, Data = data };
                }
                catch (PageCaptureHelper.PageOutOfRangeException ex)
                {
                    return new ToolResult
                    {
                        Success = false,
                        Error = ex.Message,
                        Data = new Dictionary<string, object>
                        {
                            ["page_number"] = ex.RequestedPage,
                            ["total_pages"] = ex.TotalPages
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[F_capture_document_page_image] 失败: {ex.Message}");
                    return Fail("截图失败: " + ex.Message);
                }
            };
        }

        private static string RejectUnknownPrimary(IReadOnlyDictionary<string, object> args)
        {
            if (args != null && args.ContainsKey("page_numbers") && args["page_numbers"] != null)
            {
                return "截图不收 page_numbers，PPT 批量请用 page_from + page_to";
            }

            if (HasNonEmpty(args, "slide_id"))
            {
                return "截图请用 page_number，不要 slide_id";
            }

            if (HasNonEmpty(args, "storage_doc_uuid"))
            {
                return "客户端截图不收 storage_doc_uuid（知识库截图由服务端处理）";
            }

            if (HasNonEmpty(args, "source_channel_id"))
            {
                return "截图不收 source_channel_id，请用 channel_id";
            }

            if (HasNonEmpty(args, "document_name") || HasNonEmpty(args, "knowledge_base_uuid"))
            {
                return "截图不收名+库，请用 channel_id 或 storage_doc_uuid";
            }

            return null;
        }

        private sealed class PageCaptureRequest
        {
            public int? PageNumber;
            public int? PageFrom;
            public int? PageTo;
            public bool IsContactSheet;
        }

        private static string ValidateArgsForKind(
            ChannelKind kind,
            IReadOnlyDictionary<string, object> args,
            out PageCaptureRequest pageReq)
        {
            pageReq = new PageCaptureRequest();
            string pageErr = TryParseOptionalInt(args, "page_number", out pageReq.PageNumber);
            if (pageErr != null)
            {
                return pageErr;
            }

            string fromErr = TryParseOptionalInt(args, "page_from", out pageReq.PageFrom);
            if (fromErr != null)
            {
                return fromErr;
            }

            string toErr = TryParseOptionalInt(args, "page_to", out pageReq.PageTo);
            if (toErr != null)
            {
                return toErr;
            }

            bool hasSheet = HasNonEmpty(args, "sheet");
            bool hasRange = HasNonEmpty(args, "range");
            bool hasSpan = pageReq.PageFrom.HasValue || pageReq.PageTo.HasValue;

            if (kind == ChannelKind.Word || kind == ChannelKind.Wps)
            {
                if (hasSpan)
                {
                    return "Word/WPS 截图不收 page_from/page_to";
                }

                if (hasSheet || hasRange)
                {
                    return "Word/WPS 截图不收 sheet/range";
                }

                return null;
            }

            if (kind == ChannelKind.Excel || kind == ChannelKind.Et)
            {
                if (hasSpan)
                {
                    return "表格截图不收 page_from/page_to，请用 sheet + range";
                }

                if (pageReq.PageNumber.HasValue)
                {
                    return "表格截图不收 page_number，请用 sheet + range";
                }

                if (!hasSheet || !hasRange)
                {
                    return "表格截图必须提供 sheet 与 A1 range";
                }

                return null;
            }

            if (kind == ChannelKind.Ppt || kind == ChannelKind.Wpp)
            {
                if (hasSheet || hasRange)
                {
                    return "演示文稿截图不收 sheet/range，请用 page_number 或 page_from + page_to";
                }

                if (hasSpan)
                {
                    if (pageReq.PageNumber.HasValue)
                    {
                        return "page_number 与 page_from/page_to 不能同时传";
                    }

                    if (!pageReq.PageFrom.HasValue || !pageReq.PageTo.HasValue)
                    {
                        return "page_from 与 page_to 必须成对";
                    }

                    if (pageReq.PageFrom.Value > pageReq.PageTo.Value)
                    {
                        return "page_from 不能大于 page_to";
                    }

                    int count = pageReq.PageTo.Value - pageReq.PageFrom.Value + 1;
                    if (count > ImageCaptureCompressor.ContactSheetMaxPages)
                    {
                        return "一次最多 9 页，请分段（如 1–9、10–18）";
                    }

                    if (count == 1)
                    {
                        pageReq.PageNumber = pageReq.PageFrom;
                        pageReq.IsContactSheet = false;
                        return null;
                    }

                    pageReq.IsContactSheet = true;
                    return null;
                }

                if (!pageReq.PageNumber.HasValue)
                {
                    return "演示文稿截图必须提供 page_number，或 page_from + page_to（从概览 index 抄）";
                }

                return null;
            }

            if (kind == ChannelKind.Browser)
            {
                if (hasSpan)
                {
                    return "网页截图不收 page_from/page_to，只截当前可视区";
                }

                if (pageReq.PageNumber.HasValue || hasSheet || hasRange)
                {
                    return "网页截图不收 page_number/sheet/range，只截当前可视区";
                }

                return null;
            }

            return "unsupported: 当前渠道不支持截图";
        }

        private static async Task<ToolResult> CaptureWordAsync(
            Dictionary<string, object> args,
            object wordApplication,
            int? requestedPage)
        {
            if (!DocumentHostAdapter.TryResolveInteropDocument(
                    args,
                    wordApplication,
                    out InteropDocumentHandle docHandle,
                    out ToolResult resolveError,
                    activateDocument: false))
            {
                return resolveError;
            }

            Word.Document doc = docHandle.Document;
            Word.Application app = null;
            try
            {
                app = doc.Application;
            }
            catch (Exception)
            {
            }

            if (app == null)
            {
                app = wordApplication as Word.Application;
            }

            if (app == null)
            {
                return Fail("无法获取文档 Application（Word/WPS）");
            }

            var capture = PageCaptureHelper.CaptureDocumentPage(app, requestedPage, doc);
            await Task.CompletedTask;
            string sourceKind = string.Equals(docHandle.HostName, "wps", StringComparison.OrdinalIgnoreCase)
                ? "wps"
                : "word";
            return new ToolResult
            {
                Success = true,
                Data = new Dictionary<string, object>
                {
                    ["source_kind"] = sourceKind,
                    ["page_number"] = capture.PageNumber,
                    ["total_pages"] = capture.TotalPages,
                    ["format"] = capture.Format,
                    ["width_px"] = capture.WidthPx,
                    ["height_px"] = capture.HeightPx,
                    ["image_base64"] = capture.ImageBase64,
                    ["captured_at"] = capture.CapturedAt,
                    ["doc_title"] = capture.DocTitle,
                    ["channel_id"] = ChannelRegistry.ToPublicId(docHandle.ChannelId) ?? docHandle.ChannelId ?? "",
                    ["host"] = docHandle.HostName ?? ""
                }
            };
        }

        private static ToolResult CaptureSpreadsheet(IOperationChannel channel, Dictionary<string, object> args)
        {
            string sheet = GetString(args, "sheet");
            string range = GetString(args, "range");
            if (!SpreadsheetHostAdapter.TryCaptureRange(
                    channel,
                    sheet,
                    range,
                    out SpreadsheetCaptureResult hostResult,
                    out ToolResult errorResult))
            {
                return errorResult;
            }

            var image = hostResult.Image;
            return OkImage(
                image,
                new Dictionary<string, object>
                {
                    ["source_kind"] = hostResult.Kind ?? "",
                    ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? hostResult.ChannelId ?? "",
                    ["sheet"] = hostResult.Sheet ?? "",
                    ["range"] = hostResult.Range ?? "",
                    ["actual_range"] = hostResult.ActualRange ?? ""
                });
        }

        private static ToolResult CapturePresentation(IOperationChannel channel, int pageNumber)
        {
            if (!PresentationHostAdapter.TryCaptureSlide(
                    channel,
                    pageNumber,
                    out PresentationCaptureResult hostResult,
                    out ToolResult errorResult))
            {
                return errorResult;
            }

            var extra = new Dictionary<string, object>
            {
                ["source_kind"] = hostResult.Kind ?? "",
                ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? hostResult.ChannelId ?? "",
                ["page_number"] = hostResult.PageNumber,
                ["slide_count"] = hostResult.SlideCount
            };
            if (!string.IsNullOrEmpty(hostResult.SlideId))
            {
                extra["slide_id"] = hostResult.SlideId;
            }

            return OkImage(hostResult.Image, extra);
        }

        private static ToolResult CapturePresentationRange(
            IOperationChannel channel,
            int pageFrom,
            int pageTo)
        {
            var tiles = new List<ImageCaptureCompressor.ContactSheetTile>();
            var pages = new List<Dictionary<string, object>>();
            string kind = "";
            string channelId = "";
            int slideCount = 0;
            for (int page = pageFrom; page <= pageTo; page++)
            {
                if (!PresentationHostAdapter.TryCaptureSlide(
                        channel,
                        page,
                        out PresentationCaptureResult hostResult,
                        out ToolResult errorResult,
                        ImageCaptureCompressor.ContactSheetCellLongEdge))
                {
                    return errorResult;
                }

                kind = hostResult.Kind ?? "";
                channelId = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? hostResult.ChannelId ?? "";
                slideCount = hostResult.SlideCount;
                string slideId = hostResult.SlideId ?? "";
                tiles.Add(new ImageCaptureCompressor.ContactSheetTile
                {
                    Index = page,
                    SlideId = slideId,
                    ImageBytes = hostResult.Image == null ? null : hostResult.Image.Bytes
                });
                pages.Add(new Dictionary<string, object>
                {
                    ["index"] = page,
                    ["slide_id"] = slideId
                });
            }

            ImageCaptureCompressor.CompressedImage sheet;
            try
            {
                sheet = ImageCaptureCompressor.ComposeContactSheet(tiles);
            }
            catch (Exception ex)
            {
                return Fail("拼九宫格失败: " + ex.Message);
            }

            return OkImage(
                sheet,
                new Dictionary<string, object>
                {
                    ["source_kind"] = kind,
                    ["channel_id"] = channelId,
                    ["page_number"] = pageFrom,
                    ["page_from"] = pageFrom,
                    ["page_to"] = pageTo,
                    ["slide_count"] = slideCount,
                    ["layout"] = "contact_sheet",
                    ["pages"] = pages
                });
        }

        private static async Task<ToolResult> CaptureBrowserAsync(IOperationChannel channel)
        {
            BrowserChannel browser = channel as BrowserChannel;
            if (browser == null)
            {
                return Fail("channel_id 不是浏览器渠道");
            }

            if (!string.Equals(browser.Track, "agent", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(browser.Track, "attach", StringComparison.OrdinalIgnoreCase))
            {
                return Fail("不支持的浏览器 track: " + browser.Track);
            }

            if (!browser.IsLive())
            {
                return Fail("浏览器渠道对应的页面已关闭: " + browser.ChannelId);
            }

            BrowserCaptureResult result = await BrowserHostAdapter
                .CaptureViewportAsync(browser)
                .ConfigureAwait(true);
            if (!result.Success)
            {
                return Fail(result.Error ?? "截取可视区失败");
            }

            return OkImage(
                result.Image,
                new Dictionary<string, object>
                {
                    ["source_kind"] = "browser",
                    ["channel_id"] = ChannelRegistry.ToPublicId(result.Channel.ChannelId) ?? result.Channel.ChannelId,
                    ["url"] = result.Url ?? "",
                    ["title"] = result.Title ?? ""
                });
        }

        private static ToolResult OkImage(
            ImageCaptureCompressor.CompressedImage image,
            Dictionary<string, object> extra)
        {
            if (image == null || image.Bytes == null || image.Bytes.Length == 0)
            {
                return Fail("截图结果为空");
            }

            var data = extra ?? new Dictionary<string, object>();
            data["format"] = image.Format;
            data["width_px"] = image.Width;
            data["height_px"] = image.Height;
            data["image_base64"] = Convert.ToBase64String(image.Bytes);
            data["captured_at"] = DateTimeOffset.Now.ToString("yyyy-MM-ddTHH:mm:sszzz");
            return new ToolResult { Success = true, Data = data };
        }

        private static string TryParseOptionalInt(
            IReadOnlyDictionary<string, object> args,
            string key,
            out int? value)
        {
            value = null;
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            string rawText = Convert.ToString(args[key])?.Trim();
            if (string.IsNullOrEmpty(rawText))
            {
                return null;
            }

            object raw = args[key];
            if (raw is int i)
            {
                value = i;
            }
            else if (raw is long l)
            {
                value = (int)l;
            }
            else if (raw is double d && Math.Abs(d - Math.Round(d)) < double.Epsilon)
            {
                value = (int)Math.Round(d);
            }
            else if (!int.TryParse(rawText, out int parsed))
            {
                return key + " 必须是整数";
            }
            else
            {
                value = parsed;
            }

            if (value.Value < 1)
            {
                return key + " 必须是 ≥1 的整数";
            }

            return null;
        }

        private static bool HasNonEmpty(IReadOnlyDictionary<string, object> args, string key)
        {
            return !string.IsNullOrWhiteSpace(GetString(args, key));
        }

        private static string GetString(IReadOnlyDictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return "";
            }

            return Convert.ToString(args[key])?.Trim() ?? "";
        }

        private static byte[] TryGetBytes(Dictionary<string, object> data)
        {
            if (data == null || !data.ContainsKey("image_base64") || data["image_base64"] == null)
            {
                return null;
            }

            try
            {
                return Convert.FromBase64String(Convert.ToString(data["image_base64"]) ?? "");
            }
            catch (FormatException)
            {
                return null;
            }
        }

        private static Dictionary<string, object> ToDict(object data)
        {
            var dict = new Dictionary<string, object>();
            if (data == null)
            {
                return dict;
            }

            foreach (var prop in data.GetType().GetProperties())
            {
                dict[prop.Name] = prop.GetValue(data);
            }

            return dict;
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
