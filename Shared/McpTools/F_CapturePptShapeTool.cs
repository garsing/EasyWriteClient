using System;
using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.PresentationHost;

namespace WordAddIn1
{
    /// <summary>
    /// PPT/WPP 形状截图：单 shape_id 或 shape_ids 宫格。
    /// </summary>
    public static class F_CapturePptShapeTool
    {
        private const string MigratedAwayMessage =
            "形状截图请改用 F_capture_ppt_shape（shape_id / shape_ids）；本工具只负责整页/页九宫格等。";

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_capture_ppt_shape"] = async (args) =>
            {
                try
                {
                    string reject = RejectForbiddenArgs(args);
                    if (reject != null)
                    {
                        return Fail(reject);
                    }

                    string shapeId = GetString(args, "shape_id");
                    List<string> shapeIds = TryGetShapeIds(args, out string idsError);
                    if (idsError != null)
                    {
                        return Fail(idsError);
                    }

                    bool hasSingle = !string.IsNullOrWhiteSpace(shapeId);
                    bool hasMulti = shapeIds != null && shapeIds.Count > 0;
                    if (hasSingle == hasMulti)
                    {
                        return Fail("请只传 shape_id（单张）或 shape_ids（宫格）其一");
                    }

                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return Fail(resolveError);
                    }

                    if (channel.Kind != ChannelKind.Ppt && channel.Kind != ChannelKind.Wpp)
                    {
                        return Fail("仅已打开的 ppt/wpp 支持形状截图");
                    }

                    ToolResult captured = hasSingle
                        ? CaptureSingle(channel, shapeId.Trim())
                        : CaptureGrid(channel, shapeIds);

                    if (captured == null || !captured.Success)
                    {
                        return captured ?? Fail("形状截图失败");
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
                catch (Exception ex)
                {
                    return Fail("形状截图失败: " + ex.Message);
                }
            };
        }

        /// <summary>旧 F_capture 误传 shape_id 时复用同一文案。</summary>
        internal static string ShapeMigratedAwayMessage => MigratedAwayMessage;

        private static ToolResult CaptureSingle(IOperationChannel channel, string shapeId)
        {
            if (!PresentationHostAdapter.TryCaptureShape(
                    channel,
                    shapeId,
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
                ["slide_count"] = hostResult.SlideCount,
                ["shape_id"] = hostResult.ShapeId ?? shapeId,
                ["capture_mode"] = "shape",
                ["scale"] = hostResult.Scale
            };
            if (!string.IsNullOrEmpty(hostResult.SlideId))
            {
                extra["slide_id"] = hostResult.SlideId;
            }

            if (hostResult.BoundsLeftPct.HasValue
                && hostResult.BoundsTopPct.HasValue
                && hostResult.BoundsWidthPct.HasValue
                && hostResult.BoundsHeightPct.HasValue)
            {
                extra["bounds_pct"] = new Dictionary<string, object>
                {
                    ["left"] = Math.Round(hostResult.BoundsLeftPct.Value, 4),
                    ["top"] = Math.Round(hostResult.BoundsTopPct.Value, 4),
                    ["width"] = Math.Round(hostResult.BoundsWidthPct.Value, 4),
                    ["height"] = Math.Round(hostResult.BoundsHeightPct.Value, 4)
                };
            }

            return OkImage(hostResult.Image, extra);
        }

        private static ToolResult CaptureGrid(IOperationChannel channel, List<string> shapeIds)
        {
            if (shapeIds.Count == 1)
            {
                return CaptureSingle(channel, shapeIds[0]);
            }

            if (shapeIds.Count > ImageCaptureCompressor.ContactSheetMaxPages)
            {
                return Fail("shape_ids 一次最多 9 个，请拆批");
            }

            for (int i = 0; i < shapeIds.Count; i++)
            {
                if (!PptShapeId.TryParseShape(shapeIds[i], out _, out _))
                {
                    return Fail("形状截图须提供完整 shape_id（sid{SlideID}-s{n}）");
                }
            }

            string firstSlide = null;
            for (int i = 0; i < shapeIds.Count; i++)
            {
                PptShapeId.TryParseShape(shapeIds[i], out string slideId, out _);
                if (firstSlide == null)
                {
                    firstSlide = slideId;
                }
                else if (!string.Equals(firstSlide, slideId, StringComparison.Ordinal))
                {
                    return Fail("shape_ids 须属于同一页（SlideID 一致）");
                }
            }

            if (!PresentationHostAdapter.TryCaptureShapes(
                    channel,
                    shapeIds,
                    out List<PresentationCaptureResult> hostResults,
                    out ToolResult errorResult))
            {
                return errorResult;
            }

            var tiles = new List<ImageCaptureCompressor.ShapeContactSheetTile>(hostResults.Count);
            var echoedIds = new List<string>(hostResults.Count);
            string kind = "";
            string channelId = "";
            string slideIdOut = firstSlide ?? "";
            int pageNumber = 0;
            int slideCount = 0;
            for (int i = 0; i < hostResults.Count; i++)
            {
                PresentationCaptureResult r = hostResults[i];
                kind = r.Kind ?? kind;
                channelId = ChannelRegistry.ToPublicId(r.ChannelId) ?? r.ChannelId ?? channelId;
                slideIdOut = r.SlideId ?? slideIdOut;
                pageNumber = r.PageNumber;
                slideCount = r.SlideCount;
                string id = r.ShapeId ?? shapeIds[i];
                echoedIds.Add(id);
                tiles.Add(new ImageCaptureCompressor.ShapeContactSheetTile
                {
                    Index = i + 1,
                    ShapeId = id,
                    ImageBytes = r.Image == null ? null : r.Image.Bytes
                });
            }

            int cols = shapeIds.Count <= 4 ? 2 : 3;
            int rows = cols;
            ImageCaptureCompressor.CompressedImage sheet;
            try
            {
                sheet = ImageCaptureCompressor.ComposeShapeContactSheet(tiles);
            }
            catch (Exception ex)
            {
                return Fail("拼形状九宫格失败: " + ex.Message);
            }

            return OkImage(
                sheet,
                new Dictionary<string, object>
                {
                    ["source_kind"] = kind,
                    ["channel_id"] = channelId,
                    ["page_number"] = pageNumber,
                    ["slide_count"] = slideCount,
                    ["slide_id"] = slideIdOut,
                    ["capture_mode"] = "shape_contact_sheet",
                    ["layout"] = "contact_sheet",
                    ["shape_ids"] = echoedIds,
                    ["cell_count"] = echoedIds.Count,
                    ["grid_cols"] = cols,
                    ["grid_rows"] = rows
                });
        }

        private static string RejectForbiddenArgs(IReadOnlyDictionary<string, object> args)
        {
            if (HasNonEmpty(args, "slide_id"))
            {
                return "形状截图请用 shape_id / shape_ids，不要 slide_id";
            }

            if (HasNonEmpty(args, "page_number")
                || HasNonEmpty(args, "page_from")
                || HasNonEmpty(args, "page_to"))
            {
                return "形状截图不收 page_number / page_from / page_to；整页请用 F_capture_document_page_image";
            }

            if (HasNonEmpty(args, "storage_doc_uuid"))
            {
                return "仅已打开的 ppt/wpp 支持形状截图";
            }

            if (HasNonEmpty(args, "sheet") || HasNonEmpty(args, "range"))
            {
                return "形状截图不收 sheet/range";
            }

            if (HasNonEmpty(args, "source_channel_id"))
            {
                return "截图不收 source_channel_id，请用 channel_id";
            }

            return null;
        }

        private static List<string> TryGetShapeIds(
            IReadOnlyDictionary<string, object> args,
            out string error)
        {
            error = null;
            if (args == null || !args.ContainsKey("shape_ids") || args["shape_ids"] == null)
            {
                return null;
            }

            object raw = args["shape_ids"];
            if (raw is string s)
            {
                if (string.IsNullOrWhiteSpace(s))
                {
                    return null;
                }

                error = "shape_ids 须为字符串数组";
                return null;
            }

            var list = new List<string>();
            if (raw is IEnumerable enumerable && !(raw is string))
            {
                foreach (object item in enumerable)
                {
                    string id = Convert.ToString(item)?.Trim() ?? "";
                    if (string.IsNullOrEmpty(id))
                    {
                        error = "shape_ids 中不能有空项";
                        return null;
                    }

                    list.Add(id);
                }
            }
            else
            {
                error = "shape_ids 须为字符串数组";
                return null;
            }

            if (list.Count == 0)
            {
                return null;
            }

            if (list.Count > ImageCaptureCompressor.ContactSheetMaxPages)
            {
                error = "shape_ids 一次最多 9 个，请拆批";
                return null;
            }

            return list;
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

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
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

            return Convert.ToString(args[key]) ?? "";
        }

        private static Dictionary<string, object> ToDict(object data)
        {
            if (data is Dictionary<string, object> dict)
            {
                return dict;
            }

            return null;
        }

        private static byte[] TryGetBytes(Dictionary<string, object> data)
        {
            if (data == null || !data.ContainsKey("image_base64"))
            {
                return null;
            }

            string b64 = Convert.ToString(data["image_base64"]);
            if (string.IsNullOrWhiteSpace(b64))
            {
                return null;
            }

            try
            {
                return Convert.FromBase64String(b64);
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
