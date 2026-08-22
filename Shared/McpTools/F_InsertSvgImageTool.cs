using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// SVG → WebView2 PNG → 按 target/光标插入 InlineShape 图片。
    /// 文档触点经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_InsertSvgImageTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_insert_svg_image"] = async (args) =>
            {
                string pngPath = null;
                try
                {
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_insert_svg_image] resolve failed: {resolveError?.Error}; " +
                            $"arg.channel_id={ChannelContext.TryGetChannelIdFromParameters(args) ?? "(null)"}; " +
                            $"default={ChannelRegistry.DefaultChannelId ?? "(null)"}");
                        return Fail(resolveError?.Error ?? "无法解析文档渠道");
                    }

                    Word.Document document = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_insert_svg_image] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    string pathArg = FilePathResolver.TryGetArg(args, "path");
                    string svg = args != null && args.ContainsKey("svg") ? args["svg"]?.ToString() : null;
                    if (!string.IsNullOrEmpty(pathArg) && !string.IsNullOrEmpty(svg))
                    {
                        return Fail("path 与 svg 只能传其中一个");
                    }

                    if (string.IsNullOrEmpty(pathArg) && string.IsNullOrEmpty(svg))
                    {
                        return Fail("必须提供 path 或 svg");
                    }

                    if (!string.IsNullOrEmpty(pathArg))
                    {
                        if (!FilePathResolver.TryResolve(pathArg, out ResolvedFilePath svgResolved, out string pathError))
                        {
                            return Fail(pathError);
                        }

                        var svgRead = await FilePathResolver.ReadAsync(svgResolved).ConfigureAwait(true);
                        if (!svgRead.Success)
                        {
                            return Fail(svgRead.Error);
                        }

                        svg = svgRead.Text;
                    }

                    var validation = SvgSecurityValidator.Validate(svg);
                    if (!validation.Success)
                    {
                        return Fail(validation.Error);
                    }

                    double? widthCm = TryParseWidthCm(args);

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[InsertSvg] ReadWord 警告: {ex.Message}");
                    }

                    var locate = InsertTargetLocatorHelper.TryResolveInsertRange(document, args);
                    if (!locate.Success)
                    {
                        return Fail(locate.Error);
                    }

                    var raster = await SvgWebView2Rasterizer.RasterizeAsync(
                        validation.NormalizedSvg,
                        widthCm).ConfigureAwait(true);
                    if (!raster.Success)
                    {
                        return Fail(raster.Error ?? "render_failed");
                    }

                    pngPath = raster.PngPath;

                    int pictureStart = -1;
                    try
                    {
                        Word.Range range = locate.Range;
                        Word.InlineShape picture = range.InlineShapes.AddPicture(
                            FileName: pngPath,
                            LinkToFile: false,
                            SaveWithDocument: true);
                        picture.Width = (float)(raster.WidthCm * 28.35);
                        picture.Height = (float)(raster.HeightCm * 28.35);

                        pictureStart = picture.Range.Start;

                        // 与建表/建图约定一致：光标落到图片后（并换行）
                        PostModifyNavigateHelper.NavigateAfterInsertObject(
                            document.Application, picture.Range, "insert_svg_image");
                    }
                    catch (Exception ex)
                    {
                        return Fail($"insert_failed: {ex.Message}");
                    }

                    var warnings = new List<string>();
                    string resolvedImageId = ImageResolveHelper.TryResolveImageIdAfterInsert(
                        document, pictureStart, warnings);

                    var data = new Dictionary<string, object>
                    {
                        ["inserted"] = true,
                        ["image_width_cm"] = Math.Round(raster.WidthCm, 2),
                        ["image_height_cm"] = Math.Round(raster.HeightCm, 2),
                        ["used_cursor"] = locate.UsedCursor,
                        ["image_id"] = resolvedImageId,
                        ["message"] =
                            "已插入 SVG 示意图（PNG 图片，非 Word Chart）。误插可用文档 checkpoint 回滚。"
                    };

                    if (warnings.Count > 0)
                    {
                        data["message"] = data["message"] + " 警告: " + string.Join("; ", warnings);
                    }

                    var echo = InsertTargetLocatorHelper.BuildTargetEcho(locate.Spec);
                    if (echo != null)
                    {
                        data["target_echo"] = echo;
                    }

                    System.Diagnostics.Debug.WriteLine(
                        $"[InsertSvg] ok image_id={resolvedImageId ?? "(null)"} used_cursor={locate.UsedCursor} " +
                        $"w_cm={raster.WidthCm:F2} h_cm={raster.HeightCm:F2}");

                    return new ToolResult { Success = true, Data = data };
                }
                catch (Exception ex)
                {
                    return Fail($"insert_failed: {ex.Message}");
                }
                finally
                {
                    SvgWebView2Rasterizer.TryDelete(pngPath);
                }
            };
        }

        private static double? TryParseWidthCm(Dictionary<string, object> args)
        {
            if (args == null || !args.ContainsKey("width_cm") || args["width_cm"] == null)
            {
                return null;
            }

            if (args["width_cm"] is double d)
            {
                return d > 0 ? d : (double?)null;
            }

            if (args["width_cm"] is float f)
            {
                return f > 0 ? (double)f : (double?)null;
            }

            if (args["width_cm"] is int i)
            {
                return i > 0 ? (double)i : (double?)null;
            }

            if (double.TryParse(
                    args["width_cm"].ToString(),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out double parsed)
                || double.TryParse(
                    args["width_cm"].ToString(),
                    NumberStyles.Float,
                    CultureInfo.CurrentCulture,
                    out parsed))
            {
                return parsed > 0 ? parsed : (double?)null;
            }

            return null;
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
