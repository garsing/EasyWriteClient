using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 工作区图片（Matplotlib PNG 等）→ Ensure → 按 target/光标插入 InlineShape。
    /// </summary>
    public static class F_InsertWorkspaceImageTool
    {
        private const double DefaultWidthCm = 14.0;
        private const double PointsPerCm = 28.35;

        private static readonly HashSet<string> AllowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".png",
            ".jpg",
            ".jpeg",
            ".webp",
        };

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_insert_workspace_image"] = async (args) =>
            {
                var sw = Stopwatch.StartNew();
                try
                {
                    if (wordApplication == null)
                    {
                        return Fail("Word应用程序不可用");
                    }

                    dynamic wordApp = wordApplication;
                    if (wordApp.ActiveDocument == null)
                    {
                        return Fail("没有活动的Word文档");
                    }

                    Word.Document document = wordApp.ActiveDocument;
                    DocumentState.BindAndActivate(document);

                    string rawName = args != null && args.ContainsKey("filename")
                        ? args["filename"]?.ToString()
                        : null;
                    string safeName = WorkspacePathResolver.SanitizeFilename(rawName);
                    if (string.IsNullOrEmpty(safeName))
                    {
                        return Fail("invalid_filename: 文件名无效或含路径");
                    }

                    string ext = Path.GetExtension(safeName);
                    if (string.IsNullOrEmpty(ext) || !AllowedExtensions.Contains(ext))
                    {
                        return Fail(
                            "unsupported_format: 仅支持 .png / .jpg / .jpeg / .webp");
                    }

                    double widthCm = TryParseWidthCm(args) ?? DefaultWidthCm;

                    bool hadLocal = false;
                    try
                    {
                        string existing = WorkspacePathResolver.ResolveReadPath(safeName);
                        hadLocal = !string.IsNullOrEmpty(existing) && File.Exists(existing);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[InsertWorkspaceImage] ResolveReadPath 警告: {ex.Message}");
                    }

                    var ensure = await McpToolsHelpers.EnsureWorkspaceFileAsync(safeName)
                        .ConfigureAwait(true);
                    if (!ensure.success || string.IsNullOrEmpty(ensure.localPath) || !File.Exists(ensure.localPath))
                    {
                        string detail = string.IsNullOrEmpty(ensure.error)
                            ? safeName
                            : ensure.error;
                        return Fail($"file_not_found: {detail}");
                    }

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[InsertWorkspaceImage] ReadWord 警告: {ex.Message}");
                    }

                    var locate = InsertTargetLocatorHelper.TryResolveInsertRange(document, args);
                    if (!locate.Success)
                    {
                        return Fail(locate.Error ?? "locate_failed");
                    }

                    double heightCm;
                    double? aspect = TryGetImageAspectRatio(ensure.localPath);
                    if (aspect.HasValue && aspect.Value > 0)
                    {
                        heightCm = widthCm / aspect.Value;
                    }
                    else
                    {
                        heightCm = 0;
                    }

                    int pictureStart = -1;
                    try
                    {
                        Word.Range range = locate.Range;
                        Word.InlineShape picture = range.InlineShapes.AddPicture(
                            FileName: ensure.localPath,
                            LinkToFile: false,
                            SaveWithDocument: true);
                        picture.Width = (float)(widthCm * PointsPerCm);
                        if (heightCm > 0)
                        {
                            picture.Height = (float)(heightCm * PointsPerCm);
                        }
                        else
                        {
                            // 无法读像素比时仅设宽度，保留 Word 默认纵横比
                            heightCm = picture.Height / PointsPerCm;
                        }

                        pictureStart = picture.Range.Start;

                        // 与建表/建图约定一致：光标落到图片后（并换行）
                        PostModifyNavigateHelper.NavigateAfterInsertObject(
                            document.Application, picture.Range, "insert_workspace_image");
                    }
                    catch (Exception ex)
                    {
                        return Fail($"insert_failed: {ex.Message}");
                    }

                    var warnings = new List<string>();
                    string resolvedImageId = ImageResolveHelper.TryResolveImageIdAfterInsert(
                        document, pictureStart, warnings);

                    bool downloaded = !hadLocal;
                    var data = new Dictionary<string, object>
                    {
                        ["inserted"] = true,
                        ["filename"] = safeName,
                        ["image_width_cm"] = Math.Round(widthCm, 2),
                        ["image_height_cm"] = Math.Round(heightCm, 2),
                        ["used_cursor"] = locate.UsedCursor,
                        ["downloaded"] = downloaded,
                        ["image_id"] = resolvedImageId,
                        ["message"] =
                            "已插入工作区图片（PNG/位图，非 Word Chart）。误插可用文档 checkpoint 回滚。",
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

                    Debug.WriteLine(
                        $"[InsertWorkspaceImage] ok file={safeName} downloaded={downloaded} " +
                        $"image_id={resolvedImageId ?? "(null)"} " +
                        $"used_cursor={locate.UsedCursor} w_cm={widthCm:F2} h_cm={heightCm:F2} " +
                        $"ms={sw.ElapsedMilliseconds}");

                    return new ToolResult { Success = true, Data = data };
                }
                catch (Exception ex)
                {
                    return Fail($"insert_failed: {ex.Message}");
                }
            };
        }

        /// <summary>宽/高比；失败返回 null。</summary>
        private static double? TryGetImageAspectRatio(string localPath)
        {
            try
            {
                using (var img = Image.FromFile(localPath))
                {
                    if (img.Height <= 0)
                    {
                        return null;
                    }

                    return (double)img.Width / img.Height;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[InsertWorkspaceImage] 读图尺寸失败: {ex.Message}");
                return null;
            }
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
