using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Threading.Tasks;

namespace WordAddIn1
{
    public static class F_RenderSvgTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_render_svg"] = async (args) =>
            {
                string tempPng = null;
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (TryGetContractParseFailure(args, out ToolResult contractFail))
                    {
                        return contractFail;
                    }

                    if (!TryLoadSvg(args, out string svg, out string loadError))
                    {
                        return Fail(loadError);
                    }

                    var validation = SvgSecurityValidator.Validate(svg);
                    if (!validation.Success)
                    {
                        return Fail(validation.Error ?? validation.ErrorCode ?? "invalid_svg");
                    }

                    var raster = await SvgWebView2Rasterizer.RasterizeAsync(
                        validation.NormalizedSvg,
                        widthCm: null).ConfigureAwait(true);
                    if (!raster.Success)
                    {
                        return Fail(raster.Error ?? "render_failed");
                    }

                    tempPng = raster.PngPath;
                    if (!TryPlaceWorkspacePng(args, tempPng, out string destPath, out string filename, out string placeError))
                    {
                        return Fail(placeError);
                    }

                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["path"] = filename,
                            ["filename"] = filename,
                            ["bytes"] = new FileInfo(destPath).Length
                        }
                    };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return Fail("render_failed: " + ex.Message);
                }
                finally
                {
                    SvgWebView2Rasterizer.TryDelete(tempPng);
                }
            };
        }

        public static bool TryGetContractParseFailure(Dictionary<string, object> args, out ToolResult failure)
        {
            failure = null;
            string pathArg = FilePathResolver.TryGetArg(args, "path");
            string svg = args != null && args.ContainsKey("svg") ? args["svg"]?.ToString() : null;
            bool hasPath = !string.IsNullOrWhiteSpace(pathArg);
            bool hasSvg = !string.IsNullOrWhiteSpace(svg);
            if (hasPath == hasSvg)
            {
                failure = Fail("path 与 svg 须二选一");
                return true;
            }

            string filename = GetFilenameArg(args);
            if (!string.IsNullOrEmpty(filename) && !IsSafePngName(filename, out string nameError))
            {
                failure = Fail(nameError);
                return true;
            }

            if (hasPath)
            {
                if (!FilePathResolver.TryResolve(pathArg, out ResolvedFilePath resolved, out string pathError))
                {
                    failure = Fail(pathError);
                    return true;
                }

                if (!string.Equals(Path.GetExtension(resolved.LocalPath), ".svg", StringComparison.OrdinalIgnoreCase))
                {
                    failure = Fail("path 须为 .svg 文件");
                    return true;
                }

                if (!File.Exists(resolved.LocalPath))
                {
                    failure = Fail("文件不存在: " + resolved.Display);
                    return true;
                }
            }

            return false;
        }

        private static bool TryLoadSvg(Dictionary<string, object> args, out string svg, out string error)
        {
            svg = null;
            error = null;
            string pathArg = FilePathResolver.TryGetArg(args, "path");
            if (!string.IsNullOrWhiteSpace(pathArg))
            {
                FilePathResolver.TryResolve(pathArg, out ResolvedFilePath resolved, out _);
                if (!DocumentCheckpointMutatingTools.TryReadLocalText(resolved, out svg, out error))
                {
                    return false;
                }

                return true;
            }

            svg = args != null && args.ContainsKey("svg") ? args["svg"]?.ToString() : null;
            return !string.IsNullOrWhiteSpace(svg);
        }

        private static bool TryPlaceWorkspacePng(
            Dictionary<string, object> args,
            string tempPng,
            out string destPath,
            out string filename,
            out string error)
        {
            destPath = null;
            filename = null;
            error = null;
            string sessionDir;
            try
            {
                sessionDir = WorkspacePathResolver.GetSessionDirectory();
                Directory.CreateDirectory(sessionDir);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }

            string requested = GetFilenameArg(args);
            if (!string.IsNullOrEmpty(requested))
            {
                destPath = Path.Combine(sessionDir, requested);
                if (File.Exists(destPath))
                {
                    error = "文件已存在，请换 filename: " + requested;
                    return false;
                }

                filename = requested;
                File.Copy(tempPng, destPath, false);
                return true;
            }

            for (int i = 0; i < 32; i++)
            {
                filename = "svg-" + RandomToken8() + ".png";
                destPath = Path.Combine(sessionDir, filename);
                if (!File.Exists(destPath))
                {
                    File.Copy(tempPng, destPath, false);
                    return true;
                }
            }

            error = "无法分配工作区文件名";
            return false;
        }

        private static string GetFilenameArg(Dictionary<string, object> args)
        {
            if (args == null || !args.ContainsKey("filename") || args["filename"] == null)
            {
                return "";
            }

            return args["filename"].ToString().Trim();
        }

        private static bool IsSafePngName(string filename, out string error)
        {
            error = null;
            if (filename.IndexOfAny(new[] { '/', '\\', ':' }) >= 0
                || filename.IndexOf("..", StringComparison.Ordinal) >= 0
                || filename.IndexOf(Path.DirectorySeparatorChar) >= 0)
            {
                error = "filename 须为工作区裸名";
                return false;
            }

            if (!string.Equals(Path.GetExtension(filename), ".png", StringComparison.OrdinalIgnoreCase))
            {
                error = "filename 须为 .png";
                return false;
            }

            if (string.IsNullOrEmpty(Path.GetFileNameWithoutExtension(filename)))
            {
                error = "filename 非法";
                return false;
            }

            return true;
        }

        private static string RandomToken8()
        {
            const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            byte[] buf = new byte[8];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(buf);
            }

            char[] chars = new char[8];
            for (int i = 0; i < 8; i++)
            {
                chars[i] = alphabet[buf[i] % alphabet.Length];
            }

            return new string(chars);
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
