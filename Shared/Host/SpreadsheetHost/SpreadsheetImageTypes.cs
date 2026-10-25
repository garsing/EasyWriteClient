using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace WordAddIn1.SpreadsheetHost
{
    internal sealed class SpreadsheetImageRequest
    {
        public string Action { get; set; }

        public string Sheet { get; set; }

        public string Anchor { get; set; }

        public string ImageId { get; set; }

        public double? WidthCm { get; set; }

        public string LocalPath { get; set; }

        public string PathDisplay { get; set; }
    }

    internal sealed class SpreadsheetImageInfo
    {
        public string ImageId { get; set; }

        public string Sheet { get; set; }

        public string Anchor { get; set; }

        public double WidthCm { get; set; }

        public double HeightCm { get; set; }
    }

    internal sealed class SpreadsheetImageResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string ImageId { get; set; }

        public string Sheet { get; set; }

        public string Anchor { get; set; }

        public double WidthCm { get; set; }

        public double HeightCm { get; set; }

        public string PathDisplay { get; set; }

        public List<SpreadsheetImageInfo> Images { get; set; }

        public int ImagesCount { get; set; }

        public List<string> Warnings { get; set; }
    }

    internal static class SpreadsheetImageParse
    {
        public const double DefaultWidthCm = 8.0;

        private static readonly HashSet<string> AllowedImageExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".png",
            ".jpg",
            ".jpeg",
            ".webp",
        };

        private static readonly Regex AnchorRegex = new Regex(
            @"^[A-Za-z]{1,3}[1-9][0-9]{0,6}$",
            RegexOptions.Compiled);

        private static readonly Regex ImageIdRegex = new Regex(
            @"^I_[A-Za-z0-9]{8}$",
            RegexOptions.Compiled);

        public static bool TryParseRequest(
            Dictionary<string, object> args,
            out SpreadsheetImageRequest request,
            out string error)
        {
            request = null;
            error = null;
            string action = GetString(args, "action").ToLowerInvariant();
            if (action != "list" && action != "insert" && action != "update" && action != "delete")
            {
                error = "action 须为 list | insert | update | delete";
                return false;
            }

            string sheet = GetString(args, "sheet");
            string anchor = GetString(args, "anchor");
            string imageId = GetString(args, "image_id");
            double? widthCm = TryParseWidthCm(args);

            if (widthCm.HasValue && widthCm.Value <= 0)
            {
                error = "width_cm 须大于 0";
                return false;
            }

            if (action == "list")
            {
                request = new SpreadsheetImageRequest
                {
                    Action = action,
                    Sheet = string.IsNullOrEmpty(sheet) ? null : sheet
                };
                return true;
            }

            if (action == "insert")
            {
                if (string.IsNullOrEmpty(sheet))
                {
                    error = "insert 必须提供 sheet";
                    return false;
                }

                if (!IsSingleAnchor(anchor, out error))
                {
                    return false;
                }

                if (!TryResolveBitmapPath(args, required: true, out string local, out string display, out error))
                {
                    return false;
                }

                request = new SpreadsheetImageRequest
                {
                    Action = action,
                    Sheet = sheet,
                    Anchor = anchor,
                    WidthCm = widthCm,
                    LocalPath = local,
                    PathDisplay = display
                };
                return true;
            }

            if (string.IsNullOrEmpty(imageId) || !ImageIdRegex.IsMatch(imageId))
            {
                error = "必须提供 image_id（I_ + 8 位）";
                return false;
            }

            if (action == "delete")
            {
                request = new SpreadsheetImageRequest
                {
                    Action = action,
                    ImageId = imageId,
                    Sheet = string.IsNullOrEmpty(sheet) ? null : sheet
                };
                return true;
            }

            bool hasPath = args != null && (!string.IsNullOrWhiteSpace(GetString(args, "path"))
                || !string.IsNullOrWhiteSpace(GetString(args, "filename")));
            bool hasAnchor = !string.IsNullOrEmpty(anchor);
            bool hasWidth = widthCm.HasValue;
            if (!hasPath && !hasAnchor && !hasWidth)
            {
                error = "update 须提供 path、anchor、width_cm 至少一项";
                return false;
            }

            if (hasAnchor && !IsSingleAnchor(anchor, out error))
            {
                return false;
            }

            string localPath = null;
            string pathDisplay = null;
            if (hasPath && !TryResolveBitmapPath(args, required: true, out localPath, out pathDisplay, out error))
            {
                return false;
            }

            request = new SpreadsheetImageRequest
            {
                Action = action,
                ImageId = imageId,
                Sheet = string.IsNullOrEmpty(sheet) ? null : sheet,
                Anchor = hasAnchor ? anchor : null,
                WidthCm = widthCm,
                LocalPath = localPath,
                PathDisplay = pathDisplay
            };
            return true;
        }

        private static bool IsSingleAnchor(string anchor, out string error)
        {
            error = null;
            if (string.IsNullOrEmpty(anchor) || anchor.IndexOf('!') >= 0 || anchor.IndexOf(':') >= 0
                || !AnchorRegex.IsMatch(anchor))
            {
                error = "anchor 须为单个 A1（如 B2）";
                return false;
            }

            return true;
        }

        private static bool TryResolveBitmapPath(
            Dictionary<string, object> args,
            bool required,
            out string localPath,
            out string display,
            out string error)
        {
            localPath = null;
            display = null;
            error = null;
            if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath resolved, out string pathError, "path", "filename"))
            {
                if (!required)
                {
                    return true;
                }

                error = pathError;
                return false;
            }

            string ext = Path.GetExtension(resolved.LocalPath ?? "");
            if (!AllowedImageExt.Contains(ext))
            {
                error = "path 后缀限 .png / .jpg / .jpeg / .webp";
                return false;
            }

            if (!File.Exists(resolved.LocalPath))
            {
                error = "文件不存在: " + resolved.Display;
                return false;
            }

            localPath = resolved.LocalPath;
            display = resolved.Display;
            return true;
        }

        internal static string GetString(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return "";
            }

            return Convert.ToString(args[key], CultureInfo.InvariantCulture)?.Trim() ?? "";
        }

        internal static double? TryParseWidthCm(Dictionary<string, object> args)
        {
            if (args == null || !args.ContainsKey("width_cm") || args["width_cm"] == null)
            {
                return null;
            }

            if (args["width_cm"] is double d)
            {
                return d;
            }

            if (args["width_cm"] is float f)
            {
                return (double)f;
            }

            if (args["width_cm"] is int i)
            {
                return i;
            }

            if (double.TryParse(
                    Convert.ToString(args["width_cm"], CultureInfo.InvariantCulture),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out double parsed))
            {
                return parsed;
            }

            return null;
        }
    }
}
