using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace BrowserNavigateTest
{
    internal sealed class NavSession
    {
        public ToolResult Call(Dictionary<string, object> args)
        {
            return DesktopToolPipe.Call("F_browser_navigate", args ?? new Dictionary<string, object>());
        }

        public ToolResult Nav(
            string url,
            string host = null,
            string browser = null,
            object visible = null,
            string channelId = null)
        {
            var args = new Dictionary<string, object>();
            if (url != null)
            {
                args["url"] = url;
            }

            if (host != null)
            {
                args["host"] = host;
            }

            if (browser != null)
            {
                args["browser"] = browser;
            }

            if (visible != null)
            {
                args["visible"] = visible;
            }

            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return Call(args);
        }

        public ToolResult Snap(
            string channelId = null,
            string refId = null,
            string domSupplement = null)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            if (refId != null)
            {
                args["ref"] = refId;
            }

            if (domSupplement != null)
            {
                args["dom_supplement"] = domSupplement;
            }

            return DesktopToolPipe.Call("F_browser_snapshot", args);
        }

        public ToolResult Interact(
            string channelId = null,
            string action = null,
            string refId = null,
            string text = null,
            string key = null,
            string option = null,
            string direction = null)
        {
            var args = new Dictionary<string, object>();
            if (action != null)
            {
                args["action"] = action;
            }

            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            if (refId != null)
            {
                args["ref"] = refId;
            }

            if (text != null)
            {
                args["text"] = text;
            }

            if (key != null)
            {
                args["key"] = key;
            }

            if (option != null)
            {
                args["option"] = option;
            }

            if (direction != null)
            {
                args["direction"] = direction;
            }

            return DesktopToolPipe.Call("F_browser_interact", args);
        }

        public ToolResult Download(
            string channelId = null,
            string refId = null,
            string url = null,
            string path = null)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            if (refId != null)
            {
                args["ref"] = refId;
            }

            if (url != null)
            {
                args["url"] = url;
            }

            if (path != null)
            {
                args["path"] = path;
            }

            return DesktopToolPipe.Call("F_browser_download", args);
        }

        public ToolResult ReadFile(string path)
        {
            var args = new Dictionary<string, object>();
            if (path != null)
            {
                args["path"] = path;
            }

            return DesktopToolPipe.Call("F_read_file", args);
        }

        public static bool IsShortId(string id)
        {
            return !string.IsNullOrEmpty(id)
                && Regex.IsMatch(id, @"^b\d+$", RegexOptions.CultureInvariant)
                && id.IndexOf("browser:", StringComparison.OrdinalIgnoreCase) < 0;
        }
    }
}
