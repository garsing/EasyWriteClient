using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordAddIn1;
using WordAddIn1.BrowserHost;

namespace BrowserNavigateTest
{
    internal sealed class FakeWordChannel : IOperationChannel
    {
        public string ChannelId { get { return "w1"; } }

        public ChannelKind Kind { get { return ChannelKind.Word; } }
    }

    internal sealed class NavSession
    {
        private readonly Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> _tools;

        public NavSession()
        {
            _tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.Ordinal);
            F_BrowserNavigateTool.Register(_tools, null);
            ExtensionHost.EnsureStarted();
            WaitForBridge(8000);
        }

        /// <summary>已开的 Chrome/Edge 扩展把 NM 接到本进程管道上需要一两秒。</summary>
        public static void WaitForBridge(int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (BrowserBridgePipeServer.SnapshotClients().Count > 0)
                {
                    return;
                }

                Application.DoEvents();
                Thread.Sleep(50);
            }
        }

        public ToolResult Call(Dictionary<string, object> args)
        {
            try
            {
                return Pump(_tools["F_browser_navigate"](args ?? new Dictionary<string, object>()));
            }
            catch (Exception ex)
            {
                return new ToolResult { Success = false, Error = ex.Message };
            }
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

        public static bool IsShortId(string id)
        {
            return !string.IsNullOrEmpty(id)
                && Regex.IsMatch(id, @"^b\d+$", RegexOptions.CultureInvariant)
                && id.IndexOf("browser:", StringComparison.OrdinalIgnoreCase) < 0;
        }

        private static ToolResult Pump(Task<ToolResult> task)
        {
            var sw = Stopwatch.StartNew();
            while (!task.IsCompleted)
            {
                Application.DoEvents();
                Thread.Sleep(20);
                if (sw.ElapsedMilliseconds > 120000)
                {
                    throw new TimeoutException("F_browser_navigate 超过 120s");
                }
            }

            return task.GetAwaiter().GetResult();
        }
    }
}
