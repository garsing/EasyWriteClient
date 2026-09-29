using System;
using System.Collections.Generic;
using System.Diagnostics;
using WordAddIn1;

namespace BrowserNavigateTest
{
    internal static class NavCases
    {
        public static void Run(TestRun run, NavHttp http, int chromeBefore, int edgeBefore)
        {
            var s = new NavSession();
            RunNeg(run, s);
            string agentId = RunAgent(run, s, http);
            RunAttach(run, s, http, agentId);
            RunNoQuit(run, chromeBefore, edgeBefore);
        }

        private static void RunNeg(TestRun run, NavSession s)
        {
            ChannelRegistry.ClearAll();
            ExpectFail(run, "nav-neg-no-url", s.Call(new Dictionary<string, object>()), "url");
            ExpectFail(run, "nav-neg-empty-url", s.Nav("   "), "url");
            ExpectFail(run, "nav-neg-file", s.Nav("file:///C:/Windows/win.ini"), "http");
            ExpectFail(run, "nav-neg-javascript", s.Nav("javascript:alert(1)"), "http");
            ExpectFail(run, "nav-neg-about", s.Nav("about:blank"), "http");
            ExpectFail(run, "nav-neg-chrome-scheme", s.Nav("chrome://extensions"), "http");
            ExpectFail(run, "nav-neg-bad-host", s.Nav("https://example.com", host: "360"), "host");
            ExpectFail(
                run,
                "nav-neg-hidden-chrome",
                s.Nav("https://example.com", host: "chrome", visible: false),
                "visible");
            ExpectFail(
                run,
                "nav-neg-hidden-edge",
                s.Nav("https://example.com", host: "edge", visible: false),
                "visible");

            ChannelRegistry.ClearAll();
            ChannelRegistry.Register(new FakeWordChannel(), setAsDefault: true);
            ExpectFail(run, "nav-neg-wrong-kind", s.Nav("https://example.com", channelId: "w1"), "渠道");
            ChannelRegistry.ClearAll();
            ExpectFail(run, "nav-neg-unknown-id", s.Nav("https://example.com", channelId: "b999"), "渠道");
        }

        private static string RunAgent(TestRun run, NavSession s, NavHttp http)
        {
            ChannelRegistry.ClearAll();
            ToolResult opened = s.Nav(http.One);
            if (!opened.Success)
            {
                run.CaseFail("nav-agent-open", opened.Error ?? "打开失败");
                run.CaseFail("nav-kind-browser", "依赖 agent 打开");
                run.CaseFail("nav-id-short", "依赖 agent 打开");
                run.CaseFail("nav-tab-uuid", "依赖 agent 打开");
                run.CaseFail("nav-https", "依赖 agent 打开");
                run.CaseFail("nav-title-field", "依赖 agent 打开");
                run.CaseFail("nav-agent-host-yiwrite", "依赖 agent 打开");
                run.CaseFail("nav-host-browser-alias", "依赖 agent 打开");
                run.CaseFail("nav-host-agent-alias", "依赖 agent 打开");
                run.CaseFail("nav-agent-visible", "依赖 agent 打开");
                run.CaseFail("nav-agent-hidden", "依赖 agent 打开");
                run.CaseFail("nav-agent-reuse-id", "依赖 agent 打开");
                run.CaseFail("nav-agent-reuse-default", "依赖 agent 打开");
                return null;
            }

            string id = Str(opened, "channel_id");
            string track = Str(opened, "track");
            string kind = Str(opened, "kind");
            string url = Str(opened, "url");
            string title = Str(opened, "title");
            string tab = Str(opened, "tab_uuid");

            if (string.Equals(track, "agent", StringComparison.OrdinalIgnoreCase)
                && string.Equals(kind, "browser", StringComparison.OrdinalIgnoreCase)
                && NavSession.IsShortId(id)
                && UrlHas(url, "127.0.0.1")
                && !string.IsNullOrEmpty(tab))
            {
                run.CaseOk("nav-agent-open");
            }
            else
            {
                run.CaseFail(
                    "nav-agent-open",
                    "track=" + track + " kind=" + kind + " id=" + id + " url=" + url + " tab=" + tab);
            }

            if (string.Equals(kind, "browser", StringComparison.Ordinal)
                && !string.Equals(kind, "excel", StringComparison.Ordinal)
                && !string.Equals(kind, "word", StringComparison.Ordinal)
                && !string.Equals(kind, "ppt", StringComparison.Ordinal))
            {
                run.CaseOk("nav-kind-browser");
            }
            else
            {
                run.CaseFail("nav-kind-browser", "kind=" + kind);
            }

            if (NavSession.IsShortId(id) && id.IndexOf("browser:", StringComparison.OrdinalIgnoreCase) < 0)
            {
                run.CaseOk("nav-id-short");
            }
            else
            {
                run.CaseFail("nav-id-short", "channel_id=" + id);
            }

            if (!string.IsNullOrEmpty(tab))
            {
                run.CaseOk("nav-tab-uuid");
            }
            else
            {
                run.CaseFail("nav-tab-uuid", "tab_uuid 空");
            }

            if (Uri.TryCreate(url, UriKind.Absolute, out Uri abs)
                && (abs.Scheme == Uri.UriSchemeHttp || abs.Scheme == Uri.UriSchemeHttps))
            {
                run.CaseOk("nav-https");
            }
            else
            {
                run.CaseFail("nav-https", "url=" + url);
            }

            Dictionary<string, object> data = Data(opened);
            if (data != null && data.ContainsKey("title") && title != null)
            {
                run.CaseOk("nav-title-field");
            }
            else
            {
                run.CaseFail("nav-title-field", "无 title 字段");
            }

            ToolResult yi = s.Nav(http.One, host: "yiwrite", channelId: id);
            ExpectAgentOk(run, "nav-agent-host-yiwrite", yi, id);

            ToolResult aliasBrowser = s.Nav(http.One, browser: "yiwrite", channelId: id);
            ExpectAgentOk(run, "nav-host-browser-alias", aliasBrowser, id);

            ToolResult aliasAgent = s.Nav(http.One, host: "agent", channelId: id);
            ExpectAgentOk(run, "nav-host-agent-alias", aliasAgent, id);

            ToolResult vis = s.Nav(http.One, host: "yiwrite", visible: true, channelId: id);
            if (vis.Success && Bool(vis, "visible"))
            {
                run.CaseOk("nav-agent-visible");
            }
            else
            {
                run.CaseFail("nav-agent-visible", vis.Success ? "visible=" + Bool(vis, "visible") : vis.Error);
            }

            ToolResult hid = s.Nav(http.One, host: "yiwrite", visible: false, channelId: id);
            if (hid.Success
                && string.Equals(Str(hid, "track"), "agent", StringComparison.OrdinalIgnoreCase)
                && !Bool(hid, "visible"))
            {
                run.CaseOk("nav-agent-hidden");
            }
            else
            {
                run.CaseFail("nav-agent-hidden", hid.Success ? "track/visible 不符" : hid.Error);
            }

            ToolResult jump = s.Nav(http.Two, host: "yiwrite", channelId: id);
            if (jump.Success
                && string.Equals(Str(jump, "channel_id"), id, StringComparison.Ordinal)
                && UrlHas(Str(jump, "url"), "/two"))
            {
                run.CaseOk("nav-agent-reuse-id");
            }
            else
            {
                run.CaseFail(
                    "nav-agent-reuse-id",
                    jump.Success
                        ? "id=" + Str(jump, "channel_id") + " url=" + Str(jump, "url")
                        : jump.Error);
            }

            ToolResult defJump = s.Nav(http.One);
            if (defJump.Success && string.Equals(Str(defJump, "channel_id"), id, StringComparison.Ordinal))
            {
                run.CaseOk("nav-agent-reuse-default");
            }
            else
            {
                run.CaseFail(
                    "nav-agent-reuse-default",
                    defJump.Success ? "id=" + Str(defJump, "channel_id") : defJump.Error);
            }

            return id;
        }

        private static void RunAttach(TestRun run, NavSession s, NavHttp http, string agentId)
        {
            NavSession.WaitForBridge(5000);
            ToolResult chrome = s.Nav(http.One, host: "chrome", visible: true);
            if (!chrome.Success)
            {
                run.CaseFail("nav-attach-chrome", chrome.Error ?? "打开失败");
                run.CaseFail("nav-track-attach-chrome", "依赖 chrome 附着");
                run.CaseFail("nav-attach-reuse", "依赖 chrome 附着");
            }
            else
            {
                string cid = Str(chrome, "channel_id");
                string track = Str(chrome, "track");
                if (string.Equals(track, "attach", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(Str(chrome, "kind"), "browser", StringComparison.Ordinal)
                    && NavSession.IsShortId(cid)
                    && Bool(chrome, "visible"))
                {
                    run.CaseOk("nav-attach-chrome");
                }
                else
                {
                    run.CaseFail(
                        "nav-attach-chrome",
                        "track=" + track + " id=" + cid + " visible=" + Bool(chrome, "visible"));
                }

                if (string.Equals(track, "attach", StringComparison.OrdinalIgnoreCase))
                {
                    run.CaseOk("nav-track-attach-chrome");
                }
                else
                {
                    run.CaseFail("nav-track-attach-chrome", "track=" + track);
                }

                ToolResult reuse = s.Nav(http.Two, host: "chrome", channelId: cid);
                if (reuse.Success
                    && string.Equals(Str(reuse, "channel_id"), cid, StringComparison.Ordinal)
                    && UrlHas(Str(reuse, "url"), "127.0.0.1"))
                {
                    run.CaseOk("nav-attach-reuse");
                }
                else
                {
                    run.CaseFail(
                        "nav-attach-reuse",
                        reuse.Success
                            ? "id=" + Str(reuse, "channel_id") + " url=" + Str(reuse, "url")
                            : reuse.Error);
                }
            }

            ToolResult edge = s.Nav(http.One, host: "edge", visible: true);
            if (!edge.Success)
            {
                run.CaseFail("nav-attach-edge", edge.Error ?? "打开失败");
            }
            else if (string.Equals(Str(edge, "track"), "attach", StringComparison.OrdinalIgnoreCase)
                && Bool(edge, "visible")
                && NavSession.IsShortId(Str(edge, "channel_id")))
            {
                run.CaseOk("nav-attach-edge");
            }
            else
            {
                run.CaseFail("nav-attach-edge", "track/visible/id 不符");
            }

            ToolResult force = s.Nav(http.One, host: "yiwrite");
            if (!force.Success)
            {
                run.CaseFail("nav-force-yiwrite", force.Error ?? "打开失败");
                return;
            }

            string forceTrack = Str(force, "track");
            string forceId = Str(force, "channel_id");
            bool notStolen = string.IsNullOrEmpty(agentId)
                || string.Equals(forceId, agentId, StringComparison.Ordinal)
                || string.Equals(forceTrack, "agent", StringComparison.OrdinalIgnoreCase);
            if (string.Equals(forceTrack, "agent", StringComparison.OrdinalIgnoreCase) && notStolen)
            {
                run.CaseOk("nav-force-yiwrite");
            }
            else
            {
                run.CaseFail("nav-force-yiwrite", "track=" + forceTrack + " id=" + forceId);
            }
        }

        private static void RunNoQuit(TestRun run, int chromeBefore, int edgeBefore)
        {
            int chromeNow = CountProc("chrome");
            int edgeNow = CountProc("msedge");
            if (chromeNow >= chromeBefore && edgeNow >= edgeBefore)
            {
                run.CaseOk("nav-no-quit");
            }
            else
            {
                run.CaseFail(
                    "nav-no-quit",
                    "chrome " + chromeBefore + "→" + chromeNow + "  edge " + edgeBefore + "→" + edgeNow);
            }
        }

        private static void ExpectAgentOk(TestRun run, string id, ToolResult r, string expectChannel)
        {
            if (r.Success
                && string.Equals(Str(r, "track"), "agent", StringComparison.OrdinalIgnoreCase)
                && (string.IsNullOrEmpty(expectChannel)
                    || string.Equals(Str(r, "channel_id"), expectChannel, StringComparison.Ordinal)))
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, r.Success ? "track/id 不符 id=" + Str(r, "channel_id") : r.Error);
            }
        }

        private static void ExpectFail(TestRun run, string id, ToolResult r, string mustContain)
        {
            if (r != null && !r.Success)
            {
                string err = r.Error ?? "";
                if (string.IsNullOrEmpty(mustContain)
                    || err.IndexOf(mustContain, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    run.CaseOk(id);
                    return;
                }

                run.CaseFail(id, "失败文案不含「" + mustContain + "」: " + err);
                return;
            }

            run.CaseFail(id, r != null && r.Success ? "应失败却成功" : "无回包");
        }

        private static Dictionary<string, object> Data(ToolResult r)
        {
            return r?.Data as Dictionary<string, object>;
        }

        private static string Str(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return "";
            }

            return Convert.ToString(d[key]) ?? "";
        }

        private static bool Bool(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return false;
            }

            object raw = d[key];
            if (raw is bool b)
            {
                return b;
            }

            bool.TryParse(Convert.ToString(raw), out bool parsed);
            return parsed;
        }

        private static bool UrlHas(string url, string part)
        {
            return !string.IsNullOrEmpty(url)
                && url.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static int CountProc(string name)
        {
            try
            {
                return Process.GetProcessesByName(name).Length;
            }
            catch
            {
                return 0;
            }
        }
    }
}
