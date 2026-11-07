using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 页面元素到 DOM 节点号。快照补控件和编号截图都走这里，再发号。
    /// </summary>
    internal static class BrowserDomBackend
    {
        internal const string MarkAttribute = "data-yw-clk";

        internal static async Task<int?> TryCreateFrameWorldAsync(
            IBrowserCdp cdp,
            string frameId,
            string sessionId)
        {
            if (cdp == null || string.IsNullOrWhiteSpace(frameId))
            {
                return null;
            }

            try
            {
                var payload = new JObject
                {
                    ["frameId"] = frameId,
                    ["worldName"] = "yiwrite_clickable"
                };
                string json = await cdp
                    .CallCdpAsync("Page.createIsolatedWorld", payload.ToString(Formatting.None), sessionId)
                    .ConfigureAwait(true);
                JObject jo = JObject.Parse(json ?? "{}");
                if (jo["executionContextId"] != null
                    && jo["executionContextId"].Type != JTokenType.Null)
                {
                    return jo["executionContextId"].Value<int>();
                }
            }
            catch
            {
            }

            return null;
        }

        internal static Task<int?> FromMarkAsync(
            IBrowserCdp cdp,
            string mark,
            int? contextId,
            string sessionId)
        {
            string expr = "document.querySelector('[" + MarkAttribute + "=" + Quote(mark) + "]')";
            return FromExpressionAsync(cdp, expr, contextId, sessionId);
        }

        internal static Task<int?> FromCssAsync(
            IBrowserCdp cdp,
            string css,
            int? contextId,
            string sessionId)
        {
            string expression = "document.querySelector(" + JsonConvert.SerializeObject(css ?? "") + ")";
            return FromExpressionAsync(cdp, expression, contextId, sessionId);
        }

        internal static Task<int?> FromXPathAsync(
            IBrowserCdp cdp,
            string xpath,
            int? contextId,
            string sessionId)
        {
            string expression = "document.evaluate("
                + JsonConvert.SerializeObject(xpath ?? "")
                + ", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue";
            return FromExpressionAsync(cdp, expression, contextId, sessionId);
        }

        internal static async Task<int?> FromExpressionAsync(
            IBrowserCdp cdp,
            string expression,
            int? contextId,
            string sessionId)
        {
            if (cdp == null || string.IsNullOrWhiteSpace(expression))
            {
                return null;
            }

            var payload = new JObject
            {
                ["expression"] = expression,
                ["returnByValue"] = false
            };
            if (contextId.HasValue)
            {
                payload["contextId"] = contextId.Value;
            }

            string json;
            try
            {
                json = await cdp
                    .CallCdpAsync("Runtime.evaluate", payload.ToString(Formatting.None), sessionId)
                    .ConfigureAwait(true);
            }
            catch
            {
                return null;
            }

            string objectId = (string)JObject.Parse(json ?? "{}")["result"]?["objectId"];
            if (string.IsNullOrEmpty(objectId))
            {
                return null;
            }

            try
            {
                string described = await cdp.CallCdpAsync(
                    "DOM.describeNode",
                    new JObject { ["objectId"] = objectId }.ToString(Formatting.None),
                    sessionId).ConfigureAwait(true);
                JToken id = JObject.Parse(described ?? "{}")["node"]?["backendNodeId"];
                if (id == null || id.Type == JTokenType.Null)
                {
                    return null;
                }

                return id.Value<int?>();
            }
            catch
            {
                return null;
            }
        }

        internal static async Task ClearSheetMarksAsync(
            IBrowserCdp cdp,
            int? contextId,
            string sessionId)
        {
            if (cdp == null)
            {
                return;
            }

            var payload = new JObject
            {
                ["expression"] =
                    "(function(){var list=document.querySelectorAll('[" + MarkAttribute + "]');"
                    + "for(var i=0;i<list.length;i++){var v=list[i].getAttribute('" + MarkAttribute + "')||'';"
                    + "if(v.indexOf('sh')===0)list[i].removeAttribute('" + MarkAttribute + "');}return true;})()",
                ["returnByValue"] = true
            };
            if (contextId.HasValue)
            {
                payload["contextId"] = contextId.Value;
            }

            try
            {
                await cdp
                    .CallCdpAsync("Runtime.evaluate", payload.ToString(Formatting.None), sessionId)
                    .ConfigureAwait(true);
            }
            catch
            {
            }
        }

        private static string Quote(string s)
        {
            return "\"" + (s ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }
    }
}
