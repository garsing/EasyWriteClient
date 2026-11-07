using System.Collections.Generic;
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

        internal static async Task<int?> FromXPathAsync(
            IBrowserCdp cdp,
            string xpath,
            int? contextId,
            string sessionId)
        {
            if (cdp == null || string.IsNullOrWhiteSpace(xpath))
            {
                return null;
            }

            var parts = SplitDocuments(xpath);
            if (parts.Count == 0)
            {
                return null;
            }

            if (parts.Count == 1)
            {
                return await EvalSegmentAsync(cdp, parts[0], contextId, sessionId).ConfigureAwait(true);
            }

            string objectId = await EvalSegmentObjectAsync(cdp, parts[0], null, null).ConfigureAwait(true);
            string session = null;
            for (int i = 1; i < parts.Count && !string.IsNullOrEmpty(objectId); i++)
            {
                objectId = await EvalInsideFrameAsync(cdp, objectId, parts[i], session).ConfigureAwait(true);
            }

            if (string.IsNullOrEmpty(objectId))
            {
                return null;
            }

            return await DescribeBackendAsync(cdp, objectId, session).ConfigureAwait(true);
        }

        private static List<string> SplitDocuments(string xpath)
        {
            var parts = new List<string>();
            int start = 0;
            while (start < xpath.Length)
            {
                int next = start + 1 < xpath.Length
                    ? xpath.IndexOf("/html", start + 1, System.StringComparison.Ordinal)
                    : -1;
                if (next < 0)
                {
                    parts.Add(xpath.Substring(start));
                    break;
                }

                parts.Add(xpath.Substring(start, next - start));
                start = next;
            }

            return parts;
        }

        private static Task<int?> EvalSegmentAsync(
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

        private static async Task<string> EvalSegmentObjectAsync(
            IBrowserCdp cdp,
            string xpath,
            int? contextId,
            string sessionId)
        {
            string expression = "document.evaluate("
                + JsonConvert.SerializeObject(xpath ?? "")
                + ", document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue";
            var payload = new JObject
            {
                ["expression"] = expression,
                ["returnByValue"] = false
            };
            if (contextId.HasValue)
            {
                payload["contextId"] = contextId.Value;
            }

            try
            {
                string json = await cdp
                    .CallCdpAsync("Runtime.evaluate", payload.ToString(Formatting.None), sessionId)
                    .ConfigureAwait(true);
                return (string)JObject.Parse(json ?? "{}")["result"]?["objectId"];
            }
            catch
            {
                return null;
            }
        }

        private static async Task<string> EvalInsideFrameAsync(
            IBrowserCdp cdp,
            string frameObjectId,
            string xpath,
            string sessionId)
        {
            try
            {
                var args = new JArray { new JObject { ["value"] = xpath ?? "" } };
                string json = await BrowserInteractEngine.CallFunctionOnAsync(
                    cdp,
                    frameObjectId,
                    "function(xpath){var doc=null;try{doc=this.contentDocument;}catch(e){doc=null;}"
                    + "if(!doc)return null;try{return doc.evaluate(xpath,doc,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;}"
                    + "catch(e2){return null;}}",
                    args.ToString(Formatting.None),
                    sessionId,
                    returnByValue: false).ConfigureAwait(true);
                string objectId = (string)JObject.Parse(json ?? "{}")["result"]?["objectId"];
                if (!string.IsNullOrEmpty(objectId))
                {
                    return objectId;
                }
            }
            catch
            {
            }

            try
            {
                string described = await cdp.CallCdpAsync(
                    "DOM.describeNode",
                    new JObject { ["objectId"] = frameObjectId }.ToString(Formatting.None),
                    sessionId).ConfigureAwait(true);
                string frameId = (string)JObject.Parse(described ?? "{}")["node"]?["frameId"];
                if (string.IsNullOrWhiteSpace(frameId))
                {
                    return null;
                }

                int? world = await TryCreateFrameWorldAsync(cdp, frameId, sessionId).ConfigureAwait(true);
                return await EvalSegmentObjectAsync(cdp, xpath, world, sessionId).ConfigureAwait(true);
            }
            catch
            {
                return null;
            }
        }

        private static async Task<int?> DescribeBackendAsync(
            IBrowserCdp cdp,
            string objectId,
            string sessionId)
        {
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
