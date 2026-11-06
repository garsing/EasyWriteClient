using System;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 把 iframe 内的视口坐标换成截图/点击所用画面上的坐标。同源 iframe 没有单独会话，按框架树把每一层框的位置都加上。
    /// </summary>
    internal static class BrowserFrameOrigin
    {
        internal struct Offset
        {
            public double X;
            public double Y;
        }

        internal static async Task<Offset> ContentAsync(
            IBrowserCdp cdp,
            string objectId,
            string sessionId,
            string frameId)
        {
            if (string.IsNullOrEmpty(sessionId) && !string.IsNullOrWhiteSpace(frameId))
            {
                return await FrameTreeOriginAsync(cdp, frameId).ConfigureAwait(true);
            }

            return await FrameElementOriginAsync(cdp, objectId, sessionId).ConfigureAwait(true);
        }

        private static async Task<Offset> FrameTreeOriginAsync(IBrowserCdp cdp, string frameId)
        {
            var origin = new Offset();
            if (cdp == null || string.IsNullOrWhiteSpace(frameId))
            {
                return origin;
            }

            BrowserFrameMap map;
            try
            {
                map = await BrowserFrameResolver.LoadAsync(cdp).ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                ScriptDialogTrace.Write("frame-origin-error " + ex.GetType().Name);
                return origin;
            }

            string id = frameId;
            int guard = 0;
            var parts = new StringBuilder();
            while (guard < 8 && !string.IsNullOrEmpty(id))
            {
                if (!string.IsNullOrEmpty(map.RootFrameId)
                    && string.Equals(id, map.RootFrameId, StringComparison.Ordinal))
                {
                    break;
                }

                if (!map.ParentOf.ContainsKey(id))
                {
                    break;
                }

                Offset hop = await IframeBoxInParentAsync(cdp, id).ConfigureAwait(true);
                origin.X += hop.X;
                origin.Y += hop.Y;
                parts.Append(" +")
                    .Append(hop.X.ToString("0.##", CultureInfo.InvariantCulture))
                    .Append(",")
                    .Append(hop.Y.ToString("0.##", CultureInfo.InvariantCulture));
                id = map.ParentOf[id];
                guard++;
            }

            ScriptDialogTrace.Write("frame-origin" + parts);
            return origin;
        }

        private static async Task<Offset> IframeBoxInParentAsync(IBrowserCdp cdp, string frameId)
        {
            var origin = new Offset();
            try
            {
                var payload = new JObject { ["frameId"] = frameId };
                string ownerJson = await cdp
                    .CallCdpAsync("DOM.getFrameOwner", payload.ToString(Newtonsoft.Json.Formatting.None))
                    .ConfigureAwait(true);
                JObject owner = JObject.Parse(ownerJson ?? "{}");
                if (owner["backendNodeId"] == null
                    || !int.TryParse(owner["backendNodeId"].ToString(), out int backend)
                    || backend == 0)
                {
                    ScriptDialogTrace.Write("frame-origin-miss frame=" + frameId);
                    return origin;
                }

                string objectId = await ResolveInParentAsync(cdp, backend).ConfigureAwait(true);
                string json = await BrowserInteractEngine.CallFunctionOnAsync(
                    cdp,
                    objectId,
                    @"function() {
  var rect = this.getBoundingClientRect();
  return {
    x: rect.left + (this.clientLeft || 0),
    y: rect.top + (this.clientTop || 0)
  };
}",
                    null,
                    null).ConfigureAwait(true);
                var value = JObject.Parse(json ?? "{}")["result"]?["value"] as JObject;
                if (value != null)
                {
                    if (value["x"] != null)
                    {
                        origin.X = (double)value["x"];
                    }

                    if (value["y"] != null)
                    {
                        origin.Y = (double)value["y"];
                    }
                }
            }
            catch (Exception ex)
            {
                ScriptDialogTrace.Write(
                    "frame-origin-miss frame=" + frameId
                    + " " + ex.GetType().Name
                    + " " + (ex.Message ?? ""));
            }

            return origin;
        }

        private static async Task<Offset> FrameElementOriginAsync(
            IBrowserCdp cdp,
            string objectId,
            string sessionId)
        {
            var origin = new Offset();
            if (cdp == null || string.IsNullOrEmpty(objectId))
            {
                return origin;
            }

            try
            {
                string json = await BrowserInteractEngine.CallFunctionOnAsync(
                    cdp,
                    objectId,
                    @"function() {
  var dx = 0, dy = 0, win = null, guard = 0;
  try { win = this.ownerDocument && this.ownerDocument.defaultView; } catch (e0) {}
  while (win && guard < 8) {
    var frameEl = null;
    try { frameEl = win.frameElement; } catch (e1) { break; }
    if (!frameEl || !frameEl.getBoundingClientRect) break;
    var rect = frameEl.getBoundingClientRect();
    dx += rect.left + (frameEl.clientLeft || 0);
    dy += rect.top + (frameEl.clientTop || 0);
    var parent = null;
    try { parent = win.parent; } catch (e2) { break; }
    if (!parent || parent === win) break;
    win = parent;
    guard++;
  }
  return { dx: dx, dy: dy };
}",
                    null,
                    sessionId).ConfigureAwait(true);
                var jo = JObject.Parse(json ?? "{}");
                var value = jo["result"]?["value"] as JObject;
                if (value != null)
                {
                    if (value["dx"] != null)
                    {
                        origin.X = (double)value["dx"];
                    }

                    if (value["dy"] != null)
                    {
                        origin.Y = (double)value["dy"];
                    }
                }
            }
            catch (Exception ex)
            {
                ScriptDialogTrace.Write("frame-origin-error " + ex.GetType().Name);
            }

            return origin;
        }

        private static async Task<string> ResolveInParentAsync(IBrowserCdp cdp, int backendNodeId)
        {
            string json = await cdp.CallCdpAsync(
                "DOM.resolveNode",
                "{\"backendNodeId\":" + backendNodeId.ToString(CultureInfo.InvariantCulture) + "}",
                null).ConfigureAwait(true);
            var jo = JObject.Parse(json ?? "{}");
            string objectId = (string)jo["object"]?["objectId"];
            if (string.IsNullOrEmpty(objectId))
            {
                throw new InvalidOperationException("无法解析 iframe 框");
            }

            return objectId;
        }
    }
}
