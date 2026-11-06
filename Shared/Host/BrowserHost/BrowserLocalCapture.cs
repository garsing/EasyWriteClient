using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 按快照里已打印的祖先裁一块图。易写浏览器与附着页都走 CDP。
    /// </summary>
    internal static class BrowserLocalCapture
    {
        private const string NoRegion = "没有可截区域";
        private const string NotEnough = "往上没有那么多层";
        private const string StaleRef = "ref 无效或已过期，请重新 F_browser_snapshot";

        public static async Task<BrowserCaptureResult> CaptureAsync(
            IBrowserCdp cdp,
            BrowserChannel channel,
            string refId,
            int up)
        {
            if (cdp == null || channel == null)
            {
                return BrowserCaptureResult.Fail("无浏览器渠道");
            }

            if (up < 1)
            {
                return BrowserCaptureResult.Fail("up 须为从 1 起的正整数");
            }

            if (!BrowserRefStore.TryGet(channel.ChannelId, refId, out BrowserRefEntry entry) || entry == null)
            {
                return BrowserCaptureResult.Fail(StaleRef);
            }

            if (entry.IsScriptDialog)
            {
                return BrowserCaptureResult.Fail(StaleRef);
            }

            if (entry.PrintedAncestors == null || entry.PrintedAncestors.Count < up)
            {
                return BrowserCaptureResult.Fail(NotEnough);
            }

            BrowserPrintedAncestor ancestor = entry.PrintedAncestors[up - 1];
            if (ancestor == null || !ancestor.HasBox)
            {
                return BrowserCaptureResult.Fail(NoRegion);
            }

            string sessionId = string.IsNullOrWhiteSpace(entry.CdpSessionId) ? null : entry.CdpSessionId;
            string objectId = null;
            try
            {
                objectId = await ResolveObjectAsync(cdp, sessionId, ancestor).ConfigureAwait(true);
                if (string.IsNullOrEmpty(objectId))
                {
                    return BrowserCaptureResult.Fail(NoRegion);
                }

                JObject box = await CallOnAsync(cdp, sessionId, objectId, MeasureFn, null).ConfigureAwait(true);
                if (!BoxOk(box))
                {
                    return BrowserCaptureResult.Fail(NoRegion);
                }

                bool scrolled = false;
                byte[] jpeg = null;
                try
                {
                    if (box.Value<bool?>("clipped") == true)
                    {
                        scrolled = true;
                        JObject marked = await CallOnAsync(cdp, sessionId, objectId, ScrollFn, null).ConfigureAwait(true);
                        string token = marked?.Value<string>("token");
                        if (string.IsNullOrEmpty(token))
                        {
                            return BrowserCaptureResult.Fail(NoRegion);
                        }

                        scrolled = true;
                        box = await CallOnAsync(cdp, sessionId, objectId, MeasureFn, null).ConfigureAwait(true);
                        if (!BoxOk(box) || box.Value<bool?>("clipped") == true)
                        {
                            return BrowserCaptureResult.Fail(NoRegion);
                        }

                        jpeg = await ShootAsync(
                            cdp,
                            sessionId,
                            box.Value<double?>("left") ?? 0,
                            box.Value<double?>("top") ?? 0,
                            box.Value<double?>("w") ?? 0,
                            box.Value<double?>("h") ?? 0,
                            beyond: false).ConfigureAwait(true);
                    }
                    else
                    {
                        jpeg = await ShootAsync(
                            cdp,
                            sessionId,
                            box.Value<double?>("docX") ?? 0,
                            box.Value<double?>("docY") ?? 0,
                            box.Value<double?>("w") ?? 0,
                            box.Value<double?>("h") ?? 0,
                            beyond: true).ConfigureAwait(true);
                    }
                }
                finally
                {
                    if (scrolled)
                    {
                        try
                        {
                            await CallOnAsync(cdp, sessionId, objectId, RestoreFn, null).ConfigureAwait(true);
                        }
                        catch
                        {
                            // 图已经拿到时，写回失败仍算成功。
                        }
                    }
                }

                if (jpeg == null || jpeg.Length == 0)
                {
                    return BrowserCaptureResult.Fail(NoRegion);
                }

                using (var stream = new MemoryStream(jpeg))
                using (var bitmap = new Bitmap(stream))
                {
                    ImageCaptureCompressor.CompressedImage image = ImageCaptureCompressor.Compress(bitmap);
                    return BrowserCaptureResult.Ok(channel, cdp.CurrentUrl, cdp.CurrentTitle, image);
                }
            }
            catch (Exception ex)
            {
                return BrowserCaptureResult.Fail(string.IsNullOrWhiteSpace(ex.Message) ? NoRegion : ex.Message);
            }
            finally
            {
                if (!string.IsNullOrEmpty(objectId))
                {
                    try
                    {
                        var release = new JObject { ["objectId"] = objectId };
                        await cdp.CallCdpAsync("Runtime.releaseObject", release.ToString(Formatting.None), sessionId)
                            .ConfigureAwait(true);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static bool BoxOk(JObject box)
        {
            if (box == null || box.Value<bool?>("ok") != true)
            {
                return false;
            }

            double w = box.Value<double?>("w") ?? 0;
            double h = box.Value<double?>("h") ?? 0;
            return w > 0 && h > 0;
        }

        private static async Task<string> ResolveObjectAsync(
            IBrowserCdp cdp,
            string sessionId,
            BrowserPrintedAncestor ancestor)
        {
            if (ancestor.BackendDomNodeId.HasValue)
            {
                try
                {
                    await cdp.CallCdpAsync(
                        "DOM.getDocument",
                        "{\"depth\":0}",
                        sessionId).ConfigureAwait(true);
                }
                catch
                {
                }

                var resolve = new JObject { ["backendNodeId"] = ancestor.BackendDomNodeId.Value };
                string json = await cdp.CallCdpAsync("DOM.resolveNode", resolve.ToString(Formatting.None), sessionId)
                    .ConfigureAwait(true);
                JObject jo = Parse(json);
                string objectId = (string)jo["object"]?["objectId"];
                if (!string.IsNullOrEmpty(objectId))
                {
                    return objectId;
                }
            }

            if (string.IsNullOrWhiteSpace(ancestor.AttachCssPath))
            {
                return null;
            }

            string expression = "document.querySelector(" + JsonConvert.SerializeObject(ancestor.AttachCssPath) + ")";
            var eval = new JObject
            {
                ["expression"] = expression,
                ["returnByValue"] = false
            };
            string evalJson = await cdp.CallCdpAsync("Runtime.evaluate", eval.ToString(Formatting.None), sessionId)
                .ConfigureAwait(true);
            JObject evalObj = Parse(evalJson);
            if (evalObj["exceptionDetails"] != null && evalObj["exceptionDetails"].Type != JTokenType.Null)
            {
                return null;
            }

            return (string)evalObj["result"]?["objectId"];
        }

        private static async Task<JObject> CallOnAsync(
            IBrowserCdp cdp,
            string sessionId,
            string objectId,
            string functionDeclaration,
            string token)
        {
            var payload = new JObject
            {
                ["objectId"] = objectId,
                ["functionDeclaration"] = functionDeclaration,
                ["returnByValue"] = true
            };
            if (token != null)
            {
                payload["arguments"] = new JArray { new JObject { ["value"] = token } };
            }

            string json = await cdp.CallCdpAsync("Runtime.callFunctionOn", payload.ToString(Formatting.None), sessionId)
                .ConfigureAwait(true);
            JObject jo = Parse(json);
            JToken ex = jo["exceptionDetails"];
            if (ex != null && ex.Type != JTokenType.Null)
            {
                string msg = (string)ex["exception"]?["description"] ?? (string)ex["text"] ?? NoRegion;
                throw new InvalidOperationException(msg);
            }

            return jo["result"]?["value"] as JObject ?? new JObject();
        }

        private static async Task<byte[]> ShootAsync(
            IBrowserCdp cdp,
            string sessionId,
            double x,
            double y,
            double width,
            double height,
            bool beyond)
        {
            if (width <= 0 || height <= 0)
            {
                throw new InvalidOperationException(NoRegion);
            }

            if (x < 0)
            {
                width += x;
                x = 0;
            }

            if (y < 0)
            {
                height += y;
                y = 0;
            }

            if (width <= 0 || height <= 0)
            {
                throw new InvalidOperationException(NoRegion);
            }

            var clip = new JObject
            {
                ["x"] = Math.Round(x, 2),
                ["y"] = Math.Round(y, 2),
                ["width"] = Math.Round(width, 2),
                ["height"] = Math.Round(height, 2),
                ["scale"] = 1
            };
            var payload = new JObject
            {
                ["format"] = "jpeg",
                ["quality"] = 80,
                ["fromSurface"] = true,
                ["clip"] = clip
            };
            if (beyond)
            {
                payload["captureBeyondViewport"] = true;
            }

            Task<string> cdpTask = cdp.CallCdpAsync(
                "Page.captureScreenshot",
                payload.ToString(Formatting.None),
                sessionId);
            Task winner = await Task.WhenAny(cdpTask, Task.Delay(TimeSpan.FromSeconds(20))).ConfigureAwait(true);
            if (!ReferenceEquals(winner, cdpTask))
            {
                throw new TimeoutException("局部截图超时（20 秒未返回）");
            }

            JObject shot = Parse(await cdpTask.ConfigureAwait(true));
            string data = (string)shot["data"];
            if (string.IsNullOrEmpty(data))
            {
                throw new InvalidOperationException(NoRegion);
            }

            return Convert.FromBase64String(data);
        }

        private static JObject Parse(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return new JObject();
            }

            return JObject.Parse(json);
        }

        private const string MeasureFn =
            "function(){"
            + "var el=this;"
            + "if(el&&el.nodeType===9)el=el.documentElement;"
            + "if(el&&el.nodeType===3)el=el.parentElement;"
            + "if(!el||!el.getBoundingClientRect)return{ok:false};"
            + "var rect=el.getBoundingClientRect();var clipped=false;var n=el.parentElement;"
            + "while(n){"
            + "var st=window.getComputedStyle(n);var oy=st.overflowY;var ox=st.overflowX;"
            + "var can=oy==='auto'||oy==='scroll'||oy==='hidden'||ox==='auto'||ox==='scroll'||ox==='hidden';"
            + "if(can){"
            + "var cr=n.getBoundingClientRect();"
            + "var top=rect.top-cr.top;var left=rect.left-cr.left;"
            + "var ih=n.clientHeight;var iw=n.clientWidth;"
            + "if(top+rect.height<=1||left+rect.width<=1||top>=ih-1||left>=iw-1)clipped=true;"
            + "}"
            + "n=n.parentElement;}"
            + "var sx=window.scrollX||0;var sy=window.scrollY||0;"
            + "return{ok:true,w:rect.width,h:rect.height,top:rect.top,left:rect.left,docX:rect.left+sx,docY:rect.top+sy,clipped:clipped};"
            + "}";

        private const string ScrollFn =
            "function(){"
            + "var el=this;"
            + "if(el&&el.nodeType===9)el=el.documentElement;"
            + "if(el&&el.nodeType===3)el=el.parentElement;"
            + "if(!el||!el.parentElement&&el.nodeType!==1)return{};"
            + "var token='yw'+Date.now()+Math.floor(Math.random()*1000);"
            + "function mark(root){var n=root;while(n){"
            + "var st;try{st=n.ownerDocument.defaultView.getComputedStyle(n);}catch(e){st=null;}"
            + "if(st){var oy=st.overflowY;var ox=st.overflowX;"
            + "var can=oy==='auto'||oy==='scroll'||oy==='hidden'||ox==='auto'||ox==='scroll'||ox==='hidden';"
            + "if(can&&(n.scrollHeight>n.clientHeight+1||n.scrollWidth>n.clientWidth+1)){"
            + "n.setAttribute('data-yw-shot',token);n.setAttribute('data-yw-top',String(n.scrollTop||0));n.setAttribute('data-yw-left',String(n.scrollLeft||0));}}"
            + "n=n.parentElement;}}"
            + "mark(el.parentElement);"
            + "var w=window;var guard=0;"
            + "while(w&&guard<8){"
            + "w.__ywShot=token;w.__ywX=w.scrollX||0;w.__ywY=w.scrollY||0;"
            + "try{if(w.frameElement)mark(w.frameElement.parentElement);}catch(e1){}"
            + "guard++;try{if(!w.parent||w.parent===w)break;w=w.parent;}catch(e2){break;}}"
            + "try{el.scrollIntoView({block:'nearest',inline:'nearest'});}catch(e3){}"
            + "return{token:token};"
            + "}";

        private const string RestoreFn =
            "function(){"
            + "function restoreDoc(doc){if(!doc||!doc.querySelectorAll)return;"
            + "var nodes=doc.querySelectorAll('[data-yw-shot]');"
            + "for(var i=0;i<nodes.length;i++){var n=nodes[i];"
            + "n.scrollTop=Number(n.getAttribute('data-yw-top')||0);"
            + "n.scrollLeft=Number(n.getAttribute('data-yw-left')||0);"
            + "n.removeAttribute('data-yw-shot');n.removeAttribute('data-yw-top');n.removeAttribute('data-yw-left');}}"
            + "var w=window;var guard=0;"
            + "while(w&&guard<8){"
            + "if(w.__ywShot){try{w.scrollTo(w.__ywX||0,w.__ywY||0);}catch(e){}try{delete w.__ywShot;delete w.__ywX;delete w.__ywY;}catch(e2){}}"
            + "try{restoreDoc(w.document);if(w.frameElement)restoreDoc(w.frameElement.ownerDocument);}catch(e3){}"
            + "guard++;try{if(!w.parent||w.parent===w)break;w=w.parent;}catch(e4){break;}}"
            + "return{ok:true};"
            + "}";
    }
}
