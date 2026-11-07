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
            int? up)
        {
            LocalPlate plate = await CapturePlateAsync(cdp, channel, refId, up, collectNodes: false)
                .ConfigureAwait(true);
            if (plate.Error != null)
            {
                return BrowserCaptureResult.Fail(plate.Error);
            }

            return FinishImage(cdp, channel, plate.Jpeg, null);
        }

        internal static async Task<LocalPlate> CapturePlateAsync(
            IBrowserCdp cdp,
            BrowserChannel channel,
            string refId,
            int? up,
            bool collectNodes)
        {
            if (cdp == null || channel == null)
            {
                return LocalPlate.Fail("无浏览器渠道");
            }

            if (up.HasValue && up.Value < 1)
            {
                return LocalPlate.Fail("up 须为从 1 起的正整数");
            }

            if (!BrowserRefStore.TryGet(channel.ChannelId, refId, out BrowserRefEntry entry) || entry == null)
            {
                return LocalPlate.Fail(StaleRef);
            }

            if (entry.IsScriptDialog)
            {
                return LocalPlate.Fail(StaleRef);
            }

            BrowserPrintedAncestor ancestor;
            if (!up.HasValue)
            {
                ancestor = OwnBox(entry);
                if (ancestor == null)
                {
                    return LocalPlate.Fail(NoRegion);
                }
            }
            else
            {
                if (entry.PrintedAncestors == null || entry.PrintedAncestors.Count < up.Value)
                {
                    return LocalPlate.Fail(NotEnough);
                }

                ancestor = entry.PrintedAncestors[up.Value - 1];
                if (ancestor == null || !ancestor.HasBox)
                {
                    return LocalPlate.Fail(NoRegion);
                }
            }

            string sessionId = string.IsNullOrWhiteSpace(entry.CdpSessionId) ? null : entry.CdpSessionId;
            string objectId = null;
            try
            {
                objectId = await ResolveObjectAsync(cdp, sessionId, ancestor).ConfigureAwait(true);
                if (string.IsNullOrEmpty(objectId))
                {
                    return LocalPlate.Fail(NoRegion);
                }

                if (collectNodes)
                {
                    string widened = await CallElementAsync(cdp, sessionId, objectId, WidenFn).ConfigureAwait(true);
                    if (!string.IsNullOrEmpty(widened) && widened != objectId)
                    {
                        await ReleaseObjectAsync(cdp, sessionId, objectId).ConfigureAwait(true);
                        objectId = widened;
                    }
                }

                JObject box = await CallOnAsync(cdp, sessionId, objectId, MeasureFn, null).ConfigureAwait(true);
                if (!BoxOk(box))
                {
                    return LocalPlate.Fail(NoRegion);
                }

                bool scrolled = false;
                byte[] jpeg = null;
                JArray items = null;
                try
                {
                    if (box.Value<bool?>("clipped") == true)
                    {
                        scrolled = true;
                        JObject marked = await CallOnAsync(cdp, sessionId, objectId, ScrollFn, null).ConfigureAwait(true);
                        string token = marked?.Value<string>("token");
                        if (string.IsNullOrEmpty(token))
                        {
                            return LocalPlate.Fail(NoRegion);
                        }

                        box = await CallOnAsync(cdp, sessionId, objectId, MeasureFn, null).ConfigureAwait(true);
                        if (!BoxOk(box) || box.Value<bool?>("clipped") == true)
                        {
                            return LocalPlate.Fail(NoRegion);
                        }
                    }

                    if (collectNodes)
                    {
                        JObject listed = await CallOnAsync(cdp, sessionId, objectId, ListFn, null).ConfigureAwait(true);
                        if (listed.Value<bool?>("ok") != true)
                        {
                            return LocalPlate.Fail(NoRegion);
                        }

                        items = listed["items"] as JArray ?? new JArray();
                    }

                    BrowserFrameOrigin.Offset origin = await BrowserFrameOrigin
                        .ContentAsync(cdp, objectId, sessionId, entry.FrameId)
                        .ConfigureAwait(true);
                    bool nested = Math.Abs(origin.X) > 0.5 || Math.Abs(origin.Y) > 0.5;
                    if (nested || box.Value<bool?>("clipped") == true)
                    {
                        jpeg = await ShootAsync(
                            cdp,
                            sessionId,
                            (box.Value<double?>("left") ?? 0) + origin.X,
                            (box.Value<double?>("top") ?? 0) + origin.Y,
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
                    return LocalPlate.Fail(NoRegion);
                }

                return new LocalPlate
                {
                    Jpeg = jpeg,
                    CssW = box.Value<double?>("w") ?? 0,
                    CssH = box.Value<double?>("h") ?? 0,
                    Items = items,
                    Entry = entry
                };
            }
            catch (Exception ex)
            {
                return LocalPlate.Fail(string.IsNullOrWhiteSpace(ex.Message) ? NoRegion : ex.Message);
            }
            finally
            {
                await ReleaseObjectAsync(cdp, sessionId, objectId).ConfigureAwait(true);
            }
        }

        private static async Task ReleaseObjectAsync(IBrowserCdp cdp, string sessionId, string objectId)
        {
            if (cdp == null || string.IsNullOrEmpty(objectId))
            {
                return;
            }

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

        private static async Task<string> CallElementAsync(
            IBrowserCdp cdp,
            string sessionId,
            string objectId,
            string functionDeclaration)
        {
            var payload = new JObject
            {
                ["objectId"] = objectId,
                ["functionDeclaration"] = functionDeclaration,
                ["returnByValue"] = false
            };
            string json = await cdp.CallCdpAsync("Runtime.callFunctionOn", payload.ToString(Formatting.None), sessionId)
                .ConfigureAwait(true);
            JObject jo = Parse(json);
            JToken ex = jo["exceptionDetails"];
            if (ex != null && ex.Type != JTokenType.Null)
            {
                string msg = (string)ex["exception"]?["description"] ?? (string)ex["text"] ?? NoRegion;
                throw new InvalidOperationException(msg);
            }

            return (string)jo["result"]?["objectId"];
        }

        internal static BrowserCaptureResult FinishImage(
            IBrowserCdp cdp,
            BrowserChannel channel,
            byte[] jpeg,
            string note)
        {
            if (jpeg == null || jpeg.Length == 0)
            {
                return BrowserCaptureResult.Fail(NoRegion);
            }

            using (var stream = new MemoryStream(jpeg))
            using (var bitmap = new Bitmap(stream))
            {
                ImageCaptureCompressor.CompressedImage image = ImageCaptureCompressor.Compress(bitmap);
                return BrowserCaptureResult.Ok(channel, cdp.CurrentUrl, cdp.CurrentTitle, image, note);
            }
        }

        internal sealed class LocalPlate
        {
            public string Error;
            public byte[] Jpeg;
            public double CssW;
            public double CssH;
            public JArray Items;
            public BrowserRefEntry Entry;

            public static LocalPlate Fail(string error)
            {
                return new LocalPlate { Error = error ?? NoRegion };
            }
        }

        private static BrowserPrintedAncestor OwnBox(BrowserRefEntry entry)
        {
            if (entry == null)
            {
                return null;
            }

            if (entry.BackendDomNodeId.HasValue)
            {
                return new BrowserPrintedAncestor { BackendDomNodeId = entry.BackendDomNodeId };
            }

            if (!string.IsNullOrWhiteSpace(entry.AttachCssPath))
            {
                return new BrowserPrintedAncestor { AttachCssPath = entry.AttachCssPath };
            }

            return null;
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

        /// <summary>
        /// 文字节点的盒子包不住同一行左侧的箭头和方框。编号时把总图扩到仍不超过 15 个可编号元素的那一层，页面本身不扩进去。15 与 BrowserNodeSheet.MaxNodes 一致。
        /// </summary>
        private const string WidenFn =
            "function(){"
            + "var el=this;"
            + "if(el&&el.nodeType===9)el=el.documentElement;"
            + "if(el&&el.nodeType===3)el=el.parentElement;"
            + "if(!el||!el.parentElement)return el;"
            + "function hidden(node){var st;try{st=getComputedStyle(node);}catch(e){return true;}"
            + "return !st||st.display==='none'||st.visibility==='hidden';}"
            + "function svgPart(node){var tag=(node.tagName||'').toLowerCase();"
            + "return tag==='path'||tag==='rect'||tag==='circle'||tag==='line'||tag==='polyline'||tag==='polygon'||tag==='ellipse';}"
            + "function skipTag(node){var tag=(node.tagName||'').toLowerCase();"
            + "return tag==='script'||tag==='style'||tag==='noscript'||tag==='br';}"
            + "function count(root){var rr=root.getBoundingClientRect();"
            + "function intersects(r){return r.width>1&&r.height>1&&r.left<rr.right-0.5&&r.right>rr.left+0.5&&r.top<rr.bottom-0.5&&r.bottom>rr.top+0.5;}"
            + "var all=root.querySelectorAll?root.querySelectorAll('*'):[];var els=[];"
            + "for(var i=0;i<all.length;i++){var node=all[i];if(node===root||skipTag(node)||svgPart(node)||hidden(node))continue;"
            + "if(!intersects(node.getBoundingClientRect()))continue;els.push(node);}"
            + "function isWrap(node){var kids=node.children;if(!kids||!kids.length)return false;"
            + "var minL=1e9,minT=1e9,maxR=-1e9,maxB=-1e9,n=0;"
            + "for(var k=0;k<kids.length;k++){var c=kids[k];if(skipTag(c)||svgPart(c)||hidden(c))continue;"
            + "var kb=c.getBoundingClientRect();if(kb.width<=1||kb.height<=1)continue;n++;"
            + "if(kb.left<minL)minL=kb.left;if(kb.top<minT)minT=kb.top;if(kb.right>maxR)maxR=kb.right;if(kb.bottom>maxB)maxB=kb.bottom;}"
            + "if(!n)return false;var rb=node.getBoundingClientRect();"
            + "return Math.abs(rb.left-minL)<2&&Math.abs(rb.top-minT)<2&&Math.abs(rb.right-maxR)<2&&Math.abs(rb.bottom-maxB)<2;}"
            + "var kept=0;for(var j=0;j<els.length;j++){if(!isWrap(els[j]))kept++;}return kept;}"
            + "var best=el;var parent=el.parentElement;var guard=0;"
            + "var startH=el.getBoundingClientRect().height||0;"
            + "while(parent&&parent.nodeType===1&&guard++<8){"
            + "var tag=(parent.tagName||'').toLowerCase();if(tag==='body'||tag==='html')break;"
            + "var pr=parent.getBoundingClientRect();if(pr.width<2||pr.height<2)break;"
            + "if(pr.height>Math.max(480,startH*6))break;"
            + "var n=count(parent);if(n>15||n<1)break;"
            + "best=parent;parent=parent.parentElement;}"
            + "return best;"
            + "}";

        private const string ListFn =
            "function(){"
            + "var root=this;"
            + "if(root&&root.nodeType===9)root=root.documentElement;"
            + "if(root&&root.nodeType===3)root=root.parentElement;"
            + "if(!root||!root.getBoundingClientRect)return{ok:false};"
            + "var rr=root.getBoundingClientRect();"
            + "function hidden(el){var st;try{st=getComputedStyle(el);}catch(e){return true;}"
            + "return !st||st.display==='none'||st.visibility==='hidden';}"
            + "function svgPart(el){var tag=(el.tagName||'').toLowerCase();"
            + "return tag==='path'||tag==='rect'||tag==='circle'||tag==='line'||tag==='polyline'||tag==='polygon'||tag==='ellipse';}"
            + "function skipTag(el){var tag=(el.tagName||'').toLowerCase();"
            + "return tag==='script'||tag==='style'||tag==='noscript'||tag==='br';}"
            + "function intersects(r){return r.width>1&&r.height>1&&r.left<rr.right-0.5&&r.right>rr.left+0.5&&r.top<rr.bottom-0.5&&r.bottom>rr.top+0.5;}"
            + "var all=root.querySelectorAll?root.querySelectorAll('*'):[];"
            + "var els=[];"
            + "for(var i=0;i<all.length;i++){var el=all[i];if(el===root||skipTag(el)||svgPart(el)||hidden(el))continue;"
            + "var r=el.getBoundingClientRect();if(!intersects(r))continue;els.push(el);}"
            + "function isWrap(el){var kids=el.children;if(!kids||!kids.length)return false;"
            + "var minL=1e9,minT=1e9,maxR=-1e9,maxB=-1e9,n=0;"
            + "for(var k=0;k<kids.length;k++){var c=kids[k];if(skipTag(c)||svgPart(c)||hidden(c))continue;"
            + "var kb=c.getBoundingClientRect();if(kb.width<=1||kb.height<=1)continue;n++;"
            + "if(kb.left<minL)minL=kb.left;if(kb.top<minT)minT=kb.top;if(kb.right>maxR)maxR=kb.right;if(kb.bottom>maxB)maxB=kb.bottom;}"
            + "if(!n)return false;var rb=el.getBoundingClientRect();"
            + "return Math.abs(rb.left-minL)<2&&Math.abs(rb.top-minT)<2&&Math.abs(rb.right-maxR)<2&&Math.abs(rb.bottom-maxB)<2;}"
            + "function cssPath(el){var parts=[];var n=el;var guard=0;"
            + "while(n&&n.nodeType===1&&guard++<40){var tag=String(n.tagName||'').toLowerCase();"
            + "if(!tag||tag==='html'){parts.unshift('html');break;}var parent=n.parentElement;if(!parent)break;"
            + "var idx=1;var sib=n;while((sib=sib.previousElementSibling)){if(String(sib.tagName||'').toLowerCase()===tag)idx++;}"
            + "parts.unshift(tag+':nth-of-type('+idx+')');n=parent;}return parts.join(' > ');}"
            + "function xpath(el){var parts=[];var n=el;var guard=0;"
            + "while(n&&n.nodeType===1&&guard++<80){var tag=String(n.tagName||'').toLowerCase();if(!tag)break;"
            + "if(tag==='html'){parts.unshift('html');break;}var parent=n.parentElement;if(!parent)break;"
            + "var idx=1;var seen=0;for(var c=parent.firstElementChild;c;c=c.nextElementSibling){"
            + "if(String(c.tagName||'').toLowerCase()===tag){seen++;if(c===n)idx=seen;}}"
            + "parts.unshift(tag+'['+idx+']');n=parent;}return parts.length?('/'+parts.join('/')):'';}"
            + "function textOf(el){var t=String(el.innerText||el.textContent||'').replace(/\\s+/g,' ').trim();"
            + "if(!t||t.length>24)return '';"
            + "for(var i=0;i<t.length;i++){var c=t.charCodeAt(i);if(c>=0xE000&&c<=0xF8FF)return '';}return t;}"
            + "function stateOf(el){var c=String(el.className||'').toLowerCase();"
            + "if(!/checkbox|radio|duox|option-icon|banxuan/.test(c))return '';"
            + "var blob=c;var p=el.parentElement;if(p)blob+=' '+String(p.className||'').toLowerCase();"
            + "var aria=el.getAttribute?el.getAttribute('aria-checked'):'';"
            + "if(aria==='mixed'||/indeterminate|banxuan|\\bhalf\\b/.test(blob))return '半选';"
            + "if(aria==='true'||/\\bselected\\b|\\bchecked\\b/.test(blob))return '已选';"
            + "if(aria==='false')return '未选';return '';}"
            + "var items=[];"
            + "for(var j=0;j<els.length;j++){var el=els[j];if(isWrap(el))continue;var b=el.getBoundingClientRect();"
            + "var label=textOf(el)||stateOf(el);"
            + "items.push({x:b.left-rr.left,y:b.top-rr.top,w:b.width,h:b.height,text:label,css:cssPath(el),xpath:xpath(el)});}"
            + "return{ok:true,items:items};"
            + "}";

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
