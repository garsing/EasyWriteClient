using System;
using System.Globalization;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
        /// <summary>
        /// 认整页 HTML 弹出：浮层 + 全页遮罩。命中则给 z 最高的一层根打标，供快照只出这一层。
        /// Cookie 窄条、无遮罩侧栏不命中。嵌套小窗取最上面那层，不要求占主界面面积。
        /// </summary>
    internal static class BrowserHtmlModalProbe
    {
        public const string MarkerAttr = "data-yiwrite-modal-root";
        public const string MarkerSelector = "[data-yiwrite-modal-root=\"1\"]";

        public const string MarkExpression =
            @"(function(){
  function wipe(){
    var old=document.querySelectorAll('[data-yiwrite-modal-root]');
    for (var i=0;i<old.length;i++) old[i].removeAttribute('data-yiwrite-modal-root');
  }
  function cls(el){
    if (!el) return '';
    var c=el.className;
    var name=c && typeof c==='object' && c.baseVal!=null ? String(c.baseVal) : String(c||'');
    return (name+' '+(el.id||'')).toLowerCase();
  }
  function vw(){ return window.innerWidth || document.documentElement.clientWidth || 0; }
  function vh(){ return window.innerHeight || document.documentElement.clientHeight || 0; }
  function zOf(el){
    try {
      var z=parseInt(window.getComputedStyle(el).zIndex,10);
      return isNaN(z)?0:z;
    } catch (e) { return 0; }
  }
  function isShade(el, W, H){
    if (!el) return false;
    if (/\blayui-layer-shade\b|\blayer-shade\b/.test(cls(el))) return true;
    var st;
    try { st=window.getComputedStyle(el); } catch (eS) { return false; }
    if (!st || st.position!=='fixed') return false;
    var r=el.getBoundingClientRect();
    if (r.width<W*0.9 || r.height<H*0.9) return false;
    if (r.left>8 || r.top>8) return false;
    var zi=parseInt(st.zIndex,10);
    if (isNaN(zi) || zi<1) return false;
    var op=parseFloat(st.opacity);
    if (!isNaN(op) && op===0) return false;
    return true;
  }
  function isTiny(el){
    var r=el.getBoundingClientRect();
    return r.width<8 || r.height<8;
  }
  function isFloating(el){
    var st;
    try { st=window.getComputedStyle(el); } catch (eF) { return false; }
    if (!st) return false;
    if (st.position==='fixed') return true;
    var c=cls(el);
    if (/\blayui-layer\b/.test(c) && !/\blayui-layer-shade\b/.test(c)) return true;
    var role=((el.getAttribute&&el.getAttribute('role'))||'').toLowerCase();
    if (role==='dialog') return true;
    if (((el.getAttribute&&el.getAttribute('aria-modal'))||'').toLowerCase()==='true') return true;
    return false;
  }
  wipe();
  var W=vw(), H=vh();
  if (W<50 || H<50 || !document.body) return {found:false};
  var all=document.body.querySelectorAll('*');
  var shadeZ=-1;
  var i;
  for (i=0;i<all.length && i<5000;i++){
    if (isShade(all[i], W, H)) {
      var zs=zOf(all[i]);
      if (zs>shadeZ) shadeZ=zs;
    }
  }
  if (shadeZ<0) return {found:false};
  var cands=[];
  for (i=0;i<all.length && i<5000;i++){
    var el=all[i];
    if (isShade(el, W, H)) continue;
    if (!isFloating(el)) continue;
    if (isTiny(el)) continue;
    var z=zOf(el);
    if (z<shadeZ) continue;
    cands.push(el);
  }
  var best=null, bestZ=-1, a, b, skip;
  for (a=0;a<cands.length;a++){
    skip=false;
    for (b=0;b<cands.length;b++){
      if (a===b) continue;
      if (cands[b].contains(cands[a])) { skip=true; break; }
    }
    if (skip) continue;
    z=zOf(cands[a]);
    if (z>=bestZ) { best=cands[a]; bestZ=z; }
  }
  if (!best) return {found:false};
  best.setAttribute('data-yiwrite-modal-root','1');
  return {found:true};
})()";

        public static async Task<int?> TryGetBackendNodeIdAsync(IBrowserCdp cdp, string sessionId)
        {
            if (cdp == null)
            {
                return null;
            }

            try
            {
                var evalPayload = new JObject
                {
                    ["expression"] = MarkExpression,
                    ["returnByValue"] = true
                };
                string evalJson = await cdp.CallCdpAsync(
                        "Runtime.evaluate",
                        evalPayload.ToString(Newtonsoft.Json.Formatting.None),
                        sessionId)
                    .ConfigureAwait(true);
                var eval = JObject.Parse(evalJson ?? "{}");
                if (eval["exceptionDetails"] != null && eval["exceptionDetails"].Type != JTokenType.Null)
                {
                    return null;
                }

                bool found = eval["result"]?["value"]?["found"]?.Value<bool>() ?? false;
                if (!found)
                {
                    return null;
                }

                string docJson = await cdp.CallCdpAsync(
                        "DOM.getDocument",
                        "{\"depth\":0}",
                        sessionId)
                    .ConfigureAwait(true);
                var doc = JObject.Parse(docJson ?? "{}");
                if (doc["root"]?["nodeId"] == null)
                {
                    return null;
                }

                int rootId = doc["root"]["nodeId"].Value<int>();
                var qsPayload = new JObject
                {
                    ["nodeId"] = rootId,
                    ["selector"] = MarkerSelector
                };
                string qsJson = await cdp.CallCdpAsync(
                        "DOM.querySelector",
                        qsPayload.ToString(Newtonsoft.Json.Formatting.None),
                        sessionId)
                    .ConfigureAwait(true);
                var qs = JObject.Parse(qsJson ?? "{}");
                if (qs["nodeId"] == null || qs["nodeId"].Type == JTokenType.Null)
                {
                    return null;
                }

                int nodeId = qs["nodeId"].Value<int>();
                if (nodeId == 0)
                {
                    return null;
                }

                string descJson = await cdp.CallCdpAsync(
                        "DOM.describeNode",
                        "{\"nodeId\":" + nodeId.ToString(CultureInfo.InvariantCulture) + "}",
                        sessionId)
                    .ConfigureAwait(true);
                var desc = JObject.Parse(descJson ?? "{}");
                if (desc["node"]?["backendNodeId"] == null
                    || desc["node"]["backendNodeId"].Type == JTokenType.Null)
                {
                    return null;
                }

                return desc["node"]["backendNodeId"].Value<int>();
            }
            catch
            {
                return null;
            }
        }
    }
}
