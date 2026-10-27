using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 补全 hover 才显现的「查看详情」块（含 .imgtextbtn / display:none）、课程卡、侧栏短菜单。
    /// 仍不放行无详情文案的大段 display:none 模板。失败不阻断 AX snapshot。
    /// </summary>
    internal static class BrowserDomClickableSupplement
    {
        private const int MaxExtra = 120;

        private const string CollectJs =
            @"(function(){
  function ownText(el){
    var t='';
    for (var i=0;i<el.childNodes.length;i++){
      var n=el.childNodes[i];
      if (n.nodeType===3) t+=n.textContent;
    }
    return String(t||'').replace(/\s+/g,' ').trim();
  }
  function isDetailName(s){
    if (!s) return false;
    s=String(s).replace(/\s+/g,' ').trim();
    if (s.indexOf('查看详情')>=0) return true;
    if (s.indexOf('查看更多')>=0) return true;
    return s==='详情';
  }
  function isHoverRole(el){
    var tag=(el.tagName||'').toLowerCase();
    if (tag==='a'||tag==='button') return true;
    var role=((el.getAttribute&&el.getAttribute('role'))||'').toLowerCase();
    return role==='button'||role==='link'||role==='menuitem'||role==='tab';
  }
  function classBlob(el){
    return (classNameOf(el)+' '+(el.id||'')).toLowerCase();
  }
  function isNamedIcon(el){
    return /delete|remove|trash|\bdel\b|shanchu|icon-del|\bedit\b|pencil|modify|bianji|icon-edit/.test(classBlob(el));
  }
  function isGlyphIcon(el){
    var tag=(el.tagName||'').toLowerCase();
    if (tag!=='i' && tag!=='span' && tag!=='em') return false;
    var inner=String(el.textContent||'').replace(/\s+/g,' ').trim();
    if (inner.length>4) return false;
    if (/[\uE000-\uF8FF]/.test(inner)) return true;
    return tag==='i' && inner.length<=1;
  }
  function isClickableIcon(el){
    var tag=(el.tagName||'').toLowerCase();
    if (tag!=='i' && tag!=='span' && tag!=='em') return false;
    var role=((el.getAttribute&&el.getAttribute('role'))||'').toLowerCase();
    if (role==='button'||role==='link'||role==='menuitem') return true;
    if (el.getAttribute && el.getAttribute('onclick')) return true;
    return isNamedIcon(el) || isGlyphIcon(el);
  }
  function insideHiddenTemplate(el){
    var p=el.parentElement;
    while (p && p!==document.body && p!==document.documentElement){
      var pst;
      try { pst=window.getComputedStyle(p); } catch (eH) { break; }
      if (pst && (pst.display==='none' || pst.visibility==='hidden')) return true;
      p=p.parentElement;
    }
    return false;
  }
  function boxMissing(el){
    try {
      var r=el.getBoundingClientRect();
      return !r || r.width<=0 || r.height<=0;
    } catch (eB) { return true; }
  }
  function classHint(el){
    var blob=classBlob(el);
    if (/delete|remove|trash|shanchu|icon-del|icon-close|\bclose\b|btn-del|del-btn/.test(blob)) return '删除';
    if (/\bedit\b|pencil|modify|bianji|icon-edit/.test(blob)) return '编辑';
    var parts=blob.match(/[a-z][a-z0-9-]{2,18}/g)||[];
    var skip={icon:1,iconfont:1,icons:1,button:1,btn:1,fl:1,fr:1};
    var picked='';
    for (var i=0;i<parts.length;i++){ if (!skip[parts[i]]) picked=parts[i]; }
    return picked;
  }
  function rowHint(el){
    var p=el&&el.parentElement;
    for (var d=0;p&&d<8;d++){
      var kids=p.children||[];
      for (var i=0;i<kids.length;i++){
        var k=kids[i];
        if (k===el) continue;
        if (k.contains&&k.contains(el)) continue;
        var t=String(k.innerText||k.textContent||'').replace(/[\uE000-\uF8FF]/g,'').replace(/\s+/g,' ').trim();
        if (t&&t.length>=2&&t.length<=20) return t;
      }
      p=p.parentElement;
    }
    return '';
  }
  function iconLabel(el){
    var aria=el.getAttribute&&(el.getAttribute('aria-label')||el.getAttribute('title'));
    aria=String(aria||'').replace(/\s+/g,' ').trim();
    if (aria) return aria.slice(0,40);
    var hinted=classHint(el);
    if (hinted==='删除'||hinted==='编辑') return hinted;
    var ownT=ownText(el);
    if (ownT && ownT.length<=16 && !/[\uE000-\uF8FF]/.test(ownT)) return ownT;
    var name='图标';
    if (hinted) name=name+' '+hinted;
    var rh=rowHint(el);
    if (rh && name.indexOf(rh)<0) name=(name+' '+rh).slice(0,40);
    return name;
  }
  function classNameOf(el){
    var c=el.className;
    if (c && typeof c==='object' && c.baseVal!=null) return String(c.baseVal);
    return String(c||'');
  }
  function isImgtextBtn(el){
    return /\bimgtextbtn\b/.test(classNameOf(el));
  }
  function isImgtextCard(el){
    return /\bimgtext\b/.test(classNameOf(el)) && el.querySelector && !!el.querySelector('img');
  }
  function isPointerCard(el, st){
    if (!st || st.cursor!=='pointer') return false;
    if (!el.querySelector) return false;
    if (!el.querySelector('img')) return false;
    return el.offsetWidth>=100 && el.offsetHeight>=60;
  }
  function nameOf(el){
    return (el.getAttribute&&(el.getAttribute('aria-label')||el.getAttribute('title'))) || ownText(el) || '';
  }
  function isCardGroup(el){
    if (!el || !el.querySelector) return false;
    var tag=(el.tagName||'').toLowerCase();
    if (tag==='a'||tag==='button') return false;
    if (!el.querySelector('a[href], [role=link]')) return false;
    if (!el.querySelector('img')) return false;
    var n=String(nameOf(el)||'').replace(/\s+/g,' ').trim();
    if (!n || n.length<8) return false;
    try {
      var w=el.offsetWidth, h=el.offsetHeight;
      if (w<100||h<60||w>720||h>900) return false;
    } catch (eC) { return false; }
    var links=el.querySelectorAll('a[href], [role=link]');
    return links.length>=1 && links.length<=8;
  }
  function closestCardGroup(el){
    var p=el;
    while (p && p!==document.body){
      if (isCardGroup(p)) return p;
      p=p.parentElement;
    }
    return null;
  }
  function isCardCoverImg(el){
    if (!el || (el.tagName||'').toLowerCase()!=='img') return false;
    try {
      var r=el.getBoundingClientRect();
      if (r.width<40||r.height<40) return false;
    } catch (eI) { return false; }
    return !!closestCardGroup(el.parentElement);
  }
  var items=[];
  var all=document.querySelectorAll('body *');
  for (var i=0;i<all.length && items.length<120;i++){
    var el=all[i];
    var tag=(el.tagName||'').toLowerCase();
    if (tag==='script'||tag==='style'||tag==='svg'||tag==='path'||tag==='noscript') continue;
    var st;
    try { st=window.getComputedStyle(el); } catch (e) { continue; }
    var own=ownText(el);
    var displayNone=!!(st && st.display==='none');
    var innerNow=String(el.textContent||'').replace(/\s+/g,' ').trim();
    var rowHover=(isHoverRole(el)||isClickableIcon(el)) && innerNow.length<=16 && !insideHiddenTemplate(el);
    if (displayNone && !isDetailName(own) && !isImgtextBtn(el) && !isImgtextCard(el) && !rowHover) continue;
    var hidden=displayNone || (st && (st.visibility==='hidden' || parseFloat(st.opacity)===0));
    var inner=String(el.textContent||'').replace(/\s+/g,' ').trim();
    var shortInner=inner.slice(0,24);
    var kind=null;
    var name=own;
    var role='button';
    if (isImgtextBtn(el) || isDetailName(own) || (isDetailName(shortInner) && shortInner.length<=12 && !el.querySelector('img'))){
      kind=hidden?'hidden_detail':'detail';
      name=own||shortInner||'查看详情';
      role='button';
    } else if ((hidden || boxMissing(el)) && (isHoverRole(el) || isClickableIcon(el)) && !insideHiddenTemplate(el)){
      kind='hidden_detail';
      name=iconLabel(el);
      role=(tag==='a')?'link':'button';
      hidden=true;
    } else if ((isImgtextCard(el) || (!hidden && isPointerCard(el, st)))
        && !(el.querySelector && el.querySelector('a[href], [role=link]'))){
      kind='card';
      var alt=(el.querySelector('img')&&el.querySelector('img').alt)||'';
      name=(own||'').slice(0,40) || alt || inner.slice(0,40) || '卡片';
      role='generic';
    } else if (!hidden && isCardCoverImg(el)){
      kind='pic';
      var g2=closestCardGroup(el.parentElement);
      name=(el.alt||'').slice(0,40) || (g2&&nameOf(g2)||'').slice(0,40) || '商品图';
      role='image';
    } else if (!hidden && st && st.cursor==='pointer' && own.length>=2 && own.length<=16){
      var rect=el.getBoundingClientRect();
      if (rect.width>0 && rect.height>0){
        kind='pointer_label';
        name=own;
        role='button';
      }
    } else if (!hidden && own.length>=2 && own.length<=16){
      var hint=!!(el.getAttribute&&el.getAttribute('onclick'));
      var ti=el.getAttribute&&el.getAttribute('tabindex');
      if (!hint && ti!=null && ti!=='' && !isNaN(Number(ti)) && Number(ti)>=0) hint=true;
      if (hint){
        kind='click_hint';
        name=own;
        role='button';
      }
    }
    if (!kind && !hidden){
      var rowRole=((el.getAttribute&&el.getAttribute('role'))||'').toLowerCase();
      if (tag==='tr' || rowRole==='row' || /\bel-table__row\b/.test(classNameOf(el)) || /\bant-table-row\b/.test(classNameOf(el))){
        var cells=el.querySelectorAll?el.querySelectorAll(""td, th, [role='cell'], [role='gridcell']""):[];
        var parts=[];
        for (var ci=0; ci<cells.length && parts.length<6; ci++){
          var ct=ownText(cells[ci]) || String(cells[ci].innerText||'').replace(/\s+/g,' ').trim();
          if (ct && ct.length<=40) parts.push(ct);
        }
        var rn=parts.join(' ').slice(0,80);
        if (rn){
          kind='row';
          name=rn;
          role='row';
        }
      }
    }
    if (!kind) continue;
    var mark='yw'+items.length;
    try { el.setAttribute('data-yw-clk', mark); } catch (e2) { continue; }
    items.push({mark:mark, kind:kind, name:name, role:role, hidden:!!hidden});
  }
  return items;
})()";

        private const string CleanupJs =
            @"(function(){
  var list=document.querySelectorAll('[data-yw-clk]');
  for (var i=0;i<list.length;i++){ list[i].removeAttribute('data-yw-clk'); }
  return true;
})()";

        private const string RowHiddenJs =
            @"function() {
  try {
    var el = this;
    if (el && el.nodeType !== 1) el = el.parentElement;
    if (!el) return true;
    if (el.isConnected === false) return true;
    var n = el;
    while (n && n.nodeType === 1) {
      var st = null;
      try { st = window.getComputedStyle(n); } catch (e1) {}
      if (st && (st.display === 'none' || st.visibility === 'hidden')) return true;
      var cls = (n.className && typeof n.className === 'string') ? n.className : '';
      if (/\bpanel\b/.test(cls) && !/\bon\b/.test(cls)) return true;
      n = n.parentElement;
    }
    return false;
  } catch (e) {
    return false;
  }
}";

        /// <summary>
        /// AX 会把 display:none 面板里的表行也发 ref。D5：藏着的表不当可见行。
        /// </summary>
        public static async Task PruneHiddenRowsAsync(
            IBrowserCdp cdp,
            BrowserAxBuildResult built,
            string sessionId)
        {
            if (cdp == null || built == null || !built.Success || built.Refs == null || built.Refs.Count == 0)
            {
                return;
            }

            var drop = new List<string>();
            foreach (var kv in built.Refs)
            {
                BrowserRefEntry e = kv.Value;
                if (e == null || !e.BackendDomNodeId.HasValue)
                {
                    continue;
                }

                string role = e.Role ?? "";
                if (!string.Equals(role, "row", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(role, "LayoutTableRow", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string sid = string.IsNullOrWhiteSpace(e.CdpSessionId) ? sessionId : e.CdpSessionId;
                try
                {
                    string objectId = await BrowserInteractEngine
                        .ResolveObjectIdAsync(cdp, e.BackendDomNodeId.Value, sid)
                        .ConfigureAwait(true);
                    string json = await BrowserInteractEngine
                        .CallFunctionOnAsync(cdp, objectId, RowHiddenJs, null, sid)
                        .ConfigureAwait(true);
                    if (ReadReturnedBool(json))
                    {
                        drop.Add(kv.Key);
                    }
                }
                catch
                {
                }
            }

            if (drop.Count == 0)
            {
                return;
            }

            foreach (string id in drop)
            {
                built.Refs.Remove(id);
            }

            built.TreeText = DropRefLines(built.TreeText, drop);
        }

        public static async Task MergeAsync(
            IBrowserCdp cdp,
            BrowserAxBuildResult built,
            string frameId,
            string sessionId)
        {
            if (cdp == null || built == null || !built.Success || built.Refs == null)
            {
                return;
            }

            int? contextId = null;
            try
            {
                contextId = await TryCreateFrameWorldAsync(cdp, frameId, sessionId).ConfigureAwait(true);
                JArray items = await EvaluateArrayAsync(cdp, CollectJs, contextId, sessionId)
                    .ConfigureAwait(true);
                if (items == null || items.Count == 0)
                {
                    return;
                }

                var already = new HashSet<int>();
                int maxRef = 0;
                foreach (var kv in built.Refs)
                {
                    if (kv.Value?.BackendDomNodeId != null)
                    {
                        already.Add(kv.Value.BackendDomNodeId.Value);
                    }

                    if (kv.Key != null
                        && kv.Key.Length > 1
                        && kv.Key[0] == 'e'
                        && int.TryParse(kv.Key.Substring(1), out int n)
                        && n > maxRef)
                    {
                        maxRef = n;
                    }
                }

                var sb = new StringBuilder(built.TreeText ?? "");
                bool header = false;
                int nextRef = maxRef + 1;
                int added = 0;

                foreach (JToken tok in items)
                {
                    if (added >= MaxExtra
                        || built.Refs.Count >= BrowserAxTreeBuilder.MaxRefNodes
                        || sb.Length >= BrowserAxTreeBuilder.MaxTreeChars)
                    {
                        built.Truncated = true;
                        built.TruncatedReason = built.Refs.Count >= BrowserAxTreeBuilder.MaxRefNodes
                            ? "ref_nodes>" + BrowserAxTreeBuilder.MaxRefNodes
                            : "tree_chars>" + BrowserAxTreeBuilder.MaxTreeChars;
                        break;
                    }

                    JObject item = tok as JObject;
                    if (item == null)
                    {
                        continue;
                    }

                    string mark = (string)item["mark"];
                    if (string.IsNullOrWhiteSpace(mark))
                    {
                        continue;
                    }

                    int? backend = await ResolveMarkedBackendAsync(cdp, mark, contextId, sessionId)
                        .ConfigureAwait(true);
                    if (!backend.HasValue || already.Contains(backend.Value))
                    {
                        continue;
                    }

                    string role = (string)item["role"] ?? "generic";
                    string name = (string)item["name"] ?? "";
                    bool hidden = item["hidden"] != null
                        && item["hidden"].Type == JTokenType.Boolean
                        && (bool)item["hidden"];

                    if (!header)
                    {
                        if (sb.Length > 0 && !sb.ToString().EndsWith("\n"))
                        {
                            sb.AppendLine();
                        }

                        sb.AppendLine("- (补全) 可点控件");
                        header = true;
                    }

                    string refId = "e" + nextRef.ToString(CultureInfo.InvariantCulture);
                    nextRef++;
                    built.Refs[refId] = new BrowserRefEntry
                    {
                        BackendDomNodeId = backend.Value,
                        Role = role,
                        Name = name,
                        FrameId = frameId,
                        CdpSessionId = sessionId
                    };
                    already.Add(backend.Value);
                    added++;

                    sb.Append("  - ").Append(role).Append(" ref=").Append(refId);
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        sb.Append(" \"").Append(Escape(name.Trim())).Append('"');
                    }

                    if (hidden)
                    {
                        sb.Append(" (隐藏)");
                    }

                    sb.AppendLine();
                }

                if (header)
                {
                    built.TreeText = sb.ToString().TrimEnd();
                }
            }
            catch
            {
            }
            finally
            {
                try
                {
                    await EvaluateAsync(cdp, CleanupJs, contextId, sessionId).ConfigureAwait(true);
                }
                catch
                {
                }
            }
        }

        private static async Task<int?> TryCreateFrameWorldAsync(
            IBrowserCdp cdp,
            string frameId,
            string sessionId)
        {
            if (string.IsNullOrWhiteSpace(frameId))
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
                JObject jo = JObject.Parse(json);
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

        private static async Task<JArray> EvaluateArrayAsync(
            IBrowserCdp cdp,
            string expression,
            int? contextId,
            string sessionId)
        {
            string json = await EvaluateAsync(cdp, expression, contextId, sessionId).ConfigureAwait(true);
            JObject jo = JObject.Parse(json);
            JToken value = jo["result"]?["value"];
            return value as JArray;
        }

        private static async Task<string> EvaluateAsync(
            IBrowserCdp cdp,
            string expression,
            int? contextId,
            string sessionId)
        {
            var payload = new JObject
            {
                ["expression"] = expression,
                ["returnByValue"] = true,
                ["awaitPromise"] = true
            };
            if (contextId.HasValue)
            {
                payload["contextId"] = contextId.Value;
            }

            return await cdp
                .CallCdpAsync("Runtime.evaluate", payload.ToString(Formatting.None), sessionId)
                .ConfigureAwait(true);
        }

        private static async Task<int?> ResolveMarkedBackendAsync(
            IBrowserCdp cdp,
            string mark,
            int? contextId,
            string sessionId)
        {
            string expr = "document.querySelector('[data-yw-clk=" + JsonEscape(mark) + "]')";
            var payload = new JObject
            {
                ["expression"] = expr,
                ["returnByValue"] = false
            };
            if (contextId.HasValue)
            {
                payload["contextId"] = contextId.Value;
            }

            string json = await cdp
                .CallCdpAsync("Runtime.evaluate", payload.ToString(Formatting.None), sessionId)
                .ConfigureAwait(true);
            JObject jo = JObject.Parse(json);
            string objectId = (string)jo["result"]?["objectId"];
            if (string.IsNullOrEmpty(objectId))
            {
                return null;
            }

            var descPayload = new JObject { ["objectId"] = objectId };
            string descJson = await cdp
                .CallCdpAsync("DOM.describeNode", descPayload.ToString(Formatting.None), sessionId)
                .ConfigureAwait(true);
            JObject desc = JObject.Parse(descJson);
            JToken backend = desc["node"]?["backendNodeId"];
            if (backend == null || backend.Type == JTokenType.Null)
            {
                return null;
            }

            return backend.Value<int>();
        }

        private static bool ReadReturnedBool(string resultJson)
        {
            try
            {
                JObject jo = JObject.Parse(resultJson ?? "{}");
                JToken tok = jo["result"]?["value"];
                if (tok == null || tok.Type == JTokenType.Null)
                {
                    return false;
                }

                if (tok.Type == JTokenType.Boolean)
                {
                    return (bool)tok;
                }

                return string.Equals(tok.ToString(), "true", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static string DropRefLines(string tree, List<string> refIds)
        {
            if (string.IsNullOrEmpty(tree) || refIds == null || refIds.Count == 0)
            {
                return tree;
            }

            var keep = new List<string>();
            foreach (string line in tree.Replace("\r\n", "\n").Split('\n'))
            {
                bool drop = false;
                for (int i = 0; i < refIds.Count; i++)
                {
                    if (LineHasRef(line, refIds[i]))
                    {
                        drop = true;
                        break;
                    }
                }

                if (!drop)
                {
                    keep.Add(line);
                }
            }

            return string.Join("\n", keep.ToArray()).TrimEnd();
        }

        private static bool LineHasRef(string line, string refId)
        {
            if (string.IsNullOrEmpty(line) || string.IsNullOrEmpty(refId))
            {
                return false;
            }

            string needle = "ref=" + refId;
            int at = line.IndexOf(needle, StringComparison.Ordinal);
            if (at < 0)
            {
                return false;
            }

            int end = at + needle.Length;
            return end >= line.Length || !char.IsDigit(line[end]);
        }

        private static string JsonEscape(string s)
        {
            return "\"" + (s ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }

        private static string Escape(string s)
        {
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", " ").Replace("\n", " ");
        }
    }
}
