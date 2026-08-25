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
    if (displayNone && !isDetailName(own) && !isImgtextBtn(el) && !isImgtextCard(el)) continue;
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
    } else if (hidden && isHoverRole(el)){
      kind='hidden_detail';
      name=own||shortInner||'button';
      role='button';
    } else if (isImgtextCard(el) || (!hidden && isPointerCard(el, st))){
      kind='card';
      var alt=(el.querySelector('img')&&el.querySelector('img').alt)||'';
      name=(own||'').slice(0,40) || alt || inner.slice(0,40) || '卡片';
      role='generic';
    } else if (!hidden && st && st.cursor==='pointer' && own.length>=2 && own.length<=16){
      var rect=el.getBoundingClientRect();
      if (rect.left>=0 && rect.left<240 && rect.width>0 && rect.height>0){
        kind='menu';
        name=own;
        role='menuitem';
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

        public static async Task MergeAsync(
            YiWriteBrowserForm form,
            BrowserAxBuildResult built,
            string frameId,
            string sessionId)
        {
            if (form == null || built == null || !built.Success || built.Refs == null)
            {
                return;
            }

            int? contextId = null;
            try
            {
                contextId = await TryCreateFrameWorldAsync(form, frameId, sessionId).ConfigureAwait(true);
                JArray items = await EvaluateArrayAsync(form, CollectJs, contextId, sessionId)
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

                    int? backend = await ResolveMarkedBackendAsync(form, mark, contextId, sessionId)
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
                    await EvaluateAsync(form, CleanupJs, contextId, sessionId).ConfigureAwait(true);
                }
                catch
                {
                }
            }
        }

        private static async Task<int?> TryCreateFrameWorldAsync(
            YiWriteBrowserForm form,
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
                string json = await form
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
            YiWriteBrowserForm form,
            string expression,
            int? contextId,
            string sessionId)
        {
            string json = await EvaluateAsync(form, expression, contextId, sessionId).ConfigureAwait(true);
            JObject jo = JObject.Parse(json);
            JToken value = jo["result"]?["value"];
            return value as JArray;
        }

        private static async Task<string> EvaluateAsync(
            YiWriteBrowserForm form,
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

            return await form
                .CallCdpAsync("Runtime.evaluate", payload.ToString(Formatting.None), sessionId)
                .ConfigureAwait(true);
        }

        private static async Task<int?> ResolveMarkedBackendAsync(
            YiWriteBrowserForm form,
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

            string json = await form
                .CallCdpAsync("Runtime.evaluate", payload.ToString(Formatting.None), sessionId)
                .ConfigureAwait(true);
            JObject jo = JObject.Parse(json);
            string objectId = (string)jo["result"]?["objectId"];
            if (string.IsNullOrEmpty(objectId))
            {
                return null;
            }

            var descPayload = new JObject { ["objectId"] = objectId };
            string descJson = await form
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
