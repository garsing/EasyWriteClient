using System;
using System.Globalization;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>WebView2 CDP：click / hover / type / scroll / press / select。</summary>
    internal static class BrowserInteractEngine
    {
        internal const int HoverSettleMs = 150;
        public static async Task<DomNodeProbe> ProbeAsync(
            YiWriteBrowserForm form,
            int backendNodeId,
            string sessionId = null)
        {
            string objectId = await ResolveObjectIdAsync(form, backendNodeId, sessionId).ConfigureAwait(true);
            string resultJson = await CallFunctionOnAsync(
                form,
                objectId,
                @"function() {
  var tag = (this.tagName || '').toUpperCase();
  var type = (this.getAttribute && this.getAttribute('type')) || this.type || '';
  var ac = (this.getAttribute && this.getAttribute('autocomplete')) || '';
  var role = (this.getAttribute && this.getAttribute('role')) || '';
  var name = (this.getAttribute && (this.getAttribute('aria-label') || this.getAttribute('name') || this.getAttribute('placeholder'))) || '';
  var href = '';
  try { href = this.href || (this.getAttribute && this.getAttribute('href')) || ''; } catch (eH) {}
  try {
    if ((!href || href === '#' || (href + '').toLowerCase().indexOf('javascript:') === 0) && this.closest) {
      var a = this.closest('a[href]');
      if (a) href = a.href || a.getAttribute('href') || href;
    }
  } catch (eC) {}
  var dataUrl = '';
  try {
    if (this.getAttribute) {
      dataUrl = this.getAttribute('data-url')
        || this.getAttribute('data-href')
        || this.getAttribute('data-download')
        || this.getAttribute('data-file')
        || '';
    }
  } catch (eDu) {}
  var hasDownloadAttr = false;
  try { hasDownloadAttr = !!(this.getAttribute && this.getAttribute('download') != null); } catch (eD) {}
  var val = (this.getAttribute && this.getAttribute('value')) || '';
  var text = (this.innerText || this.textContent || '').trim().slice(0, 120);
  var pageHasPwd = !!(document.querySelector && document.querySelector('input[type=password]'));
  var disabled = false;
  var readOnly = false;
  var ariaDisabled = false;
  try {
    if (this.disabled) disabled = true;
    if (this.matches && this.matches(':disabled')) disabled = true;
  } catch (eDis) {}
  try { if (this.readOnly) readOnly = true; } catch (eRo) {}
  try {
    var ad = this.getAttribute && this.getAttribute('aria-disabled');
    if (ad != null && String(ad).toLowerCase() !== 'false') ariaDisabled = true;
  } catch (eAd) {}
  var isFileChooser = false;
  try {
    var tLow = (tag || '').toLowerCase();
    var typLow = String(type || '').toLowerCase();
    if (tLow === 'input' && typLow === 'file') isFileChooser = true;
    else if (this.querySelector && this.querySelector('input[type=file]')) isFileChooser = true;
    else if (this.closest) {
      var lab = this.closest('label');
      if (lab) {
        if (lab.querySelector && lab.querySelector('input[type=file]')) isFileChooser = true;
        else {
          var fid = lab.htmlFor || (lab.getAttribute && lab.getAttribute('for')) || '';
          if (fid) {
            var fe = document.getElementById(fid);
            if (fe && String(fe.type || '').toLowerCase() === 'file') isFileChooser = true;
          }
        }
      }
    }
  } catch (eF) {}
  return {
    tag: tag,
    type: type,
    autocomplete: ac,
    role: role,
    accessibleName: name,
    valueAttr: val,
    innerText: text,
    href: href,
    dataUrl: dataUrl,
    hasDownloadAttr: hasDownloadAttr,
    pageHasPasswordInput: pageHasPwd,
    isFileChooser: isFileChooser,
    disabled: disabled,
    readOnly: readOnly,
    ariaDisabled: ariaDisabled
  };
}",
                null,
                sessionId).ConfigureAwait(true);

            var probe = new DomNodeProbe();
            try
            {
                var jo = JObject.Parse(resultJson);
                var v = jo["result"]?["value"] as JObject;
                if (v != null)
                {
                    probe.Tag = (string)v["tag"];
                    probe.InputType = (string)v["type"];
                    probe.Autocomplete = (string)v["autocomplete"];
                    probe.Role = (string)v["role"];
                    probe.AccessibleName = (string)v["accessibleName"];
                    probe.ValueAttr = (string)v["valueAttr"];
                    probe.InnerText = (string)v["innerText"];
                    probe.Href = (string)v["href"];
                    probe.DataUrl = (string)v["dataUrl"];
                    probe.HasDownloadAttr = v["hasDownloadAttr"] != null
                        && v["hasDownloadAttr"].Type == JTokenType.Boolean
                        && (bool)v["hasDownloadAttr"];
                    probe.PageHasPasswordInput = v["pageHasPasswordInput"] != null
                        && v["pageHasPasswordInput"].Type == JTokenType.Boolean
                        && (bool)v["pageHasPasswordInput"];
                    probe.IsFileChooser = v["isFileChooser"] != null
                        && v["isFileChooser"].Type == JTokenType.Boolean
                        && (bool)v["isFileChooser"];
                    probe.Disabled = v["disabled"] != null
                        && v["disabled"].Type == JTokenType.Boolean
                        && (bool)v["disabled"];
                    probe.ReadOnly = v["readOnly"] != null
                        && v["readOnly"].Type == JTokenType.Boolean
                        && (bool)v["readOnly"];
                    probe.AriaDisabled = v["ariaDisabled"] != null
                        && v["ariaDisabled"].Type == JTokenType.Boolean
                        && (bool)v["ariaDisabled"];
                }
            }
            catch
            {
            }

            return probe;
        }

        public static async Task<DomNodeProbe> ProbePagePasswordAsync(
            YiWriteBrowserForm form,
            string sessionId = null)
        {
            string json = await form.CallCdpAsync(
                "Runtime.evaluate",
                "{\"expression\":\"!!document.querySelector('input[type=password]')\",\"returnByValue\":true}",
                sessionId)
                .ConfigureAwait(true);
            var probe = new DomNodeProbe();
            try
            {
                var jo = JObject.Parse(json);
                probe.PageHasPasswordInput = jo["result"]?["value"] != null
                    && jo["result"]["value"].Type == JTokenType.Boolean
                    && (bool)jo["result"]["value"];
            }
            catch
            {
            }

            return probe;
        }

        public static async Task ClickAsync(
            YiWriteBrowserForm form,
            int backendNodeId,
            string sessionId = null,
            string clickMode = null,
            string role = null)
        {
            if (IsCheckLikeRole(role))
            {
                await ToggleCheckAsync(form, backendNodeId, sessionId).ConfigureAwait(true);
                return;
            }

            string mode = ConfigManager.NormalizeBrowserClickMode(clickMode);
            string objectId = await ResolveObjectIdAsync(form, backendNodeId, sessionId).ConfigureAwait(true);
            string resultJson = await CallFunctionOnAsync(
                form,
                objectId,
                PointerPrepareJs,
                BuildCallObject(new JObject
                {
                    ["kind"] = "click",
                    ["mode"] = mode
                }),
                sessionId).ConfigureAwait(true);
            ThrowIfJsNotOk(resultJson, "点击被挡住");
            if (string.Equals(mode, "legacy", StringComparison.Ordinal)
                || ReadJsFallback(resultJson))
            {
                return;
            }

            ReadJsPoint(resultJson, out double x, out double y);
            await DispatchCdpClickAsync(form, objectId, x, y, sessionId).ConfigureAwait(true);
        }

        public static async Task HoverAsync(
            YiWriteBrowserForm form,
            int backendNodeId,
            string sessionId = null)
        {
            string objectId = await ResolveObjectIdAsync(form, backendNodeId, sessionId).ConfigureAwait(true);
            string resultJson = await CallFunctionOnAsync(
                form,
                objectId,
                PointerPrepareJs,
                BuildCallObject(new JObject
                {
                    ["kind"] = "hover",
                    ["mode"] = "pointer"
                }),
                sessionId).ConfigureAwait(true);
            ThrowIfJsNotOk(resultJson, "没有可悬停区域。请悬停可见的父节点或卡片。");
            ReadJsPoint(resultJson, out double x, out double y);
            await DispatchCdpHoverAsync(form, objectId, x, y, sessionId).ConfigureAwait(true);
            await Task.Delay(HoverSettleMs).ConfigureAwait(true);
        }

        public static async Task TypeAsync(
            YiWriteBrowserForm form,
            int backendNodeId,
            string text,
            string sessionId = null)
        {
            string want = text ?? "";
            string objectId = await ResolveObjectIdAsync(form, backendNodeId, sessionId).ConfigureAwait(true);

            // 1) 真实聚焦 + 点一下 + 原生 setter 清空（百度等受控框需要）
            string prepJson = await CallFunctionOnAsync(
                form,
                objectId,
                @"function() {
  this.scrollIntoView({block:'center', inline:'center'});
  try {
    this.dispatchEvent(new MouseEvent('mousedown', {bubbles:true, cancelable:true, view:window}));
    this.dispatchEvent(new MouseEvent('mouseup', {bubbles:true, cancelable:true, view:window}));
    this.dispatchEvent(new MouseEvent('click', {bubbles:true, cancelable:true, view:window}));
  } catch (e0) {}
  this.focus && this.focus();
  if (typeof this.select === 'function') { try { this.select(); } catch(e1) {} }
  if ('value' in this) {
    try {
      var proto = this.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      var desc = Object.getOwnPropertyDescriptor(proto, 'value');
      if (desc && desc.set) desc.set.call(this, '');
      else this.value = '';
    } catch (e2) { try { this.value = ''; } catch (e3) {} }
    try {
      this.dispatchEvent(new InputEvent('input', {bubbles:true, cancelable:true, inputType:'deleteContentBackward', data:null}));
    } catch (e4) {
      this.dispatchEvent(new Event('input', {bubbles:true}));
    }
    return { ok:true, kind:'value' };
  }
  if (this.isContentEditable) {
    this.innerText = '';
    this.textContent = '';
    this.dispatchEvent(new Event('input', {bubbles:true}));
    return { ok:true, kind:'ce' };
  }
  return { ok:false, reason:'not_editable' };
}",
                null,
                sessionId).ConfigureAwait(true);

            EnsureTypePrepOk(prepJson);

            // 2) CDP insertText：走浏览器输入通道，框架更容易收到（比单纯改 value 稳）
            var insertPayload = new JObject { ["text"] = want };
            await form.CallCdpAsync(
                "Input.insertText",
                insertPayload.ToString(Newtonsoft.Json.Formatting.None),
                sessionId)
                .ConfigureAwait(true);

            // 3) 再补一枪：原生 setter + InputEvent（insertText 偶发未同步时兜底）
            string argsJson = BuildCallArgs(want);
            string syncJson = await CallFunctionOnAsync(
                form,
                objectId,
                @"function(text) {
  if ('value' in this) {
    var cur = String(this.value || '');
    if (cur !== text) {
      try {
        var proto = this.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
        var desc = Object.getOwnPropertyDescriptor(proto, 'value');
        if (desc && desc.set) desc.set.call(this, text);
        else this.value = text;
      } catch (e) { this.value = text; }
    }
    try {
      this.dispatchEvent(new InputEvent('input', {bubbles:true, cancelable:true, inputType:'insertText', data:text}));
    } catch (e2) {
      this.dispatchEvent(new Event('input', {bubbles:true}));
    }
    this.dispatchEvent(new Event('change', {bubbles:true}));
    return { ok:true, value: String(this.value || '') };
  }
  if (this.isContentEditable) {
    var t = String(this.innerText || this.textContent || '');
    if (t !== text) {
      this.innerText = text;
      this.textContent = text;
      this.dispatchEvent(new Event('input', {bubbles:true}));
    }
    return { ok:true, value: String(this.innerText || this.textContent || '') };
  }
  return { ok:false, reason:'not_editable' };
}",
                argsJson,
                sessionId).ConfigureAwait(true);

            string got = ReadTypedValue(syncJson);
            if (!string.Equals(got, want, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "输入未生效（读回 value=\"" + Truncate(got, 40)
                    + "\"，期望=\"" + Truncate(want, 40) + "\"）");
            }

            // 4) 收起下拉联想：否则按 Enter 常会选中推荐热词而不是刚输入的内容
            await DispatchKeyAsync(form, "keyDown", "Escape", sessionId).ConfigureAwait(true);
            await DispatchKeyAsync(form, "keyUp", "Escape", sessionId).ConfigureAwait(true);
            await CallFunctionOnAsync(
                form,
                objectId,
                "function(){ this.focus && this.focus(); }",
                null,
                sessionId).ConfigureAwait(true);
        }

        private static void EnsureTypePrepOk(string resultJson)
        {
            try
            {
                var jo = JObject.Parse(resultJson);
                var v = jo["result"]?["value"] as JObject;
                if (v == null || v["ok"] == null || !(bool)v["ok"])
                {
                    throw new InvalidOperationException("目标不是可 type 的输入控件");
                }
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("type 失败: " + ex.Message);
            }
        }

        private static string ReadTypedValue(string resultJson)
        {
            try
            {
                var jo = JObject.Parse(resultJson);
                var v = jo["result"]?["value"] as JObject;
                if (v == null || v["ok"] == null || !(bool)v["ok"])
                {
                    throw new InvalidOperationException("目标不是可 type 的输入控件");
                }

                return (string)v["value"] ?? "";
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("type 失败: " + ex.Message);
            }
        }

        private static string Truncate(string s, int max)
        {
            if (string.IsNullOrEmpty(s) || s.Length <= max)
            {
                return s ?? "";
            }

            return s.Substring(0, max) + "…";
        }

        public static async Task ScrollIntoViewAsync(
            YiWriteBrowserForm form,
            int backendNodeId,
            string sessionId = null)
        {
            string objectId = await ResolveObjectIdAsync(form, backendNodeId, sessionId).ConfigureAwait(true);
            string resultJson = await CallFunctionOnAsync(
                form,
                objectId,
                ScrollIntoOverflowJs,
                null,
                sessionId).ConfigureAwait(true);
            ThrowIfJsNotOk(resultJson, "滚动失败");
        }

        public static async Task ScrollByDirectionAsync(YiWriteBrowserForm form, string direction)
        {
            string dir = (direction ?? "down").Trim().ToLowerInvariant();
            string dx = "0";
            string dy = "0";
            switch (dir)
            {
                case "up":
                    dy = "-Math.floor(window.innerHeight*0.85)";
                    break;
                case "left":
                    dx = "-Math.floor(window.innerWidth*0.85)";
                    break;
                case "right":
                    dx = "Math.floor(window.innerWidth*0.85)";
                    break;
                default:
                    dy = "Math.floor(window.innerHeight*0.85)";
                    break;
            }

            string expr = "window.scrollBy(" + dx + "," + dy + "); true";
            await form.CallCdpAsync(
                "Runtime.evaluate",
                "{\"expression\":" + QuoteJson(expr) + ",\"returnByValue\":true}")
                .ConfigureAwait(true);
        }

        public static async Task PressAsync(
            YiWriteBrowserForm form,
            string key,
            int? backendNodeId,
            string sessionId = null)
        {
            string keyName = NormalizeKey(key);

            if (backendNodeId.HasValue)
            {
                string objectId = await ResolveObjectIdAsync(form, backendNodeId.Value, sessionId).ConfigureAwait(true);
                await CallFunctionOnAsync(
                    form,
                    objectId,
                    "function(){ this.focus && this.focus(); }",
                    null,
                    sessionId).ConfigureAwait(true);

                // 带 ref 的 Enter：先 Esc 收联想，避免提交高亮推荐词
                if (string.Equals(keyName, "Enter", StringComparison.Ordinal))
                {
                    await DispatchKeyAsync(form, "keyDown", "Escape", sessionId).ConfigureAwait(true);
                    await DispatchKeyAsync(form, "keyUp", "Escape", sessionId).ConfigureAwait(true);
                    await CallFunctionOnAsync(
                        form,
                        objectId,
                        "function(){ this.focus && this.focus(); }",
                        null,
                        sessionId).ConfigureAwait(true);
                }
            }

            await DispatchKeyAsync(form, "keyDown", keyName, sessionId).ConfigureAwait(true);
            await DispatchKeyAsync(form, "keyUp", keyName, sessionId).ConfigureAwait(true);
        }

        public static async Task SelectAsync(
            YiWriteBrowserForm form,
            int backendNodeId,
            string option,
            string sessionId = null)
        {
            string objectId = await ResolveObjectIdAsync(form, backendNodeId, sessionId).ConfigureAwait(true);
            string argsJson = BuildCallArgs(option ?? "");
            string resultJson = await CallFunctionOnAsync(
                form,
                objectId,
                @"function(option) {
  function findSelect(el) {
    if (!el) return null;
    if ((el.tagName || '').toUpperCase() === 'SELECT') return el;
    try {
      if (el.querySelector) {
        var inner = el.querySelector('select');
        if (inner) return inner;
      }
      var prev = el.previousElementSibling;
      if (prev && (prev.tagName || '').toUpperCase() === 'SELECT') return prev;
      var next = el.nextElementSibling;
      if (next && (next.tagName || '').toUpperCase() === 'SELECT') return next;
      if (el.closest) {
        var g = el.closest('.input-group');
        if (g && g.querySelector) {
          var inGroup = g.querySelector('select');
          if (inGroup) return inGroup;
        }
      }
    } catch (eF) {}
    return null;
  }
  function optionPreview(sel) {
    var out = [];
    if (!sel || !sel.options) return '';
    for (var i = 0; i < sel.options.length && out.length < 8; i++) {
      var tx = String(sel.options[i].text || '').replace(/\s+/g, ' ').trim();
      if (tx) out.push(tx);
    }
    return out.join('、');
  }
  var sel = findSelect(this);
  if (!sel) return { ok:false, reason:'not_select', tag: (this.tagName || '?') };
  sel.scrollIntoView({block:'center', inline:'center'});
  sel.focus && sel.focus();
  var exact = -1, contains = -1;
  for (var i = 0; i < sel.options.length; i++) {
    var o = sel.options[i];
    var t = (o.text || '').trim();
    var v = (o.value || '').trim();
    if (v === option || t === option) { exact = i; break; }
    if (contains < 0 && (t.indexOf(option) >= 0 || v.indexOf(option) >= 0)) contains = i;
  }
  var idx = exact >= 0 ? exact : contains;
  if (idx < 0) return { ok:false, reason:'no_match', options: optionPreview(sel) };
  sel.selectedIndex = idx;
  sel.dispatchEvent(new Event('input', {bubbles:true}));
  sel.dispatchEvent(new Event('change', {bubbles:true}));
  return { ok:true };
}",
                argsJson,
                sessionId).ConfigureAwait(true);

            try
            {
                var jo = JObject.Parse(resultJson);
                var v = jo["result"]?["value"] as JObject;
                if (v == null || v["ok"] == null || !(bool)v["ok"])
                {
                    string reason = v != null ? (string)v["reason"] : "failed";
                    if (reason == "no_match")
                    {
                        string hint = v != null ? (string)v["options"] : null;
                        throw new InvalidOperationException(
                            string.IsNullOrWhiteSpace(hint)
                                ? "未找到匹配的下拉选项"
                                : "未找到匹配的下拉选项。可选：" + hint);
                    }

                    string tag = v != null ? (string)v["tag"] : null;
                    throw new InvalidOperationException(
                        string.IsNullOrWhiteSpace(tag)
                            ? "select 未落到原生下拉。请重新 snapshot，对带 options= 的那一行再 select，不要改 click。"
                            : "select 未落到原生下拉（当前是 " + tag
                                + "）。请重新 snapshot，对带 options= 的那一行再 select，不要改 click。");
                }
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("select 失败: " + ex.Message);
            }
        }

        private static async Task DispatchKeyAsync(
            YiWriteBrowserForm form,
            string type,
            string key,
            string sessionId = null)
        {
            int vk;
            string code;
            string text = "";
            switch (key)
            {
                case "Enter":
                    vk = 13; code = "Enter"; text = "\r";
                    break;
                case "Tab":
                    vk = 9; code = "Tab";
                    break;
                case "Escape":
                    vk = 27; code = "Escape";
                    break;
                case "Backspace":
                    vk = 8; code = "Backspace";
                    break;
                case "Delete":
                    vk = 46; code = "Delete";
                    break;
                case "ArrowUp":
                    vk = 38; code = "ArrowUp";
                    break;
                case "ArrowDown":
                    vk = 40; code = "ArrowDown";
                    break;
                case "ArrowLeft":
                    vk = 37; code = "ArrowLeft";
                    break;
                case "ArrowRight":
                    vk = 39; code = "ArrowRight";
                    break;
                case "Home":
                    vk = 36; code = "Home";
                    break;
                case "End":
                    vk = 35; code = "End";
                    break;
                case "PageUp":
                    vk = 33; code = "PageUp";
                    break;
                case "PageDown":
                    vk = 34; code = "PageDown";
                    break;
                default:
                    throw new InvalidOperationException("不支持的按键: " + key);
            }

            var payload = new JObject
            {
                ["type"] = type,
                ["windowsVirtualKeyCode"] = vk,
                ["nativeVirtualKeyCode"] = vk,
                ["code"] = code,
                ["key"] = key
            };
            if (!string.IsNullOrEmpty(text))
            {
                payload["text"] = text;
                payload["unmodifiedText"] = text;
            }

            await form.CallCdpAsync(
                    "Input.dispatchKeyEvent",
                    payload.ToString(Newtonsoft.Json.Formatting.None),
                    sessionId)
                .ConfigureAwait(true);
        }

        private static string NormalizeKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return key;
            }

            string k = key.Trim();
            if (string.Equals(k, "Esc", StringComparison.OrdinalIgnoreCase)
                || string.Equals(k, "Escape", StringComparison.OrdinalIgnoreCase))
            {
                return "Escape";
            }

            if (string.Equals(k, "Return", StringComparison.OrdinalIgnoreCase)
                || string.Equals(k, "Enter", StringComparison.OrdinalIgnoreCase))
            {
                return "Enter";
            }

            // 统一首字母大写的白名单名
            string[] allowed =
            {
                "Enter", "Tab", "Escape", "Backspace", "Delete",
                "ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight",
                "Home", "End", "PageUp", "PageDown"
            };
            foreach (string a in allowed)
            {
                if (string.Equals(a, k, StringComparison.OrdinalIgnoreCase))
                {
                    return a;
                }
            }

            return k;
        }

        private const string PointerPrepareJs =
            @"function(opts) {
  function classNameOf(el) {
    var c = el && el.className;
    if (c && typeof c === 'object' && c.baseVal != null) return String(c.baseVal);
    return String(c || '');
  }
  function ownText(el) {
    var t = '';
    if (!el || !el.childNodes) return t;
    for (var i = 0; i < el.childNodes.length; i++) {
      var n = el.childNodes[i];
      if (n.nodeType === 3) t += n.textContent || '';
    }
    return String(t || '').replace(/\s+/g, ' ').trim();
  }
  function isDetailRevealName(s) {
    if (!s) return false;
    var n = String(s).replace(/\s+/g, ' ').trim();
    if (n.indexOf('查看详情') >= 0) return true;
    if (n.indexOf('查看更多') >= 0) return true;
    return n === '详情';
  }
  function isForceTarget(el) {
    if (!el) return false;
    if (/\bimgtextbtn\b/.test(classNameOf(el))) return true;
    if (isDetailRevealName(ownText(el))) return true;
    var inner = '';
    try { inner = String(el.innerText || '').replace(/\s+/g, ' ').trim(); } catch (eI) {}
    return inner.length <= 12 && isDetailRevealName(inner);
  }
  function parentComposed(el) {
    if (!el) return null;
    if (el.assignedSlot) return el.assignedSlot;
    if (el.parentElement) return el.parentElement;
    var root = el.getRootNode && el.getRootNode();
    if (root && root.host) return root.host;
    return null;
  }
  function isOnTarget(hit, t) {
    var n = hit;
    while (n) {
      if (n === t) return true;
      n = parentComposed(n);
    }
    try { return !!(t && t.contains && hit && t.contains(hit)); } catch (eC) { return false; }
  }
  function previewHit(hit) {
    var tag = ((hit && hit.tagName) || 'div').toLowerCase();
    var text = '';
    try {
      text = (hit.getAttribute && (hit.getAttribute('aria-label') || hit.getAttribute('title'))) || '';
      if (!text) text = String(hit.innerText || hit.textContent || '').replace(/\s+/g, ' ').trim();
    } catch (eP) {}
    if (text.length > 40) text = text.slice(0, 40);
    return text ? (tag + ' ' + text) : tag;
  }
  function doClick(t) {
    if (typeof t.click === 'function') t.click();
    else t.dispatchEvent(new MouseEvent('click', {bubbles:true, cancelable:true, view:window}));
  }
  function scrollIfNeeded(el) {
    if (!el || !el.getBoundingClientRect) return;
    try {
      var box = el.getBoundingClientRect();
      var vw = window.innerWidth || document.documentElement.clientWidth || 0;
      var vh = window.innerHeight || document.documentElement.clientHeight || 0;
      var visible = box.width > 1 && box.height > 1
        && box.bottom > 0 && box.right > 0 && box.top < vh && box.left < vw;
      if (!visible) el.scrollIntoView({block:'nearest', inline:'nearest'});
    } catch (eSc) {}
  }

  opts = opts || {};
  var kind = String(opts.kind || 'click');
  var mode = String(opts.mode || 'pointer');
  var isHover = kind === 'hover';
  var isLegacy = !isHover && mode === 'legacy';

  function imageHrefLink(el) {
    if (!el || !el.querySelector) return null;
    var img = el.querySelector('a[href] img');
    return img && img.closest ? img.closest('a[href]') : null;
  }

  var t = this;
  var tag = (this.tagName || '').toUpperCase();
  var role = '';
  try { role = ((this.getAttribute && this.getAttribute('role')) || '').toLowerCase(); } catch (eR) {}
  var isImg = tag === 'IMG' || role === 'img' || role === 'image';
  if (isImg && !isHover && !isLegacy) {
    scrollIfNeeded(this);
    var ir = { width: 0, height: 0, left: 0, top: 0 };
    try { ir = this.getBoundingClientRect(); } catch (eIb) {}
    if (ir.width <= 0 || ir.height <= 0) {
      return { ok: false, error: '点击被挡住：目标没有可点区域。' };
    }
    return { ok: true, x: ir.left + ir.width / 2, y: ir.top + ir.height / 2 };
  }
  if (!isHover && tag !== 'A' && role !== 'link' && tag !== 'BUTTON' && role !== 'button') {
    var imgA = imageHrefLink(this);
    if (imgA) t = imgA;
  }
  if (isImg) {
    try {
      if (this.closest) {
        var card = this.closest('.imgtext');
        if (card) {
          var btn = card.querySelector('.imgtextbtn');
          t = btn || card;
        } else {
          var hitA = this.closest('a[href], button, [role=button], [role=link], [onclick]');
          if (hitA) t = hitA;
        }
      }
      if (t === this) {
        var p = this.parentElement;
        for (var i = 0; i < 5 && p; i++) {
          try {
            var st = window.getComputedStyle(p);
            if (st && st.cursor === 'pointer') { t = p; break; }
          } catch (eS) {}
          p = p.parentElement;
        }
      }
    } catch (eC) {}
  }
  scrollIfNeeded(t);
  var force = !isHover && isForceTarget(t);
  var r = { width: 0, height: 0, left: 0, top: 0 };
  try { r = t.getBoundingClientRect(); } catch (eB) {}
  var x = r.left + r.width / 2;
  var y = r.top + r.height / 2;
  if (force) {
    if (isLegacy || r.width <= 0 || r.height <= 0) {
      doClick(t);
      return { ok: true, fallback: true };
    }
    return { ok: true, x: x, y: y };
  }
  if (r.width <= 0 || r.height <= 0) {
    return {
      ok: false,
      error: isHover
        ? '没有可悬停区域。请悬停可见的父节点或卡片。'
        : '点击被挡住：目标没有可点区域。'
    };
  }
  var hit = null;
  try { hit = document.elementFromPoint(x, y); } catch (eH) {}
  if (!hit) {
    return { ok: false, error: '点击被挡住：目标不在当前视口内，请先 scroll 再点。' };
  }
  if (!isOnTarget(hit, t)) {
    return { ok: false, error: '点击被挡住：挡在上面的是「' + previewHit(hit) + '」。请先操作该蒙层（如接受 Cookie / 关闭），再点原来的目标。' };
  }
  if (isLegacy) {
    doClick(t);
    return { ok: true };
  }
  return { ok: true, x: x, y: y };
}";

        private const string ScrollIntoOverflowJs =
            @"function() {
  function isScrollable(el) {
    if (!el || el === document.body || el === document.documentElement) return false;
    var st;
    try { st = window.getComputedStyle(el); } catch (e) { return false; }
    if (!st) return false;
    var oy = st.overflowY, ox = st.overflowX, o = st.overflow;
    var y = (oy === 'auto' || oy === 'scroll' || o === 'auto' || o === 'scroll') && el.scrollHeight > el.clientHeight + 1;
    var x = (ox === 'auto' || ox === 'scroll' || o === 'auto' || o === 'scroll') && el.scrollWidth > el.clientWidth + 1;
    return y || x;
  }
  function nearestScrollable(el) {
    var p = el && el.parentElement;
    while (p && p !== document.body && p !== document.documentElement) {
      if (isScrollable(p)) return p;
      p = p.parentElement;
    }
    return null;
  }
  var t = this;
  if (!t || !t.isConnected) return { ok: false, error: '滚动目标已不在页面上' };
  try { t.scrollIntoView({block:'center', inline:'nearest'}); } catch (eV) {}
  var box = nearestScrollable(t);
  if (box) {
    var er = t.getBoundingClientRect();
    var br = box.getBoundingClientRect();
    box.scrollTop += (er.top + er.height / 2) - (br.top + br.height / 2);
    box.scrollLeft += (er.left + er.width / 2) - (br.left + br.width / 2);
  }
  return { ok: true };
}";

        private static void ThrowIfJsNotOk(string resultJson, string fallback)
        {
            try
            {
                var jo = JObject.Parse(resultJson ?? "{}");
                var ex = jo["exceptionDetails"];
                if (ex != null && ex.Type != JTokenType.Null)
                {
                    string msg = (string)ex["exception"]?["description"]
                        ?? (string)ex["text"]
                        ?? fallback;
                    throw new InvalidOperationException(msg);
                }

                var v = jo["result"]?["value"] as JObject;
                if (v != null
                    && v["ok"] != null
                    && v["ok"].Type == JTokenType.Boolean
                    && !(bool)v["ok"])
                {
                    throw new InvalidOperationException(
                        (string)v["error"] ?? fallback);
                }
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("工具结果无法解析: " + (ex.Message ?? ""));
            }
        }

        internal static async Task<string> ResolveObjectIdAsync(
            YiWriteBrowserForm form,
            int backendNodeId,
            string sessionId = null)
        {
            string json;
            try
            {
                json = await form.CallCdpAsync(
                    "DOM.resolveNode",
                    "{\"backendNodeId\":" + backendNodeId.ToString(CultureInfo.InvariantCulture) + "}",
                    sessionId)
                    .ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "无法解析页面节点（可能跨进程 iframe）: " + (ex.Message ?? ""));
            }

            var jo = JObject.Parse(json);
            string objectId = (string)jo["object"]?["objectId"];
            if (string.IsNullOrEmpty(objectId))
            {
                throw new InvalidOperationException("无法解析页面节点（backendNodeId=" + backendNodeId + "）");
            }

            return objectId;
        }

        internal static async Task<string> CallFunctionOnAsync(
            YiWriteBrowserForm form,
            string objectId,
            string functionDeclaration,
            string argumentsJsonArrayOrNull,
            string sessionId = null)
        {
            var payload = new JObject
            {
                ["objectId"] = objectId,
                ["functionDeclaration"] = functionDeclaration,
                ["returnByValue"] = true,
                ["awaitPromise"] = true
            };
            if (!string.IsNullOrEmpty(argumentsJsonArrayOrNull))
            {
                payload["arguments"] = JArray.Parse(argumentsJsonArrayOrNull);
            }

            return await form.CallCdpAsync(
                    "Runtime.callFunctionOn",
                    payload.ToString(Newtonsoft.Json.Formatting.None),
                    sessionId)
                .ConfigureAwait(true);
        }

        private static string BuildCallArgs(string value)
        {
            var arr = new JArray { new JObject { ["value"] = value ?? "" } };
            return arr.ToString(Newtonsoft.Json.Formatting.None);
        }

        private static string BuildCallObject(JObject value)
        {
            var arr = new JArray { new JObject { ["value"] = value ?? new JObject() } };
            return arr.ToString(Newtonsoft.Json.Formatting.None);
        }

        private static bool ReadJsFallback(string resultJson)
        {
            try
            {
                var jo = JObject.Parse(resultJson ?? "{}");
                var v = jo["result"]?["value"] as JObject;
                return v != null
                    && v["fallback"] != null
                    && v["fallback"].Type == JTokenType.Boolean
                    && (bool)v["fallback"];
            }
            catch
            {
                return false;
            }
        }

        private static void ReadJsPoint(string resultJson, out double x, out double y)
        {
            x = 0;
            y = 0;
            try
            {
                var jo = JObject.Parse(resultJson ?? "{}");
                var v = jo["result"]?["value"] as JObject;
                if (v == null || v["x"] == null || v["y"] == null)
                {
                    throw new InvalidOperationException("目标没有可点区域");
                }

                x = (double)v["x"];
                y = (double)v["y"];
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("无法读取点击坐标: " + (ex.Message ?? ""));
            }
        }

        private static async Task DispatchCdpMouseAsync(
            YiWriteBrowserForm form,
            string sessionId,
            string type,
            double x,
            double y,
            bool pressed)
        {
            var payload = new JObject
            {
                ["type"] = type,
                ["x"] = Math.Round(x, 2),
                ["y"] = Math.Round(y, 2),
                ["pointerType"] = "mouse"
            };
            if (string.Equals(type, "mousePressed", StringComparison.Ordinal)
                || string.Equals(type, "mouseReleased", StringComparison.Ordinal))
            {
                payload["button"] = "left";
                payload["clickCount"] = 1;
                payload["buttons"] = pressed ? 1 : 0;
            }

            await form.CallCdpAsync(
                    "Input.dispatchMouseEvent",
                    payload.ToString(Newtonsoft.Json.Formatting.None),
                    sessionId)
                .ConfigureAwait(true);
        }

        private static bool IsCheckLikeRole(string role)
        {
            if (string.IsNullOrWhiteSpace(role))
            {
                return false;
            }

            return string.Equals(role, "checkbox", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "radio", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "switch", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "checkable", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "menuitemcheckbox", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "menuitemradio", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Vue / Element 复选框的真实 input 常是 0 盒。指针打在「我的学生」文案上不会改 checked。
        /// 对原生 input 或包裹 label 发一次 click。
        /// </summary>
        private static async Task ToggleCheckAsync(
            YiWriteBrowserForm form,
            int backendNodeId,
            string sessionId)
        {
            string objectId = await ResolveObjectIdAsync(form, backendNodeId, sessionId).ConfigureAwait(true);
            string resultJson = await CallFunctionOnAsync(
                form,
                objectId,
                @"function() {
  function findInput(el) {
    if (!el) return null;
    var tag = (el.tagName || '').toUpperCase();
    var typ = String(el.type || '').toLowerCase();
    if (tag === 'INPUT' && (typ === 'checkbox' || typ === 'radio')) return el;
    try {
      var q = el.querySelector && el.querySelector('input[type=checkbox], input[type=radio]');
      if (q) return q;
    } catch (e0) {}
    try {
      var lab = el.closest && el.closest('label');
      if (lab) {
        var q2 = lab.querySelector && lab.querySelector('input[type=checkbox], input[type=radio]');
        if (q2) return q2;
        var id = lab.htmlFor || (lab.getAttribute && lab.getAttribute('for')) || '';
        if (id) {
          var by = document.getElementById(id);
          if (by && (by.type === 'checkbox' || by.type === 'radio')) return by;
        }
      }
    } catch (e1) {}
    return null;
  }
  var t = findInput(this) || this;
  try { if (t.focus) t.focus(); } catch (eF) {}
  try {
    if (typeof t.click === 'function') t.click();
    else t.dispatchEvent(new MouseEvent('click', {bubbles:true, cancelable:true, view:window}));
  } catch (eC) {}
  return { ok: true };
}",
                null,
                sessionId).ConfigureAwait(true);
            ThrowIfJsNotOk(resultJson, "无法切换复选框");
        }

        private static async Task DispatchCdpClickAsync(
            YiWriteBrowserForm form,
            string objectId,
            double x,
            double y,
            string sessionId)
        {
            await DispatchCdpMouseAsync(form, sessionId, "mouseMoved", x, y, false).ConfigureAwait(true);
            await DispatchCdpMouseAsync(form, sessionId, "mousePressed", x, y, true).ConfigureAwait(true);
            await DispatchCdpMouseAsync(form, sessionId, "mouseReleased", x, y, false).ConfigureAwait(true);
            if (form.HasNewWindowDuringTool)
            {
                return;
            }

            try
            {
                await CallFunctionOnAsync(
                    form,
                    objectId,
                    @"function(x, y) {
  if (!this || !this.isConnected) return false;
  var hit = null;
  try { hit = document.elementFromPoint(x, y); } catch (eH) {}
  var same = !hit || hit === this;
  try {
    same = same || (this.contains && this.contains(hit)) || (hit.contains && hit.contains(this));
  } catch (eC) {}
  if (!same) return false;
  var n = hit || this;
  try {
    n.dispatchEvent(new MouseEvent('click', {
      bubbles: true, cancelable: true, view: window,
      clientX: x, clientY: y, buttons: 0
    }));
  } catch (e) {}
  return true;
}",
                    BuildCallArgsPair(x, y),
                    sessionId).ConfigureAwait(true);
            }
            catch
            {
            }
        }

        private static async Task DispatchCdpHoverAsync(
            YiWriteBrowserForm form,
            string objectId,
            double x,
            double y,
            string sessionId)
        {
            await DispatchCdpMouseAsync(form, sessionId, "mouseMoved", x, y, false).ConfigureAwait(true);
            await CallFunctionOnAsync(
                form,
                objectId,
                @"function(x, y) {
  var opts = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y };
  try { this.dispatchEvent(new MouseEvent('mousemove', opts)); } catch (e0) {}
  try {
    if (typeof PointerEvent === 'function') {
      this.dispatchEvent(new PointerEvent('pointermove', {
        bubbles: true, cancelable: true, view: window,
        clientX: x, clientY: y, pointerId: 1, pointerType: 'mouse'
      }));
    }
  } catch (e1) {}
  return true;
}",
                BuildCallArgsPair(x, y),
                sessionId).ConfigureAwait(true);
        }

        private static string BuildCallArgsPair(double x, double y)
        {
            var arr = new JArray
            {
                new JObject { ["value"] = x },
                new JObject { ["value"] = y }
            };
            return arr.ToString(Newtonsoft.Json.Formatting.None);
        }

        private static string QuoteJson(string s)
        {
            return JToken.FromObject(s ?? "").ToString(Newtonsoft.Json.Formatting.None);
        }
    }
}
