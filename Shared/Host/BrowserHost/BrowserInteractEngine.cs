using System;
using System.Globalization;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>WebView2 CDP：click / type / scroll / press / select。</summary>
    internal static class BrowserInteractEngine
    {
        public static async Task<DomNodeProbe> ProbeAsync(YiWriteBrowserForm form, int backendNodeId)
        {
            string objectId = await ResolveObjectIdAsync(form, backendNodeId).ConfigureAwait(true);
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
    pageHasPasswordInput: pageHasPwd
  };
}",
                null).ConfigureAwait(true);

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
                }
            }
            catch
            {
            }

            return probe;
        }

        public static async Task<DomNodeProbe> ProbePagePasswordAsync(YiWriteBrowserForm form)
        {
            string json = await form.CallCdpAsync(
                "Runtime.evaluate",
                "{\"expression\":\"!!document.querySelector('input[type=password]')\",\"returnByValue\":true}")
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

        public static async Task ClickAsync(YiWriteBrowserForm form, int backendNodeId)
        {
            string objectId = await ResolveObjectIdAsync(form, backendNodeId).ConfigureAwait(true);
            // 临时用简单 click：验证「完整鼠标序列 + 等导航」是否可去掉
            await CallFunctionOnAsync(
                form,
                objectId,
                @"function() {
  this.scrollIntoView({block:'center', inline:'center'});
  if (typeof this.click === 'function') { this.click(); }
  else {
    var e = new MouseEvent('click', {bubbles:true, cancelable:true, view:window});
    this.dispatchEvent(e);
  }
}",
                null).ConfigureAwait(true);

            // 临时关闭：点击后短等导航
            // try
            // {
            //     await form.WaitNavigationAsync(TimeSpan.FromSeconds(3)).ConfigureAwait(true);
            // }
            // catch
            // {
            // }

            /* 增强版（完整鼠标序列 + closest a[href] + 等导航）暂存对照：
            await CallFunctionOnAsync(
                form,
                objectId,
                @"function() {
  this.scrollIntoView({block:'center', inline:'center'});
  var t = this;
  try {
    if (t.closest) {
      var a = t.closest('a[href]');
      if (a) t = a;
    }
  } catch (e) {}
  function fire(type) {
    try {
      t.dispatchEvent(new MouseEvent(type, {bubbles:true, cancelable:true, view:window, buttons:1}));
    } catch (e2) {}
  }
  fire('pointerdown');
  fire('mousedown');
  fire('mouseup');
  fire('click');
  try {
    if (typeof t.click === 'function') t.click();
  } catch (e3) {}
  return !!(t && (t.href || t.tagName));
}",
                null).ConfigureAwait(true);
            try
            {
                await form.WaitNavigationAsync(TimeSpan.FromSeconds(3)).ConfigureAwait(true);
            }
            catch
            {
            }
            */
        }

        public static async Task TypeAsync(YiWriteBrowserForm form, int backendNodeId, string text)
        {
            string want = text ?? "";
            string objectId = await ResolveObjectIdAsync(form, backendNodeId).ConfigureAwait(true);

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
                null).ConfigureAwait(true);

            EnsureTypePrepOk(prepJson);

            // 2) CDP insertText：走浏览器输入通道，框架更容易收到（比单纯改 value 稳）
            var insertPayload = new JObject { ["text"] = want };
            await form.CallCdpAsync(
                "Input.insertText",
                insertPayload.ToString(Newtonsoft.Json.Formatting.None))
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
                argsJson).ConfigureAwait(true);

            string got = ReadTypedValue(syncJson);
            if (!string.Equals(got, want, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "输入未生效（读回 value=\"" + Truncate(got, 40)
                    + "\"，期望=\"" + Truncate(want, 40) + "\"）");
            }

            // 4) 收起下拉联想：否则按 Enter 常会选中推荐热词而不是刚输入的内容
            await DispatchKeyAsync(form, "keyDown", "Escape").ConfigureAwait(true);
            await DispatchKeyAsync(form, "keyUp", "Escape").ConfigureAwait(true);
            await CallFunctionOnAsync(
                form,
                objectId,
                "function(){ this.focus && this.focus(); }",
                null).ConfigureAwait(true);
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

        public static async Task ScrollIntoViewAsync(YiWriteBrowserForm form, int backendNodeId)
        {
            string objectId = await ResolveObjectIdAsync(form, backendNodeId).ConfigureAwait(true);
            await CallFunctionOnAsync(
                form,
                objectId,
                "function(){ this.scrollIntoView({block:'center', inline:'nearest'}); }",
                null).ConfigureAwait(true);
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

        public static async Task PressAsync(YiWriteBrowserForm form, string key, int? backendNodeId)
        {
            string keyName = NormalizeKey(key);

            if (backendNodeId.HasValue)
            {
                string objectId = await ResolveObjectIdAsync(form, backendNodeId.Value).ConfigureAwait(true);
                await CallFunctionOnAsync(
                    form,
                    objectId,
                    "function(){ this.focus && this.focus(); }",
                    null).ConfigureAwait(true);

                // 带 ref 的 Enter：先 Esc 收联想，避免提交高亮推荐词
                if (string.Equals(keyName, "Enter", StringComparison.Ordinal))
                {
                    await DispatchKeyAsync(form, "keyDown", "Escape").ConfigureAwait(true);
                    await DispatchKeyAsync(form, "keyUp", "Escape").ConfigureAwait(true);
                    await CallFunctionOnAsync(
                        form,
                        objectId,
                        "function(){ this.focus && this.focus(); }",
                        null).ConfigureAwait(true);
                }
            }

            await DispatchKeyAsync(form, "keyDown", keyName).ConfigureAwait(true);
            await DispatchKeyAsync(form, "keyUp", keyName).ConfigureAwait(true);
        }

        public static async Task SelectAsync(YiWriteBrowserForm form, int backendNodeId, string option)
        {
            string objectId = await ResolveObjectIdAsync(form, backendNodeId).ConfigureAwait(true);
            string argsJson = BuildCallArgs(option ?? "");
            string resultJson = await CallFunctionOnAsync(
                form,
                objectId,
                @"function(option) {
  this.scrollIntoView({block:'center', inline:'center'});
  this.focus && this.focus();
  var tag = (this.tagName || '').toUpperCase();
  if (tag === 'SELECT') {
    var exact = -1, contains = -1;
    for (var i = 0; i < this.options.length; i++) {
      var o = this.options[i];
      var t = (o.text || '').trim();
      var v = (o.value || '').trim();
      if (v === option || t === option) { exact = i; break; }
      if (contains < 0 && (t.indexOf(option) >= 0 || v.indexOf(option) >= 0)) contains = i;
    }
    var idx = exact >= 0 ? exact : contains;
    if (idx < 0) return { ok:false, reason:'no_match' };
    this.selectedIndex = idx;
    this.dispatchEvent(new Event('input', {bubbles:true}));
    this.dispatchEvent(new Event('change', {bubbles:true}));
    return { ok:true };
  }
  return { ok:false, reason:'not_select' };
}",
                argsJson).ConfigureAwait(true);

            try
            {
                var jo = JObject.Parse(resultJson);
                var v = jo["result"]?["value"] as JObject;
                if (v == null || v["ok"] == null || !(bool)v["ok"])
                {
                    string reason = v != null ? (string)v["reason"] : "failed";
                    if (reason == "no_match")
                    {
                        throw new InvalidOperationException("未找到匹配的下拉选项");
                    }

                    throw new InvalidOperationException("目标不是可 select 的下拉控件");
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

        private static async Task DispatchKeyAsync(YiWriteBrowserForm form, string type, string key)
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

            await form.CallCdpAsync("Input.dispatchKeyEvent", payload.ToString(Newtonsoft.Json.Formatting.None))
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

        private static async Task<string> ResolveObjectIdAsync(YiWriteBrowserForm form, int backendNodeId)
        {
            string json = await form.CallCdpAsync(
                "DOM.resolveNode",
                "{\"backendNodeId\":" + backendNodeId.ToString(CultureInfo.InvariantCulture) + "}")
                .ConfigureAwait(true);
            var jo = JObject.Parse(json);
            string objectId = (string)jo["object"]?["objectId"];
            if (string.IsNullOrEmpty(objectId))
            {
                throw new InvalidOperationException("无法解析页面节点（backendNodeId=" + backendNodeId + "）");
            }

            return objectId;
        }

        private static async Task<string> CallFunctionOnAsync(
            YiWriteBrowserForm form,
            string objectId,
            string functionDeclaration,
            string argumentsJsonArrayOrNull)
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

            return await form.CallCdpAsync("Runtime.callFunctionOn", payload.ToString(Newtonsoft.Json.Formatting.None))
                .ConfigureAwait(true);
        }

        private static string BuildCallArgs(string value)
        {
            var arr = new JArray { new JObject { ["value"] = value ?? "" } };
            return arr.ToString(Newtonsoft.Json.Formatting.None);
        }

        private static string QuoteJson(string s)
        {
            return JToken.FromObject(s ?? "").ToString(Newtonsoft.Json.Formatting.None);
        }
    }
}
