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
        }

        public static async Task TypeAsync(YiWriteBrowserForm form, int backendNodeId, string text)
        {
            string objectId = await ResolveObjectIdAsync(form, backendNodeId).ConfigureAwait(true);
            // 参数经 CDP 传值，避免拼进脚本字符串
            string argsJson = BuildCallArgs(text ?? "");
            await CallFunctionOnAsync(
                form,
                objectId,
                @"function(text) {
  this.scrollIntoView({block:'center', inline:'center'});
  this.focus();
  if (typeof this.select === 'function') { try { this.select(); } catch(e) {} }
  if ('value' in this) {
    this.value = '';
    this.dispatchEvent(new Event('input', {bubbles:true}));
    this.value = text;
    this.dispatchEvent(new Event('input', {bubbles:true}));
    this.dispatchEvent(new Event('change', {bubbles:true}));
  } else if (this.isContentEditable) {
    this.innerText = '';
    this.textContent = text;
    this.dispatchEvent(new Event('input', {bubbles:true}));
  }
}",
                argsJson).ConfigureAwait(true);
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
            if (backendNodeId.HasValue)
            {
                string objectId = await ResolveObjectIdAsync(form, backendNodeId.Value).ConfigureAwait(true);
                await CallFunctionOnAsync(
                    form,
                    objectId,
                    "function(){ this.focus && this.focus(); }",
                    null).ConfigureAwait(true);
            }

            string keyName = NormalizeKey(key);
            // keyDown + keyUp
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
