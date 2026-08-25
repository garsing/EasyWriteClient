using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// AX 无名时补 aria-label / title / alt（图标按钮）。不穿透 Shadow，不上坐标。
    /// </summary>
    internal static class BrowserAccessibleNameFill
    {
        private const int MaxAttempts = 30;

        private const string ReadNameJs =
            @"function() {
  function t(s) { return String(s || '').replace(/\s+/g, ' ').trim().slice(0, 120); }
  var el = this;
  if (!el) return '';
  try {
    var a = el.getAttribute && el.getAttribute('aria-label');
    function bad(s) {
      s = t(s);
      if (!s) return true;
      if (s === '不能为空' || s === '必填' || s === '此项必填') return true;
      if (s.indexOf('输入格式不正确') >= 0) return true;
      return false;
    }
    if (t(a) && !bad(a)) return t(a);
    if (t(el.title) && !bad(el.title)) return t(el.title);
    if (t(el.alt)) return t(el.alt);
    var img = el.querySelector && el.querySelector('img[alt]');
    if (img && t(img.alt)) return t(img.alt);
    var st = el.querySelector && el.querySelector('svg title');
    if (st && t(st.textContent)) return t(st.textContent);
    var area = el.querySelector && el.querySelector('area[alt]');
    if (area && t(area.alt)) return t(area.alt);
  } catch (e) {}
  return '';
}";

        public static async Task FillEmptyAsync(
            YiWriteBrowserForm form,
            BrowserAxBuildResult built,
            string sessionId)
        {
            if (form == null || built == null || built.Refs == null || built.Refs.Count == 0)
            {
                return;
            }

            int attempts = 0;
            var filled = new List<KeyValuePair<string, string>>();
            foreach (var kv in built.Refs)
            {
                if (attempts >= MaxAttempts)
                {
                    break;
                }

                BrowserRefEntry e = kv.Value;
                if (e == null || !e.BackendDomNodeId.HasValue)
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(e.Name))
                {
                    continue;
                }

                attempts++;
                string sid = string.IsNullOrWhiteSpace(e.CdpSessionId) ? sessionId : e.CdpSessionId;
                try
                {
                    string objectId = await BrowserInteractEngine
                        .ResolveObjectIdAsync(form, e.BackendDomNodeId.Value, sid)
                        .ConfigureAwait(true);
                    string json = await BrowserInteractEngine
                        .CallFunctionOnAsync(form, objectId, ReadNameJs, null, sid)
                        .ConfigureAwait(true);
                    string name = ReadReturnedString(json);
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    if (BrowserAxTreeBuilder.IsValidatorAccessibleName(name))
                    {
                        continue;
                    }

                    e.Name = name;
                    filled.Add(new KeyValuePair<string, string>(kv.Key, name));
                }
                catch
                {
                }
            }

            if (filled.Count == 0 || string.IsNullOrEmpty(built.TreeText))
            {
                return;
            }

            string tree = built.TreeText;
            foreach (var kv in filled)
            {
                tree = InsertNameAfterRef(tree, kv.Key, kv.Value);
            }

            built.TreeText = tree;
        }

        /// <summary>树格式：<c>- button ref=e5</c> → <c>- button ref=e5 "设置"</c></summary>
        private static string InsertNameAfterRef(string tree, string refId, string name)
        {
            string needle = "ref=" + refId;
            int at = 0;
            while (at < tree.Length)
            {
                int found = tree.IndexOf(needle, at, StringComparison.Ordinal);
                if (found < 0)
                {
                    return tree;
                }

                int end = found + needle.Length;
                if (end < tree.Length && char.IsDigit(tree[end]))
                {
                    at = end;
                    continue;
                }

                if (end < tree.Length - 1 && tree[end] == ' ' && tree[end + 1] == '"')
                {
                    return tree;
                }

                string quoted = " \"" + Escape(name) + "\"";
                return tree.Substring(0, end) + quoted + tree.Substring(end);
            }

            return tree;
        }

        private static string Escape(string s)
        {
            return (s ?? "")
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", " ")
                .Replace("\n", " ");
        }

        private static string ReadReturnedString(string resultJson)
        {
            try
            {
                var jo = JObject.Parse(resultJson ?? "{}");
                var tok = jo["result"]?["value"];
                return tok == null || tok.Type == JTokenType.Null ? "" : tok.ToString();
            }
            catch
            {
                return "";
            }
        }
    }
}
