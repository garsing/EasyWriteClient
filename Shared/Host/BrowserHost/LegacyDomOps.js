/**
 * ClickMode=legacy 共用 DOM 树 / 操作（WebView2 Runtime.evaluate；扩展 content-script 仍走原实现）。
 */
(function (root) {
  var MAX_CHARS = 36000;
  var MAX_NODES = 300;

  function parentComposedTop(el) {
    if (!el) return null;
    if (el.assignedSlot) return el.assignedSlot;
    if (el.parentElement) return el.parentElement;
    var r = el.getRootNode && el.getRootNode();
    if (r && r.host) return r.host;
    return null;
  }

  function hostControl(el) {
    var n = el;
    while (n) {
      var tag = (n.tagName || "").toUpperCase();
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return n;
      n = parentComposedTop(n);
    }
    return el;
  }

  function isTypeableInput(el) {
    if (!el || el.nodeType !== 1) return false;
    if (el.isContentEditable) return true;
    var tag = (el.tagName || "").toLowerCase();
    if (tag === "textarea") return true;
    if (tag !== "input") return false;
    var typ = String((el.getAttribute && el.getAttribute("type")) || el.type || "text").toLowerCase();
    return !/^(button|submit|reset|checkbox|radio|file|image|hidden)$/.test(typ);
  }

  function onlyTypeableIn(root) {
    if (!root || !root.querySelectorAll) return null;
    var list = root.querySelectorAll("input, textarea, [contenteditable='true'], [contenteditable='']");
    var hit = null;
    for (var i = 0; i < list.length; i++) {
      if (!isTypeableInput(list[i])) continue;
      if (hit) return null;
      hit = list[i];
    }
    return hit;
  }

  function labelControl(el) {
    if (!el) return null;
    try {
      var lab = (el.tagName || "").toLowerCase() === "label" ? el : (el.closest && el.closest("label"));
      if (!lab) return null;
      var id = lab.htmlFor || (lab.getAttribute && lab.getAttribute("for"));
      if (id) {
        var by = document.getElementById(id);
        if (isTypeableInput(by)) return by;
      }
      return onlyTypeableIn(lab);
    } catch (e) {
      return null;
    }
  }

  function nameHint(el) {
    if (!el) return "";
    try {
      var aria = el.getAttribute && el.getAttribute("aria-label");
      if (aria && String(aria).trim()) return String(aria).trim();
    } catch (e1) {}
    var own = "";
    var nodes = el.childNodes || [];
    for (var i = 0; i < nodes.length; i++) {
      if (nodes[i].nodeType === 3) own += nodes[i].textContent || "";
    }
    own = own.replace(/\s+/g, " ").trim();
    return own.length > 0 && own.length <= 40 ? own : "";
  }

  function findTypeableByName(hint) {
    if (!hint) return null;
    var list = document.querySelectorAll("input, textarea, [contenteditable='true'], [contenteditable='']");
    var hits = [];
    for (var i = 0; i < list.length; i++) {
      var c = list[i];
      if (!isTypeableInput(c)) continue;
      var an = String((c.getAttribute && (c.getAttribute("aria-label") || c.getAttribute("name"))) || "").trim();
      if (an && an === hint) hits.push(c);
    }
    return hits.length === 1 ? hits[0] : null;
  }

  function findTypeable(el, entry) {
    if (!el) return null;
    var up = hostControl(el);
    if (isTypeableInput(up)) return up;
    if (el.isContentEditable) return el;
    var fromLabel = labelControl(el);
    if (fromLabel) return fromLabel;
    var hint = (entry && (entry.name || (entry.probe && entry.probe.accessibleName))) || nameHint(el);
    var byName = findTypeableByName(hint);
    if (byName) return byName;
    var only = onlyTypeableIn(el);
    if (only) return only;
    try {
      if (isTypeableInput(el.previousElementSibling)) return el.previousElementSibling;
      if (isTypeableInput(el.nextElementSibling)) return el.nextElementSibling;
      var g = el.closest && el.closest(".input-group");
      if (g) {
        var inG = onlyTypeableIn(g);
        if (inG) return inG;
      }
    } catch (e2) {}
    return null;
  }

  function looksLikePassword(el) {
    if (!el) return false;
    var typ = String((el.getAttribute && el.getAttribute("type")) || el.type || "").toLowerCase();
    if (typ === "password") return true;
    var ac = (el.getAttribute && el.getAttribute("autocomplete")) || "";
    return /password/i.test(ac);
  }

  function looksLikeFileChooser(node) {
    if (!node) return false;
    try {
      var tag = (node.tagName || "").toLowerCase();
      var typ = String((node.getAttribute && node.getAttribute("type")) || node.type || "").toLowerCase();
      if (tag === "input" && typ === "file") return true;
      if (node.querySelector && node.querySelector('input[type="file"]')) return true;
      var lab = node.closest && node.closest("label");
      if (lab) {
        if (lab.querySelector && lab.querySelector('input[type="file"]')) return true;
        var id = lab.htmlFor || lab.getAttribute("for");
        if (id) {
          var t = document.getElementById(id);
          if (t && String(t.type || "").toLowerCase() === "file") return true;
        }
      }
    } catch (e3) {}
    return false;
  }

  function looksLikeNotEditable(el) {
    if (!el) return true;
    try {
      if (el.disabled) return true;
      if (el.matches && el.matches(":disabled")) return true;
      if (el.readOnly) return true;
      var ad = el.getAttribute && el.getAttribute("aria-disabled");
      if (ad != null && String(ad).toLowerCase() !== "false") return true;
    } catch (e4) {}
    return false;
  }

  function isTemporalInput(el) {
    if (!el || (el.tagName || "").toUpperCase() !== "INPUT") return false;
    return /^(date|time|datetime-local|month|week)$/i.test(String(el.type || ""));
  }

  function pad2(n) {
    return String(n).length < 2 ? "0" + n : String(n);
  }

  function ymd(d) {
    return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
  }

  function fillTemporal(el, raw) {
    var s = String(raw || "").trim();
    var today = new Date();
    if (s === "今天" || /^today$/i.test(s)) s = ymd(today);
    else if (s === "昨天" || /^yesterday$/i.test(s)) {
      var y = new Date(today);
      y.setDate(y.getDate() - 1);
      s = ymd(y);
    } else if (s === "明天" || /^tomorrow$/i.test(s)) {
      var t = new Date(today);
      t.setDate(t.getDate() + 1);
      s = ymd(t);
    }
    try {
      var proto = window.HTMLInputElement.prototype;
      var desc = Object.getOwnPropertyDescriptor(proto, "value");
      if (desc && desc.set) desc.set.call(el, s);
      else el.value = s;
    } catch (e5) {
      el.value = s;
    }
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    return { ok: true, value: el.value };
  }

  function setNativeValue(el, value) {
    var tag = (el.tagName || "").toLowerCase();
    var proto = tag === "textarea" ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
    var desc = Object.getOwnPropertyDescriptor(proto, "value");
    if (desc && desc.set) desc.set.call(el, value);
    else el.value = value;
  }

  function fillTextControl(el, text) {
    el.focus();
    try { el.select && el.select(); } catch (e6) {}
    setNativeValue(el, "");
    el.dispatchEvent(new Event("input", { bubbles: true }));
    setNativeValue(el, text);
    try {
      el.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: true, inputType: "insertText", data: text }));
    } catch (e7) {
      el.dispatchEvent(new Event("input", { bubbles: true }));
    }
    el.dispatchEvent(new Event("change", { bubbles: true }));
    return { ok: true, value: el.value };
  }

  function roleOf(el) {
    var r = el.getAttribute && el.getAttribute("role");
    if (r) return r;
    var tag = (el.tagName || "").toLowerCase();
    if (tag === "a") return "link";
    if (tag === "button") return "button";
    if (tag === "input") {
      var t = (el.getAttribute("type") || "text").toLowerCase();
      if (t === "checkbox") return "checkbox";
      if (t === "radio") return "radio";
      if (t === "submit" || t === "button") return "button";
      return "textbox";
    }
    if (tag === "textarea") return "textbox";
    if (tag === "select") return "combobox";
    if (el.isContentEditable) return "textbox";
    if (tag === "img") return "image";
    if (tag === "iframe" || tag === "frame") return tag;
    if (/^h[1-6]$/.test(tag)) return "heading";
    if (tag === "li") return "listitem";
    if (tag === "tr") return "row";
    if (tag === "td" || tag === "th") return "cell";
    return tag || "generic";
  }

  function isValidatorText(s) {
    var n = String(s || "").replace(/\s+/g, " ").trim();
    if (!n) return false;
    if (n === "不能为空" || n === "必填" || n === "此项必填") return true;
    if (n.indexOf("输入格式不正确") >= 0) return true;
    var low = n.toLowerCase();
    return low === "required" || low === "this field is required";
  }

  function nameOf(el) {
    var tag = (el.tagName || "").toLowerCase();
    function t(s) { return String(s || "").replace(/\s+/g, " ").trim(); }
    if (tag === "html" || tag === "body") return t(document.title).slice(0, 120);
    if (isRowHoverControl(el)) return hoverIconLabel(el);
    var aria = el.getAttribute && el.getAttribute("aria-label");
    if (t(aria) && !isValidatorText(aria)) return t(aria).slice(0, 120);
    if (t(el.title) && !isValidatorText(el.title)) return t(el.title).slice(0, 120);
    if (t(el.alt)) return t(el.alt).slice(0, 120);
    if (el.placeholder && !isValidatorText(el.placeholder)) return String(el.placeholder).trim().slice(0, 120);
    var labels = el.labels ? Array.prototype.slice.call(el.labels) : [];
    for (var i = 0; i < labels.length; i++) {
      var n = t(labels[i].innerText);
      if (n && !isValidatorText(n)) return n.slice(0, 120);
    }
    var text = t(el.innerText || el.textContent);
    return isValidatorText(text) ? "" : text.slice(0, 120);
  }

  function cssPath(el) {
    if (!el || !el.tagName) return "";
    if (el.id) {
      try { return "#" + CSS.escape(el.id); } catch (e8) { return "#" + el.id; }
    }
    var parts = [];
    var cur = el;
    while (cur && cur.nodeType === 1 && parts.length < 6) {
      var part = cur.tagName.toLowerCase();
      if (cur.id) {
        try { parts.unshift("#" + CSS.escape(cur.id)); } catch (e9) { parts.unshift("#" + cur.id); }
        break;
      }
      var parent = cur.parentElement;
      if (parent) {
        var siblings = Array.prototype.filter.call(parent.children, function (c) { return c.tagName === cur.tagName; });
        if (siblings.length > 1) part += ":nth-of-type(" + (siblings.indexOf(cur) + 1) + ")";
      }
      parts.unshift(part);
      cur = parent;
    }
    return parts.join(" > ");
  }

  function probeOf(el) {
    var tag = (el.tagName || "").toLowerCase();
    var ctrl = labelControl(el);
    var src = ctrl || el;
    var srcTag = (src.tagName || "").toLowerCase();
    return {
      tag: tag,
      inputType: srcTag === "input" ? (src.getAttribute("type") || "text") : null,
      autocomplete: (src.getAttribute && src.getAttribute("autocomplete")) || null,
      role: roleOf(el),
      accessibleName: nameOf(el),
      valueAttr: src.value != null ? String(src.value).slice(0, 200) : null,
      innerText: (tag === "html" || tag === "body")
        ? String(document.title || "").replace(/\s+/g, " ").trim().slice(0, 120)
        : String(el.innerText || "").replace(/\s+/g, " ").trim().slice(0, 200),
      href: el.href || (el.getAttribute && el.getAttribute("href")) || null,
      dataUrl: el.src || (el.getAttribute && el.getAttribute("src")) || null,
      hasDownloadAttr: !!(el.hasAttribute && el.hasAttribute("download")),
      pageHasPasswordInput: !!(document.querySelector && document.querySelector('input[type="password"]')),
      isFileChooser: looksLikeFileChooser(el) || looksLikeFileChooser(src),
      disabled: !!(src.disabled),
      readOnly: !!src.readOnly,
      ariaDisabled: (function () {
        var ad = src.getAttribute && src.getAttribute("aria-disabled");
        return ad != null && String(ad).toLowerCase() !== "false";
      })()
    };
  }

  function classBlob(el) {
    var c = el.className;
    var name = c && typeof c === "object" && c.baseVal != null ? String(c.baseVal) : String(c || "");
    return (name + " " + (el.id || "")).toLowerCase();
  }

  function isNamedIcon(el) {
    return /delete|remove|trash|\bdel\b|shanchu|icon-del|\bedit\b|pencil|modify|bianji|icon-edit/.test(classBlob(el));
  }

  function isClickableIcon(el) {
    var tag = (el.tagName || "").toLowerCase();
    if (tag !== "i" && tag !== "span" && tag !== "em") return false;
    var role = ((el.getAttribute && el.getAttribute("role")) || "").toLowerCase();
    if (role === "button" || role === "link" || role === "menuitem") return true;
    if (el.getAttribute && el.getAttribute("onclick")) return true;
    return isNamedIcon(el) || isGlyphIcon(el);
  }

  function isGlyphIcon(el) {
    var tag = (el.tagName || "").toLowerCase();
    if (tag !== "i" && tag !== "span" && tag !== "em") return false;
    var inner = String(el.textContent || "").replace(/\s+/g, " ").trim();
    if (inner.length > 4) return false;
    if (/[\uE000-\uF8FF]/.test(inner)) return true;
    return tag === "i" && inner.length <= 1;
  }

  function insideHiddenTemplate(el) {
    var p = el.parentElement;
    while (p && p !== document.body && p !== document.documentElement) {
      var st = null;
      try { st = window.getComputedStyle(p); } catch (eH) { break; }
      if (st && (st.display === "none" || st.visibility === "hidden")) return true;
      p = p.parentElement;
    }
    return false;
  }

  function isRowHoverControl(el) {
    if (!el || el.nodeType !== 1 || insideHiddenTemplate(el)) return false;
    var tag = (el.tagName || "").toLowerCase();
    var role = ((el.getAttribute && el.getAttribute("role")) || "").toLowerCase();
    var hoverRole = tag === "a" || tag === "button" || role === "button" || role === "link" || role === "menuitem" || role === "tab";
    if (!hoverRole && !isClickableIcon(el)) return false;
    var inner = String(el.textContent || "").replace(/\s+/g, " ").trim();
    if (inner.length > 16) return false;
    var st = null;
    try { st = window.getComputedStyle(el); } catch (eS) {}
    if (st && (st.display === "none" || st.visibility === "hidden")) return true;
    if (st && parseFloat(st.opacity) === 0) return true;
    try {
      var r = el.getBoundingClientRect();
      return !r || r.width <= 0 || r.height <= 0;
    } catch (eB) { return true; }
  }

  function hoverIconLabel(el) {
    var aria = el.getAttribute && (el.getAttribute("aria-label") || el.getAttribute("title"));
    aria = String(aria || "").replace(/\s+/g, " ").trim();
    if (aria) return aria.slice(0, 40);
    var blob = classBlob(el);
    if (/delete|remove|trash|\bdel\b|shanchu|icon-del|icon-close|\bclose\b|btn-del|del-btn/.test(blob)) return "删除";
    if (/\bedit\b|pencil|modify|bianji|icon-edit/.test(blob)) return "编辑";
    var own = String(el.textContent || "").replace(/\s+/g, " ").trim();
    if (own && own.length <= 16 && !/[\uE000-\uF8FF]/.test(own)) return own;
    var name = "图标";
    var parts = blob.match(/[a-z][a-z0-9-]{2,18}/g) || [];
    var skip = { icon: 1, iconfont: 1, icons: 1, button: 1, btn: 1, fl: 1, fr: 1 };
    var picked = "";
    for (var i = 0; i < parts.length; i++) if (!skip[parts[i]]) picked = parts[i];
    if (picked) name += " " + picked;
    var p = el.parentElement;
    for (var d = 0; p && d < 8; d++) {
      var kids = p.children || [];
      for (var k = 0; k < kids.length; k++) {
        var node = kids[k];
        if (node === el || (node.contains && node.contains(el))) continue;
        var t = String(node.innerText || node.textContent || "").replace(/[\uE000-\uF8FF]/g, "").replace(/\s+/g, " ").trim();
        if (t && t.length >= 2 && t.length <= 20 && name.indexOf(t) < 0) return (name + " " + t).slice(0, 40);
      }
      p = p.parentElement;
    }
    return name.slice(0, 40);
  }

  function visibleHost(el) {
    var p = el && el.parentElement;
    for (var d = 0; p && p !== document.body && p !== document.documentElement && d < 6; d++) {
      var st = null, box = null;
      try { st = window.getComputedStyle(p); } catch (e1) {}
      try { box = p.getBoundingClientRect(); } catch (e2) {}
      if (st && st.display !== "none" && st.visibility !== "hidden" && box && box.width > 1 && box.height > 1) return p;
      p = p.parentElement;
    }
    return null;
  }

  function shouldSkip(el) {
    if (!el || el.nodeType !== 1) return true;
    var tag = (el.tagName || "").toLowerCase();
    if (tag === "script" || tag === "style" || tag === "noscript" || tag === "meta" || tag === "link") return true;
    try {
      var st = window.getComputedStyle(el);
      if (st && (st.display === "none" || st.visibility === "hidden")) {
        if (tag === "input" || tag === "textarea" || tag === "select") return false;
      }
    } catch (e10) {}
    return false;
  }

  function interesting(el) {
    var tag = (el.tagName || "").toLowerCase();
    if (isRowHoverControl(el)) return true;
    if (tag === "a" || tag === "button" || tag === "input" || tag === "textarea" || tag === "select") return true;
    if (tag === "label" || tag === "iframe" || tag === "frame" || tag === "img") return true;
    if (el.isContentEditable) return true;
    var aria = el.getAttribute && el.getAttribute("role");
    if (aria) return true;
    if (/^h[1-6]$/.test(tag) || tag === "li" || tag === "tr") return true;
    return false;
  }

  function valueOf(el) {
    var tag = (el.tagName || "").toLowerCase();
    if (tag === "input" && (el.getAttribute("type") || "").toLowerCase() === "password") return null;
    if (tag === "select") {
      var o = el.options && el.selectedIndex >= 0 ? el.options[el.selectedIndex] : null;
      return o ? String(o.text || o.value || "").replace(/\s+/g, " ").trim() : "";
    }
    if (tag === "input" || tag === "textarea") {
      return el.value != null ? String(el.value).slice(0, 80) : "";
    }
    return null;
  }

  function optionsOf(el) {
    if ((el.tagName || "").toLowerCase() !== "select" || !el.options) return null;
    var out = [];
    for (var i = 0; i < el.options.length && out.length < 8; i++) {
      var tx = String(el.options[i].text || "").replace(/\s+/g, " ").trim();
      if (tx) out.push(tx);
    }
    return out.length ? out.join("|") : null;
  }

  function snapshot(msg) {
    msg = msg || {};
    var nodes = {};
    var counter = 0;
    var lines = [];
    var truncated = false;
    var truncatedReason = null;
    var chromeLeft = 0;
    var pageHasRuns = false;
    var charStop = false;
    var refTrunc = false;
    var focusRef = null;
    var windowApplied = false;

    function nextRef(chrome) {
      if (chrome && chromeLeft > 0) chromeLeft -= 1;
      var limit = (!chrome && pageHasRuns) ? MAX_NODES - Math.max(0, chromeLeft) : MAX_NODES;
      if (counter >= limit) return null;
      counter += 1;
      return "e" + counter;
    }

    function isNoiseRole(role) {
      var r = String(role || "").toLowerCase();
      return r === "div" || r === "generic";
    }

    function isChromeRole(role) {
      var r = String(role || "").toLowerCase();
      return r === "button" || r === "link" || r === "textbox" || r === "checkbox"
        || r === "radio" || r === "combobox" || r === "searchbox";
    }

    function buildNode(el) {
      if (shouldSkip(el)) return null;
      var children = [];
      var raw = el.children ? Array.prototype.slice.call(el.children) : [];
      for (var i = 0; i < raw.length; i++) {
        var child = buildNode(raw[i]);
        if (child) children.push(child);
      }
      return { el: el, role: roleOf(el), children: children, print: interesting(el) };
    }

    function sigOf(node, stack) {
      if (node._sig) return node._sig;
      if (stack.indexOf(node) >= 0) {
        node._sig = { signature: (node.role || "generic") + "()", parts: [], desc: 1 };
        return node._sig;
      }
      stack.push(node);
      var desc = 1;
      var rawParts = [];
      for (var i = 0; i < node.children.length; i++) {
        var child = node.children[i];
        var cs = sigOf(child, stack);
        desc += cs.desc;
        if (isNoiseRole(child.role)) {
          for (var p = 0; p < cs.parts.length; p++) rawParts.push(cs.parts[p]);
        } else rawParts.push(cs.signature);
      }
      stack.pop();
      var parts = collapseParts(rawParts);
      node._sig = { signature: (node.role || "generic") + "(" + parts.join(",") + ")", parts: parts, desc: desc };
      return node._sig;
    }

    function stripTimes(value) {
      if (value && value.charAt(value.length - 1) === "×") return value.slice(0, -1);
      return value || "";
    }

    function collapseParts(parts) {
      var rle = [];
      for (var i = 0; i < parts.length; i++) {
        var part = parts[i];
        if (!rle.length) { rle.push(part); continue; }
        var last = rle[rle.length - 1];
        if (stripTimes(last) === stripTimes(part)) rle[rle.length - 1] = stripTimes(part) + "×";
        else rle.push(part);
      }
      return collapsePeriodic(rle);
    }

    function collapsePeriodic(seq) {
      var result = [];
      var i = 0;
      while (i < seq.length) {
        var bestP = 0, bestK = 0;
        var maxP = Math.floor((seq.length - i) / 2);
        for (var p = 1; p <= maxP; p++) {
          var k = 1;
          while (i + (k + 1) * p <= seq.length && blockEquals(seq, i, i + k * p, p)) k += 1;
          if (k >= 2 && k * p > bestK * bestP) { bestP = p; bestK = k; }
        }
        if (bestK >= 2) {
          var block = [];
          for (var j = 0; j < bestP; j++) block.push(seq[i + j]);
          result.push("(" + block.join(",") + ")×");
          i += bestK * bestP;
        } else { result.push(seq[i]); i += 1; }
      }
      return result;
    }

    function blockEquals(seq, a, b, len) {
      for (var j = 0; j < len; j++) if (seq[a + j] !== seq[b + j]) return false;
      return true;
    }

    function markUnder(node) {
      node._under = true;
      for (var i = 0; i < node.children.length; i++) markUnder(node.children[i]);
    }

    function analyze(node) {
      sigOf(node, []);
      var runs = [];
      var ordinal = 0;
      var i = 0;
      var kids = node.children;
      while (i < kids.length) {
        var sig = kids[i]._sig;
        if (!sig || sig.desc < 12) { i += 1; continue; }
        var j = i + 1;
        while (j < kids.length && kids[j]._sig && kids[j]._sig.desc >= 12 && kids[j]._sig.signature === sig.signature) j += 1;
        if (j - i >= 3) {
          var members = kids.slice(i, j);
          for (var m = 0; m < members.length; m++) markUnder(members[m]);
          runs.push({ start: i, count: j - i, ordinal: ordinal, signature: sig.signature, members: members });
          ordinal += 1;
          i = j;
        } else i += 1;
      }
      node._runs = runs;
      for (var c = 0; c < kids.length; c++) analyze(kids[c]);
    }

    function countChrome(node) {
      var n = (isChromeRole(node.role) && node.print && !node._under) ? 1 : 0;
      n += (node._runs || []).length;
      for (var i = 0; i < node.children.length; i++) n += countChrome(node.children[i]);
      return n;
    }

    function hasRuns(node) {
      if (node._runs && node._runs.length) return true;
      for (var i = 0; i < node.children.length; i++) if (hasRuns(node.children[i])) return true;
      return false;
    }

    function quoteName(name) {
      return name ? ' "' + String(name).replace(/"/g, "'") + '"' : "";
    }

    function appendLine(depth, role, name, ref, el, extra) {
      var indent = "";
      for (var d = 0; d < Math.min(depth, 12); d++) indent += "  ";
      var line = indent + "- " + role + quoteName(name);
      if (ref) line += " [ref=" + ref + "]";
      if (extra) line += extra;
      lines.push(line);
      if (lines.join("\n").length > MAX_CHARS) charStop = true;
      if (ref && el) {
        nodes[ref] = { role: role, name: name, cssPath: cssPath(el), probe: probeOf(el), tag: (el.tagName || "").toLowerCase() };
      }
    }

    function emitRepeat(parent, run, depth) {
      var start = 1;
      var end = Math.min(300, run.count);
      var ranged = msg.repeat_parent_css && msg.range_start != null && msg.range_end != null;
      var matched = ranged && cssPath(parent.el) === msg.repeat_parent_css
        && Number(msg.repeat_ordinal) === run.ordinal && msg.repeat_signature === run.signature;
      if (matched) {
        windowApplied = true;
        if (Number(msg.range_start) > run.count) { windowApplied = false; return; }
        start = Number(msg.range_start);
        end = Math.min(Number(msg.range_end), run.count);
        if (end - start + 1 > 300) end = start + 300 - 1;
      }
      var label = "共 " + run.count + " 项，当前 " + start + "-" + end;
      var headerRef = nextRef(true);
      if (headerRef) {
        nodes[headerRef] = {
          role: "repeat", name: label, isRepeat: true,
          repeatParentCss: cssPath(parent.el), repeatSignature: run.signature,
          repeatOrdinal: run.ordinal, repeatCount: run.count
        };
        if (matched) focusRef = headerRef;
      }
      var indent = "";
      for (var d = 0; d < Math.min(depth, 12); d++) indent += "  ";
      lines.push(indent + "- repeat" + quoteName(label) + (headerRef ? " [ref=" + headerRef + "]" : ""));
      var prevChar = charStop;
      emit(run.members[start - 1], depth + 1);
      var childChar = charStop;
      charStop = prevChar;
      var printedEnd = start;
      if (!childChar) {
        for (var n = start + 1; n <= end; n++) {
          if (charStop || lines.join("\n").length > MAX_CHARS) { charStop = true; break; }
          var item = run.members[n - 1];
          var itemName = nameOf(item.el) || "";
          if (itemName.length > 40) itemName = itemName.slice(0, 40) + "…";
          var oneRef = nextRef(false);
          if (!oneRef) refTrunc = true;
          appendLine(depth + 1, item.role || "generic", itemName, oneRef, item.el, "");
          printedEnd = n;
          if (charStop) break;
        }
      }
      if (printedEnd < end) {
        var actual = "共 " + run.count + " 项，当前 " + start + "-" + printedEnd;
        for (var li = 0; li < lines.length; li++) lines[li] = lines[li].split(label).join(actual);
        if (headerRef && nodes[headerRef]) nodes[headerRef].name = actual;
      }
      if (childChar) charStop = true;
    }

    function emit(node, depth) {
      if (!node) return;
      var chrome = isChromeRole(node.role) && node.print && !node._under;
      if (charStop && !chrome) { emitChildren(node, depth); return; }
      var print = node.print || depth === 0;
      if (print) {
        var tag = (node.el.tagName || "").toLowerCase();
        var ref = nextRef(chrome);
        if (!ref) refTrunc = true;
        var extra = "";
        var val = valueOf(node.el);
        if (val != null && val !== "") extra += ' value="' + String(val).replace(/"/g, "'") + '"';
        var opts = optionsOf(node.el);
        if (opts) extra += " options=" + opts;
        if (tag === "input" && (node.el.getAttribute("type") || "") === "password") extra += " (password)";
        if (isRowHoverControl(node.el)) extra += " (隐藏)";
        appendLine(depth, node.role, nameOf(node.el), ref, node.el, extra);
        if (tag === "iframe" || tag === "frame") return;
      }
      emitChildren(node, print ? depth + 1 : depth);
    }

    function emitChildren(node, depth) {
      var runs = node._runs || [];
      var kids = node.children;
      for (var i = 0; i < kids.length; i++) {
        var run = null;
        for (var r = 0; r < runs.length; r++) if (runs[r].start === i) run = runs[r];
        if (run) {
          var prev = charStop;
          emitRepeat(node, run, depth);
          charStop = prev;
          i += run.count - 1;
          continue;
        }
        emit(kids[i], depth);
      }
    }

    var rootEl = document.body || document.documentElement;
    if (String(msg.mode || "") === "detail" && msg.css_path) {
      rootEl = document.querySelector(msg.css_path);
      if (!rootEl) return { ok: false, error: "ref 无效或已过期，请重新 F_browser_snapshot" };
    }
    var built = buildNode(rootEl);
    if (built) {
      analyze(built);
      pageHasRuns = hasRuns(built);
      chromeLeft = countChrome(built);
      emit(built, 0);
    }
    if (msg.range_start != null && msg.range_end != null && !windowApplied) {
      return { ok: false, error: "ref 无效或已过期，请重新 F_browser_snapshot" };
    }
    if (charStop) {
      truncated = true;
      truncatedReason = truncatedReason || "char_limit";
    } else if (refTrunc) {
      truncated = true;
      truncatedReason = truncatedReason || "node_limit";
    }
    return {
      ok: true,
      snapshot: lines.join("\n"),
      nodes: nodes,
      truncated: truncated,
      truncated_reason: truncatedReason,
      focus_ref: focusRef,
      url: document.location.href,
      title: document.title
    };
  }

  function resolveEl(cssPathHint) {
    var path = cssPathHint || "";
    if (!path) throw new Error("未知 ref；请先 snapshot");
    var el = document.querySelector(path);
    if (!el) throw new Error("ref 对应元素已失效。请重新 snapshot 后再操作");
    return el;
  }

  function findSelect(el) {
    if (!el) return null;
    if ((el.tagName || "").toUpperCase() === "SELECT") return el;
    try {
      if (el.querySelector) {
        var inner = el.querySelector("select");
        if (inner) return inner;
      }
      if (el.previousElementSibling && (el.previousElementSibling.tagName || "").toUpperCase() === "SELECT") {
        return el.previousElementSibling;
      }
      if (el.nextElementSibling && (el.nextElementSibling.tagName || "").toUpperCase() === "SELECT") {
        return el.nextElementSibling;
      }
      var g = el.closest && el.closest(".input-group");
      if (g && g.querySelector) return g.querySelector("select");
    } catch (e11) {}
    return null;
  }

  function interact(msg) {
    msg = msg || {};
    var action = String(msg.action || "").toLowerCase();
    if (action === "scroll" && !msg.css_path) {
      var dir = String(msg.direction || "down").toLowerCase();
      var dy = dir === "up" ? -400 : dir === "down" ? 400 : 0;
      var dx = dir === "left" ? -400 : dir === "right" ? 400 : 0;
      window.scrollBy(dx, dy);
      return { ok: true, message: "scrolled " + dir, url: document.location.href, title: document.title };
    }

    var el = null;
    var entry = { name: msg.name || "", probe: msg.probe || {} };
    if (msg.css_path) {
      el = resolveEl(msg.css_path);
    } else if (action === "press") {
      el = document.activeElement || document.body;
    } else {
      throw new Error("未知 ref；请先 snapshot");
    }

    if (looksLikeFileChooser(el)) {
      throw new Error("本批不支持文件选择。请用户自己在窗里选文件。");
    }

    var typeTarget = el;
    if (action === "type") {
      typeTarget = findTypeable(el, entry);
      if (!typeTarget) throw new Error("目标不可输入");
      if (looksLikePassword(typeTarget)) {
        throw new Error("已拒绝向密码框输入；请用户在看得见的窗里自己填写密码");
      }
      if (looksLikeNotEditable(typeTarget) && !isTemporalInput(typeTarget)) {
        throw new Error("目标不可编辑");
      }
    }

    try {
      (action === "type" ? typeTarget : el).scrollIntoView({ block: "center", inline: "nearest" });
    } catch (e12) {}

    if (action === "click") {
      if ((el.tagName || "").toLowerCase() === "select") {
        throw new Error("这是原生下拉，请用 action=select，option 填选项原文或 value。");
      }
      var box0 = { width: 0, height: 0 };
      try { box0 = el.getBoundingClientRect(); } catch (eBox) {}
      if (!box0 || box0.width <= 0 || box0.height <= 0) {
        var host = visibleHost(el);
        if (host) {
          var hb = host.getBoundingClientRect();
          var hx = hb.left + hb.width / 2;
          var hy = hb.top + hb.height / 2;
          var hopts = { bubbles: true, cancelable: true, view: window, clientX: hx, clientY: hy };
          try { host.dispatchEvent(new MouseEvent("mouseover", hopts)); } catch (eH1) {}
          try { host.dispatchEvent(new MouseEvent("mousemove", hopts)); } catch (eH2) {}
        }
      }
      try { el.click(); } catch (e13) {}
      return { ok: true, message: "clicked", url: document.location.href, title: document.title };
    }

    if (action === "hover") {
      var box = el.getBoundingClientRect();
      var opts = { bubbles: true, cancelable: true, view: window, clientX: box.left + box.width / 2, clientY: box.top + box.height / 2 };
      try { el.dispatchEvent(new MouseEvent("mouseover", opts)); } catch (e14) {}
      try { el.dispatchEvent(new MouseEvent("mouseenter", opts)); } catch (e15) {}
      return { ok: true, message: "hovered", url: document.location.href, title: document.title };
    }

    if (action === "type") {
      var text = msg.text == null ? "" : String(msg.text);
      if (isTemporalInput(typeTarget)) {
        var filledDate = fillTemporal(typeTarget, text);
        return { ok: true, message: "typed value=" + JSON.stringify(filledDate.value), url: document.location.href, title: document.title };
      }
      var tag = (typeTarget.tagName || "").toLowerCase();
      typeTarget.focus();
      try { typeTarget.click(); } catch (e16) {}
      if (tag === "input" || tag === "textarea") {
        var filled = fillTextControl(typeTarget, text);
        try {
          typeTarget.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
          typeTarget.dispatchEvent(new KeyboardEvent("keyup", { key: "Escape", bubbles: true }));
        } catch (e17) {}
        return { ok: true, message: "typed value=" + JSON.stringify(filled.value), url: document.location.href, title: document.title };
      }
      if (typeTarget.isContentEditable) {
        typeTarget.focus();
        try { document.execCommand("selectAll", false, null); document.execCommand("delete", false, null); } catch (e18) {}
        var ok = false;
        try { ok = document.execCommand("insertText", false, text); } catch (e19) {}
        if (!ok) typeTarget.textContent = text;
        try {
          typeTarget.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: true, inputType: "insertText", data: text }));
        } catch (e20) {
          typeTarget.dispatchEvent(new Event("input", { bubbles: true }));
        }
        return { ok: true, message: "typed contenteditable", url: document.location.href, title: document.title };
      }
      throw new Error("目标不可输入");
    }

    if (action === "press") {
      var key = msg.key || "Enter";
      el.focus();
      el.dispatchEvent(new KeyboardEvent("keydown", { key: key, bubbles: true }));
      el.dispatchEvent(new KeyboardEvent("keyup", { key: key, bubbles: true }));
      if (key === "Enter" && (el.tagName || "").toLowerCase() === "form" && el.requestSubmit) {
        el.requestSubmit();
      }
      return { ok: true, message: "pressed " + key, url: document.location.href, title: document.title };
    }

    if (action === "select") {
      var sel = findSelect(el);
      if (!sel) {
        throw new Error("select 未落到原生下拉。请重新 snapshot，对带 options= 的那一行再 select。");
      }
      if (looksLikeNotEditable(sel)) throw new Error("目标不可编辑");
      var option = String(msg.option || "");
      var matched = false;
      var preview = [];
      for (var i = 0; i < sel.options.length; i++) {
        var opt = sel.options[i];
        var tx = (opt.textContent || "").trim();
        if (preview.length < 8 && tx) preview.push(tx);
        if (opt.value === option || tx === option || tx.indexOf(option) >= 0 || String(opt.value).indexOf(option) >= 0) {
          sel.selectedIndex = i;
          matched = true;
          break;
        }
      }
      if (!matched) {
        throw new Error("未找到选项: " + option + (preview.length ? "。可选：" + preview.join("、") : ""));
      }
      sel.dispatchEvent(new Event("input", { bubbles: true }));
      sel.dispatchEvent(new Event("change", { bubbles: true }));
      return { ok: true, message: "selected", url: document.location.href, title: document.title };
    }

    if (action === "scroll") {
      try { el.scrollIntoView({ block: "center", inline: "nearest" }); } catch (e21) {}
      return { ok: true, message: "scrolled to ref", url: document.location.href, title: document.title };
    }

    throw new Error("不支持的 action: " + action);
  }

  root.__yiwriteLegacyDom = { snapshot: snapshot, interact: interact };
})(typeof globalThis !== "undefined" ? globalThis : this);
