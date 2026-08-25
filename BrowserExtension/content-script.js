/**
 * 易写 — 页内 snapshot / interact / download 辅助
 */
(function () {
  if (window.__yiwriteContentLoaded) {
    return;
  }
  window.__yiwriteContentLoaded = true;

  const MAX_CHARS = 36000;
  const MAX_NODES = 300;

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    handle(msg).then(sendResponse).catch((e) => {
      sendResponse({ ok: false, error: String(e && e.message ? e.message : e) });
    });
    return true;
  });

  async function handle(msg) {
    if (!msg || !msg.type) {
      return { ok: false, error: "bad message" };
    }
    switch (msg.type) {
      case "ping":
        return { ok: true };
      case "snapshot":
        return doSnapshot(msg);
      case "interact":
        return doInteract(msg);
      case "download_prep":
        return doDownloadPrep(msg.ref, msg.css_path);
      case "download_click":
        return doDownloadClick(msg.ref, msg.css_path);
      default:
        return { ok: false, error: "unknown: " + msg.type };
    }
  }

  function looksLikeNotEditable(node) {
    if (!node) return false;
    try {
      if (node.disabled) return true;
      if (node.matches && node.matches(":disabled")) return true;
    } catch (_) {}
    try {
      if (node.readOnly) return true;
    } catch (_) {}
    try {
      const ad = node.getAttribute && node.getAttribute("aria-disabled");
      if (ad != null && String(ad).toLowerCase() !== "false") return true;
    } catch (_) {}
    return false;
  }

  function looksLikeFileChooser(node) {
    if (!node) return false;
    try {
      const tag = (node.tagName || "").toLowerCase();
      const typ = String((node.getAttribute && node.getAttribute("type")) || node.type || "").toLowerCase();
      if (tag === "input" && typ === "file") return true;
      if (node.querySelector && node.querySelector('input[type="file"]')) return true;
      const lab = node.closest && node.closest("label");
      if (lab) {
        if (lab.querySelector && lab.querySelector('input[type="file"]')) return true;
        const id = lab.htmlFor || lab.getAttribute("for");
        if (id) {
          const t = document.getElementById(id);
          if (t && String(t.type || "").toLowerCase() === "file") return true;
        }
      }
    } catch (_) {}
    return false;
  }

  function doSnapshot(msg) {
    const nodes = {};
    let counter = 0;
    const lines = [];
    let truncated = false;
    let truncatedReason = null;

    function nextRef() {
      counter += 1;
      return "e" + counter;
    }

    function roleOf(el) {
      const r = el.getAttribute("role");
      if (r) return r;
      const tag = (el.tagName || "").toLowerCase();
      if (tag === "a") return "link";
      if (tag === "button") return "button";
      if (tag === "input") {
        const t = (el.getAttribute("type") || "text").toLowerCase();
        if (t === "password") return "textbox";
        if (t === "checkbox") return "checkbox";
        if (t === "radio") return "radio";
        if (t === "submit" || t === "button") return "button";
        return "textbox";
      }
      if (tag === "textarea") return "textbox";
      if (tag === "select") return "combobox";
      if (el.isContentEditable) return "textbox";
      if (tag === "img") return "image";
      if (tag === "area") return "link";
      if (tag === "iframe" || tag === "frame") return tag;
      if (/^h[1-6]$/.test(tag)) return "heading";
      if (tag === "li") return "listitem";
      return tag || "generic";
    }

    function srcTail(src) {
      if (!src) return "";
      try {
        const u = new URL(src, document.baseURI);
        const last = (u.pathname || "").split("/").filter(Boolean).pop();
        return last || u.hostname || "";
      } catch (_) {
        const s = String(src);
        const i = s.lastIndexOf("/");
        return i >= 0 && i < s.length - 1 ? s.slice(i + 1) : s.slice(-40);
      }
    }

    function nameOf(el) {
      const tag = (el.tagName || "").toLowerCase();
      const t = (s) => String(s || "").replace(/\s+/g, " ").trim();
      if (tag === "iframe" || tag === "frame") {
        const aria = el.getAttribute("aria-label");
        if (t(aria)) return t(aria);
        if (t(el.title)) return t(el.title);
        const src = el.getAttribute("src") || el.src || "";
        return srcTail(src) || "iframe";
      }
      const aria = el.getAttribute("aria-label");
      if (t(aria)) return t(aria).slice(0, 120);
      if (t(el.title)) return t(el.title).slice(0, 120);
      if (t(el.alt)) return t(el.alt).slice(0, 120);
      try {
        const img = el.querySelector && el.querySelector("img[alt]");
        if (img && t(img.alt)) return t(img.alt).slice(0, 120);
        const st = el.querySelector && el.querySelector("svg title");
        if (st && t(st.textContent)) return t(st.textContent).slice(0, 120);
        const area = el.querySelector && el.querySelector("area[alt]");
        if (area && t(area.alt)) return t(area.alt).slice(0, 120);
      } catch (_) {}
      if (el.placeholder) return String(el.placeholder).trim().slice(0, 120);
      const label = el.labels && el.labels[0];
      if (label) return t(label.innerText).slice(0, 120);
      return t(el.innerText || el.textContent).slice(0, 120);
    }

    function probeOf(el) {
      const tag = (el.tagName || "").toLowerCase();
      return {
        tag,
        inputType: tag === "input" ? (el.getAttribute("type") || "text") : null,
        autocomplete: el.getAttribute("autocomplete") || null,
        role: roleOf(el),
        accessibleName: nameOf(el),
        valueAttr: el.value != null ? String(el.value).slice(0, 200) : null,
        innerText: (el.innerText || "").replace(/\s+/g, " ").trim().slice(0, 200),
        href: el.href || el.getAttribute("href") || null,
        dataUrl: el.src || el.getAttribute("src") || null,
        hasDownloadAttr: el.hasAttribute("download"),
        pageHasPasswordInput: !!document.querySelector('input[type="password"]'),
        isFileChooser: looksLikeFileChooser(el),
        disabled: (() => {
          try {
            if (el.disabled) return true;
            if (el.matches && el.matches(":disabled")) return true;
          } catch (_) {}
          return false;
        })(),
        readOnly: !!el.readOnly,
        ariaDisabled: (() => {
          const ad = el.getAttribute && el.getAttribute("aria-disabled");
          return ad != null && String(ad).toLowerCase() !== "false";
        })()
      };
    }

    function ownText(el) {
      let t = "";
      const nodes = el.childNodes || [];
      for (let i = 0; i < nodes.length; i++) {
        if (nodes[i].nodeType === 3) t += nodes[i].textContent || "";
      }
      return t.replace(/\s+/g, " ").trim();
    }

    function isDetailRevealName(s) {
      if (!s) return false;
      const n = String(s).replace(/\s+/g, " ").trim();
      if (n.indexOf("查看详情") >= 0) return true;
      if (n.indexOf("查看更多") >= 0) return true;
      return n === "详情";
    }

    function hasDetailLabel(el) {
      const own = ownText(el);
      if (isDetailRevealName(own)) return true;
      const inner = (el.innerText || "").replace(/\s+/g, " ").trim();
      return inner.length <= 12 && isDetailRevealName(inner);
    }

    function isHoverRevealControl(el) {
      if (!el || el.nodeType !== 1) return false;
      if (hasDetailLabel(el)) return true;
      const tag = (el.tagName || "").toLowerCase();
      if (tag === "a" || tag === "button") return true;
      const role = (el.getAttribute("role") || "").toLowerCase();
      return role === "button" || role === "link" || role === "menuitem"
        || role === "menuitemcheckbox" || role === "menuitemradio" || role === "tab";
    }

    function classNameOf(el) {
      const c = el.className;
      if (c && typeof c === "object" && c.baseVal != null) return String(c.baseVal);
      return String(c || "");
    }

    function isImgtextBtn(el) {
      return /\bimgtextbtn\b/.test(classNameOf(el));
    }

    function isImgtextCard(el) {
      return /\bimgtext\b/.test(classNameOf(el)) && !!(el.querySelector && el.querySelector("img"));
    }

    function isPointerCard(el) {
      if (isImgtextCard(el)) return true;
      try {
        const st = window.getComputedStyle(el);
        if (!st || st.cursor !== "pointer" || st.display === "none") return false;
        if (!el.querySelector || !el.querySelector("img")) return false;
        return el.offsetWidth >= 100 && el.offsetHeight >= 60;
      } catch (_) {
        return false;
      }
    }

    function closestPointerCard(el) {
      let p = el;
      while (p && p.nodeType === 1 && p !== document.body) {
        if (isPointerCard(p)) return p;
        p = p.parentElement;
      }
      return null;
    }

    function isPointerMenu(el) {
      try {
        const st = window.getComputedStyle(el);
        if (!st || st.cursor !== "pointer" || st.display === "none") return false;
        const own = ownText(el);
        if (own.length < 2 || own.length > 16) return false;
        const rect = el.getBoundingClientRect();
        return rect.left >= 0 && rect.left < 240 && rect.width > 0 && rect.height > 0;
      } catch (_) {
        return false;
      }
    }

    function isHoverHiddenStyle(el) {
      try {
        const st = window.getComputedStyle(el);
        if (!st) return false;
        if (st.visibility === "hidden") return true;
        const op = parseFloat(st.opacity);
        return !isNaN(op) && op === 0;
      } catch (_) {
        return false;
      }
    }

    function shouldSkip(el) {
      if (!el || el.nodeType !== 1) return true;
      const tag = (el.tagName || "").toLowerCase();
      if (tag === "iframe" || tag === "frame") return false;
      if (tag === "map" || tag === "area") return false;
      if (tag === "script" || tag === "style" || tag === "noscript" || tag === "svg") return true;
      const st = window.getComputedStyle(el);
      if (st && st.display === "none") {
        if (hasDetailLabel(el) || isImgtextBtn(el) || isDetailRevealName(ownText(el))) return false;
        if (el.querySelector && el.querySelector(".imgtextbtn, iframe, frame")) return false;
        return true;
      }
      if (st && st.visibility === "hidden") {
        if (isHoverRevealControl(el) || hasDetailLabel(el)) return false;
        const inner = (el.innerText || "").replace(/\s+/g, " ").trim();
        if (inner.length <= 16 && isDetailRevealName(inner)) return false;
        if (el.querySelector && el.querySelector("a, button, [role='button'], [role='link'], [role='menuitem'], iframe, frame")) {
          return false;
        }
        return true;
      }
      return false;
    }

    function interesting(el) {
      const tag = (el.tagName || "").toLowerCase();
      if (hasDetailLabel(el) || isImgtextBtn(el) || isPointerCard(el) || isPointerMenu(el)) return true;
      if (tag === "img" && !closestPointerCard(el.parentElement)) return true;
      if (["a", "button", "input", "textarea", "select", "label", "iframe", "frame"].indexOf(tag) >= 0) {
        return true;
      }
      if (tag === "area" && (el.getAttribute("alt") || el.title || el.getAttribute("aria-label"))) {
        return true;
      }
      if (el.isContentEditable) return true;
      if (el.getAttribute("role")) return true;
      if (/^h[1-6]$/.test(tag)) return true;
      return false;
    }

    function walk(el, depth) {
      if (truncated || counter >= MAX_NODES) {
        truncated = true;
        truncatedReason = "node_limit";
        return;
      }
      if (shouldSkip(el)) return;

      if (interesting(el) || depth === 0) {
        const ref = nextRef();
        const role = roleOf(el);
        const name = nameOf(el);
        const indent = "  ".repeat(Math.min(depth, 12));
        let line = indent + "- " + role;
        if (name) line += ' "' + name.replace(/"/g, "'") + '"';
        line += " [ref=" + ref + "]";
        if (el.tagName && el.tagName.toLowerCase() === "input" && (el.getAttribute("type") || "") === "password") {
          line += " (password)";
        }
        if ((isHoverRevealControl(el) || hasDetailLabel(el) || isImgtextBtn(el))
          && (isHoverHiddenStyle(el) || (window.getComputedStyle(el).display === "none"))) {
          line += " (隐藏)";
        }
        lines.push(line);
        const tag = (el.tagName || "").toLowerCase();
        const rec = {
          role,
          name,
          cssPath: cssPath(el),
          probe: probeOf(el),
          tag
        };
        if (tag === "iframe" || tag === "frame") {
          rec.src = el.getAttribute("src") || el.src || "";
          rec.iframeIndex = iframeIndexOf(el);
        }
        nodes[ref] = rec;

        if (lines.join("\n").length > MAX_CHARS) {
          truncated = true;
          truncatedReason = "char_limit";
          return;
        }

        if (tag === "iframe" || tag === "frame") {
          return;
        }
      }

      const children = el.children ? Array.from(el.children) : [];
      for (let i = 0; i < children.length; i++) {
        walk(children[i], depth + 1);
        if (truncated) return;
      }
    }

    function iframeIndexOf(el) {
      const all = Array.from(document.querySelectorAll("iframe, frame"));
      return all.indexOf(el);
    }

    function collectMissedIframes() {
      const seen = new Set();
      Object.keys(nodes).forEach((k) => {
        if (nodes[k] && nodes[k].cssPath) seen.add(nodes[k].cssPath);
      });
      Array.from(document.querySelectorAll("iframe, frame")).forEach((el) => {
        if (truncated || counter >= MAX_NODES) {
          truncated = true;
          truncatedReason = truncatedReason || "node_limit";
          return;
        }
        const path = cssPath(el);
        if (seen.has(path)) return;
        const ref = nextRef();
        const role = roleOf(el);
        const name = nameOf(el);
        lines.push("- " + role + (name ? ' "' + name.replace(/"/g, "'") + '"' : "") + " [ref=" + ref + "]");
        nodes[ref] = {
          role,
          name,
          cssPath: path,
          probe: probeOf(el),
          tag: (el.tagName || "").toLowerCase(),
          src: el.getAttribute("src") || el.src || "",
          iframeIndex: iframeIndexOf(el)
        };
        seen.add(path);
      });
    }

    const root = document.body || document.documentElement;
    walk(root, 0);
    collectMissedIframes();

    // DOM supplement
    let domSupplement = null;
    const spec = (msg.dom_supplement || "").trim();
    if (spec && spec.toLowerCase() !== "off") {
      const kinds = spec.split(/[,|;、]/).map((s) => s.trim().toLowerCase()).filter(Boolean);
      const extras = [];
      kinds.forEach((k) => {
        let list = [];
        if (k === "input") list = Array.from(document.querySelectorAll("input"));
        else if (k === "textarea") list = Array.from(document.querySelectorAll("textarea"));
        else if (k === "select") list = Array.from(document.querySelectorAll("select"));
        else if (k === "contenteditable") list = Array.from(document.querySelectorAll("[contenteditable='true'], [contenteditable='']"));
        list.forEach((el) => {
          if (shouldSkip(el)) return;
          const ref = nextRef();
          const role = roleOf(el);
          const name = nameOf(el);
          const ph = el.placeholder ? " placeholder=" + el.placeholder : "";
          extras.push('- ' + role + (name ? ' "' + name.replace(/"/g, "'") + '"' : "") + " [ref=" + ref + "]" + ph);
          nodes[ref] = { role, name, cssPath: cssPath(el), probe: probeOf(el) };
        });
      });
      if (extras.length) {
        lines.push("", "# dom_supplement", ...extras);
        domSupplement = spec;
      }
    }

    window.__yiwriteNodes = nodes;
    return {
      ok: true,
      snapshot: lines.join("\n"),
      nodes,
      truncated,
      truncated_reason: truncatedReason,
      dom_supplement: domSupplement
    };
  }

  function cssPath(el) {
    if (!el || !el.tagName) return "";
    if (el.id) return "#" + CSS.escape(el.id);
    const parts = [];
    let cur = el;
    while (cur && cur.nodeType === 1 && parts.length < 6) {
      let part = cur.tagName.toLowerCase();
      if (cur.id) {
        parts.unshift("#" + CSS.escape(cur.id));
        break;
      }
      const parent = cur.parentElement;
      if (parent) {
        const siblings = Array.from(parent.children).filter((c) => c.tagName === cur.tagName);
        if (siblings.length > 1) {
          part += ":nth-of-type(" + (siblings.indexOf(cur) + 1) + ")";
        }
      }
      parts.unshift(part);
      cur = parent;
    }
    return parts.join(" > ");
  }

  function resolveRef(ref, cssPathHint) {
    const nodes = window.__yiwriteNodes || {};
    const entry = nodes[ref] || {};
    const path = entry.cssPath || cssPathHint;
    if (!path) {
      throw new Error("未知 ref: " + ref + "；请重新 snapshot");
    }
    const el = document.querySelector(path);
    if (!el) {
      throw new Error("ref 对应元素已失效: " + ref);
    }
    return { el, entry: entry.cssPath ? entry : { cssPath: path, probe: {} } };
  }

  function isForceClickTarget(el) {
    if (!el) return false;
    const cls = el.className && typeof el.className === "object" && el.className.baseVal != null
      ? String(el.className.baseVal)
      : String(el.className || "");
    if (/\bimgtextbtn\b/.test(cls)) return true;
    function revealName(s) {
      if (!s) return false;
      const n = String(s).replace(/\s+/g, " ").trim();
      if (n.indexOf("查看详情") >= 0) return true;
      if (n.indexOf("查看更多") >= 0) return true;
      return n === "详情";
    }
    let own = "";
    const nodes = el.childNodes || [];
    for (let i = 0; i < nodes.length; i++) {
      if (nodes[i].nodeType === 3) own += nodes[i].textContent || "";
    }
    own = own.replace(/\s+/g, " ").trim();
    if (revealName(own)) return true;
    const inner = (el.innerText || "").replace(/\s+/g, " ").trim();
    return inner.length <= 12 && revealName(inner);
  }

  function parentComposed(el) {
    if (!el) return null;
    if (el.assignedSlot) return el.assignedSlot;
    if (el.parentElement) return el.parentElement;
    const root = el.getRootNode && el.getRootNode();
    if (root && root.host) return root.host;
    return null;
  }

  function isOnTarget(hit, t) {
    let n = hit;
    while (n) {
      if (n === t) return true;
      n = parentComposed(n);
    }
    try {
      return !!(t && t.contains && hit && t.contains(hit));
    } catch (_) {
      return false;
    }
  }

  function previewHit(hit) {
    const tag = ((hit && hit.tagName) || "div").toLowerCase();
    let text = "";
    try {
      text = (hit.getAttribute && (hit.getAttribute("aria-label") || hit.getAttribute("title"))) || "";
      if (!text) text = String(hit.innerText || hit.textContent || "").replace(/\s+/g, " ").trim();
    } catch (_) {}
    if (text.length > 40) text = text.slice(0, 40);
    return text ? (tag + " " + text) : tag;
  }

  function hitTestError(t) {
    if (isForceClickTarget(t)) return null;
    let r = { width: 0, height: 0, left: 0, top: 0 };
    try { r = t.getBoundingClientRect(); } catch (_) {}
    if (r.width <= 0 || r.height <= 0) {
      return "点击被挡住：目标没有可点区域。";
    }
    const x = r.left + r.width / 2;
    const y = r.top + r.height / 2;
    let hit = null;
    try { hit = document.elementFromPoint(x, y); } catch (_) {}
    if (!hit) {
      return "点击被挡住：目标不在当前视口内，请先 scroll 再点。";
    }
    if (isOnTarget(hit, t)) return null;
    return "点击被挡住：挡在上面的是「" + previewHit(hit) + "」。请先操作该蒙层（如接受 Cookie / 关闭），再点原来的目标。";
  }

  function isOverflowScrollable(el) {
    if (!el || el === document.body || el === document.documentElement) return false;
    let st;
    try { st = window.getComputedStyle(el); } catch (_) { return false; }
    if (!st) return false;
    const oy = st.overflowY, ox = st.overflowX, o = st.overflow;
    const y = (oy === "auto" || oy === "scroll" || o === "auto" || o === "scroll")
      && el.scrollHeight > el.clientHeight + 1;
    const x = (ox === "auto" || ox === "scroll" || o === "auto" || o === "scroll")
      && el.scrollWidth > el.clientWidth + 1;
    return y || x;
  }

  function nearestOverflowBox(el) {
    let p = el && el.parentElement;
    while (p && p !== document.body && p !== document.documentElement) {
      if (isOverflowScrollable(p)) return p;
      p = p.parentElement;
    }
    return null;
  }

  function scrollRefIntoView(el) {
    if (!el || !el.isConnected) {
      return { ok: false, error: "滚动目标已不在页面上" };
    }
    try { el.scrollIntoView({ block: "center", inline: "nearest" }); } catch (_) {}
    const box = nearestOverflowBox(el);
    if (box) {
      const er = el.getBoundingClientRect();
      const br = box.getBoundingClientRect();
      box.scrollTop += (er.top + er.height / 2) - (br.top + br.height / 2);
      box.scrollLeft += (er.left + er.width / 2) - (br.left + br.width / 2);
    }
    return { ok: true };
  }

  function clickWithHitTest(el) {
    const err = hitTestError(el);
    if (err) return { ok: false, error: err };
    try { el.focus && el.focus(); } catch (_) {}
    if (typeof el.click === "function") el.click();
    else el.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));
    return { ok: true };
  }

  function resolveClickTarget(el) {
    if (!el) return el;
    const tag = (el.tagName || "").toLowerCase();
    const role = ((el.getAttribute && el.getAttribute("role")) || "").toLowerCase();
    const isImg = tag === "img" || role === "img" || role === "image";
    if (!isImg) return el;
    try {
      if (el.closest) {
        const card = el.closest(".imgtext");
        if (card) {
          const btn = card.querySelector(".imgtextbtn");
          if (btn) return btn;
          return card;
        }
        const hit = el.closest("a[href], button, [role='button'], [role='link'], [onclick]");
        if (hit) return hit;
      }
      let p = el.parentElement;
      for (let i = 0; i < 5 && p; i++) {
        try {
          const st = window.getComputedStyle(p);
          if (st && st.cursor === "pointer") return p;
        } catch (_) {}
        p = p.parentElement;
      }
    } catch (_) {}
    return el;
  }

  function doInteract(msg) {
    const action = (msg.action || "").toLowerCase();
    if (action === "scroll" && !msg.ref) {
      const dir = (msg.direction || "down").toLowerCase();
      const dy = dir === "up" ? -400 : dir === "down" ? 400 : 0;
      const dx = dir === "left" ? -400 : dir === "right" ? 400 : 0;
      window.scrollBy(dx, dy);
      return { ok: true, message: "scrolled " + dir };
    }

    const { el, entry } = resolveRef(msg.ref, msg.css_path);
    const target = action === "click" ? resolveClickTarget(el) : el;
    if (looksLikeFileChooser(el) || looksLikeFileChooser(target)) {
      throw new Error("本批不支持文件选择。请用户自己在窗里选文件。");
    }
    if ((action === "type" || action === "select") && looksLikeNotEditable(el)) {
      throw new Error("目标不可编辑");
    }
    target.scrollIntoView({ block: "center", inline: "nearest" });

    if (action === "click") {
      const clicked = clickWithHitTest(target);
      if (!clicked.ok) return clicked;
      return { ok: true, message: "clicked", probe: entry.probe };
    }

    if (action === "type") {
      const text = msg.text == null ? "" : String(msg.text);
      const tag = (el.tagName || "").toLowerCase();
      el.focus();
      try {
        el.click();
      } catch (_) {}

      if (tag === "input" || tag === "textarea") {
        const filled = fillTextControl(el, text);
        if (!filled.ok) {
          throw new Error(filled.error || "输入未生效");
        }
        // Esc 收联想，避免挡「百度一下」
        el.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
        el.dispatchEvent(new KeyboardEvent("keyup", { key: "Escape", bubbles: true }));
        return {
          ok: true,
          message: "typed value=" + JSON.stringify(filled.value),
          probe: entry.probe
        };
      }

      if (el.isContentEditable) {
        el.focus();
        document.execCommand("selectAll", false, null);
        document.execCommand("delete", false, null);
        const ok = document.execCommand("insertText", false, text);
        if (!ok) {
          el.textContent = text;
        }
        el.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          cancelable: true,
          inputType: "insertText",
          data: text
        }));
        return { ok: true, message: "typed contenteditable", probe: entry.probe };
      }

      throw new Error("目标不可输入");
    }

    if (action === "press") {
      const key = msg.key || "Enter";
      el.focus();
      el.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
      el.dispatchEvent(new KeyboardEvent("keyup", { key, bubbles: true }));
      if (key === "Enter" && (el.tagName || "").toLowerCase() === "form") {
        el.requestSubmit && el.requestSubmit();
      }
      return { ok: true, message: "pressed " + key, probe: entry.probe };
    }

    if (action === "select") {
      if ((el.tagName || "").toLowerCase() !== "select") {
        throw new Error("select 仅用于 <select>");
      }
      const option = String(msg.option || "");
      let matched = false;
      for (let i = 0; i < el.options.length; i++) {
        const opt = el.options[i];
        if (opt.value === option || (opt.textContent || "").trim() === option) {
          el.selectedIndex = i;
          matched = true;
          break;
        }
      }
      if (!matched) throw new Error("未找到选项: " + option);
      el.dispatchEvent(new Event("change", { bubbles: true }));
      return { ok: true, message: "selected", probe: entry.probe };
    }

    if (action === "scroll") {
      const scrolled = scrollRefIntoView(el);
      if (!scrolled.ok) return scrolled;
      return { ok: true, message: "scrolled to ref", probe: entry.probe };
    }

    throw new Error("不支持的 action: " + action);
  }

  function setNativeValue(el, value) {
    const tag = (el.tagName || "").toLowerCase();
    const proto = tag === "textarea"
      ? window.HTMLTextAreaElement.prototype
      : window.HTMLInputElement.prototype;
    const desc = Object.getOwnPropertyDescriptor(proto, "value");
    if (desc && desc.set) {
      desc.set.call(el, value);
    } else {
      el.value = value;
    }
  }

  function fillTextControl(el, text) {
    if (looksLikeNotEditable(el)) {
      return { ok: false, error: "目标不可编辑" };
    }
    // 1) 全选清空（对百度等受控框更稳）
    try {
      el.select();
    } catch (_) {}
    setNativeValue(el, "");
    try {
      el.dispatchEvent(new InputEvent("input", {
        bubbles: true,
        cancelable: true,
        inputType: "deleteContentBackward",
        data: null
      }));
    } catch (_) {
      el.dispatchEvent(new Event("input", { bubbles: true }));
    }

    // 2) 优先 execCommand insertText（走编辑器通道）
    let usedExec = false;
    try {
      usedExec = document.execCommand("insertText", false, text);
    } catch (_) {
      usedExec = false;
    }

    // 3) 原生 setter + InputEvent 兜底
    if (!usedExec || String(el.value || "") !== text) {
      setNativeValue(el, text);
      try {
        el.dispatchEvent(new InputEvent("input", {
          bubbles: true,
          cancelable: true,
          inputType: "insertText",
          data: text
        }));
      } catch (_) {
        el.dispatchEvent(new Event("input", { bubbles: true }));
      }
      el.dispatchEvent(new Event("change", { bubbles: true }));
    }

    const got = String(el.value || "");
    if (got !== text) {
      // 4) 逐字输入再试一次
      setNativeValue(el, "");
      el.dispatchEvent(new Event("input", { bubbles: true }));
      for (let i = 0; i < text.length; i++) {
        const ch = text.charAt(i);
        const next = text.slice(0, i + 1);
        setNativeValue(el, next);
        try {
          el.dispatchEvent(new InputEvent("input", {
            bubbles: true,
            cancelable: true,
            inputType: "insertText",
            data: ch
          }));
        } catch (_) {
          el.dispatchEvent(new Event("input", { bubbles: true }));
        }
      }
      el.dispatchEvent(new Event("change", { bubbles: true }));
    }

    const finalVal = String(el.value || "");
    if (finalVal !== text) {
      return {
        ok: false,
        error: "输入未生效（读回 value=" + JSON.stringify(finalVal) + "，期望 " + JSON.stringify(text) + "）"
      };
    }
    return { ok: true, value: finalVal };
  }

  function doDownloadPrep(ref, cssPathHint) {
    const { el, entry } = resolveRef(ref, cssPathHint);
    const probe = entry.probe || {};
    let href = probe.href || el.href || null;
    if (href && href.indexOf("javascript:") === 0) href = null;
    return { ok: true, href, probe };
  }

  function doDownloadClick(ref, cssPathHint) {
    const { el } = resolveRef(ref, cssPathHint);
    const target = resolveClickTarget(el);
    if (looksLikeFileChooser(el) || looksLikeFileChooser(target)) {
      throw new Error("本批不支持文件选择。请用户自己在窗里选文件。");
    }
    try { target.scrollIntoView({ block: "center" }); } catch (_) {}
    return clickWithHitTest(target);
  }
})();
