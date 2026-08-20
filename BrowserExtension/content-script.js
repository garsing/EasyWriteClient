/**
 * 易写 — 页内 snapshot / interact / download 辅助
 */
(function () {
  if (window.__yiwriteContentLoaded) {
    return;
  }
  window.__yiwriteContentLoaded = true;

  const MAX_CHARS = 120000;
  const MAX_NODES = 800;

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
        return doDownloadPrep(msg.ref);
      case "download_click":
        return doDownloadClick(msg.ref);
      default:
        return { ok: false, error: "unknown: " + msg.type };
    }
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
      if (/^h[1-6]$/.test(tag)) return "heading";
      if (tag === "li") return "listitem";
      return tag || "generic";
    }

    function nameOf(el) {
      const aria = el.getAttribute("aria-label");
      if (aria) return aria.trim();
      if (el.alt) return String(el.alt).trim();
      if (el.title) return String(el.title).trim();
      if (el.placeholder) return String(el.placeholder).trim();
      const label = el.labels && el.labels[0];
      if (label) return (label.innerText || "").trim().slice(0, 120);
      const text = (el.innerText || el.textContent || "").replace(/\s+/g, " ").trim();
      return text.slice(0, 120);
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
        pageHasPasswordInput: !!document.querySelector('input[type="password"]')
      };
    }

    function shouldSkip(el) {
      if (!el || el.nodeType !== 1) return true;
      const tag = (el.tagName || "").toLowerCase();
      if (tag === "script" || tag === "style" || tag === "noscript" || tag === "svg") return true;
      const st = window.getComputedStyle(el);
      if (st && (st.display === "none" || st.visibility === "hidden")) return true;
      return false;
    }

    function interesting(el) {
      const tag = (el.tagName || "").toLowerCase();
      if (["a", "button", "input", "textarea", "select", "img", "label"].indexOf(tag) >= 0) {
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
        lines.push(line);
        nodes[ref] = {
          role,
          name,
          cssPath: cssPath(el),
          probe: probeOf(el)
        };

        if (lines.join("\n").length > MAX_CHARS) {
          truncated = true;
          truncatedReason = "char_limit";
          return;
        }
      }

      const children = el.children ? Array.from(el.children) : [];
      for (let i = 0; i < children.length; i++) {
        walk(children[i], depth + 1);
        if (truncated) return;
      }
    }

    const root = document.body || document.documentElement;
    walk(root, 0);

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

  function resolveRef(ref) {
    const nodes = window.__yiwriteNodes || {};
    const entry = nodes[ref];
    if (!entry || !entry.cssPath) {
      throw new Error("未知 ref: " + ref + "；请重新 snapshot");
    }
    const el = document.querySelector(entry.cssPath);
    if (!el) {
      throw new Error("ref 对应元素已失效: " + ref);
    }
    return { el, entry };
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

    const { el, entry } = resolveRef(msg.ref);
    el.scrollIntoView({ block: "center", inline: "nearest" });

    if (action === "click") {
      el.focus();
      el.click();
      return { ok: true, message: "clicked", probe: entry.probe };
    }

    if (action === "type") {
      const text = msg.text == null ? "" : String(msg.text);
      el.focus();
      const tag = (el.tagName || "").toLowerCase();
      if (tag === "input" || tag === "textarea") {
        const proto = tag === "input" ? window.HTMLInputElement.prototype : window.HTMLTextAreaElement.prototype;
        const setter = Object.getOwnPropertyDescriptor(proto, "value");
        if (setter && setter.set) {
          setter.set.call(el, text);
        } else {
          el.value = text;
        }
        el.dispatchEvent(new Event("input", { bubbles: true }));
        el.dispatchEvent(new Event("change", { bubbles: true }));
      } else if (el.isContentEditable) {
        el.textContent = text;
        el.dispatchEvent(new Event("input", { bubbles: true }));
      } else {
        throw new Error("目标不可输入");
      }
      return { ok: true, message: "typed", probe: entry.probe };
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
      el.scrollIntoView({ block: "center" });
      return { ok: true, message: "scrolled to ref", probe: entry.probe };
    }

    throw new Error("不支持的 action: " + action);
  }

  function doDownloadPrep(ref) {
    const { el, entry } = resolveRef(ref);
    const probe = entry.probe || {};
    let href = probe.href || el.href || null;
    if (href && href.indexOf("javascript:") === 0) href = null;
    return { ok: true, href, probe };
  }

  function doDownloadClick(ref) {
    const { el } = resolveRef(ref);
    el.scrollIntoView({ block: "center" });
    el.click();
    return { ok: true };
  }
})();
