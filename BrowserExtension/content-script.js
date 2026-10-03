/**
 * 易写 — 页内 snapshot / interact / download 辅助
 */
(function () {
  const CONTENT_VER = 37;
  if (window.__yiwriteContentVer === CONTENT_VER && window.__yiwriteOnMsg) {
    return;
  }
  if (window.__yiwriteOnMsg) {
    try { chrome.runtime.onMessage.removeListener(window.__yiwriteOnMsg); } catch (_) {}
  }
  window.__yiwriteContentVer = CONTENT_VER;
  window.__yiwriteContentLoaded = true;

  const MAX_CHARS = 36000;
  const MAX_NODES = 300;

  function onMsg(msg, _sender, sendResponse) {
    handle(msg).then(sendResponse).catch((e) => {
      sendResponse({ ok: false, error: String(e && e.message ? e.message : e) });
    });
    return true;
  }
  window.__yiwriteOnMsg = onMsg;
  chrome.runtime.onMessage.addListener(onMsg);

  async function handle(msg) {
    if (!msg || !msg.type) {
      return { ok: false, error: "bad message" };
    }
    switch (msg.type) {
      case "ping":
        return { ok: true, ver: CONTENT_VER };
      case "snapshot":
        return doSnapshot(msg);
      case "interact":
        return doInteract(msg);
      case "download_prep":
        return doDownloadPrep(msg.ref, msg.css_path);
      case "download_click":
        return doDownloadClick(msg);
      default:
        return { ok: false, error: "unknown: " + msg.type };
    }
  }

  function parentComposedTop(el) {
    if (!el) return null;
    if (el.assignedSlot) return el.assignedSlot;
    if (el.parentElement) return el.parentElement;
    const root = el.getRootNode && el.getRootNode();
    if (root && root.host) return root.host;
    return null;
  }

  function hostControl(el) {
    let n = el;
    while (n) {
      const tag = (n.tagName || "").toUpperCase();
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return n;
      n = parentComposedTop(n);
    }
    return el;
  }

  function isTemporalInput(el) {
    if (!el || (el.tagName || "").toUpperCase() !== "INPUT") return false;
    return /^(date|time|datetime-local|month|week)$/i.test(String(el.type || ""));
  }

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function ymd(d) {
    return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
  }

  function parseDateText(raw) {
    const s = String(raw || "").trim();
    let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    m = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return null;
  }

  function normalizeTemporal(kind, text, current, segment) {
    const raw = String(text || "").trim();
    if (!raw) return "";
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    if (raw === "今天" || /^today$/i.test(raw)) return formatTemporalKind(kind, today);
    function formatTemporalKind(kind, d) {
      if (kind === "month") return ymd(d).slice(0, 7);
      if (kind === "week") return isoWeek(d);
      if (kind === "time") return "00:00";
      if (kind === "datetime-local") return ymd(d) + "T00:00";
      return ymd(d);
    }
    function isoWeek(d) {
      const day = (d.getDay() + 6) % 7;
      const thu = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 3 - day);
      const jan4 = new Date(thu.getFullYear(), 0, 4);
      const week = 1 + Math.round(((thu - jan4) / 86400000 - 3 + ((jan4.getDay() + 6) % 7)) / 7);
      return thu.getFullYear() + "-W" + pad2(week);
    }
    if (raw === "昨天" || /^yesterday$/i.test(raw)) {
      const d = new Date(today);
      d.setDate(d.getDate() - 1);
      return formatTemporalKind(kind, d);
    }
    if (raw === "明天" || /^tomorrow$/i.test(raw)) {
      const d = new Date(today);
      d.setDate(d.getDate() + 1);
      return formatTemporalKind(kind, d);
    }
    if (kind === "date") {
      const full = parseDateText(raw);
      if (full && !isNaN(full.getTime())) return ymd(full);
      let basis = parseDateText(current) || today;
      const n = Number(raw);
      if (String(n) === raw || /^-?\d+$/.test(raw)) {
        if (segment === "year" && n >= 1000 && n <= 9999) {
          basis = new Date(n, basis.getMonth(), basis.getDate());
          return ymd(basis);
        }
        if (segment === "month" && n >= 1 && n <= 12) {
          basis = new Date(basis.getFullYear(), n - 1, basis.getDate());
          return ymd(basis);
        }
        if (segment === "day" && n >= 1 && n <= 31) {
          basis = new Date(basis.getFullYear(), basis.getMonth(), n);
          return ymd(basis);
        }
      }
    }
    if (kind === "month") {
      const m = raw.match(/^(\d{4})[-/](\d{1,2})$/);
      if (m) return m[1] + "-" + pad2(m[2]);
    }
    if (kind === "week") {
      const w = raw.match(/^(\d{4})-W(\d{1,2})$/i);
      if (w) return w[1] + "-W" + pad2(w[2]);
    }
    if (kind === "time") {
      const tm = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (tm) {
        return pad2(tm[1]) + ":" + tm[2] + (tm[3] ? ":" + tm[3] : "");
      }
    }
    if (kind === "datetime-local") {
      const t = raw.replace(" ", "T");
      if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(t)) return t;
    }
    return "";
  }

  function temporalSegment(el) {
    let lab = "";
    try {
      lab = String((el.getAttribute && el.getAttribute("aria-label")) || "") + " " + String(el.textContent || "");
    } catch (_) {}
    lab = lab.replace(/\s+/g, " ").toLowerCase();
    if (/年|year/.test(lab)) return "year";
    if (/月|month/.test(lab) && !/日|day/.test(lab)) return "month";
    if (/日|day/.test(lab)) return "day";
    return "";
  }

  function fillTemporalControl(el, text) {
    const host = hostControl(el);
    if (!isTemporalInput(host)) return { ok: false, error: "不是日期/时间控件" };
    const want = normalizeTemporal(
      String(host.type || "").toLowerCase(),
      text,
      host.value || "",
      temporalSegment(el)
    );
    if (!want) {
      return { ok: false, error: "日期/时间格式无法识别，请用 YYYY-MM-DD（今天/昨天/明天也可以）" };
    }
    setNativeValue(host, want);
    try { host.dispatchEvent(new Event("input", { bubbles: true })); } catch (_) {}
    try { host.dispatchEvent(new Event("change", { bubbles: true })); } catch (_) {}
    try { host.blur && host.blur(); } catch (_) {}
    const got = String(host.value || "");
    if (got !== want) {
      return {
        ok: false,
        error: "输入未生效（读回 value=" + JSON.stringify(got) + "，期望 " + JSON.stringify(want) + "）"
      };
    }
    return { ok: true, value: got };
  }

  function looksLikeNotEditable(node) {
    if (!node) return false;
    function locked(el) {
      if (!el) return false;
      try {
        if (el.disabled) return true;
        if (el.matches && el.matches(":disabled")) return true;
      } catch (_) {}
      try {
        if (el.readOnly) return true;
      } catch (_) {}
      try {
        const ad = el.getAttribute && el.getAttribute("aria-disabled");
        if (ad != null && String(ad).toLowerCase() !== "false") return true;
      } catch (_) {}
      return false;
    }
    return locked(node) || locked(hostControl(node));
  }

  function findSelect(el) {
    if (!el) return null;
    if ((el.tagName || "").toLowerCase() === "select") return el;
    try {
      if (el.querySelector) {
        const inner = el.querySelector("select");
        if (inner) return inner;
      }
      const prev = el.previousElementSibling;
      if (prev && (prev.tagName || "").toLowerCase() === "select") return prev;
      const next = el.nextElementSibling;
      if (next && (next.tagName || "").toLowerCase() === "select") return next;
      const g = el.closest && el.closest(".input-group");
      if (g && g.querySelector) {
        const inGroup = g.querySelector("select");
        if (inGroup) return inGroup;
      }
    } catch (_) {}
    return null;
  }

  function isTypeableInput(el) {
    if (!el || el.nodeType !== 1) return false;
    if (el.isContentEditable) return true;
    const tag = (el.tagName || "").toLowerCase();
    if (tag === "textarea") return true;
    if (tag !== "input") return false;
    const typ = String((el.getAttribute && el.getAttribute("type")) || el.type || "text").toLowerCase();
    return !/^(button|submit|reset|checkbox|radio|file|image|hidden)$/.test(typ);
  }

  function onlyTypeableIn(root) {
    if (!root || !root.querySelectorAll) return null;
    const list = root.querySelectorAll("input, textarea, [contenteditable='true'], [contenteditable='']");
    let hit = null;
    for (let i = 0; i < list.length; i++) {
      if (!isTypeableInput(list[i])) continue;
      if (hit) return null;
      hit = list[i];
    }
    return hit;
  }

  function labelControl(el) {
    if (!el) return null;
    try {
      const lab = (el.tagName || "").toLowerCase() === "label"
        ? el
        : (el.closest && el.closest("label"));
      if (!lab) return null;
      const id = lab.htmlFor || (lab.getAttribute && lab.getAttribute("for"));
      if (id) {
        const by = document.getElementById(id);
        if (isTypeableInput(by)) return by;
      }
      return onlyTypeableIn(lab);
    } catch (_) {
      return null;
    }
  }

  function nameHint(el) {
    if (!el) return "";
    try {
      const aria = el.getAttribute && el.getAttribute("aria-label");
      if (aria && String(aria).trim()) return String(aria).trim();
    } catch (_) {}
    let own = "";
    const nodes = el.childNodes || [];
    for (let i = 0; i < nodes.length; i++) {
      if (nodes[i].nodeType === 3) own += nodes[i].textContent || "";
    }
    own = own.replace(/\s+/g, " ").trim();
    return own.length > 0 && own.length <= 40 ? own : "";
  }

  function findTypeableByName(hint) {
    if (!hint) return null;
    const list = document.querySelectorAll("input, textarea, [contenteditable='true'], [contenteditable='']");
    const exact = [];
    const contained = [];
    for (let i = 0; i < list.length; i++) {
      const c = list[i];
      if (!isTypeableInput(c)) continue;
      const an = String((c.getAttribute && (c.getAttribute("aria-label") || c.getAttribute("name"))) || "").trim();
      if (!an) continue;
      if (an === hint) exact.push(c);
      else if (an.length >= 2 && hint.indexOf(an) >= 0) contained.push(c);
    }
    if (exact.length === 1) return exact[0];
    if (exact.length === 0 && contained.length === 1) return contained[0];
    return null;
  }

  function findTypeable(el, entry) {
    if (!el) return null;
    const up = hostControl(el);
    if (isTypeableInput(up)) return up;
    if (el.isContentEditable) return el;
    const fromLabel = labelControl(el);
    if (fromLabel) return fromLabel;
    const hint = (entry && (entry.name || (entry.probe && entry.probe.accessibleName))) || nameHint(el);
    const byName = findTypeableByName(hint);
    if (byName) return byName;
    const only = onlyTypeableIn(el);
    if (only) return only;
    try {
      const prev = el.previousElementSibling;
      if (isTypeableInput(prev)) return prev;
      const next = el.nextElementSibling;
      if (isTypeableInput(next)) return next;
      const g = el.closest && el.closest(".input-group");
      if (g) {
        const inG = onlyTypeableIn(g);
        if (inG) return inG;
      }
    } catch (_) {}
    return null;
  }

  function looksLikePassword(el) {
    if (!el) return false;
    const typ = String((el.getAttribute && el.getAttribute("type")) || el.type || "").toLowerCase();
    if (typ === "password") return true;
    const ac = (el.getAttribute && el.getAttribute("autocomplete")) || "";
    return /password/i.test(ac);
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
      if (tag === "tr") return "row";
      if (tag === "td" || tag === "th") return "cell";
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

    function isValidatorText(s) {
      const n = String(s || "").replace(/\s+/g, " ").trim();
      if (!n) return false;
      if (n === "不能为空" || n === "必填" || n === "此项必填") return true;
      if (n.indexOf("输入格式不正确") >= 0) return true;
      const low = n.toLowerCase();
      return low === "required" || low === "this field is required"
        || low === "cannot be empty" || low === "must not be empty";
    }

    function isValidatorLabel(lab) {
      if (!lab) return false;
      if (lab.getAttribute && lab.getAttribute("data-bv-validator")) return true;
      if (/\bhelp-block\b/.test(String(lab.className || ""))) return true;
      try {
        const st = window.getComputedStyle(lab);
        if (st && st.display === "none" && isValidatorText(lab.innerText)) return true;
      } catch (_) {}
      return isValidatorText(lab.innerText);
    }

    function nameOf(el) {
      const tag = (el.tagName || "").toLowerCase();
      const t = (s) => String(s || "").replace(/\s+/g, " ").trim();
      if (tag === "html" || tag === "body") {
        return t(document.title).slice(0, 120);
      }
      if (isDataRow(el)) {
        const row = rowName(el);
        if (row) return row;
      }
      if (tag === "iframe" || tag === "frame") {
        const aria = el.getAttribute("aria-label");
        if (t(aria)) return t(aria);
        if (t(el.title)) return t(el.title);
        const src = el.getAttribute("src") || el.src || "";
        return srcTail(src) || "iframe";
      }
      const aria = el.getAttribute("aria-label");
      if (t(aria) && !isValidatorText(aria)) return t(aria).slice(0, 120);
      if (t(el.title) && !isValidatorText(el.title)) return t(el.title).slice(0, 120);
      if (t(el.alt)) return t(el.alt).slice(0, 120);
      try {
        const img = el.querySelector && el.querySelector("img[alt]");
        if (img && t(img.alt)) return t(img.alt).slice(0, 120);
        const st = el.querySelector && el.querySelector("svg title");
        if (st && t(st.textContent)) return t(st.textContent).slice(0, 120);
        const area = el.querySelector && el.querySelector("area[alt]");
        if (area && t(area.alt)) return t(area.alt).slice(0, 120);
      } catch (_) {}
      try {
        const g = el.closest && el.closest(".input-group");
        if (g) {
          const addon = g.querySelector(".input-group-addon");
          const n = t(addon && addon.innerText).replace(/^\*+\s*/, "");
          if (n && !isValidatorText(n)) return n.slice(0, 120);
        }
      } catch (_) {}
      if (el.placeholder && !isValidatorText(el.placeholder)) {
        return String(el.placeholder).trim().slice(0, 120);
      }
      const labels = el.labels ? Array.from(el.labels) : [];
      for (let i = 0; i < labels.length; i++) {
        if (isValidatorLabel(labels[i])) continue;
        const n = t(labels[i].innerText);
        if (n) return n.slice(0, 120);
      }
      const text = t(el.innerText || el.textContent);
      return isValidatorText(text) ? "" : text.slice(0, 120);
    }

    function optionsOf(el) {
      if ((el.tagName || "").toLowerCase() !== "select" || !el.options) return null;
      const out = [];
      for (let i = 0; i < el.options.length && out.length < 8; i++) {
        const tx = String(el.options[i].text || "").replace(/\s+/g, " ").trim();
        if (tx) out.push(tx);
      }
      return out.length ? out.join("|") : null;
    }

    function valueOf(el) {
      const tag = (el.tagName || "").toLowerCase();
      if (tag === "input" && (el.getAttribute("type") || "").toLowerCase() === "password") {
        return null;
      }
      if (tag === "select") {
        const o = el.options && el.selectedIndex >= 0 ? el.options[el.selectedIndex] : null;
        return o ? String(o.text || o.value || "").replace(/\s+/g, " ").trim() : "";
      }
      if (tag === "input" || tag === "textarea") {
        return el.value != null ? String(el.value).slice(0, 120) : "";
      }
      if (el.isContentEditable) {
        return String(el.innerText || "").replace(/\s+/g, " ").trim().slice(0, 120);
      }
      return null;
    }

    function probeOf(el) {
      const tag = (el.tagName || "").toLowerCase();
      const ctrl = labelControl(el);
      const src = ctrl || el;
      const srcTag = (src.tagName || "").toLowerCase();
      return {
        tag,
        inputType: srcTag === "input" ? (src.getAttribute("type") || "text") : null,
        autocomplete: src.getAttribute("autocomplete") || el.getAttribute("autocomplete") || null,
        role: roleOf(el),
        accessibleName: nameOf(el),
        valueAttr: src.value != null ? String(src.value).slice(0, 200) : (el.value != null ? String(el.value).slice(0, 200) : null),
        innerText: (tag === "html" || tag === "body")
          ? String(document.title || "").replace(/\s+/g, " ").trim().slice(0, 120)
          : (el.innerText || "").replace(/\s+/g, " ").trim().slice(0, 200),
        href: el.href || el.getAttribute("href") || null,
        dataUrl: el.src || el.getAttribute("src") || null,
        hasDownloadAttr: el.hasAttribute("download"),
        pageHasPasswordInput: !!document.querySelector('input[type="password"]'),
        isFileChooser: looksLikeFileChooser(el) || looksLikeFileChooser(src),
        disabled: (() => {
          try {
            if (src.disabled) return true;
            if (src.matches && src.matches(":disabled")) return true;
          } catch (_) {}
          return false;
        })(),
        readOnly: !!src.readOnly,
        ariaDisabled: (() => {
          const ad = src.getAttribute && src.getAttribute("aria-disabled");
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

    function isNativeInteractive(el) {
      const tag = ((el && el.tagName) || "").toLowerCase();
      if (["a", "button", "input", "textarea", "select"].indexOf(tag) >= 0) return true;
      const role = ((el.getAttribute && el.getAttribute("role")) || "").toLowerCase();
      return role === "link" || role === "button";
    }

    function isPointerCard(el) {
      if (isNativeInteractive(el)) return false;
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

    function isCardGroup(el) {
      if (!el || el.nodeType !== 1 || isNativeInteractive(el) || !el.querySelector) return false;
      if (!el.querySelector("a[href], [role='link']")) return false;
      if (!el.querySelector("img") && !isPointerCard(el)) return false;
      const n = nameOf(el);
      if (!n || n.length < 8) return false;
      try {
        const w = el.offsetWidth;
        const h = el.offsetHeight;
        if (w < 100 || h < 60 || w > 720 || h > 900) return false;
      } catch (_) {
        return false;
      }
      const links = el.querySelectorAll("a[href], [role='link']");
      return links.length >= 1 && links.length <= 8;
    }

    function closestCardGroup(el) {
      let p = el;
      while (p && p.nodeType === 1 && p !== document.body) {
        if (isCardGroup(p)) return p;
        p = p.parentElement;
      }
      return null;
    }

    function isOutermostCardGroup(el) {
      return isCardGroup(el) && !closestCardGroup(el.parentElement);
    }

    function inCardNotShell(el) {
      return !!closestCardGroup(el.parentElement);
    }

    function isVisibleImage(el) {
      if (!el || (el.tagName || "").toLowerCase() !== "img") return false;
      try {
        const r = el.getBoundingClientRect();
        return r.width >= 40 && r.height >= 40;
      } catch (_) {
        return false;
      }
    }

    function isNamedIcon(el) {
      if (!el || el.nodeType !== 1 || isNativeInteractive(el)) return false;
      const tag = (el.tagName || "").toLowerCase();
      if (tag !== "svg" && tag !== "i" && tag !== "use" && tag !== "path") return false;
      const n = (
        (el.getAttribute && (el.getAttribute("aria-label") || el.getAttribute("title")))
        || (el.querySelector && el.querySelector("title") && el.querySelector("title").textContent)
        || ""
      ).replace(/\s+/g, " ").trim();
      return n.length > 0;
    }

    function isTextLeaf(el) {
      if (!el || isNativeInteractive(el)) return false;
      if (el.querySelector && el.querySelector("img, a, button, input, select, textarea, svg")) {
        return false;
      }
      const own = ownText(el);
      return own.length >= 2 && own.length <= 40;
    }

    function isShortOwnText(el) {
      const own = ownText(el);
      return own.length >= 2 && own.length <= 16;
    }

    function isVisibleBox(el, st) {
      if (!st || st.display === "none" || st.visibility === "hidden") return false;
      try {
        const rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      } catch (_) {
        return false;
      }
    }

    /// 人能点的短标签：pointer + 自己有短文案。不限左侧（中间栏 Tab 也要进树）。
    function isPointerLabel(el) {
      try {
        if (closestCardGroup(el)) return false;
        const st = window.getComputedStyle(el);
        if (!st || st.cursor !== "pointer") return false;
        if (!isShortOwnText(el)) return false;
        return isVisibleBox(el, st);
      } catch (_) {
        return false;
      }
    }

    function isIgnoredAriaRole(role) {
      const r = String(role || "").toLowerCase();
      return r === "none" || r === "presentation";
    }

    function isDataRow(el) {
      if (!el || el.nodeType !== 1) return false;
      const tag = (el.tagName || "").toLowerCase();
      if (tag === "tr") return true;
      const role = (el.getAttribute("role") || "").toLowerCase();
      if (role === "row") return true;
      const cls = classNameOf(el);
      return /\bel-table__row\b/.test(cls) || /\bant-table-row\b/.test(cls);
    }

    function rowName(el) {
      const cells = el.querySelectorAll
        ? el.querySelectorAll("td, th, [role='cell'], [role='gridcell'], [role='columnheader']")
        : [];
      const parts = [];
      for (let i = 0; i < cells.length && parts.length < 6; i++) {
        const n = (ownText(cells[i]) || String(cells[i].innerText || "").replace(/\s+/g, " ").trim());
        if (n && n.length <= 40) parts.push(n);
      }
      return parts.join(" ").slice(0, 80);
    }

    function hasClickHint(el) {
      if (!el || el.nodeType !== 1) return false;
      if (el.getAttribute && el.getAttribute("onclick")) return true;
      if (typeof el.onclick === "function") return true;
      const ti = el.getAttribute ? el.getAttribute("tabindex") : null;
      if (ti != null && ti !== "" && !isNaN(Number(ti)) && Number(ti) >= 0 && isShortOwnText(el)) {
        return true;
      }
      return false;
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
      if (isNativeInteractive(el)) return true;
      if (isOutermostCardGroup(el)) return true;
      if (inCardNotShell(el)) {
        if (hasDetailLabel(el) || isImgtextBtn(el)) return true;
        if (isVisibleImage(el) || isNamedIcon(el) || isTextLeaf(el)) return true;
        return false;
      }
      if (isPointerCard(el) && closestCardGroup(el.parentElement)) return false;
      if (hasDetailLabel(el) || isImgtextBtn(el)
        || (isPointerCard(el) && !el.querySelector("a[href], [role='link']"))
        || isPointerLabel(el) || hasClickHint(el)) return true;
      if (tag === "img" && !closestPointerCard(el.parentElement)) return true;
      if (["label", "iframe", "frame"].indexOf(tag) >= 0) {
        return true;
      }
      if (tag === "area" && (el.getAttribute("alt") || el.title || el.getAttribute("aria-label"))) {
        return true;
      }
      if (el.isContentEditable) return true;
      if (isDataRow(el)) return true;
      const ariaRole = el.getAttribute("role");
      if (ariaRole && !isIgnoredAriaRole(ariaRole)) return true;
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

      const groupOnly = isOutermostCardGroup(el);
      const print = interesting(el) || depth === 0;
      if (print) {
        const role = roleOf(el);
        let name = nameOf(el);
        const tag = (el.tagName || "").toLowerCase();
        if (!name && tag === "a" && el.querySelector && el.querySelector("img")) {
          const g = closestCardGroup(el);
          if (g) name = nameOf(g);
        }
        if (!name && isVisibleImage(el)) {
          name = (el.getAttribute("alt") || "").trim();
          if (!name) {
            const g = closestCardGroup(el);
            if (g) name = nameOf(g);
          }
        }
        const indent = "  ".repeat(Math.min(depth, 12));
        let line = indent + "- " + role;
        if (name) line += ' "' + name.replace(/"/g, "'") + '"';
        if (!groupOnly) {
          const ref = nextRef();
          line += " [ref=" + ref + "]";
          const val = valueOf(el);
          if (val != null) line += ' value="' + String(val).replace(/"/g, "'") + '"';
          const opts = optionsOf(el);
          if (opts) line += " options=" + opts;
          if (tag === "input" && (el.getAttribute("type") || "") === "password") {
            line += " (password)";
          }
          if ((isHoverRevealControl(el) || hasDetailLabel(el) || isImgtextBtn(el))
            && (isHoverHiddenStyle(el) || (window.getComputedStyle(el).display === "none"))) {
            line += " (隐藏)";
          }
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
        }
        lines.push(line);

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
      const childDepth = print ? depth + 1 : depth;
      for (let i = 0; i < children.length; i++) {
        walk(children[i], childDepth);
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
    const path = cssPathHint || entry.cssPath;
    if (!path) {
      throw new Error("未知 ref: " + ref + "；请重新 snapshot");
    }
    const el = document.querySelector(path);
    if (!el) {
      throw new Error("ref 对应元素已失效: " + ref + "。请重新 snapshot 后再操作，不要改用 click。");
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
    n = t;
    while (n) {
      if (n === hit) return true;
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

  function hitTestError(t, opts) {
    const hover = !!(opts && opts.hover);
    if (!hover && isForceClickTarget(t)) return null;
    let r = { width: 0, height: 0, left: 0, top: 0 };
    try { r = t.getBoundingClientRect(); } catch (_) {}
    if (r.width <= 0 || r.height <= 0) {
      return hover
        ? "没有可悬停区域。请悬停可见的父节点或卡片。"
        : "点击被挡住：目标没有可点区域。";
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

  function normalizeClickMode(v) {
    return String(v || "").toLowerCase() === "legacy" ? "legacy" : "pointer";
  }

  function normalizeAttachPointer(v) {
    return String(v || "").toLowerCase() === "cdp" ? "cdp" : "synthetic";
  }

  function isImageNode(el) {
    const tag = ((el && el.tagName) || "").toUpperCase();
    let role = "";
    try { role = ((el.getAttribute && el.getAttribute("role")) || "").toLowerCase(); } catch (_) {}
    return tag === "IMG" || role === "img" || role === "image";
  }

  function boxCenter(t) {
    let r = { width: 0, height: 0, left: 0, top: 0 };
    try { r = t.getBoundingClientRect(); } catch (_) {}
    return { x: r.left + r.width / 2, y: r.top + r.height / 2, w: r.width, h: r.height };
  }

  function toTopViewport(x, y) {
    let ox = x;
    let oy = y;
    try {
      let win = window;
      while (win !== win.top) {
        const frame = win.frameElement;
        if (!frame) break;
        const r = frame.getBoundingClientRect();
        ox += r.left;
        oy += r.top;
        win = win.parent;
      }
    } catch (_) {}
    return { x: ox, y: oy };
  }

  function fireSyntheticClick(t, x, y) {
    const down = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, buttons: 1 };
    const up = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, buttons: 0 };
    try { t.dispatchEvent(new MouseEvent("mousemove", down)); } catch (_) {}
    try {
      if (typeof PointerEvent === "function") {
        t.dispatchEvent(new PointerEvent("pointermove", Object.assign({}, down, { pointerId: 1, pointerType: "mouse" })));
      }
    } catch (_) {}
    try { t.dispatchEvent(new MouseEvent("mousedown", down)); } catch (_) {}
    try {
      if (typeof PointerEvent === "function") {
        t.dispatchEvent(new PointerEvent("pointerdown", Object.assign({}, down, { pointerId: 1, pointerType: "mouse" })));
      }
    } catch (_) {}
    try { t.dispatchEvent(new MouseEvent("mouseup", up)); } catch (_) {}
    try {
      if (typeof PointerEvent === "function") {
        t.dispatchEvent(new PointerEvent("pointerup", Object.assign({}, up, { pointerId: 1, pointerType: "mouse" })));
      }
    } catch (_) {}
    try { t.dispatchEvent(new MouseEvent("click", up)); } catch (_) {}
  }

  function fireSyntheticHover(t, x, y) {
    const opts = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y };
    try { t.dispatchEvent(new MouseEvent("mousemove", opts)); } catch (_) {}
    try {
      if (typeof PointerEvent === "function") {
        t.dispatchEvent(new PointerEvent("pointermove", Object.assign({}, opts, { pointerId: 1, pointerType: "mouse" })));
      }
    } catch (_) {}
  }

  async function requestCdpPointer(kind, x, y) {
    const top = kind === "prepare" || kind === "detach"
      ? { x: 0, y: 0 }
      : toTopViewport(x, y);
    const resp = await chrome.runtime.sendMessage({
      type: "yiwrite.pointer",
      kind,
      x: top.x,
      y: top.y
    });
    if (!resp || !resp.ok) {
      throw new Error("未授予调试权限或无法附加调试器，无法使用 CDP 点击。");
    }
  }

  function viewportHeight() {
    try {
      if (window.visualViewport && window.visualViewport.height) {
        return window.visualViewport.height;
      }
    } catch (_) {}
    return window.innerHeight;
  }

  function sleepMs(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  async function waitForDebuggerViewport() {
    const before = viewportHeight();
    const deadline = Date.now() + 350;
    while (Date.now() < deadline) {
      await sleepMs(16);
      if (Math.abs(viewportHeight() - before) >= 8) {
        await sleepMs(32);
        return;
      }
    }
  }

  function hitAt(x, y, fallback) {
    try {
      return document.elementFromPoint(x, y) || fallback;
    } catch (_) {
      return fallback;
    }
  }

  function emptyBoxFallback(el) {
    if (isForceClickTarget(el)) {
      if (typeof el.click === "function") el.click();
      else el.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));
      return { ok: true };
    }
    return { ok: false, error: "点击被挡住：目标没有可点区域。" };
  }

  async function clickWithPointer(el, clickMode, attachPointer, opts) {
    const skipHit = !!(opts && opts.skipHit);
    if (!skipHit) {
      const err = hitTestError(el);
      if (err) return { ok: false, error: err };
    }
    try { el.focus && el.focus(); } catch (_) {}
    if (clickMode === "legacy") {
      if (typeof el.click === "function") el.click();
      else el.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));
      return { ok: true };
    }
    const box = boxCenter(el);
    if (box.w <= 0 || box.h <= 0) {
      return emptyBoxFallback(el);
    }
    if (attachPointer === "cdp") {
      await requestCdpPointer("prepare", 0, 0);
      try {
        await waitForDebuggerViewport();
        if (!skipHit) {
          const err = hitTestError(el);
          if (err) return { ok: false, error: err };
        }
        const live = boxCenter(el);
        if (live.w <= 0 || live.h <= 0) {
          return emptyBoxFallback(el);
        }
        const receiver = skipHit ? hitAt(live.x, live.y, el) : el;
        await requestCdpPointer("click", live.x, live.y);
        try {
          receiver.dispatchEvent(new MouseEvent("click", {
            bubbles: true, cancelable: true, view: window,
            clientX: live.x, clientY: live.y, buttons: 0
          }));
        } catch (_) {}
        return { ok: true };
      } finally {
        try { await requestCdpPointer("detach", 0, 0); } catch (_) {}
      }
    }
    const receiver = skipHit ? hitAt(box.x, box.y, el) : el;
    fireSyntheticClick(receiver, box.x, box.y);
    return { ok: true };
  }

  async function hoverTarget(el, attachPointer) {
    const err = hitTestError(el, { hover: true });
    if (err) return { ok: false, error: err };
    if (attachPointer === "cdp") {
      await requestCdpPointer("prepare", 0, 0);
      try {
        await waitForDebuggerViewport();
        const again = hitTestError(el, { hover: true });
        if (again) return { ok: false, error: again };
        const live = boxCenter(el);
        await requestCdpPointer("hover", live.x, live.y);
        fireSyntheticHover(el, live.x, live.y);
      } finally {
        try { await requestCdpPointer("detach", 0, 0); } catch (_) {}
      }
      await new Promise((resolve) => setTimeout(resolve, 150));
      return { ok: true };
    }
    const box = boxCenter(el);
    fireSyntheticHover(el, box.x, box.y);
    await new Promise((resolve) => setTimeout(resolve, 150));
    return { ok: true };
  }

  function imageHrefLink(el) {
    if (!el || !el.querySelector) return null;
    const img = el.querySelector("a[href] img");
    return img && img.closest ? img.closest("a[href]") : null;
  }

  function resolveImageAncestor(el) {
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

  function resolveClickTarget(el) {
    if (!el) return el;
    const tag = (el.tagName || "").toLowerCase();
    const role = ((el.getAttribute && el.getAttribute("role")) || "").toLowerCase();
    if (tag !== "a" && role !== "link" && tag !== "button" && role !== "button") {
      const imgA = imageHrefLink(el);
      if (imgA) return imgA;
    }
    return resolveImageAncestor(el);
  }

  function resolveHoverTarget(el) {
    return resolveImageAncestor(el);
  }

  async function doInteract(msg) {
    const action = (msg.action || "").toLowerCase();
    if (action === "scroll" && !msg.ref) {
      const dir = (msg.direction || "down").toLowerCase();
      const dy = dir === "up" ? -400 : dir === "down" ? 400 : 0;
      const dx = dir === "left" ? -400 : dir === "right" ? 400 : 0;
      window.scrollBy(dx, dy);
      return { ok: true, message: "scrolled " + dir };
    }

    if (action === "press" && !msg.ref) {
      const key = msg.key || "Enter";
      const el = document.activeElement && document.activeElement !== document.body
        ? document.activeElement
        : document.documentElement;
      try { el.focus && el.focus(); } catch (_) {}
      el.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
      el.dispatchEvent(new KeyboardEvent("keyup", { key, bubbles: true }));
      return { ok: true, message: "pressed " + key };
    }

    const { el, entry } = resolveRef(msg.ref, msg.css_path);
    const target = action === "click"
      ? resolveClickTarget(el)
      : action === "hover"
        ? resolveHoverTarget(el)
        : el;
    if (looksLikeFileChooser(el) || looksLikeFileChooser(target)) {
      throw new Error("本批不支持文件选择。请用户自己在窗里选文件。");
    }
    if (action === "click") {
      const tagEl = ((el && el.tagName) || "").toLowerCase();
      const tagTg = ((target && target.tagName) || "").toLowerCase();
      if (tagEl === "select" || tagTg === "select") {
        throw new Error("这是原生下拉，请用 action=select，option 填选项原文或 value。");
      }
    }
    let typeTarget = el;
    if (action === "type") {
      typeTarget = findTypeable(el, entry);
      if (!typeTarget) {
        throw new Error("目标不可输入");
      }
      if (looksLikePassword(typeTarget)) {
        throw new Error("已拒绝向密码框输入；请用户在看得见的窗里自己填写密码");
      }
      if (looksLikeNotEditable(typeTarget) && !isTemporalInput(typeTarget)) {
        throw new Error("目标不可编辑");
      }
    }
    (action === "type" ? typeTarget : target).scrollIntoView({ block: "center", inline: "nearest" });

    if (action === "click") {
      const mode = normalizeClickMode(msg.click_mode);
      const regionImage = mode !== "legacy" && isImageNode(el);
      if (regionImage) {
        try { el.scrollIntoView({ block: "center", inline: "nearest" }); } catch (_) {}
      }
      const clicked = await clickWithPointer(
        regionImage ? el : target,
        mode,
        normalizeAttachPointer(msg.attach_pointer),
        { skipHit: regionImage }
      );
      if (!clicked.ok) return clicked;
      const used = regionImage ? el : target;
      const box = boxCenter(used);
      return {
        ok: true,
        message: "clicked",
        probe: entry.probe,
        resolved: previewHit(used),
        hit_at: previewHit(hitAt(box.x, box.y, used)),
        path: String(msg.css_path || entry.cssPath || "")
      };
    }

    if (action === "hover") {
      const hovered = await hoverTarget(target, normalizeAttachPointer(msg.attach_pointer));
      if (!hovered.ok) return hovered;
      return { ok: true, message: "hovered", probe: entry.probe };
    }

    if (action === "type") {
      const text = msg.text == null ? "" : String(msg.text);
      if (isTemporalInput(typeTarget)) {
        const filledDate = fillTemporalControl(typeTarget, text);
        if (!filledDate.ok) {
          throw new Error(filledDate.error || "输入未生效");
        }
        return {
          ok: true,
          message: "typed value=" + JSON.stringify(filledDate.value),
          probe: entry.probe
        };
      }
      const tag = (typeTarget.tagName || "").toLowerCase();
      typeTarget.focus();
      try {
        typeTarget.click();
      } catch (_) {}

      if (tag === "input" || tag === "textarea") {
        const filled = fillTextControl(typeTarget, text);
        if (!filled.ok) {
          throw new Error(filled.error || "输入未生效");
        }
        // Esc 收联想，避免挡「百度一下」
        typeTarget.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
        typeTarget.dispatchEvent(new KeyboardEvent("keyup", { key: "Escape", bubbles: true }));
        return {
          ok: true,
          message: "typed value=" + JSON.stringify(filled.value),
          probe: entry.probe
        };
      }

      if (typeTarget.isContentEditable) {
        typeTarget.focus();
        document.execCommand("selectAll", false, null);
        document.execCommand("delete", false, null);
        const ok = document.execCommand("insertText", false, text);
        if (!ok) {
          typeTarget.textContent = text;
        }
        typeTarget.dispatchEvent(new InputEvent("input", {
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
      const sel = findSelect(el);
      if (!sel) {
        const tag = ((el && el.tagName) || "?").toLowerCase();
        throw new Error(
          "select 未落到原生下拉（当前是 " + tag
            + "）。请重新 snapshot，对带 options= 的那一行再 select，不要改 click。"
        );
      }
      if (looksLikeNotEditable(sel)) {
        throw new Error("目标不可编辑");
      }
      const option = String(msg.option || "");
      let matched = false;
      const preview = [];
      for (let i = 0; i < sel.options.length; i++) {
        const opt = sel.options[i];
        const text = (opt.textContent || "").trim();
        if (preview.length < 8 && text) preview.push(text);
        if (opt.value === option || text === option || text.indexOf(option) >= 0 || String(opt.value).indexOf(option) >= 0) {
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

  async function doDownloadClick(msg) {
    const { el } = resolveRef(msg.ref, msg.css_path);
    const target = resolveClickTarget(el);
    if (looksLikeFileChooser(el) || looksLikeFileChooser(target)) {
      throw new Error("本批不支持文件选择。请用户自己在窗里选文件。");
    }
    try { target.scrollIntoView({ block: "center" }); } catch (_) {}
    return clickWithPointer(
      target,
      normalizeClickMode(msg.click_mode),
      normalizeAttachPointer(msg.attach_pointer)
    );
  }
})();
