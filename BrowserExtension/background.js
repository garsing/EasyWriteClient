/**
 * 易写浏览器助手 — MV3 service worker
 * Native Messaging host: com.yiwrite.browser_bridge
 */
const NATIVE_HOST = "com.yiwrite.browser_bridge";

let port = null;
let browserKind = detectBrowserKind();
let reconnectTimer = null;
const pendingRpc = new Map();

function detectBrowserKind() {
  try {
    const ua = (self.navigator && self.navigator.userAgent) || "";
    // Chromium Edge 带 Edg/；Chrome 不会
    if (/Edg\//.test(ua)) {
      return "edge";
    }
  } catch (_) {
    /* ignore */
  }
  return "chrome";
}

function tabUuid(tabId) {
  return browserKind + ":" + String(tabId);
}

function connectNative() {
  if (port) {
    try {
      port.disconnect();
    } catch (_) {}
    port = null;
  }

  try {
    port = chrome.runtime.connectNative(NATIVE_HOST);
  } catch (e) {
    console.warn("[yiwrite] connectNative failed", e);
    scheduleReconnect();
    return;
  }

  port.onMessage.addListener(onHostMessage);
  port.onDisconnect.addListener(() => {
    const err = chrome.runtime.lastError;
    console.warn("[yiwrite] native disconnected", err && err.message);
    port = null;
    rejectAllPending("native disconnected");
    scheduleReconnect();
  });

  sendToHost({
    type: "hello",
    browser: browserKind,
    version: chrome.runtime.getManifest().version,
    extensionId: chrome.runtime.id
  });

  pushTabsSnapshot();
}

function scheduleReconnect() {
  if (reconnectTimer) {
    clearTimeout(reconnectTimer);
  }
  reconnectTimer = setTimeout(() => {
    reconnectTimer = null;
    connectNative();
  }, 3000);
}

function sendToHost(msg) {
  if (!port) {
    return false;
  }
  try {
    port.postMessage(msg);
    return true;
  } catch (e) {
    console.warn("[yiwrite] postMessage failed", e);
    return false;
  }
}

function rejectAllPending(reason) {
  for (const [id, entry] of pendingRpc.entries()) {
    entry.reject(new Error(reason));
    pendingRpc.delete(id);
  }
}

async function collectTabs() {
  const tabs = await chrome.tabs.query({});
  return tabs
    .filter((t) => t.id != null && !String(t.url || "").startsWith("chrome://")
      && !String(t.url || "").startsWith("edge://")
      && !String(t.url || "").startsWith("devtools://"))
    .map((t) => ({
      tab_uuid: tabUuid(t.id),
      tab_id: t.id,
      title: t.title || "",
      url: t.url || "",
      active: !!t.active,
      window_id: t.windowId,
      browser: browserKind
    }));
}

async function pushTabsSnapshot() {
  const tabs = await collectTabs();
  sendToHost({ type: "tabs.snapshot", tabs });
}

function onHostMessage(msg) {
  if (!msg || typeof msg !== "object") {
    return;
  }

  const type = msg.type;
  if (type === "ping") {
    sendToHost({ type: "pong", id: msg.id });
    return;
  }

  if (type === "rpc.snapshot" || type === "rpc.interact"
    || type === "rpc.navigate" || type === "rpc.download"
    || type === "rpc.activate_tab" || type === "rpc.screenshot") {
    handleRpc(msg).catch((e) => {
      sendToHost({
        type: "rpc.result",
        id: msg.id,
        ok: false,
        error: String(e && e.message ? e.message : e)
      });
    });
  }
}

async function handleRpc(msg) {
  const id = msg.id;
  try {
    let result;
    switch (msg.type) {
      case "rpc.activate_tab":
        result = await activateTab(msg.tab_uuid);
        break;
      case "rpc.navigate":
        result = await navigateTab(msg.tab_uuid, msg.url);
        break;
      case "rpc.snapshot":
        result = await snapshotTab(msg);
        break;
      case "rpc.interact":
        result = await interactTab(msg);
        break;
      case "rpc.download":
        result = await downloadTab(msg);
        break;
      case "rpc.screenshot":
        result = await screenshotTab(msg.tab_uuid);
        break;
      default:
        throw new Error("unknown rpc: " + msg.type);
    }
    sendToHost({ type: "rpc.result", id, ok: true, result });
  } catch (e) {
    sendToHost({
      type: "rpc.result",
      id,
      ok: false,
      error: String(e && e.message ? e.message : e)
    });
  }
}

function parseTabId(tabUuid) {
  if (!tabUuid || typeof tabUuid !== "string") {
    throw new Error("tab_uuid 无效");
  }
  const i = tabUuid.lastIndexOf(":");
  if (i < 0) {
    throw new Error("tab_uuid 格式错误: " + tabUuid);
  }
  const id = Number(tabUuid.slice(i + 1));
  if (!Number.isFinite(id)) {
    throw new Error("tab_uuid 无法解析: " + tabUuid);
  }
  return id;
}

async function activateTab(tabUuid) {
  const tabId = parseTabId(tabUuid);
  await chrome.tabs.update(tabId, { active: true });
  return { activated: true };
}

async function screenshotTab(tabUuid) {
  const tabId = parseTabId(tabUuid);
  const tab = await chrome.tabs.get(tabId);
  await chrome.tabs.update(tabId, { active: true });
  if (tab.windowId != null) {
    try {
      await chrome.windows.update(tab.windowId, { focused: true });
    } catch (_) {
      /* ignore */
    }
  }
  // Native Messaging 单条消息约 1MB；PNG 全页很容易超，用 JPEG 阶梯压到能回传
  let quality = 70;
  let dataUrl = await chrome.tabs.captureVisibleTab(tab.windowId, {
    format: "jpeg",
    quality
  });
  while (dataUrl && dataUrl.length > 900000 && quality > 40) {
    quality -= 15;
    dataUrl = await chrome.tabs.captureVisibleTab(tab.windowId, {
      format: "jpeg",
      quality
    });
  }
  if (!dataUrl || typeof dataUrl !== "string" || dataUrl.indexOf(",") < 0) {
    throw new Error("captureVisibleTab 无数据");
  }
  if (dataUrl.length > 1200000) {
    throw new Error("截图过大，无法经扩展回传");
  }
  const refreshed = await chrome.tabs.get(tabId);
  return {
    image_base64: dataUrl.slice(dataUrl.indexOf(",") + 1),
    format: "jpeg",
    url: refreshed.url || tab.url || "",
    title: refreshed.title || tab.title || ""
  };
}

async function navigateTab(tabUuid, url) {
  const tabId = parseTabId(tabUuid);
  await chrome.tabs.update(tabId, { active: true, url });
  await waitTabComplete(tabId, 30000);
  const tab = await chrome.tabs.get(tabId);
  return { url: tab.url || url, title: tab.title || "" };
}

function waitTabComplete(tabId, timeoutMs) {
  return new Promise((resolve, reject) => {
    const t0 = Date.now();
    const timer = setInterval(async () => {
      try {
        const tab = await chrome.tabs.get(tabId);
        if (tab.status === "complete") {
          clearInterval(timer);
          resolve();
          return;
        }
        if (Date.now() - t0 > timeoutMs) {
          clearInterval(timer);
          reject(new Error("导航超时"));
        }
      } catch (e) {
        clearInterval(timer);
        reject(e);
      }
    }, 200);
  });
}

async function ensureContentScript(tabId) {
  try {
    await chrome.tabs.sendMessage(tabId, { type: "ping" });
    return;
  } catch (_) {
    /* inject */
  }
  await chrome.scripting.executeScript({
    target: { tabId },
    files: ["content-script.js"]
  });
}

async function snapshotTab(msg) {
  const tabId = parseTabId(msg.tab_uuid);
  await chrome.tabs.update(tabId, { active: true });
  await ensureContentScript(tabId);
  const resp = await chrome.tabs.sendMessage(tabId, {
    type: "snapshot",
    mode: msg.mode || "overview",
    ref: msg.ref || null,
    dom_supplement: msg.dom_supplement || null
  });
  if (!resp || !resp.ok) {
    throw new Error((resp && resp.error) || "snapshot 失败");
  }
  const tab = await chrome.tabs.get(tabId);
  return {
    url: tab.url || "",
    title: tab.title || "",
    snapshot: resp.snapshot,
    nodes: resp.nodes || {},
    truncated: !!resp.truncated,
    truncated_reason: resp.truncated_reason || null,
    dom_supplement: resp.dom_supplement || null
  };
}

async function interactTab(msg) {
  const tabId = parseTabId(msg.tab_uuid);
  await chrome.tabs.update(tabId, { active: true });
  await ensureContentScript(tabId);
  const resp = await chrome.tabs.sendMessage(tabId, {
    type: "interact",
    action: msg.action,
    ref: msg.ref,
    text: msg.text,
    direction: msg.direction,
    key: msg.key,
    option: msg.option
  });
  if (!resp || !resp.ok) {
    throw new Error((resp && resp.error) || "interact 失败");
  }
  const tab = await chrome.tabs.get(tabId);
  return {
    url: tab.url || "",
    title: tab.title || "",
    message: resp.message || "ok",
    probe: resp.probe || null
  };
}

async function downloadTab(msg) {
  const tabId = parseTabId(msg.tab_uuid);
  await chrome.tabs.update(tabId, { active: true });

  if (msg.url) {
    const downloadId = await chrome.downloads.download({
      url: msg.url,
      conflictAction: "uniquify",
      saveAs: false
    });
    const file = await waitDownload(downloadId, 30000);
    return file;
  }

  await ensureContentScript(tabId);
  const prep = await chrome.tabs.sendMessage(tabId, {
    type: "download_prep",
    ref: msg.ref
  });
  if (!prep || !prep.ok) {
    throw new Error((prep && prep.error) || "无法定位下载目标");
  }

  if (prep.href) {
    const downloadId = await chrome.downloads.download({
      url: prep.href,
      conflictAction: "uniquify",
      saveAs: false
    });
    const file = await waitDownload(downloadId, 30000);
    file.probe = prep.probe || null;
    return file;
  }

  const downloadPromise = waitNextDownload(30000);
  const click = await chrome.tabs.sendMessage(tabId, {
    type: "download_click",
    ref: msg.ref
  });
  if (!click || !click.ok) {
    throw new Error((click && click.error) || "点击下载失败");
  }
  const file = await downloadPromise;
  file.probe = prep.probe || null;
  return file;
}

function waitDownload(downloadId, timeoutMs) {
  return new Promise((resolve, reject) => {
    const t0 = Date.now();
    const onChanged = (delta) => {
      if (delta.id !== downloadId) {
        return;
      }
      if (delta.state && delta.state.current === "complete") {
        chrome.downloads.onChanged.removeListener(onChanged);
        chrome.downloads.search({ id: downloadId }).then((items) => {
          const it = items && items[0];
          if (!it || !it.filename) {
            reject(new Error("下载完成但无文件路径"));
            return;
          }
          resolve({
            filename: it.filename,
            bytes: it.fileSize || 0,
            content_type: it.mime || null,
            source_url: it.url || null
          });
        }).catch(reject);
      } else if (delta.state && delta.state.current === "interrupted") {
        chrome.downloads.onChanged.removeListener(onChanged);
        reject(new Error("下载中断"));
      } else if (Date.now() - t0 > timeoutMs) {
        chrome.downloads.onChanged.removeListener(onChanged);
        reject(new Error("下载超时"));
      }
    };
    chrome.downloads.onChanged.addListener(onChanged);
    setTimeout(() => {
      chrome.downloads.onChanged.removeListener(onChanged);
      reject(new Error("下载超时"));
    }, timeoutMs + 50);
  });
}

function waitNextDownload(timeoutMs) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      chrome.downloads.onCreated.removeListener(onCreated);
      reject(new Error("等待下载超时"));
    }, timeoutMs);

    const onCreated = (item) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      chrome.downloads.onCreated.removeListener(onCreated);
      waitDownload(item.id, timeoutMs).then(resolve, reject);
    };
    chrome.downloads.onCreated.addListener(onCreated);
  });
}

chrome.tabs.onCreated.addListener(() => pushTabsSnapshot());
chrome.tabs.onRemoved.addListener(() => pushTabsSnapshot());
chrome.tabs.onUpdated.addListener(() => pushTabsSnapshot());
chrome.tabs.onActivated.addListener(() => pushTabsSnapshot());

chrome.runtime.onStartup.addListener(connectNative);
chrome.runtime.onInstalled.addListener(connectNative);
connectNative();
