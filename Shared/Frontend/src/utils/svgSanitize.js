/**
 * 聊天面板 SVG 门禁，拒绝项对齐 SvgSecurityValidator。
 * 属性名上的 \son 会命中 opacity，这里只把 opacity 留作普通属性，其余 on* 仍拒绝。
 */

const MAX_BYTES = 512 * 1024
const EVENT_ATTR = /\son(?!opacity\b)[a-zA-Z]+\s*=/i
const EXTERNAL_URL = /(?:href|xlink:href)\s*=\s*["']\s*(?:https?:|\/\/|javascript:)/i
const EXTERNAL_VALUE = /^\s*(?:https?:|\/\/|javascript:)/i

function byteLength (text) {
  if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(text).length
  return unescape(encodeURIComponent(text)).length
}

function isEventAttr (name) {
  const normalized = String(name || '').toLowerCase()
  if (normalized === 'opacity') return false
  return /^on[a-z]/.test(normalized)
}

function isExternalHref (name, value) {
  const normalized = String(name || '').toLowerCase()
  const isHref = normalized === 'href' || normalized === 'xlink:href' || normalized.endsWith(':href')
  if (!isHref) return false
  return EXTERNAL_VALUE.test(String(value || ''))
}

function precheck (trimmed) {
  if (!trimmed) return false
  if (byteLength(trimmed) > MAX_BYTES) return false
  const lower = trimmed.toLowerCase()
  if (lower.includes('<script') || lower.includes('</script')) return false
  if (lower.includes('<foreignobject') || lower.includes('<iframe')) return false
  if (EVENT_ATTR.test(trimmed)) return false
  if (EXTERNAL_URL.test(trimmed) || lower.includes('javascript:')) return false
  if (lower.includes('<!doctype') || lower.includes('<!entity')) return false
  return true
}

function domRejects (root) {
  const elements = [root, ...root.querySelectorAll('*')]
  for (const el of elements) {
    const tag = (el.localName || '').toLowerCase()
    if (tag === 'script' || tag === 'foreignobject' || tag === 'iframe' || tag === 'parsererror') {
      return true
    }
    for (const attr of [...el.attributes]) {
      if (isEventAttr(attr.name) || isEventAttr(attr.localName)) return true
      if (isExternalHref(attr.name, attr.value) || isExternalHref(attr.localName, attr.value)) return true
    }
  }
  return false
}

/** 通过时返回可插入的根 svg 字符串，否则返回空字符串。 */
export function sanitizeSvg (svg) {
  const trimmed = String(svg || '').trim()
  if (!precheck(trimmed)) return ''
  if (typeof DOMParser === 'undefined' || typeof XMLSerializer === 'undefined') return ''
  let doc
  try {
    doc = new DOMParser().parseFromString(trimmed, 'image/svg+xml')
  } catch {
    return ''
  }
  const root = doc && doc.documentElement
  if (!root || (root.localName || '').toLowerCase() !== 'svg') return ''
  if (doc.querySelector('parsererror')) return ''
  if (domRejects(root)) return ''
  try {
    return new XMLSerializer().serializeToString(root)
  } catch {
    return ''
  }
}
