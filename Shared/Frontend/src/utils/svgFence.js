/**
 * 助手正文里的 svg 围栏。
 * 闭合：另有一行，反引号不少于开头，且该行只有反引号。文末不算闭合。
 * 信息字符串的第一个词是 svg（大小写不区分）才算 svg 围栏。
 */

const FENCE_OPEN = /^( {0,3})(`{3,})([^`]*)$/
const FENCE_CLOSE = /^( {0,3})(`{3,})[ \t]*$/

function normalizeNewlines (text) {
  return String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n')
}

function lineSpans (text) {
  const spans = []
  let start = 0
  for (let i = 0; i < text.length; i++) {
    if (text[i] === '\n') {
      spans.push([start, i])
      start = i + 1
    }
  }
  spans.push([start, text.length])
  return spans
}

function matchOpen (line) {
  const matched = FENCE_OPEN.exec(line)
  if (!matched) return null
  const info = matched[3].trim()
  const first = info ? info.split(/[ \t]+/)[0] : ''
  return {
    ticks: matched[2].length,
    lang: first.toLowerCase()
  }
}

function matchClose (line, ticks) {
  const matched = FENCE_CLOSE.exec(line)
  return !!(matched && matched[2].length >= ticks)
}

export function scanFences (text) {
  const src = normalizeNewlines(text)
  const spans = lineSpans(src)
  const fences = []
  let lineIdx = 0
  while (lineIdx < spans.length) {
    const [lineStart, lineEnd] = spans[lineIdx]
    const open = matchOpen(src.slice(lineStart, lineEnd))
    if (!open) {
      lineIdx += 1
      continue
    }
    let closeLine = -1
    for (let j = lineIdx + 1; j < spans.length; j++) {
      const [closeStart, closeEnd] = spans[j]
      if (matchClose(src.slice(closeStart, closeEnd), open.ticks)) {
        closeLine = j
        break
      }
    }
    if (closeLine >= 0) {
      const innerStart = lineIdx + 1 < spans.length ? spans[lineIdx + 1][0] : src.length
      const closeLineStart = spans[closeLine][0]
      const end = closeLine + 1 < spans.length ? spans[closeLine + 1][0] : src.length
      fences.push({
        start: lineStart,
        end,
        closed: true,
        lang: open.lang,
        inner: src.slice(innerStart, closeLineStart).replace(/\n$/, '')
      })
      lineIdx = closeLine + 1
    } else {
      const innerStart = lineIdx + 1 < spans.length ? spans[lineIdx + 1][0] : src.length
      fences.push({
        start: lineStart,
        end: src.length,
        closed: false,
        lang: open.lang,
        inner: src.slice(innerStart)
      })
      break
    }
  }
  return fences
}

function isNameBoundary (text, index) {
  if (index >= text.length) return true
  return /[\s/>]/.test(text[index])
}

function endOfTag (text, start) {
  let quote = ''
  for (let i = start; i < text.length; i++) {
    const ch = text[i]
    if (quote) {
      if (ch === quote) quote = ''
      continue
    }
    if (ch === '"' || ch === "'") {
      quote = ch
      continue
    }
    if (ch === '>') return i + 1
  }
  return -1
}

function findSvgOpen (text, from) {
  const lower = text.toLowerCase()
  let i = from
  while (i < text.length) {
    const at = lower.indexOf('<svg', i)
    if (at < 0) return -1
    if (isNameBoundary(text, at + 4)) return at
    i = at + 4
  }
  return -1
}

/** 从 `<svg` 起找到配对的结束位置。没闭合返回 -1。 */
function findSvgClose (text, start) {
  const lower = text.toLowerCase()
  let i = start
  let depth = 0
  while (i < text.length) {
    if (text.startsWith('<!--', i)) {
      const end = text.indexOf('-->', i + 4)
      if (end < 0) return -1
      i = end + 3
      continue
    }
    if (text.startsWith('<![CDATA[', i)) {
      const end = text.indexOf(']]>', i + 9)
      if (end < 0) return -1
      i = end + 3
      continue
    }
    if (lower.startsWith('<svg', i) && isNameBoundary(text, i + 4)) {
      const tagEnd = endOfTag(text, i)
      if (tagEnd < 0) return -1
      const selfClose = /\/\s*>$/.test(text.slice(i, tagEnd))
      if (selfClose) {
        if (depth === 0) return tagEnd
      } else {
        depth += 1
      }
      i = tagEnd
      continue
    }
    if (lower.startsWith('</svg', i) && isNameBoundary(text, i + 5)) {
      const tagEnd = endOfTag(text, i)
      if (tagEnd < 0) return -1
      depth -= 1
      i = tagEnd
      if (depth === 0) return i
      continue
    }
    i += 1
  }
  return -1
}

function svgPieces (text) {
  const src = normalizeNewlines(text)
  const fences = scanFences(src)
  const pieces = []
  for (const fence of fences) {
    if (fence.lang !== 'svg') continue
    pieces.push({
      start: fence.start,
      end: fence.end,
      closed: fence.closed,
      inner: fence.inner
    })
  }
  let i = 0
  while (i < src.length) {
    const covering = fences.find((fence) => i >= fence.start && i < fence.end)
    if (covering) {
      i = covering.end
      continue
    }
    const start = findSvgOpen(src, i)
    if (start < 0) break
    const inside = fences.find((fence) => start >= fence.start && start < fence.end)
    if (inside) {
      i = inside.end
      continue
    }
    const end = findSvgClose(src, start)
    if (end < 0) {
      pieces.push({
        start,
        end: src.length,
        closed: false,
        inner: src.slice(start)
      })
      break
    }
    pieces.push({
      start,
      end,
      closed: true,
      inner: src.slice(start, end)
    })
    i = end
  }
  pieces.sort((a, b) => a.start - b.start)
  return { src, pieces }
}

function removePieces (src, pieces) {
  if (!pieces.length) return src.trim().replace(/\n{3,}/g, '\n\n')
  let out = ''
  let cursor = 0
  for (const piece of pieces) {
    out += src.slice(cursor, piece.start)
    cursor = piece.end
  }
  out += src.slice(cursor)
  return out.trim().replace(/\n{3,}/g, '\n\n')
}

export function stripSvgFences (text) {
  const { src, pieces } = svgPieces(text)
  return removePieces(src, pieces)
}

const PENDING_HTML = '\n\n<div class="md-svg-pending">正在画图</div>\n\n'

/**
 * 通过门禁的完整 SVG 换成图片占位。
 * 还没闭合时换成「正在画图」，不把源码留在正文里。
 * pending 为 false 时，没画成的片段直接拿掉，仍不露出源码。
 */
export function prepareSvgFences (text, sanitize, pending = false) {
  const { src, pieces } = svgPieces(text)
  const slots = []
  let out = ''
  let cursor = 0
  for (const piece of pieces) {
    out += src.slice(cursor, piece.start)
    if (!piece.closed) {
      if (pending) out += PENDING_HTML
      cursor = piece.end
      break
    }
    const safe = sanitize(piece.inner)
    if (!safe) {
      cursor = piece.end
      continue
    }
    const id = slots.length
    slots.push(safe)
    out += `\n\n<div class="md-svg" data-svg-slot="${id}"></div>\n\n`
    cursor = piece.end
  }
  out += src.slice(cursor)
  return { text: out, slots }
}

export function fillSvgSlots (html, slots) {
  return String(html || '').replace(
    /<div class="md-svg" data-svg-slot="(\d+)"\s*>\s*<\/div>/g,
    (full, id) => {
      const svg = slots[Number(id)]
      if (!svg) return full
      return `<div class="md-svg">${svg}</div>`
    }
  )
}
