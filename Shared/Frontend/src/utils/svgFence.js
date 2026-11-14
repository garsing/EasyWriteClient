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

function svgFences (text) {
  return scanFences(text).filter((fence) => fence.lang === 'svg')
}

export function stripSvgFences (text) {
  const src = normalizeNewlines(text)
  const fences = svgFences(src)
  if (!fences.length) return src.trim().replace(/\n{3,}/g, '\n\n')
  let out = ''
  let cursor = 0
  for (const fence of fences) {
    out += src.slice(cursor, fence.start)
    cursor = fence.end
  }
  out += src.slice(cursor)
  return out.trim().replace(/\n{3,}/g, '\n\n')
}

/**
 * 闭合且 sanitize(inner) 有结果的 svg 围栏换成单行占位。
 * 未闭合、非 svg、未通过门禁的原文留下。
 */
export function prepareSvgFences (text, sanitize) {
  const src = normalizeNewlines(text)
  const fences = svgFences(src)
  const slots = []
  let out = ''
  let cursor = 0
  for (const fence of fences) {
    if (!fence.closed) {
      out += src.slice(cursor)
      cursor = src.length
      break
    }
    const safe = sanitize(fence.inner)
    if (!safe) continue
    const id = slots.length
    slots.push(safe)
    out += src.slice(cursor, fence.start)
    out += `\n\n<div class="md-svg" data-svg-slot="${id}"></div>\n\n`
    cursor = fence.end
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
