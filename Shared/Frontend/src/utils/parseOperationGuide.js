/**
 * 解析操作指南：
 * - {{icon:filename}} 内联图标（I1；含多余属性则跳过）
 * - 行首 `- ` / `* ` / `• ` 为 bullet list
 *
 * @param {string} text
 * @returns {Array<
 *   | { type: 'paragraph', parts: Array<{ type: 'text', value: string } | { type: 'icon', filename: string }> }
 *   | { type: 'list', items: Array<Array<{ type: 'text', value: string } | { type: 'icon', filename: string }>> }
 * >}
 */
export function parseOperationGuide (text) {
  if (!text || typeof text !== 'string') return []

  const blocks = []
  let currentList = null

  for (const rawLine of text.split(/\r?\n/)) {
    const trimmed = rawLine.trim()
    if (!trimmed) {
      currentList = null
      continue
    }

    const bulletMatch = /^[-*•]\s+/.exec(trimmed)
    if (bulletMatch) {
      const content = trimmed.slice(bulletMatch[0].length)
      const parts = parseInlineParts(content)
      if (!parts.length) continue
      if (!currentList) {
        currentList = { type: 'list', items: [] }
        blocks.push(currentList)
      }
      currentList.items.push(parts)
    } else {
      currentList = null
      const parts = parseInlineParts(trimmed)
      if (parts.length) {
        blocks.push({ type: 'paragraph', parts })
      }
    }
  }

  return blocks
}

/**
 * @param {string} text
 * @returns {Array<{ type: 'text', value: string } | { type: 'icon', filename: string }>}
 */
function parseInlineParts (text) {
  if (!text) return []

  const re = /\{\{([^}]+)\}\}/g
  const parts = []
  let lastIndex = 0
  let match

  while ((match = re.exec(text)) !== null) {
    if (match.index > lastIndex) {
      parts.push({ type: 'text', value: text.slice(lastIndex, match.index) })
    }
    const inner = match[1].trim()
    const iconMatch = /^icon:([^\s]+)$/.exec(inner)
    if (iconMatch) {
      const filename = iconMatch[1]
      if (filename && !filename.includes('/') && !filename.includes('\\') && !filename.includes('..')) {
        parts.push({ type: 'icon', filename })
      } else {
        console.warn('[parseOperationGuide] 非法图标名，跳过:', inner)
      }
    } else {
      console.warn('[parseOperationGuide] 未知 token，跳过:', inner)
    }
    lastIndex = match.index + match[0].length
  }

  if (lastIndex < text.length) {
    parts.push({ type: 'text', value: text.slice(lastIndex) })
  }

  return parts
}

/** 收集所有 icon 文件名（供加载用） */
export function collectGuideIconFilenames (blocks) {
  const names = new Set()
  for (const block of blocks || []) {
    if (block.type === 'paragraph') {
      for (const p of block.parts || []) {
        if (p.type === 'icon') names.add(p.filename)
      }
    } else if (block.type === 'list') {
      for (const item of block.items || []) {
        for (const p of item || []) {
          if (p.type === 'icon') names.add(p.filename)
        }
      }
    }
  }
  return [...names]
}
