/** Desktop「打开文件」选中集合：拼接 / 剥离 / prune（见实现文档 I15） */

export const MAX_SELECTED_OPEN_FILES = 15
export const SELECTED_OPEN_FILES_MARKER = '【用户已选中的打开文件 / 渠道】'

export function openFileSelectionKey (item) {
  if (item?.id != null && String(item.id).length > 0) return String(item.id)
  const type = item?.appType || item?.app_type || ''
  const name = item?.fullPath || item?.full_path || item?.displayName || item?.display_name || ''
  return `${type}|${name}`
}

export function toSelectedOpenFile (item) {
  return {
    id: openFileSelectionKey(item),
    displayName: item.displayName || item.display_name || '未命名文档',
    appType: String(item.appType || item.app_type || 'word').toLowerCase(),
    channelId: item.channelId || item.channel_id || '',
    fullPath: item.fullPath || item.full_path || ''
  }
}

export function buildSelectedOpenFilesAppendix (selected) {
  if (!selected?.length) return ''
  const lines = selected.map((s, i) => {
    const name = s.displayName || '未命名文档'
    const ch = (s.channelId || '').trim()
    // 有真实 channel_id（含 wps:）则写出；不再因 appType===wps 强制「无渠道」
    if (ch) return `${i + 1}. ${name} | channel_id=${ch}`
    return `${i + 1}. ${name} | 无操作渠道（渠道未建立）`
  })
  return (
    `\n\n${SELECTED_OPEN_FILES_MARKER}\n` +
    lines.join('\n') +
    '\n请优先针对上述有渠道的项理解与操作；若需切换默认渠道请使用工具显式指定。'
  )
}

export function appendToUserContent (raw, selected) {
  const text = (raw || '').trimEnd()
  const appendix = buildSelectedOpenFilesAppendix(selected)
  return appendix ? text + appendix : text
}

/** 历史/展示：去掉附加段，避免气泡泄露 channel_id */
export function stripSelectedOpenFilesAppendix (content) {
  if (!content || typeof content !== 'string') return content || ''
  const idx = content.indexOf(`\n\n${SELECTED_OPEN_FILES_MARKER}`)
  if (idx >= 0) return content.slice(0, idx).trimEnd()
  const idx2 = content.indexOf(SELECTED_OPEN_FILES_MARKER)
  if (idx2 === 0) return ''
  if (idx2 > 0) return content.slice(0, idx2).trimEnd()
  return content
}

export function pruneSelectionByOpenFiles (selected, openFiles) {
  const alive = new Set((openFiles || []).map(openFileSelectionKey))
  return (selected || []).filter((s) => alive.has(s.id))
}
