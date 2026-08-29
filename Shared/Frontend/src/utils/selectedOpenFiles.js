/** Desktop「打开文件」选中集合：拼接 / 剥离 / prune（见实现文档 I15） */
import { ref } from 'vue'

export const MAX_SELECTED_OPEN_FILES = 15
/** 模块级选中，避免 App 重挂载时被 ref([]) 冲掉；切会话不重置 */
export const selectedOpenFilesState = ref([])
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
    // 助手短号 w1/x1/p1/b1（与 system open_channels / 工具 channel_id 一致）
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
  const byKey = new Map()
  for (const item of openFiles || []) {
    byKey.set(openFileSelectionKey(item), item)
  }
  const next = []
  for (const s of selected || []) {
    const live = byKey.get(s.id)
    if (!live) continue
    const fresh = toSelectedOpenFile(live)
    next.push({
      ...s,
      displayName: fresh.displayName || s.displayName,
      appType: fresh.appType || s.appType,
      channelId: fresh.channelId,
      fullPath: fresh.fullPath || s.fullPath
    })
  }
  return next
}
