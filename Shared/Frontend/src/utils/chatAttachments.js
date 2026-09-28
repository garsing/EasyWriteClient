export const MAX_CHAT_ATTACHMENTS = 10
export const CHAT_ATTACHMENTS_MARKER = '【对话附件】'

export function buildChatAttachmentsAppendix (items) {
  const ready = (items || []).filter((x) => x.phase === 'ready' && x.storageDocUuid)
  if (!ready.length) return ''
  const lines = ready.map((x) =>
    `storage_doc_uuid=${x.storageDocUuid}\nfile_name=${x.fileLabel?.name || '文件'}`
  )
  return `\n\n${CHAT_ATTACHMENTS_MARKER}\n` + lines.join('\n')
}

export function appendChatAttachmentsToUserContent (raw, items) {
  const text = raw || ''
  const appendix = buildChatAttachmentsAppendix(items)
  return appendix ? text + appendix : text
}

export function stripChatAttachmentsAppendix (content) {
  if (!content || typeof content !== 'string') return content || ''
  const idx = content.indexOf(`\n\n${CHAT_ATTACHMENTS_MARKER}`)
  if (idx >= 0) return content.slice(0, idx).trimEnd()
  const idx2 = content.indexOf(CHAT_ATTACHMENTS_MARKER)
  if (idx2 >= 0) return content.slice(0, idx2).trimEnd()
  return content
}

const VISION_DISPLAY_DROP = new Set([
  '[图片附件已用于视觉分析]',
  '[图片附件已用于视觉分析，图像未持久化]',
  '[页面截图已落盘]',
  '[文件图片已落盘]',
  '[页面截图已用于视觉分析，图像未持久化]',
  '[文件图片已用于视觉分析，图像未持久化]',
  '[vision_turn_consumed]'
])

/** 历史/展示：去掉落库用的 consume / path，不进气泡。 */
export function stripVisionPersistNotes (content) {
  if (!content || typeof content !== 'string') return content || ''
  const kept = content.split(/\r?\n/).filter((line) => {
    const t = line.trim().replace(/^-\s*/, '')
    if (!t) return true
    if (VISION_DISPLAY_DROP.has(t)) return false
    if (/^path:\s*conversation_vision\//i.test(t)) return false
    return true
  })
  return kept.join('\n').replace(/\n{3,}/g, '\n\n').trim()
}

export function snapshotAttachments (items) {
  return (items || [])
    .filter((x) => x.phase === 'ready' && x.storageDocUuid)
    .map((x) => {
      const out = {
        storage_doc_uuid: x.storageDocUuid,
        fileName: x.fileLabel?.name || '文件',
        ext: x.fileLabel?.ext || '',
        sizeText: x.fileLabel?.sizeText || ''
      }
      if (x.knowledgeBaseUuid) {
        out.knowledge_base_uuid = x.knowledgeBaseUuid
      }
      return out
    })
}

export function normalizeHistoryAttachments (raw) {
  if (Array.isArray(raw) && raw.length) return raw
  return undefined
}
