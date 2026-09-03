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
