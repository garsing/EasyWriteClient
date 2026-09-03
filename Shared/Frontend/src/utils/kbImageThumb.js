import { getImageUrl } from '../services/knowledgeBaseApi.js'

export const IMAGE_DOCUMENT_EXTS = ['png', 'jpg', 'jpeg', 'webp', 'bmp', 'gif']

export function isImageFileName (fileName, ext) {
  const fromExt = String(ext || '').replace(/^\./, '').toLowerCase()
  const name = String(fileName || '')
  const fromName = name.includes('.') ? name.split('.').pop().toLowerCase() : ''
  return IMAGE_DOCUMENT_EXTS.includes(fromExt || fromName)
}

export function visionSidecarRelPath (storageDocUuid) {
  return `${storageDocUuid}/vision.jpeg`
}

export async function loadKbImageThumbUrl (storageDocUuid, knowledgeBaseUuid) {
  if (!storageDocUuid || !knowledgeBaseUuid) {
    throw new Error('缺少 storage_doc_uuid 或 knowledge_base_uuid')
  }
  return getImageUrl(visionSidecarRelPath(storageDocUuid), knowledgeBaseUuid)
}
