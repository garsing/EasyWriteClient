import { getImageUrl, getStorageVisionImageUrl } from '../services/knowledgeBaseApi.js'

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
  if (!storageDocUuid) {
    throw new Error('缺少 storage_doc_uuid')
  }
  try {
    return await getStorageVisionImageUrl(storageDocUuid)
  } catch (e) {
    if (!knowledgeBaseUuid) throw e
    return getImageUrl(visionSidecarRelPath(storageDocUuid), knowledgeBaseUuid)
  }
}
