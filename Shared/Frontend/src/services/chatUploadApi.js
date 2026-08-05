/**
 * 对话框上传：GET /chat-upload/by-document/{storage_doc_uuid}（含 409 退避重试）
 */
import { buildHeaders } from './knowledgeBaseApi.js'

/**
 * @param {string} storageDocUuid
 * @returns {Promise<{ upload_id: string, storage_doc_uuid?: string, [k: string]: any }>}
 */
export async function getChatUploadByDocument (storageDocUuid) {
  const delays = [300, 600, 1200, 2400, 2400]
  const maxAttempts = delays.length

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const { headers, baseUrl } = await buildHeaders()
    const url = `${baseUrl}/chat-upload/by-document/${encodeURIComponent(storageDocUuid)}`

    const res = await fetch(url, {
      method: 'GET',
      headers
    })
    const body = await res.json().catch(() => ({}))

    if (res.ok && body.success && body.data && body.data.upload_id) {
      return body.data
    }

    if (res.status === 409 && attempt < maxAttempts - 1) {
      await new Promise((r) => setTimeout(r, delays[attempt]))
      continue
    }

    const msg = body.message || body.detail || `GET by-document 失败 (${res.status})`
    throw new Error(msg)
  }

  throw new Error('GET by-document：重试次数已用尽')
}
