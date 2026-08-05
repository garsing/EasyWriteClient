/**
 * 聊天空状态：操作指南 + 输入推荐 + 图标（公开接口，无需登录）
 */
import { buildHeaders } from './knowledgeBaseApi.js'

/**
 * @param {boolean} documentEmpty
 * @returns {Promise<{ operation_guide: string, recommendations: string[], document_empty: boolean }>}
 */
export async function fetchChatEmptyState (documentEmpty) {
  const { headers, baseUrl } = await buildHeaders()
  const qs = new URLSearchParams({
    document_empty: documentEmpty ? 'true' : 'false'
  })
  const res = await fetch(`${baseUrl}/chat-empty-state?${qs}`, {
    method: 'GET',
    headers
  })
  const body = await res.json().catch(() => ({}))
  if (!res.ok || !body.success) {
    throw new Error(body.message || body.detail || `GET /chat-empty-state 失败 (${res.status})`)
  }
  return body.data || { operation_guide: '', recommendations: [], document_empty: documentEmpty }
}

/**
 * 拉取图标，返回 blob URL；失败抛错（调用方按 I3 跳过）。
 * @param {string} filename
 * @returns {Promise<string>}
 */
export async function fetchChatEmptyStateIconUrl (filename) {
  const { headers, baseUrl } = await buildHeaders(false)
  const url = `${baseUrl}/chat-empty-state/icons/${encodeURIComponent(filename)}`
  const res = await fetch(url, { method: 'GET', headers })
  if (!res.ok) {
    throw new Error(`图标加载失败: ${filename} (${res.status})`)
  }
  const blob = await res.blob()
  return URL.createObjectURL(blob)
}
