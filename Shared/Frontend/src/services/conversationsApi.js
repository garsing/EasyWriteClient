/**
 * 任务（会话）列表 / 详情 — Desktop 侧栏用，经 getApiConfig 鉴权直连后端。
 */
import { buildHeaders } from './knowledgeBaseApi.js'

export const CONVERSATION_PAGE_SIZE = 50

export async function listConversations ({ page = 1, pageSize = CONVERSATION_PAGE_SIZE } = {}) {
  const { headers, baseUrl } = await buildHeaders(true)
  const qs = new URLSearchParams({
    sort_by: 'last_message_at',
    sort_order: 'desc',
    page: String(page),
    page_size: String(pageSize)
  })
  const url = `${baseUrl.replace(/\/$/, '')}/conversations/?${qs.toString()}`
  const res = await fetch(url, { method: 'GET', headers })
  if (!res.ok) {
    const text = await res.text().catch(() => '')
    throw new Error(`获取任务列表失败: HTTP ${res.status} ${text}`)
  }
  const data = await res.json()
  const conversations = Array.isArray(data?.conversations) ? data.conversations : []
  const total = Number(data?.total) || 0
  const currentPage = Number(data?.page) || page
  const totalPages = Number(data?.total_pages) || 0
  return {
    conversations,
    total,
    page: currentPage,
    pageSize: Number(data?.page_size) || pageSize,
    totalPages,
    hasMore: totalPages > 0 ? currentPage < totalPages : conversations.length >= pageSize
  }
}

/**
 * 后端会话时间多为 utcnow() 的 naive ISO（无时区后缀）。
 * 浏览器会把无后缀的 ISO 当成本地时间，中国区会固定偏早 8 小时 → 「刚提问却显示 8 小时前」。
 * 无时区时按 UTC 解析。
 */
function parseApiDateMs (isoOrLocal) {
  if (!isoOrLocal) return NaN
  let s = String(isoOrLocal).trim().replace(' ', 'T')
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(s) && !/[zZ]|[+-]\d{2}:?\d{2}$/.test(s)) {
    s += 'Z'
  }
  return Date.parse(s)
}

/**
 * 相对时间展示（侧栏副标题）。
 */
export function formatRelativeTime (isoOrLocal) {
  if (!isoOrLocal) return ''
  const t = parseApiDateMs(isoOrLocal)
  if (Number.isNaN(t)) return String(isoOrLocal)
  const diffMs = Date.now() - t
  const sec = Math.floor(diffMs / 1000)
  if (sec < 60) return '刚刚'
  const min = Math.floor(sec / 60)
  if (min < 60) return `${min} 分钟前`
  const hr = Math.floor(min / 60)
  if (hr < 24) return `${hr} 小时前`
  const day = Math.floor(hr / 24)
  if (day < 30) return `${day} 天前`
  try {
    return new Date(t).toLocaleDateString('zh-CN')
  } catch {
    return ''
  }
}
