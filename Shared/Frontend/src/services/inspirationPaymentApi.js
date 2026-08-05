/**
 * 灵感值支付 API（设置页 Web 直调 Backend）
 */
import { buildHeaders } from './knowledgeBaseApi'

async function apiRequest (method, endpoint, body = null) {
  const { headers, baseUrl } = await buildHeaders(body !== null)
  const url = `${baseUrl}${endpoint}`
  const options = { method, headers }
  if (body !== null) {
    options.body = JSON.stringify(body)
  }
  const response = await fetch(url, options)
  let responseData
  try {
    responseData = await response.json()
  } catch {
    throw new Error(`${method} ${endpoint} 响应无效`)
  }
  if (response.ok && responseData.success) {
    return responseData.data
  }
  const detail = responseData.detail || responseData.message
  const msg = typeof detail === 'string' ? detail : (detail?.message || `${method} ${endpoint} 失败`)
  throw new Error(msg)
}

/** @param {{ productType: string, planCode?: string, amountYuan?: number, channel?: string }} opts */
export function createInspirationOrder ({ productType, planCode, amountYuan, channel = 'alipay_page' }) {
  const body = { product_type: productType, channel }
  if (planCode) body.plan_code = planCode
  if (amountYuan != null) body.amount_yuan = amountYuan
  return apiRequest('POST', '/payments/inspiration/orders', body)
}

export function getInspirationOrder (orderId) {
  return apiRequest('GET', `/payments/inspiration/orders/${orderId}`)
}

export function confirmInspirationPaid (orderId) {
  return apiRequest('POST', `/payments/inspiration/orders/${orderId}/confirm-paid`, {})
}
