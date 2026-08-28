/**
 * WebView2 通信桥接
 * 用于 C# 和 Vue 之间的双向通信
 */

let messageHandlers = []
const pendingMessages = new Map() // 存储待响应的消息

// 检查是否在 WebView2 环境中
const isWebView2 = () => {
  return typeof window.chrome !== 'undefined' && 
         window.chrome.webview !== undefined &&
         typeof window.chrome.webview.postMessage === 'function'
}

// 处理消息响应（由全局消息监听器调用）
const settlePending = (messageId, pending, responseData) => {
  clearTimeout(pending.timeout)
  pendingMessages.delete(messageId)

  if (responseData.data && responseData.data.success !== undefined) {
    pending.resolve(responseData.data)
  } else {
    pending.resolve(responseData.data || { success: true })
  }
}

const handleMessageResponse = (responseData) => {
  if (responseData.type !== 'messageResponse') {
    return
  }

  if (responseData.messageId && pendingMessages.has(responseData.messageId)) {
    settlePending(responseData.messageId, pendingMessages.get(responseData.messageId), responseData)
    return
  }

  // 兼容未回传 messageId 的旧宿主；同类型并发仍可能对串
  if (responseData.originalType) {
    for (const [messageId, pending] of pendingMessages.entries()) {
      if (pending.type === responseData.originalType) {
        settlePending(messageId, pending, responseData)
        return
      }
    }
  }
}

// 发送消息到 C# 并等待响应
export const sendMessage = (type, data = {}) => {
  return new Promise((resolve, reject) => {
    if (!isWebView2()) {
      console.warn('不在 WebView2 环境中，无法发送消息')
      // 开发模式下，模拟成功
      if (import.meta.env.DEV) {
        console.log('[DEV] 模拟发送消息:', type, data)
        resolve({ success: true })
        return
      }
      reject(new Error('不在 WebView2 环境中'))
      return
    }

    const messageId = `${type}_${Date.now()}_${Math.random()}`
    const message = {
      type,
      data,
      timestamp: Date.now(),
      messageId // 添加消息ID用于匹配响应
    }

    // 设置超时
    const timeout = setTimeout(() => {
      pendingMessages.delete(messageId)
      reject(new Error('消息响应超时'))
    }, 30000) // 30秒超时

    // 存储待响应的消息
    pendingMessages.set(messageId, { resolve, reject, timeout, type })

    try {
      window.chrome.webview.postMessage(JSON.stringify(message))
    } catch (error) {
      pendingMessages.delete(messageId)
      clearTimeout(timeout)
      console.error('发送消息失败:', error)
      reject(error)
    }
  })
}

// 监听来自 C# 的消息
export const onMessage = (handler) => {
  if (!isWebView2()) {
    console.warn('不在 WebView2 环境中，无法监听消息')
    // 开发模式下，设置一个模拟监听器
    if (import.meta.env.DEV) {
      console.log('[DEV] 模拟消息监听器已注册')
      // 返回一个空的取消函数
      return () => {}
    }
    return () => {}
  }

  const wrappedHandler = (event) => {
    try {
      console.log('[useWebViewBridge] 收到消息事件:', event)
      let data
      if (typeof event === 'string') {
        // 直接是字符串
        console.log('[useWebViewBridge] 消息是字符串，解析 JSON')
        data = JSON.parse(event)
      } else if (event.data) {
        // 事件对象
        console.log('[useWebViewBridge] 消息有 data 属性:', typeof event.data)
        data = typeof event.data === 'string' ? JSON.parse(event.data) : event.data
      } else {
        // 直接是对象
        console.log('[useWebViewBridge] 消息是对象')
        data = event
      }
      
      console.log('[useWebViewBridge] 解析后的数据:', data)
      
      // 先处理消息响应（如果有）
      // C# 发送的格式是 { type: "messageResponse", data: { type, originalType, data, timestamp } }
      if (data.type === 'messageResponse' && data.data) {
        console.log('[useWebViewBridge] 检测到消息响应，originalType:', data.data.originalType)
        // 提取嵌套的响应数据
        const responseData = {
          type: data.data.type || 'messageResponse',
          originalType: data.data.originalType,
          messageId: data.data.messageId || data.messageId,
          data: data.data.data || data.data,
          timestamp: data.data.timestamp || data.timestamp
        }
        console.log('[useWebViewBridge] 提取的响应数据:', responseData)
        handleMessageResponse(responseData)
      }
      
      // 然后调用用户提供的处理器
      handler(data)
    } catch (error) {
      console.error('[useWebViewBridge] 解析消息失败:', error, event)
    }
  }

  // WebView2 使用 window.chrome.webview.addEventListener
  if (window.chrome.webview.addEventListener) {
    window.chrome.webview.addEventListener('message', wrappedHandler)
  } else {
    // 备用方案：使用全局消息事件
    window.addEventListener('message', (e) => {
      if (e.source === window) {
        wrappedHandler(e)
      }
    })
  }
  
  messageHandlers.push({ handler, wrappedHandler })

  // 返回取消监听的函数
  return () => {
    if (window.chrome.webview && window.chrome.webview.removeEventListener) {
      window.chrome.webview.removeEventListener('message', wrappedHandler)
    }
    messageHandlers = messageHandlers.filter(h => h.handler !== handler)
  }
}

// 组合式函数
export const useWebViewBridge = () => {
  return {
    sendMessage,
    onMessage,
    isWebView2: isWebView2()
  }
}

