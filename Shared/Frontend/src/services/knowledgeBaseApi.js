/**
 * 知识库 API 服务
 * 前端直接调用后端 API，无需通过 C# WebViewBridge
 */

// API 配置缓存
let apiConfigCache = null

/** 登录/退出后清掉鉴权缓存，避免侧栏仍用旧 token。 */
export function clearApiConfigCache() {
  apiConfigCache = null
}

/**
 * 获取 API 配置（baseUrl 和认证头）
 * 未登录也可返回 baseUrl（headers 可为空）；无 Authorization 时不缓存，避免登录后仍用空头。
 */
async function getApiConfig() {
  if (apiConfigCache) {
    return apiConfigCache
  }
  
  // 首次调用需要通过 WebViewBridge 获取配置
  const { sendMessage } = await import('../composables/useWebViewBridge')
  try {
    const response = await sendMessage('getApiConfig', {})
    if (response.success && response.data) {
      // 仅在已带鉴权时缓存；未登录每次向宿主取最新配置
      if (response.data.headers?.Authorization) {
        apiConfigCache = response.data
      }
      return response.data
    }
    throw new Error(response.message || '获取 API 配置失败')
  } catch (error) {
    console.error('获取 API 配置失败:', error)
    throw error
  }
}

/**
 * 构建请求头（供其他模块复用，如 chat-upload）
 */
export async function buildHeaders(includeContentType = true) {
  const apiConfig = await getApiConfig()
  const { headers } = apiConfig
  
  const requestHeaders = new Headers()
  if (includeContentType) {
    requestHeaders.append('Content-Type', 'application/json')
  }
  if (headers.Authorization) {
    requestHeaders.append('Authorization', headers.Authorization)
  }
  if (headers['X-Username']) {
    requestHeaders.append('X-Username', headers['X-Username'])
  }
  
  return { headers: requestHeaders, baseUrl: apiConfig.baseUrl }
}

/**
 * 通用 API 请求方法
 */
async function apiRequest(method, endpoint, body = null) {
  const { headers, baseUrl } = await buildHeaders(body !== null)
  const url = `${baseUrl}${endpoint}`
  
  const options = {
    method,
    headers
  }
  
  if (body !== null) {
    options.body = JSON.stringify(body)
  }
  
  const response = await fetch(url, options)
  const responseData = await response.json()
  
  if (response.ok && responseData.success) {
    return responseData.data
  }
  
  throw new Error(responseData.message || responseData.detail || `${method} ${endpoint} 失败`)
}

/**
 * 获取知识库列表
 * @param {Object} params - 查询参数
 * @param {string} params.tab - 标签页类型 ('all' | 'favorites' | 'recent')
 * @param {string} params.search - 搜索关键词
 * @param {string} params.group - 分组筛选
 * @param {string} params.tag - 标签筛选
 * @param {string} params.status - 状态筛选
 * @returns {Promise<Array>} 知识库列表
 */
export async function getKnowledgeBaseList(params = {}) {
  try {
    const { baseUrl } = await buildHeaders()
    const queryParams = new URLSearchParams()
    
    if (params.tab) queryParams.append('tab', params.tab)
    if (params.search) queryParams.append('search', params.search)
    if (params.group) queryParams.append('group', params.group)
    if (params.tag) queryParams.append('tag', params.tag)
    if (params.status) queryParams.append('status', params.status)
    
    const queryString = queryParams.toString()
    const endpoint = `/knowledge/kb/list${queryString ? '?' + queryString : ''}`
    
    const data = await apiRequest('GET', endpoint)
    // API 返回格式: {"data": {"knowledge_bases": [...]}}
    // 需要提取 knowledge_bases 数组
    if (data.knowledge_bases && Array.isArray(data.knowledge_bases)) {
      return data.knowledge_bases
    }
    return Array.isArray(data) ? data : []
  } catch (error) {
    console.error('获取知识库列表失败:', error)
    throw error
  }
}

/**
 * 创建知识库
 * @param {string} name - 知识库名称（1-200字符）
 * @param {string} [groupUuid] - 分组UUID（可选），如果为null或undefined则不发送该字段
 * @returns {Promise<Object>} 创建的知识库信息
 */
export async function createKnowledgeBase(name, groupUuid = null) {
  try {
    const body = { name }
    // 只有当 groupUuid 是有效的字符串时才添加到请求体
    // 当 groupUuid 为 null、undefined、空字符串或非字符串类型时，不发送 group_uuid 字段
    if (groupUuid !== null && groupUuid !== undefined && groupUuid !== '' && typeof groupUuid === 'string') {
      body.group_uuid = groupUuid
    }
    console.log('[createKnowledgeBase] 请求体:', body)
    return await apiRequest('POST', '/knowledge/kb/create', body)
  } catch (error) {
    console.error('创建知识库失败:', error)
    throw error
  }
}

/**
 * 删除知识库
 * @param {string} uuid - 知识库UUID
 * @returns {Promise<void>}
 */
export async function deleteKnowledgeBase(uuid) {
  try {
    await apiRequest('DELETE', `/knowledge/kb/delete/${uuid}`)
  } catch (error) {
    console.error('删除知识库失败:', error)
    throw error
  }
}

/**
 * 更新知识库关注状态
 * @param {string} uuid - 知识库UUID
 * @param {boolean} isFollowed - 是否关注
 * @returns {Promise<Object>} 更新后的知识库信息
 */
export async function updateKnowledgeBaseFollow(uuid, isFollowed) {
  try {
    return await apiRequest('PUT', `/knowledge/kb/${uuid}/follow`, { is_followed: isFollowed })
  } catch (error) {
    console.error('更新关注状态失败:', error)
    throw error
  }
}

/**
 * 更新知识库标签
 * @param {string} uuid - 知识库UUID
 * @param {string|null} tags - 标签（单个标签，可为空字符串或null）
 * @returns {Promise<Object>} 更新后的知识库信息
 */
export async function updateKnowledgeBaseTags(uuid, tags) {
  try {
    return await apiRequest('PUT', `/knowledge/kb/${uuid}/tags`, { tags })
  } catch (error) {
    console.error('更新标签失败:', error)
    throw error
  }
}

/**
 * 获取知识库分组列表
 * @returns {Promise<Array>} 分组列表
 */
export async function getGroupList() {
  try {
    const data = await apiRequest('GET', '/knowledge/group/list')
    // API 返回格式: {"data": {"groups": [...]}}
    if (data.groups && Array.isArray(data.groups)) {
      return data.groups
    }
    return Array.isArray(data) ? data : []
  } catch (error) {
    console.error('获取分组列表失败:', error)
    throw error
  }
}

/**
 * 创建知识库分组
 * @param {string} groupName - 分组名称（1-200字符）
 * @returns {Promise<Object>} 创建的分组信息
 */
export async function createGroup(groupName) {
  try {
    return await apiRequest('POST', '/knowledge/group/create', { group_name: groupName })
  } catch (error) {
    console.error('创建分组失败:', error)
    throw error
  }
}

/**
 * 更新知识库分组名称
 * @param {string} groupUuid - 分组UUID
 * @param {string} groupName - 新的分组名称（1-200字符）
 * @returns {Promise<Object>} 更新后的分组信息
 */
export async function updateGroupName(groupUuid, groupName) {
  try {
    return await apiRequest('PUT', `/knowledge/group/${groupUuid}/name`, { group_name: groupName })
  } catch (error) {
    console.error('更新分组名称失败:', error)
    throw error
  }
}

/**
 * 删除知识库分组
 * @param {string} groupUuid - 分组UUID
 * @returns {Promise<void>}
 */
export async function deleteGroup(groupUuid) {
  try {
    await apiRequest('DELETE', `/knowledge/group/delete/${groupUuid}`)
  } catch (error) {
    console.error('删除分组失败:', error)
    throw error
  }
}

/**
 * 移动知识库到分组
 * @param {string} knowledgeBaseUuid - 知识库UUID
 * @param {string|null} groupUuid - 分组UUID（可为null，表示清空分组）
 * @returns {Promise<Object>} 更新后的知识库信息
 */
export async function moveKnowledgeBaseToGroup(knowledgeBaseUuid, groupUuid = null) {
  try {
    const body = {}
    // 如果 groupUuid 不为 null，才添加到请求体
    if (groupUuid !== null && groupUuid !== undefined && groupUuid !== '') {
      body.group_uuid = groupUuid
    }
    return await apiRequest('PUT', `/knowledge/kb/${knowledgeBaseUuid}/group_uuid`, body)
  } catch (error) {
    console.error('移动知识库失败:', error)
    throw error
  }
}

/**
 * 获取文档列表
 * @param {string} knowledgeBaseUuid - 知识库UUID
 * @param {number} startIndex - 起始序号（从0开始，默认为0）
 * @param {number} endIndex - 结束序号（不包含，如果为空则返回从startIndex开始的所有文档）
 * @returns {Promise<Object>} 文档列表数据
 */
export async function getDocumentList(knowledgeBaseUuid, startIndex = 0, endIndex = null) {
  try {
    const queryParams = new URLSearchParams()
    queryParams.append('start_index', startIndex.toString())
    if (endIndex !== null && endIndex !== undefined) {
      queryParams.append('end_index', endIndex.toString())
    }
    
    const endpoint = `/knowledge/document/list/${knowledgeBaseUuid}?${queryParams.toString()}`
    console.log('[getDocumentList] 请求文档列表:', endpoint)
    const result = await apiRequest('GET', endpoint)
    console.log('[getDocumentList] 文档列表响应:', result)
    return result
  } catch (error) {
    console.error('获取文档列表失败:', error)
    throw error
  }
}

/**
 * 上传文档到知识库（前端直接上传）
 * @param {string} knowledgeBaseUuid - 知识库UUID
 * @param {File} file - 要上传的文件
 * @returns {Promise<Object>} 上传后的文档信息
 */
export async function uploadDocument(knowledgeBaseUuid, file) {
  try {
    console.log('[uploadDocument] 原始文件名:', file.name)
    console.log('[uploadDocument] 文件类型:', file.type)
    
    // 获取 API 配置
    const apiConfig = await getApiConfig()
    const { baseUrl, headers } = apiConfig
    
    // 直接使用 fetch API 上传文件（不需要 Base64 编码）
    const formData = new FormData()
    formData.append('file', file)
    formData.append('knowledge_base_uuid', knowledgeBaseUuid)
    
    const apiUrl = `${baseUrl}/knowledge/document/upload`
    console.log('[uploadDocument] 直接上传到:', apiUrl)
    
    // 构建请求头（不要设置 Content-Type，浏览器会自动为 FormData 设置）
    const requestHeaders = new Headers()
    if (headers.Authorization) {
      requestHeaders.append('Authorization', headers.Authorization)
    }
    if (headers['X-Username']) {
      requestHeaders.append('X-Username', headers['X-Username'])
    }
    
    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: requestHeaders, // 只包含认证头，不包含 Content-Type
      body: formData // FormData 会自动设置 Content-Type: multipart/form-data
    })
    
    const responseData = await response.json()
    
    if (response.ok && responseData.success) {
      return responseData.data
    }
    
    throw new Error(responseData.message || responseData.detail || '上传文档失败')
  } catch (error) {
    console.error('上传文档失败:', error)
    throw error
  }
}

/**
 * 处理文档
 * @param {string} storageDocUuid - storage 文档 UUID
 * @returns {Promise<Object>} 处理后的文档信息
 */
export async function processDocument(storageDocUuid) {
  try {
    return await apiRequest('POST', '/knowledge/document/process', {
      storage_doc_uuid: storageDocUuid
    })
  } catch (error) {
    console.error('处理文档失败:', error)
    throw error
  }
}

/** 对话场景 default_upload；上传超时 300s */
const CHAT_DOC_TIMEOUT_MS = 300000

/**
 * 对话上传：上传到用户默认对话知识库（字面量 default_upload）
 * @param {File} file
 * @returns {Promise<Object>} data（含 storage_doc_uuid）
 */
export async function uploadDocumentForChat(file) {
  const { headers, baseUrl } = await buildHeaders(false)
  const formData = new FormData()
  formData.append('file', file)
  formData.append('knowledge_base_uuid', 'default_upload')

  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), CHAT_DOC_TIMEOUT_MS)

  try {
    const requestHeaders = new Headers()
    if (headers.get('Authorization')) requestHeaders.append('Authorization', headers.get('Authorization'))
    if (headers.get('X-Username')) requestHeaders.append('X-Username', headers.get('X-Username'))

    const response = await fetch(`${baseUrl}/knowledge/document/upload`, {
      method: 'POST',
      headers: requestHeaders,
      body: formData,
      signal: controller.signal
    })
    const responseData = await response.json()
    if (response.ok && responseData.success) {
      return responseData.data
    }
    throw new Error(responseData.message || responseData.detail || '上传失败')
  } finally {
    clearTimeout(timer)
  }
}

/**
 * 对话场景文档处理（长超时）
 */
export async function processDocumentForChat(storageDocUuid) {
  const { headers, baseUrl } = await buildHeaders(true)
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), CHAT_DOC_TIMEOUT_MS)
  try {
    const response = await fetch(`${baseUrl}/knowledge/document/process`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ storage_doc_uuid: storageDocUuid }),
      signal: controller.signal
    })
    const responseData = await response.json()
    if (response.ok && responseData.success) {
      return responseData.data
    }
    throw new Error(responseData.message || responseData.detail || '处理文档失败')
  } finally {
    clearTimeout(timer)
  }
}

/**
 * 批量删除文档
 * @param {string[]} storageDocUuids - 要删除的 storage 文档 UUID 列表
 * @returns {Promise<Object>} 删除结果
 */
export async function deleteDocuments(storageDocUuids) {
  try {
    if (!storageDocUuids || storageDocUuids.length === 0) {
      throw new Error('至少需要提供一个 storage 文档 UUID')
    }
    
    return await apiRequest('POST', '/knowledge/document/delete', {
      storage_doc_uuids: storageDocUuids
    })
  } catch (error) {
    console.error('删除文档失败:', error)
    throw error
  }
}

/**
 * 设置文档启用状态
 * @param {string} storageDocUuid - storage 文档 UUID
 * @param {boolean} isEnabled - 是否启用（true-启用，false-禁用）
 * @returns {Promise<Object>} 更新后的文档信息
 */
export async function setDocumentEnableStatus(storageDocUuid, isEnabled) {
  try {
    if (!storageDocUuid) {
      throw new Error('storage 文档 UUID 不能为空')
    }
    
    return await apiRequest('PUT', `/knowledge/document/${storageDocUuid}/enable`, {
      is_enable: isEnabled ? 1 : 0
    })
  } catch (error) {
    console.error('设置文档启用状态失败:', error)
    throw error
  }
}

/**
 * 获取文档chunk列表
 * @param {string} storageDocUuid - storage 文档 UUID
 * @param {number} startIndex - 起始序号（从0开始）
 * @param {number} endIndex - 结束序号（不包含）
 * @returns {Promise<Object>} chunk列表数据
 */
export async function getDocumentChunks(storageDocUuid, startIndex = 0, endIndex = null) {
  try {
    if (!storageDocUuid) {
      throw new Error('storage 文档 UUID 不能为空')
    }
    
    const params = new URLSearchParams()
    params.append('start_index', startIndex.toString())
    if (endIndex !== null) {
      params.append('end_index', endIndex.toString())
    }
    
    const queryString = params.toString()
    const endpoint = `/knowledge/document/${storageDocUuid}/chunks${queryString ? `?${queryString}` : ''}`
    
    return await apiRequest('GET', endpoint)
  } catch (error) {
    console.error('获取文档chunk列表失败:', error)
    throw error
  }
}

/**
 * 当前用户知识库能力（分块调试等）
 * @returns {Promise<{ can_view_document_chunks: boolean }>}
 */
export async function getMyKbCapabilities() {
  return await apiRequest('GET', '/knowledge/me/capabilities')
}

/**
 * 获取文档预览 Blob 与类型
 * @param {string} storageDocUuid
 * @returns {Promise<{ blob: Blob, previewKind: string }>}
 */
export async function getDocumentPreview(storageDocUuid) {
  if (!storageDocUuid) {
    throw new Error('storage 文档 UUID 不能为空')
  }
  const { headers, baseUrl } = await buildHeaders(false)
  const url = `${baseUrl}/knowledge/document/${storageDocUuid}/preview`
  const response = await fetch(url, { method: 'GET', headers })
  if (!response.ok) {
    let detail = `${response.status} ${response.statusText}`
    try {
      const body = await response.json()
      detail = body.detail || body.message || detail
    } catch (_) {
      /* ignore */
    }
    throw new Error(detail)
  }
  const previewKindHeader = (response.headers.get('X-Preview-Kind') || '').toLowerCase()
  const contentType = (response.headers.get('Content-Type') || '').toLowerCase()
  let previewKind = previewKindHeader
  if (!previewKind) {
    if (contentType.includes('pdf')) previewKind = 'pdf'
    else if (contentType.includes('wordprocessingml') || contentType.includes('officedocument.word')) previewKind = 'docx'
    else if (contentType.includes('sheet') || contentType.includes('excel')) previewKind = 'xlsx'
    else if (contentType.startsWith('text/') || contentType.includes('xml')) previewKind = 'text'
    else previewKind = 'pdf'
  }
  const blob = await response.blob()
  return { blob, previewKind }
}

/**
 * 获取图片的 Blob URL（带认证头）
 * @param {string} imagePath - 图片相对路径（如：96a5d6cf9e/image_11.jpeg）
 * @param {string} knowledgeBaseUuid - 知识库UUID
 * @returns {Promise<string>} 图片的 Blob URL
 */
export async function getImageUrl(imagePath, knowledgeBaseUuid) {
  try {
    if (!imagePath) {
      throw new Error('图片路径不能为空')
    }
    if (!knowledgeBaseUuid) {
      throw new Error('知识库UUID不能为空')
    }
    
    // 构建完整的图片路径：/kb/{knowledge_base_uuid}/{image_path}
    const fullPath = `/kb/${knowledgeBaseUuid}/${imagePath}`
    
    // 获取 API 配置以构建完整的 URL 和请求头
    const apiConfig = await getApiConfig()
    const { baseUrl, headers } = apiConfig
    
    // 构建完整的图片 URL
    const imageUrl = `${baseUrl}/knowledge/image${fullPath}`
    
    // 构建请求头
    const requestHeaders = new Headers()
    if (headers.Authorization) {
      requestHeaders.append('Authorization', headers.Authorization)
    }
    if (headers['X-Username']) {
      requestHeaders.append('X-Username', headers['X-Username'])
    }
    
    // 使用 fetch 获取图片（带认证头）
    const response = await fetch(imageUrl, {
      method: 'GET',
      headers: requestHeaders
    })
    
    if (!response.ok) {
      throw new Error(`获取图片失败: ${response.status} ${response.statusText}`)
    }
    
    // 将响应转换为 Blob
    const blob = await response.blob()
    
    // 创建 Blob URL
    const blobUrl = URL.createObjectURL(blob)
    
    return blobUrl
  } catch (error) {
    console.error('获取图片失败:', error)
    throw error
  }
}

/**
 * 创建文档处理进度的 WebSocket 连接
 * @param {string} storageDocUuid - storage 文档 UUID
 * @param {Function} onProgress - 进度更新回调函数 (progress) => {}
 * @param {Function} onComplete - 完成回调函数 (data) => {}
 * @param {Function} onError - 错误回调函数 (error) => {}
 * @returns {Promise<WebSocket>} WebSocket 连接对象
 */
export async function createDocumentProcessWebSocket(storageDocUuid, onProgress, onComplete, onError) {
  try {
    const apiConfig = await getApiConfig()
    const { baseUrl, headers } = apiConfig
    
    let wsUrl = baseUrl.replace(/^http:/, 'ws:').replace(/^https:/, 'wss:')
    wsUrl = `${wsUrl}/knowledge/document/process/ws/${storageDocUuid}`
    
    // WebSocket 不支持自定义 HTTP 头，需要通过 URL 参数传递认证信息
    const params = new URLSearchParams()
    if (headers.Authorization) {
      // 提取 Bearer token
      const token = headers.Authorization.replace('Bearer ', '')
      params.append('token', token)
    }
    if (headers['X-Username']) {
      params.append('username', headers['X-Username'])
    }
    
    if (params.toString()) {
      wsUrl += `?${params.toString()}`
    }
    
    console.log('[createDocumentProcessWebSocket] 连接 WebSocket:', wsUrl)
    
    return new Promise((resolve, reject) => {
      const ws = new WebSocket(wsUrl)
      
      ws.onopen = () => {
        console.log('[WebSocket] 连接已建立:', storageDocUuid)
        resolve(ws)
        // 发送心跳（可选）
        // ws.send('ping')
      }
      
      ws.onerror = (error) => {
        console.error('[WebSocket] 连接错误:', error)
        reject(new Error('WebSocket 连接错误'))
        if (onError) {
          onError({ error: 'WebSocket 连接错误' })
        }
      }
      
      ws.onmessage = (event) => {
      try {
        // 处理心跳响应
        if (event.data === 'pong') {
          console.log('[WebSocket] 收到心跳响应')
          return
        }
        
        const message = JSON.parse(event.data)
        console.log('[WebSocket] 收到消息:', message)
        
        switch (message.type) {
          case 'progress':
            if (onProgress) {
              onProgress({
                storageDocUuid: message.storage_doc_uuid,
                progress: message.processing_progress,
                status: message.processing_status,
                message: message.message,
                timestamp: message.timestamp
              })
            }
            break
            
          case 'complete':
            if (onComplete) {
              onComplete({
                storageDocUuid: message.storage_doc_uuid,
                data: message.data,
                timestamp: message.timestamp
              })
            }
            // 处理完成后可以选择关闭连接
            // ws.close()
            break
            
          case 'error':
            if (onError) {
              onError({
                storageDocUuid: message.storage_doc_uuid,
                error: message.error,
                timestamp: message.timestamp
              })
            }
            break
            
          default:
            console.warn('[WebSocket] 未知消息类型:', message.type)
        }
      } catch (error) {
        console.error('[WebSocket] 解析消息失败:', error)
        if (onError) {
          onError({ error: '解析消息失败', raw: event.data })
        }
      }
    }
    
    ws.onclose = (event) => {
      console.log('[WebSocket] 连接已关闭:', event.code, event.reason)
    }
    })
  } catch (error) {
    console.error('创建 WebSocket 连接失败:', error)
    throw error
  }
}
