import { sendMessage } from '../composables/useWebViewBridge'

// 工具别名缓存
const aliasCache = new Map()
let allAliasesPromise = null

function applyAliasMap(aliases) {
  if (!aliases || typeof aliases !== 'object') {
    return
  }
  for (const [name, alias] of Object.entries(aliases)) {
    if (name) {
      aliasCache.set(name, alias || name)
    }
  }
}

async function ensureAllAliases() {
  if (allAliasesPromise) {
    return allAliasesPromise
  }

  allAliasesPromise = (async () => {
    const response = await sendMessage('getToolAlias', { all: true })
    if (response && response.success && response.aliases) {
      applyAliasMap(response.aliases)
    }
  })().catch((error) => {
    console.error('预加载工具别名失败:', error)
    allAliasesPromise = null
  })

  return allAliasesPromise
}

/**
 * 获取工具别名
 * @param {string} toolName - 工具名称
 * @returns {Promise<string>} - 工具别名（如果找不到则返回工具名称）
 */
export async function getToolAlias(toolName) {
  if (!toolName) {
    return toolName
  }

  // 检查缓存
  if (aliasCache.has(toolName)) {
    return aliasCache.get(toolName)
  }

  try {
    await ensureAllAliases()
    if (aliasCache.has(toolName)) {
      return aliasCache.get(toolName)
    }

    // 从后端查询别名
    const response = await sendMessage('getToolAlias', { toolName })
    
    if (response && response.success && response.aliases) {
      applyAliasMap(response.aliases)
      if (aliasCache.has(toolName)) {
        return aliasCache.get(toolName)
      }
    }

    if (response && response.success && response.alias) {
      const alias = response.alias
      // 缓存结果
      aliasCache.set(toolName, alias)
      return alias
    }
  } catch (error) {
    console.error('获取工具别名失败:', error)
  }

  // 如果查询失败，返回工具名称
  aliasCache.set(toolName, toolName)
  return toolName
}

/**
 * 批量获取工具别名（用于预加载）
 * @param {string[]} toolNames - 工具名称数组
 */
export async function preloadToolAliases(toolNames) {
  if (!Array.isArray(toolNames) || toolNames.length === 0) {
    return
  }

  // 过滤掉已缓存的工具
  const uncachedTools = toolNames.filter(name => !aliasCache.has(name))
  
  if (uncachedTools.length === 0) {
    return
  }

  // 批量查询（可以优化为一次请求多个）
  const promises = uncachedTools.map(name => getToolAlias(name))
  await Promise.all(promises)
}

