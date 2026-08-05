/**
 * 解析文本中的工具调用
 * 格式: [TOOL_CALL:工具名] {JSON内容}
 */
export function parseToolCalls(text) {
  if (!text) return []

  const toolCalls = []
  const pattern = /\[TOOL_CALL:([^\]]+)\]/g
  let match

  while ((match = pattern.exec(text)) !== null) {
    const toolName = match[1].trim()
    const toolCallStartIndex = match.index
    const toolCallEndIndex = match.index + match[0].length

    // 查找工具调用标记后的JSON内容
    let jsonContent = ''
    let isComplete = false
    let jsonStartIndex = toolCallEndIndex
    let jsonEndIndex = -1

    // 跳过空白字符
    while (jsonStartIndex < text.length && /\s/.test(text[jsonStartIndex])) {
      jsonStartIndex++
    }

    // 如果找到 { 开始，尝试解析JSON
    if (jsonStartIndex < text.length && text[jsonStartIndex] === '{') {
      let braceCount = 0

      // 计算大括号匹配，找到完整的JSON
      for (let i = jsonStartIndex; i < text.length; i++) {
        if (text[i] === '{') {
          braceCount++
        } else if (text[i] === '}') {
          braceCount--
          if (braceCount === 0) {
            jsonEndIndex = i
            break
          }
        }
      }

      if (jsonEndIndex >= jsonStartIndex) {
        // JSON完整
        jsonContent = text.substring(jsonStartIndex, jsonEndIndex + 1)
        isComplete = true
      } else {
        // JSON不完整（还在流式输出中）
        jsonContent = text.substring(jsonStartIndex)
        isComplete = false
      }
    }

    toolCalls.push({
      toolName,
      content: jsonContent,
      startIndex: toolCallStartIndex,
      endIndex: isComplete && jsonEndIndex >= 0 ? jsonEndIndex + 1 : text.length,
      isComplete
    })
  }

  return toolCalls
}

/**
 * 将文本中的工具调用替换为占位符，并返回分段信息
 */
export function processTextWithToolCalls(text) {
  if (!text) return { processedText: '', segments: [] }

  const toolCalls = parseToolCalls(text)
  if (toolCalls.length === 0) {
    return { processedText: text, segments: [{ type: 'text', content: text }] }
  }

  const segments = []
  let lastIndex = 0

  // 按位置排序工具调用
  const sortedToolCalls = [...toolCalls].sort((a, b) => a.startIndex - b.startIndex)

  for (const toolCall of sortedToolCalls) {
    // 添加工具调用前的文本
    if (toolCall.startIndex > lastIndex) {
      const textSegment = text.substring(lastIndex, toolCall.startIndex)
      if (textSegment) {
        segments.push({ type: 'text', content: textSegment })
      }
    }

    // 添加工具调用
    segments.push({
      type: 'toolCall',
      toolName: toolCall.toolName,
      content: toolCall.content,
      isComplete: toolCall.isComplete,
      startIndex: toolCall.startIndex // 用于生成稳定的key
    })

    lastIndex = toolCall.endIndex
  }

  // 添加最后的文本
  if (lastIndex < text.length) {
    const textSegment = text.substring(lastIndex)
    if (textSegment) {
      segments.push({ type: 'text', content: textSegment })
    }
  }

  return { processedText: text, segments }
}

