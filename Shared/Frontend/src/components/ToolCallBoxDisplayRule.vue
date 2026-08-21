<template>
  <!-- toolsHidden：整行不渲染（无标题、无折叠框） -->
  <ToolCallBox
    v-if="shouldRenderBox"
    :tool-name="toolName"
    :content="formattedContent"
    :is-complete="isComplete"
    :allow-expand="shouldShowDetails"
    :language="languageType"
    :icon="toolIcon"
    :result="result"
  />
</template>

<script setup>
import { computed } from 'vue'
import ToolCallBox from './ToolCallBox.vue'
import manipulateToolIcon from '../assets/images/manipulate_tool.png'
import readToolIcon from '../assets/images/read_tool.png'

const props = defineProps({
  toolName: {
    type: String,
    required: true
  },
  content: {
    type: String,
    required: true
  },
  isComplete: {
    type: Boolean,
    default: false
  },
  result: {
    type: Object,
    default: null
  }
})

// 使用 manipulate_tool.png 图标的工具列表
const manipulateTools = [
  'F_create_chart_from_xml',
  'F_create_table_from_yaml',
  'F_csv_to_xml',
  'F_extract_table_format',
  'F_extract_table_data',
  'F_apply_table_data_from_xml',
  'F_greet',
  'F_modify_yaml_file',
  'F_process_paragraph_actions',
  'F_process_document_actions',
  'F_write_format_file',
  'B_calculate',
  'B_run_python',
  'B_write_python',
  'F_run_terminal',
  'F_close_terminal',
  'F_close_document'
]

// 使用 read_tool.png 图标的工具列表
const readTools = [
  'F_read_file',
  'F_split_document_by_paragraphs',
  'F_get_document_content'
]

// 根据工具名称返回对应的图标路径
const toolIcon = computed(() => {
  if (manipulateTools.includes(props.toolName)) {
    return manipulateToolIcon
  } else if (readTools.includes(props.toolName)) {
    return readToolIcon
  }
  return null // 没有图标的工具返回 null
})

// 完全不显示（连折叠框标题也不出）。优先于 toolsWithDetails。
const toolsHidden = [
  'B_todo'
]

// 需要显示详细内容、可展开的工具列表
const toolsWithDetails = [
  'F_process_paragraph_actions',
  'F_read_file',
  'B_write_python',
  'F_run_terminal',
  'F_close_terminal',
  'F_close_document'
]

const shouldRenderBox = computed(() => {
  return !toolsHidden.includes(props.toolName)
})

// 判断是否应该显示详细内容（可展开）
const shouldShowDetails = computed(() => {
  if (!shouldRenderBox.value) return false
  return toolsWithDetails.includes(props.toolName)
})

// 根据工具类型确定语言类型
const languageType = computed(() => {
  switch (props.toolName) {
    case 'B_write_python':
      return 'python'
    case 'F_write_format_file':
      return 'yaml'
    case 'F_process_paragraph_actions':
    case 'F_process_document_actions':
    case 'F_run_terminal':
    case 'F_close_terminal':
    case 'F_close_document':
      return 'json'
    default:
      return 'text'
  }
})

// 根据工具类型格式化内容
const formattedContent = computed(() => {
  // 如果内容为空或不完整，直接返回
  if (!props.content || props.content.trim() === '') {
    return props.content
  }

  try {
    // 尝试解析 JSON
    const parsed = JSON.parse(props.content)
    
    switch (props.toolName) {
      case 'F_read_file':
        // 只显示 filename
        if (parsed.filename !== undefined && parsed.filename !== null) {
          return String(parsed.filename)
        }
        // 如果没有 filename，显示完整 JSON
        return JSON.stringify(parsed, null, 2)
        
      case 'F_write_format_file':
        // 只显示 content
        if (parsed.content !== undefined && parsed.content !== null) {
          // 直接返回 content 的值，不管是否为空
          return String(parsed.content)
        }
        // 如果没有 content 字段，显示完整 JSON（用于调试）
        console.warn('[ToolCallBoxDisplayRule] F_write_format_file: content 字段不存在', parsed)
        return JSON.stringify(parsed, null, 2)
        
      case 'F_run_terminal':
        if (parsed.command !== undefined && parsed.command !== null) {
          return String(parsed.command)
        }
        return JSON.stringify(parsed, null, 2)

      case 'B_write_python':
        // 显示 "写入[filename]文件" + python_code 内容
        let result = ''
        if (parsed.filename !== undefined && parsed.filename !== null) {
          result = `写入 "${parsed.filename}" 文件\n\n`
        }
        // 尝试多种可能的字段名
        const pythonCode = parsed.python_code || parsed.pythonCode || parsed.code
        if (pythonCode !== undefined && pythonCode !== null) {
          result += String(pythonCode)
        } else {
          // 如果没有找到代码字段，显示完整 JSON
          result += JSON.stringify(parsed, null, 2)
        }
        return result
        
      case 'F_process_paragraph_actions':
      case 'F_process_document_actions':
        // 显示完整 JSON
        return JSON.stringify(parsed, null, 2)
        
      default:
        // 其他工具显示完整 JSON（虽然这些工具通常不会展开）
        return JSON.stringify(parsed, null, 2)
    }
  } catch (error) {
    // 如果不是有效的JSON，可能是还在流式输出中
    // 对于特定工具，尝试提取部分内容
    
    if (props.toolName === 'F_write_format_file') {
      // 尝试从原始字符串中提取 content 字段
      // content 可能是多行字符串，需要处理转义字符
      // 使用更强大的正则表达式来匹配可能包含换行符和转义字符的字符串
      const contentMatch = props.content.match(/"content"\s*:\s*"((?:[^"\\]|\\.)*)"/s)
      
      if (contentMatch && contentMatch[1]) {
        // 解码转义的字符串
        const decodedContent = contentMatch[1]
          .replace(/\\n/g, '\n')
          .replace(/\\t/g, '\t')
          .replace(/\\r/g, '\r')
          .replace(/\\"/g, '"')
          .replace(/\\\\/g, '\\')
        return decodedContent
      }
      
      // 如果正则匹配失败，尝试查找 content 字段的位置
      const contentIndex = props.content.indexOf('"content"')
      if (contentIndex !== -1) {
        // 找到 content 字段，尝试提取值
        const afterContent = props.content.substring(contentIndex + 9) // 跳过 "content"
        const colonIndex = afterContent.indexOf(':')
        if (colonIndex !== -1) {
          const afterColon = afterContent.substring(colonIndex + 1).trim()
          // 如果以引号开始，尝试提取字符串值
          if (afterColon.startsWith('"')) {
            let extracted = ''
            let i = 1 // 跳过第一个引号
            let escaped = false
            while (i < afterColon.length) {
              const char = afterColon[i]
              if (escaped) {
                if (char === 'n') extracted += '\n'
                else if (char === 't') extracted += '\t'
                else if (char === 'r') extracted += '\r'
                else if (char === '"') extracted += '"'
                else if (char === '\\') extracted += '\\'
                else extracted += char
                escaped = false
              } else if (char === '\\') {
                escaped = true
              } else if (char === '"') {
                // 找到结束引号
                return extracted
              } else {
                extracted += char
              }
              i++
            }
            return extracted
          }
        }
      }
    }
    
    if (props.toolName === 'B_write_python') {
      // 使用手动解析方法提取 python_code（适用于流式输出和不完整 JSON）
      let pythonCodeMatch = null
      let isIncompleteMatch = false
      
      const pythonCodeIndex = props.content.indexOf('"python_code"')
      if (pythonCodeIndex !== -1) {
        // 找到 python_code 字段，手动提取值
        const afterPythonCode = props.content.substring(pythonCodeIndex + 13) // 跳过 "python_code"
        const colonIndex = afterPythonCode.indexOf(':')
        if (colonIndex !== -1) {
          const afterColon = afterPythonCode.substring(colonIndex + 1).trim()
          // 如果以引号开始，尝试提取字符串值
          if (afterColon.startsWith('"')) {
            let extracted = ''
            let i = 1 // 跳过第一个引号
            let escaped = false
            while (i < afterColon.length) {
              const char = afterColon[i]
              if (escaped) {
                // 处理转义字符
                if (char === 'n') extracted += '\n'
                else if (char === 't') extracted += '\t'
                else if (char === 'r') extracted += '\r'
                else if (char === '"') extracted += '"'
                else if (char === '\\') extracted += '\\'
                else extracted += char
                escaped = false
              } else if (char === '\\') {
                escaped = true
              } else if (char === '"') {
                // 找到结束引号，字符串完整
                pythonCodeMatch = { 1: extracted }
                break
              } else {
                extracted += char
              }
              i++
            }
            // 如果没有找到结束引号，说明还在流式输出中（不完整匹配）
            if (i >= afterColon.length && !pythonCodeMatch) {
              pythonCodeMatch = { 1: extracted }
              isIncompleteMatch = true
            }
          }
        }
      }
      
      const filenameMatch = props.content.match(/"filename"\s*:\s*"([^"]*)"/)
      
      if (pythonCodeMatch || filenameMatch) {
        let result = ''
        if (filenameMatch) {
          result = `写入 "${filenameMatch[1]}" 文件\n\n`
        }
        if (pythonCodeMatch && pythonCodeMatch[1] !== undefined) {
          // 手动解析时转义字符已经处理过了，直接使用
          result += pythonCodeMatch[1]
          
          // 调试日志（仅在开发环境）
          if (process.env.NODE_ENV === 'development') {
            console.log('[ToolCallBoxDisplayRule] B_write_python: 提取 python_code', {
              contentLength: props.content.length,
              pythonCodeLength: pythonCodeMatch[1].length,
              isIncomplete: isIncompleteMatch
            })
          }
        }
        // 如果找到了匹配（pythonCodeMatch 或 filenameMatch），返回结果
        // 即使 pythonCodeMatch[1] 为空字符串，也返回（可能是空代码）
        return result
      } else {
        // 调试日志：无法识别 python_code
        if (process.env.NODE_ENV === 'development') {
          console.warn('[ToolCallBoxDisplayRule] B_write_python: 无法识别 python_code', {
            contentLength: props.content.length,
            contentPreview: props.content.substring(0, 200),
            hasPythonCodeField: props.content.includes('"python_code"')
          })
        }
      }
    }
    
    // 如果解析失败且无法提取，对于 B_write_python 工具，返回截断的预览以避免性能问题
    // 其他工具返回原内容
    if (props.toolName === 'B_write_python') {
      // 如果无法识别 python_code，返回一个简短的提示，避免对大量 JSON 进行语法高亮
      const preview = props.content.length > 500 
        ? props.content.substring(0, 500) + '...' 
        : props.content
      return preview
    }
    
    // 其他工具返回原内容
    return props.content
  }
})
</script>

