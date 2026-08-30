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
    :output-text="outputText"
    :output-done="outputDone"
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
  },
  outputText: {
    type: String,
    default: ''
  },
  outputDone: {
    type: Boolean,
    default: false
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
  'F_write_file',
  'B_calculate',
  'B_run_python',
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
// F_write_file 不在此列：只有写 .py 才展开（对齐旧 B_write_python）
const toolsWithDetails = [
  'F_process_paragraph_actions',
  'F_read_file',
  'F_run_terminal',
  'F_close_terminal',
  'F_close_document',
  'B_run_python'
]

const shouldRenderBox = computed(() => {
  return !toolsHidden.includes(props.toolName)
})

// 判断是否应该显示详细内容（可展开）
const shouldShowDetails = computed(() => {
  if (!shouldRenderBox.value) return false
  if (props.toolName === 'F_write_file') {
    return isPythonWritePath(tryParsePath(props.content))
  }
  return toolsWithDetails.includes(props.toolName)
})

// 根据工具类型确定语言类型
function isPythonWritePath(path) {
  return String(path || '').toLowerCase().endsWith('.py')
}

function guessWriteLanguage(path) {
  const p = String(path || '').toLowerCase()
  if (isPythonWritePath(p)) return 'python'
  if (p.endsWith('.xml')) return 'xml'
  if (p.endsWith('.yaml') || p.endsWith('.yml')) return 'yaml'
  if (p.endsWith('.json')) return 'json'
  return 'text'
}

function tryParsePath(raw) {
  try {
    const parsed = JSON.parse(raw)
    return parsed.path || parsed.filename || ''
  } catch {
    const match = String(raw || '').match(/"path"\s*:\s*"([^"]*)"/)
    return match ? match[1] : ''
  }
}

const languageType = computed(() => {
  switch (props.toolName) {
    case 'F_write_file':
      return guessWriteLanguage(tryParsePath(props.content))
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
        if (parsed.path !== undefined && parsed.path !== null) {
          return String(parsed.path)
        }
        if (parsed.filename !== undefined && parsed.filename !== null) {
          return String(parsed.filename)
        }
        return JSON.stringify(parsed, null, 2)
        
      case 'F_write_file': {
        const writePath = parsed.path || parsed.filename || ''
        const body = parsed.content !== undefined && parsed.content !== null
          ? String(parsed.content)
          : JSON.stringify(parsed, null, 2)
        return writePath ? `写入 "${writePath}"\n\n${body}` : body
      }
        
      case 'F_run_terminal':
        if (parsed.command !== undefined && parsed.command !== null) {
          return String(parsed.command)
        }
        return JSON.stringify(parsed, null, 2)

      case 'B_run_python':
        if (parsed.file_path !== undefined && parsed.file_path !== null) {
          return String(parsed.file_path)
        }
        if (parsed.path !== undefined && parsed.path !== null) {
          return String(parsed.path)
        }
        return JSON.stringify(parsed, null, 2)

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
    
    if (props.toolName === 'F_write_file') {
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
        const pathMatch = props.content.match(/"path"\s*:\s*"([^"]*)"/)
        return pathMatch ? `写入 "${pathMatch[1]}"\n\n${decodedContent}` : decodedContent
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
            const pathMatch = props.content.match(/"path"\s*:\s*"([^"]*)"/)
            return pathMatch ? `写入 "${pathMatch[1]}"\n\n${extracted}` : extracted
          }
        }
      }
    }
    
    if (props.toolName === 'F_write_file') {
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

