<template>
  <div class="tool-call-box">
    <div class="tool-call-header" :class="{ 'no-expand': !allowExpand }" @click="allowExpand && toggleExpand()">
      <span
        class="chevron"
        :class="allowExpand && isExpanded ? 'chevron-down' : 'chevron-right'"
        aria-hidden="true"
      />
      <img v-if="icon" :src="icon" class="tool-icon" alt="" />
      <span class="tool-name">{{ displayName }}</span>
    </div>
    <div v-if="isExpanded && allowExpand" class="tool-call-content">
      <pre :class="['code-content', `language-${language}`]"><code ref="codeElement" v-html="highlightedContent"></code></pre>
      <div v-if="resultText && !hideNestedResult" class="tool-output">
        <div class="tool-output-label">输出</div>
        <pre class="tool-output-body">{{ resultText }}</pre>
      </div>
    </div>
    <div v-if="hasOutputText" class="tool-process-output">
      <div class="tool-process-output-header" @click="toggleOutput">
        <span
          class="chevron"
          :class="outputExpanded ? 'chevron-down' : 'chevron-right'"
          aria-hidden="true"
        />
        <span class="tool-process-output-label">输出</span>
      </div>
      <div
        v-show="outputExpanded"
        ref="outputBody"
        class="tool-process-output-body"
        @scroll="handleOutputScroll"
      >{{ outputText }}</div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, watch, nextTick, onMounted, onUnmounted } from 'vue'
import { getToolAlias } from '../utils/toolAlias'
import {
  escapeHtml,
  highlightByLanguage,
  highlightPythonIncremental,
  resetPythonLineCache
} from '../utils/codeSyntaxHighlight'

const HIGHLIGHT_DEBOUNCE_MS = 150
const HIGHLIGHT_MIN_CHARS_DELTA = 200

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
  allowExpand: {
    type: Boolean,
    default: true // 默认允许展开
  },
  language: {
    type: String,
    default: 'text' // 语言类型：python, yaml, json, text
  },
  icon: {
    type: String,
    default: null // 图标路径，可选
  },
  result: {
    type: Object,
    default: null
  },
  outputText: {
    type: String,
    default: ''
  }
})

const hideNestedResult = computed(() => {
  return props.toolName === 'B_run_python' || props.toolName === 'F_run_terminal'
})

const outputText = computed(() => String(props.outputText || ''))
const hasOutputText = computed(() => outputText.value.length > 0)
const outputExpanded = ref(false)
const outputBody = ref(null)
const shouldAutoScroll = ref(true)
let suppressScrollUntil = 0

function toggleOutput () {
  outputExpanded.value = !outputExpanded.value
  if (outputExpanded.value && shouldAutoScroll.value) {
    nextTick(pinOutput)
  }
}

function pinOutput () {
  const el = outputBody.value
  if (!el) return
  suppressScrollUntil = performance.now() + 300
  el.scrollTop = el.scrollHeight
}

function handleOutputScroll () {
  const el = outputBody.value
  if (!el || performance.now() < suppressScrollUntil) return
  shouldAutoScroll.value = el.scrollHeight - el.scrollTop - el.clientHeight < 50
}

watch(outputText, (next, prev) => {
  const had = !!(prev && String(prev).length)
  const has = !!(next && String(next).length)
  if (!had && has && !props.isComplete) {
    outputExpanded.value = true
  }
  if (has && outputExpanded.value && shouldAutoScroll.value) {
    nextTick(pinOutput)
  }
})

watch(() => props.isComplete, (next, prev) => {
  if (next && !prev && hasOutputText.value) {
    outputExpanded.value = false
  }
})

const resultText = computed(() => {
  const r = props.result
  if (!r) return ''
  const data = r.data || {}
  const parts = []
  if (r.success === false && r.error) {
    parts.push(String(r.error))
  }
  if (data.exit_code !== undefined && data.exit_code !== null) {
    parts.push('exit_code=' + data.exit_code)
  }
  if (data.stdout) parts.push(String(data.stdout))
  if (data.stderr) parts.push(String(data.stderr))
  if (data.path) parts.push(String(data.path))
  if (data.closed && !data.stdout) {
    parts.push(data.already ? '终端已关闭（原本没有会话）' : '已关闭')
  }
  return parts.join('\n').trim()
})

const toolAlias = ref('')

// 显示名称：优先显示别名，如果没有别名则显示工具名称
const displayName = computed(() => {
  return toolAlias.value || props.toolName
})

// 加载工具别名
const loadToolAlias = async (name) => {
  try {
    toolAlias.value = await getToolAlias(name)
  } catch (error) {
    console.error('加载工具别名失败:', error)
    toolAlias.value = name
  }
}

// 组件挂载时加载别名
onMounted(() => {
  loadToolAlias(props.toolName)
})

// 监听 toolName 变化，重新加载别名
watch(() => props.toolName, (newName) => {
  if (newName) {
    loadToolAlias(newName)
  }
})

// 只有允许展开的工具才初始化展开状态
const isExpanded = ref(props.allowExpand && !props.isComplete) // 完整的默认折叠，不完整的默认展开

// 进行中展开（让用户看见写入过程）；完成后折叠。
// F_write_file 常先吐出合法 {"path":"x.py"}，随后才流 content：complete→incomplete 也要重新展开。
watch(() => props.isComplete, (newComplete, oldComplete) => {
  if (!props.allowExpand) return
  if (newComplete && !oldComplete) {
    isExpanded.value = false
  } else if (!newComplete && oldComplete) {
    isExpanded.value = true
  }
})

// 中途才识别出可展开（例如 path 刚落到 .py）且仍在流式：立刻展开
watch(() => props.allowExpand, (newAllowExpand) => {
  if (!newAllowExpand) {
    isExpanded.value = false
  } else if (!props.isComplete) {
    isExpanded.value = true
  }
})

const formattedContent = computed(() => {
  // 直接返回原始内容，不做 JSON 格式化
  // 因为格式化已经在 ToolCallBoxDisplayRule 中完成
  return props.content
})

const codeElement = ref(null)
const highlightedContent = ref('')

const pythonLineCache = {
  stableLineCount: 0,
  prefixHtml: ''
}

let highlightDebounceTimer = null
let lastHighlightedLength = 0
let pendingHighlightText = ''

function clearHighlightDebounce() {
  if (highlightDebounceTimer !== null) {
    clearTimeout(highlightDebounceTimer)
    highlightDebounceTimer = null
  }
}

function applyFullHighlight(text) {
  resetPythonLineCache(pythonLineCache)
  highlightedContent.value = highlightByLanguage(text, props.language)
  lastHighlightedLength = text.length
}

function applyIncrementalPythonHighlight(text) {
  highlightedContent.value = highlightPythonIncremental(text, pythonLineCache)
  lastHighlightedLength = text.length
}

function runScheduledHighlight(text) {
  if (props.isComplete || props.language !== 'python') {
    applyFullHighlight(text)
    return
  }
  applyIncrementalPythonHighlight(text)
}

function scheduleHighlightUpdate(text) {
  pendingHighlightText = text
  clearHighlightDebounce()

  if (props.isComplete) {
    applyFullHighlight(text)
    return
  }

  // 流式阶段先即时展示纯文本，避免每个 chunk 触发高亮
  highlightedContent.value = escapeHtml(text)

  const charDelta = text.length - lastHighlightedLength
  if (charDelta >= HIGHLIGHT_MIN_CHARS_DELTA) {
    runScheduledHighlight(text)
    return
  }

  highlightDebounceTimer = setTimeout(() => {
    highlightDebounceTimer = null
    runScheduledHighlight(pendingHighlightText)
  }, HIGHLIGHT_DEBOUNCE_MS)
}

watch(
  () => [formattedContent.value, props.language, props.isComplete],
  ([text]) => {
    scheduleHighlightUpdate(text || '')
  },
  { immediate: true }
)

onUnmounted(() => {
  clearHighlightDebounce()
})

const toggleExpand = () => {
  isExpanded.value = !isExpanded.value
}
</script>

<style scoped>
.tool-call-box {
  margin: 4px 0;
  border: none;
  border-radius: 0;
  background-color: #f7f7f5;
  overflow: hidden;
}

.tool-call-header {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 10px;
  cursor: pointer;
  user-select: none;
  background-color: transparent;
}

.tool-call-header:hover {
  background-color: rgba(0, 0, 0, 0.03);
}

.tool-call-header.no-expand {
  cursor: default;
}

.tool-call-header.no-expand:hover {
  background-color: transparent;
}

.tool-icon {
  width: 16px;
  height: 16px;
  object-fit: contain;
  flex-shrink: 0;
  opacity: 0.75;
}

/* 与 Todo 进度条一致：90° 直角折角 */
.chevron {
  flex: 0 0 auto;
  width: 6px;
  height: 6px;
  box-sizing: border-box;
  border-right: 1.5px solid #999;
  border-bottom: 1.5px solid #999;
  display: inline-block;
}

.chevron-right {
  transform: rotate(-45deg);
}

.chevron-down {
  transform: rotate(45deg);
}

.tool-name {
  flex: 1;
  font-size: 13px;
  font-weight: 500;
  color: #666;
}

.tool-call-content {
  padding: 0 10px 10px 24px;
  background-color: transparent;
  border-top: none;
}

.tool-output {
  margin-top: 8px;
}

.tool-output-label {
  font-size: 12px;
  color: #999;
  margin-bottom: 4px;
}

.tool-output-body {
  margin: 0;
  font-family: 'Courier New', 'Consolas', monospace;
  font-size: 12px;
  line-height: 1.5;
  color: #555;
  white-space: pre-wrap;
  word-wrap: break-word;
}

.tool-process-output {
  padding: 0 10px 10px 24px;
}

.tool-process-output-header {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 4px 0;
  cursor: pointer;
  user-select: none;
}

.tool-process-output-label {
  font-size: 12px;
  color: #999;
}

.tool-process-output-body {
  max-height: 100px;
  overflow: auto;
  margin: 0;
  font-family: 'Courier New', 'Consolas', monospace;
  font-size: 12px;
  line-height: 1.5;
  color: #555;
  white-space: pre-wrap;
  word-wrap: break-word;
}

.code-content {
  margin: 0;
  font-family: 'Courier New', 'Consolas', monospace;
  font-size: 12px;
  line-height: 1.5;
  color: #666;
  white-space: pre-wrap;
  word-wrap: break-word;
  overflow-x: auto;
  background-color: transparent;
  padding: 4px 0;
  border-radius: 0;
}

.code-content code {
  display: block;
  font-family: inherit;
  font-size: inherit;
  line-height: inherit;
  color: inherit;
  background: transparent;
  padding: 0;
  border: none;
}

/* 语法高亮样式 - 使用 :deep() 穿透 scoped 样式 */
/* 参照 CodeSyntaxHighlighter.cs 的颜色方案 */
.code-content :deep(.keyword) {
  color: #800080; /* Purple - 紫色，与 C# 实现一致 */
  font-weight: 500;
}

.code-content :deep(.string) {
  color: #a0522d; /* Brown - 棕色，与 C# 实现一致 */
}

.code-content :deep(.comment) {
  color: #008000; /* Green - 绿色，与 C# 实现一致 */
  font-style: normal;
}

.code-content :deep(.number) {
  color: #0000ff; /* Blue - 蓝色，与 C# 实现一致 */
}

.code-content :deep(.boolean) {
  color: #800080; /* Purple - 布尔值使用关键字颜色 */
}

.code-content :deep(.key) {
  color: #6f42c1;
  font-weight: 500;
}

.code-content :deep(.builtin) {
  color: #0000ff; /* Blue - 蓝色，与数字一致 */
  font-weight: normal;
}

.code-content :deep(.function) {
  color: #333333; /* Black - 黑色，标识符颜色 */
}

.code-content :deep(.class-name) {
  color: #800080; /* Purple - 紫色，与关键字一致 */
  font-weight: 500;
}

.code-content :deep(.decorator) {
  color: #008000; /* Green - 绿色，与注释一致 */
  font-style: normal;
}

.code-content :deep(.operator) {
  color: #ff4500; /* OrangeRed - 橙红色，与 C# 实现一致 */
}

.language-python,
.language-yaml,
.language-json,
.language-text {
  background-color: transparent;
}

.language-python code,
.language-yaml code,
.language-json code,
.language-text code {
  color: #666;
}
</style>

