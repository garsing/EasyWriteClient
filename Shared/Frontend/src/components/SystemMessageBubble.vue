<template>
  <div class="system-message" :class="{ 'system-message--hint': message.isHint, 'system-message--compact': compact }">
    <div v-if="message.isHint" class="hint-text">{{ message.content }}</div>
    <div v-else class="message-content">
      <div class="message-text">
        <template
          v-for="(segment, index) in messageSegments"
          :key="segment.type === 'toolCall'
            ? `tool-${segment.startIndex}`
            : segment.type === 'thinking'
              ? `thinking-${segment.startIndex ?? index}`
              : `text-${index}`"
        >
          <ThinkingBox
            v-if="segment.type === 'thinking'"
            :content="segment.content"
            :is-complete="segment.isComplete"
          />
          <div
            v-else-if="segment.type === 'text'"
            class="text-segment"
            :class="{ 'text-segment--clamped': isClampedText(index) }"
            v-html="renderMarkdown(normalizeTextContent(segment.content))"
          ></div>
          <ToolCallBoxDisplayRule
            v-else-if="segment.type === 'toolCall'"
            :tool-name="segment.toolName"
            :content="segment.content"
            :is-complete="segment.isComplete"
            :result="segment.result"
            :output-text="segment.outputText"
            :output-done="segment.outputDone"
          />
        </template>
      </div>
      <div v-if="showStreamingSpinner" class="streaming-indicator">
        <a-spin size="small" />
      </div>
    </div>
  </div>
</template>

<script setup>
import { computed } from 'vue'
import { marked } from 'marked'
import ToolCallBoxDisplayRule from './ToolCallBoxDisplayRule.vue'
import ThinkingBox from './ThinkingBox.vue'

const props = defineProps({
  message: {
    type: Object,
    required: true
  },
  /** 嵌在「奋力工作」里：去掉外层左右 padding */
  compact: {
    type: Boolean,
    default: false
  },
  /** 归档过程正文：全部按未确认收尾（限高 + 灰字） */
  forceProcess: {
    type: Boolean,
    default: false
  }
})

function isBlankText (content) {
  return !content || !String(content).replace(/\s/g, '')
}

function normalizeTextContent (text) {
  if (!text) return ''
  return String(text).replace(/^\s+|\s+$/g, '').replace(/\n{3,}/g, '\n\n')
}

const messageSegments = computed(() => {
  if (props.message.segments && props.message.segments.length > 0) {
    return props.message.segments.filter((s) => {
      if (!s || s.type === 'todoList') return false
      if (s.type === 'text' && isBlankText(s.content)) return false
      return true
    })
  }
  if (!props.message.content) return []
  return [{ type: 'text', content: props.message.content }]
})

/** 最后一个工具卡下标；没有工具则为 -1。其前的正文都限高 */
const lastToolIndex = computed(() => {
  const segs = messageSegments.value
  for (let i = segs.length - 1; i >= 0; i--) {
    if (segs[i].type === 'toolCall') return i
  }
  return -1
})

function isClampedText (index) {
  if (props.forceProcess) return true
  if (lastToolIndex.value > index) return true
  // 流式中还不能确认自然结束，尾段也先限高；整轮结束后再展开
  return !!props.message.isStreaming
}

// 工具卡已出齐后不再在卡后挂加载点（整轮仍 isStreaming 时的误导）
const showStreamingSpinner = computed(() => {
  if (!props.message.isStreaming) {
    return false
  }
  const segs = messageSegments.value
  const last = segs.length ? segs[segs.length - 1] : null
  if (last && last.type === 'toolCall' && last.isComplete) {
    return false
  }
  return true
})

const renderMarkdown = (text) => {
  if (!text) return ''
  try {
    return marked.parse(text)
  } catch (error) {
    return text
  }
}
</script>

<style scoped>
.system-message {
  display: flex;
  justify-content: flex-start;
  margin-bottom: 4px;
}

.system-message--hint {
  justify-content: center;
}

.system-message--compact .message-content {
  padding: 0;
}

.hint-text {
  padding: 4px 14px 8px;
  font-size: 14px;
  line-height: 1.5;
  color: #999;
  text-align: center;
}

.message-content {
  max-width: 100%;
  background-color: transparent;
  color: #333;
  padding: 10px 14px;
  word-wrap: break-word;
}

.message-text {
  font-size: 14px;
  line-height: 1.6;
}

.message-text :deep(p) {
  margin: 0 0 8px 0;
}

.text-segment :deep(p:last-child) {
  margin-bottom: 0;
}

.message-text :deep(p:empty),
.message-text :deep(p:has(> br:only-child)) {
  display: none;
}

/* 工具卡与限高盒紧贴，避免 4px margin 叠出缝 */
.message-text :deep(.tool-call-box) {
  margin-top: 0;
  margin-bottom: 0;
}

.text-segment--clamped {
  max-height: 80px;
  overflow-y: auto;
  overflow-x: hidden;
  color: #999;
  scrollbar-width: none;
  -ms-overflow-style: none;
}

.text-segment--clamped :deep(*) {
  color: inherit;
}

.text-segment--clamped::-webkit-scrollbar {
  display: none;
  width: 0;
  height: 0;
}

.message-text :deep(code) {
  background-color: #f5f5f5;
  padding: 2px 4px;
  border-radius: 3px;
  font-family: 'Courier New', monospace;
  font-size: 13px;
}

.message-text :deep(pre) {
  background-color: #f5f5f5;
  padding: 10px;
  border-radius: 4px;
  overflow-x: auto;
  margin: 8px 0;
}

.message-text :deep(pre code) {
  background-color: transparent;
  padding: 0;
}

.streaming-indicator {
  margin-top: 8px;
  display: flex;
  align-items: center;
}
</style>
