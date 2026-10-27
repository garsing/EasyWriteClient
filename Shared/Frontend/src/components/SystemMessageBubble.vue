<template>
  <div class="system-message" :class="{ 'system-message--hint': message.isHint, 'system-message--compact': compact }">
    <div v-if="message.isHint" class="hint-text">{{ message.content }}</div>
    <div v-else class="message-content">
      <div class="message-text">
        <template
          v-for="(segment, index) in displaySegments"
          :key="segmentKey(segment, index)"
        >
          <ThinkingBox
            v-if="segment.type === 'thinking'"
            :content="segment.content"
            :is-complete="segment.isComplete"
          />
          <ClampedTextSegment
            v-else-if="segment.type === 'text'"
            :content="segment.content"
            :clamped="isClampedText(segment, index)"
          />
          <ToolCallBoxDisplayRule
            v-else-if="segment.type === 'toolCall'"
            :tool-name="segment.toolName"
            :content="segment.content"
            :is-complete="segment.isComplete"
            :result="segment.result"
            :output-text="segment.outputText"
            :output-done="segment.outputDone"
          />
          <ToolRunFold
            v-else-if="segment.type === 'toolRun'"
            :segments="segment.segments"
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
import ToolCallBoxDisplayRule from './ToolCallBoxDisplayRule.vue'
import ToolRunFold from './ToolRunFold.vue'
import ThinkingBox from './ThinkingBox.vue'
import ClampedTextSegment from './ClampedTextSegment.vue'
import { groupToolRuns, toolRunKey } from '../utils/toolRunGroup'

const props = defineProps({
  message: {
    type: Object,
    required: true
  },
  /** 嵌在「奋力工作的记录」里：去掉外层左右 padding */
  compact: {
    type: Boolean,
    default: false
  },
  /** 归档过程正文：全部按未确认收尾（限高 + 灰字） */
  forceProcess: {
    type: Boolean,
    default: false
  },
  /** 当前轮：连续工具收成探索/操作/思考折叠。归档进奋力工作后不要开 */
  groupTools: {
    type: Boolean,
    default: false
  }
})

function isBlankText (content) {
  return !content || !String(content).replace(/\s/g, '')
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

const displaySegments = computed(() => {
  if (!props.groupTools) return messageSegments.value
  return groupToolRuns(messageSegments.value)
})

/** 最后一个工具卡下标；没有工具则为 -1。其前的正文都限高 */
const lastToolIndex = computed(() => {
  const segs = messageSegments.value
  for (let i = segs.length - 1; i >= 0; i--) {
    if (segs[i].type === 'toolCall') return i
  }
  return -1
})

function isClampedText (segment, index) {
  if (props.forceProcess) return true
  if (props.groupTools) return !!segment.clamped
  if (lastToolIndex.value > index) return true
  // 流式中还不能确认自然结束，尾段也先限高；整轮结束后再展开
  return !!props.message.isStreaming
}

function segmentKey (segment, index) {
  if (segment.type === 'toolRun') return toolRunKey(segment, index)
  if (segment.type === 'toolCall') return `tool-${segment.startIndex}`
  if (segment.type === 'thinking') return `thinking-${segment.startIndex ?? index}`
  return `text-${index}`
}

function endsWithCompletedTool (segs) {
  const last = segs.length ? segs[segs.length - 1] : null
  if (!last) return false
  if (last.type === 'toolCall' && last.isComplete) return true
  if (last.type === 'toolRun') {
    const inner = last.segments || []
    const tail = inner.length ? inner[inner.length - 1] : null
    return !!(tail && tail.type === 'toolCall' && tail.isComplete)
  }
  return false
}

// 工具卡已出齐后不再在卡后挂加载点（整轮仍 isStreaming 时的误导）
const showStreamingSpinner = computed(() => {
  if (!props.message.isStreaming) return false
  if (endsWithCompletedTool(displaySegments.value)) return false
  return true
})
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
  font-family: "Segoe UI", "Microsoft YaHei UI", "Microsoft YaHei", sans-serif;
  font-size: 14px;
  font-weight: 500;
  line-height: 1.8;
}

.message-text :deep(p) {
  margin: 0 0 8px 0;
}

.message-text :deep(.text-segment-inner p:last-child) {
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

/* 行内 `code`：字形与正文一致，只靠浅灰底提示 */
.message-text :deep(code) {
  font-family: inherit;
  font-size: inherit;
  font-weight: inherit;
  line-height: inherit;
  color: inherit;
  background-color: #ececec;
  padding: 0 4px;
  border-radius: 3px;
}

.message-text :deep(pre) {
  background-color: #f5f5f5;
  padding: 10px;
  border-radius: 4px;
  overflow-x: auto;
  margin: 8px 0;
  font-family: Consolas, "Courier New", monospace;
  font-size: 13px;
  font-weight: 400;
}

.message-text :deep(pre code) {
  background-color: transparent;
  padding: 0;
  font-family: inherit;
  font-size: inherit;
  font-weight: inherit;
}

.streaming-indicator {
  margin-top: 8px;
  display: flex;
  align-items: center;
}
</style>
