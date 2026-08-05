<template>
  <div class="system-message" :class="{ 'system-message--hint': message.isHint }">
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
          <div v-else-if="segment.type === 'text'" v-html="renderMarkdown(segment.content)"></div>
          <ToolCallBoxDisplayRule
            v-else-if="segment.type === 'toolCall'"
            :tool-name="segment.toolName"
            :content="segment.content"
            :is-complete="segment.isComplete"
          />
        </template>
      </div>
      <div v-if="message.isStreaming" class="streaming-indicator">
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
  }
})

const messageSegments = computed(() => {
  if (props.message.segments && props.message.segments.length > 0) {
    // 忽略历史遗留的 todoList segment（改由底部进度条展示）
    return props.message.segments.filter((s) => s && s.type !== 'todoList')
  }
  if (!props.message.content) return []
  return [{ type: 'text', content: props.message.content }]
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

.message-text :deep(p:last-child) {
  margin-bottom: 0;
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
