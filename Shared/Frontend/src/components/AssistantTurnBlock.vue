<template>
  <div class="assistant-turn">
    <template v-if="!archived">
      <div class="assistant-reply-wrap">
        <SystemMessageBubble
          v-for="m in messages"
          :key="`${m.role}-${m.id}`"
          :message="m"
        />
        <MessageCopyBar :text="copyText" align="left" />
      </div>
    </template>
    <div v-else class="system-message">
      <div class="message-content">
        <WorkFold v-if="workMessage" :duration-ms="durationMs">
          <SystemMessageBubble :message="workMessage" compact force-process />
        </WorkFold>
        <div v-if="finalMessage" class="assistant-reply-wrap">
          <SystemMessageBubble :message="finalMessage" compact />
          <MessageCopyBar :text="finalCopyText" align="left" />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { computed } from 'vue'
import SystemMessageBubble from './SystemMessageBubble.vue'
import WorkFold from './WorkFold.vue'
import MessageCopyBar from './MessageCopyBar.vue'

const props = defineProps({
  messages: {
    type: Array,
    required: true
  },
  archived: {
    type: Boolean,
    default: false
  },
  durationMs: {
    type: Number,
    default: null
  }
})

function isBlankText (content) {
  return !content || !String(content).replace(/\s/g, '')
}

function collectSegments (messages) {
  const segs = []
  for (const m of messages) {
    let list = []
    if (m.segments && m.segments.length > 0) {
      list = m.segments.filter((s) => s && s.type !== 'todoList')
    } else if (m.content) {
      list = [{ type: 'text', content: m.content }]
    }
    for (const s of list) {
      if (s.type === 'text' && isBlankText(s.content)) continue
      segs.push({ ...s })
    }
  }
  return segs
}

const split = computed(() => {
  const segs = collectSegments(props.messages)
  const aborted = props.messages.some((m) => m.aborted)
  if (aborted) {
    return { work: segs, final: null }
  }
  let lastTool = -1
  segs.forEach((s, i) => {
    if (s.type === 'toolCall') lastTool = i
  })
  let finalIdx = -1
  for (let i = segs.length - 1; i > lastTool; i--) {
    if (segs[i].type === 'text') {
      finalIdx = i
      break
    }
  }
  if (finalIdx < 0) {
    return { work: segs, final: null }
  }
  return {
    work: segs.filter((_, i) => i !== finalIdx),
    final: segs[finalIdx]
  }
})

const workMessage = computed(() => {
  if (!split.value.work.length) return null
  return {
    id: `work-${props.messages[0]?.id}`,
    role: 'system',
    segments: split.value.work,
    isStreaming: false
  }
})

const finalMessage = computed(() => {
  if (!split.value.final) return null
  return {
    id: `final-${props.messages[0]?.id}`,
    role: 'system',
    segments: [split.value.final],
    isStreaming: false
  }
})

function collectText (messages) {
  const parts = []
  for (const m of messages || []) {
    if (m?.segments && m.segments.length > 0) {
      for (const s of m.segments) {
        if (s && s.type === 'text' && s.content && String(s.content).trim()) {
          parts.push(String(s.content).trim())
        }
      }
    } else if (m?.content && String(m.content).trim()) {
      parts.push(String(m.content).trim())
    }
  }
  return parts.join('\n\n')
}

const copyText = computed(() => collectText(props.messages))

const finalCopyText = computed(() => {
  const seg = split.value.final
  if (seg?.content && String(seg.content).trim()) return String(seg.content).trim()
  return copyText.value
})
</script>

<style scoped>
.assistant-turn > .assistant-reply-wrap :deep(.message-copy-bar) {
  padding: 0 14px;
}

.assistant-reply-wrap:hover :deep(.message-copy-bar),
.assistant-reply-wrap:focus-within :deep(.message-copy-bar) {
  opacity: 1;
  pointer-events: auto;
}

.assistant-turn :deep(.system-message + .system-message .message-content) {
  padding-top: 0;
}

.assistant-turn :deep(.system-message:has(+ .system-message) .message-content) {
  padding-bottom: 0;
}

.assistant-turn :deep(.system-message:has(+ .system-message)) {
  margin-bottom: 0;
}

.system-message {
  display: flex;
  justify-content: flex-start;
}

.message-content {
  max-width: 100%;
  padding: 10px 14px;
}
</style>
