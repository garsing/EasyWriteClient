<template>
  <div class="assistant-turn">
    <template v-if="!archived">
      <SystemMessageBubble
        v-for="m in messages"
        :key="`${m.role}-${m.id}`"
        :message="m"
      />
    </template>
    <div v-else class="system-message">
      <div class="message-content">
        <WorkFold v-if="workMessage">
          <SystemMessageBubble :message="workMessage" compact force-process />
        </WorkFold>
        <SystemMessageBubble v-if="finalMessage" :message="finalMessage" compact />
      </div>
    </div>
  </div>
</template>

<script setup>
import { computed } from 'vue'
import SystemMessageBubble from './SystemMessageBubble.vue'
import WorkFold from './WorkFold.vue'

const props = defineProps({
  messages: {
    type: Array,
    required: true
  },
  archived: {
    type: Boolean,
    default: false
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
</script>

<style scoped>
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
