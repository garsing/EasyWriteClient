<template>
  <div class="assistant-turn">
    <template v-if="!archived">
      <SystemMessageBubble
        v-if="liveWorkMessage"
        :key="`live-work-${messages[0]?.id}`"
        :message="liveWorkMessage"
        group-tools
      />
      <div v-if="liveFinalMessage" class="assistant-reply-wrap">
        <SystemMessageBubble
          :key="`live-final-${messages[0]?.id}`"
          :message="liveFinalMessage"
        />
        <MessageCopyBar :text="finalCopyText" align="left" />
      </div>
    </template>
    <div v-else class="system-message">
      <div class="message-content">
        <WorkFold v-if="workMessage" :duration-ms="durationMs">
          <SystemMessageBubble :message="workMessage" compact force-process />
        </WorkFold>
        <div v-else-if="ranText" class="ran-line">{{ ranText }}</div>
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

function markClamp (message) {
  let list = []
  if (message.segments && message.segments.length > 0) {
    list = message.segments.filter((s) => s && s.type !== 'todoList')
  } else if (message.content) {
    list = [{ type: 'text', content: message.content }]
  }
  const kept = []
  for (const s of list) {
    if (s.type === 'text' && isBlankText(s.content)) continue
    kept.push({ ...s })
  }
  let lastTool = -1
  kept.forEach((s, i) => {
    if (s.type === 'toolCall') lastTool = i
  })
  // 无工具信号：尾段正文按终稿字体流。出现 toolCall 后，它前面的正文才改成过程灰字。
  kept.forEach((s, i) => {
    if (s.type === 'text') s.clamped = lastTool > i
  })
  return kept
}

const streaming = computed(() => {
  const last = props.messages[props.messages.length - 1]
  return !!last?.isStreaming
})

const liveWorkMessage = computed(() => {
  if (!split.value.work.length) return null
  return {
    id: `live-work-${props.messages[0]?.id}`,
    role: 'system',
    segments: markClamp({
      segments: split.value.work,
      isStreaming: streaming.value && !split.value.final
    }),
    isStreaming: streaming.value && !split.value.final
  }
})

const liveFinalMessage = computed(() => {
  if (!split.value.final) return null
  return {
    id: `live-final-${props.messages[0]?.id}`,
    role: 'system',
    segments: [{ ...split.value.final, clamped: false }],
    isStreaming: streaming.value
  }
})

const copyText = computed(() => collectText(props.messages))

const finalCopyText = computed(() => {
  const seg = split.value.final
  if (seg?.content && String(seg.content).trim()) return String(seg.content).trim()
  return copyText.value
})

/** 没有可收进「奋力工作的记录」的过程时，回复上面写后端耗时。 */
const ranText = computed(() => {
  const text = formatWorkDuration(props.durationMs)
  return text ? `已运行 ${text}` : ''
})

function formatWorkDuration (ms) {
  const n = Number(ms)
  if (!Number.isFinite(n) || n < 500) return ''
  const totalSec = Math.round(n / 1000)
  if (totalSec < 1) return ''
  const hours = Math.floor(totalSec / 3600)
  const minutes = Math.floor((totalSec % 3600) / 60)
  const seconds = totalSec % 60
  if (hours > 0) {
    return minutes > 0 ? `${hours}小时${minutes}分` : `${hours}小时`
  }
  if (minutes > 0) {
    return seconds > 0 ? `${minutes}分${seconds}秒` : `${minutes}分`
  }
  return `${seconds}秒`
}
</script>

<style scoped>
.assistant-turn > .assistant-reply-wrap :deep(.message-copy-bar) {
  padding: 0 14px;
}

.ran-line {
  margin-bottom: 4px;
  font-size: 13px;
  font-weight: 400;
  line-height: 1.4;
  color: #999;
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

.assistant-turn,
.assistant-reply-wrap {
  min-width: 0;
  max-width: 100%;
}

.system-message {
  display: flex;
  justify-content: flex-start;
  min-width: 0;
  max-width: 100%;
  width: 100%;
}

.message-content {
  min-width: 0;
  max-width: 100%;
  width: 100%;
  box-sizing: border-box;
  padding: 10px 14px;
}
</style>
