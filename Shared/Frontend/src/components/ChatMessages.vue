<template>
  <div class="chat-messages" ref="messagesContainer">
    <div v-for="message in messages" :key="`${message.role}-${message.id}`" class="message-wrapper">
      <UserMessageBubble v-if="message.role === 'user'" :message="message" />
      <SystemMessageBubble v-else :message="message" />
    </div>
  </div>
</template>

<script setup>
import { ref, watch, nextTick, onMounted, onUnmounted } from 'vue'
import UserMessageBubble from './UserMessageBubble.vue'
import SystemMessageBubble from './SystemMessageBubble.vue'

const props = defineProps({
  messages: {
    type: Array,
    default: () => []
  },
  /** 递增则强制钉到最新（切历史 / 切形态）；不沿用上一会话的上翻态 */
  pinToLatestToken: {
    type: [Number, String],
    default: 0
  }
})

const messagesContainer = ref(null)
const shouldAutoScroll = ref(true)
let suppressScrollUntil = 0
let resizeObserver = null
const pinTimers = []

function clearPinTimers () {
  while (pinTimers.length) {
    clearTimeout(pinTimers.pop())
  }
}

function pinNow () {
  const el = messagesContainer.value
  if (!el) return
  suppressScrollUntil = performance.now() + 300
  el.scrollTop = el.scrollHeight
}

/** 强制跟到底：恢复跟随，并多次钉底以等 WebView2 / 原生改尺寸落稳 */
function forcePinToLatest () {
  shouldAutoScroll.value = true
  clearPinTimers()
  pinNow()
  nextTick(pinNow)
  requestAnimationFrame(() => {
    requestAnimationFrame(pinNow)
  })
  pinTimers.push(setTimeout(pinNow, 80))
  pinTimers.push(setTimeout(pinNow, 200))
}

const handleScroll = () => {
  if (!messagesContainer.value) return
  if (performance.now() < suppressScrollUntil) return

  const container = messagesContainer.value
  const scrollTop = container.scrollTop
  const scrollHeight = container.scrollHeight
  const clientHeight = container.clientHeight

  const isNearBottom = scrollHeight - scrollTop - clientHeight < 50
  shouldAutoScroll.value = isNearBottom
}

const scrollToBottom = () => {
  if (shouldAutoScroll.value) {
    nextTick(() => {
      if (messagesContainer.value && shouldAutoScroll.value) {
        messagesContainer.value.scrollTop = messagesContainer.value.scrollHeight
      }
    })
  }
}

onMounted(() => {
  nextTick(() => {
    if (messagesContainer.value) {
      messagesContainer.value.addEventListener('scroll', handleScroll, { passive: true })
      if (typeof ResizeObserver !== 'undefined') {
        resizeObserver = new ResizeObserver(() => {
          if (shouldAutoScroll.value) pinNow()
        })
        resizeObserver.observe(messagesContainer.value)
      }
    }
    forcePinToLatest()
  })
})

onUnmounted(() => {
  clearPinTimers()
  if (resizeObserver) {
    resizeObserver.disconnect()
    resizeObserver = null
  }
  if (messagesContainer.value) {
    messagesContainer.value.removeEventListener('scroll', handleScroll)
  }
})

watch(() => props.pinToLatestToken, (next, prev) => {
  if (next === prev) return
  forcePinToLatest()
})

watch(() => props.messages.length, (newLength, oldLength) => {
  console.log('[ChatMessages] 消息数量变化:', oldLength, '->', newLength, '消息列表:', props.messages.map(m => ({ id: m.id, role: m.role, content: m.content?.substring(0, 20) })))
  scrollToBottom()
})

watch(() => props.messages, (newMessages, oldMessages) => {
  console.log('[ChatMessages] 消息列表变化，当前消息:', newMessages.map(m => ({ id: m.id, role: m.role, contentLength: m.content?.length || 0 })))
  scrollToBottom()
}, { deep: true })
</script>

<style scoped>
.chat-messages {
  flex: 1;
  min-height: 0;
  overflow-y: auto;
  padding: 16px;
  display: flex;
  flex-direction: column;
  gap: 0;
  background-color: #f7f7f5;
  scrollbar-width: none; /* Firefox */
  -ms-overflow-style: none; /* 旧 Edge */
}

/* 用户气泡 / 提示与助手之间留缝；连续助手气泡（多轮工具拆成多条消息）贴在一起 */
.message-wrapper + .message-wrapper {
  margin-top: 12px;
}

.message-wrapper:has(.system-message:not(.system-message--hint))
  + .message-wrapper:has(.system-message:not(.system-message--hint)) {
  margin-top: 0;
}

.message-wrapper:has(.system-message:not(.system-message--hint))
  + .message-wrapper:has(.system-message:not(.system-message--hint))
  :deep(.message-content) {
  padding-top: 0;
}

.message-wrapper:has(.system-message:not(.system-message--hint)):has(
    + .message-wrapper .system-message:not(.system-message--hint)
  )
  :deep(.message-content) {
  padding-bottom: 0;
}

.message-wrapper:has(.system-message:not(.system-message--hint)):has(
    + .message-wrapper .system-message:not(.system-message--hint)
  )
  :deep(.system-message) {
  margin-bottom: 0;
}

.chat-messages::-webkit-scrollbar {
  display: none; /* Chrome / WebView2 */
  width: 0;
  height: 0;
}

.message-wrapper {
  display: flex;
  flex-direction: column;
}
</style>

