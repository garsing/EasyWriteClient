<template>
  <div class="chat-messages" ref="messagesContainer">
    <div v-for="turn in displayTurns" :key="turn.key" class="chat-turn">
      <div v-if="turn.user" class="chat-turn-user">
        <UserMessageBubble :message="turn.user" />
      </div>
      <div
        v-for="item in turn.rest"
        :key="item.key"
        class="message-wrapper"
      >
        <SystemMessageBubble v-if="item.kind === 'hint'" :message="item.message" />
        <AssistantTurnBlock
          v-else
          :messages="item.messages"
          :archived="item.archived"
        />
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, watch, nextTick, onMounted, onUnmounted } from 'vue'
import UserMessageBubble from './UserMessageBubble.vue'
import SystemMessageBubble from './SystemMessageBubble.vue'
import AssistantTurnBlock from './AssistantTurnBlock.vue'

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

/** 用户气泡 / 提示单独一项；连续助手消息合成一轮。
 * 归档：后面已有用户消息，或从历史加载（切对话 / 重开应用）——最后一轮也收进「奋力工作的记录」。
 * 只有本会话刚生成完 / 刚点停止的那一轮保持展开。 */
const displayItems = computed(() => {
  const items = []
  const list = props.messages || []
  let i = 0
  while (i < list.length) {
    const m = list[i]
    if (m.role === 'user') {
      items.push({ kind: 'user', key: `user-${m.id}`, message: m })
      i += 1
      continue
    }
    if (m.isHint) {
      items.push({ kind: 'hint', key: `hint-${m.id}`, message: m })
      i += 1
      continue
    }
    const run = []
    while (i < list.length && list[i].role !== 'user' && !list[i].isHint) {
      run.push(list[i])
      i += 1
    }
    const hasUserAfter = list.slice(i).some((x) => x.role === 'user')
    const fromHistory = run.every((x) => x.fromHistory)
    items.push({
      kind: 'assistant',
      key: `asst-${run.map((x) => x.id).join('-')}`,
      messages: run,
      archived: hasUserAfter || fromHistory
    })
  }
  return items
})

/** 一轮 = 一条用户消息 + 其后的助手/提示；用户气泡在轮内 sticky，滚过本轮才离开顶部 */
const displayTurns = computed(() => {
  const turns = []
  for (const item of displayItems.value) {
    if (item.kind === 'user') {
      turns.push({ key: item.key, user: item.message, rest: [] })
      continue
    }
    if (!turns.length) {
      turns.push({ key: `lead-${item.key}`, user: null, rest: [item] })
    } else {
      turns[turns.length - 1].rest.push(item)
    }
  }
  return turns
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
  /* 顶 padding 会在吸顶条上方留出一条可透视缝，改由吸顶条自己垫 */
  padding: 0 16px 16px;
  display: block;
  background-color: #f7f7f5;
  scrollbar-width: none; /* Firefox */
  -ms-overflow-style: none; /* 旧 Edge */
}

.chat-turn {
  display: flex;
  flex-direction: column;
  min-width: 0;
}

.chat-turn + .chat-turn {
  margin-top: 16px;
}

/* 用户气泡贴在滚动区顶部，只在本轮还在视口内时吸顶 */
.chat-turn-user {
  position: sticky;
  top: 0;
  z-index: 5;
  margin: 0 -16px;
  padding: 8px 16px 10px;
  background: #f7f7f5;
}

/* 盖住滚动容器边缘 / 圆角处可能露出的正文 */
.chat-turn-user::before {
  content: '';
  position: absolute;
  left: 0;
  right: 0;
  top: -20px;
  height: 20px;
  background: #f7f7f5;
}

.chat-turn-user :deep(.user-message) {
  margin-bottom: 0;
}

.chat-turn:first-child:not(:has(.chat-turn-user)) {
  padding-top: 16px;
}

.chat-turn-user + .message-wrapper {
  margin-top: 12px;
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

