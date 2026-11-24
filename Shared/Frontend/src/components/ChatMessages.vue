<template>
  <div class="chat-messages" ref="messagesContainer">
    <div class="chat-main-column">
    <div
      v-for="turn in displayTurns"
      :key="turn.key"
      class="chat-turn"
      :class="{ 'chat-turn--latest': turn.pinStatus }"
    >
      <div v-if="turn.user" class="chat-turn-user">
        <UserMessageBubble :message="turn.user" />
      </div>
      <div v-if="turn.liveTimerText" class="live-turn-timer">{{ turn.liveTimerText }}</div>
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
          :duration-ms="item.durationMs"
          :pin-status="!!turn.pinStatus"
        />
      </div>
    </div>
    <div class="user-pin-spacer" :style="{ height: userPinSpacerHeight + 'px' }"></div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, watch, nextTick, onMounted, onUnmounted } from 'vue'
import UserMessageBubble from './UserMessageBubble.vue'
import SystemMessageBubble from './SystemMessageBubble.vue'
import AssistantTurnBlock from './AssistantTurnBlock.vue'
import { applyConversationSearch, setConversationSearchCurrent } from '../utils/conversationSearchHighlight.js'

const props = defineProps({
  messages: {
    type: Array,
    default: () => []
  },
  /** 当前看见的会话这一轮还在跑。别的会话在处理时不要当成这一轮 */
  running: {
    type: Boolean,
    default: false
  },
  /** 进行中的「已处理 …」。空则不画 */
  liveTimerText: {
    type: String,
    default: ''
  },
  /** 递增则强制钉到最新（切历史 / 切形态 / 发送新消息）；不沿用上翻取消跟随 */
  pinToLatestToken: {
    type: [Number, String],
    default: 0
  },
  /** 递增则把刚发出的用户气泡对齐到可见区顶部，并在下面留空 */
  pinUserToTopToken: {
    type: [Number, String],
    default: 0
  },
  /** 这次挂载是刚发送的用户消息，而不是打开历史 */
  pendingUserPin: {
    type: Boolean,
    default: false
  },
  searchQuery: { type: String, default: '' },
  searchIndex: { type: Number, default: 0 },
  searchScrollToken: { type: Number, default: 0 }
})

const emit = defineEmits(['search-count'])

/** 用户气泡 / 提示单独一项；连续助手消息合成一轮。
 * 归档：后面已有用户消息；或这一轮已经不在跑（含历史）。刚点停止的本轮先展开。 */
/** 只认后端 duration_ms（回推或历史回放）。没有或不足 500ms 不展示。 */
function computeTurnDurationMs (user) {
  const n = Number(user?.duration_ms ?? user?.durationMs)
  return Number.isFinite(n) && n >= 500 ? n : null
}

const displayItems = computed(() => {
  const items = []
  const list = props.messages || []
  let i = 0
  let lastUser = null
  while (i < list.length) {
    const m = list[i]
    if (m.role === 'user') {
      lastUser = m
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
    const aborted = run.some((m) => m.aborted)
    const streaming = run.some((m) => m.isStreaming)
    const stillRunning = !hasUserAfter && (props.running || streaming)
    let archived = true
    if (stillRunning) archived = false
    else if (aborted && !fromHistory && !hasUserAfter) archived = false
    items.push({
      kind: 'assistant',
      key: `asst-${run.map((x) => x.id).join('-')}`,
      messages: run,
      archived,
      durationMs: computeTurnDurationMs(lastUser)
    })
  }
  return items
})

/** 一轮 = 一条用户消息 + 其后的助手/提示；用户气泡在轮内 sticky，滚过本轮才离开顶部。
 * 最新一条已归档的系统回复，「已运行」或「奋力工作的记录」吸在该轮气泡下面。 */
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
  if (props.liveTimerText && turns.length) {
    turns[turns.length - 1].liveTimerText = props.liveTimerText
  }
  for (let i = turns.length - 1; i >= 0; i--) {
    if (turns[i].rest.some((item) => item.kind === 'assistant' && item.archived)) {
      turns[i].pinStatus = true
      break
    }
  }
  return turns
})

const messagesContainer = ref(null)
const shouldAutoScroll = ref(true)
const userPinArmed = ref(false)
const userPinLatched = ref(false)
const userPinSpacerHeight = ref(0)
const NEAR_BOTTOM_PX = 50
let suppressScrollUntil = 0
let resizeObserver = null
let statusStickyObserver = null
const pinTimers = []
let scrollLock = 0
let lastScrollTop = 0
let alignGen = 0
let pointerDown = false
let gestureActive = false
let gestureStartTop = 0
let gestureDelta = 0
let gestureTimer = null

function clearPinTimers () {
  while (pinTimers.length) {
    clearTimeout(pinTimers.pop())
  }
}

function pinNow () {
  const el = messagesContainer.value
  if (!el || gestureActive) return
  setScrollTop(el, el.scrollHeight)
}

function setScrollTop (el, value) {
  scrollLock += 1
  suppressScrollUntil = performance.now() + 300
  el.scrollTop = value
  lastScrollTop = el.scrollTop
  scrollLock -= 1
}

function disarmUserPin () {
  if (!userPinLatched.value || !userPinArmed.value) return
  userPinArmed.value = false
  shouldAutoScroll.value = false
  alignGen += 1
}

function rearmUserPin () {
  if (!userPinLatched.value || userPinArmed.value) return
  userPinArmed.value = true
  shouldAutoScroll.value = true
  alignUserBubbleToTop()
}

function finishUserGesture () {
  gestureTimer = null
  gestureActive = false
  const el = messagesContainer.value
  if (!el || !userPinLatched.value) return
  const top = el.scrollTop
  const movedUp = top < gestureStartTop - 4
  const movedDown = top > gestureStartTop + 4
  const nearBottom = el.scrollHeight - top - el.clientHeight < NEAR_BOTTOM_PX
  if (userPinArmed.value && movedUp) {
    disarmUserPin()
    return
  }
  if (!userPinArmed.value && movedDown && nearBottom) {
    rearmUserPin()
    return
  }
  if (userPinArmed.value) alignUserBubbleToTop()
}

function beginUserGesture (startTop) {
  if (!gestureActive) {
    gestureActive = true
    gestureStartTop = startTop
    gestureDelta = 0
    alignGen += 1
  }
  if (gestureTimer) clearTimeout(gestureTimer)
  gestureTimer = setTimeout(finishUserGesture, 180)
}

function lastUserTurn () {
  const root = messagesContainer.value
  if (!root) return null
  const bubbles = root.querySelectorAll('.chat-turn-user')
  const bubble = bubbles[bubbles.length - 1]
  if (!bubble) return null
  const turn = bubble.closest('.chat-turn')
  if (!turn) return null
  return { bubble, turn }
}

function alignUserBubbleToTop () {
  if (gestureActive || !userPinArmed.value) return
  const gen = alignGen
  nextTick(() => {
    requestAnimationFrame(() => {
      if (gen !== alignGen || gestureActive || !userPinArmed.value) return
      const el = messagesContainer.value
      const found = lastUserTurn()
      if (!el || !found || el.clientHeight <= 0) return
      const { turn } = found
      if (turn.offsetHeight >= el.clientHeight) {
        userPinSpacerHeight.value = 0
        nextTick(() => {
          const box = messagesContainer.value
          if (!box || gen !== alignGen || gestureActive || !userPinArmed.value) return
          setScrollTop(box, box.scrollHeight)
        })
        return
      }
      userPinSpacerHeight.value = Math.max(0, el.clientHeight - turn.offsetHeight)
      nextTick(() => {
        const box = messagesContainer.value
        const again = lastUserTurn()
        if (!box || !again || gen !== alignGen || gestureActive || !userPinArmed.value) return
        const boxRect = box.getBoundingClientRect()
        const bubbleRect = again.bubble.getBoundingClientRect()
        setScrollTop(box, box.scrollTop + (bubbleRect.top - boxRect.top))
      })
    })
  })
}

function armUserPin () {
  userPinLatched.value = true
  userPinArmed.value = true
  shouldAutoScroll.value = true
  alignUserBubbleToTop()
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

function onUserWheel (event) {
  if (!userPinLatched.value || !event.deltaY) return
  beginUserGesture(messagesContainer.value?.scrollTop ?? lastScrollTop)
  gestureDelta += event.deltaY
  if (gestureDelta < -8) disarmUserPin()
}

function onPointerDown () {
  pointerDown = true
}

function onPointerUp () {
  pointerDown = false
}

const handleScroll = () => {
  const container = messagesContainer.value
  if (!container) return
  const scrollTop = container.scrollTop
  const prev = lastScrollTop
  if (scrollLock > 0) {
    lastScrollTop = scrollTop
    return
  }
  lastScrollTop = scrollTop
  if (!userPinLatched.value) {
    shouldAutoScroll.value = container.scrollHeight - scrollTop - container.clientHeight < NEAR_BOTTOM_PX
    return
  }
  if (performance.now() < suppressScrollUntil && !pointerDown && !gestureActive) return
  const moved = Math.abs(scrollTop - prev) > 1
  if (pointerDown && moved) beginUserGesture(prev)
  else if (gestureActive) beginUserGesture(gestureStartTop)
  if (gestureActive && userPinArmed.value && scrollTop < gestureStartTop - 4) disarmUserPin()
}

const scrollToBottom = () => {
  if (gestureActive || !shouldAutoScroll.value) return
  nextTick(() => {
    const el = messagesContainer.value
    if (!el || gestureActive || !shouldAutoScroll.value) return
    setScrollTop(el, el.scrollHeight)
  })
}

function applySearchHighlight () {
  nextTick(() => {
    const count = applyConversationSearch(messagesContainer.value, props.searchQuery)
    emit('search-count', count)
    setConversationSearchCurrent(props.searchIndex, false)
  })
}

watch(() => props.searchQuery, applySearchHighlight)
watch(() => props.messages, applySearchHighlight, { deep: true })
watch(() => props.searchScrollToken, () => {
  nextTick(() => setConversationSearchCurrent(props.searchIndex, true))
})

function syncLatestStatusSticky () {
  const root = messagesContainer.value
  if (!root) return
  const turn = root.querySelector('.chat-turn--latest')
  if (!turn) return
  const user = turn.querySelector('.chat-turn-user')
  const h = user ? user.offsetHeight : 0
  turn.style.setProperty('--turn-user-sticky-h', `${h}px`)
}

function bindStatusStickyObserver () {
  syncLatestStatusSticky()
  if (typeof ResizeObserver === 'undefined') return
  if (!statusStickyObserver) {
    statusStickyObserver = new ResizeObserver(() => syncLatestStatusSticky())
  }
  const root = messagesContainer.value
  const user = root?.querySelector('.chat-turn--latest .chat-turn-user')
  if (user) statusStickyObserver.observe(user)
}

function scheduleStatusSticky () {
  nextTick(() => {
    requestAnimationFrame(bindStatusStickyObserver)
  })
}

onMounted(() => {
  nextTick(() => {
    if (messagesContainer.value) {
      messagesContainer.value.addEventListener('scroll', handleScroll, { passive: true })
      messagesContainer.value.addEventListener('wheel', onUserWheel, { passive: true })
      messagesContainer.value.addEventListener('pointerdown', onPointerDown)
      window.addEventListener('pointerup', onPointerUp)
      if (typeof ResizeObserver !== 'undefined') {
        resizeObserver = new ResizeObserver(() => {
          if (gestureActive) return
          if (userPinArmed.value) alignUserBubbleToTop()
          else if (shouldAutoScroll.value) pinNow()
        })
        resizeObserver.observe(messagesContainer.value)
      }
    }
    if (props.pendingUserPin) armUserPin()
    else forcePinToLatest()
    applySearchHighlight()
    scheduleStatusSticky()
  })
})

onUnmounted(() => {
  clearPinTimers()
  if (gestureTimer) clearTimeout(gestureTimer)
  window.removeEventListener('pointerup', onPointerUp)
  if (resizeObserver) {
    resizeObserver.disconnect()
    resizeObserver = null
  }
  if (statusStickyObserver) {
    statusStickyObserver.disconnect()
    statusStickyObserver = null
  }
  if (messagesContainer.value) {
    messagesContainer.value.removeEventListener('scroll', handleScroll)
    messagesContainer.value.removeEventListener('wheel', onUserWheel)
    messagesContainer.value.removeEventListener('pointerdown', onPointerDown)
  }
})

watch(() => props.pinToLatestToken, (next, prev) => {
  if (next === prev) return
  userPinArmed.value = false
  userPinLatched.value = false
  userPinSpacerHeight.value = 0
  forcePinToLatest()
  scheduleStatusSticky()
})

watch(() => props.pinUserToTopToken, (next, prev) => {
  if (next === prev) return
  armUserPin()
}, { flush: 'post' })

watch(() => props.messages.length, (newLength, oldLength) => {
  console.log('[ChatMessages] 消息数量变化:', oldLength, '->', newLength, '消息列表:', props.messages.map(m => ({ id: m.id, role: m.role, content: m.content?.substring(0, 20) })))
  const last = props.messages[props.messages.length - 1]
  if (newLength > (oldLength ?? 0) && last?.role === 'user') return
  if (userPinArmed.value || props.pendingUserPin) return
  scrollToBottom()
})

watch(() => props.messages, () => {
  scheduleStatusSticky()
  if (userPinArmed.value) {
    alignUserBubbleToTop()
    return
  }
  if (props.pendingUserPin) return
  scrollToBottom()
}, { deep: true })
</script>

<style>
.chat-messages mark.conv-search-hit {
  background: #d6e4ff;
  color: inherit;
  border-radius: 2px;
}

.chat-messages mark.conv-search-hit--current {
  background: #3b82f6;
  color: #fff;
}
</style>

<style scoped>
.chat-messages {
  flex: 1;
  min-height: 0;
  min-width: 0;
  overflow-x: hidden;
  overflow-y: auto;
  overflow-anchor: none;
  /* 顶 padding 会在吸顶条上方留出一条可透视缝，改由吸顶条自己垫 */
  padding: 0 16px 16px;
  display: block;
  background-color: #f7f7f5;
  scrollbar-width: none; /* Firefox */
  -ms-overflow-style: none; /* 旧 Edge */
}

.chat-main-column {
  width: 100%;
  max-width: 100%;
  min-width: 0;
  margin: 0 auto;
  box-sizing: border-box;
}

.user-pin-spacer {
  width: 100%;
  flex-shrink: 0;
  pointer-events: none;
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
  padding: 8px 16px 0;
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

/* 最新一轮的「已运行 / 奋力工作的记录」吸在该轮用户气泡下面，下滑时整行停住 */
.chat-turn--latest {
  --turn-user-sticky-h: 0px;
}

.chat-turn--latest :deep(.work-fold-header.is-pin),
.chat-turn--latest :deep(.ran-line.is-pin) {
  position: sticky;
  top: var(--turn-user-sticky-h);
  z-index: 4;
  background: #f7f7f5;
  box-shadow: 0 -6px 0 #f7f7f5;
}

.chat-turn--latest :deep(.ran-line.is-pin) {
  margin-bottom: 0;
  padding-bottom: 4px;
}

.chat-turn:first-child:not(:has(.chat-turn-user)) {
  padding-top: 16px;
}

.chat-turn-user + .message-wrapper {
  margin-top: 4px;
}

.live-turn-timer {
  margin-top: 4px;
  margin-bottom: 4px;
  padding: 0 14px;
  font-size: 13px;
  font-weight: 400;
  line-height: 1.4;
  color: #999;
}

/* 折叠标题自己有内边距，紧跟用户气泡时不再叠一层顶 padding */
.chat-turn-user + .message-wrapper :deep(.assistant-turn > .system-message > .message-content),
.chat-turn-user + .message-wrapper :deep(.assistant-reply-wrap > .system-message > .message-content) {
  padding-top: 0;
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
  min-width: 0;
  max-width: 100%;
}
</style>

