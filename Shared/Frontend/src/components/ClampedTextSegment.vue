<template>
  <div
    ref="viewportRef"
    class="text-segment"
    :class="{ 'text-segment--clamped': clamped }"
    @scroll.passive="handleScroll"
  >
    <div class="text-segment-inner" v-html="html"></div>
  </div>
</template>

<script setup>
import { computed, nextTick, onMounted, onUnmounted, ref, watch } from 'vue'
import { marked } from 'marked'

const props = defineProps({
  content: {
    type: String,
    default: ''
  },
  clamped: {
    type: Boolean,
    default: false
  }
})

function normalizeTextContent (text) {
  if (!text) return ''
  return String(text).replace(/^\s+|\s+$/g, '').replace(/\n{3,}/g, '\n\n')
}

// GFM 自动链接用 [^\s<]*，中文和全角标点不算空白，会把网址后面的正文一起吞进 <a>。
// 在网址和后续非网址字符之间插入发丝空格，让链接停在网址末尾。
const BARE_URL = /((?:https?:\/\/|ftp:\/\/|www\.)[A-Za-z0-9\-._~:/?#\[\]@!$&'()*+,;=%]+)/g

function splitUrlFromFollowingText (text) {
  return text.replace(BARE_URL, (url, _match, offset, source) => {
    const next = source[offset + url.length]
    if (!next || /[\sA-Za-z0-9\-._~:/?#\[\]@!$&'()*+,;=%]/.test(next)) return url
    return url + '\u200a'
  })
}

function wrapMarkdownTables (html) {
  return html.replace(
    /<table\b[^>]*>[\s\S]*?<\/table>/g,
    (table) => `<div class="md-table-wrap">${table}</div>`
  )
}

const html = computed(() => {
  const text = splitUrlFromFollowingText(normalizeTextContent(props.content))
  if (!text) return ''
  try {
    return wrapMarkdownTables(marked.parse(text))
  } catch {
    return text
  }
})

const viewportRef = ref(null)
const shouldAutoScroll = ref(true)
let suppressScrollUntil = 0
let resizeObserver = null

function pinNow () {
  const el = viewportRef.value
  if (!el || !props.clamped || !shouldAutoScroll.value) return
  suppressScrollUntil = performance.now() + 300
  el.scrollTop = el.scrollHeight
}

function pinSoon () {
  nextTick(pinNow)
  requestAnimationFrame(() => {
    requestAnimationFrame(pinNow)
  })
}

function handleScroll () {
  const el = viewportRef.value
  if (!el || !props.clamped) return
  if (performance.now() < suppressScrollUntil) return
  shouldAutoScroll.value = el.scrollHeight - el.scrollTop - el.clientHeight < 50
}

watch(
  () => [props.content, props.clamped],
  () => {
    if (!props.clamped) return
    pinSoon()
  },
  { flush: 'post' }
)

onMounted(() => {
  const inner = viewportRef.value?.firstElementChild
  if (inner && typeof ResizeObserver !== 'undefined') {
    resizeObserver = new ResizeObserver(() => {
      if (props.clamped && shouldAutoScroll.value) pinNow()
    })
    resizeObserver.observe(inner)
  }
  if (props.clamped) pinSoon()
})

onUnmounted(() => {
  if (resizeObserver) {
    resizeObserver.disconnect()
    resizeObserver = null
  }
})
</script>

<style scoped>
.text-segment--clamped {
  max-height: 80px;
  overflow-y: auto;
  overflow-x: hidden;
  overflow-anchor: none;
  color: #999;
  scrollbar-width: none;
  -ms-overflow-style: none;
}

.text-segment-inner :deep(.md-table-wrap) {
  overflow-x: auto;
  margin: 10px 0;
  border: 1px solid #e6e6e6;
  border-radius: 12px;
  background: #fff;
}

.text-segment-inner :deep(table) {
  width: max-content;
  min-width: 100%;
  border-collapse: collapse;
  margin: 0;
  font-weight: 400;
  line-height: 1.5;
}

.text-segment-inner :deep(th),
.text-segment-inner :deep(td) {
  padding: 12px 16px;
  border-bottom: 1px solid #ececec;
  text-align: left;
  vertical-align: middle;
  white-space: nowrap;
}

.text-segment-inner :deep(th) {
  font-weight: 600;
  color: #111;
}

.text-segment-inner :deep(td) {
  font-weight: 400;
  color: #222;
}

.text-segment-inner :deep(tr:last-child td) {
  border-bottom: none;
}

.text-segment-inner :deep(th[align='center']),
.text-segment-inner :deep(td[align='center']) {
  text-align: center;
}

.text-segment-inner :deep(th[align='right']),
.text-segment-inner :deep(td[align='right']) {
  text-align: right;
}

.text-segment-inner :deep(a) {
  color: #1677ff;
  text-decoration: underline;
  cursor: pointer;
}

.text-segment--clamped :deep(*) {
  color: inherit;
}

.text-segment--clamped :deep(a) {
  color: inherit;
}

.text-segment--clamped::-webkit-scrollbar {
  display: none;
  width: 0;
  height: 0;
}
</style>
