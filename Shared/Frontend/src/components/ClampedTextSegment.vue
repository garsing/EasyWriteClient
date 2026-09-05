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

const html = computed(() => {
  const text = normalizeTextContent(props.content)
  if (!text) return ''
  try {
    return marked.parse(text)
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

.text-segment--clamped :deep(*) {
  color: inherit;
}

.text-segment--clamped::-webkit-scrollbar {
  display: none;
  width: 0;
  height: 0;
}
</style>
