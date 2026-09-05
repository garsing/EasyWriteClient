<template>
  <Teleport to="body">
    <div
      class="lb-mask"
      role="dialog"
      aria-modal="true"
      :aria-label="title || '图片预览'"
      @click.self="close"
      @wheel.prevent="onWheel"
    >
      <div class="lb-toolbar">
        <button type="button" class="lb-btn" title="放大" aria-label="放大" @click="zoomBy(1.25)">
          <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="6.5" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M11 8v6M8 11h6M16.5 16.5l4 4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>
        </button>
        <button type="button" class="lb-btn" title="缩小" aria-label="缩小" @click="zoomBy(0.8)">
          <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="6.5" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M8 11h6M16.5 16.5l4 4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>
        </button>
        <button type="button" class="lb-btn" title="原始比例" aria-label="原始比例" @click="resetZoom">
          <span class="lb-11">1:1</span>
        </button>
        <button type="button" class="lb-btn" title="下载" aria-label="下载" @click="download">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v10M8 10l4 4 4-4M6 18h12" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
        </button>
        <button type="button" class="lb-btn" title="关闭" aria-label="关闭" @click="close">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 7l10 10M17 7L7 17" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>
        </button>
      </div>
      <div
        class="lb-stage"
        :class="{ 'lb-stage--grab': scale > 1, 'lb-stage--grabbing': dragging }"
        @pointerdown="onPointerDown"
        @pointermove="onPointerMove"
        @pointerup="onPointerUp"
        @pointercancel="onPointerUp"
      >
        <img
          :src="src"
          :alt="title || ''"
          class="lb-img"
          :style="imgStyle"
          draggable="false"
          @click.stop
        />
      </div>
    </div>
  </Teleport>
</template>

<script setup>
import { computed, onMounted, onUnmounted, ref } from 'vue'

const props = defineProps({
  src: { type: String, required: true },
  title: { type: String, default: '' }
})

const emit = defineEmits(['close'])

const MIN = 0.25
const MAX = 8
const scale = ref(1)
const panX = ref(0)
const panY = ref(0)
const dragging = ref(false)
let dragStart = null

const imgStyle = computed(() => ({
  transform: `translate(${panX.value}px, ${panY.value}px) scale(${scale.value})`
}))

function clampScale (n) {
  return Math.min(MAX, Math.max(MIN, n))
}

function zoomBy (factor) {
  scale.value = clampScale(scale.value * factor)
}

function resetZoom () {
  scale.value = 1
  panX.value = 0
  panY.value = 0
}

function close () {
  emit('close')
}

function onWheel (e) {
  zoomBy(e.deltaY < 0 ? 1.12 : 0.89)
}

function onPointerDown (e) {
  if (scale.value <= 1) return
  dragging.value = true
  dragStart = { x: e.clientX - panX.value, y: e.clientY - panY.value, id: e.pointerId }
  e.currentTarget.setPointerCapture?.(e.pointerId)
}

function onPointerMove (e) {
  if (!dragging.value || !dragStart) return
  panX.value = e.clientX - dragStart.x
  panY.value = e.clientY - dragStart.y
}

function onPointerUp () {
  dragging.value = false
  dragStart = null
}

function download () {
  const a = document.createElement('a')
  a.href = props.src
  a.download = props.title || 'image.jpg'
  a.rel = 'noopener'
  document.body.appendChild(a)
  a.click()
  a.remove()
}

function onKeydown (e) {
  if (e.key === 'Escape') {
    e.preventDefault()
    close()
  }
}

onMounted(() => {
  window.addEventListener('keydown', onKeydown)
})

onUnmounted(() => {
  window.removeEventListener('keydown', onKeydown)
})
</script>

<style scoped>
.lb-mask {
  position: fixed;
  inset: 0;
  z-index: 4000;
  background: rgba(12, 12, 12, 0.86);
  display: flex;
  align-items: center;
  justify-content: center;
}

.lb-toolbar {
  position: absolute;
  top: 12px;
  right: 12px;
  display: flex;
  gap: 4px;
  z-index: 1;
}

.lb-btn {
  width: 34px;
  height: 34px;
  border: none;
  border-radius: 6px;
  background: rgba(255, 255, 255, 0.08);
  color: #f3f3f3;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 0;
}

.lb-btn:hover {
  background: rgba(255, 255, 255, 0.18);
}

.lb-btn svg {
  width: 18px;
  height: 18px;
  display: block;
}

.lb-11 {
  font-size: 11px;
  font-weight: 600;
  letter-spacing: -0.04em;
  line-height: 1;
}

.lb-stage {
  width: 100%;
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: hidden;
}

.lb-stage--grab {
  cursor: grab;
}

.lb-stage--grabbing {
  cursor: grabbing;
}

.lb-img {
  max-width: 88%;
  max-height: 88%;
  object-fit: contain;
  user-select: none;
  transform-origin: center center;
  pointer-events: none;
}
</style>
