<template>
  <div class="user-message">
    <div
      v-for="(att, i) in (message.attachments || [])"
      :key="att.storage_doc_uuid || att.fileName || i"
      class="attachment-bubble"
      :class="{ 'attachment-bubble--image': canPreview(att) }"
      :role="canPreview(att) ? 'button' : undefined"
      :tabindex="canPreview(att) ? 0 : undefined"
      :title="canPreview(att) ? '点击查看大图' : undefined"
      @click="openPreview(att)"
      @keydown.enter.prevent="openPreview(att)"
      @keydown.space.prevent="openPreview(att)"
    >
      <img
        v-if="isImageAtt(att) && thumbUrls[att.storage_doc_uuid]"
        :src="thumbUrls[att.storage_doc_uuid]"
        alt=""
        class="doc-icon doc-thumb"
        @error="onThumbError(att)"
      />
      <img
        v-else
        :src="attachmentIcon(att)"
        alt=""
        class="doc-icon"
        aria-hidden="true"
      />
      <div class="attach-text">
        <div class="attach-title" :title="att.fileName">
          {{ truncate(att.fileName) }}
        </div>
        <div class="attach-sub">
          {{ formatSub(att) }}
        </div>
      </div>
    </div>
    <div v-if="displayContent" class="user-text-wrap">
      <div
        ref="shellRef"
        class="message-shell"
        :class="{ 'message-shell--scrollable': overflowing }"
      >
        <div
          ref="contentRef"
          class="message-content"
          :class="{ 'message-content--scrollable': overflowing }"
          @scroll="syncThumb"
        >
          <div class="message-text">{{ displayContent }}</div>
        </div>
        <div
          v-if="overflowing"
          ref="trackRef"
          class="bubble-scroll-track"
          @pointerdown.stop="onTrackPointerDown"
        >
          <div class="bubble-scroll-thumb" :style="thumbStyle" />
        </div>
      </div>
      <MessageCopyBar :text="displayContent" align="right" />
    </div>
    <ImageLightbox
      v-if="preview"
      :src="preview.src"
      :title="preview.title"
      @close="preview = null"
    />
  </div>
</template>

<script setup>
import { computed, nextTick, onMounted, onUnmounted, ref, watch } from 'vue'
import { stripChatAttachmentsAppendix, stripVisionPersistNotes } from '../utils/chatAttachments.js'
import { stripSelectedOpenFilesAppendix } from '../utils/selectedOpenFiles.js'
import { resolveFileAppIconByName } from '../utils/openFileAppIcon.js'
import { isImageFileName, loadKbImageThumbUrl } from '../utils/kbImageThumb.js'
import ImageLightbox from './ImageLightbox.vue'
import MessageCopyBar from './MessageCopyBar.vue'

const props = defineProps({
  message: {
    type: Object,
    required: true
  }
})

const displayContent = computed(() => {
  const raw = props.message?.content
  if (typeof raw !== 'string') return ''
  return stripVisionPersistNotes(
    stripSelectedOpenFilesAppendix(stripChatAttachmentsAppendix(raw))
  )
})

const thumbUrls = ref({})
const preview = ref(null)
const contentRef = ref(null)
const shellRef = ref(null)
const trackRef = ref(null)
const overflowing = ref(false)
const thumb = ref({ height: 28, top: 0 })
let blobUrls = []
let dragState = null
let overflowObserver = null

const thumbStyle = computed(() => ({
  height: `${thumb.value.height}px`,
  transform: `translateY(${thumb.value.top}px)`
}))

function refreshOverflow () {
  const el = contentRef.value
  if (!el) {
    overflowing.value = false
    return
  }
  const next = el.scrollHeight - el.clientHeight > 1
  overflowing.value = next
  nextTick(() => {
    const node = contentRef.value
    if (!node) return
    overflowing.value = node.scrollHeight - node.clientHeight > 1
    if (overflowing.value) syncThumb()
  })
}

function trackSize () {
  return trackRef.value?.clientHeight || contentRef.value?.clientHeight || 0
}

function syncThumb () {
  const el = contentRef.value
  if (!el) return
  const view = el.clientHeight
  const total = el.scrollHeight
  const trackH = trackSize()
  if (total <= view + 1 || trackH <= 0) {
    thumb.value = { height: Math.max(28, trackH), top: 0 }
    return
  }
  const height = Math.max(28, (view / total) * trackH)
  const maxTop = Math.max(0, trackH - height)
  const top = (el.scrollTop / (total - view)) * maxTop
  thumb.value = { height, top }
}

function scrollFromThumbTop (top) {
  const el = contentRef.value
  if (!el) return
  const view = el.clientHeight
  const total = el.scrollHeight
  const trackH = trackSize()
  const maxTop = Math.max(1, trackH - thumb.value.height)
  const ratio = Math.min(1, Math.max(0, top / maxTop))
  el.scrollTop = ratio * (total - view)
}

function onTrackPointerDown (e) {
  const track = e.currentTarget
  const rect = track.getBoundingClientRect()
  const y = e.clientY - rect.top
  const { height, top } = thumb.value
  if (y < top || y > top + height) {
    scrollFromThumbTop(y - height / 2)
    syncThumb()
  }
  dragState = {
    pointerId: e.pointerId,
    startY: e.clientY,
    startTop: thumb.value.top
  }
  track.setPointerCapture(e.pointerId)
  window.addEventListener('pointermove', onDragMove)
  window.addEventListener('pointerup', onDragUp)
}

function onDragMove (e) {
  if (!dragState) return
  scrollFromThumbTop(dragState.startTop + (e.clientY - dragState.startY))
}

function onDragUp () {
  dragState = null
  window.removeEventListener('pointermove', onDragMove)
  window.removeEventListener('pointerup', onDragUp)
}

watch(displayContent, () => nextTick(refreshOverflow))

watch(overflowing, (on) => {
  if (on) nextTick(syncThumb)
})

/** 超高时先滚气泡，到顶/底再交给外层列表 */
function onBubbleWheel (e) {
  const el = contentRef.value
  if (!el) return
  const overflow = el.scrollHeight - el.clientHeight
  if (overflow <= 1) return

  let dy = e.deltaY
  if (e.deltaMode === 1) dy *= 16
  if (e.deltaMode === 2) dy *= el.clientHeight

  const atTop = el.scrollTop <= 0
  const atBottom = el.scrollTop >= overflow - 1
  const scrollingUp = dy < 0
  const scrollingDown = dy > 0

  e.preventDefault()
  e.stopPropagation()

  if ((scrollingUp && !atTop) || (scrollingDown && !atBottom)) {
    el.scrollTop += dy
    return
  }

  const list = el.closest('.chat-messages')
  if (list) list.scrollTop += dy
}

function isImageAtt (att) {
  return isImageFileName(att?.fileName, att?.ext)
}

function canPreview (att) {
  return !!(isImageAtt(att) && att?.storage_doc_uuid && thumbUrls.value[att.storage_doc_uuid])
}

function openPreview (att) {
  if (!canPreview(att)) return
  preview.value = {
    src: thumbUrls.value[att.storage_doc_uuid],
    title: att.fileName || '图片'
  }
}

function attachmentIcon (att) {
  return resolveFileAppIconByName(att?.fileName, att?.ext)
}

function revokeThumbs () {
  blobUrls.forEach((u) => {
    try { URL.revokeObjectURL(u) } catch (_) { /* ignore */ }
  })
  blobUrls = []
}

async function loadThumbs (attachments) {
  revokeThumbs()
  const next = {}
  await Promise.all((attachments || []).map(async (att) => {
    if (!isImageAtt(att) || !att.storage_doc_uuid) return
    try {
      const url = await loadKbImageThumbUrl(att.storage_doc_uuid, att.knowledge_base_uuid)
      blobUrls.push(url)
      next[att.storage_doc_uuid] = url
    } catch (e) {
      console.warn('[UserMessageBubble] 缩略图加载失败', att.fileName, e?.message || e)
    }
  }))
  thumbUrls.value = next
}

function onThumbError (att) {
  if (!att?.storage_doc_uuid) return
  const copy = { ...thumbUrls.value }
  delete copy[att.storage_doc_uuid]
  thumbUrls.value = copy
}

watch(
  () => props.message?.attachments,
  (atts) => { loadThumbs(atts) },
  { immediate: true, deep: true }
)

onMounted(() => {
  const el = contentRef.value
  if (el) {
    el.addEventListener('wheel', onBubbleWheel, { passive: false })
    if (typeof ResizeObserver !== 'undefined') {
      overflowObserver = new ResizeObserver(() => refreshOverflow())
      overflowObserver.observe(el)
    }
  }
  nextTick(refreshOverflow)
})

onUnmounted(() => {
  const el = contentRef.value
  if (el) el.removeEventListener('wheel', onBubbleWheel)
  if (overflowObserver) {
    overflowObserver.disconnect()
    overflowObserver = null
  }
  onDragUp()
  revokeThumbs()
})

function truncate (name) {
  if (!name) return ''
  return name.length > 28 ? name.slice(0, 25) + '…' : name
}

function formatSub (att) {
  const ext = att.ext || ''
  const size = att.sizeText || ''
  if (ext && size) return `${ext}  ${size}`
  return ext || size || ''
}
</script>

<style scoped>
.user-message {
  display: flex;
  flex-direction: column;
  align-items: flex-end;
  gap: 6px;
  width: 100%;
  min-width: 0;
  margin-bottom: 4px;
}

.attachment-bubble {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  max-width: 85%;
  padding: 8px 12px;
  background-color: #f0f0f0;
  color: #333;
  border-radius: 10px;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.06);
}

.attachment-bubble--image {
  cursor: zoom-in;
}

.attachment-bubble--image:hover {
  background-color: #e8e8e8;
}

.doc-icon {
  width: 18px;
  height: 18px;
  object-fit: contain;
  flex-shrink: 0;
  display: block;
  margin-top: 1px;
}

.doc-thumb {
  width: 40px;
  height: 40px;
  object-fit: cover;
  border-radius: 4px;
}

.attach-text {
  min-width: 0;
  text-align: left;
}

.attach-title {
  font-size: 13px;
  font-weight: 500;
  word-break: break-all;
}

.attach-sub {
  font-size: 12px;
  color: #888;
  margin-top: 2px;
}

.user-text-wrap {
  display: inline-flex;
  flex-direction: column;
  align-items: stretch;
  max-width: 100%;
  min-width: 0;
}

.user-text-wrap:hover :deep(.message-copy-bar),
.user-text-wrap:focus-within :deep(.message-copy-bar) {
  opacity: 1;
  pointer-events: auto;
}

.message-shell {
  display: flex;
  flex-direction: row;
  align-items: stretch;
  max-width: 100%;
  max-height: 160px;
  min-width: 0;
  background-color: #1890ff;
  color: #fff;
  border-radius: 8px;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
  overflow: hidden;
}

.message-content {
  box-sizing: border-box;
  flex: 1;
  max-width: 100%;
  max-height: 160px;
  min-width: 0;
  min-height: 0;
  overflow-x: hidden;
  overflow-y: hidden;
  background: transparent;
  padding: 10px 14px;
  word-wrap: break-word;
  scrollbar-width: none;
}

.message-content--scrollable {
  overflow-y: auto;
}

.message-content::-webkit-scrollbar,
.message-content--scrollable::-webkit-scrollbar {
  width: 0;
  height: 0;
  display: none;
}

.bubble-scroll-track {
  flex: 0 0 12px;
  width: 12px;
  margin: 8px 6px 8px 0;
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.22);
  position: relative;
  cursor: pointer;
}

.bubble-scroll-thumb {
  position: absolute;
  left: 1px;
  right: 1px;
  top: 0;
  border-radius: 7px;
  background: rgba(80, 80, 80, 0.45);
  pointer-events: none;
}

.message-text {
  font-family: "Segoe UI", "Microsoft YaHei UI", "Microsoft YaHei", sans-serif;
  font-size: 14px;
  font-weight: 500;
  line-height: 1.7;
  white-space: pre-wrap;
  text-align: left;
}
</style>
