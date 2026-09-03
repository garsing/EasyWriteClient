<template>
  <div class="user-message">
    <div
      v-for="(att, i) in (message.attachments || [])"
      :key="att.storage_doc_uuid || att.fileName || i"
      class="attachment-bubble"
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
    <div class="message-content">
      <div class="message-text">{{ message.content }}</div>
    </div>
  </div>
</template>

<script setup>
import { onUnmounted, ref, watch } from 'vue'
import { resolveFileAppIconByName } from '../utils/openFileAppIcon.js'
import { isImageFileName, loadKbImageThumbUrl } from '../utils/kbImageThumb.js'

const props = defineProps({
  message: {
    type: Object,
    required: true
  }
})

const thumbUrls = ref({})
let blobUrls = []

function isImageAtt (att) {
  return isImageFileName(att?.fileName, att?.ext)
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

onUnmounted(revokeThumbs)

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

.message-content {
  max-width: 85%;
  background-color: #1890ff;
  color: #fff;
  padding: 10px 14px;
  border-radius: 8px;
  word-wrap: break-word;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

.message-text {
  font-size: 14px;
  line-height: 1.5;
  white-space: pre-wrap;
  text-align: left;
}
</style>
