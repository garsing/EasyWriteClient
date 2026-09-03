<template>
  <div class="user-message">
    <div
      v-for="(att, i) in (message.attachments || [])"
      :key="att.storage_doc_uuid || att.fileName || i"
      class="attachment-bubble"
    >
      <img
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
import { resolveFileAppIconByName } from '../utils/openFileAppIcon.js'

defineProps({
  message: {
    type: Object,
    required: true
  }
})

function attachmentIcon (att) {
  return resolveFileAppIconByName(att?.fileName, att?.ext)
}

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
