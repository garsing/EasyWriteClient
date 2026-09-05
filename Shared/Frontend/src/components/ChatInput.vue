<template>
  <div
    class="chat-input-container"
    :class="{ 'chat-input-container--desktop': desktop }"
    @dragover.capture="onFileDragOverCapture"
    @drop.capture="onFileDropCapture"
    @paste.capture="onPasteCapture"
  >
    <div v-if="showAttachmentStrip" class="attachment-strip">
      <div
        v-for="item in pendingAttachments"
        :key="item.id"
        class="attachment-card"
        :class="{ 'attachment-card--image': canPreview(item) }"
        :role="canPreview(item) ? 'button' : undefined"
        :tabindex="canPreview(item) ? 0 : undefined"
        :title="canPreview(item) ? '点击查看大图' : undefined"
        @click="openPreview(item)"
        @keydown.enter.prevent="openPreview(item)"
        @keydown.space.prevent="openPreview(item)"
      >
        <img
          v-if="thumbSrc(item)"
          :src="thumbSrc(item)"
          alt=""
          class="attach-icon attach-thumb"
          @error="onThumbError(item)"
        />
        <img
          v-else
          :src="itemIcon(item)"
          alt=""
          class="attach-icon"
          aria-hidden="true"
        />
        <div class="attach-meta">
          <div class="attach-name" :title="item.fileLabel?.name">{{ truncateName(item.fileLabel?.name) }}</div>
          <div v-if="isItemBusy(item)" class="attach-progress">
            <span class="spin" />
            <span>{{ itemProgressLabel(item) }}</span>
          </div>
          <div v-else-if="item.phase === 'ready'" class="attach-ready">就绪</div>
          <div v-else-if="item.phase === 'error'" class="attach-error">{{ item.errorMessage || '失败' }}</div>
        </div>
        <button
          v-if="!isItemBusy(item)"
          type="button"
          class="attach-remove"
          @click.stop="$emit('remove-attachment', item.id)"
        >
          ×
        </button>
      </div>
    </div>
    <ImageLightbox
      v-if="preview"
      :src="preview.src"
      :title="preview.title"
      @close="preview = null"
    />

    <!-- Desktop：外框与内框之间 — 已选打开文件芯片 -->
    <div v-if="showSelectedOpenFiles" class="selected-open-files-strip">
      <div
        v-for="f in selectedOpenFiles"
        :key="f.id"
        class="selected-open-file-chip"
        :title="f.displayName"
      >
        <img :src="chipAppIcon(f)" alt="" class="chip-app-icon" aria-hidden="true" />
        <span class="chip-name">{{ truncateName(f.displayName) }}</span>
        <button
          type="button"
          class="chip-remove"
          :aria-label="'取消选中 ' + (f.displayName || '')"
          @click="$emit('remove-selected-open-file', f.id)"
        >
          ×
        </button>
      </div>
    </div>

    <div class="input-wrapper">
      <a-textarea
        v-model:value="inputValue"
        :placeholder="placeholder"
        :auto-size="autoSize"
        :disabled="loading"
        class="input-textarea"
        @keydown.enter.exact.prevent="handleEnter"
        @keydown.shift.enter.exact="handleShiftEnter"
      />
      <button
        v-if="desktop || inputValue.trim().length > 0 || isProcessing"
        type="button"
        :disabled="sendDisabled"
        :title="isProcessing ? '停止' : '发送'"
        :aria-label="isProcessing ? '停止' : '发送'"
        @click="handleButtonClick"
        class="send-button"
      >
        <img :src="currentIcon" :alt="isProcessing ? '停止' : '发送'" class="send-icon" />
      </button>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, onUnmounted, watch } from 'vue'
import submitIcon from '../assets/images/submit.png'
import stopIcon from '../assets/images/stop.png'
import { resolveFileAppIconByName, resolveOpenFileAppIcon } from '../utils/openFileAppIcon.js'
import { useWebViewBridge } from '../composables/useWebViewBridge'
import { syncLiveDraftInput } from '../utils/draftChatInput.js'
import { collectClipboardFiles } from '../utils/clipboardChatImages.js'
import { isImageFileName, loadKbImageThumbUrl } from '../utils/kbImageThumb.js'
import ImageLightbox from './ImageLightbox.vue'

const { sendMessage } = useWebViewBridge()

const props = defineProps({
  loading: {
    type: Boolean,
    default: false
  },
  isProcessing: {
    type: Boolean,
    default: false
  },
  /** 桌面端加高输入区；插件侧栏保持紧凑 */
  desktop: {
    type: Boolean,
    default: false
  },
  /** 输入区待发送附件 items[] */
  attachments: {
    type: Array,
    default: () => []
  },
  /** Desktop 已选打开文件芯片 */
  selectedOpenFiles: {
    type: Array,
    default: () => []
  },
  /** 是否展示打开文件芯片（与 desktop 密度解耦；缩小版仍可显示） */
  showOpenFileChips: {
    type: Boolean,
    default: false
  }
})

const autoSize = computed(() =>
  props.desktop ? { minRows: 3, maxRows: 8 } : { minRows: 1, maxRows: 4 }
)

const emit = defineEmits(['send', 'stop', 'remove-attachment', 'dropped-files', 'remove-selected-open-file'])

const showSelectedOpenFiles = computed(
  () =>
    props.showOpenFileChips &&
    Array.isArray(props.selectedOpenFiles) &&
    props.selectedOpenFiles.length > 0
)

function chipAppIcon (f) {
  return resolveOpenFileAppIcon(f)
}

const pendingAttachments = computed(() =>
  Array.isArray(props.attachments) ? props.attachments : []
)

function itemIcon (item) {
  return resolveFileAppIconByName(item?.fileLabel?.name, item?.fileLabel?.ext)
}

function isImageItem (item) {
  return isImageFileName(item?.fileLabel?.name, item?.fileLabel?.ext)
}

const kbThumbUrls = ref({})
const thumbFailed = ref({})
const preview = ref(null)
let kbBlobUrls = []

function revokeKbThumbs () {
  kbBlobUrls.forEach((u) => {
    try { URL.revokeObjectURL(u) } catch (_) { /* ignore */ }
  })
  kbBlobUrls = []
}

function thumbSrc (item) {
  if (!item || thumbFailed.value[item.id]) return ''
  if (item.localPreviewUrl) return item.localPreviewUrl
  if (item.storageDocUuid && kbThumbUrls.value[item.storageDocUuid]) {
    return kbThumbUrls.value[item.storageDocUuid]
  }
  return ''
}

function canPreview (item) {
  return !!(isImageItem(item) && thumbSrc(item))
}

function openPreview (item) {
  if (!canPreview(item)) return
  preview.value = {
    src: thumbSrc(item),
    title: item.fileLabel?.name || '图片'
  }
}

function onThumbError (item) {
  if (!item?.id) return
  thumbFailed.value = { ...thumbFailed.value, [item.id]: true }
}

async function loadKbThumbs (list) {
  const next = { ...kbThumbUrls.value }
  await Promise.all((list || []).map(async (item) => {
    if (!isImageItem(item) || item.localPreviewUrl || !item.storageDocUuid) return
    if (item.phase !== 'ready') return
    if (next[item.storageDocUuid]) return
    try {
      const url = await loadKbImageThumbUrl(item.storageDocUuid, item.knowledgeBaseUuid)
      kbBlobUrls.push(url)
      next[item.storageDocUuid] = url
    } catch (e) {
      console.warn('[ChatInput] 附件缩略图加载失败', item.fileLabel?.name, e?.message || e)
    }
  }))
  kbThumbUrls.value = next
}

watch(
  () => pendingAttachments.value.map((x) => `${x.id}:${x.phase}:${x.storageDocUuid || ''}:${x.localPreviewUrl || ''}`).join('|'),
  () => { loadKbThumbs(pendingAttachments.value) }
)

onUnmounted(() => {
  revokeKbThumbs()
})

function isItemBusy (item) {
  return item && (item.phase === 'uploading' || item.phase === 'processing')
}

function itemProgressLabel (item) {
  const p = item?.progress
  if (typeof p === 'number' && !Number.isNaN(p)) {
    return `${Math.min(100, Math.round(p))}%`
  }
  return '处理中…'
}

function dataTransferHasFiles (dt) {
  if (!dt) return false
  try {
    if (dt.files && dt.files.length > 0) return true
    if (dt.items && dt.items.length) {
      for (let i = 0; i < dt.items.length; i++) {
        if (dt.items[i].kind === 'file') return true
      }
    }
    if (!dt.types) return false
    if (typeof dt.types.includes === 'function' && dt.types.includes('Files')) return true
    for (let j = 0; j < dt.types.length; j++) {
      const t = String(dt.types[j]).toLowerCase()
      if (t === 'files' || t.includes('file')) return true
    }
  } catch (_) {
    /* ignore */
  }
  return false
}

/** 捕获阶段：在落到 textarea 前拦截「文件」拖放，避免 WebView2 走打开/下载默认行为 */
function onFileDragOverCapture (e) {
  if (!dataTransferHasFiles(e.dataTransfer)) return
  e.preventDefault()
  try {
    e.dataTransfer.dropEffect = 'copy'
  } catch (_) {
    /* ignore */
  }
}

function onFileDropCapture (e) {
  if (!dataTransferHasFiles(e.dataTransfer)) return
  e.preventDefault()
  e.stopPropagation()
  const files = Array.from(e.dataTransfer?.files || []).filter(Boolean)
  if (files.length) {
    emit('dropped-files', files)
  }
}

/** 剪贴板图片或文件走与拖放相同的上传链路；纯文本粘贴不拦截 */
function onPasteCapture (e) {
  const files = collectClipboardFiles(e.clipboardData)
  if (!files.length) return
  e.preventDefault()
  e.stopPropagation()
  emit('dropped-files', files)
}

const inputValue = ref('')
const placeholder = ref('输入消息')

watch(inputValue, (v) => {
  syncLiveDraftInput(v)
}, { immediate: true })

const showAttachmentStrip = computed(() => pendingAttachments.value.length > 0)

const sendDisabled = computed(() => {
  if (props.isProcessing) return false
  if (props.loading && !props.isProcessing) return true
  const text = inputValue.value.trim()
  if (!text) return true
  if (!showAttachmentStrip.value) {
    return props.loading
  }
  if (pendingAttachments.value.some((x) => isItemBusy(x) || x.phase === 'error')) {
    return true
  }
  return props.loading
})

const canSend = computed(() => !sendDisabled.value && inputValue.value.trim().length > 0)

const currentIcon = computed(() => {
  return props.isProcessing ? stopIcon : submitIcon
})

function truncateName (name) {
  if (!name) return ''
  return name.length > 36 ? name.slice(0, 33) + '…' : name
}

onMounted(() => {
  window.addEventListener('restoreInput', (event) => {
    if (event.detail && typeof event.detail.value === 'string') {
      inputValue.value = event.detail.value
    }
  })

  window.addEventListener('clearInput', () => {
    inputValue.value = ''
  })
})


const handleSend = () => {
  if (!canSend.value) return

  const content = inputValue.value.trim()
  if (content) {
    emit('send', content)
  }
}

const handleStop = async () => {
  emit('stop')
  try {
    await sendMessage('stopRequest', {})
  } catch (error) {
    console.error('停止请求失败:', error)
  }
}

const handleButtonClick = () => {
  if (props.isProcessing) {
    handleStop()
  } else {
    handleSend()
  }
}

const handleEnter = () => {
  if (!props.isProcessing) {
    handleSend()
  }
}

const handleShiftEnter = () => {}
</script>

<style scoped>
/* 插件默认：紧凑单行输入 */
.chat-input-container {
  padding: 12px 8px 12px 16px;
  background-color: #f7f7f5;
  border-top: 1px solid transparent;
  flex-shrink: 0;
}

.attachment-strip {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin-bottom: 8px;
}

.selected-open-files-strip {
  display: flex;
  flex-wrap: wrap;
  align-content: flex-start;
  gap: 6px;
  margin-bottom: 8px;
  max-width: 100%;
  /* 约两行芯片（单行高 ~28px + 行间距 6px） */
  max-height: calc(28px * 2 + 6px);
  overflow-x: hidden;
  overflow-y: auto;
  scrollbar-width: thin;
}

.selected-open-files-strip::-webkit-scrollbar {
  width: 4px;
}

.selected-open-files-strip::-webkit-scrollbar-thumb {
  background: rgba(0, 0, 0, 0.18);
  border-radius: 2px;
}

.selected-open-file-chip {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  max-width: 100%;
  padding: 4px 6px 4px 8px;
  border-radius: 6px;
  background: #ebeae6;
  min-width: 0;
}

.chip-app-icon {
  width: 14px;
  height: 14px;
  object-fit: contain;
  flex-shrink: 0;
  display: block;
}

.chip-name {
  font-size: 12px;
  color: #1f1e1c;
  max-width: 160px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  min-width: 0;
}

.chip-remove {
  flex-shrink: 0;
  border: none;
  background: transparent;
  font-size: 16px;
  line-height: 1;
  cursor: pointer;
  color: #8a877f;
  padding: 0 2px;
}

.chip-remove:hover {
  color: #1f1e1c;
}

.attachment-card {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  padding: 8px 10px;
  background: #f5f5f5;
  border-radius: 8px;
  max-width: 100%;
}

.attachment-card--image {
  cursor: zoom-in;
}

.attachment-card--image:hover {
  background: #ececec;
}

.attach-icon {
  width: 18px;
  height: 18px;
  object-fit: contain;
  flex-shrink: 0;
  display: block;
  margin-top: 1px;
}

.attach-thumb {
  width: 40px;
  height: 40px;
  object-fit: cover;
  border-radius: 4px;
}

.attach-meta {
  flex: 1;
  min-width: 0;
}

.attach-name {
  font-size: 13px;
  font-weight: 500;
  color: #333;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.attach-progress {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  color: #666;
  margin-top: 4px;
}

.spin {
  width: 14px;
  height: 14px;
  border: 2px solid #ccc;
  border-top-color: #1890ff;
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}

@keyframes spin {
  to {
    transform: rotate(360deg);
  }
}

.attach-ready {
  font-size: 12px;
  color: #52c41a;
  margin-top: 4px;
}

.attach-error {
  font-size: 12px;
  color: #ff4d4f;
  margin-top: 4px;
}

.attach-remove {
  flex-shrink: 0;
  border: none;
  background: transparent;
  font-size: 18px;
  line-height: 1;
  cursor: pointer;
  color: #999;
  padding: 0 4px;
}

.attach-remove:hover {
  color: #333;
}

.input-wrapper {
  display: flex;
  gap: 4px;
  align-items: center;
}

.input-textarea {
  flex: 1;
  min-width: 0;
}

.input-textarea :deep(textarea.ant-input),
.input-textarea :deep(.ant-input) {
  background-color: #fff;
  border-color: #e0e0e0;
}

.send-button {
  flex-shrink: 0;
  background: none;
  border: none;
  padding: 6px;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
}

.send-button:disabled {
  cursor: not-allowed;
  opacity: 0.5;
}

.send-icon {
  width: 24px;
  height: 24px;
  object-fit: contain;
}

/* 桌面端：外层与面板同色，仅内框白色；外层 padding 收紧使内框更贴边 */
.chat-input-container--desktop {
  padding: 6px 10px 10px;
  background-color: #f7f7f5;
}

.chat-input-container--desktop .input-wrapper {
  gap: 8px;
  /* 顶对齐，保证「输入消息」距上边与左边空隙一致 */
  align-items: flex-start;
  min-height: 96px;
  padding: 12px;
  box-sizing: border-box;
  background: #ffffff;
  border: 1px solid #e6e6e6;
  border-radius: 16px;
}

.chat-input-container--desktop .input-textarea {
  border: none !important;
  box-shadow: none !important;
  background: transparent !important;
  margin: 0 !important;
  padding: 0 !important;
}

.chat-input-container--desktop .input-wrapper :deep(.ant-input-textarea),
.chat-input-container--desktop .input-wrapper :deep(.ant-input),
.chat-input-container--desktop .input-wrapper :deep(textarea.ant-input),
.chat-input-container--desktop .input-wrapper :deep(.ant-input-affix-wrapper) {
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
  outline: none !important;
  margin: 0 !important;
}

.chat-input-container--desktop .input-textarea :deep(.ant-input),
.chat-input-container--desktop .input-textarea :deep(textarea.ant-input),
.chat-input-container--desktop .input-textarea :deep(.ant-input-textarea textarea) {
  padding: 0 !important;
  text-indent: 0 !important;
  min-height: 60px;
  resize: none;
  font-size: 14px;
  line-height: 1.55;
}

.chat-input-container--desktop .input-textarea :deep(.ant-input:hover),
.chat-input-container--desktop .input-textarea :deep(.ant-input:focus),
.chat-input-container--desktop .input-textarea :deep(.ant-input-focused),
.chat-input-container--desktop .input-textarea :deep(.ant-input-affix-wrapper:hover),
.chat-input-container--desktop .input-textarea :deep(.ant-input-affix-wrapper-focused),
.chat-input-container--desktop .input-textarea :deep(.ant-input-affix-wrapper:focus) {
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
  outline: none !important;
}

.chat-input-container--desktop .send-button {
  padding: 4px;
  /* 发送钮仍靠右下 */
  align-self: flex-end;
  margin-top: auto;
}
</style>
