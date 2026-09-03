import { computed, ref } from 'vue'
import {
  uploadDocumentForChat,
  processDocumentForChat,
  createDocumentProcessWebSocket
} from '../services/knowledgeBaseApi.js'
import { MAX_CHAT_ATTACHMENTS } from '../utils/chatAttachments.js'

function formatSize (bytes) {
  if (!bytes && bytes !== 0) return ''
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(2)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

function makeId () {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`
}

function fileLabelFrom (name, sizeBytes) {
  const display = name || '文件'
  const ext = display.includes('.') ? display.split('.').pop().toUpperCase() : ''
  const sizeText =
    typeof sizeBytes === 'number' && sizeBytes >= 0 ? formatSize(sizeBytes) : ''
  return { name: display, sizeText, ext }
}

/**
 * 对话附件：upload →（可选）WS + process。就绪 = 有 storage_doc_uuid 且 process 完成。
 */
export function useChatFileUpload () {
  const items = ref([])
  const sockets = new Map()

  const allReady = computed(
    () => items.value.length === 0 || items.value.every((x) => x.phase === 'ready')
  )
  const hasBusy = computed(() =>
    items.value.some((x) => x.phase === 'uploading' || x.phase === 'processing')
  )
  const hasError = computed(() => items.value.some((x) => x.phase === 'error'))

  function clearItemWs (id) {
    const ws = sockets.get(id)
    if (!ws) return
    try {
      ws.close()
    } catch (e) {
      console.warn('[useChatFileUpload] ws close', e)
    }
    sockets.delete(id)
  }

  function clearAllWs () {
    for (const id of [...sockets.keys()]) {
      clearItemWs(id)
    }
  }

  function patchItem (id, patch) {
    const idx = items.value.findIndex((x) => x.id === id)
    if (idx < 0) return
    items.value.splice(idx, 1, { ...items.value[idx], ...patch })
  }

  function reset () {
    clearAllWs()
    items.value = []
  }

  function remove (id) {
    clearItemWs(id)
    items.value = items.value.filter((x) => x.id !== id)
  }

  function hasUuid (uuid) {
    return items.value.some((x) => x.storageDocUuid && x.storageDocUuid === uuid)
  }

  async function runProcessPipeline (id, uuid) {
    patchItem(id, { phase: 'processing', storageDocUuid: uuid, progress: 8 })
    const ws = await createDocumentProcessWebSocket(
      uuid,
      (p) => {
        const raw = p.progress
        let n = typeof raw === 'number' ? raw : parseFloat(raw)
        if (!Number.isNaN(n)) {
          if (n > 1) n = n / 100
          patchItem(id, { progress: Math.min(92, 8 + Math.round(n * 84)) })
        }
      },
      () => {
        patchItem(id, { progress: Math.max(items.value.find((x) => x.id === id)?.progress || 0, 92) })
      },
      (err) => {
        console.warn('[useChatFileUpload] WebSocket 错误（仍继续处理）', err)
      }
    )
    sockets.set(id, ws)
    await processDocumentForChat(uuid)
    clearItemWs(id)
    patchItem(id, { phase: 'ready', storageDocUuid: uuid, progress: 100 })
  }

  async function runItemPipeline (id, uuid, skipProcess) {
    if (skipProcess) {
      patchItem(id, { phase: 'ready', storageDocUuid: uuid, progress: 100 })
      return
    }
    await runProcessPipeline(id, uuid)
  }

  async function startUpload (file) {
    if (!file) return 'ok'
    if (items.value.length >= MAX_CHAT_ATTACHMENTS) return 'limit'
    const item = {
      id: makeId(),
      phase: 'uploading',
      storageDocUuid: null,
      progress: 3,
      fileLabel: fileLabelFrom(file.name, file.size),
      errorMessage: ''
    }
    items.value = [...items.value, item]
    try {
      const data = await uploadDocumentForChat(file)
      const uuid = data.storage_doc_uuid
      if (!uuid) {
        throw new Error('上传响应缺少 storage_doc_uuid')
      }
      if (hasUuid(uuid) && items.value.find((x) => x.id === item.id)?.storageDocUuid !== uuid) {
        remove(item.id)
        return 'dup'
      }
      const skipProcess = Number(data.processing_status) === 2
      await runItemPipeline(item.id, uuid, skipProcess)
      return 'ok'
    } catch (e) {
      console.error('[useChatFileUpload]', e)
      clearItemWs(item.id)
      patchItem(item.id, {
        phase: 'error',
        errorMessage: e.message || '上传或处理失败'
      })
      return 'error'
    }
  }

  async function startUploadMany (fileList) {
    const files = Array.from(fileList || []).filter(Boolean)
    let hitLimit = false
    for (const file of files) {
      if (items.value.length >= MAX_CHAT_ATTACHMENTS) {
        hitLimit = true
        break
      }
      const result = await startUpload(file)
      if (result === 'limit') {
        hitLimit = true
        break
      }
    }
    return hitLimit ? 'limit' : 'ok'
  }

  async function startFromHostUploadedDocument (uuid, displayName, fileSizeBytes, skipProcess = false) {
    if (!uuid) return 'ok'
    if (hasUuid(uuid)) return 'dup'
    if (items.value.length >= MAX_CHAT_ATTACHMENTS) return 'limit'
    const item = {
      id: makeId(),
      phase: 'processing',
      storageDocUuid: uuid,
      progress: 5,
      fileLabel: fileLabelFrom(displayName, fileSizeBytes),
      errorMessage: ''
    }
    items.value = [...items.value, item]
    try {
      await runItemPipeline(item.id, uuid, skipProcess)
      return 'ok'
    } catch (e) {
      console.error('[useChatFileUpload] host path', e)
      clearItemWs(item.id)
      patchItem(item.id, {
        phase: 'error',
        errorMessage: e.message || '处理失败'
      })
      return 'error'
    }
  }

  return {
    items,
    allReady,
    hasBusy,
    hasError,
    startUpload,
    startUploadMany,
    startFromHostUploadedDocument,
    remove,
    reset
  }
}
