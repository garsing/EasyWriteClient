import { ref } from 'vue'
import {
  uploadDocumentForChat,
  processDocumentForChat,
  createDocumentProcessWebSocket
} from '../services/knowledgeBaseApi.js'
import { getChatUploadByDocument } from '../services/chatUploadApi.js'

function formatSize (bytes) {
  if (!bytes && bytes !== 0) return ''
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(2)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

/**
 * 对话附件：upload → WS → process → GET by-document；
 * 若 index 已处理完成（processing_status=2）则跳过 process。
 */
export function useChatFileUpload () {
  const phase = ref('idle')
  const errorMessage = ref('')
  const storageDocUuid = ref(null)
  const uploadId = ref(null)
  const progress = ref(0)
  const fileLabel = ref({ name: '', sizeText: '', ext: '' })
  let wsConnection = null

  function clearWebSocket () {
    if (wsConnection) {
      try {
        wsConnection.close()
      } catch (e) {
        console.warn('[useChatFileUpload] ws close', e)
      }
      wsConnection = null
    }
  }

  function reset () {
    phase.value = 'idle'
    errorMessage.value = ''
    storageDocUuid.value = null
    uploadId.value = null
    progress.value = 0
    fileLabel.value = { name: '', sizeText: '', ext: '' }
    clearWebSocket()
  }

  async function completeRegistration (uuid) {
    phase.value = 'registering'
    progress.value = 96
    const reg = await getChatUploadByDocument(uuid)
    uploadId.value = reg.upload_id
    if (!uploadId.value) {
      throw new Error('未返回 upload_id')
    }
    progress.value = 100
    phase.value = 'ready'
  }

  async function runSkipProcessPipeline (uuid) {
    storageDocUuid.value = uuid
    progress.value = Math.max(progress.value, 50)
    try {
      await completeRegistration(uuid)
    } catch (firstErr) {
      console.log('[useChatFileUpload] 已处理 index：首次 by-document 未就绪，等待处理完成', firstErr)
      phase.value = 'processing'
      progress.value = 8
      await waitProcessingCompleteWebSocket(uuid)
      clearWebSocket()
      await completeRegistration(uuid)
    }
  }

  function waitProcessingCompleteWebSocket (uuid) {
    return new Promise((resolve, reject) => {
      createDocumentProcessWebSocket(
        uuid,
        (p) => {
          const raw = p.progress
          let n = typeof raw === 'number' ? raw : parseFloat(raw)
          if (!Number.isNaN(n)) {
            if (n > 1) n = n / 100
            progress.value = Math.min(92, 8 + Math.round(n * 84))
          }
        },
        () => {
          progress.value = Math.max(progress.value, 92)
          resolve()
        },
        (err) => {
          reject(new Error(err?.error || '文档处理失败'))
        }
      )
        .then((ws) => {
          wsConnection = ws
        })
        .catch(reject)
    })
  }

  async function runProcessRegisterPipeline (uuid) {
    storageDocUuid.value = uuid

    phase.value = 'processing'
    progress.value = 8

    wsConnection = await createDocumentProcessWebSocket(
      uuid,
      (p) => {
        const raw = p.progress
        let n = typeof raw === 'number' ? raw : parseFloat(raw)
        if (!Number.isNaN(n)) {
          if (n > 1) n = n / 100
          progress.value = Math.min(92, 8 + Math.round(n * 84))
        }
      },
      () => {
        progress.value = Math.max(progress.value, 92)
      },
      (err) => {
        console.warn('[useChatFileUpload] WebSocket 错误（仍继续处理）', err)
      }
    )

    await processDocumentForChat(uuid)
    progress.value = 96
    clearWebSocket()

    await completeRegistration(uuid)
  }

  async function startUpload (file) {
    reset()
    if (!file) return

    const name = file.name || '文件'
    const ext = name.includes('.') ? name.split('.').pop().toUpperCase() : ''
    fileLabel.value = {
      name,
      sizeText: formatSize(file.size),
      ext
    }

    try {
      phase.value = 'uploading'
      progress.value = 3

      const data = await uploadDocumentForChat(file)
      const uuid = data.storage_doc_uuid
      if (!uuid) {
        throw new Error('上传响应缺少 storage_doc_uuid')
      }

      const skipProcess = Number(data.processing_status) === 2
      if (skipProcess) {
        await runSkipProcessPipeline(uuid)
      } else {
        await runProcessRegisterPipeline(uuid)
      }
    } catch (e) {
      console.error('[useChatFileUpload]', e)
      phase.value = 'error'
      errorMessage.value = e.message || '上传或处理失败'
      clearWebSocket()
    }
  }

  /**
   * WebView2 宿主已上传并得到 storage_doc_uuid 时，从处理阶段继续
   */
  async function startFromHostUploadedDocument (uuid, displayName, fileSizeBytes, skipProcess = false) {
    reset()
    if (!uuid) return

    const name = displayName || '文件'
    const ext = name.includes('.') ? name.split('.').pop().toUpperCase() : ''
    fileLabel.value = {
      name,
      sizeText: typeof fileSizeBytes === 'number' && fileSizeBytes >= 0 ? formatSize(fileSizeBytes) : '—',
      ext
    }

    try {
      progress.value = 5
      if (skipProcess) {
        await runSkipProcessPipeline(uuid)
      } else {
        await runProcessRegisterPipeline(uuid)
      }
    } catch (e) {
      console.error('[useChatFileUpload] host path', e)
      phase.value = 'error'
      errorMessage.value = e.message || '处理或注册失败'
      clearWebSocket()
    }
  }

  return {
    phase,
    errorMessage,
    storageDocUuid,
    uploadId,
    progress,
    fileLabel,
    startUpload,
    startFromHostUploadedDocument,
    reset
  }
}
