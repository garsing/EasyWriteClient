<template>
  <div class="original-preview" v-loading="loading">
    <div v-if="errorMessage" class="preview-error">
      <p>{{ errorMessage }}</p>
    </div>
    <div v-else-if="previewKind === 'text'" class="preview-text">
      <pre>{{ textContent }}</pre>
    </div>
    <div v-else-if="previewKind === 'xlsx'" class="preview-xlsx">
      <div v-if="sheetNames.length > 1" class="sheet-tabs">
        <button
          v-for="name in sheetNames"
          :key="name"
          type="button"
          class="sheet-tab"
          :class="{ active: name === activeSheet }"
          @click="selectSheet(name)"
        >
          {{ name }}
        </button>
      </div>
      <div class="sheet-table-wrap" v-html="sheetHtml"></div>
    </div>
    <div v-else-if="previewKind === 'image'" class="preview-image">
      <img :src="imageUrl" :alt="documentName" />
    </div>
    <div v-else-if="previewKind === 'pdf'" class="preview-pdf" ref="pdfContainerRef"></div>
    <div v-else-if="previewKind === 'docx'" class="preview-docx" ref="docxContainerRef"></div>
    <div v-else-if="!loading" class="preview-error">
      <p>暂不支持预览</p>
    </div>
  </div>
</template>

<script setup>
import { ref, watch, nextTick, onBeforeUnmount } from 'vue'
import { renderAsync } from 'docx-preview'
import * as pdfjsLib from 'pdfjs-dist'
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import * as XLSX from 'xlsx'
import { getDocumentPreview } from '../services/knowledgeBaseApi'

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker

const props = defineProps({
  storageDocId: { type: String, required: true },
  documentName: { type: String, default: '' }
})

const loading = ref(false)
const errorMessage = ref('')
const previewKind = ref('')
const textContent = ref('')
const sheetNames = ref([])
const activeSheet = ref('')
const sheetHtml = ref('')
const workbookRef = ref(null)
const imageUrl = ref('')
const docxContainerRef = ref(null)
const pdfContainerRef = ref(null)
let pdfDoc = null
let imageBlobUrl = ''

function revokeImageUrl () {
  if (imageBlobUrl) {
    try { URL.revokeObjectURL(imageBlobUrl) } catch (_) { /* ignore */ }
    imageBlobUrl = ''
  }
  imageUrl.value = ''
}

async function renderPdf(arrayBuffer) {
  await nextTick()
  const container = pdfContainerRef.value
  if (!container) return
  container.innerHTML = ''
  pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise
  const scale = 1.15
  for (let i = 1; i <= pdfDoc.numPages; i++) {
    const page = await pdfDoc.getPage(i)
    const viewport = page.getViewport({ scale })
    const canvas = document.createElement('canvas')
    const ctx = canvas.getContext('2d')
    canvas.width = viewport.width
    canvas.height = viewport.height
    canvas.className = 'pdf-page'
    container.appendChild(canvas)
    await page.render({ canvasContext: ctx, viewport }).promise
  }
}

async function renderDocx(arrayBuffer) {
  await nextTick()
  const container = docxContainerRef.value
  if (!container) return
  container.innerHTML = ''
  await renderAsync(arrayBuffer, container, undefined, {
    className: 'docx-preview-body',
    inWrapper: true,
    ignoreWidth: false,
    breakPages: true
  })
}

function renderSheet(name) {
  const wb = workbookRef.value
  if (!wb) return
  activeSheet.value = name
  const sheet = wb.Sheets[name]
  sheetHtml.value = XLSX.utils.sheet_to_html(sheet, { editable: false })
}

function selectSheet(name) {
  renderSheet(name)
}

async function loadPreview() {
  if (!props.storageDocId) return
  loading.value = true
  errorMessage.value = ''
  previewKind.value = ''
  textContent.value = ''
  sheetHtml.value = ''
  sheetNames.value = []
  workbookRef.value = null
  revokeImageUrl()
  if (pdfContainerRef.value) pdfContainerRef.value.innerHTML = ''
  if (docxContainerRef.value) docxContainerRef.value.innerHTML = ''

  try {
    const { blob, previewKind: kind } = await getDocumentPreview(props.storageDocId)
    previewKind.value = kind

    if (kind === 'image') {
      imageBlobUrl = URL.createObjectURL(blob)
      imageUrl.value = imageBlobUrl
      return
    }

    const buffer = await blob.arrayBuffer()

    if (kind === 'text') {
      const decoder = new TextDecoder('utf-8')
      textContent.value = decoder.decode(buffer)
    } else if (kind === 'xlsx') {
      const wb = XLSX.read(buffer, { type: 'array' })
      workbookRef.value = wb
      sheetNames.value = wb.SheetNames || []
      if (sheetNames.value.length) {
        renderSheet(sheetNames.value[0])
      }
    } else if (kind === 'pdf') {
      previewKind.value = 'pdf'
      await nextTick()
      await renderPdf(buffer)
    } else if (kind === 'docx') {
      previewKind.value = 'docx'
      await nextTick()
      await renderDocx(buffer)
    } else {
      errorMessage.value = '暂不支持预览'
    }
  } catch (e) {
    console.error('[DocumentOriginalPreview]', e)
    errorMessage.value = e?.message || '预览失败'
    previewKind.value = ''
  } finally {
    loading.value = false
  }
}

watch(
  () => props.storageDocId,
  () => {
    loadPreview()
  },
  { immediate: true }
)

onBeforeUnmount(() => {
  pdfDoc = null
  revokeImageUrl()
})
</script>

<style scoped>
.original-preview {
  min-height: 200px;
  width: 100%;
}

.preview-error {
  padding: 48px 24px;
  text-align: center;
  color: #666;
  font-size: 14px;
}

.preview-text {
  background: #fff;
  border-radius: 4px;
  padding: 16px;
  overflow: auto;
  max-height: calc(100vh - 180px);
}

.preview-text pre {
  margin: 0;
  white-space: pre-wrap;
  word-break: break-word;
  font-family: Consolas, 'Courier New', monospace;
  font-size: 13px;
  line-height: 1.5;
  color: #222;
}

.preview-xlsx {
  background: #fff;
  border-radius: 4px;
  overflow: auto;
  max-height: calc(100vh - 180px);
}

.sheet-tabs {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  padding: 8px;
  border-bottom: 1px solid #e8e8e8;
  position: sticky;
  top: 0;
  background: #fff;
  z-index: 1;
}

.sheet-tab {
  border: 1px solid #d9d9d9;
  background: #fafafa;
  border-radius: 4px;
  padding: 4px 10px;
  font-size: 12px;
  cursor: pointer;
}

.sheet-tab.active {
  border-color: #1677ff;
  color: #1677ff;
  background: #e6f4ff;
}

.sheet-table-wrap {
  padding: 8px;
  overflow: auto;
}

.sheet-table-wrap :deep(table) {
  border-collapse: collapse;
  font-size: 12px;
  width: max-content;
  min-width: 100%;
}

.sheet-table-wrap :deep(td),
.sheet-table-wrap :deep(th) {
  border: 1px solid #ddd;
  padding: 4px 8px;
  white-space: pre-wrap;
}

.preview-image {
  display: flex;
  justify-content: center;
  align-items: flex-start;
  overflow: auto;
  max-height: calc(100vh - 180px);
  padding: 8px 0;
}

.preview-image img {
  max-width: 100%;
  height: auto;
  display: block;
  box-shadow: 0 1px 6px rgba(0, 0, 0, 0.12);
  background: #fff;
}

.preview-pdf {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 12px;
  padding: 8px 0;
  overflow: auto;
  max-height: calc(100vh - 180px);
}

.preview-pdf :deep(.pdf-page) {
  max-width: 100%;
  height: auto;
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.12);
  background: #fff;
}

.preview-docx {
  background: #fff;
  border-radius: 4px;
  padding: 8px;
  overflow: auto;
  max-height: calc(100vh - 180px);
}

.preview-docx :deep(.docx-wrapper) {
  background: #fff;
  padding: 0;
}
</style>
