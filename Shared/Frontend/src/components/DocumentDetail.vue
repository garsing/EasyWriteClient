<template>
  <div class="document-detail-container">
    <!-- 顶部导航栏 -->
    <div class="top-header">
      <div class="header-left">
        <el-icon class="back-icon" @click="handleBack">
          <ArrowLeft />
        </el-icon>
        <div v-if="headerThumbUrl" class="doc-icon-header doc-icon-thumb">
          <img :src="headerThumbUrl" alt="" @error="headerThumbUrl = ''" />
        </div>
        <div v-else class="doc-icon-header" :class="getDocumentIconClass(documentName)">
          {{ getDocumentIcon(documentName) }}
        </div>
        <h1 class="doc-title">{{ documentName }}</h1>
        <el-button 
          type="default" 
          size="small" 
          class="add-tag-btn-header"
          @click="handleAddTag"
        >
          添加标签
        </el-button>
        <el-radio-group
          v-if="canViewChunks"
          v-model="viewMode"
          size="small"
          class="view-mode-switch"
        >
          <el-radio-button label="chunks">分块调试</el-radio-button>
          <el-radio-button label="preview">原件预览</el-radio-button>
        </el-radio-group>
      </div>
    </div>

    <!-- 搜索和筛选区域（仅分块视图） -->
    <div v-if="viewMode === 'chunks'" class="search-filter-section">
      <el-input
        v-model="searchText"
        placeholder="搜索内容"
        class="search-input"
        :prefix-icon="Search"
        clearable
      />
      
      <el-select v-model="selectedType" placeholder="全部类型" class="filter-select">
        <el-option label="全部类型" value="all"></el-option>
      </el-select>
    </div>

    <!-- 原件预览 -->
    <div v-if="viewMode === 'preview'" class="content-section preview-section">
      <DocumentOriginalPreview
        v-if="storageDocId"
        :storage-doc-id="storageDocId"
        :document-name="documentName"
      />
    </div>

    <!-- 分块调试 -->
    <template v-else>
      <div class="content-section" ref="contentSectionRef" v-loading="loading">
        <div class="document-content">
          <template v-if="currentPageChunks && currentPageChunks.length > 0">
            <div 
              v-for="(chunk, index) in currentPageChunks" 
              :key="chunk.uuid"
              class="chunk-item"
              @mouseenter="handleChunkEnter(index)"
              @mouseleave="handleChunkLeave(index)"
            >
              <div class="chunk-header">
                <div 
                  class="chunk-number"
                  :class="{ 'is-hidden': chunk.showSummary }"
                >
                  {{ getChunkNumber(index) }} / {{ totalChunks }}
                </div>
                <div 
                  class="chunk-summary"
                  :class="{ 'is-visible': chunk.showSummary && chunk.summary }"
                >
                  <span class="summary-label">摘要</span>
                  <span class="summary-text">{{ chunk.summary }}</span>
                </div>
              </div>
              <div class="chunk-content" v-html="chunk.content"></div>
            </div>
          </template>
          <div v-else-if="!loading" class="empty-state">
            <p>暂无chunk数据 (chunk数量: {{ currentPageChunks?.length || 0 }}, 总数: {{ totalChunks }})</p>
          </div>
        </div>
      </div>

      <div class="footer-section">
        <div class="footer-left">
          <span class="total-count">共计{{ totalChunks }}chunk</span>
        </div>
        <div class="footer-right">
          <el-pagination
            v-model:current-page="currentPage"
            v-model:page-size="pageSize"
            :page-sizes="[20, 50, 100]"
            :total="totalChunks"
            layout="prev, pager, next, sizes"
            small
            :disabled="loading"
            @size-change="handlePageSizeChange"
            @current-change="handlePageChange"
          />
        </div>
      </div>
    </template>
  </div>
</template>

<script setup>
import { ref, onMounted, onUnmounted, computed, watch, nextTick } from 'vue'
import { ElMessage } from 'element-plus'
import {
  ArrowLeft,
  Search
} from '@element-plus/icons-vue'
import { getDocumentChunks, getImageUrl, getMyKbCapabilities } from '../services/knowledgeBaseApi'
import { isImageFileName, loadKbImageThumbUrl } from '../utils/kbImageThumb'
import DocumentOriginalPreview from './DocumentOriginalPreview.vue'

// Props
const props = defineProps({
  documentName: {
    type: String,
    default: '司法数据解析项目报告 - 20260101.docx'
  },
  storageDocId: {
    type: String,
    default: ''
  },
  knowledgeBaseUuid: {
    type: String,
    default: ''
  }
})

// Emits
const emit = defineEmits(['back'])

// 搜索和筛选
const searchText = ref('')
const selectedType = ref('all')

// 视图：preview | chunks
const canViewChunks = ref(false)
const viewMode = ref('preview')
const capabilitiesLoaded = ref(false)

// chunk信息
const totalChunks = ref(0)
const currentPage = ref(1)
const pageSize = ref(50) // 默认每页50个chunk

// 加载状态
const loading = ref(false)

// 内容区域引用
const contentSectionRef = ref(null)

// 文档chunk数据（只存储当前页的chunk）
const chunks = ref([])

// 计算总页数
const totalPages = computed(() => {
  return Math.ceil(totalChunks.value / pageSize.value)
})

// 当前页的chunk（直接使用 chunks，因为 chunks 只包含当前页的数据）
const currentPageChunks = computed(() => {
  const result = chunks.value || []
  console.log('[currentPageChunks] 计算属性被调用，chunks.value.length:', chunks.value?.length || 0, '返回长度:', result.length)
  return result
})

/**
 * 获取文档图标文字（W、X、PDF、PPT 等）
 */
function getDocumentIcon(fileName) {
  if (!fileName) return 'W'
  if (isImageFileName(fileName)) return 'IMG'
  const lowerName = fileName.toLowerCase()
  if (lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls')) {
    return 'X'
  }
  if (lowerName.endsWith('.pdf')) {
    return 'PDF'
  }
  if (lowerName.endsWith('.pptx') || lowerName.endsWith('.ppt')) {
    return 'PPT'
  }
  return 'W'
}

/**
 * 获取文档图标样式类
 */
function getDocumentIconClass(fileName) {
  if (!fileName) return 'doc-icon-word'
  if (isImageFileName(fileName)) return 'doc-icon-image'
  const lowerName = fileName.toLowerCase()
  if (lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls')) {
    return 'doc-icon-excel'
  }
  if (lowerName.endsWith('.pdf')) {
    return 'doc-icon-pdf'
  }
  if (lowerName.endsWith('.pptx') || lowerName.endsWith('.ppt')) {
    return 'doc-icon-ppt'
  }
  return 'doc-icon-word'
}

const headerThumbUrl = ref('')
let headerThumbBlob = ''

function revokeHeaderThumb () {
  if (headerThumbBlob) {
    try { URL.revokeObjectURL(headerThumbBlob) } catch (_) { /* ignore */ }
    headerThumbBlob = ''
  }
  headerThumbUrl.value = ''
}

async function loadHeaderThumb () {
  revokeHeaderThumb()
  if (!isImageFileName(props.documentName) || !props.storageDocId || !props.knowledgeBaseUuid) {
    return
  }
  try {
    const url = await loadKbImageThumbUrl(props.storageDocId, props.knowledgeBaseUuid)
    headerThumbBlob = url
    headerThumbUrl.value = url
  } catch (e) {
    console.warn('[DocumentDetail] 缩略图加载失败', e?.message || e)
  }
}

watch(
  () => [props.storageDocId, props.knowledgeBaseUuid, props.documentName],
  () => { loadHeaderThumb() },
  { immediate: true }
)
onUnmounted(revokeHeaderThumb)

// 获取chunk在整个文档中的编号
const getChunkNumber = (index) => {
  const start = (currentPage.value - 1) * pageSize.value
  return start + index + 1
}

// 鼠标进入chunk
const handleChunkEnter = (index) => {
  if (chunks.value[index]) {
    chunks.value[index].showSummary = true
  }
}

// 鼠标离开chunk
const handleChunkLeave = (index) => {
  if (chunks.value[index]) {
    chunks.value[index].showSummary = false
  }
}

/**
 * 加载chunk数据
 */
const loadChunks = async () => {
  if (!props.storageDocId) {
    console.warn('[loadChunks] 文档UUID为空，跳过加载')
    return
  }
  
  loading.value = true
  
  try {
    const startIndex = (currentPage.value - 1) * pageSize.value
    const endIndex = startIndex + pageSize.value
    
    console.log('[loadChunks] 开始加载chunk:', {
      storageDocUuid: props.storageDocId,
      startIndex,
      endIndex,
      page: currentPage.value,
      pageSize: pageSize.value
    })
    
    const result = await getDocumentChunks(props.storageDocId, startIndex, endIndex)
    
    console.log('[loadChunks] 加载成功:', result)
    
    // apiRequest 已经提取了 responseData.data，所以 result 就是数据对象本身
    if (result && result.chunks) {
      // 更新总数
      totalChunks.value = result.total || 0
      
      // 转换chunk数据格式
      const newChunks = await Promise.all((result.chunks || []).map(async (chunk) => {
        // 处理换行符：将 \n 转换为 <br>，以便在 HTML 中正确显示
        let content = chunk.chunk_content || ''
        
        const knowledgeBaseUuid =
          chunk.knowledge_base_uuid || props.knowledgeBaseUuid

        // 处理图片标签：将 <image src="xxx" /> 转换为 <img> 标签
        if (content.includes('<image') && knowledgeBaseUuid) {
          const imageRegex = /<image\s+src=["']([^"']+)["']\s*\/?>/gi
          let match
          const imageMatches = []
          while ((match = imageRegex.exec(content)) !== null) {
            imageMatches.push({
              fullMatch: match[0],
              imagePath: match[1]
            })
          }
          for (const imageMatch of imageMatches) {
            if (imageMatch.imagePath.startsWith('lightweight/')) {
              content = content.replace(
                imageMatch.fullMatch,
                '<span style="color: #999;">[轻量模式占位图，未提取原图]</span>'
              )
              continue
            }
            try {
              const imageUrl = await getImageUrl(imageMatch.imagePath, knowledgeBaseUuid)
              content = content.replace(
                imageMatch.fullMatch,
                `<img src="${imageUrl}" alt="图片" style="max-width: 40%; height: auto; display: block; margin: 12px 0;" />`
              )
            } catch (error) {
              console.error('[loadChunks] 处理图片失败:', error, '图片路径:', imageMatch.imagePath)
              content = content.replace(
                imageMatch.fullMatch,
                `<span style="color: #999;">[图片加载失败: ${imageMatch.imagePath}]</span>`
              )
            }
          }
        }
        
        // 转换自定义表格格式 <table><row><cell> 为标准 HTML 表格格式
        if (content.includes('<table') && content.includes('<row>')) {
          // 将 <table> 标签（可能带属性）转换为标准 <table> 标签
          // 匹配 <table 属性> 格式，保留 id 和 description 属性（如果需要）
          content = content.replace(/<table\s+([^>]*?)>/gi, (match, attributes) => {
            // 提取 id 和 description 属性（可选，用于调试）
            const idMatch = attributes.match(/id\s*=\s*["']([^"']+)["']/i)
            const descMatch = attributes.match(/description\s*=\s*["']([^"']+)["']/i)
            
            let tableAttrs = ''
            if (idMatch) {
              tableAttrs += ` data-table-id="${idMatch[1]}"`
            }
            if (descMatch) {
              tableAttrs += ` data-table-desc="${descMatch[1]}"`
            }
            
            return `<table${tableAttrs}>`
          })
          // 处理没有属性的 <table> 标签
          content = content.replace(/<table>/gi, '<table>')
          
          // 将 <row> 转换为 <tr>
          content = content.replace(/<row>/gi, '<tr>')
          content = content.replace(/<\/row>/gi, '</tr>')
          
          // 将 <cell> 转换为 <td>，保留 rowspan 和 colspan 属性
          // 匹配 <cell> 或 <cell 属性> 格式，包括 rowspan 和 colspan
          content = content.replace(/<cell\s+([^>]*?)>/gi, (match, attributes) => {
            // 提取 rowspan 和 colspan 属性
            const rowspanMatch = attributes.match(/rowspan\s*=\s*["']?(\d+)["']?/i)
            const colspanMatch = attributes.match(/colspan\s*=\s*["']?(\d+)["']?/i)
            
            let tdAttrs = ''
            if (rowspanMatch) {
              tdAttrs += ` rowspan="${rowspanMatch[1]}"`
            }
            if (colspanMatch) {
              tdAttrs += ` colspan="${colspanMatch[1]}"`
            }
            
            return `<td${tdAttrs}>`
          })
          // 处理没有属性的 <cell> 标签
          content = content.replace(/<cell>/gi, '<td>')
          content = content.replace(/<\/cell>/gi, '</td>')
          
          // 将表格内容中的 \r 转换为 <br>（在 <td> 标签内）
          // 使用正则表达式匹配 <td> 标签内的内容，将 \r 转换为 <br>
          content = content.replace(/(<td[^>]*>)([\s\S]*?)(<\/td>)/gi, (match, openTag, cellContent, closeTag) => {
            // 将 cellContent 中的 \r 转换为 <br>
            const processedContent = cellContent.replace(/\r/g, '<br>')
            return openTag + processedContent + closeTag
          })
        }
        
        // 如果内容是纯文本（不包含 HTML 标签），则将换行符转换为 <br>
        if (content && !content.includes('<') && !content.includes('>')) {
          content = content.replace(/\n/g, '<br>')
        }
        // 如果内容包含 HTML 标签但仍有换行符，也需要处理
        else if (content) {
          // 对于 HTML 内容，将不在标签内的换行符转换为 <br>
          // 使用正则表达式匹配不在 HTML 标签内的换行符
          content = content.replace(/(?<!>)\n(?!<)/g, '<br>')
        }
        
        return {
          uuid: chunk.uuid,
          content: content,
          summary: chunk.summary || null,
          chunkIndex: chunk.chunk_index,
          showSummary: false
        }
      }))
      
      // 按照 chunk_index 排序（虽然后端应该已经排序，但为了安全起见）
      newChunks.sort((a, b) => a.chunkIndex - b.chunkIndex)
      
      // 直接替换当前页的chunk数据
      console.log('[loadChunks] 准备更新chunk数据，newChunks.length:', newChunks.length)
      
      // 使用 nextTick 确保响应式更新
      await nextTick()
      chunks.value = newChunks
      await nextTick()
      
      console.log('[loadChunks] chunk数据已更新:', {
        total: totalChunks.value,
        chunksCount: chunks.value.length,
        chunksValue: chunks.value,
        currentPageChunksCount: currentPageChunks.value.length,
        currentPageChunksValue: currentPageChunks.value
      })
    } else {
      console.warn('[loadChunks] 响应数据格式不正确:', result)
      // 如果没有chunk数据，清空列表
      chunks.value = []
      totalChunks.value = 0
    }
  } catch (error) {
    console.error('[loadChunks] 加载失败:', error)
    ElMessage.error(error.message || '加载chunk列表失败')
  } finally {
    loading.value = false
  }
}

// 初始化标志，避免 watch 在初始化时重复调用
const isInitialized = ref(false)

async function loadCapabilitiesAndInit() {
  try {
    const caps = await getMyKbCapabilities()
    canViewChunks.value = !!(caps && caps.can_view_document_chunks)
  } catch (e) {
    console.warn('[DocumentDetail] capabilities 失败，按普通用户处理:', e)
    canViewChunks.value = false
  }
  viewMode.value = (canViewChunks.value && !isImageFileName(props.documentName))
    ? 'chunks'
    : 'preview'
  capabilitiesLoaded.value = true
  if (props.storageDocId && viewMode.value === 'chunks') {
    isInitialized.value = true
    await loadChunks()
  }
}

// 监听 documentId 变化，重新加载数据
watch(() => props.storageDocId, (newId) => {
  if (!newId || !capabilitiesLoaded.value) return
  chunks.value = []
  totalChunks.value = 0
  currentPage.value = 1
  if (viewMode.value === 'chunks') {
    isInitialized.value = true
    loadChunks()
  }
})

// 切换到分块时加载
watch(viewMode, (mode) => {
  if (mode === 'chunks' && props.storageDocId && capabilitiesLoaded.value) {
    isInitialized.value = true
    loadChunks()
  }
})

// 监听页码和每页大小变化，重新加载数据（只在初始化后触发）
watch([currentPage, pageSize], () => {
  if (props.storageDocId && isInitialized.value && viewMode.value === 'chunks') {
    loadChunks()
  }
})

onMounted(() => {
  loadCapabilitiesAndInit()
})

// 返回文档列表
const handleBack = () => {
  emit('back')
}

// 添加标签
const handleAddTag = () => {
  console.log('添加标签')
  // TODO: 实现添加标签逻辑
}

// 分页大小改变
const handlePageSizeChange = (newSize) => {
  pageSize.value = newSize
  currentPage.value = 1 // 重置到第一页
  // 滚动到顶部
  if (contentSectionRef.value) {
    contentSectionRef.value.scrollTop = 0
  }
  // loadChunks 会在 watch 中自动调用
}

// 页码改变
const handlePageChange = (newPage) => {
  currentPage.value = newPage
  // 滚动到顶部
  if (contentSectionRef.value) {
    contentSectionRef.value.scrollTop = 0
  }
  // loadChunks 会在 watch 中自动调用
}
</script>

<style scoped>
.document-detail-container {
  display: flex;
  flex-direction: column;
  height: 100vh;
  width: 100%;
  background-color: #f5f5f5;
  margin: 0;
  padding: 0;
  box-sizing: border-box;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
}

/* 顶部导航栏 */
.top-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 16px 24px;
  background-color: #fff;
  border-bottom: 1px solid #e0e0e0;
}

.header-left {
  display: flex;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
}

.view-mode-switch {
  margin-left: 8px;
}

.preview-section {
  flex: 1;
  overflow: auto;
  padding: 16px 24px;
}

.back-icon {
  font-size: 20px;
  color: #666;
  cursor: pointer;
  transition: color 0.2s;
}

.back-icon:hover {
  color: #333;
}

.doc-icon-header {
  width: 32px;
  height: 32px;
  min-width: 32px;
  min-height: 32px;
  max-width: 32px;
  max-height: 32px;
  border-radius: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  color: #fff;
  font-weight: 600;
  font-size: 16px;
}

.doc-icon-word {
  background-color: #2B579A;
}

.doc-icon-excel {
  background-color: #217346;
}

.doc-icon-pdf {
  background-color: #D93025;
  width: auto;
  min-width: 32px;
  max-width: none;
  padding: 0 4px;
  font-size: 11px;
}

.doc-icon-ppt {
  background-color: #D24726;
  width: auto;
  min-width: 32px;
  max-width: none;
  padding: 0 4px;
  font-size: 11px;
}

.doc-icon-image {
  background-color: #7B61FF;
  font-size: 9px;
}

.doc-icon-thumb {
  background: #f0f0f0;
  overflow: hidden;
  padding: 0;
}

.doc-icon-thumb img {
  width: 100%;
  height: 100%;
  object-fit: cover;
  display: block;
}

.doc-title {
  margin: 0;
  font-size: 18px;
  font-weight: 500;
  color: #333;
  line-height: 1.5;
}

.add-tag-btn-header {
  margin-left: 12px;
}

/* 搜索和筛选区域 */
.search-filter-section {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 16px 24px;
  background-color: #fff;
  border-bottom: 1px solid #e0e0e0;
}

.search-input {
  width: 300px;
}

.filter-select {
  width: 140px;
}

/* 底部信息 */
.footer-section {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 16px 24px;
  background-color: #fff;
  border-top: 1px solid #e0e0e0;
}

.footer-left {
  display: flex;
  align-items: center;
  gap: 16px;
}

.total-count {
  color: #666;
  font-size: 14px;
}

/* 文档内容区域 */
.content-section {
  flex: 1;
  padding: 24px;
  overflow: auto;
  background-color: #f5f5f5;
}

.document-content {
  max-width: 100%;
  line-height: 1.5;
  color: #333;
}

.document-content :deep(br) {
  display: block;
  content: "";
  margin-top: 0.5em;
}

.chunk-content {
  font-size: 14px;
  line-height: 1.8;
  color: #333;
}

.chunk-content :deep(*) {
  font-size: 14px;
  line-height: 1.5;
}

.chunk-content :deep(table) {
  border: 1px solid #000 !important;
  border-collapse: collapse !important;
}

.chunk-content :deep(table td),
.chunk-content :deep(table th) {
  border: 1px solid #000 !important;
  padding: 12px !important;
}

.chunk-content :deep(table tr) {
  border-bottom: 1px solid #000 !important;
}

.chunk-item {
  padding: 8px 24px 16px 24px;
  margin-bottom: 16px;
  background-color: #fff;
  border-radius: 4px;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
  transition: all 0.2s;
  position: relative;
  font-size: 14px;
  line-height: 1.5;
  color: #333;
}

.chunk-item:last-child {
  margin-bottom: 0;
}

.chunk-header {
  position: relative;
  margin-bottom: 8px;
  min-height: 20px;
}

.chunk-number {
  display: block;
  font-size: 12px;
  color: #d9d9d9;
  font-weight: 500;
  transition: opacity 0.2s, visibility 0.2s;
  position: absolute;
  top: 0;
  left: 0;
}

.chunk-number.is-hidden {
  opacity: 0;
  visibility: hidden;
}

.chunk-summary {
  display: none;
  font-size: 14px !important;
  color: #333;
  font-weight: 500;
  transition: opacity 0.2s, visibility 0.2s;
  position: absolute;
  top: 0;
  left: 0;
  align-items: center;
  gap: 8px;
  line-height: 1.5;
}

.chunk-summary.is-visible {
  display: flex;
  opacity: 1;
  visibility: visible;
}

.summary-label {
  color: #333;
  font-weight: 500;
  white-space: nowrap;
  font-size: 14px !important;
}

.summary-text {
  color: #666;
  font-weight: normal;
  line-height: 1.5;
  font-size: 14px !important;
}

.chunk-loading {
  padding: 16px;
  text-align: center;
  color: #999;
  font-size: 14px;
}

.empty-state {
  padding: 40px;
  text-align: center;
  color: #999;
  font-size: 14px;
}

.doc-content-wrapper {
  padding: 0;
}

.doc-content-wrapper p {
  margin: 8px 0;
  font-size: 14px;
  line-height: 1.6;
  color: #333;
}

.section-title {
  font-size: 16px;
  font-weight: 600;
  color: #333;
  margin: 16px 0 8px 0;
}

.performance-table {
  width: 100%;
  border-collapse: collapse;
  margin: 12px 0;
  font-size: 14px;
}

.performance-table th,
.performance-table td {
  border: 1px solid #e0e0e0;
  padding: 12px;
  text-align: center;
}

.performance-table th {
  background-color: #f5f5f5;
  font-weight: 500;
  color: #333;
}

.performance-table td {
  background-color: #fff;
}

.performance-table tbody tr:hover {
  background-color: #f9f9f9;
}

/* chunk内容中的表格样式 */
.chunk-content table {
  width: 100% !important;
  border-collapse: collapse !important;
  margin: 12px 0 !important;
  font-size: 14px !important;
  border: 1px solid #000 !important;
}

.chunk-content table tr {
  border-bottom: 1px solid #000 !important;
}

.chunk-content table tr:last-child {
  border-bottom: 1px solid #000 !important;
}

.chunk-content table td {
  border: 1px solid #000 !important;
  padding: 12px !important;
  text-align: left !important;
  vertical-align: top !important;
  line-height: 1.5 !important;
  font-size: 14px !important;
}

.chunk-content table th {
  border: 1px solid #000 !important;
  padding: 12px !important;
  text-align: left !important;
  vertical-align: top !important;
  line-height: 1.5 !important;
  font-size: 14px !important;
  background-color: #f5f5f5 !important;
  font-weight: 500 !important;
  color: #333 !important;
}

.chunk-content table tr:hover {
  background-color: #f9f9f9 !important;
}

.chunk-content table tr:first-child td,
.chunk-content table tr:first-child th {
  background-color: #f5f5f5 !important;
  font-weight: 500 !important;
  color: #333 !important;
  border: 1px solid #000 !important;
}
</style>

