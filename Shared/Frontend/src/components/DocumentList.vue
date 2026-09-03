<template>
  <div class="document-list-container">
    <!-- 文档详情页 -->
    <DocumentDetail
      v-if="showDocumentDetail"
      :document-name="selectedDocument?.name || ''"
      :storage-doc-id="selectedDocument?.storageDocUuid || ''"
      :knowledge-base-uuid="knowledgeBaseUuid"
      @back="handleBackFromDetail"
    />
    
    <!-- 文档列表页 -->
    <template v-else>
    <!-- 顶部导航栏 -->
    <div class="top-header">
      <div class="header-left">
        <el-icon class="back-icon" @click="handleBack">
          <ArrowLeft />
        </el-icon>
        <div class="kb-icon-small">N</div>
        <h1 class="kb-title">{{ knowledgeBaseName }}</h1>
        <div class="doc-count-info">
          <el-icon><Document /></el-icon>
          <span>{{ documentCount }}</span>
        </div>
        <el-icon 
          class="star-icon-header" 
          :class="{ 'is-favorite': isFavorite }" 
          @click="toggleFavorite"
        >
          <StarFilled v-if="isFavorite" />
          <Star v-else />
        </el-icon>
      </div>
    </div>

    <!-- 搜索和筛选区域 -->
    <div class="search-filter-section">
      <el-input
        v-model="searchText"
        placeholder="搜索文档名称"
        class="search-input"
        :prefix-icon="Search"
        clearable
      />
      
      <el-select v-model="selectedTag" placeholder="未选择标签" class="filter-select">
        <el-option label="未选择标签" value=""></el-option>
      </el-select>
      
      <el-select v-model="selectedStatus" placeholder="全部状态" class="filter-select">
        <el-option label="全部状态" value="all"></el-option>
        <el-option label="可用" value="available"></el-option>
        <el-option label="不可用" value="unavailable"></el-option>
      </el-select>
      
      <el-button type="primary" :icon="Plus" class="add-doc-btn" @click="showUploadDialog = true">
        添加文档
      </el-button>
    </div>

    <!-- 文档列表表格 -->
    <div class="table-section">
      <el-table
        ref="tableRef"
        :data="documentList"
        style="width: 100%"
        class="doc-table"
        v-loading="loading"
        row-key="storageDocUuid"
        :header-cell-style="{ background: '#f5f5f5', color: '#333', fontWeight: 'normal' }"
        @row-mouse-enter="handleRowEnter"
        @row-mouse-leave="handleRowLeave"
        @row-click="handleRowClick"
        @selection-change="handleSelectionChange"
      >
        <el-table-column type="selection" width="55"></el-table-column>
        
        <el-table-column prop="name" label="文档" min-width="250">
          <template #default="{ row }">
            <div class="doc-name-cell" @click.stop="handleOpenDocumentDetail(row)">
              <div
                v-if="isImageFileName(row.name) && thumbUrls[row.storageDocUuid]"
                class="doc-icon doc-icon-thumb"
              >
                <img
                  :src="thumbUrls[row.storageDocUuid]"
                  alt=""
                  @error="onThumbError(row)"
                />
              </div>
              <div v-else class="doc-icon" :class="getDocumentIconClass(row.name)">
                {{ getDocumentIcon(row.name) }}
              </div>
              <span class="doc-name">{{ row.name }}</span>
            </div>
          </template>
        </el-table-column>
        
        <el-table-column prop="tags" label="标签" min-width="120">
          <template #default="{ row }">
            <div class="tags-cell">
              <span v-if="row.tags && row.tags.length > 0" class="tags-list">
                <el-tag v-for="tag in row.tags" :key="tag" size="small" class="tag-item">
                  {{ tag }}
                </el-tag>
              </span>
              <div 
                v-if="row.showAddTag || (!row.tags || row.tags.length === 0)"
                class="add-tag-container"
                :class="{ 'is-hover': row.showAddTag }"
                @click.stop="handleAddTag(row)"
              >
                <span class="add-tag-text">添加标签</span>
              </div>
            </div>
          </template>
        </el-table-column>
        
        <el-table-column prop="size" label="大小" min-width="100">
          <template #default="{ row }">
            {{ row.size }}
          </template>
        </el-table-column>
        
        <el-table-column prop="chunkCount" label="chunk数" min-width="100">
          <template #default="{ row }">
            {{ row.chunkCount }}
          </template>
        </el-table-column>
        
        <el-table-column prop="charCount" label="字符数" min-width="100">
          <template #default="{ row }">
            {{ row.charCount }}
          </template>
        </el-table-column>
        
        <el-table-column prop="updater" label="更新人" min-width="120">
          <template #default="{ row }">
            {{ row.updater }}
          </template>
        </el-table-column>
        
        <el-table-column prop="updateTime" label="更新时间" min-width="150" sortable>
          <template #default="{ row }">
            {{ row.updateTime }}
          </template>
        </el-table-column>
        
        <el-table-column prop="status" label="状态" min-width="150" fixed="right">
          <template #default="{ row }">
            <div v-if="row.status === '处理中' || (row.processingProgress !== undefined && row.processingProgress !== null && row.processingProgress < 1)" class="status-processing">
              <el-progress 
                :percentage="getProgressPercentage(row.processingProgress)" 
                :stroke-width="6"
                :show-text="true"
                :format="(percentage) => `${percentage}%`"
                color="#409eff"
              />
            </div>
            <el-tag v-else :type="row.status === '可用' ? 'success' : 'info'" size="small">
              {{ row.status }}
            </el-tag>
          </template>
        </el-table-column>
        
        <el-table-column prop="enabled" label="启用" min-width="100" fixed="right">
          <template #default="{ row }">
            <el-switch
              v-model="row.enabled"
              active-color="#722ed1"
              @change="handleToggleEnable(row)"
            />
          </template>
        </el-table-column>
      </el-table>
    </div>

    <!-- 底部信息 -->
    <div class="footer-section">
      <div class="footer-left">
        <span class="total-count">共计{{ total }}文档</span>
        <div v-if="selectedDocuments.length > 0" class="selection-info">
          <span class="selected-count">已选 {{ selectedDocuments.length }}/{{ documentList.length }} 数据</span>
          <el-button 
            :icon="Delete" 
            type="danger"
            size="small"
            class="delete-btn"
            @click="handleDeleteClick"
          >
            删除
          </el-button>
        </div>
      </div>
      <div class="footer-right">
        <el-pagination
          v-model:current-page="currentPage"
          v-model:page-size="pageSize"
          :page-sizes="[20, 50, 100]"
          :total="total"
          layout="prev, pager, next, sizes"
          small
          @size-change="loadDocumentList"
          @current-change="loadDocumentList"
        />
      </div>
    </div>

    <!-- 上传文档对话框 -->
    <el-dialog
      v-model="showUploadDialog"
      title="上传文档"
      width="600px"
      :close-on-click-modal="false"
      @close="handleCloseUploadDialog"
    >
      <el-upload
        ref="uploadRef"
        class="upload-container"
        drag
        :auto-upload="false"
        :on-change="handleFileChange"
        :on-remove="handleRemoveFile"
        multiple
        :file-list="[]"
        :show-file-list="false"
      >
        <el-icon class="el-icon--upload"><upload-filled /></el-icon>
        <div class="el-upload__text">
          将文件拖到此处，或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            支持文档与图片拖拽上传，文件选择后会自动上传
          </div>
        </template>
      </el-upload>

      <!-- 文件列表 -->
      <div v-if="uploadedFiles.length > 0" class="upload-file-list">
        <div
          v-for="file in uploadedFiles"
          :key="file.uid"
          class="upload-file-item"
        >
          <div class="file-info">
            <span class="file-name">{{ file.name }}</span>
            <span class="file-size">{{ formatFileSize(file.size / 1024) }}</span>
          </div>
          <div class="file-status">
            <el-icon v-if="file.status === 'success'" class="success-icon">
              <circle-check />
            </el-icon>
            <el-icon v-else-if="file.status === 'error'" class="error-icon">
              <circle-close />
            </el-icon>
            <span v-else class="uploading-text">上传中...</span>
          </div>
        </div>
      </div>

      <template #footer>
        <div class="dialog-footer">
          <el-button @click="handleCloseUploadDialog">取消</el-button>
          <el-button
            type="primary"
            :disabled="uploadedFiles.filter(f => f.status === 'success').length === 0 || isProcessing"
            :loading="isProcessing"
            @click="handleCompleteUpload"
          >
            完成
          </el-button>
        </div>
      </template>
    </el-dialog>
    </template>
  </div>
</template>

<script setup>
import { ref, onMounted, onUnmounted, watch } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import {
  ArrowLeft,
  Document,
  Star,
  StarFilled,
  Plus,
  Search,
  Delete
} from '@element-plus/icons-vue'
import DocumentDetail from './DocumentDetail.vue'
import { getDocumentList, uploadDocument, processDocument, createDocumentProcessWebSocket, deleteDocuments, setDocumentEnableStatus } from '../services/knowledgeBaseApi'
import { isImageFileName, loadKbImageThumbUrl } from '../utils/kbImageThumb'

// Props
const props = defineProps({
  knowledgeBaseName: {
    type: String,
    default: 'Agent未命名应用1747文档知识库'
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
const selectedTag = ref('')
const selectedStatus = ref('all')

// 关注状态
const isFavorite = ref(false)

// 文档数量
const documentCount = ref(0)

// 页面状态
const showDocumentDetail = ref(false)
const selectedDocument = ref(null)
const loading = ref(false)

// 文档列表数据
const documentList = ref([])

// 分页
const currentPage = ref(1)
const pageSize = ref(20)
const total = ref(0)

// 上传对话框
const showUploadDialog = ref(false)
const uploadedFiles = ref([])
const isProcessing = ref(false)
const uploadRef = ref(null)

// 表格引用和选中状态
const tableRef = ref(null)
const selectedDocuments = ref([])

const thumbUrls = ref({})
let thumbBlobUrls = []

function revokeThumbs () {
  thumbBlobUrls.forEach((u) => {
    try { URL.revokeObjectURL(u) } catch (_) { /* ignore */ }
  })
  thumbBlobUrls = []
}

async function loadThumbs (docs) {
  revokeThumbs()
  thumbUrls.value = {}
  const kb = props.knowledgeBaseUuid
  if (!kb) return
  const next = {}
  await Promise.all((docs || []).map(async (row) => {
    if (!isImageFileName(row.name) || !row.storageDocUuid) return
    try {
      const url = await loadKbImageThumbUrl(row.storageDocUuid, kb)
      thumbBlobUrls.push(url)
      next[row.storageDocUuid] = url
    } catch (e) {
      console.warn('[DocumentList] 缩略图加载失败', row.name, e?.message || e)
    }
  }))
  thumbUrls.value = next
}

function onThumbError (row) {
  if (!row?.storageDocUuid) return
  const copy = { ...thumbUrls.value }
  delete copy[row.storageDocUuid]
  thumbUrls.value = copy
}

watch(documentList, (docs) => {
  loadThumbs(docs)
})

onUnmounted(revokeThumbs)

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

/**
 * 格式化文件大小
 */
function formatFileSize(sizeKb) {
  if (sizeKb < 1024) {
    return `${sizeKb}KB`
  } else if (sizeKb < 1024 * 1024) {
    return `${(sizeKb / 1024).toFixed(2)}MB`
  } else {
    return `${(sizeKb / (1024 * 1024)).toFixed(2)}GB`
  }
}

/**
 * 格式化字符数
 */
function formatCharCount(count) {
  if (count < 1000) {
    return count.toString()
  } else if (count < 1000000) {
    return `${(count / 1000).toFixed(1)}k`
  } else {
    return `${(count / 1000000).toFixed(1)}M`
  }
}

/**
 * 格式化时间
 */
function formatTime(timeString) {
  if (!timeString) return '未知'
  try {
    const date = new Date(timeString)
    const year = date.getFullYear()
    const month = String(date.getMonth() + 1).padStart(2, '0')
    const day = String(date.getDate()).padStart(2, '0')
    const hours = String(date.getHours()).padStart(2, '0')
    const minutes = String(date.getMinutes()).padStart(2, '0')
    return `${year}-${month}-${day} ${hours}:${minutes}`
  } catch (error) {
    return timeString
  }
}

/**
 * 解析标签字符串（逗号分隔）
 */
function parseTags(tagsString) {
  if (!tagsString || !tagsString.trim()) {
    return []
  }
  return tagsString.split(',').map(tag => tag.trim()).filter(tag => tag.length > 0)
}

/**
 * 将0-1的进度值转换为0-100的百分比
 */
function getProgressPercentage(progress) {
  if (progress === undefined || progress === null) {
    return 0
  }
  // 如果已经是0-100的值，直接返回
  if (progress > 1) {
    return Math.min(100, Math.max(0, progress))
  }
  // 如果是0-1的值，转换为0-100
  return Math.min(100, Math.max(0, progress * 100))
}

/**
 * 根据 processing_status 和 processing_progress 获取状态文本
 * 状态列只根据 processing_status 来显示，与 is_enabled 无关
 */
function getStatusFromProcessingStatus(processingStatus, processingProgress) {
  // 如果 processing_status 是 processing 或 pending，显示"处理中"
  if (processingStatus === 'processing' || processingStatus === 'pending') {
    return '处理中'
  }
  
  // 如果 processing_status 是 completed 或 success，显示"可用"
  if (processingStatus === 'completed' || processingStatus === 'success') {
    return '可用'
  }
  
  // 如果 processing_status 是 failed 或 error，显示"处理失败"
  if (processingStatus === 'failed' || processingStatus === 'error') {
    return '处理失败'
  }
  
  // 如果没有 processing_status，但 processing_progress 存在且小于1，显示"处理中"
  if (processingProgress !== undefined && processingProgress !== null && processingProgress < 1) {
    return '处理中'
  }
  
  // 默认显示"可用"（假设处理完成）
  return '可用'
}

/**
 * 加载文档列表
 */
async function loadDocumentList() {
  const listStartTime = performance.now()
  console.log(`[loadDocumentList] ===== 开始加载文档列表 =====`, new Date().toISOString())
  
  if (!props.knowledgeBaseUuid) {
    console.warn('[loadDocumentList] knowledgeBaseUuid 为空，跳过加载')
    return
  }
  
  console.log('[loadDocumentList] knowledgeBaseUuid:', props.knowledgeBaseUuid)
  loading.value = true
  
  const loadingSetTime = performance.now()
  console.log(`[loadDocumentList] ✓ 加载状态已设置`, `耗时: ${(loadingSetTime - listStartTime).toFixed(2)}ms`)
  
  try {
    const startIndex = (currentPage.value - 1) * pageSize.value
    const endIndex = startIndex + pageSize.value
    
    console.log(`[loadDocumentList] 调用 API, startIndex: ${startIndex}, endIndex: ${endIndex}`)
    const apiCallTime = performance.now()
    
    const data = await getDocumentList(props.knowledgeBaseUuid, startIndex, endIndex)
    
    const apiReturnTime = performance.now()
    console.log(`[loadDocumentList] ✓ API 返回数据`, `API耗时: ${(apiReturnTime - apiCallTime).toFixed(2)}ms, 总耗时: ${(apiReturnTime - listStartTime).toFixed(2)}ms`)
    console.log('[loadDocumentList] 数据内容:', data)
    
    // 更新总数和文档数量
    total.value = data.total || 0
    documentCount.value = data.total || 0
    
    // 转换数据格式
    const transformStartTime = performance.now()
    documentList.value = (data.documents || []).map(doc => ({
      storageDocUuid: doc.storage_doc_uuid,
      name: doc.document_name || '',
      tags: parseTags(doc.tags),
      size: formatFileSize(doc.document_size_kb || 0),
      chunkCount: doc.chunk_count || 0,
      charCount: formatCharCount(doc.character_count || 0),
      updater: doc.owner_username || '',
      updateTime: formatTime(doc.updated_at || doc.uploaded_at),
      // 状态列只根据 processing_status 来设置，与 is_enabled 无关
      status: getStatusFromProcessingStatus(doc.processing_status, doc.processing_progress),
      // 启用列只根据 is_enabled 来设置
      enabled: doc.is_enabled || false,
      showAddTag: false,
      processingProgress: doc.processing_progress !== undefined && doc.processing_progress !== null ? doc.processing_progress : 0
    }))
    
    const transformEndTime = performance.now()
    console.log(`[loadDocumentList] ✓ 数据转换完成`, `转换耗时: ${(transformEndTime - transformStartTime).toFixed(2)}ms`)
    
    const listEndTime = performance.now()
    console.log(`[loadDocumentList] ===== 文档列表加载完成 =====`, `总耗时: ${(listEndTime - listStartTime).toFixed(2)}ms`)
    
  } catch (error) {
    const errorTime = performance.now()
    console.error(`[loadDocumentList] ✗ 加载文档列表失败:`, error, `耗时: ${(errorTime - listStartTime).toFixed(2)}ms`)
    ElMessage.error(error.message || '加载文档列表失败')
    documentList.value = []
    total.value = 0
    documentCount.value = 0
  } finally {
    loading.value = false
    const finalTime = performance.now()
    console.log(`[loadDocumentList] ✓ 加载状态已清除`, `总耗时: ${(finalTime - listStartTime).toFixed(2)}ms`)
  }
}

// 监听分页变化
watch([currentPage, pageSize], () => {
  loadDocumentList()
})

// 组件挂载时加载数据
onMounted(() => {
  loadDocumentList()
})

/**
 * 处理文件选择
 */
const handleFileChange = (file, fileList) => {
  // 文件选择后自动上传
  uploadFile(file.raw || file)
}

/**
 * 上传文件
 */
async function uploadFile(file) {
  if (!file) return
  
  // 添加到上传列表
  const fileItem = {
    uid: file.uid || Date.now() + Math.random(),
    name: file.name,
    size: file.size,
    status: 'uploading',
    progress: 0
  }
  uploadedFiles.value.push(fileItem)
  
  try {
    // 调用上传 API
    const result = await uploadDocument(props.knowledgeBaseUuid, file)

    const index = uploadedFiles.value.findIndex(f => f.uid === fileItem.uid)

    if (index > -1) {
      uploadedFiles.value[index].status = 'success'
      uploadedFiles.value[index].progress = 100
      uploadedFiles.value[index].storageDocUuid = result.storage_doc_uuid
    }

    if (result?.is_renamed) {
      ElMessage.info(`文件名已自动重命名为「${result.document_name}」`)
    }

    ElMessage.success(`文件"${file.name}"上传成功`)
  } catch (error) {
    // 更新错误状态
    const index = uploadedFiles.value.findIndex(f => f.uid === fileItem.uid)
    if (index > -1) {
      uploadedFiles.value[index].status = 'error'
    }
    ElMessage.error(`文件"${file.name}"上传失败: ${error.message || '未知错误'}`)
  }
}

/**
 * 移除文件
 */
const handleRemoveFile = (file) => {
  const index = uploadedFiles.value.findIndex(f => f.uid === file.uid)
  if (index > -1) {
    uploadedFiles.value.splice(index, 1)
  }
}

/**
 * 完成上传并处理文档
 */
const handleCompleteUpload = async () => {
  const startTime = performance.now()
  console.log('[handleCompleteUpload] ===== 开始处理文档 =====', new Date().toISOString())
  
  const successFiles = uploadedFiles.value.filter(f => f.status === 'success' && f.storageDocUuid)
  
  if (successFiles.length === 0) {
    ElMessage.warning('没有成功上传的文件')
    return
  }
  
  console.log(`[handleCompleteUpload] 成功上传的文件数: ${successFiles.length}`)
  
  isProcessing.value = true
  
  // 存储所有 WebSocket 连接，用于后续关闭
  const wsConnections = []
  
  // 跟踪每个文档是否已收到第一次 progress 消息
  const receivedFirstProgress = new Set()
  let listLoaded = false // 标记列表是否已加载
  
  try {
    // 为每个文档建立 WebSocket 连接并处理
    successFiles.forEach(async (fileItem) => {
      const fileStartTime = performance.now()
      console.log(`[handleCompleteUpload] 开始处理文件: ${fileItem.name}`, `耗时: ${(fileStartTime - startTime).toFixed(2)}ms`)
      
      try {
        // 先建立 WebSocket 连接
        const wsStartTime = performance.now()
        console.log(`[${fileItem.name}] 开始建立 WebSocket 连接...`)
        
        const ws = await createDocumentProcessWebSocket(
          fileItem.storageDocUuid,
          // onProgress - 进度更新
          (progress) => {
            const progressTime = performance.now()
            console.log(`[${fileItem.name}] 处理进度:`, progress.progress, progress.status, `耗时: ${(progressTime - startTime).toFixed(2)}ms`)
            
            // 更新文档列表中的进度和状态
            updateDocumentProgress(fileItem.storageDocUuid, progress.progress || 0, progress.status)
            
            // 如果是第一次收到 progress 消息，标记为已收到
            if (!receivedFirstProgress.has(fileItem.storageDocUuid)) {
              const firstProgressTime = performance.now()
              receivedFirstProgress.add(fileItem.storageDocUuid)
              console.log(`[${fileItem.name}] ✓ 收到第一次进度消息`, `耗时: ${(firstProgressTime - startTime).toFixed(2)}ms`)
              
              // 检查是否所有文档都收到了第一次 progress 消息
              if (!listLoaded && receivedFirstProgress.size === successFiles.length) {
                const allProgressTime = performance.now()
                console.log(`[handleCompleteUpload] ✓ 所有文档都已收到第一次进度消息，开始加载列表`, `耗时: ${(allProgressTime - startTime).toFixed(2)}ms`)
                listLoaded = true
                loadDocumentList()
              }
            }
          },
          // onComplete - 处理完成（异步更新状态）
          (complete) => {
            const completeTime = performance.now()
            console.log(`[${fileItem.name}] ✓ 处理完成:`, complete, `耗时: ${(completeTime - startTime).toFixed(2)}ms`)
            // 更新文档列表中的状态（局部更新，不触发全量刷新）
            updateDocumentStatus(fileItem.storageDocUuid, complete.data)
          },
          // onError - 处理错误
          (error) => {
            const errorTime = performance.now()
            console.error(`[${fileItem.name}] ✗ 处理错误:`, error, `耗时: ${(errorTime - startTime).toFixed(2)}ms`)
            ElMessage.error(`处理文档"${fileItem.name}"失败: ${error.error || '未知错误'}`)
            
            // 即使出错，也标记为已收到（避免一直等待）
            if (!receivedFirstProgress.has(fileItem.storageDocUuid)) {
              receivedFirstProgress.add(fileItem.storageDocUuid)
              if (!listLoaded && receivedFirstProgress.size === successFiles.length) {
                const allProgressTime = performance.now()
                console.log(`[handleCompleteUpload] ✓ 所有文档都已处理（包括错误），开始加载列表`, `耗时: ${(allProgressTime - startTime).toFixed(2)}ms`)
                listLoaded = true
                loadDocumentList()
              }
            }
          }
        )
        
        const wsEndTime = performance.now()
        console.log(`[${fileItem.name}] ✓ WebSocket 连接建立完成`, `耗时: ${(wsEndTime - wsStartTime).toFixed(2)}ms (总耗时: ${(wsEndTime - startTime).toFixed(2)}ms)`)
        
        wsConnections.push(ws)
        
        // WebSocket 连接建立后，调用处理接口（不等待返回）
        const apiStartTime = performance.now()
        console.log(`[${fileItem.name}] 开始调用处理接口...`)
        
        processDocument(fileItem.storageDocUuid).then(() => {
          const apiEndTime = performance.now()
          console.log(`[${fileItem.name}] ✓ 处理接口调用成功`, `耗时: ${(apiEndTime - apiStartTime).toFixed(2)}ms (总耗时: ${(apiEndTime - startTime).toFixed(2)}ms)`)
        }).catch(error => {
          const apiErrorTime = performance.now()
          console.error(`[${fileItem.name}] ✗ 调用处理接口失败:`, error, `耗时: ${(apiErrorTime - apiStartTime).toFixed(2)}ms (总耗时: ${(apiErrorTime - startTime).toFixed(2)}ms)`)
          ElMessage.error(`启动处理"${fileItem.name}"失败: ${error.message || '未知错误'}`)
          
          // API 调用失败时，也标记为已处理（避免一直等待）
          if (!receivedFirstProgress.has(fileItem.storageDocUuid)) {
            receivedFirstProgress.add(fileItem.storageDocUuid)
            if (!listLoaded && receivedFirstProgress.size === successFiles.length) {
              const allProgressTime = performance.now()
              console.log(`[handleCompleteUpload] ✓ 所有文档都已处理（包括API失败），开始加载列表`, `耗时: ${(allProgressTime - startTime).toFixed(2)}ms`)
              listLoaded = true
              loadDocumentList()
            }
          }
        })
      } catch (error) {
        const wsErrorTime = performance.now()
        console.error(`[${fileItem.name}] ✗ 建立 WebSocket 连接失败:`, error, `耗时: ${(wsErrorTime - startTime).toFixed(2)}ms`)
        ElMessage.error(`建立连接"${fileItem.name}"失败: ${error.message || '未知错误'}`)
        
        // WebSocket 连接失败时，也标记为已处理（避免一直等待）
        if (!receivedFirstProgress.has(fileItem.storageDocUuid)) {
          receivedFirstProgress.add(fileItem.storageDocUuid)
          if (!listLoaded && receivedFirstProgress.size === successFiles.length) {
            const allProgressTime = performance.now()
            console.log(`[handleCompleteUpload] ✓ 所有文档都已处理（包括连接失败），开始加载列表`, `耗时: ${(allProgressTime - startTime).toFixed(2)}ms`)
            listLoaded = true
            loadDocumentList()
          }
        }
      }
    })
    
    const dialogCloseTime = performance.now()
    console.log(`[handleCompleteUpload] 准备关闭对话框...`, `耗时: ${(dialogCloseTime - startTime).toFixed(2)}ms`)
    
    ElMessage.success('文档处理已开始')
    
    // 立即关闭对话框（不等待 API 返回）
    handleCloseUploadDialog()
    
    const afterCloseTime = performance.now()
    console.log(`[handleCompleteUpload] ✓ 对话框已关闭`, `耗时: ${(afterCloseTime - startTime).toFixed(2)}ms`)
    
    // 显示加载状态（使用和列表加载相同的 loading 状态）
    loading.value = true
    console.log(`[handleCompleteUpload] ✓ 加载状态已显示`, `耗时: ${(performance.now() - startTime).toFixed(2)}ms`)
  } catch (error) {
    console.error('处理文档失败:', error)
    ElMessage.error('处理文档失败')
    // 发生错误时，也要隐藏加载状态
    loading.value = false
  } finally {
    isProcessing.value = false
    // 可以选择保持 WebSocket 连接打开，或者在这里关闭
    // wsConnections.forEach(ws => {
    //   if (ws.readyState === WebSocket.OPEN) {
    //     ws.close()
    //   }
    // })
  }
}

/**
 * 更新文档列表中的进度和状态
 */
function updateDocumentProgress(storageDocUuid, progress, status) {
  const docIndex = documentList.value.findIndex(doc => doc.storageDocUuid === storageDocUuid)
  if (docIndex > -1) {
    console.log(`[updateDocumentProgress] 更新文档 ${storageDocUuid} 的进度:`, progress, '状态:', status)
    
    // 更新进度（确保是数字类型，保持0-1的原始值）
    const progressValue = typeof progress === 'number' ? progress : (parseFloat(progress) || 0)
    documentList.value[docIndex].processingProgress = progressValue
    
    // 更新状态
    if (status) {
      // 根据 processing_status 判断状态
      if (status === 'processing' || status === 'pending') {
        documentList.value[docIndex].status = '处理中'
      } else if (status === 'completed' || status === 'success') {
        documentList.value[docIndex].status = '可用'
        documentList.value[docIndex].processingProgress = 1 // 完成时设置为1（100%）
      } else if (status === 'failed' || status === 'error') {
        documentList.value[docIndex].status = '处理失败'
      }
    } else {
      // 如果没有状态，但进度小于1，设置为处理中
      if (progressValue < 1) {
        documentList.value[docIndex].status = '处理中'
      }
    }
    
    console.log(`[updateDocumentProgress] 更新后:`, {
      status: documentList.value[docIndex].status,
      progress: documentList.value[docIndex].processingProgress,
      percentage: getProgressPercentage(progressValue)
    })
  } else {
    console.warn(`[updateDocumentProgress] 未找到文档: ${storageDocUuid}`)
  }
}

/**
 * 更新文档列表中的状态
 * 注意：状态列只根据 processing_status 更新，启用列只根据 is_enabled 更新
 */
function updateDocumentStatus(storageDocUuid, documentData) {
  const docIndex = documentList.value.findIndex(doc => doc.storageDocUuid === storageDocUuid)
  if (docIndex > -1 && documentData) {
    // 更新进度
    if (documentData.hasOwnProperty('processing_progress')) {
      documentList.value[docIndex].processingProgress = documentData.processing_progress || 0
    }
    
    // 更新状态列（只根据 processing_status）
    if (documentData.hasOwnProperty('processing_status')) {
      documentList.value[docIndex].status = getStatusFromProcessingStatus(
        documentData.processing_status,
        documentData.processing_progress
      )
    }
    
    // 更新启用列（只根据 is_enabled）
    if (documentData.hasOwnProperty('is_enabled')) {
      documentList.value[docIndex].enabled = documentData.is_enabled || false
    }
  }
}

/**
 * 关闭上传对话框
 */
const handleCloseUploadDialog = () => {
  showUploadDialog.value = false
  uploadedFiles.value = []
  if (uploadRef.value) {
    uploadRef.value.clearFiles()
  }
}

// 返回知识库列表
const handleBack = () => {
  emit('back')
}

// 切换关注状态
const toggleFavorite = () => {
  isFavorite.value = !isFavorite.value
  console.log('切换关注状态:', isFavorite.value)
  // TODO: 实现关注逻辑
}

// 添加标签
const handleAddTag = (row) => {
  console.log('添加标签:', row)
  // TODO: 实现添加标签逻辑
}

// 切换启用状态
const handleToggleEnable = async (row) => {
  const newEnabled = row.enabled
  const storageDocUuid = row.storageDocUuid
  
  console.log('[handleToggleEnable] 切换文档启用状态:', {
    storageDocUuid,
    documentName: row.name,
    newEnabled
  })
  
  try {
    // 调用 API 设置启用状态
    const result = await setDocumentEnableStatus(storageDocUuid, newEnabled)
    
    console.log('[handleToggleEnable] 设置成功:', result)
    
    // 更新本地状态
    const docIndex = documentList.value.findIndex(doc => doc.storageDocUuid === storageDocUuid)
    if (docIndex > -1 && result) {
      if (result.hasOwnProperty('is_enabled')) {
        documentList.value[docIndex].enabled = result.is_enabled
      }
    }
    
    ElMessage.success(newEnabled ? '文档已启用' : '文档已禁用')
  } catch (error) {
    console.error('[handleToggleEnable] 设置失败:', error)
    
    // 恢复开关状态（因为操作失败）
    row.enabled = !newEnabled
    
    ElMessage.error(error.message || '设置文档启用状态失败')
  }
}

// 鼠标进入行
const handleRowEnter = (row) => {
  row.showAddTag = true
}

// 鼠标离开行
const handleRowLeave = (row) => {
  row.showAddTag = false
}

// 打开文档详情页
const handleOpenDocumentDetail = (row) => {
  selectedDocument.value = row
  showDocumentDetail.value = true
}

// 从文档详情页返回
const handleBackFromDetail = () => {
  showDocumentDetail.value = false
  selectedDocument.value = null
}

// 表格行点击事件
const handleRowClick = (row, column, event) => {
  // 检查点击的目标元素，如果是按钮、链接、下拉菜单等交互元素，则不触发跳转
  const target = event.target
  const isInteractiveElement = target.closest('button') || 
                                target.closest('a') || 
                                target.closest('.el-link') ||
                                target.closest('.el-dropdown') ||
                                target.closest('.el-icon') ||
                                target.closest('.el-switch') ||
                                target.closest('.el-checkbox') ||
                                target.closest('.add-tag-container') ||
                                target.closest('.doc-name-cell')
  
  if (!isInteractiveElement) {
    handleOpenDocumentDetail(row)
  }
}

// 表格选中变化事件
const handleSelectionChange = (selection) => {
  selectedDocuments.value = selection
  console.log('[handleSelectionChange] 选中的文档:', selection.length, selection)
}

// 删除按钮点击事件
const handleDeleteClick = async () => {
  if (selectedDocuments.value.length === 0) {
    ElMessage.warning('请先选择要删除的文档')
    return
  }
  
  try {
    // 构建确认消息
    const documentNames = selectedDocuments.value.map(doc => doc.name).join('、')
    const confirmMessage = selectedDocuments.value.length === 1
      ? `确定要删除文档"${documentNames}"吗？删除后该文档及其关联数据将被永久删除，操作不可恢复。`
      : `确定要删除选中的 ${selectedDocuments.value.length} 个文档吗？删除后这些文档及其关联数据将被永久删除，操作不可恢复。`
    
    await ElMessageBox.confirm(
      confirmMessage,
      '确认删除',
      {
        confirmButtonText: '确定',
        cancelButtonText: '取消',
        type: 'warning'
      }
    )
    
    // 获取要删除的文档UUID列表
    const storageDocUuids = selectedDocuments.value.map(doc => doc.storageDocUuid)
    
    console.log('[handleDeleteClick] 开始删除文档:', storageDocUuids)
    
    // 调用删除API
    const result = await deleteDocuments(storageDocUuids)
    
    console.log('[handleDeleteClick] 删除结果:', result)
    
    // 显示删除结果消息
    if (result.success_count > 0) {
      let message = `成功删除 ${result.success_count} 个文档`
      const details = []
      
      if (result.not_found_count > 0) {
        details.push(`${result.not_found_count} 个文档不存在`)
      }
      if (result.permission_denied_count > 0) {
        details.push(`${result.permission_denied_count} 个文档无权限删除`)
      }
      if (result.failed_count > 0) {
        details.push(`${result.failed_count} 个文档删除失败`)
      }
      
      if (details.length > 0) {
        message += `；${details.join('；')}`
      }
      
      ElMessage.success(message)
    } else {
      // 全部失败的情况
      let errorMessage = '删除失败'
      if (result.not_found_count > 0) {
        errorMessage = `${result.not_found_count} 个文档不存在`
      } else if (result.permission_denied_count > 0) {
        errorMessage = `无权限删除 ${result.permission_denied_count} 个文档`
      } else if (result.failed_count > 0) {
        errorMessage = `${result.failed_count} 个文档删除失败`
      }
      ElMessage.warning(errorMessage)
    }
    
    // 清除选中状态
    if (tableRef.value) {
      tableRef.value.clearSelection()
    }
    selectedDocuments.value = []
    
    // 重新加载文档列表
    await loadDocumentList()
    
  } catch (error) {
    if (error !== 'cancel') {
      console.error('删除文档失败:', error)
      ElMessage.error(error.message || '删除文档失败')
    }
  }
}
</script>

<style scoped>
.document-list-container {
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

.kb-icon-small {
  width: 24px;
  height: 24px;
  border-radius: 50%;
  background: linear-gradient(135deg, #b37feb 0%, #722ed1 100%);
  display: flex;
  align-items: center;
  justify-content: center;
  color: #fff;
  font-weight: 600;
  font-size: 12px;
}

.kb-title {
  margin: 0;
  font-size: 18px;
  font-weight: 500;
  color: #333;
  line-height: 1.5;
}

.doc-count-info {
  display: flex;
  align-items: center;
  gap: 4px;
  color: #666;
  font-size: 14px;
}

.star-icon-header {
  font-size: 18px;
  color: #999;
  cursor: pointer;
  transition: color 0.2s;
}

.star-icon-header:hover {
  color: #f5a623;
}

.star-icon-header.is-favorite {
  color: #f5a623;
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

.add-doc-btn {
  margin-left: auto;
}

/* 表格区域 */
.table-section {
  flex: 1;
  padding: 24px;
  overflow: auto;
  background-color: #fff;
}

.doc-table {
  background-color: #fff;
}

/* 确保表格可以横向滚动 */
.doc-table :deep(.el-table__body-wrapper) {
  overflow-x: auto;
}

.doc-table :deep(.el-table) {
  width: 100%;
  min-width: 100%;
}

.doc-table :deep(.el-table__header-wrapper) {
  background-color: #f5f5f5;
}

.doc-table :deep(.el-table__body-wrapper) {
  background-color: #fff;
}

.doc-table :deep(.el-table th) {
  background-color: #f5f5f5 !important;
  color: #333;
  font-weight: normal;
}

.doc-table :deep(.el-table td) {
  background-color: #fff;
  padding-top: 16px;
  padding-bottom: 16px;
}

.doc-table :deep(.el-table tr:hover > td) {
  background-color: #f5f7fa;
}

.doc-table :deep(.el-table__row) {
  cursor: pointer;
  height: auto;
}

.doc-name-cell {
  display: flex;
  align-items: center;
  gap: 12px;
  cursor: pointer;
}

.doc-icon {
  width: 32px;
  height: 32px;
  min-width: 32px;
  min-height: 32px;
  max-width: 32px;
  max-height: 32px;
  flex-shrink: 0;
  border-radius: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  color: #fff;
  font-weight: 600;
  font-size: 14px;
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
  font-size: 10px;
}

.doc-icon-ppt {
  background-color: #D24726;
  width: auto;
  min-width: 32px;
  max-width: none;
  padding: 0 4px;
  font-size: 10px;
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

.doc-name {
  color: #333;
  font-size: 14px;
}

.tags-cell {
  display: flex;
  align-items: center;
  gap: 8px;
  position: relative;
}

.tags-list {
  display: flex;
  align-items: center;
  gap: 4px;
}

.add-tag-container {
  display: inline-block;
  padding: 2px 8px;
  border: 1px dashed #d9d9d9;
  border-radius: 4px;
  background-color: #fff;
  cursor: pointer;
  transition: all 0.2s;
  opacity: 0;
}

.add-tag-container.is-hover {
  opacity: 1;
}

.add-tag-text {
  font-size: 11px;
  color: #666;
  line-height: 1.4;
}

.add-tag-container:hover {
  border-color: #409eff;
  color: #409eff;
}

.add-tag-container:hover .add-tag-text {
  color: #409eff;
}

/* 确保 hover 时显示添加标签容器 */
.doc-table :deep(.el-table__row:hover .add-tag-container) {
  opacity: 1 !important;
}

/* 处理中状态样式 */
.status-processing {
  width: 100%;
}

.status-processing :deep(.el-progress) {
  width: 100%;
}

.status-processing :deep(.el-progress__text) {
  font-size: 12px;
  color: #666;
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

.total-count {
  color: #666;
  font-size: 14px;
}

.selection-info {
  display: flex;
  align-items: center;
  gap: 12px;
}

.selected-count {
  color: #333;
  font-size: 14px;
  font-weight: 500;
}

.delete-btn {
  background-color: #f56c6c;
  border-color: #f56c6c;
  color: #fff;
}

.delete-btn:hover {
  background-color: #f78989;
  border-color: #f78989;
}

.footer-right {
  display: flex;
  align-items: center;
}

/* 上传对话框样式 */
.upload-container {
  margin-bottom: 20px;
}

.upload-file-list {
  margin-top: 20px;
  max-height: 300px;
  overflow-y: auto;
}

.upload-file-item {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 12px;
  border: 1px solid #e4e7ed;
  border-radius: 4px;
  margin-bottom: 8px;
}

.file-info {
  display: flex;
  flex-direction: column;
  flex: 1;
}

.file-name {
  font-size: 14px;
  color: #303133;
  margin-bottom: 4px;
}

.file-size {
  font-size: 12px;
  color: #909399;
}

.file-status {
  width: 120px;
  display: flex;
  justify-content: flex-end;
  align-items: center;
}

.success-icon {
  color: #67c23a;
  font-size: 20px;
}

.error-icon {
  color: #f56c6c;
  font-size: 20px;
}

.uploading-text {
  color: #909399;
  font-size: 14px;
}

.dialog-footer {
  display: flex;
  justify-content: flex-end;
  gap: 12px;
}
</style>

