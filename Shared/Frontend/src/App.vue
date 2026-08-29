<template>
  <div
    class="app-shell"
    :class="{
      'host-desktop': isDesktopHost,
      'layout-compact': isDesktopHost && layoutMode === 'compact'
    }"
  >
    <!-- 完整版：任务侧栏；缩小版对齐插件，无侧栏 -->
    <TaskSidebar
      v-if="isDesktopHost && layoutMode === 'expanded'"
      :collapsed="sidebarCollapsed"
      :tasks="displayTaskList"
      :open-files="openFiles"
      :selected-open-file-ids="selectedOpenFileIds"
      :active-id="activeTaskId"
      :loading="taskListLoading"
      :error="taskListError"
      :loading-more="taskListLoadingMore"
      :has-more="taskListHasMore"
      @toggle="handleSidebarToggle"
      @new-task="handleDesktopNewTask"
      @select="handleDesktopSelectTask"
      @select-open-file="handleSelectOpenFile"
      @load-more="loadMoreTasks"
    />
    <div
      class="chat-container"
      @dragenter.prevent="onChatDragEnter"
      @dragleave.prevent="onChatDragLeave"
      @dragover.prevent="onChatDragOver"
      @drop.prevent="onChatDrop"
    >
      <div v-if="chatDragOver" class="chat-drop-hint">拖放到此处以上传为对话附件</div>
      <!-- 插件顶栏；缩小版顶栏：历史 / 新建 / 设置 -->
      <ChatHeader v-if="!isDesktopHost" />
      <div
        v-else-if="layoutMode === 'compact'"
        class="compact-chrome"
        @keydown.esc.stop="historyPopoverOpen = false"
      >
        <ChatHeader
          variant="compact"
          @history="toggleHistoryPopover"
          @add="handleDesktopNewTask"
        />
        <ConversationHistoryPopover
          v-if="historyPopoverOpen"
          :tasks="displayTaskList"
          :loading="taskListLoading"
          :error="taskListError"
          :loading-more="taskListLoadingMore"
          :has-more="taskListHasMore"
          :active-id="activeTaskId"
          @select="handleCompactSelectTask"
          @close="historyPopoverOpen = false"
          @load-more="loadMoreTasks"
        />
      </div>
      <ChatMessages
        v-if="messages.length > 0"
        :messages="messages"
        :pin-to-latest-token="pinToLatestToken"
      />
      <ChatEmptyState
        v-else
        :operation-guide="emptyGuide"
        :recommendations="emptyRecommendations"
        :disabled="isProcessing"
        @select-recommendation="handleSend"
      />
      <div v-if="panelHint" class="panel-hint">{{ panelHint }}</div>
      <TodoProgressBar
        :visible="todoBarVisible"
        :expanded="todoBarExpanded"
        :todos="todoBarTodos"
        @toggle="todoBarExpanded = !todoBarExpanded"
      />
      <ChatInput
        @send="handleSend"
        @stop="handleUserStop"
        @clear-attachment="clearChatAttachment"
        @dropped-file="onInputAreaFileDrop"
        @remove-selected-open-file="handleRemoveSelectedOpenFile"
        :loading="isProcessing"
        :isProcessing="isProcessing"
        :desktop="isDesktopHost && layoutMode === 'expanded'"
        :show-open-file-chips="isDesktopHost"
        :attachment="attachmentView"
        :selected-open-files="selectedOpenFiles"
      />
    </div>
  </div>
</template>

<script setup>
import { ref, computed, watch, onMounted, onUnmounted, nextTick } from 'vue'
import ChatHeader from './components/ChatHeader.vue'
import ChatMessages from './components/ChatMessages.vue'
import ChatEmptyState from './components/ChatEmptyState.vue'
import ChatInput from './components/ChatInput.vue'
import TodoProgressBar from './components/TodoProgressBar.vue'
import TaskSidebar from './components/TaskSidebar.vue'
import ConversationHistoryPopover from './components/ConversationHistoryPopover.vue'
import { useWebViewBridge } from './composables/useWebViewBridge'
import { useChatFileUpload } from './composables/useChatFileUpload'
import { ToolCallAccumulator } from './utils/toolCallAccumulator'
import { fetchChatEmptyState } from './services/chatEmptyStateApi.js'
import { listConversations, CONVERSATION_PAGE_SIZE } from './services/conversationsApi.js'
import { clearApiConfigCache } from './services/knowledgeBaseApi.js'
import {
  MAX_SELECTED_OPEN_FILES,
  selectedOpenFilesState,
  toSelectedOpenFile,
  appendToUserContent,
  stripSelectedOpenFilesAppendix,
  pruneSelectionByOpenFiles
} from './utils/selectedOpenFiles.js'
import {
  DRAFT_TASK_ID,
  makeDraftTaskItem,
  stashDraftInputFromLive,
  peekStashedDraftInput,
  clearStashedDraftInput,
  clearAllDraftInput
} from './utils/draftChatInput.js'

const { sendMessage, onMessage } = useWebViewBridge()
const chatFile = useChatFileUpload()

function detectDesktopHost () {
  try {
    return new URLSearchParams(window.location.search).get('host') === 'desktop'
  } catch {
    return false
  }
}

const isDesktopHost = detectDesktopHost()
/** Desktop 窗口形态：完整版 expanded / 缩小版 compact（由 C# layoutModeChanged 驱动） */
const layoutMode = ref('expanded')
const sidebarCollapsed = ref(false)
/** 递增后 ChatMessages 强制钉到最新（切历史 / 切形态） */
const pinToLatestToken = ref(0)

const historyPopoverOpen = ref(false)

function reportCompactUiBusy (busy) {
  if (!isDesktopHost || layoutMode.value !== 'compact') return
  sendMessage('compactUiBusy', { busy: !!busy }).catch(() => {})
}

function applyLayoutMode (mode) {
  if (mode !== 'compact' && mode !== 'expanded') return
  const changed = layoutMode.value !== mode
  const wasCompact = layoutMode.value === 'compact'
  layoutMode.value = mode
  if (historyPopoverOpen.value) {
    historyPopoverOpen.value = false
  } else if (wasCompact && mode !== 'compact') {
    reportCompactUiBusy(false)
  }
  // 缩小版无侧栏；完整版默认展开侧栏
  sidebarCollapsed.value = mode !== 'expanded'
  if (!changed) return
  // 形态一变容器高度先抖：立刻重新钉底，避免 scroll 把跟随打掉后流式不再贴最新
  nextTick(() => {
    pinToLatestToken.value++
  })
}

watch(historyPopoverOpen, (open) => {
  reportCompactUiBusy(open)
})

function handleSidebarToggle () {
  if (!isDesktopHost || layoutMode.value !== 'expanded') return
  sidebarCollapsed.value = !sidebarCollapsed.value
}

async function toggleHistoryPopover () {
  if (historyPopoverOpen.value) {
    historyPopoverOpen.value = false
    return
  }
  historyPopoverOpen.value = true
  await refreshTaskList()
  // watch 会 reportCompactUiBusy(true)
}

async function handleCompactSelectTask (item) {
  await handleDesktopSelectTask(item)
  historyPopoverOpen.value = false
}
const taskList = ref([])
const taskListLoading = ref(false)
const taskListLoadingMore = ref(false)
const taskListHasMore = ref(false)
const taskListPage = ref(0)
const taskListError = ref('')
const activeTaskId = ref(null)
/** 从「新对话」点进历史后，列表顶保留可切回的「新对话」 */
const draftSlotVisible = ref(false)
/** addConversation 会 clearMessages，延后到清空后再还原输入框 */
let pendingRestoreDraftInput = null

const displayTaskList = computed(() => {
  const list = Array.isArray(taskList.value) ? taskList.value : []
  if (!draftSlotVisible.value) return list
  if (list.some((t) => String(t.id) === DRAFT_TASK_ID || t.isDraft)) return list
  return [makeDraftTaskItem(), ...list]
})

function isOnNewChatDraft () {
  if (String(activeTaskId.value) === DRAFT_TASK_ID) return true
  return activeTaskId.value == null
    || activeTaskId.value === ''
    || String(activeTaskId.value) === '-1'
}

function clearInputBox () {
  window.dispatchEvent(new CustomEvent('clearInput'))
}

function restoreInputBox (value) {
  restoreInputValue(value == null ? '' : String(value))
}
const openFiles = ref([])
/** Desktop：已选打开文件（芯片 / 发送附加段）；不落库、不跟会话走 */
const selectedOpenFiles = selectedOpenFilesState
const selectedOpenFileIds = computed(() => selectedOpenFiles.value.map((x) => x.id))

watch(openFiles, (list) => {
  // 仅在拿到真实列表时 prune；缺字段/非数组不碰选中，避免切会话误清芯片
  if (!Array.isArray(list)) return
  selectedOpenFiles.value = pruneSelectionByOpenFiles(selectedOpenFiles.value, list)
})

function handleSelectOpenFile (item) {
  if (!isDesktopHost) return
  const sel = toSelectedOpenFile(item)
  if (selectedOpenFiles.value.some((x) => x.id === sel.id)) return
  if (selectedOpenFiles.value.length >= MAX_SELECTED_OPEN_FILES) {
    panelHint.value = '最多选择 15 个文件'
    setTimeout(() => {
      if (panelHint.value === '最多选择 15 个文件') panelHint.value = ''
    }, 2000)
    return
  }
  selectedOpenFiles.value = [...selectedOpenFiles.value, sel]
}

function handleRemoveSelectedOpenFile (id) {
  selectedOpenFiles.value = selectedOpenFiles.value.filter((x) => x.id !== id)
}

async function refreshOpenFiles () {
  if (!isDesktopHost) return
  try {
    const res = await sendMessage('getOpenFiles', {})
    const items = res?.items ?? res?.data?.items
    openFiles.value = Array.isArray(items) ? items : []
  } catch (e) {
    console.warn('[App] getOpenFiles failed:', e?.message || e)
  }
}

async function refreshTaskList () {
  if (!isDesktopHost) return
  taskListLoading.value = true
  taskListError.value = ''
  try {
    const data = await listConversations({ page: 1, pageSize: CONVERSATION_PAGE_SIZE })
    taskList.value = data.conversations
    taskListPage.value = data.page
    taskListHasMore.value = !!data.hasMore
  } catch (e) {
    const msg = String(e?.message || '')
    const authFailed = /401|403|Not authenticated|未登录|未认证/i.test(msg)
    taskListError.value = authFailed ? '' : (e?.message || '加载任务失败')
    taskList.value = []
    taskListPage.value = 0
    taskListHasMore.value = false
  } finally {
    taskListLoading.value = false
  }
}

async function loadMoreTasks () {
  if (!isDesktopHost || taskListLoading.value || taskListLoadingMore.value || !taskListHasMore.value) {
    return
  }

  taskListLoadingMore.value = true
  try {
    const data = await listConversations({
      page: taskListPage.value + 1,
      pageSize: CONVERSATION_PAGE_SIZE
    })
    const seen = new Set(taskList.value.map((t) => String(t.id)))
    const extra = data.conversations.filter((t) => !seen.has(String(t.id)))
    taskList.value = [...taskList.value, ...extra]
    taskListPage.value = data.page
    taskListHasMore.value = !!data.hasMore
  } catch (e) {
    console.warn('[App] loadMoreTasks failed:', e?.message || e)
  } finally {
    taskListLoadingMore.value = false
  }
}

async function handleDesktopNewTask () {
  try {
    draftSlotVisible.value = false
    clearAllDraftInput()
    clearInputBox()
    await sendMessage('addConversation', {})
    activeTaskId.value = DRAFT_TASK_ID
    await refreshTaskList()
  } catch (e) {
    console.error('[App] 新建任务失败:', e)
  }
}

async function handleDesktopSelectTask (item) {
  if (!item?.id) return

  // 切回「新对话」草稿
  if (item.isDraft || String(item.id) === DRAFT_TASK_ID) {
    try {
      const text = peekStashedDraftInput()
      pendingRestoreDraftInput = text
      await sendMessage('addConversation', {})
      activeTaskId.value = DRAFT_TASK_ID
      draftSlotVisible.value = false
      await refreshTaskList()
      // 若宿主未推 clearMessages，仍兜底还原
      await nextTick()
      if (pendingRestoreDraftInput != null) {
        restoreInputBox(pendingRestoreDraftInput)
        pendingRestoreDraftInput = null
        clearStashedDraftInput()
      }
    } catch (e) {
      pendingRestoreDraftInput = null
      console.error('[App] 切回新对话失败:', e)
    }
    return
  }

  try {
    if (isOnNewChatDraft()) {
      stashDraftInputFromLive()
      draftSlotVisible.value = true
      clearInputBox()
    }
    activeTaskId.value = item.id
    const res = await sendMessage('openConversation', { id: item.id })
    if (!res?.success) {
      console.error('[App] 打开任务失败:', res?.message)
      return
    }
  } catch (e) {
    console.error('[App] 打开任务失败:', e)
  }
}

const emptyGuide = ref('')
const emptyRecommendations = ref([])

async function loadEmptyState () {
  emptyGuide.value = ''
  emptyRecommendations.value = []
  try {
    let documentEmpty = true
    try {
      const res = await sendMessage('getDocumentEmptyState', {})
      if (res?.success && res.data && typeof res.data.documentEmpty === 'boolean') {
        documentEmpty = res.data.documentEmpty
      }
    } catch (e) {
      console.warn('[App] getDocumentEmptyState 失败，按空白文档处理:', e)
    }
    const data = await fetchChatEmptyState(documentEmpty)
    emptyGuide.value = data.operation_guide || ''
    emptyRecommendations.value = Array.isArray(data.recommendations) ? data.recommendations : []
  } catch (e) {
    console.warn('[App] 拉取空状态引导失败（静默）:', e?.message || e)
  }
}

const chatDragOver = ref(false)
let chatDragDepth = 0

const attachmentView = computed(() => ({
  phase: chatFile.phase.value,
  progress: chatFile.progress.value,
  fileLabel: chatFile.fileLabel.value,
  errorMessage: chatFile.errorMessage.value,
  /** 与输入区附件条展示绑定：ready 态仅在有 upload_id 时显示 */
  uploadId: chatFile.uploadId.value
}))

function onChatDragEnter () {
  chatDragDepth++
  chatDragOver.value = true
}

function onChatDragOver (e) {
  e.preventDefault()
  if (e.dataTransfer) {
    try {
      e.dataTransfer.dropEffect = 'copy'
    } catch (_) {
      /* ignore */
    }
  }
}

function onChatDragLeave (e) {
  const next = e.relatedTarget
  if (next && e.currentTarget && e.currentTarget.contains(next)) {
    return
  }
  chatDragDepth = Math.max(0, chatDragDepth - 1)
  if (chatDragDepth === 0) chatDragOver.value = false
}

function onChatDrop (e) {
  e.preventDefault()
  chatDragDepth = 0
  chatDragOver.value = false
  const file = e.dataTransfer?.files?.[0]
  if (file) {
    chatFile.startUpload(file)
  }
}

/** 输入区在捕获阶段已拦截文件拖放并上报，与消息区拖放共用上传逻辑 */
function onInputAreaFileDrop (file) {
  chatDragDepth = 0
  chatDragOver.value = false
  if (file) {
    chatFile.startUpload(file)
  }
}

function clearChatAttachment () {
  chatFile.reset()
}

const messages = ref([])
const toolCallAccumulators = new Map()
const isProcessing = ref(false)
const panelHint = ref('')
const pendingInput = ref('') // 保存待发送的消息（用于登录失败时恢复）
const receivedBackendMessages = ref(new Set()) // 保存已收到后端消息的用户消息ID（用于判断是否应该移除）

// 清除输入框的方法
const clearInput = () => {
  const event = new CustomEvent('clearInput')
  window.dispatchEvent(event)
}

/** 用户点停止：当前助手轮没有自然收尾，下一问时整轮收进「奋力工作」 */
function handleUserStop () {
  for (let i = messages.value.length - 1; i >= 0; i--) {
    const m = messages.value[i]
    if (m.role === 'user') break
    if (m.role === 'system' && !m.isHint) {
      m.aborted = true
    }
  }
}

// 处理发送消息
const handleSend = async (content) => {
  if (!content.trim() || isProcessing.value) return

  if (chatFile.uploadId.value && chatFile.phase.value !== 'ready') {
    console.warn('[App] 附件未就绪，跳过发送')
    return
  }

  const raw = content.trim()

  // 保存输入内容（用于登录失败时恢复）
  pendingInput.value = raw
  panelHint.value = ''

  // 立即设置处理状态，确保按钮显示停止图标
  isProcessing.value = true
  console.log('[App] 开始发送消息，设置 isProcessing = true')

  const uploadIdSnapshot = chatFile.uploadId.value
  const attachmentSnap =
    uploadIdSnapshot && chatFile.fileLabel.value?.name
      ? {
          fileName: chatFile.fileLabel.value.name,
          ext: chatFile.fileLabel.value.ext,
          sizeText: chatFile.fileLabel.value.sizeText
        }
      : null

  // 气泡仅原文；附加段只进 sendPayload（选中保留，不在发送后清空）
  const userMessage = {
    id: Date.now(),
    role: 'user',
    content: raw,
    timestamp: new Date(),
    attachment: attachmentSnap || undefined
  }
  console.log('[App] 添加用户消息:', { id: userMessage.id, content: userMessage.content.substring(0, 20), messageCount: messages.value.length })
  messages.value.push(userMessage)
  console.log('[App] 用户消息已添加，当前消息数量:', messages.value.length)

  const apiContent = isDesktopHost
    ? appendToUserContent(raw, selectedOpenFiles.value)
    : raw
  const sendPayload = { content: apiContent }
  if (uploadIdSnapshot) {
    sendPayload.uploadId = uploadIdSnapshot
  }

  // 发送到 C# 后端
  try {
    const response = await sendMessage('sendMessage', sendPayload)

    // 检查响应
    if (!response || !response.success) {
      if (response && response.requiresLogin) {
        // 需要登录，处理登录流程
        handleLoginRequired(response.userInput || content.trim())
      } else {
        // 其他错误，但不移除消息（因为可能已经收到后端消息）
        console.error('发送消息失败:', response?.message || '未知错误')
        // 不设置 isProcessing = false，等待后端的 requestStateChanged 消息
        pendingInput.value = ''
      }
    } else {
      // 消息发送成功，清除输入框
      console.log('[App] 消息发送成功，清除输入框')
      clearInput()
      pendingInput.value = ''
      // C# 与 Web 同步：首个带正文的流式 chunk 时发 uploadAttachmentContextConsumed（清空 upload_id）；首轮结束仍会兜底调用，无 upload 时不重复通知
    }
  } catch (error) {
    console.error('发送消息失败:', error)
    // 不设置 isProcessing = false，等待后端的 requestStateChanged 消息
    // isProcessing 状态应该完全由后端的 requestStateChanged 消息控制

    // 检查是否已经收到后端消息（userMessage 或 systemMessage）
    const hasReceivedBackendMessage = receivedBackendMessages.value.has(userMessage.id)

    if (!hasReceivedBackendMessage) {
      // 如果没有收到任何后端消息，说明消息确实发送失败，移除消息
      console.log('[App] 未收到后端消息，移除用户消息:', userMessage.id)
      const index = messages.value.findIndex(m => m.id === userMessage.id)
      if (index >= 0) {
        messages.value.splice(index, 1)
      }
      // 只有在确认没有收到后端消息时才设置 isProcessing = false
      isProcessing.value = false
      pendingInput.value = ''
    } else {
      console.log('[App] 已收到后端消息，不移除用户消息:', userMessage.id)
      // 已收到后端消息，说明后端正在处理，不设置 isProcessing = false
      // 等待后端的 requestStateChanged 消息来更新状态
      pendingInput.value = ''
    }
  }
}

// 处理需要登录的情况
const handleLoginRequired = async (userInput) => {
  console.log('需要登录，处理登录流程')
  
  // 1. 移除最后一条用户消息
  const lastMessage = messages.value[messages.value.length - 1]
  if (lastMessage && lastMessage.role === 'user') {
    messages.value.pop()
  }
  
  // 2. 恢复输入框内容
  if (userInput) {
    restoreInputValue(userInput)
  }
  
  // 3. 打开登录窗口
  try {
    await sendMessage('openLoginWindow', {})
  } catch (error) {
    console.error('打开登录窗口失败:', error)
  }
  
  isProcessing.value = false
  pendingInput.value = ''
}

const todoBarTodos = ref([])
const todoBarVersion = ref(null)
const todoBarVisible = ref(false)
const todoBarExpanded = ref(false)
let todoBarHideTimer = null
let todoBarLastRequestId = null
let todoBarPreviousSnapshot = null

function normalizeTodos(raw) {
  let list = raw
  if (!Array.isArray(list) && list && typeof list === 'object') {
    list = Object.keys(list)
      .filter((k) => /^\d+$/.test(k))
      .sort((a, b) => Number(a) - Number(b))
      .map((k) => list[k])
  }
  if (!Array.isArray(list)) return []
  return list.map((item) => ({
    id: String(item?.id ?? ''),
    content: String(item?.content ?? ''),
    status: String(item?.status ?? 'pending')
  }))
}

function clearTodoBarHideTimer() {
  if (todoBarHideTimer != null) {
    clearTimeout(todoBarHideTimer)
    todoBarHideTimer = null
  }
}

function resetTodoProgressBar() {
  clearTodoBarHideTimer()
  todoBarTodos.value = []
  todoBarVersion.value = null
  todoBarVisible.value = false
  todoBarExpanded.value = false
  todoBarLastRequestId = null
  todoBarPreviousSnapshot = null
}

function todosAreAllCompleted(todos) {
  return Array.isArray(todos) && todos.length > 0 && todos.every((t) => t.status === 'completed')
}

function canShowTodoBar(todos) {
  if (!Array.isArray(todos) || todos.length === 0) return false
  if (todosAreAllCompleted(todos)) return true
  if (todos.some((t) => t.status === 'in_progress')) return true
  if (todos.some((t) => t.status === 'pending')) return true
  return false
}

function applyTodoBarTodos(todos, version) {
  clearTodoBarHideTimer()
  todoBarTodos.value = todos
  if (version != null && version !== '') {
    const v = Number(version)
    if (!Number.isNaN(v)) todoBarVersion.value = v
  }

  if (!canShowTodoBar(todos)) {
    todoBarVisible.value = false
    todoBarExpanded.value = false
    return
  }

  todoBarVisible.value = true

  if (todosAreAllCompleted(todos)) {
    todoBarHideTimer = setTimeout(() => {
      todoBarHideTimer = null
      resetTodoProgressBar()
    }, 10000)
  }
}

async function handleTodoListUpdated(payload) {
  const requestId = payload?.requestId || payload?.request_id
  if (!requestId) {
    console.warn('[App] todoListUpdated missing requestId')
    return
  }

  const todos = normalizeTodos(payload?.todos)
  const versionRaw = payload?.version
  const version = versionRaw == null || versionRaw === '' ? null : Number(versionRaw)

  // 乱序：旧 version 仍 ack，但不回退 UI
  if (
    version != null &&
    !Number.isNaN(version) &&
    todoBarVersion.value != null &&
    version < todoBarVersion.value
  ) {
    try {
      await nextTick()
      await sendMessage('todoListReply', { requestId, ok: true })
    } catch (err) {
      console.error('[App] todoListReply (stale) failed:', err)
      try {
        await sendMessage('todoListReply', {
          requestId,
          ok: false,
          error: err?.message || 'todoListReply failed'
        })
      } catch (e2) {
        console.error('[App] todoListReply stale error ack failed:', e2)
      }
    }
    return
  }

  todoBarPreviousSnapshot = {
    todos: todoBarTodos.value.slice(),
    version: todoBarVersion.value,
    visible: todoBarVisible.value,
    expanded: todoBarExpanded.value
  }
  todoBarLastRequestId = requestId
  applyTodoBarTodos(todos, version)

  try {
    await nextTick()
    await sendMessage('todoListReply', { requestId, ok: true })
  } catch (err) {
    console.error('[App] todoListReply failed:', err)
    if (todoBarLastRequestId === requestId && todoBarPreviousSnapshot) {
      const prev = todoBarPreviousSnapshot
      clearTodoBarHideTimer()
      todoBarTodos.value = prev.todos
      todoBarVersion.value = prev.version
      todoBarVisible.value = prev.visible
      todoBarExpanded.value = prev.expanded
      if (prev.visible && todosAreAllCompleted(prev.todos)) {
        todoBarHideTimer = setTimeout(() => {
          todoBarHideTimer = null
          resetTodoProgressBar()
        }, 10000)
      }
    }
    try {
      await sendMessage('todoListReply', {
        requestId,
        ok: false,
        error: err?.message || 'todoListReply failed'
      })
    } catch (e2) {
      console.error('[App] todoListReply error ack failed:', e2)
    }
  }
}

function handleTodoListRevert(payload) {
  const requestId = payload?.requestId || payload?.request_id
  if (!requestId) return
  if (todoBarLastRequestId !== requestId) return
  if (!todoBarPreviousSnapshot) {
    resetTodoProgressBar()
    return
  }
  const prev = todoBarPreviousSnapshot
  clearTodoBarHideTimer()
  todoBarTodos.value = prev.todos
  todoBarVersion.value = prev.version
  todoBarVisible.value = prev.visible
  todoBarExpanded.value = prev.expanded
  todoBarLastRequestId = null
  todoBarPreviousSnapshot = null
  if (prev.visible && todosAreAllCompleted(prev.todos)) {
    todoBarHideTimer = setTimeout(() => {
      todoBarHideTimer = null
      resetTodoProgressBar()
    }, 10000)
  }
}

/** 侧栏历史 → Vue 消息；C# 已带 segments 时直接用，否则从扁平字段补 thinking */
function normalizeHistoryMessage(m) {
  if (!m || typeof m !== 'object') {
    return {
      id: Date.now(),
      role: 'system',
      content: '',
      segments: [],
      timestamp: new Date(),
      isStreaming: false
    }
  }

  const id = Number(m.id) || Date.now()
  const timestamp = m.timestamp ? new Date(m.timestamp) : new Date()
  const role = m.role || 'system'
  let content = m.content || ''
  if (role === 'user') {
    content = stripSelectedOpenFilesAppendix(content)
  }

  if (Array.isArray(m.segments) && m.segments.length > 0) {
    const segments = m.segments.map((seg) => {
      if (role === 'user' && seg?.type === 'text' && typeof seg.content === 'string') {
        return { ...seg, content: stripSelectedOpenFilesAppendix(seg.content) }
      }
      return seg
    })
    return {
      id,
      role,
      content,
      segments,
      timestamp,
      isStreaming: false,
      isHint: !!m.isHint
    }
  }

  const reasoning = m.reasoningContent || m.reasoning_content || m.reasoning || ''
  const segments = []
  if (reasoning && String(reasoning).length > 0) {
    segments.push({
      type: 'thinking',
      content: String(reasoning),
      isComplete: true
    })
  }
  if (content) {
    segments.push({ type: 'text', content })
  }

  return {
    id,
    role,
    content,
    segments,
    timestamp,
    isStreaming: false,
    isHint: !!m.isHint
  }
}

// 监听来自 C# 的消息
onMounted(() => {
  loadEmptyState()
  if (isDesktopHost) {
    refreshTaskList()
    refreshOpenFiles()
  }

  onMessage((data) => {
    console.log('收到 C# 消息:', data)
    
    if (data.type === 'userMessage') {
      // 用户消息已确认（后端已收到消息）
      // 标记已收到后端消息，这样即使 sendMessage 超时也不会删除用户消息
      const messageData = data.data || data
      // 尝试从确认消息中获取用户消息ID，如果找不到，标记所有用户消息
      if (messageData && messageData.id) {
        receivedBackendMessages.value.add(Number(messageData.id))
      } else if (data.id) {
        receivedBackendMessages.value.add(Number(data.id))
      } else {
        // 如果找不到ID，标记最后一条用户消息
        const lastUserMessage = messages.value.filter(m => m.role === 'user').pop()
        if (lastUserMessage) {
          receivedBackendMessages.value.add(lastUserMessage.id)
        }
      }
      console.log('[App] 收到用户消息确认，已标记为收到后端消息:', Array.from(receivedBackendMessages.value))
      // 不清除输入框和isProcessing状态，等待后端的requestStateChanged消息来管理
      // 输入框会在发送成功时清除，isProcessing会在后端开始处理时设置为true
      clearInput() // 清除输入框
      pendingInput.value = '' // 清空待发送消息
    } else if (data.type === 'systemMessage') {
      // 系统消息（AI 回复）
      // 消息数据可能在 data.data 中（如果 C# 发送的是 { type: "systemMessage", data: {...} }）
      // 也可能直接在 data 中（如果 C# 发送的是 { type: "systemMessage", content: ... }）
      const messageData = data.data || data

      // 确保 messageId 是数字类型（C# 发送的是 long，需要转换为 number）
      const messageId = Number(messageData.id) || Date.now()
      const messageContent = messageData.content || ''
      const reasoningContent = messageData.reasoningContent || messageData.reasoning_content || ''
      const isStreaming = messageData.isStreaming || false
      const isUpdate = messageData.isUpdate !== undefined ? messageData.isUpdate : (data.isUpdate !== undefined ? data.isUpdate : false)
      const toolCallsDelta = messageData.toolCallsDelta
      const finishReason = messageData.finishReason || messageData.finish_reason || null

      if (!toolCallAccumulators.has(messageId)) {
        toolCallAccumulators.set(messageId, new ToolCallAccumulator())
      }
      const acc = toolCallAccumulators.get(messageId)
      // thinking 按到达顺序插入时间线（多轮工具时落在当轮，不永久置顶）
      if (reasoningContent) {
        acc.feedReasoningFromCumulative(String(reasoningContent), messageContent)
      }
      if (toolCallsDelta && toolCallsDelta.length) {
        acc.feedDeltaToolCalls(toolCallsDelta)
      }
      if (finishReason) {
        if (finishReason === 'tool_calls' || finishReason === 'function_call') {
          acc.feedFinishReason(finishReason, messageContent)
        } else {
          acc.feedFinishReason(finishReason)
        }
      }
      if (!isStreaming) {
        acc.markStreamComplete()
      }
      const segments = acc.buildSegments(messageContent)
      
      // 标记已收到后端消息（系统消息说明后端已经处理了用户消息）
      // 找到最后一条用户消息并标记
      const lastUserMessage = messages.value.filter(m => m.role === 'user').pop()
      if (lastUserMessage) {
        receivedBackendMessages.value.add(lastUserMessage.id)
        console.log('[App] 收到系统消息，标记用户消息已收到后端响应:', lastUserMessage.id)
      }
      
      console.log('[App] 处理系统消息:', { 
        messageId, 
        messageContent: messageContent ? messageContent.substring(0, 20) : '(空)', 
        contentLength: messageContent ? messageContent.length : 0,
        isStreaming, 
        isUpdate
      })
      
      // 查找现有消息（确保ID类型一致，且角色是system）
      console.log('[App] 查找系统消息，ID:', messageId, '当前消息列表:', messages.value.map(m => ({ id: m.id, role: m.role, content: m.content?.substring(0, 20) })))
      const index = messages.value.findIndex(m => Number(m.id) === messageId && m.role === 'system')
      
      if (index >= 0) {
        // 更新现有消息（流式更新）
        console.log('[App] 更新现有消息，索引:', index, '消息ID:', messageId, '角色:', messages.value[index].role)
        // 只有当新内容不为空，或者现有内容为空时才更新（避免空内容覆盖非空内容）
        if (messageContent || segments.length > 0 || !messages.value[index].content) {
          messages.value[index].content = messageContent
          messages.value[index].segments = segments
          messages.value[index].isStreaming = isStreaming
          messages.value[index].timestamp = new Date(messageData.timestamp || data.timestamp || Date.now())
          console.log('[App] 消息已更新，segments:', segments.length, 'content长度:', messageContent.length)
        }
      } else {
        // 找不到现有消息，创建新消息
        console.log('[App] 创建新消息，内容长度:', messageContent.length)
        // 只有当内容不为空时才创建（避免创建空的消息气泡）
        if (messageContent || segments.length > 0) {
          const newMessage = {
            id: messageId,
            role: 'system',
            content: messageContent,
            segments,
            timestamp: new Date(messageData.timestamp || data.timestamp || Date.now()),
            isStreaming: isStreaming,
            isHint: !!messageData.isHint
          }
          messages.value.push(newMessage)
          console.log('[App] 消息已添加，当前消息数量:', messages.value.length, '添加后消息列表:', messages.value.map(m => ({ id: m.id, role: m.role, content: m.content?.substring(0, 20) })))
        } else {
          console.log('[App] 跳过空消息')
        }
      }
      
      // 定期检查消息列表，确保用户消息没有被意外删除
      console.log('[App] 当前所有消息:', messages.value.map(m => ({ id: m.id, role: m.role, contentLength: m.content?.length || 0 })))

      // 注意：isProcessing状态由requestStateChanged消息控制，不在此处设置
    } else if (data.type === 'requiresLogin') {
      // 认证过期，需要登录
      handleLoginRequired(data.userInput || pendingInput.value)
    } else if (data.type === 'requestStateChanged') {
      // 请求状态变化（开始/停止处理）
      const stateData = data.data || data
      const newState = stateData.isProcessing !== undefined ? stateData.isProcessing : (data.isProcessing !== undefined ? data.isProcessing : false)
      console.log('[App] 收到requestStateChanged消息，准备改变状态:', {
        旧状态: isProcessing.value,
        新状态: newState,
        stateData,
        data
      })
      isProcessing.value = newState
      console.log('[App] 状态已改变，当前isProcessing:', isProcessing.value)
    } else if (data.type === 'uploadAttachmentContextConsumed') {
      chatFile.reset()
    } else if (data.type === 'orchestratorMaxRoundsReached') {
      const payload = data.data || data
      const maxRounds = payload.maxRounds
      panelHint.value = typeof maxRounds === 'number' && maxRounds > 0
        ? `本轮工具调用已达上限（${maxRounds} 轮），可发送新消息继续`
        : '本轮工具调用已达上限，可发送新消息继续'
    } else if (data.type === 'panelHint') {
      const payload = data.data || data
      panelHint.value = payload.message || ''
    } else if (data.type === 'clearMessages') {
      console.log('[App] 清空消息列表')
      messages.value = []
      toolCallAccumulators.clear()
      panelHint.value = ''
      chatFile.reset()
      resetTodoProgressBar()
      loadEmptyState()
      if (pendingRestoreDraftInput != null) {
        const text = pendingRestoreDraftInput
        pendingRestoreDraftInput = null
        clearStashedDraftInput()
        nextTick(() => restoreInputBox(text))
      }
    } else if (data.type === 'todoListUpdated') {
      handleTodoListUpdated(data.data || data)
    } else if (data.type === 'openFilesUpdated') {
      const payload = data.data || data
      const items = payload?.items
      if (Array.isArray(items)) {
        openFiles.value = items
      }
    } else if (data.type === 'todoListRevert') {
      handleTodoListRevert(data.data || data)
    } else if (data.type === 'conversationHistory') {
      // 加载对话历史（C# 载荷在 data.data.messages）
      const payload = data.data || data
      const raw = payload.messages || []
      toolCallAccumulators.clear()
      panelHint.value = ''
      resetTodoProgressBar()
      messages.value = raw.map((m) => normalizeHistoryMessage(m))
      if (payload.conversationId && payload.conversationId !== '-1') {
        activeTaskId.value = payload.conversationId
        // 保留 draftSlotVisible：从新对话点进历史后，列表顶仍显示「新对话」
      }
      if (!messages.value.length) {
        loadEmptyState()
      } else {
        pinToLatestToken.value++
      }
    } else if (data.type === 'conversationIdChanged') {
      const payload = data.data || data
      if (payload?.conversationId && payload.conversationId !== '-1') {
        activeTaskId.value = payload.conversationId
        // 仅在未挂「新对话」入口时清理（例如当前草稿首次落成正式会话）
        if (!draftSlotVisible.value) {
          clearStashedDraftInput()
        }
        if (isDesktopHost) refreshTaskList()
      }
    } else if (data.type === 'conversationTitleUpdated') {
      const payload = data.data || data
      const id = payload?.conversationId ?? payload?.conversation_id
      const title = payload?.title
      if (id != null && title) {
        const item = taskList.value.find((t) => String(t.id) === String(id))
        if (item) {
          item.title = title
        } else if (isDesktopHost) {
          refreshTaskList()
        }
      }
    } else if (data.type === 'hostChatDocumentUploaded') {
      // WebView2 宿主拦截拖放后已 multipart 上传，前端从 process 继续（重复命中则跳过 process，仅 by-document）
      const payload = data.data || data
      const uuid = payload.storageDocUuid || payload.storage_doc_uuid
      const fileName = payload.fileName || payload.file_name
      const fileSize = typeof payload.fileSize === 'number' ? payload.fileSize : payload.file_size
      const skipProcess = payload.skipProcess === true || payload.skip_process === true
      if (uuid) {
        chatFile.startFromHostUploadedDocument(uuid, fileName, fileSize, skipProcess)
      }
    } else if (data.type === 'hostFileDropError') {
      const payload = data.data || data
      const msg = payload.message || '拖放上传失败'
      console.warn('[App] hostFileDropError:', msg)
      chatFile.reset()
      chatFile.phase.value = 'error'
      chatFile.errorMessage.value = msg
    } else if (data.type === 'clientToolResult') {
      const payload = data.data || data
      const toolCallId = payload.tool_call_id || payload.toolCallId
      if (toolCallId) {
        for (const msg of messages.value) {
          if (!msg.segments) continue
          const seg = msg.segments.find(
            (s) => s && s.type === 'toolCall' && s.toolCallId === toolCallId
          )
          if (seg) {
            seg.result = {
              success: payload.success,
              error: payload.error,
              data: payload.data || {}
            }
            break
          }
        }
      }
    } else if (data.type === 'layoutModeChanged') {
      const payload = data.data || data
      const mode = payload?.mode
      if (isDesktopHost) applyLayoutMode(mode)
    } else if (data.type === 'authChanged') {
      if (isDesktopHost) {
        clearApiConfigCache()
        draftSlotVisible.value = false
        clearAllDraftInput()
        refreshTaskList()
      }
    }
  })
})

onUnmounted(() => {
  clearTodoBarHideTimer()
})

// 恢复输入框内容（通过事件传递给 ChatInput）
const restoreInputValue = (value) => {
  // 触发自定义事件，让 ChatInput 组件恢复输入值
  const event = new CustomEvent('restoreInput', { detail: { value } })
  window.dispatchEvent(event)
}
</script>

<style scoped>
.app-shell {
  display: flex;
  height: 100vh;
  width: 100%;
  overflow: hidden;
  background-color: #f7f7f5;
  box-sizing: border-box;
}

/* WorkBuddy：顶栏/侧栏同色一体框；主区圆角浮层与空态同色灰底（有消息后也不变白） */
.app-shell.host-desktop {
  background: #f0f0f0;
  padding: 0 10px 10px 0;
}

.app-shell.host-desktop :deep(.task-sidebar) {
  background: #f0f0f0;
  border-right: none;
  /* 侧栏贴底，主区单独留底边距；底栏不再额外垫高 */
  margin-bottom: -10px;
  padding-bottom: 4px;
}

.chat-container {
  position: relative;
  display: flex;
  flex-direction: column;
  height: 100vh;
  width: 100%;
  min-width: 0;
  flex: 1;
  background-color: #f7f7f5;
}

.app-shell.host-desktop .chat-container {
  height: auto;
  align-self: stretch;
  min-height: 0;
  background: #f7f7f5;
  border-radius: 14px;
  overflow: hidden;
}

/* 缩小版：聊天区铺满标题栏下方（无工作台留白与圆角） */
.app-shell.host-desktop.layout-compact {
  padding: 0;
  background: #f7f7f5;
}

.app-shell.host-desktop.layout-compact .chat-container {
  border-radius: 0;
  background: #f7f7f5;
}

.compact-chrome {
  position: relative;
  flex-shrink: 0;
}

.app-shell.host-desktop :deep(.chat-messages),
.app-shell.host-desktop :deep(.chat-header),
.app-shell.host-desktop :deep(.todo-progress-card),
.app-shell.host-desktop :deep(.todo-progress-list) {
  background-color: #f7f7f5;
}

/* 输入区外层与聊天面板同色，仅内层圆角输入框为白 */
.app-shell.host-desktop :deep(.chat-input-container) {
  background-color: #f7f7f5;
}

/* 空状态与有消息时同灰底；仅推荐气泡为白（桌面端） */
.app-shell.host-desktop :deep(.chat-empty-state) {
  background-color: #f7f7f5;
  /* 与 .chat-input-container--desktop 左右 10px 对齐，推荐气泡与输入框同宽 */
  padding-left: 10px;
  padding-right: 10px;
}

.app-shell.host-desktop :deep(.chat-empty-state .recommendations) {
  margin-bottom: 4px;
  width: 100%;
  box-sizing: border-box;
}

.app-shell.host-desktop :deep(.chat-empty-state .rec-item) {
  appearance: none;
  -webkit-appearance: none;
  background-color: #ffffff !important;
  background: #ffffff !important;
  border-color: #e8e8e8;
}

.app-shell.host-desktop :deep(.chat-empty-state .rec-item:hover:not(:disabled)) {
  background-color: #ffffff !important;
  background: #ffffff !important;
  border-color: #d9d9d9;
}

.panel-hint {
  flex-shrink: 0;
  padding: 10px 16px 12px;
  margin-bottom: 4px;
  font-size: 12px;
  line-height: 1.5;
  color: #999;
  text-align: center;
}

.chat-drop-hint {
  position: absolute;
  left: 12px;
  right: 12px;
  top: 48px;
  z-index: 5;
  padding: 14px;
  text-align: center;
  font-size: 14px;
  color: #333;
  background: rgba(255, 255, 255, 0.92);
  border: 2px dashed #1890ff;
  border-radius: 10px;
  pointer-events: none;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
}
</style>

<!-- 全局隐藏滚动条（仍可滚轮/触控滚动）；不用 scoped，避免 WebView2 伪元素匹配失败 -->
<style>
html,
body,
#app {
  height: 100%;
  margin: 0;
  overflow: hidden;
}

.app-shell.host-desktop,
.app-shell.host-desktop * {
  scrollbar-width: none;
  -ms-overflow-style: none;
}

.app-shell.host-desktop *::-webkit-scrollbar {
  display: none !important;
  width: 0 !important;
  height: 0 !important;
  background: transparent !important;
}
</style>

