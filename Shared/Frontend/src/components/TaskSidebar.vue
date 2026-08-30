<template>
  <aside class="task-sidebar" :class="{ collapsed: collapsed }">
    <div class="sidebar-top">
      <div v-if="!collapsed" class="brand">易写</div>
      <button
        type="button"
        class="icon-btn sidebar-toggle-btn"
        :title="collapsed ? '展开侧栏' : '收起侧栏'"
        :aria-label="collapsed ? '展开侧栏' : '收起侧栏'"
        @click="$emit('toggle')"
      >
        <img
          :src="sidebarToggleIcon"
          alt=""
          class="sidebar-toggle-icon"
          :class="{ collapsed: collapsed }"
          aria-hidden="true"
        />
      </button>
    </div>

    <!-- 上部：新建任务、知识库；图标与插件 ChatHeader 同一套 -->
    <button
      v-if="!collapsed"
      type="button"
      class="new-task-btn"
      @click="$emit('new-task')"
    >
      <img :src="addIcon" alt="" class="new-task-icon-img" aria-hidden="true" />
      新建任务
    </button>
    <button
      v-if="!collapsed"
      type="button"
      class="new-task-btn"
      @click="handleOpenKnowledgeBase"
    >
      <img :src="knowledgeBaseIcon" alt="" class="new-task-icon-img" aria-hidden="true" />
      知识库
    </button>
    <button
      v-if="collapsed"
      type="button"
      class="icon-btn new-task-icon"
      title="新建任务"
      aria-label="新建任务"
      @click="$emit('new-task')"
    >
      <img :src="addIcon" alt="新建任务" class="new-task-icon-img" />
    </button>
    <button
      v-if="collapsed"
      type="button"
      class="icon-btn new-task-icon"
      title="知识库"
      aria-label="知识库"
      @click="handleOpenKnowledgeBase"
    >
      <img :src="knowledgeBaseIcon" alt="知识库" class="new-task-icon-img" />
    </button>

    <!-- 新建任务与历史任务之间：打开文件（只读列表） -->
    <div v-if="!collapsed" class="open-files-section">
      <button
        type="button"
        class="section-header"
        :aria-expanded="openFilesExpanded"
        :title="openFilesExpanded ? '收起打开文件' : '展开打开文件'"
        @click="openFilesExpanded = !openFilesExpanded"
      >
        <span class="section-title">打开文件（{{ openFilesCount }}）</span>
        <span class="section-chevron" :class="{ open: openFilesExpanded }" aria-hidden="true">›</span>
      </button>

      <div class="open-files-list-clip" :class="{ open: openFilesExpanded }">
        <div class="open-files-list-clip-inner">
          <div class="open-files-list">
            <div
              v-for="group in openFileGroups"
              :key="group.key"
              class="open-file-group"
            >
              <button
                type="button"
                class="open-file-group-header"
                :aria-expanded="isAppGroupExpanded(group.key)"
                @click="toggleAppGroup(group.key)"
              >
                <span
                  class="open-file-group-chevron"
                  :class="{ open: isAppGroupExpanded(group.key) }"
                  aria-hidden="true"
                >›</span>
                <img
                  :src="group.icon"
                  alt=""
                  class="open-file-app-icon"
                  aria-hidden="true"
                />
                <span class="open-file-group-label">{{ group.label }}</span>
                <span class="open-file-group-count">{{ group.items.length }}</span>
              </button>
              <div
                v-show="isAppGroupExpanded(group.key)"
                class="open-file-group-children"
              >
                <div
                  v-for="item in group.items"
                  :key="item.id || item.displayName"
                  class="open-file-item open-file-child"
                  :class="{ selected: isOpenFileSelected(item) }"
                  @click="onOpenFileClick(item)"
                  @mouseenter="showHoverTip($event, openFileTooltip(item))"
                  @mouseleave="hideHoverTip"
                  @contextmenu.prevent="openFileContextMenu($event, item)"
                >
                  <span class="open-file-name">{{ item.displayName || '未命名文档' }}</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!--
      任务区整体高度固定；展开/折叠只改内部列表，标题栏纵向位置不变
    -->
    <div v-if="!collapsed" class="tasks-section">
      <button
        type="button"
        class="section-header"
        :aria-expanded="tasksExpanded"
        :title="tasksExpanded ? '收起历史任务' : '展开历史任务'"
        @click="tasksExpanded = !tasksExpanded"
      >
        <span class="section-title">任务</span>
        <span class="section-chevron" :class="{ open: tasksExpanded }" aria-hidden="true">›</span>
      </button>

      <div class="task-list-clip" :class="{ open: tasksExpanded }">
        <div class="task-list-clip-inner">
          <div class="task-list" @scroll.passive="onTaskListScroll">
            <div v-if="loading && !tasks.length" class="hint">加载中…</div>
            <div v-else-if="!isLoggedIn" class="hint empty">暂无对话</div>
            <div v-else-if="error && !tasks.length" class="hint error">{{ error }}</div>
            <div v-else-if="!tasks.length" class="hint empty">暂无对话</div>
            <button
              v-for="item in visibleTasks"
              :key="item.id"
              type="button"
              class="task-item"
              :class="{ active: String(item.id) === String(activeId), draft: !!item.isDraft }"
              @mouseenter="showHoverTip($event, taskTooltip(item))"
              @mouseleave="hideHoverTip"
              @click="$emit('select', item)"
            >
              <span class="task-title">{{ item.title || '未命名任务' }}</span>
              <span v-if="!item.isDraft" class="task-time">{{ formatRelativeTime(item.last_message_at || item.updated_at) }}</span>
            </button>
            <div v-if="loadingMore" class="hint">加载更多…</div>
          </div>
        </div>
      </div>
    </div>

    <!-- 底栏整条可点 → 设置（与插件 ChatHeader 同一入口） -->
    <button
      type="button"
      class="sidebar-footer"
      :class="{ collapsed: collapsed }"
      :title="footerTitle"
      :aria-label="footerTitle"
      @click="handleOpenSettings"
    >
      <img :src="userIcon" alt="" class="user-avatar-img" />
      <span v-if="!collapsed" class="user-name">{{ displayUsername }}</span>
    </button>

    <!-- 浮动完整内容提示（避开列表 overflow 裁切；WebView2 下比原生 title 可靠） -->
    <div
      v-if="hoverTip.visible && hoverTip.text"
      class="sidebar-hover-tip"
      :style="{
        top: hoverTip.top + 'px',
        left: hoverTip.left + 'px',
        maxWidth: hoverTip.maxWidth + 'px'
      }"
    >{{ hoverTip.text }}</div>

    <!-- 打开文件右键菜单：文档可「打开文件夹」；均可「复制名称」 -->
    <div
      v-if="ctxMenu.visible"
      class="open-file-ctx-menu"
      :style="{ top: ctxMenu.top + 'px', left: ctxMenu.left + 'px' }"
      @mousedown.stop
    >
      <button
        v-if="ctxMenu.canOpenFolder"
        type="button"
        class="open-file-ctx-item"
        :disabled="!ctxMenu.hasPath"
        @click="handleOpenContainingFolder"
      >
        打开文件夹
      </button>
      <button
        type="button"
        class="open-file-ctx-item"
        @click="handleCopyOpenFileName"
      >
        复制名称
      </button>
    </div>
  </aside>
</template>

<script setup>
import { computed, ref, onMounted, onUnmounted } from 'vue'
import { formatRelativeTime } from '../services/conversationsApi.js'
import { useWebViewBridge } from '../composables/useWebViewBridge'
import addIcon from '../assets/images/add.png'
import knowledgeBaseIcon from '../assets/images/knowledge_base.png'
import sidebarToggleIcon from '../assets/images/sidebar-toggle.png'
import userIcon from '../assets/images/avatar.png'
import { openFileSelectionKey } from '../utils/selectedOpenFiles.js'
import { groupOpenFilesByApp, resolveOpenFileAppGroupKey } from '../utils/openFileAppIcon.js'

const BROWSER_OPEN_FILE_GROUPS = new Set(['chrome', 'edge', 'yiwrite'])

function isBrowserOpenFile (item) {
  return BROWSER_OPEN_FILE_GROUPS.has(resolveOpenFileAppGroupKey(item))
}

const props = defineProps({
  collapsed: { type: Boolean, default: false },
  tasks: { type: Array, default: () => [] },
  openFiles: { type: Array, default: () => [] },
  selectedOpenFileIds: { type: Array, default: () => [] },
  activeId: { type: [String, Number], default: null },
  loading: { type: Boolean, default: false },
  loadingMore: { type: Boolean, default: false },
  hasMore: { type: Boolean, default: false },
  error: { type: String, default: '' }
})

const emit = defineEmits(['toggle', 'new-task', 'select', 'select-open-file', 'load-more'])

const { sendMessage, onMessage } = useWebViewBridge()

/** 历史任务列表是否展开（点「任务」标题收起/展开） */
const tasksExpanded = ref(true)
/** 打开文件列表是否展开；默认展开，不持久化 */
const openFilesExpanded = ref(true)
/** 应用分组展开态：key → boolean；缺省视为展开 */
const appGroupExpanded = ref({})

const isLoggedIn = ref(false)
const username = ref('')

const displayUsername = computed(() => {
  if (isLoggedIn.value && username.value) return username.value
  return '未登录'
})

const visibleTasks = computed(() => (isLoggedIn.value ? props.tasks : []))

const openFileGroups = computed(() => groupOpenFilesByApp(props.openFiles))

function isAppGroupExpanded (key) {
  const map = appGroupExpanded.value
  if (Object.prototype.hasOwnProperty.call(map, key)) {
    return !!map[key]
  }
  return true
}

function toggleAppGroup (key) {
  const cur = isAppGroupExpanded(key)
  appGroupExpanded.value = { ...appGroupExpanded.value, [key]: !cur }
}

const footerTitle = computed(() => {
  const name = displayUsername.value
  return name && name !== '未登录' ? `${name} · 设置` : '设置'
})

async function refreshUsername () {
  try {
    const res = await sendMessage('getCurrentUser', {})
    if (!res?.success) return
    isLoggedIn.value = !!res.isLoggedIn
    username.value = res.username || ''
  } catch (e) {
    console.warn('[TaskSidebar] getCurrentUser failed:', e?.message || e)
  }
}

const hoverTip = ref({
  visible: false,
  text: '',
  top: 0,
  left: 0,
  maxWidth: 240
})

const ctxMenu = ref({
  visible: false,
  top: 0,
  left: 0,
  path: '',
  hasPath: false,
  canOpenFolder: false,
  displayName: ''
})

function onTaskListScroll (e) {
  if (!props.hasMore || props.loadingMore || props.loading) return
  const el = e.target
  if (!el) return
  if (el.scrollHeight - el.scrollTop - el.clientHeight < 48) {
    emit('load-more')
  }
}

const openFilesCount = computed(() =>
  Array.isArray(props.openFiles) ? props.openFiles.length : 0
)

function isOpenFileSelected (item) {
  const id = openFileSelectionKey(item)
  return (props.selectedOpenFileIds || []).some((x) => String(x) === String(id))
}

function onOpenFileClick (item) {
  hideHoverTip()
  closeFileContextMenu()
  emit('select-open-file', item)
}

/** Hover 仅展示完整文件名（不显示路径） */
function openFileTooltip (item) {
  return (item?.displayName || '未命名文档').trim()
}

function taskTooltip (item) {
  return (item?.title || '未命名任务').trim()
}

function showHoverTip (event, text) {
  const content = (text || '').trim()
  if (!content) {
    hideHoverTip()
    return
  }

  const el = event.currentTarget
  if (!el || typeof el.getBoundingClientRect !== 'function') {
    return
  }

  const rect = el.getBoundingClientRect()
  const sidebar = el.closest('.task-sidebar')
  const sidebarRect = sidebar?.getBoundingClientRect()
  const left = sidebarRect ? sidebarRect.left + 8 : rect.left
  const maxWidth = Math.max(160, (sidebarRect?.width || 240) - 16)
  const gap = 6
  let top = rect.bottom + gap
  // 靠近视口底部时改到条目上方
  const estimatedHeight = Math.min(120, 24 + content.split('\n').length * 16)
  if (top + estimatedHeight > window.innerHeight - 8) {
    top = Math.max(8, rect.top - estimatedHeight - gap)
  }

  hoverTip.value = {
    visible: true,
    text: content,
    top,
    left,
    maxWidth
  }
}

function hideHoverTip () {
  hoverTip.value = {
    ...hoverTip.value,
    visible: false,
    text: ''
  }
}

function closeFileContextMenu () {
  ctxMenu.value = {
    ...ctxMenu.value,
    visible: false,
    path: '',
    hasPath: false,
    canOpenFolder: false,
    displayName: ''
  }
}

function openFileContextMenu (event, item) {
  hideHoverTip()
  const path = (item?.fullPath || item?.full_path || '').trim()
  const canOpenFolder = !isBrowserOpenFile(item)
  const displayName = (item?.displayName || '未命名文档').trim()
  const menuW = 140
  const menuH = 8 + 36 * (canOpenFolder ? 2 : 1)
  let left = event.clientX
  let top = event.clientY
  if (left + menuW > window.innerWidth - 8) {
    left = Math.max(8, window.innerWidth - menuW - 8)
  }
  if (top + menuH > window.innerHeight - 8) {
    top = Math.max(8, window.innerHeight - menuH - 8)
  }
  ctxMenu.value = {
    visible: true,
    top,
    left,
    path,
    hasPath: !!path,
    canOpenFolder,
    displayName
  }
}

function fallbackCopyText (text) {
  const ta = document.createElement('textarea')
  ta.value = text
  ta.setAttribute('readonly', '')
  ta.style.position = 'fixed'
  ta.style.left = '-9999px'
  document.body.appendChild(ta)
  ta.select()
  try {
    document.execCommand('copy')
  } finally {
    document.body.removeChild(ta)
  }
}

async function handleCopyOpenFileName () {
  const name = (ctxMenu.value.displayName || '').trim()
  closeFileContextMenu()
  if (!name) return
  try {
    if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
      await navigator.clipboard.writeText(name)
    } else {
      fallbackCopyText(name)
    }
  } catch (e) {
    try {
      fallbackCopyText(name)
    } catch (err) {
      console.warn('[TaskSidebar] 复制名称失败:', err?.message || e)
    }
  }
}

async function handleOpenContainingFolder () {
  const path = (ctxMenu.value.path || '').trim()
  closeFileContextMenu()
  if (!path) {
    console.warn('[TaskSidebar] 未保存文档无路径，无法打开文件夹')
    return
  }
  try {
    const res = await sendMessage('openContainingFolder', { path })
    if (res && res.success === false) {
      console.warn('[TaskSidebar] 打开文件夹失败:', res.message || res)
    }
  } catch (e) {
    console.error('[TaskSidebar] 打开文件夹失败:', e)
  }
}

function onGlobalPointerDown (e) {
  if (!ctxMenu.value.visible) return
  const menu = e.target?.closest?.('.open-file-ctx-menu')
  if (!menu) closeFileContextMenu()
}

function onGlobalKeyDown (e) {
  if (e.key === 'Escape') closeFileContextMenu()
}

function onAuthChanged () {
  refreshUsername()
}

let offAuthChanged = null

onMounted(() => {
  window.addEventListener('mousedown', onGlobalPointerDown, true)
  window.addEventListener('keydown', onGlobalKeyDown, true)
  window.addEventListener('blur', closeFileContextMenu)
  window.addEventListener('resize', closeFileContextMenu)
  window.addEventListener('scroll', closeFileContextMenu, true)
  offAuthChanged = onMessage((data) => {
    if (data?.type === 'authChanged') onAuthChanged()
  })
  refreshUsername()
})

onUnmounted(() => {
  window.removeEventListener('mousedown', onGlobalPointerDown, true)
  window.removeEventListener('keydown', onGlobalKeyDown, true)
  window.removeEventListener('blur', closeFileContextMenu)
  window.removeEventListener('resize', closeFileContextMenu)
  window.removeEventListener('scroll', closeFileContextMenu, true)
  if (typeof offAuthChanged === 'function') offAuthChanged()
})

async function handleOpenKnowledgeBase () {
  try {
    const res = await sendMessage('openKnowledgeBase', {})
    if (res && res.success === false) {
      console.warn('[TaskSidebar] 打开知识库失败:', res.message || res)
    }
    await refreshUsername()
  } catch (e) {
    console.error('[TaskSidebar] 打开知识库失败:', e)
  }
}

async function handleOpenSettings () {
  try {
    if (!isLoggedIn.value) {
      await sendMessage('openLoginWindow', {})
    } else {
      await sendMessage('openSettings', {})
    }
    await refreshUsername()
  } catch (e) {
    console.error('[TaskSidebar] 打开设置失败:', e)
  }
}
</script>

<style scoped>
.task-sidebar {
  width: 260px;
  flex-shrink: 0;
  display: flex;
  flex-direction: column;
  background: #f0f0f0;
  border-right: none;
  min-height: 0;
  transition: width 0.18s ease;
}

.task-sidebar.collapsed {
  width: 48px;
  align-items: center;
}

.sidebar-top {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 14px 12px 8px;
  gap: 8px;
  flex-shrink: 0;
}

.task-sidebar.collapsed .sidebar-top {
  flex-direction: column;
  padding: 12px 0;
}

.brand {
  font-family: "Segoe UI", "Microsoft YaHei UI", sans-serif;
  font-size: 18px;
  font-weight: 700;
  color: #1f1e1c;
  letter-spacing: 0.02em;
}

.icon-btn {
  border: none;
  background: transparent;
  color: #5c5a55;
  width: 32px;
  height: 32px;
  border-radius: 6px;
  cursor: pointer;
  font-size: 16px;
  line-height: 1;
}

.icon-btn:hover {
  background: rgba(0, 0, 0, 0.06);
}

.sidebar-toggle-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.sidebar-toggle-icon {
  width: 18px;
  height: 18px;
  object-fit: contain;
  display: block;
}

.sidebar-toggle-icon.collapsed {
  transform: scaleX(-1);
}

.new-task-btn {
  margin: 4px 12px 0;
  border: none;
  background: transparent;
  border-radius: 8px;
  padding: 10px 12px;
  text-align: left;
  cursor: pointer;
  font-size: 14px;
  color: #1f1e1c;
  display: flex;
  align-items: center;
  gap: 8px;
  flex-shrink: 0;
}

.new-task-btn:hover {
  background: rgba(0, 0, 0, 0.04);
}

.new-task-icon-img {
  width: 18px;
  height: 18px;
  object-fit: contain;
  flex-shrink: 0;
  display: block;
}

.new-task-icon {
  margin-top: 4px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.new-task-icon .new-task-icon-img {
  width: 22px;
  height: 22px;
}

/* 打开文件吃掉顶栏/历史任务/用户按钮之外的剩余高度 */
.open-files-section {
  flex: 1 1 auto;
  min-height: 36px;
  display: flex;
  flex-direction: column;
  padding: 8px 8px 0;
  min-width: 0;
}

/* 历史任务区约 42%；折叠只藏列表，flex-basis 不塌 */
.tasks-section {
  flex: 0 0 42%;
  max-height: 42%;
  min-height: 36px;
  display: flex;
  flex-direction: column;
  padding: 0 8px 4px;
}

.sidebar-footer {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  gap: 8px;
  width: 100%;
  margin: 0;
  border: none;
  background: transparent;
  padding: 8px 12px 8px;
  min-height: 0;
  line-height: 1;
  cursor: pointer;
  text-align: left;
  font: inherit;
  color: inherit;
  box-sizing: border-box;
  border-radius: 8px;
}

.sidebar-footer:hover {
  background: rgba(0, 0, 0, 0.04);
}

.sidebar-footer.collapsed {
  justify-content: center;
  padding: 8px 0;
  border-radius: 0;
}

.user-avatar-img {
  width: 28px;
  height: 28px;
  object-fit: contain;
  display: block;
  border-radius: 50%;
  flex-shrink: 0;
}

.user-name {
  flex: 1;
  min-width: 0;
  font-size: 14px;
  font-weight: 600;
  color: #1f1e1c;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  line-height: 1.2;
}

.section-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  width: 100%;
  margin: 0;
  border: none;
  background: transparent;
  padding: 8px 10px 6px;
  flex-shrink: 0;
  user-select: none;
  cursor: pointer;
  border-radius: 6px;
  font: inherit;
  color: inherit;
}

.section-header:hover {
  background: rgba(0, 0, 0, 0.04);
}

.section-title {
  font-size: 12px;
  font-weight: 600;
  color: #6b6963;
  letter-spacing: 0.02em;
}

.section-chevron {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 1em;
  font-size: 14px;
  font-weight: 600;
  color: #9a978f;
  line-height: 1;
  transform: rotate(0deg);
  transition: transform 0.18s ease;
}

/* 收起 ›；展开旋转 90° 朝下 */
.section-chevron.open {
  transform: rotate(90deg);
}

/* 手风琴：在固定任务区内向上收起列表，标题位置不动 */
.task-list-clip,
.open-files-list-clip {
  flex: 1 1 auto;
  min-height: 0;
  display: grid;
  grid-template-rows: 0fr;
  transition: grid-template-rows 0.22s ease;
}

.task-list-clip.open,
.open-files-list-clip.open {
  grid-template-rows: 1fr;
}

.task-list-clip-inner,
.open-files-list-clip-inner {
  overflow: hidden;
  min-height: 0;
}

.open-files-list {
  overflow-y: auto;
  max-height: 100%;
  min-height: 0;
  padding-bottom: 8px;
  scrollbar-width: none;
  -ms-overflow-style: none;
}

.open-files-list::-webkit-scrollbar {
  display: none;
  width: 0;
  height: 0;
}

.open-file-item {
  width: 100%;
  padding: 8px 10px;
  border-radius: 8px;
  margin-bottom: 1px;
  min-width: 0;
  cursor: pointer;
  user-select: none;
  display: flex;
  align-items: center;
  gap: 8px;
}

.open-file-item:hover {
  background: rgba(0, 0, 0, 0.04);
}

.open-file-item.selected {
  background: #e4e2dc;
}

.open-file-item.selected:hover {
  background: #e4e2dc;
}

.open-file-app-icon {
  width: 16px;
  height: 16px;
  object-fit: contain;
  flex-shrink: 0;
  display: block;
}

.open-file-group {
  margin-bottom: 4px;
  min-width: 0;
}

.open-file-group-header {
  display: flex;
  align-items: center;
  gap: 6px;
  width: 100%;
  margin: 0;
  border: none;
  background: transparent;
  padding: 6px 8px;
  border-radius: 8px;
  cursor: pointer;
  font: inherit;
  color: inherit;
  text-align: left;
  box-sizing: border-box;
}

.open-file-group-header:hover {
  background: rgba(0, 0, 0, 0.04);
}

.open-file-group-chevron {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 12px;
  flex-shrink: 0;
  font-size: 12px;
  font-weight: 600;
  color: #9a978f;
  line-height: 1;
  transform: rotate(0deg);
  transition: transform 0.18s ease;
}

.open-file-group-chevron.open {
  transform: rotate(90deg);
}

.open-file-group-label {
  flex: 1;
  min-width: 0;
  font-size: 12px;
  font-weight: 600;
  color: #3d3c38;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.open-file-group-count {
  flex-shrink: 0;
  font-size: 11px;
  color: #9a978f;
  font-variant-numeric: tabular-nums;
}

.open-file-group-children {
  padding-left: 10px;
  min-width: 0;
}

.open-file-item.open-file-child {
  padding: 6px 10px 6px 18px;
  gap: 0;
}

.sidebar-hover-tip {
  position: fixed;
  z-index: 1000;
  box-sizing: border-box;
  padding: 6px 8px;
  border-radius: 6px;
  background: #2c2c2c;
  color: #f5f5f5;
  font-size: 12px;
  line-height: 1.4;
  white-space: pre-wrap;
  word-break: break-all;
  pointer-events: none;
  box-shadow: 0 2px 10px rgba(0, 0, 0, 0.18);
}

.open-file-ctx-menu {
  position: fixed;
  z-index: 1100;
  min-width: 128px;
  padding: 4px;
  border-radius: 8px;
  background: #ffffff;
  border: 1px solid rgba(0, 0, 0, 0.08);
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
}

.open-file-ctx-item {
  display: block;
  width: 100%;
  border: none;
  background: transparent;
  text-align: left;
  padding: 8px 10px;
  border-radius: 6px;
  font: inherit;
  font-size: 13px;
  color: #1f1e1c;
  cursor: pointer;
}

.open-file-ctx-item:hover:not(:disabled) {
  background: rgba(0, 0, 0, 0.06);
}

.open-file-ctx-item:disabled {
  color: #9a978f;
  cursor: not-allowed;
}

.open-file-name {
  flex: 1;
  min-width: 0;
  font-size: 13px;
  color: #1f1e1c;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.task-list {
  overflow-y: auto;
  max-height: 100%;
  min-height: 0;
  padding-bottom: 12px;
  display: flex;
  flex-direction: column;
  gap: 8px;
  scrollbar-width: none;
  -ms-overflow-style: none;
}

.task-list::-webkit-scrollbar {
  display: none;
  width: 0;
  height: 0;
}

.task-item {
  width: 100%;
  border: none;
  background: transparent;
  text-align: left;
  padding: 9px 10px;
  border-radius: 8px;
  cursor: pointer;
  margin-bottom: 0;
  display: flex;
  align-items: center;
  gap: 8px;
  min-width: 0;
}

.task-item:hover {
  background: rgba(0, 0, 0, 0.04);
}

.task-item.active {
  background: #e4e2dc;
}

.task-title {
  flex: 1;
  min-width: 0;
  font-size: 13px;
  color: #1f1e1c;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.task-time {
  flex-shrink: 0;
  font-size: 11px;
  color: #9a978f;
  white-space: nowrap;
}

.hint {
  padding: 12px 10px;
  font-size: 12px;
  color: #8a877f;
}

.hint.empty {
  color: #999;
}

.hint.error {
  color: #b42318;
}
</style>
