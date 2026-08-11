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

    <!-- 上部：仅新建任务；图标与插件 ChatHeader 同一 add.png -->
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
      v-else
      type="button"
      class="icon-btn new-task-icon"
      title="新建任务"
      aria-label="新建任务"
      @click="$emit('new-task')"
    >
      <img :src="addIcon" alt="新建任务" class="new-task-icon-img" />
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
              v-for="item in openFiles"
              :key="item.id || item.displayName"
              class="open-file-item"
              :title="item.fullPath || item.displayName || ''"
            >
              <span class="open-file-name">{{ item.displayName || '未命名文档' }}</span>
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
        <span class="section-title">任务{{ taskCountLabel }}</span>
        <span class="section-chevron" :class="{ open: tasksExpanded }" aria-hidden="true">›</span>
      </button>

      <div class="task-list-clip" :class="{ open: tasksExpanded }">
        <div class="task-list-clip-inner">
          <div class="task-list">
            <div v-if="loading" class="hint">加载中…</div>
            <div v-else-if="error" class="hint error">{{ error }}</div>
            <div v-else-if="!tasks.length" class="hint">暂无任务</div>
            <button
              v-for="item in tasks"
              :key="item.id"
              type="button"
              class="task-item"
              :class="{ active: String(item.id) === String(activeId) }"
              :title="item.title || '未命名任务'"
              @click="$emit('select', item)"
            >
              <span class="task-title">{{ item.title || '未命名任务' }}</span>
              <span class="task-time">{{ formatRelativeTime(item.last_message_at || item.updated_at) }}</span>
            </button>
          </div>
        </div>
      </div>
    </div>

    <!-- 底栏整条可点 → 设置（与插件 ChatHeader 同一入口） -->
    <button
      type="button"
      class="sidebar-footer"
      :class="{ collapsed: collapsed }"
      title="设置"
      aria-label="设置"
      @click="handleOpenSettings"
    >
      <img :src="userIcon" alt="" class="user-avatar-img" />
    </button>
  </aside>
</template>

<script setup>
import { computed, ref } from 'vue'
import { formatRelativeTime } from '../services/conversationsApi.js'
import { useWebViewBridge } from '../composables/useWebViewBridge'
import addIcon from '../assets/images/add.png'
import sidebarToggleIcon from '../assets/images/sidebar-toggle.png'
import userIcon from '../assets/images/user.png'

const props = defineProps({
  collapsed: { type: Boolean, default: false },
  tasks: { type: Array, default: () => [] },
  openFiles: { type: Array, default: () => [] },
  activeId: { type: [String, Number], default: null },
  loading: { type: Boolean, default: false },
  error: { type: String, default: '' }
})

defineEmits(['toggle', 'new-task', 'select'])

const { sendMessage } = useWebViewBridge()

/** 历史任务列表是否展开（点「任务 (N)」标题收起/展开） */
const tasksExpanded = ref(true)
/** 打开文件列表是否展开；默认展开，不持久化 */
const openFilesExpanded = ref(true)

const taskCountLabel = computed(() => {
  const n = Array.isArray(props.tasks) ? props.tasks.length : 0
  return n > 0 ? ` (${n})` : ''
})

const openFilesCount = computed(() =>
  Array.isArray(props.openFiles) ? props.openFiles.length : 0
)

async function handleOpenSettings () {
  try {
    await sendMessage('openSettings', {})
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

/* 打开文件区约 28%；折叠只藏列表，flex-basis 不塌 */
.open-files-section {
  flex: 0 0 28%;
  max-height: 28%;
  min-height: 36px;
  display: flex;
  flex-direction: column;
  padding: 8px 8px 0;
  min-width: 0;
}

/* 任务区约占侧栏一半高度 */
.tasks-section {
  flex: 0 0 50%;
  max-height: 50%;
  min-height: 0;
  display: flex;
  flex-direction: column;
  padding: 0 8px 8px;
}

.sidebar-footer {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  gap: 10px;
  width: 100%;
  margin: 0;
  border: none;
  background: transparent;
  padding: 10px 12px 14px;
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
  padding: 10px 0 14px;
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
  pointer-events: none;
  user-select: none;
}

.open-file-name {
  display: block;
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
  padding: 8px 10px;
  border-radius: 8px;
  cursor: pointer;
  margin-bottom: 1px;
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

.hint.error {
  color: #b42318;
}
</style>
