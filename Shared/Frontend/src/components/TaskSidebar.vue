<template>
  <aside class="task-sidebar" :class="{ collapsed: collapsed }">
    <div class="sidebar-top">
      <div v-if="!collapsed" class="brand">易写</div>
      <button
        type="button"
        class="icon-btn"
        :title="collapsed ? '展开侧栏' : '收起侧栏'"
        :aria-label="collapsed ? '展开侧栏' : '收起侧栏'"
        @click="$emit('toggle')"
      >
        {{ collapsed ? '»' : '«' }}
      </button>
    </div>

    <!-- 上部：仅新建任务；中间留空供后续导航 -->
    <button
      v-if="!collapsed"
      type="button"
      class="new-task-btn"
      @click="$emit('new-task')"
    >
      <span class="new-task-plus" aria-hidden="true">+</span>
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
      +
    </button>

    <!-- 新建任务与历史任务之间：预留给后续导航/入口 -->
    <div v-if="!collapsed" class="sidebar-mid" aria-hidden="true" />

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
  </aside>
</template>

<script setup>
import { computed, ref } from 'vue'
import { formatRelativeTime } from '../services/conversationsApi.js'

const props = defineProps({
  collapsed: { type: Boolean, default: false },
  tasks: { type: Array, default: () => [] },
  activeId: { type: [String, Number], default: null },
  loading: { type: Boolean, default: false },
  error: { type: String, default: '' }
})

defineEmits(['toggle', 'new-task', 'select'])

/** 历史任务列表是否展开（点「任务 (N)」标题收起/展开） */
const tasksExpanded = ref(true)

const taskCountLabel = computed(() => {
  const n = Array.isArray(props.tasks) ? props.tasks.length : 0
  return n > 0 ? ` (${n})` : ''
})
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

.new-task-plus {
  display: inline-flex;
  width: 18px;
  height: 18px;
  align-items: center;
  justify-content: center;
  font-size: 16px;
  line-height: 1;
  color: #5c5a55;
}

.new-task-icon {
  margin-top: 4px;
  font-size: 20px;
  font-weight: 600;
}

/* 新建任务 ↔ 历史任务：中间留空，后续可放导航等（展开/折叠时尺寸不变） */
.sidebar-mid {
  flex: 1 1 auto;
  min-height: 96px;
}

/* 任务区占位固定，避免折叠时标题栏上下跳动 */
.tasks-section {
  flex: 1 1 42%;
  min-height: 0;
  display: flex;
  flex-direction: column;
  padding: 0 8px 12px;
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
.task-list-clip {
  flex: 1 1 auto;
  min-height: 0;
  display: grid;
  grid-template-rows: 0fr;
  transition: grid-template-rows 0.22s ease;
}

.task-list-clip.open {
  grid-template-rows: 1fr;
}

.task-list-clip-inner {
  overflow: hidden;
  min-height: 0;
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
