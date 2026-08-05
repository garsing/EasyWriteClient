<template>
  <div v-if="visible" class="todo-progress-anchor">
    <div class="todo-progress-card" :class="{ expanded }">
      <ul v-show="expanded" class="todo-progress-list" aria-label="规划详情">
        <li
          v-for="item in todos"
          :key="item.id"
          class="todo-progress-item"
          :class="'status-' + (item.status || 'pending')"
        >
          <span class="todo-mark" aria-hidden="true">
            <span v-if="item.status === 'completed'" class="todo-check">✓</span>
            <span v-else-if="item.status === 'in_progress'" class="chevron chevron-right" />
          </span>
          <span class="todo-content">{{ item.content }}</span>
        </li>
      </ul>
      <button
        type="button"
        class="todo-progress-summary"
        :title="summaryTitle"
        @click="$emit('toggle')"
      >
        <span
          class="chevron"
          :class="expanded ? 'chevron-down' : 'chevron-up'"
          aria-hidden="true"
        />
        <span class="todo-summary-text">{{ summaryText }}</span>
      </button>
    </div>
  </div>
</template>

<script setup>
import { computed } from 'vue'

const props = defineProps({
  visible: {
    type: Boolean,
    default: false
  },
  expanded: {
    type: Boolean,
    default: false
  },
  todos: {
    type: Array,
    default: () => []
  }
})

defineEmits(['toggle'])

const summaryInfo = computed(() => {
  const list = Array.isArray(props.todos) ? props.todos : []
  const n = list.length
  if (!n) return null

  const allDone = list.every((t) => t.status === 'completed')
  if (allDone) {
    const text = `${n}/${n} 全部完成`
    return { text, title: text }
  }

  const ipIdx = list.findIndex((t) => t.status === 'in_progress')
  if (ipIdx >= 0) {
    const content = list[ipIdx].content || ''
    const text = `${ipIdx + 1}/${n} ${content} 进行中`
    return { text, title: text }
  }

  const pIdx = list.findIndex((t) => t.status === 'pending')
  if (pIdx >= 0) {
    const content = list[pIdx].content || ''
    const text = `${pIdx + 1}/${n} ${content} 待开始`
    return { text, title: text }
  }

  return null
})

const summaryText = computed(() => summaryInfo.value?.text || '')
const summaryTitle = computed(() => summaryInfo.value?.title || '')
</script>

<style scoped>
.todo-progress-anchor {
  flex-shrink: 0;
  padding: 0 8px 0 16px;
  margin-bottom: 4px;
}

.todo-progress-card {
  background: #f7f7f5;
  border: none;
  border-radius: 0;
  display: flex;
  flex-direction: column;
}

/* 文档流展开：与长条同在「聊天区 ↔ 输入框」之间，向上顶开消息区 */
.todo-progress-list {
  position: static;
  margin: 0;
  padding: 8px 10px 2px;
  list-style: none;
  max-height: min(320px, 50vh);
  overflow-y: auto;
  background: #f7f7f5;
  border: none;
  box-sizing: border-box;
  order: 0;
}

.todo-progress-item {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  font-size: 13px;
  line-height: 1.45;
  padding: 4px 0;
  color: #666;
}

.todo-mark {
  flex: 0 0 auto;
  width: 12px;
  height: 1.45em;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  color: #999;
}

.todo-check {
  font-size: 11px;
  line-height: 1;
  color: #999;
}

.todo-content {
  flex: 1;
  min-width: 0;
  white-space: pre-wrap;
  word-break: break-word;
}

.todo-progress-item.status-completed {
  opacity: 0.55;
}

.todo-progress-item.status-completed .todo-content {
  text-decoration: line-through;
}

.todo-progress-item.status-in_progress .todo-content {
  font-weight: 600;
  color: #555;
}

.todo-progress-summary {
  order: 1;
  display: flex;
  align-items: center;
  gap: 8px;
  width: 100%;
  margin: 0;
  padding: 8px 10px;
  border: none;
  background: transparent;
  cursor: pointer;
  text-align: left;
  color: #666;
  font-size: 13px;
  line-height: 1.4;
  flex-shrink: 0;
}

.todo-progress-summary:hover {
  background: rgba(0, 0, 0, 0.03);
}

/* 90° 直角折角（非字体 > 的锐角） */
.chevron {
  flex: 0 0 auto;
  width: 6px;
  height: 6px;
  box-sizing: border-box;
  border-right: 1.5px solid #999;
  border-bottom: 1.5px solid #999;
  display: inline-block;
}

.chevron-right {
  transform: rotate(-45deg);
  margin-top: 1px;
}

.chevron-up {
  transform: rotate(225deg);
}

.chevron-down {
  transform: rotate(45deg);
}

.todo-summary-text {
  flex: 1;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
</style>
