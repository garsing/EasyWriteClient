<template>
  <div class="history-popover" role="dialog" aria-label="历史对话" @mousedown.stop>
    <div class="history-popover-head">
      <span class="history-popover-title">历史对话</span>
      <button type="button" class="history-popover-close" title="关闭" aria-label="关闭" @click="$emit('close')">×</button>
    </div>
    <div class="history-popover-body">
      <div v-if="loading" class="hint">加载中…</div>
      <div v-else-if="error" class="hint error">{{ error }}</div>
      <div v-else-if="!tasks.length" class="hint">暂无对话</div>
      <button
        v-for="item in tasks"
        :key="item.id"
        type="button"
        class="history-item"
        :class="{ active: String(item.id) === String(activeId), draft: !!item.isDraft }"
        @click="$emit('select', item)"
      >
        <span class="history-item-title">{{ item.title || '未命名对话' }}</span>
        <span v-if="!item.isDraft" class="history-item-time">{{ formatRelativeTime(item.last_message_at || item.updated_at) }}</span>
      </button>
    </div>
  </div>
</template>

<script setup>
import { formatRelativeTime } from '../services/conversationsApi.js'

defineProps({
  tasks: { type: Array, default: () => [] },
  loading: { type: Boolean, default: false },
  error: { type: String, default: '' },
  activeId: { type: [String, Number], default: null }
})

defineEmits(['select', 'close'])
</script>

<style scoped>
.history-popover {
  position: absolute;
  top: 40px;
  right: 12px;
  z-index: 40;
  width: min(320px, calc(100% - 24px));
  max-height: min(420px, 70vh);
  display: flex;
  flex-direction: column;
  background: #fff;
  border: 1px solid rgba(0, 0, 0, 0.08);
  border-radius: 10px;
  box-shadow: 0 8px 28px rgba(0, 0, 0, 0.12);
  overflow: hidden;
}

.history-popover-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 10px 12px;
  border-bottom: 1px solid rgba(0, 0, 0, 0.06);
  flex-shrink: 0;
}

.history-popover-title {
  font-size: 13px;
  font-weight: 600;
  color: #1f1e1c;
}

.history-popover-close {
  border: none;
  background: transparent;
  width: 28px;
  height: 28px;
  border-radius: 6px;
  cursor: pointer;
  font-size: 18px;
  line-height: 1;
  color: #5c5a55;
}

.history-popover-close:hover {
  background: rgba(0, 0, 0, 0.06);
}

.history-popover-body {
  overflow-y: auto;
  padding: 6px;
  min-height: 80px;
}

.hint {
  padding: 16px 10px;
  text-align: center;
  font-size: 12px;
  color: #8a8780;
}

.hint.error {
  color: #b42318;
}

.history-item {
  width: 100%;
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: 2px;
  border: none;
  background: transparent;
  text-align: left;
  padding: 10px 10px;
  border-radius: 8px;
  cursor: pointer;
}

.history-item:hover {
  background: rgba(0, 0, 0, 0.04);
}

.history-item.active {
  background: rgba(0, 0, 0, 0.06);
}

.history-item-title {
  font-size: 13px;
  color: #1f1e1c;
  line-height: 1.35;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}

.history-item-time {
  font-size: 11px;
  color: #8a8780;
}
</style>
