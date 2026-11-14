<template>
  <div class="history-popover" role="dialog" aria-label="历史对话" @mousedown.stop>
    <div class="history-popover-head">
      <span class="history-popover-title">历史对话</span>
      <button type="button" class="history-popover-close" title="关闭" aria-label="关闭" @click="$emit('close')">×</button>
    </div>
    <div class="history-popover-body" @scroll.passive="onHistoryScroll">
      <div v-if="loading && !tasks.length" class="hint">加载中…</div>
      <div v-else-if="error && !tasks.length" class="hint error">{{ error }}</div>
      <div v-else-if="!tasks.length" class="hint empty">暂无对话</div>
      <div
        v-for="item in tasks"
        :key="item.id"
        class="history-item"
        :class="{ active: String(item.id) === String(activeId), draft: !!item.isDraft, editing: isEditing(item) }"
        role="button"
        :tabindex="isEditing(item) ? -1 : 0"
        @click="onItemClick(item)"
        @keydown.enter.prevent="onItemClick(item)"
        @contextmenu.prevent="openMenu($event, item)"
      >
        <input
          v-if="isEditing(item)"
          v-model="editingTitle"
          class="history-title-input"
          type="text"
          :aria-label="'重命名 ' + (item.title || '未命名对话')"
          @click.stop
          @mousedown.stop
          @keydown.enter.stop.prevent="commitRename"
          @keydown.esc.stop.prevent="commitRename"
          @blur="commitRename"
          @contextmenu.prevent
        />
        <span v-else class="history-item-title">{{ item.title || '未命名对话' }}</span>
        <span v-if="!item.isDraft" class="history-item-time">{{ formatRelativeTime(item.last_message_at || item.updated_at) }}</span>
      </div>
      <div v-if="loadingMore" class="hint">加载更多…</div>
    </div>
  </div>
  <Teleport to="body">
    <div
      v-if="menu.visible"
      class="history-ctx-menu"
      :style="{ top: menu.top + 'px', left: menu.left + 'px' }"
      @mousedown.stop
    >
      <button type="button" class="history-ctx-item" @click="startRename">
        重命名
      </button>
    </div>
  </Teleport>
</template>

<script setup>
import { ref, nextTick, onMounted, onUnmounted } from 'vue'
import { formatRelativeTime, prepareRenameTitle } from '../services/conversationsApi.js'

const props = defineProps({
  tasks: { type: Array, default: () => [] },
  loading: { type: Boolean, default: false },
  loadingMore: { type: Boolean, default: false },
  hasMore: { type: Boolean, default: false },
  error: { type: String, default: '' },
  activeId: { type: [String, Number], default: null }
})

const emit = defineEmits(['select', 'close', 'load-more', 'rename'])

const menu = ref({
  visible: false,
  top: 0,
  left: 0,
  id: null,
  title: ''
})
const editingId = ref(null)
const editingTitle = ref('')
const editingOriginal = ref('')
let renameLock = false

function isEditing (item) {
  return editingId.value != null && String(item?.id) === String(editingId.value)
}

function closeMenu () {
  menu.value = { visible: false, top: 0, left: 0, id: null, title: '' }
}

function onItemClick (item) {
  if (isEditing(item)) return
  closeMenu()
  emit('select', item)
}

function openMenu (event, item) {
  if (item?.isDraft) return
  const menuW = 128
  const menuH = 44
  let left = event.clientX
  let top = event.clientY
  if (left + menuW > window.innerWidth - 8) {
    left = Math.max(8, window.innerWidth - menuW - 8)
  }
  if (top + menuH > window.innerHeight - 8) {
    top = Math.max(8, window.innerHeight - menuH - 8)
  }
  menu.value = {
    visible: true,
    top,
    left,
    id: item.id,
    title: item.title || ''
  }
}

async function startRename () {
  const id = menu.value.id
  const title = menu.value.title || ''
  closeMenu()
  if (id == null) return
  editingId.value = String(id)
  editingTitle.value = title
  editingOriginal.value = title
  await nextTick()
  const el = document.querySelector('.history-popover .history-title-input')
  if (!el) return
  el.focus()
  el.select()
}

function commitRename () {
  if (editingId.value == null || renameLock) return
  renameLock = true
  const id = editingId.value
  const previous = editingOriginal.value
  const title = prepareRenameTitle(editingTitle.value, previous)
  editingId.value = null
  editingTitle.value = ''
  editingOriginal.value = ''
  emit('rename', { id, title, previous })
  renameLock = false
}

function onPointerDown (e) {
  if (e.target?.closest?.('.history-ctx-menu')) return
  if (editingId.value != null && !e.target?.closest?.('.history-title-input')) {
    commitRename()
  }
  if (!e.target?.closest?.('.history-popover')) closeMenu()
}

function onKeyDown (e) {
  if (e.key !== 'Escape') return
  if (editingId.value != null) {
    e.stopPropagation()
    commitRename()
    return
  }
  closeMenu()
}

onMounted(() => {
  window.addEventListener('mousedown', onPointerDown, true)
  window.addEventListener('keydown', onKeyDown, true)
})

onUnmounted(() => {
  window.removeEventListener('mousedown', onPointerDown, true)
  window.removeEventListener('keydown', onKeyDown, true)
  commitRename()
})

function onHistoryScroll (e) {
  if (!props.hasMore || props.loadingMore || props.loading) return
  const el = e.target
  if (!el) return
  if (el.scrollHeight - el.scrollTop - el.clientHeight < 48) {
    emit('load-more')
  }
}
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

.hint.empty {
  color: #999;
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
  font: inherit;
  color: inherit;
  box-sizing: border-box;
  user-select: none;
}

.history-title-input {
  width: 100%;
  height: 24px;
  margin: 0;
  padding: 0 6px;
  border: 1px solid #3b82f6;
  border-radius: 6px;
  background: #fff;
  font: inherit;
  font-size: 13px;
  line-height: 22px;
  color: #1f1e1c;
  outline: none;
  box-sizing: border-box;
  user-select: text;
  cursor: text;
}

.history-title-input::selection {
  background: #3b82f6;
  color: #fff;
}

.history-ctx-menu {
  position: fixed;
  z-index: 1100;
  min-width: 128px;
  padding: 4px;
  border-radius: 8px;
  background: #ffffff;
  border: 1px solid rgba(0, 0, 0, 0.08);
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
}

.history-ctx-item {
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

.history-ctx-item:hover {
  background: rgba(0, 0, 0, 0.06);
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
