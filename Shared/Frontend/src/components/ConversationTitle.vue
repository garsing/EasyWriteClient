<template>
  <div
    class="conv-title"
    :class="['conv-title--' + size, { 'is-editing': editing }]"
    @mouseenter="hovered = true"
    @mouseleave="hovered = false"
  >
    <input
      v-if="editing"
      ref="inputEl"
      v-model="draft"
      class="conv-title-input"
      type="text"
      :aria-label="'重命名 ' + (title || '未命名对话')"
      @click.stop
      @mousedown.stop
      @keydown.enter.stop.prevent="commit"
      @keydown.esc.stop.prevent="commit"
      @blur="commit"
    />
    <template v-else>
      <span class="conv-title-text">{{ title }}</span>
      <button
        v-if="editable && hovered"
        type="button"
        class="conv-title-edit"
        title="重命名"
        aria-label="重命名"
        @click.stop="startEdit"
      >
        <img :src="renameEditIcon" alt="" />
      </button>
    </template>
  </div>
</template>

<script setup>
import { nextTick, onUnmounted, ref } from 'vue'
import renameEditIcon from '../assets/images/rename-edit.png'
import { prepareRenameTitle } from '../services/conversationsApi.js'

const props = defineProps({
  title: { type: String, default: '' },
  editable: { type: Boolean, default: false },
  conversationId: { type: [String, Number], default: null },
  size: {
    type: String,
    default: 'expanded',
    validator: (v) => v === 'expanded' || v === 'compact'
  }
})

const emit = defineEmits(['rename'])

const hovered = ref(false)
const editing = ref(false)
const draft = ref('')
const original = ref('')
const inputEl = ref(null)
let renameLock = false

function onPointerDown (e) {
  if (!editing.value) return
  if (e.target?.closest?.('.conv-title-input')) return
  commit()
}

function bindPointer () {
  window.addEventListener('mousedown', onPointerDown, true)
}

function unbindPointer () {
  window.removeEventListener('mousedown', onPointerDown, true)
}

async function startEdit () {
  if (!props.editable || props.conversationId == null) return
  original.value = props.title
  draft.value = props.title
  editing.value = true
  bindPointer()
  await nextTick()
  const el = inputEl.value
  if (!el) return
  el.focus()
  el.select()
}

function commit () {
  if (!editing.value || renameLock) return
  renameLock = true
  const id = props.conversationId
  const previous = original.value
  const title = prepareRenameTitle(draft.value, previous)
  editing.value = false
  draft.value = ''
  original.value = ''
  unbindPointer()
  emit('rename', { id, title, previous })
  renameLock = false
}

onUnmounted(() => {
  unbindPointer()
  commit()
})
</script>

<style scoped>
.conv-title {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  max-width: 100%;
  min-width: 0;
}

.conv-title.is-editing {
  display: flex;
  width: 100%;
}

.conv-title-text {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  color: #1f1e1c;
  font-weight: 700;
  line-height: 1.3;
}

.conv-title--expanded .conv-title-text {
  font-size: 16px;
}

.conv-title--compact .conv-title-text {
  font-size: 14px;
}

.conv-title-edit {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  margin: 0;
  padding: 0;
  border: none;
  background: none;
  cursor: pointer;
}

.conv-title-edit img {
  width: 16px;
  height: 16px;
  object-fit: contain;
  display: block;
}

.conv-title-input {
  width: 100%;
  min-width: 0;
  height: 28px;
  margin: 0;
  padding: 0 8px;
  border: 1px solid #3b82f6;
  border-radius: 6px;
  background: #fff;
  font: inherit;
  font-size: 16px;
  font-weight: 700;
  line-height: 26px;
  color: #1f1e1c;
  outline: none;
  box-sizing: border-box;
}

.conv-title--compact .conv-title-input {
  height: 24px;
  font-size: 14px;
  line-height: 22px;
}

.conv-title-input::selection {
  background: #3b82f6;
  color: #fff;
}
</style>
