<template>
  <div class="conv-search">
    <button
      v-if="!open"
      type="button"
      class="conv-search-icon-btn"
      title="搜索"
      aria-label="搜索"
      @click="$emit('open')"
    >
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <circle cx="11" cy="11" r="6.25" fill="none" stroke="currentColor" stroke-width="1.8" />
        <path d="M16 16.5 L20 20.5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" />
      </svg>
    </button>
    <div v-else class="conv-search-box">
      <svg class="conv-search-glyph" viewBox="0 0 24 24" aria-hidden="true">
        <circle cx="11" cy="11" r="6.25" fill="none" stroke="currentColor" stroke-width="1.8" />
        <path d="M16 16.5 L20 20.5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" />
      </svg>
      <input
        ref="inputEl"
        class="conv-search-input"
        type="text"
        :value="query"
        placeholder="搜索对话内容"
        aria-label="搜索对话内容"
        @input="$emit('query', $event.target.value)"
        @keydown.enter.prevent="$emit('enter')"
        @keydown.esc.stop.prevent="$emit('close')"
      />
      <span v-if="showCount" class="conv-search-count">{{ countText }}</span>
      <button
        type="button"
        class="conv-search-nav"
        title="上一处"
        aria-label="上一处"
        :disabled="total < 1"
        @click="$emit('prev')"
      >
        <svg viewBox="0 0 16 16" aria-hidden="true"><path d="M4 10 L8 6 L12 10" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <button
        type="button"
        class="conv-search-nav"
        title="下一处"
        aria-label="下一处"
        :disabled="total < 1"
        @click="$emit('next')"
      >
        <svg viewBox="0 0 16 16" aria-hidden="true"><path d="M4 6 L8 10 L12 6" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" /></svg>
      </button>
      <button
        type="button"
        class="conv-search-nav conv-search-close"
        title="关闭"
        aria-label="关闭"
        @click="$emit('close')"
      >
        <svg viewBox="0 0 16 16" aria-hidden="true"><path d="M4.5 4.5 L11.5 11.5 M11.5 4.5 L4.5 11.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" /></svg>
      </button>
    </div>
  </div>
</template>

<script setup>
import { computed, nextTick, ref, watch } from 'vue'

const props = defineProps({
  open: { type: Boolean, default: false },
  query: { type: String, default: '' },
  index: { type: Number, default: 0 },
  total: { type: Number, default: 0 }
})

defineEmits(['open', 'close', 'query', 'enter', 'prev', 'next'])

const inputEl = ref(null)

const showCount = computed(() => String(props.query || '').trim().length > 0)

const countText = computed(() => {
  if (props.total < 1) return '0/0'
  return `${props.index + 1}/${props.total}`
})

function focus () {
  const el = inputEl.value
  if (!el) return
  el.focus()
  el.select()
}

watch(() => props.open, async (open) => {
  if (!open) return
  await nextTick()
  focus()
})

defineExpose({ focus })
</script>

<style scoped>
.conv-search {
  flex-shrink: 0;
  margin-left: auto;
}

.conv-search-icon-btn,
.conv-search-nav {
  border: none;
  background: none;
  padding: 0;
  color: #6b6963;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
}

.conv-search-icon-btn {
  width: 28px;
  height: 28px;
  border-radius: 6px;
}

.conv-search-icon-btn:hover,
.conv-search-nav:hover:not(:disabled) {
  background: rgba(0, 0, 0, 0.06);
}

.conv-search-icon-btn svg {
  width: 18px;
  height: 18px;
}

.conv-search-box {
  display: flex;
  align-items: center;
  gap: 2px;
  height: 32px;
  padding: 0 6px 0 10px;
  border-radius: 16px;
  background: #f3f3f1;
  box-sizing: border-box;
  max-width: 360px;
}

.conv-search-glyph {
  width: 16px;
  height: 16px;
  flex-shrink: 0;
  color: #8a8884;
}

.conv-search-input {
  width: 140px;
  min-width: 0;
  border: none;
  outline: none;
  background: transparent;
  font: inherit;
  font-size: 13px;
  color: #1f1e1c;
  padding: 0 6px;
}

.conv-search-input::placeholder {
  color: #a3a29e;
}

.conv-search-count {
  flex-shrink: 0;
  margin-right: 2px;
  font-size: 12px;
  color: #8a8884;
  font-variant-numeric: tabular-nums;
}

.conv-search-nav {
  width: 22px;
  height: 22px;
  border-radius: 4px;
  flex-shrink: 0;
}

.conv-search-nav svg {
  width: 14px;
  height: 14px;
}

.conv-search-nav:disabled {
  color: #c8c6c2;
  cursor: default;
}
</style>
