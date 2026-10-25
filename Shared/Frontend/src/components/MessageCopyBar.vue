<template>
  <div v-if="text" class="message-copy-bar" :class="'is-' + align">
    <button
      type="button"
      class="copy-btn"
      :title="copied ? '已复制' : '复制'"
      :aria-label="copied ? '已复制' : '复制'"
      @click.stop="copy"
    >
      <img
        :src="copyIcon"
        alt=""
        class="copy-icon"
        :class="{ 'copy-icon--copied': copied }"
        aria-hidden="true"
      />
    </button>
  </div>
</template>

<script setup>
import { onUnmounted, ref } from 'vue'
import copyIcon from '../assets/images/copy.png'

const props = defineProps({
  text: {
    type: String,
    default: ''
  },
  align: {
    type: String,
    default: 'left'
  }
})

const copied = ref(false)
let copiedTimer = 0

function writeClipboard (value) {
  if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
    return navigator.clipboard.writeText(value)
  }
  const ta = document.createElement('textarea')
  ta.value = value
  ta.setAttribute('readonly', '')
  ta.style.position = 'fixed'
  ta.style.left = '-9999px'
  document.body.appendChild(ta)
  ta.select()
  document.execCommand('copy')
  document.body.removeChild(ta)
  return Promise.resolve()
}

async function copy () {
  const value = props.text
  if (!value) return
  try {
    await writeClipboard(value)
    copied.value = true
    window.clearTimeout(copiedTimer)
    copiedTimer = window.setTimeout(() => {
      copied.value = false
    }, 1500)
  } catch (e) {
    console.warn('[MessageCopyBar] 复制失败', e?.message || e)
  }
}

onUnmounted(() => {
  window.clearTimeout(copiedTimer)
})
</script>

<style scoped>
.message-copy-bar {
  display: flex;
  align-items: center;
  height: 22px;
  margin-top: 2px;
  opacity: 0;
  pointer-events: none;
  transition: opacity 0.12s ease;
}

.message-copy-bar.is-left {
  justify-content: flex-start;
}

.message-copy-bar.is-right {
  justify-content: flex-end;
}

.copy-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 22px;
  height: 22px;
  padding: 0;
  border: none;
  border-radius: 6px;
  background: transparent;
  color: #5f5f5f;
  cursor: pointer;
}

.copy-btn:hover {
  background: rgba(0, 0, 0, 0.06);
  color: #222;
}

.copy-icon {
  width: 15px;
  height: 15px;
  display: block;
  object-fit: contain;
}

.copy-icon--copied {
  opacity: 0.55;
}
</style>
