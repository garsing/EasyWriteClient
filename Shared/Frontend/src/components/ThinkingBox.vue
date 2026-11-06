<template>
  <div v-if="content" class="thinking-box">
    <div class="thinking-header" @click="toggleExpand">
      <span class="thinking-title">思考</span>
      <span
        class="chevron"
        :class="isExpanded ? 'chevron-down' : 'chevron-right'"
        aria-hidden="true"
      />
    </div>
    <div
      v-if="isExpanded"
      ref="bodyRef"
      class="thinking-body"
    >{{ content }}</div>
  </div>
</template>

<script setup>
import { ref, watch, nextTick } from 'vue'

const props = defineProps({
  content: {
    type: String,
    default: ''
  },
  isComplete: {
    type: Boolean,
    default: false
  }
})

// 与工具卡一致：进行中展开，完成后折叠
const isExpanded = ref(!props.isComplete)

watch(
  () => props.isComplete,
  (complete, wasComplete) => {
    if (complete && !wasComplete) {
      isExpanded.value = false
    } else if (!complete && wasComplete) {
      isExpanded.value = true
    }
  }
)

const bodyRef = ref(null)

watch(
  () => props.content,
  async () => {
    if (!isExpanded.value || props.isComplete) return
    await nextTick()
    const el = bodyRef.value
    if (el) {
      el.scrollTop = el.scrollHeight
    }
  }
)

const toggleExpand = () => {
  isExpanded.value = !isExpanded.value
}
</script>

<style scoped>
.thinking-box {
  margin: 4px 0;
  border: none;
  border-radius: 0;
  background-color: #f7f7f5;
  overflow: hidden;
}

.thinking-header {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 10px 8px 0;
  cursor: pointer;
  user-select: none;
  background-color: transparent;
}

.chevron {
  flex: 0 0 auto;
  width: 6px;
  height: 6px;
  box-sizing: border-box;
  border-right: 1.5px solid #999;
  border-bottom: 1.5px solid #999;
  display: inline-block;
  opacity: 0;
}

.thinking-header:hover .chevron {
  opacity: 1;
}

.chevron-right {
  transform: rotate(-45deg);
}

.chevron-down {
  transform: rotate(45deg);
}

.thinking-title {
  flex: 0 1 auto;
  font-size: 13px;
  font-weight: 500;
  color: #666;
  min-width: 0;
}

.thinking-body {
  padding: 0 10px 10px 24px;
  max-height: 120px;
  overflow-y: auto;
  overflow-x: hidden;
  margin: 0;
  font-family: 'Courier New', 'Consolas', monospace;
  font-size: 12px;
  line-height: 1.5;
  color: #666;
  white-space: pre-wrap;
  word-wrap: break-word;
  background-color: transparent;
  scrollbar-width: none;
  -ms-overflow-style: none;
}

.thinking-body::-webkit-scrollbar {
  display: none;
  width: 0;
  height: 0;
}
</style>
