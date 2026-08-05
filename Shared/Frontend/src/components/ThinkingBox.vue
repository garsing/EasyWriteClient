<template>
  <div v-if="content" class="thinking-box">
    <div class="thinking-box-header" @click="toggleExpand">
      <span
        class="chevron"
        :class="isExpanded ? 'chevron-down' : 'chevron-right'"
        aria-hidden="true"
      />
      <!-- 与 ToolCallBox 的 tool-icon 同宽，保证标题与工具名左对齐 -->
      <span class="thinking-icon-spacer" aria-hidden="true" />
      <span class="thinking-title">思考中</span>
    </div>
    <div
      v-if="isExpanded"
      ref="bodyRef"
      class="thinking-box-body"
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

// 进行中默认展开（让用户看见在吐字）；完成后默认折叠
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

// 流式进行中：固定高度区域内自动滚到底，呈现「一直在出字」
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
/* 外壳与 ToolCallBox 对齐，避免左右错位 */
.thinking-box {
  margin: 4px 0;
  border: none;
  border-radius: 0;
  background-color: #f7f7f5;
  overflow: hidden;
}

.thinking-box-header {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 10px;
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
}

.chevron-right {
  transform: rotate(-45deg);
}

.chevron-down {
  transform: rotate(45deg);
}

.thinking-icon-spacer {
  width: 16px;
  height: 16px;
  flex-shrink: 0;
}

.thinking-title {
  flex: 1;
  font-size: 13px;
  font-weight: 500;
  color: #999;
}

.thinking-box-body {
  /* 与 tool-call-content 左缩进一致；固定高度 + 滚动（滚动条隐藏，仍可滚轮/触控滚动） */
  padding: 0 10px 10px 24px;
  max-height: 120px;
  overflow-y: auto;
  overflow-x: hidden;
  font-size: 13px;
  line-height: 1.6;
  color: #999;
  word-wrap: break-word;
  white-space: pre-wrap;
  background-color: transparent;
  scrollbar-width: none; /* Firefox */
  -ms-overflow-style: none; /* 旧 Edge */
}

.thinking-box-body::-webkit-scrollbar {
  display: none; /* Chrome / WebView2 */
  width: 0;
  height: 0;
}
</style>
