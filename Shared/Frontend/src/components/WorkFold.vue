<template>
  <div class="work-fold-header" :class="{ 'is-pin': pin }" @click="toggleExpand">
    <span class="chat-title">奋力工作的记录</span>
    <span v-if="durationText" class="work-fold-duration">{{ durationText }}</span>
    <span
      class="chevron"
      :class="isExpanded ? 'chevron-down' : 'chevron-right'"
      aria-hidden="true"
    />
  </div>
  <div v-if="isExpanded" class="work-fold work-fold-body">
    <slot />
  </div>
</template>

<script setup>
import { computed, ref } from 'vue'

const props = defineProps({
  /** 本轮从提问到收尾的毫秒数；过短或未知则不展示 */
  durationMs: {
    type: Number,
    default: null
  },
  /** 最新一轮：标题吸在用户气泡下面 */
  pin: {
    type: Boolean,
    default: false
  }
})

const isExpanded = ref(false)

const durationText = computed(() => formatWorkDuration(props.durationMs))

function formatWorkDuration (ms) {
  const n = Number(ms)
  if (!Number.isFinite(n) || n < 500) return ''
  const totalSec = Math.round(n / 1000)
  if (totalSec < 1) return ''
  const hours = Math.floor(totalSec / 3600)
  const minutes = Math.floor((totalSec % 3600) / 60)
  const seconds = totalSec % 60
  if (hours > 0) {
    return minutes > 0 ? `${hours}小时${minutes}分` : `${hours}小时`
  }
  if (minutes > 0) {
    return seconds > 0 ? `${minutes}分${seconds}秒` : `${minutes}分`
  }
  return `${seconds}秒`
}

const toggleExpand = () => {
  isExpanded.value = !isExpanded.value
}
</script>

<style scoped>
.work-fold-header {
  display: flex;
  align-items: center;
  gap: 8px;
  width: 100%;
  min-width: 0;
  max-width: 100%;
  box-sizing: border-box;
  padding: 8px 10px 8px 0;
  cursor: pointer;
  user-select: none;
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

.work-fold-header:hover .chevron {
  opacity: 1;
}

.chevron-right {
  transform: rotate(-45deg);
}

.chevron-down {
  transform: rotate(45deg);
}

.work-fold-duration {
  flex-shrink: 0;
  font-size: 13px;
  font-weight: 400;
  color: #999;
}

.work-fold-body {
  /* 比标题再缩一档，展开后是子级，不要和「奋力工作的记录」齐平 */
  padding: 0 10px 8px 24px;
}
</style>
