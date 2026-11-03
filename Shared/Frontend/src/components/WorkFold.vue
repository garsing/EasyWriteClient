<template>
  <div class="work-fold">
    <div class="work-fold-header" @click="toggleExpand">
      <span class="work-fold-title">
        奋力工作的记录
        <span v-if="durationText" class="work-fold-duration">{{ durationText }}</span>
      </span>
      <span
        class="chevron"
        :class="isExpanded ? 'chevron-down' : 'chevron-right'"
        aria-hidden="true"
      />
    </div>
    <div v-if="isExpanded" class="work-fold-body">
      <slot />
    </div>
  </div>
</template>

<script setup>
import { computed, ref } from 'vue'

const props = defineProps({
  /** 本轮从提问到收尾的毫秒数；过短或未知则不展示 */
  durationMs: {
    type: Number,
    default: null
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
.work-fold {
  margin: 0;
}

.work-fold-header {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 10px;
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

.work-fold-title {
  flex: 0 1 auto;
  font-size: 13px;
  font-weight: 500;
  color: #666;
}

.work-fold-duration {
  margin-left: 8px;
  font-weight: 400;
  color: #999;
}

.work-fold-body {
  /* 比标题再缩一档，展开后是子级，不要和「奋力工作的记录」齐平 */
  padding: 0 10px 8px 24px;
}
</style>
