<template>
  <div class="work-fold">
    <div class="work-fold-header" @click="toggleExpand">
      <span
        class="chevron"
        :class="isExpanded ? 'chevron-down' : 'chevron-right'"
        aria-hidden="true"
      />
      <span class="work-fold-icon-spacer" aria-hidden="true" />
      <span class="work-fold-title">奋力工作</span>
    </div>
    <div v-if="isExpanded" class="work-fold-body">
      <slot />
    </div>
  </div>
</template>

<script setup>
import { ref } from 'vue'

const isExpanded = ref(false)

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
}

.chevron-right {
  transform: rotate(-45deg);
}

.chevron-down {
  transform: rotate(45deg);
}

.work-fold-icon-spacer {
  width: 16px;
  height: 16px;
  flex-shrink: 0;
}

.work-fold-title {
  flex: 1;
  font-size: 13px;
  font-weight: 500;
  color: #666;
}

.work-fold-body {
  /* 比标题再缩一档，展开后是子级，不要和「奋力工作」齐平 */
  padding: 0 10px 8px 28px;
}
</style>
