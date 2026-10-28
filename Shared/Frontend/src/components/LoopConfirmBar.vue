<template>
  <div v-if="visible" class="loop-confirm">
    <ul v-show="expanded" class="loop-confirm-steps" aria-label="重复的操作">
      <li v-for="(step, index) in steps" :key="index">{{ step }}</li>
    </ul>
    <div class="loop-confirm-row">
      <button type="button" class="loop-confirm-summary" @click="expanded = !expanded">
        <span
          class="chevron"
          :class="expanded ? 'chevron-down' : 'chevron-right'"
          aria-hidden="true"
        />
        <span class="loop-confirm-text">同一段操作已重复 {{ repeats }} 遍</span>
      </button>
      <div class="loop-confirm-actions">
        <button type="button" class="loop-confirm-action" @click="$emit('stop')">停止</button>
        <button type="button" class="loop-confirm-action" @click="$emit('ignore')">忽略</button>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, watch } from 'vue'

const props = defineProps({
  visible: {
    type: Boolean,
    default: false
  },
  repeats: {
    type: Number,
    default: 0
  },
  steps: {
    type: Array,
    default: () => []
  }
})

defineEmits(['stop', 'ignore'])

const expanded = ref(false)

watch(() => props.visible, (showing) => {
  if (showing) expanded.value = false
})
</script>

<style scoped>
.loop-confirm {
  flex-shrink: 0;
  padding: 0 8px 0 16px;
  margin-bottom: 4px;
  background: #f7f7f5;
}

.loop-confirm-steps {
  margin: 0;
  padding: 8px 10px 2px;
  list-style: none;
  max-height: min(160px, 30vh);
  overflow-y: auto;
  color: #666;
  font-size: 13px;
  line-height: 1.45;
}

.loop-confirm-steps li {
  padding: 2px 0;
  word-break: break-word;
}

.loop-confirm-row {
  display: flex;
  align-items: center;
  gap: 8px;
  min-width: 0;
}

.loop-confirm-summary {
  display: flex;
  align-items: center;
  gap: 8px;
  flex: 1;
  min-width: 0;
  margin: 0;
  padding: 8px 10px;
  border: none;
  background: transparent;
  cursor: pointer;
  text-align: left;
  color: #666;
  font-size: 13px;
  line-height: 1.4;
}

.loop-confirm-summary:hover {
  background: rgba(0, 0, 0, 0.03);
}

.loop-confirm-text {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.loop-confirm-actions {
  display: flex;
  align-items: center;
  flex-shrink: 0;
  gap: 4px;
  padding-right: 8px;
}

.loop-confirm-action {
  margin: 0;
  padding: 4px 8px;
  border: none;
  background: transparent;
  color: #666;
  font-size: 13px;
  line-height: 1.4;
  cursor: pointer;
}

.loop-confirm-action:hover {
  color: #333;
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
  margin-top: 1px;
}

.chevron-down {
  transform: rotate(45deg);
}
</style>
