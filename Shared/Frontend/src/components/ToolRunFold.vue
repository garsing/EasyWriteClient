<template>
  <div class="tool-run-fold">
    <div
      class="tool-run-header"
      :class="{ 'tool-run-header--busy': busy }"
      @click="toggleExpand"
    >
      <span
        class="chevron"
        :class="isExpanded ? 'chevron-down' : 'chevron-right'"
        aria-hidden="true"
      />
      <span class="tool-run-title">{{ title }}</span>
    </div>
    <div v-if="isExpanded" class="tool-run-body">
      <template v-for="(segment, index) in segments" :key="childKey(segment, index)">
        <ThinkingBox
          v-if="segment.type === 'thinking'"
          :content="segment.content"
          :is-complete="segment.isComplete"
        />
        <ToolCallBoxDisplayRule
          v-else-if="segment.type === 'toolCall'"
          :tool-name="segment.toolName"
          :content="segment.content"
          :is-complete="segment.isComplete"
          :result="segment.result"
          :output-text="segment.outputText"
          :output-done="segment.outputDone"
          :progress-current="segment.progressCurrent"
          :progress-total="segment.progressTotal"
        />
      </template>
    </div>
  </div>
</template>

<script setup>
import { computed, ref, watch } from 'vue'
import ToolCallBoxDisplayRule from './ToolCallBoxDisplayRule.vue'
import ThinkingBox from './ThinkingBox.vue'
import { currentToolName } from '../utils/toolRunGroup'
import { getToolAlias } from '../utils/toolAlias'

const props = defineProps({
  segments: {
    type: Array,
    default: () => []
  },
  /** 当前这组还在跑：标题亮光从左扫到右 */
  busy: {
    type: Boolean,
    default: false
  }
})

const isExpanded = ref(false)
const toolAlias = ref('')

function toggleExpand () {
  isExpanded.value = !isExpanded.value
}

const toolName = computed(() => currentToolName(props.segments))

const title = computed(() => toolAlias.value || toolName.value)

watch(
  toolName,
  async (name) => {
    if (!name) {
      toolAlias.value = ''
      return
    }
    try {
      toolAlias.value = await getToolAlias(name)
    } catch {
      toolAlias.value = name
    }
  },
  { immediate: true }
)

function childKey (segment, index) {
  if (segment?.type === 'toolCall') return `tool-${segment.startIndex ?? index}`
  if (segment?.type === 'thinking') return `thinking-${segment.startIndex ?? index}`
  return `seg-${index}`
}
</script>

<style scoped>
.tool-run-fold {
  margin: 0;
}

.tool-run-header {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 10px;
  cursor: pointer;
  user-select: none;
  background-color: transparent;
  position: relative;
  overflow: hidden;
}

.tool-run-header--busy::after {
  content: '';
  position: absolute;
  top: 0;
  bottom: 0;
  left: 0;
  width: 42%;
  pointer-events: none;
  background: linear-gradient(
    90deg,
    transparent 0%,
    rgba(255, 255, 255, 0.15) 35%,
    rgba(255, 255, 255, 0.75) 50%,
    rgba(255, 255, 255, 0.15) 65%,
    transparent 100%
  );
  animation: tool-run-shine 1.7s ease-in-out infinite;
}

@keyframes tool-run-shine {
  0% {
    transform: translateX(-130%);
  }
  100% {
    transform: translateX(280%);
  }
}

@media (prefers-reduced-motion: reduce) {
  .tool-run-header--busy::after {
    animation: none;
    display: none;
  }
}

.tool-run-header:hover {
  background-color: rgba(0, 0, 0, 0.03);
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

.tool-run-title {
  flex: 1;
  font-size: 13px;
  font-weight: 500;
  color: #666;
}

.tool-run-body {
  padding: 0 10px 8px 24px;
}
</style>
