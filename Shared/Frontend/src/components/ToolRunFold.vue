<template>
  <div class="tool-run-fold">
    <div
      class="tool-run-header"
      :class="{ 'tool-run-header--busy': busy }"
      @click="toggleExpand"
    >
      <span class="chat-title" :class="{ 'is-live': busy }">{{ title }}</span>
      <span v-if="progressLabel" class="tool-run-progress">{{ progressLabel }}</span>
      <span
        class="chevron"
        :class="isExpanded ? 'chevron-down' : 'chevron-right'"
        aria-hidden="true"
      />
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
import { currentRunTitle } from '../utils/toolRunGroup'
import { formatActionProgressLabel } from '../utils/toolActionProgress'
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

const runTitle = computed(() => currentRunTitle(props.segments))
const toolName = computed(() => runTitle.value.toolName)

const title = computed(() => {
  if (runTitle.value.thinking) return '思考'
  return toolAlias.value || toolName.value
})

const progressLabel = computed(() => {
  if (runTitle.value.thinking) return ''
  return formatActionProgressLabel({
    progressCurrent: runTitle.value.progressCurrent,
    progressTotal: runTitle.value.progressTotal
  })
})

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

.tool-run-header:hover .chevron {
  opacity: 1;
}

.chevron-right {
  transform: rotate(-45deg);
}

.chevron-down {
  transform: rotate(45deg);
}

.tool-run-progress {
  flex-shrink: 0;
  font-size: 12px;
  font-weight: 500;
  color: #999;
  font-variant-numeric: tabular-nums;
}

.tool-run-body {
  padding: 0 10px 8px 24px;
}
</style>
