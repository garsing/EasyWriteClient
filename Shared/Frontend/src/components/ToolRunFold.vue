<template>
  <div class="tool-run-fold">
    <div
      class="tool-run-header"
      :class="{ 'tool-run-header--busy': busy }"
      @click="toggleExpand"
    >
      <span class="tool-run-title">{{ title }}</span>
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
}

.tool-run-header--busy .tool-run-title {
  display: inline-block;
  color: transparent;
  -webkit-text-fill-color: transparent;
  background-color: transparent;
  background-image: linear-gradient(
    90deg,
    #333 0%,
    #333 42%,
    #9a9a9a 50%,
    #333 58%,
    #333 100%
  );
  background-size: 280% 100%;
  background-clip: text;
  -webkit-background-clip: text;
  animation: tool-run-shine 3.4s linear infinite;
}

@keyframes tool-run-shine {
  0% {
    background-position: 100% 0;
  }
  100% {
    background-position: -120% 0;
  }
}

@media (prefers-reduced-motion: reduce) {
  .tool-run-header--busy .tool-run-title {
    color: #666;
    -webkit-text-fill-color: #666;
    background-image: none;
    animation: none;
  }
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

.tool-run-title {
  flex: 0 1 auto;
  font-size: 13px;
  font-weight: 500;
  color: #666;
}

.tool-run-body {
  padding: 0 10px 8px 24px;
}
</style>
