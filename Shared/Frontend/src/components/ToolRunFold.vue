<template>
  <div class="tool-run-fold">
    <div class="tool-run-header" @click="toggleExpand">
      <span
        class="chevron"
        :class="isExpanded ? 'chevron-down' : 'chevron-right'"
        aria-hidden="true"
      />
      <img v-if="icon" :src="icon" class="tool-icon" alt="" />
      <span v-else class="tool-icon-spacer" aria-hidden="true" />
      <span class="tool-run-title">
        <template v-if="title">{{ title }}</template>
        <span class="tool-run-count" :class="{ 'tool-run-count--solo': !title }">已调用{{ toolCount }}个工具</span>
      </span>
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
        />
      </template>
    </div>
  </div>
</template>

<script setup>
import { computed, ref } from 'vue'
import ToolCallBoxDisplayRule from './ToolCallBoxDisplayRule.vue'
import ThinkingBox from './ThinkingBox.vue'
import { resolveRunHeading } from '../utils/toolRunGroup'
import manipulateToolIcon from '../assets/images/manipulate_tool.png'
import readToolIcon from '../assets/images/read_tool.png'

const props = defineProps({
  segments: {
    type: Array,
    default: () => []
  }
})

const isExpanded = ref(false)

function toggleExpand () {
  isExpanded.value = !isExpanded.value
}

const toolCount = computed(() => {
  return (props.segments || []).filter((s) => s && s.type === 'toolCall').length
})

const heading = computed(() => resolveRunHeading(props.segments))

const title = computed(() => {
  if (heading.value === 'thinking') return '思考'
  if (heading.value === 'explore') return '探索'
  if (heading.value === 'operate') return '操作'
  return ''
})

const icon = computed(() => {
  if (heading.value === 'explore') return readToolIcon
  if (heading.value === 'operate') return manipulateToolIcon
  return null
})

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

.tool-run-header:hover {
  background-color: rgba(0, 0, 0, 0.03);
}

.tool-icon,
.tool-icon-spacer {
  width: 16px;
  height: 16px;
  flex-shrink: 0;
}

.tool-icon {
  object-fit: contain;
  opacity: 0.75;
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

.tool-run-count {
  margin-left: 8px;
  font-weight: 400;
  color: #999;
}

.tool-run-count--solo {
  margin-left: 0;
}

.tool-run-body {
  padding: 0 10px 8px 24px;
}
</style>
