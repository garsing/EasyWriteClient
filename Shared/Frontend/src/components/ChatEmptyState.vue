<template>
  <div ref="rootEl" class="chat-empty-state">
    <div class="empty-top">
      <h2 class="empty-title">Hi，今天从哪里开始？</h2>

      <div v-if="guideBlocks.length" class="operation-guide">
        <template v-for="(block, bi) in guideBlocks" :key="bi">
          <p v-if="block.type === 'paragraph'" class="guide-paragraph">
            <template v-for="(part, pi) in block.parts" :key="pi">
              <span v-if="part.type === 'text'" class="guide-text">{{ part.value }}</span>
              <img
                v-else-if="part.type === 'icon' && iconUrls[part.filename]"
                :src="iconUrls[part.filename]"
                :alt="part.filename"
                class="guide-icon"
                draggable="false"
              />
            </template>
          </p>
          <ul v-else-if="block.type === 'list'" class="guide-list">
            <li v-for="(item, ii) in block.items" :key="ii" class="guide-list-item">
              <template v-for="(part, pi) in item" :key="pi">
                <span v-if="part.type === 'text'" class="guide-text">{{ part.value }}</span>
                <img
                  v-else-if="part.type === 'icon' && iconUrls[part.filename]"
                  :src="iconUrls[part.filename]"
                  :alt="part.filename"
                  class="guide-icon"
                  draggable="false"
                />
              </template>
            </li>
          </ul>
        </template>
      </div>
    </div>

    <div class="empty-spacer" aria-hidden="true" />

    <div v-if="recommendations.length" class="recommendations">
      <button
        v-for="(item, idx) in recommendations"
        :key="idx"
        type="button"
        class="rec-item"
        :title="item"
        :disabled="disabled"
        @click="onPick(item)"
      >
        <span class="rec-item-text">{{ item }}</span>
      </button>
    </div>
  </div>
</template>

<script setup>
import { computed, onUnmounted, ref, watch } from 'vue'
import { fetchChatEmptyStateIconUrl } from '../services/chatEmptyStateApi.js'
import { collectGuideIconFilenames, parseOperationGuide } from '../utils/parseOperationGuide.js'

const props = defineProps({
  operationGuide: {
    type: String,
    default: ''
  },
  recommendations: {
    type: Array,
    default: () => []
  },
  disabled: {
    type: Boolean,
    default: false
  }
})

const emit = defineEmits(['select-recommendation'])

const iconUrls = ref({})
/** @type {string[]} */
let blobUrls = []

const guideBlocks = computed(() => parseOperationGuide(props.operationGuide))
const rootEl = ref(null)

async function loadIcons (blocks) {
  blobUrls.forEach((u) => {
    try {
      URL.revokeObjectURL(u)
    } catch (_) {
      /* ignore */
    }
  })
  blobUrls = []
  iconUrls.value = {}

  const names = collectGuideIconFilenames(blocks)
  const next = {}
  await Promise.all(
    names.map(async (name) => {
      try {
        const url = await fetchChatEmptyStateIconUrl(name)
        blobUrls.push(url)
        next[name] = url
      } catch (e) {
        console.warn('[ChatEmptyState] 缺图跳过:', name, e?.message || e)
      }
    })
  )
  iconUrls.value = next
  // 内容变化后仍从主标题处起滚
  requestAnimationFrame(() => {
    if (rootEl.value) rootEl.value.scrollTop = 0
  })
}

watch(
  guideBlocks,
  (blocks) => {
    loadIcons(blocks)
  },
  { immediate: true }
)

onUnmounted(() => {
  blobUrls.forEach((u) => {
    try {
      URL.revokeObjectURL(u)
    } catch (_) {
      /* ignore */
    }
  })
  blobUrls = []
})

function onPick (text) {
  if (props.disabled || !text) return
  emit('select-recommendation', text)
}
</script>

<style scoped>
.chat-empty-state {
  flex: 1;
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
  align-items: stretch;
  padding: 24px 16px 12px;
  min-height: 0;
  overflow-y: auto;
  background-color: #f7f7f5;
  scrollbar-width: none; /* Firefox */
  -ms-overflow-style: none; /* 旧 Edge */
}

.chat-empty-state::-webkit-scrollbar {
  display: none; /* Chrome / WebView2 */
  width: 0;
  height: 0;
}

.empty-title {
  margin: 0 0 16px;
  padding: 0;
  font-size: 22px;
  font-weight: 600;
  line-height: 1.35;
  color: #1a1a1a;
  text-align: left;
  letter-spacing: 0.02em;
}

.empty-top {
  flex-shrink: 0;
}

.empty-spacer {
  flex: 1 1 auto;
  min-height: 16px;
}

.operation-guide {
  margin-bottom: 0;
  font-size: 14px;
  font-weight: 400;
  line-height: 1.7;
  color: #444;
  word-break: break-word;
}

.guide-paragraph {
  margin: 0 0 8px;
}

.guide-paragraph:last-child {
  margin-bottom: 0;
}

.guide-list {
  margin: 0 0 8px;
  padding-left: 1.25em;
  list-style-type: disc;
}

.guide-list:last-child {
  margin-bottom: 0;
}

.guide-list-item {
  margin: 0 0 6px;
  padding-left: 0.15em;
}

.guide-list-item:last-child {
  margin-bottom: 0;
}

.guide-text {
  vertical-align: middle;
}

.guide-icon {
  display: inline-block;
  max-height: 22px;
  width: auto;
  vertical-align: middle;
  margin: 0 2px;
  pointer-events: none;
  user-select: none;
}

.recommendations {
  flex-shrink: 0;
  display: flex;
  flex-direction: column;
  gap: 6px;
  align-items: stretch;
  margin-top: 8px;
}

.rec-item {
  display: block;
  width: 100%;
  box-sizing: border-box;
  text-align: left;
  padding: 8px 12px;
  border-radius: 12px;
  background: #fff;
  /* 与指南统一；button 默认常不继承 body 字体 */
  font-family: inherit;
  font-size: 14px;
  font-weight: 400;
  line-height: 1.45;
  color: #444;
  cursor: pointer;
  border: 1px solid #ebebeb;
}

.rec-item-text {
  display: -webkit-box;
  -webkit-box-orient: vertical;
  -webkit-line-clamp: 2;
  line-clamp: 2;
  overflow: hidden;
  text-overflow: ellipsis;
  word-break: break-word;
  /* 固定约两行高，避免截断时高度抖动 */
  max-height: calc(1.45em * 2);
}

.rec-item:hover:not(:disabled) {
  background: #fafafa;
  border-color: #ddd;
}

.rec-item:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}
</style>
