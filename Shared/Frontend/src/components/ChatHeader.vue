<template>
  <div class="chat-header" :class="{ 'chat-header--compact': variant === 'compact' }">
    <div class="header-buttons">
      <!-- 插件：知识库；缩小版：历史对话弹出 -->
      <button
        v-if="variant === 'plugin'"
        class="header-button"
        type="button"
        title="知识库"
        aria-label="知识库"
        @click="handleKnowledgeBase"
      >
        <img :src="knowledgeBaseIcon" alt="" class="button-icon" />
      </button>
      <button
        v-else
        class="header-button"
        type="button"
        title="历史对话"
        aria-label="历史对话"
        @click="$emit('history')"
      >
        <img :src="historyIcon" alt="" class="button-icon" />
      </button>

      <button
        class="header-button"
        type="button"
        title="新建会话"
        aria-label="新建会话"
        @click="handleAdd"
      >
        <img :src="addIcon" alt="" class="button-icon" />
      </button>
      <button
        class="header-button"
        type="button"
        title="设置"
        aria-label="设置"
        @click="handleSettings"
      >
        <img :src="userIcon" alt="" class="button-icon" />
      </button>
    </div>
  </div>
</template>

<script setup>
import { useWebViewBridge } from '../composables/useWebViewBridge'
import knowledgeBaseIcon from '../assets/images/knowledge_base.png'
import historyIcon from '../assets/images/chat-history-line.png'
import addIcon from '../assets/images/add.png'
import userIcon from '../assets/images/user.png'

const props = defineProps({
  /** plugin = 知识库/新建/设置；compact = 历史/新建/设置 */
  variant: {
    type: String,
    default: 'plugin',
    validator: (v) => v === 'plugin' || v === 'compact'
  }
})

const emit = defineEmits(['history', 'add'])

const { sendMessage } = useWebViewBridge()

const handleAdd = async () => {
  if (props.variant === 'compact') {
    emit('add')
    return
  }
  try {
    await sendMessage('addConversation', {})
  } catch (error) {
    console.error('新建会话失败:', error)
  }
}

const handleSettings = async () => {
  try {
    await sendMessage('openSettings', {})
  } catch (error) {
    console.error('打开设置失败:', error)
  }
}

const handleKnowledgeBase = async () => {
  try {
    await sendMessage('openKnowledgeBase', {})
  } catch (error) {
    console.error('打开知识库失败:', error)
  }
}
</script>

<style scoped>
.chat-header {
  padding: 8px 16px;
  background-color: #f7f7f5;
  border-bottom: 1px solid transparent;
  display: flex;
  justify-content: flex-end;
  flex-shrink: 0;
}

.header-buttons {
  display: flex;
  gap: 0px;
}

.header-button {
  background: none;
  border: none;
  padding: 2px;
  cursor: pointer;
  border-radius: 4px;
  display: flex;
  align-items: center;
  justify-content: center;
  color: #5c5a55;
}

.header-button:hover {
  background: rgba(0, 0, 0, 0.06);
}

.button-icon {
  width: 24px;
  height: 24px;
  object-fit: contain;
}

/* 缩小版顶栏：略矮；图标约等于正文（~14–16px）的视觉大小 */
.chat-header--compact {
  padding: 4px 12px;
}

.chat-header--compact .header-button {
  padding: 1px;
}

.chat-header--compact .button-icon {
  width: 20px;
  height: 20px;
}
</style>
