<template>
  <div class="chat-header">
    <div class="header-buttons">
      <button class="header-button" @click="handleKnowledgeBase">
        <img :src="knowledgeBaseIcon" alt="知识库" class="button-icon" />
      </button>
      <button
        class="header-button"
        type="button"
        title="新建会话"
        aria-label="新建会话"
        @click="handleAdd"
      >
        <img :src="addIcon" alt="新建会话" class="button-icon" />
      </button>
      <button class="header-button" @click="handleSettings">
        <img :src="userIcon" alt="设置" class="button-icon" />
      </button>
    </div>
  </div>
</template>

<script setup>
import { useWebViewBridge } from '../composables/useWebViewBridge'
import knowledgeBaseIcon from '../assets/images/knowledge_base.png'
import addIcon from '../assets/images/add.png'
import userIcon from '../assets/images/user.png'

const { sendMessage } = useWebViewBridge()

const handleAdd = async () => {
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
}

.button-icon {
  width: 24px;
  height: 24px;
  object-fit: contain;
}
</style>
