<template>
  <div class="chat-messages" ref="messagesContainer">
    <div v-for="message in messages" :key="`${message.role}-${message.id}`" class="message-wrapper">
      <UserMessageBubble v-if="message.role === 'user'" :message="message" />
      <SystemMessageBubble v-else :message="message" />
    </div>
  </div>
</template>

<script setup>
import { ref, watch, nextTick, onMounted, onUnmounted } from 'vue'
import UserMessageBubble from './UserMessageBubble.vue'
import SystemMessageBubble from './SystemMessageBubble.vue'

const props = defineProps({
  messages: {
    type: Array,
    default: () => []
  }
})

const messagesContainer = ref(null)
const shouldAutoScroll = ref(true)

// 处理滚动事件
const handleScroll = () => {
  if (!messagesContainer.value) return

  const container = messagesContainer.value
  const scrollTop = container.scrollTop
  const scrollHeight = container.scrollHeight
  const clientHeight = container.clientHeight

  // 检查是否接近底部（距离底部50px以内）
  const isNearBottom = scrollHeight - scrollTop - clientHeight < 50
  shouldAutoScroll.value = isNearBottom
}

// 自动滚动到底部
const scrollToBottom = () => {
  if (shouldAutoScroll.value) {
    nextTick(() => {
      if (messagesContainer.value) {
        messagesContainer.value.scrollTop = messagesContainer.value.scrollHeight
      }
    })
  }
}

// 组件挂载时添加滚动监听器
onMounted(() => {
  nextTick(() => {
    if (messagesContainer.value) {
      messagesContainer.value.addEventListener('scroll', handleScroll, { passive: true })
    }
  })
})

// 组件卸载时移除滚动监听器
onUnmounted(() => {
  if (messagesContainer.value) {
    messagesContainer.value.removeEventListener('scroll', handleScroll)
  }
})

watch(() => props.messages.length, (newLength, oldLength) => {
  console.log('[ChatMessages] 消息数量变化:', oldLength, '->', newLength, '消息列表:', props.messages.map(m => ({ id: m.id, role: m.role, content: m.content?.substring(0, 20) })))
  scrollToBottom()
})

watch(() => props.messages, (newMessages, oldMessages) => {
  console.log('[ChatMessages] 消息列表变化，当前消息:', newMessages.map(m => ({ id: m.id, role: m.role, contentLength: m.content?.length || 0 })))
  scrollToBottom()
}, { deep: true })
</script>

<style scoped>
.chat-messages {
  flex: 1;
  overflow-y: auto;
  padding: 16px;
  display: flex;
  flex-direction: column;
  gap: 12px;
  background-color: #f7f7f5;
  scrollbar-width: none; /* Firefox */
  -ms-overflow-style: none; /* 旧 Edge */
}

.chat-messages::-webkit-scrollbar {
  display: none; /* Chrome / WebView2 */
  width: 0;
  height: 0;
}

.message-wrapper {
  display: flex;
  flex-direction: column;
}
</style>

