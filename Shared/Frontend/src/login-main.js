import { createApp } from 'vue'
import './styles/login.css'
import LoginPanel from './components/LoginPanel.vue'
import { onMessage } from './composables/useWebViewBridge'

onMessage((data) => {
  if (import.meta.env.DEV) {
    console.log('[login-main] 收到 C# 消息:', data)
  }
})

createApp(LoginPanel).mount('#login-app')
