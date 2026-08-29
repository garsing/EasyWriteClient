import { createApp } from 'vue'
import './utils/disableUserZoom.js'
import UserSettingsPanel from './components/UserSettingsPanel.vue'
import { onMessage } from './composables/useWebViewBridge'
import './styles/settings.css'

const app = createApp(UserSettingsPanel)

onMessage((data) => {
  console.log('[settings-main] 收到 C# 消息:', data)
})

app.mount('#settings-app')
