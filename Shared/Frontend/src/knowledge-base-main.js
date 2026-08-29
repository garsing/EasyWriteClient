import { createApp } from 'vue'
import './utils/disableUserZoom.js'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import Antd from 'ant-design-vue'
import 'ant-design-vue/dist/reset.css'
import * as ElementPlusIconsVue from '@element-plus/icons-vue'
import KnowledgeBase from './components/KnowledgeBase.vue'
import { onMessage } from './composables/useWebViewBridge'

const app = createApp(KnowledgeBase)

// 注册 Element Plus 图标
for (const [key, component] of Object.entries(ElementPlusIconsVue)) {
  app.component(key, component)
}

app.use(ElementPlus)
app.use(Antd)

// 初始化消息监听器（必须在 mount 之前）
onMessage((data) => {
  console.log('[knowledge-base-main] 收到 C# 消息:', data)
  // 消息响应处理由 useWebViewBridge 内部完成
})

app.mount('#knowledge-base-app')

