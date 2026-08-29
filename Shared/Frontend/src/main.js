import { createApp } from 'vue'
import './utils/disableUserZoom.js'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import Antd from 'ant-design-vue'
import 'ant-design-vue/dist/reset.css'
import * as ElementPlusIconsVue from '@element-plus/icons-vue'
import App from './App.vue'

const app = createApp(App)

// 注册 Element Plus 图标
for (const [key, component] of Object.entries(ElementPlusIconsVue)) {
  app.component(key, component)
}

app.use(ElementPlus)
app.use(Antd)
app.mount('#app')

/**
 * WebView2 / Chromium：从系统拖入「文件」时，默认可能触发打开/下载。
 * 仅在 dataTransfer 含 Files 时 preventDefault，避免影响从 Word 拖选中文本等到输入框。
 */
function dataTransferHasFiles (dt) {
  if (!dt) return false
  try {
    if (dt.files && dt.files.length > 0) return true
    if (dt.items && dt.items.length) {
      for (let i = 0; i < dt.items.length; i++) {
        if (dt.items[i].kind === 'file') return true
      }
    }
    if (!dt.types) return false
    if (typeof dt.types.includes === 'function') {
      if (dt.types.includes('Files')) return true
    }
    for (let j = 0; j < dt.types.length; j++) {
      const t = String(dt.types[j]).toLowerCase()
      if (t === 'files' || t.includes('file')) return true
    }
  } catch (_) {
    /* ignore */
  }
  return false
}

function preventBrowserFileDropDefault (e) {
  if (!dataTransferHasFiles(e.dataTransfer)) return
  e.preventDefault()
  if (e.dataTransfer) {
    try {
      e.dataTransfer.dropEffect = 'copy'
    } catch (_) {
      /* ignore */
    }
  }
}

document.addEventListener('dragenter', preventBrowserFileDropDefault, true)
document.addEventListener('dragover', preventBrowserFileDropDefault, true)
document.addEventListener('drop', (e) => {
  if (!dataTransferHasFiles(e.dataTransfer)) return
  e.preventDefault()
}, true)

