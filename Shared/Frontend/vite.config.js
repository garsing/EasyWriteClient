import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { resolve } from 'path'

export default defineConfig({
  plugins: [vue()],
  base: './', // 使用相对路径，适配 file:// 协议
  build: {
    outDir: 'wwwroot',
    emptyOutDir: true,
    rollupOptions: {
      input: {
        main: resolve(__dirname, 'index.html'),
        'knowledge-base': resolve(__dirname, 'knowledge-base.html'),
        settings: resolve(__dirname, 'settings.html'),
        login: resolve(__dirname, 'login.html')
      },
      output: {
        entryFileNames: 'assets/[name].js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name].[ext]'
        // 注意：多输入时不能使用 IIFE 格式，使用默认的 ES 模块格式
      }
    },
    // 禁用代码分割（登录页单独使用 login.css，见 login-main.js）
    cssCodeSplit: true
  },
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src')
    }
  }
})

