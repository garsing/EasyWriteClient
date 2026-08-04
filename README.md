# EasyWriteClient

前端总目录（由 `WordAddIn1` 改名，需求 D9/D17）。

| 目录 | 说明 |
|------|------|
| `Shared/` | 公用类库：Channels + McpTools / DocumentMapping / Client / Auth / LlmClient 等（唯一工具副本） |
| `Plugin/` | 薄 VSTO 宿主；工程文件暂仍为 `WordAddIn1.csproj`；前端源码仍在 `Plugin/frontend`（I9 迁 Shared 延后） |
| `Desktop/` | WinForms 桌面壳「易写」；WebView2 + `?host=desktop`；可单独 F5 |

解决方案：[`EasyWriteClient.sln`](./EasyWriteClient.sln)

## 调试

- **Desktop**：将 `EasyWriteClient.Desktop` 设为启动项目 → F5（加载 `index.html?host=desktop`：左栏任务 + 对话）  
  - 前端改完后先在 `Plugin/frontend` 执行 `npm run build`，再编 Desktop（会拷贝 `Plugin/wwwroot`）  
- **Plugin**：将 `Plugin`（`Plugin/WordAddIn1.csproj`）设为启动项目 → VSTO 启动 Word  
