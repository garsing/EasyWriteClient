# EasyWriteClient

前端总目录（由 `WordAddIn1` 改名，需求 D9/D17）。

| 目录 | 说明 |
|------|------|
| `Shared/` | 公用类库：Channels + McpTools / DocumentMapping / Client / Auth / LlmClient 等（唯一工具副本） |
| `Shared/Frontend/` | **共用 Vue 前端**（Desktop + Plugin）；`npm run build` → `Shared/Frontend/wwwroot` |
| `Plugin/` | 薄 VSTO 宿主；工程文件暂仍为 `WordAddIn1.csproj`；构建时同步 `Shared/Frontend/wwwroot` |
| `Desktop/` | WinForms 桌面壳「易写」；WebView2 + `?host=desktop`；可单独 F5 |

解决方案：[`EasyWriteClient.sln`](./EasyWriteClient.sln)

## 调试

- **改 UI**：在 `Shared/Frontend` 开发，改完执行 `npm run build`
- **Desktop**：将 `EasyWriteClient.Desktop` 设为启动项目 → F5（加载 `index.html?host=desktop`：左栏任务 + 对话）  
  - 编 Desktop 时会从 `Shared/Frontend/wwwroot` 拷贝到输出目录  
- **Plugin**：将 `Plugin`（`Plugin/WordAddIn1.csproj`）设为启动项目 → VSTO 启动 Word  
  - 编 Plugin 时会把 `Shared/Frontend/wwwroot` 同步到 `Plugin/wwwroot`（缺产物时自动 `npm run build`）

> 工作区需保持完整 `EasyWriteClient` 树（含 `Shared/Frontend`）。单独只打开 Plugin 子仓无法构建前端。
