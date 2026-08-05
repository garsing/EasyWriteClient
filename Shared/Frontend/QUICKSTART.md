# 快速开始指南

## 🚀 第一步：安装依赖

```bash
cd Shared/Frontend
npm install
```

## 🏗️ 第二步：构建前端

```bash
npm run build
```

构建完成后，文件会输出到 `../wwwroot` 目录。

## 🔧 第三步：切换使用 Vue 版本

在 `TaskPaneManager.cs` 中修改：

**原代码：**
```csharp
mainPanel = new TaskPaneControl();
```

**改为：**
```csharp
mainPanel = new TaskPaneControlVue();
```

## ▶️ 第四步：运行项目

在 Visual Studio 中运行 Word Add-In 项目。

## 📝 注意事项

1. **首次运行**：如果 `wwwroot` 目录不存在或为空，会显示内嵌的 HTML 提示页面
2. **WebView2 Runtime**：确保系统已安装 WebView2 Runtime
3. **调试**：查看 Visual Studio 输出窗口中的 `[TaskPaneControlVue]` 和 `[WebView2Bridge]` 日志

## 🐛 常见问题

### Q: 页面显示"正在加载 Vue 应用..."
A: 说明未找到构建后的文件。请确保：
- 已运行 `npm run build`
- `wwwroot` 目录存在于项目输出目录中
- 检查文件路径是否正确

### Q: WebView2 初始化失败
A: 请确保：
- 已安装 WebView2 Runtime
- 有足够的系统权限
- 检查用户数据文件夹路径是否正确

### Q: 消息无法发送
A: 检查：
- WebView2 是否完全加载
- 浏览器控制台是否有错误
- C# 调试输出中的错误信息

## 🔍 调试技巧

### 前端调试

1. 使用 `npm run dev` 启动开发服务器
2. 在浏览器中访问 `http://localhost:5173` 进行独立调试
3. 使用浏览器开发者工具查看控制台和网络请求

### 集成调试

1. 在 Visual Studio 中设置断点
2. 查看调试输出窗口
3. 使用 `System.Diagnostics.Debug.WriteLine` 输出日志

## 📚 下一步

- 查看 `VUE_MIGRATION_README.md` 了解详细迁移指南
- 查看 `frontend/README.md` 了解前端项目结构
- 开始迁移更多功能到 Vue 组件

