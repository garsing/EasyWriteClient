# 易写浏览器助手（Chrome / Edge 侧载联调）

本批先侧载，不依赖商店上架。正式上线仍须商店登记（见实现文档）。

## 前置

1. 编译解决方案（含 `YiWriteBrowserBridge`），确保 Desktop 输出目录有 `YiWriteBrowserBridge.exe`
2. 启动 **易写 Desktop**（会监听管道并尝试写 NM 登记）

## 侧载步骤

### Chrome

1. 打开 `chrome://extensions`
2. 打开「开发者模式」
3. 「加载已解压的扩展程序」→ 选本目录  
   `EasyWriteClient/BrowserExtension`
4. 复制扩展 **ID**

### Edge

1. 打开 `edge://extensions`
2. 同样加载本目录，复制扩展 ID

## 登记 Native Messaging

在 Desktop 输出目录创建 `browser-extension-id.txt`，每行一个扩展 ID（可同时写 Chrome 与 Edge）：

```text
# 示例
abcdefghijklmnopqrstuvwxyz123456
xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

或运行仓库脚本（推荐）：

```powershell
cd EasyWriteClient\BrowserExtension
.\register-sideload.ps1 -ExtensionId <你的ID> -BridgeExe <Desktop输出目录\YiWriteBrowserBridge.exe>
```

然后：

1. 完全退出并重启 Chrome/Edge（或在扩展页点「重新加载」）
2. 打开若干网页标签
3. 确认易写侧栏「打开文件」出现 Chrome/Edge 图标的网页项（标题不再带 [chrome]/[edge] 前缀）
4. 对该渠调用 `F_browser_snapshot` 等工具验收

## 日志

- Bridge：`%LOCALAPPDATA%\YiWrite\logs\browser-bridge.log`
- NM manifest：`%LOCALAPPDATA%\YiWrite\native-messaging\com.yiwrite.browser_bridge.json`

## 注意

- 关易写只断连，不杀浏览器
- 操作时会自动切到目标标签，但不会把浏览器窗口置顶
- `visible=false` 的 navigate 不会附着用户 Chrome/Edge
- **0.1.6** 起按教评 `.imgtext` / `.imgtextbtn` 识别「查看详情」（含 display:none）。须重新加载扩展。
- **0.1.7** 起普通 click 做命中测试：被 Cookie/蒙层挡住会失败；「查看详情」/`.imgtextbtn` 免检。须重新加载扩展。
- **0.1.8** 起 `scroll` 有 ref 时会滚内部 overflow 容器，不只滚整页。须重新加载扩展。
