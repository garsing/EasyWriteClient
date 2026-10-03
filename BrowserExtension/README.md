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
- **0.1.9** 起图标按钮取名：`aria-label` → `title` → `alt`（含按钮内 img / svg title / area）。须重新加载扩展。
- **0.1.10** 起不代点 `<input type=file>`；`chrome://` / PDF / 扩展页注入失败会写「此页无法附着」。须重新加载扩展。
- **0.1.11** 起点到包着 file 的 label 也拒绝（不只认 input 本身）。须重新加载扩展。
- **0.1.12** 起 type/select 对 disabled、readonly、fieldset、aria-disabled 失败（目标不可编辑）。须重新加载扩展。
- **0.1.13** 起 click 原生 `<select>` 失败（请用 action=select）；输入框不把「不能为空」当 name，并印 value=。须重新加载扩展。
- **0.1.14** 起 combobox 行带 options=（最多 8 项）；select 找不到选项时会列出可选原文。须重新加载扩展。
- **0.1.15** 起 interact 优先用宿主 css_path（避免 iframe 里本地 eN 指到壳上旧节点）；select 未落到 `<select>` 时不再暗示「请改用 select」。须重新加载扩展。
- **0.1.16** 起中间栏短文案 + `cursor:pointer` / `onclick` / `tabindex>=0` 也进树（不再只认左侧 240px）。须重新加载扩展。
- **0.1.17** 起表格每一行进树（`tr` / `role=row` / `.el-table__row`）；`role=none` 不再整表截成一行名字。须重新加载扩展。
- **0.1.19** 起商品卡保留父子：外壳只印名字不发 ref，图上的 `a` 继承卡名；点壳时改点里面带图的链接。评教那种没有 `a` 的卡仍给整卡 ref。须重新加载扩展。
- **0.1.20** 起卡内图上的 `<a>` 不再被当成「内层 pointer-card」丢掉；卡里继承来的 `cursor:pointer` 短文案不再人人发 ref。须重新加载扩展。
- **0.1.21** 起 click 默认按盒中心发指针四下；新增 `action=hover`；attach 默认同帧合成，可切 CDP（会黄条）。须重新加载扩展。
- **0.1.23** 起商品卡空盒 / 套娃 div 不发 ref；只给可见图、文字叶、有名图标和原生按钮/链发号。进详情点封面 `image`，不要点「联系客服」。须重新加载扩展。
- **0.1.24** 起 pointer 点封面图按图像素打，叠在图上的热区不当挡住。须重新加载扩展。
- **0.1.25** 起 `debugger` 改为必选权限（Chrome 不允许把它放进 optional）。重载后详情应能看到调试权；`AttachPointer=cdp` 才能真正 attach。须重新加载扩展。
- **0.1.26** 起 attach click 会把选中节点 / 命中点短名写入本机 `attach-click-diag.log`（不进工具返回给模型）。须重新加载扩展。
- **0.1.30** 起 attach CDP 第四下合成 `click` 失败不影响整次成功。须重新加载扩展。
- **0.1.31** 起点原生日期的年/月/日或日历按钮，不再把外面的 `input` 误报成挡住。须重新加载扩展。
- **0.1.32** 起原生日期用 type 写入 `YYYY-MM-DD`（今天/昨天/明天也可），不要点弹出月历格子。须重新加载扩展。
- **0.1.33** 起 type 从 label / 外壳 / `.input-group` 落到真正的 input、textarea、contenteditable；密码落到框上仍拒绝。须重新加载扩展。
- **0.1.34** 起扩展 debugger 提供通用 CDP（`rpc.cdp`），供 ClickMode=pointer 时拍 AX 树并走 CDP type/select/press。须重新加载扩展。
- **0.1.35** 起顶层文档回拍不再按子 iframe 拒绝。须重新加载扩展。
- **0.1.36** 起 press 可不带 ref（整页按键，与 scroll 相同）。须重新加载扩展。
