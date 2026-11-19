# 2Note v1.1.0 发布说明

## 简介

2Note（原名 MD2OneNote）是一个 OneNote COM 加载项：把剪贴板里的富文本（HTML）、Markdown、图片和内联 SVG，以 **OneNote 原生结构**（one:OE / one:OEChildren / one:Table / one:Image）直接写入当前页面——绕过 OneNote 内置 HTML 导入器（它会把嵌套列表压平），完整保留嵌套结构。

## 下载

| 文件 | 适用场景 |
|---|---|
| `2Note-Setup-1.1.0.msi` | **推荐**。双击一键安装（per-user，无需管理员），自动完成 COM 注册 |
| `2Note-1.1.0-portable.zip` | 免安装版。解压后运行 `INSTALL.ps1` 手动注册，卸载运行 `UNINSTALL.ps1` |

## 安装后

1. 重启 OneNote
2. 入口三选一：
   - 功能区「视图」右侧的 **2note** 选项卡 →「Paste as better structure」
   - 页面内容区右键 →「粘贴为 OneNote 结构 (Alt+Shift+V)」
   - 全局快捷键 **Alt+Shift+V**（可在安装目录 `hotkey.txt` 或运行 `set-hotkey.ps1` 修改）
3. 从任意应用复制内容 → 在 OneNote 页面里粘贴

## 系统要求

- Windows 10/11 x64
- OneNote 2013 及以上（桌面版，含 Office16）
- .NET Framework 4.x（系统自带）
- 图片/SVG 功能需要本机有 Microsoft Edge 或 Google Chrome（用于把 SVG 光栅化为 PNG；Windows 默认自带 Edge）

## v1.1.0 更新内容

### 图片与图表
- **剪贴板图片**：PNG / CF_DIB / CF_DIBV5 位图直接嵌入 `one:Image`（含像素→点尺寸换算、DIB→PNG 转换）
- **HTML 图片**：`<img src>` 支持 data: URI、本地文件、http(s) 下载（JPEG/GIF/BMP 自动转 PNG）
- **内联 SVG / mermaid 类图表**：解析内联 `<svg>`，并从聊天应用（WorkBuddy 等）藏在 DOM 属性/script 里的 iframe 图表载荷提取 SVG；用 headless 浏览器渲染成 PNG 后按原位置嵌入
- **图注**：SVG 自带 `<title>`/`<desc>` 时自动附加一行「[图注] ...」文字，图的语义不丢；装饰图标（aria-hidden 或 ≤24px）自动过滤，渲染失败时保留 SVG 文字
- 失败逐级降级：L3 把图片替换为 `[image omitted]` 文本，不会因图片导致整页粘贴失败

### 来源信息
- 粘贴内容末尾追加「来自 {应用} · 「{窗口/页面标题}」 · {链接}」：应用取剪贴板所有者进程，标题取顶层窗口（剥离「- 应用名」后缀），链接取 CF_HTML 的 `SourceURL`
- **链接只对 http(s) 显示**（file:// 本地地址不显示）；显示文本自动解码百分号并简化（如 `%E6%8A%80%E6%9C%AF` → `技术`），href 保留原始地址可点击

### 可靠性
- **失败降级链**：整页写入被拒时自动逐级降级（去 shadingColor → 去内联样式/CDATA 标记 → 表格转文字/去图片/去列表），每级失败转储 `failed-update-L{n}.xml` 并记录原因
- **无弹窗**：降级原因写在粘贴内容末尾「[2Note 提示] ...」，不再弹出模态窗口打断用户
- **原生粘贴回退**：所有降级级别都失败时自动发 Ctrl+V，让 OneNote 内置导入器接管（插件从不写剪贴板，内容原封未动）
- **XML 良构性**：修复图片 item 的 OE 闭合记账错误（曾导致含图页面被 OneNote 整页拒绝）；回归套件加入 XML 解析器的良构性守卫
- **内容丢失修复**：`FlushCur()` 清空槽位，解决 Qwen/Chromium 用 `<div>` 装普通段落时文本被追加到已输出条目的静默丢失问题

### 格式与解析
- 中文 CF_HTML 按 UTF-8 字节偏移正确切分（信封偏移是字节不是字符）
- 段落 `<div>` / `<section>` 等块级标签自动识别；`<li>` 内 `<div>` 不拆散列表项
- `<svg>` / `<script>` / `<style>` 等非内容子树智能跳过；引号内 `>` 不破坏标签扫描

### v1.0.0 既有能力（保留）
- 原生结构写入：嵌套列表/标题/加粗/斜体/链接/行内代码/复选框/代码块，绕过 OneNote HTML 导入器
- 晚绑定 COM，无 PIA 依赖；可配置全局热键（占用检测、专用 STA 线程）
- OneMore 式 customUI（2note 选项卡 + 页面右键菜单）；per-user MSI，无需管理员

## 已知事项

- 从「仅纯文本」的复制来源粘贴时没有格式可保留（来源限制，日志会显示 `htmlFmt=0`）；点聊天应用的「复制」按钮通常只写纯文本，图不会进剪贴板——请用鼠标拖蓝选中后 Ctrl+C
- 全局热键会截获其他应用里的同组合键；冲突时用 `set-hotkey.ps1` 换一个即可
- 早期手动安装的 HKLM 注册（`C:\ProgramData\MD2OneNote`）会遮蔽新版 HKCU 注册，运行 `installer\cleanup-legacy.ps1` 一键清理
- SVG 光栅化每次启动 headless 浏览器约 3 秒（多图串行叠加）；目标站点防盗链（403）时 http 图片会跳过并记日志
- 来源行标题依赖剪贴板所有者窗口标题；不支持窗口标题的应用只显示进程名
