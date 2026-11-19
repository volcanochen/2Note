# 2Note — 「粘贴为 OneNote 结构」插件（原名 MD2OneNote）

把剪贴板里的富文本（HTML）或 Markdown，**以 OneNote 原生结构**（`one:OE` / `one:OEChildren`）直接写入当前页面，完整保留嵌套列表缩进、标题、加粗/斜体、链接、行内代码、复选框等格式——**绕过 OneNote 内置 HTML 导入器**（它会把嵌套列表压平），这是本插件与一切「先转 HTML 再 Ctrl+V」方案的本质区别。

## 功能一览

| 入口 | 说明 |
|---|---|
| **2note 功能区选项卡** | 「视图」选项卡右侧的 2note 选项卡，内含大按钮 |
| **页面右键菜单** | 内容区右键 →「粘贴为 OneNote 结构」 |
| **全局快捷键** | 默认 **Alt+Shift+V**（安装包已配置）；可改为任意组合键，见下文「快捷键配置」 |

**支持的内容**：
- 多级嵌套列表（ul/ol 真嵌套，以及靠缩进模拟的伪嵌套自动推导层级）
- 有序列表字面编号（同层 ul/ol 混排或空行分隔时自动重置计数）
- **表格**：HTML `<table>` 与 Markdown 管道表格 → OneNote 原生表格（`one:Table`，带边框，表头行加粗）
- 标题 h1–h6、加粗/斜体、链接、行内代码
- 复选框 → ☐/☑
- 代码块 → 逐行 Consolas 等宽段落
- 中文正常（CF_HTML 按 UTF-8 字节读取）
- 纯文本/Markdown 源码兜底（剪贴板无 HTML 格式时按 Markdown 解析）
- **图片**（v1.1.0）：剪贴板 PNG / CF_DIB 位图直接嵌入；HTML `<img src>` 支持 data:/本地文件/http(s) 三种来源
- **内联 SVG / 图表卡片**（v1.1.0）：内联 `<svg>` 与聊天应用（如 WorkBuddy）藏在脚本里的 SVG 附件都能识别，用本机 Edge/Chrome headless 光栅化成 `one:Image` 嵌入；SVG 自带 `<title>`/`<desc>` 时自动附一行「[图注]」保留图的文字语义。装饰性小图标（aria-hidden 或 ≤24px）自动忽略
- **来源行**（v1.1.0）：粘贴内容末尾自动追加「来自 {应用} · 「{窗口/页面标题}」 · {链接}」；链接只在 http(s) 时显示（显示文本自动解码成可读形式，href 保留原始地址）

**插入位置**：新建独立文本容器（Outline），放在光标所在容器的正下方，多次插入自动下推、不重叠。

**失败与降级**（v1.1.0）：OneNote 拒绝整页写入时自动逐级降级（去样式 → 去内联标记 → 表格转文字/去图片），降级原因以「[2Note 提示]」形式写在内容末尾（**不弹窗打断**）；所有级别都失败时自动发送 Ctrl+V 回退到 OneNote 原生粘贴。每次尝试的 XML 转储在 `%LOCALAPPDATA%\MD2OneNote\failed-update-L{n}.xml`。

## 安装

| 方式 | 说明 |
|---|---|
| **MSI（推荐）** | 从 [GitHub Releases](https://github.com/volcanochen/2Note/releases) 下载 `2Note-Setup-1.1.0.msi`，双击安装（per-user，无需管理员），自动完成 COM 注册 |
| **便携版** | 下载 `2Note-1.1.0-portable.zip` 解压到任意目录，运行 `INSTALL.ps1` 注册；卸载运行 `UNINSTALL.ps1` |

安装完成后**重启 OneNote** 即可使用。

## 文件布局（MSI 安装后）

```
%LOCALAPPDATA%\2Note\
├─ MD2OneNote.Addin.dll   ← 插件本体（x64, .NET 4, 无外部依赖）
├─ hotkey.txt             ← 快捷键配置（首个非 # 行生效，默认 Alt+Shift+V）
├─ set-hotkey.ps1         ← 快捷键设置脚本（含占用检测）
└─ README.md

注册表（均为 HKCU，无需管理员；HKCU Classes 优先级高于旧版手动注册的 HKLM）：
HKCU\Software\Classes\CLSID\{4A9D4FBD-D100-41EF-9C9C-7CD1DB2EE81B}\InprocServer32 → mscoree.dll, CodeBase 指向上面 DLL
HKCU\Software\Microsoft\Office\OneNote\AddIns\MD2OneNote.Connect                  → LoadBehavior=3（启动自动加载）

日志：%LOCALAPPDATA%\MD2OneNote\addin.log（每次粘贴全链路日志）
```

> 旧版（1.0.0 之前手动安装）位于 `C:\ProgramData\MD2OneNote`。**注意**：其 HKLM 注册会遮蔽新版 HKCU 注册（插件 DllHost 代理由 DCOM 服务在 SYSTEM 上下文启动，只读 HKLM），导致新版更新不生效。运行 `installer\cleanup-legacy.ps1`（自提权）一键清理；清理前可通过直接覆盖 `C:\ProgramData\MD2OneNote\MD2OneNote.Addin.dll` 临时更新代码。

## 快捷键配置

安装目录下的 `set-hotkey.ps1`（MSI 安装后即 `%LOCALAPPDATA%\2Note\set-hotkey.ps1`）：

```powershell
# 测试某个组合键当前是否被占用（不保存）
& "$env:LOCALAPPDATA\2Note\set-hotkey.ps1" -Check Ctrl+Alt+P

# 保存（可用才写入）
& "$env:LOCALAPPDATA\2Note\set-hotkey.ps1" Ctrl+Alt+P
```

支持：修饰键 `Ctrl`/`Alt`/`Shift`/`Win`（至少一个）+ 主键（`A-Z`、`0-9`、`F1`-`F24`、`Home`/`End`/`Insert`/`Delete`/`PageUp`/`PageDown`/方向键/`Space`/`Tab`/`Enter`/`Backspace`/`Esc`）。保存后**重启 OneNote** 生效。

> 注意：快捷键是全局注册的，会截获其他应用里的同组合键。冲突时换一个即可，右键菜单和选项卡按钮不受影响。

## 故障排查

1. **看日志**：`%LOCALAPPDATA%\MD2OneNote\addin.log`——正常粘贴的完整链路是
   `clipboard read (htmlFmt>0)` → `parsed html: items=N` → `page found` → `page xml: N chars` → `update OK`。
2. **粘出来没格式**：日志里 `htmlFmt=0` 说明剪贴板没有 HTML 格式数据（复制来源只给了纯文本），属来源限制而非故障。
3. **快捷键没反应**：日志搜 `hotkey`——`register FAILED` 表示组合键被其他程序占用，换一个组合。
4. **升级/重装**：重新运行新版 MSI 即可（自动覆盖文件与注册）；或直接把新 DLL 覆盖到安装目录（OneNote 不锁此文件，重启 OneNote 生效）。
5. 插件代码运行在独立 COM 代理进程（DllHost）里，崩溃不影响 OneNote 本体。
6. **图片没贴上**：日志搜 `img dropped` / `svg rasterize failed`——前者是图片 URL 抓取失败（目标站防盗链/超时），后者是缺 Edge/Chrome 或 SVG 尺寸异常；日志还有 `html images: found=N resolved=M` 汇总。
7. **粘贴后末尾有「[2Note 提示]」**：说明写入触发了降级（内容已落页），日志 `update OK after degradation: Lx | first cause: ...` 有完整原因；降级 XML 转储在同目录 `failed-update-L{n}.xml`。
8. **来源行没有标题/链接**：标题取剪贴板所有者的顶层窗口标题（日志 `clipboard owner window title`），无窗口标题或复制到纯文本时不显示；链接只对 http(s) 显示（file:// 不显示）。

## 开发者：构建与打包

```powershell
# 源码在 src\（单文件 MD2OneNote.cs，无 PIA 依赖）
src\build.ps1                       # csc 编译 → src\MD2OneNote.Addin.dll

# 构建 MSI + 便携 zip 到 release\（需要 installer\wix311，WiX 3.11 二进制）
installer\build-msi.ps1
installer\make-portable.ps1
```

技术细节见 [docs/技术原理.md](docs/技术原理.md)，完整踩坑过程见 [docs/开发探索历史.md](docs/开发探索历史.md)。
