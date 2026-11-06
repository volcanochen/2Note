# 2Note — 「粘贴为 OneNote 结构」插件（原名 MD2OneNote）

把剪贴板里的富文本（HTML）或 Markdown，**以 OneNote 原生结构**（`one:OE` / `one:OEChildren`）直接写入当前页面，完整保留嵌套列表缩进、标题、加粗/斜体、链接、行内代码、复选框等格式——**绕过 OneNote 内置 HTML 导入器**（它会把嵌套列表压平），这是本插件与一切「先转 HTML 再 Ctrl+V」方案的本质区别。

## 功能一览

| 入口 | 说明 |
|---|---|
| **2note 功能区选项卡** | 「视图」选项卡右侧的 2note 选项卡，内含大按钮 |
| **页面右键菜单** | 内容区右键 →「粘贴为 OneNote 结构」 |
| **全局快捷键** | 默认 Ctrl+Alt+M；可通过配置文件改为任意组合键（本机已配置为 **Alt+Shift+V**） |

**支持的内容**：
- 多级嵌套列表（ul/ol 真嵌套，以及靠缩进模拟的伪嵌套自动推导层级）
- 有序列表字面编号（同层 ul/ol 混排或空行分隔时自动重置计数）
- 标题 h1–h6、加粗/斜体、链接、行内代码
- 复选框 → ☐/☑
- 代码块 → 逐行 Consolas 等宽段落
- 中文正常（CF_HTML 按 UTF-8 字节读取）
- 纯文本/Markdown 源码兜底（剪贴板无 HTML 格式时按 Markdown 解析）

**插入位置**：新建独立文本容器（Outline），放在光标所在容器的正下方，多次插入自动下推、不重叠。

## 文件布局（本机已安装）

```
C:\ProgramData\MD2OneNote\
├─ MD2OneNote.Addin.dll   ← 插件本体（x64, .NET 4, 无外部依赖）
├─ hotkey.txt             ← 快捷键配置（首个非 # 行生效）
├─ set-hotkey.ps1         ← 快捷键设置脚本（含占用检测）
└─ poison-backup\         ← 修复毒类型库前的注册表备份

注册表：
HKLM\...\CLSID\{4A9D4FBD-D100-41EF-9C9C-7CD1DB2EE81B}\InprocServer32   → mscoree.dll, CodeBase 指向上面 DLL
HKCU\Software\Microsoft\Office\OneNote\AddIns\MD2OneNote.Connect       → LoadBehavior=3（启动自动加载）

日志：C:\Users\<你>\AppData\Local\MD2OneNote\addin.log（每次粘贴全链路日志）
```

## 快捷键配置

```powershell
# 测试某个组合键当前是否被占用（不保存）
C:\ProgramData\MD2OneNote\set-hotkey.ps1 -Check Ctrl+Alt+P

# 保存（可用才写入）
C:\ProgramData\MD2OneNote\set-hotkey.ps1 Ctrl+Alt+P
```

支持：修饰键 `Ctrl`/`Alt`/`Shift`/`Win`（至少一个）+ 主键（`A-Z`、`0-9`、`F1`-`F24`、`Home`/`End`/`Insert`/`Delete`/`PageUp`/`PageDown`/方向键/`Space`/`Tab`/`Enter`/`Backspace`/`Esc`）。保存后**重启 OneNote** 生效。

> 注意：快捷键是全局注册的，会截获其他应用里的同组合键。冲突时换一个即可，右键菜单和选项卡按钮不受影响。

## 故障排查

1. **看日志**：`%LOCALAPPDATA%\MD2OneNote\addin.log`——正常粘贴的完整链路是
   `clipboard read (htmlFmt>0)` → `parsed html: items=N` → `page found` → `page xml: N chars` → `update OK`。
2. **粘出来没格式**：日志里 `htmlFmt=0` 说明剪贴板没有 HTML 格式数据（复制来源只给了纯文本），属来源限制而非故障。
3. **快捷键没反应**：日志搜 `hotkey`——`register FAILED` 表示组合键被其他程序占用，换一个组合。
4. **重装**：把新 DLL 覆盖到 `C:\ProgramData\MD2OneNote\MD2OneNote.Addin.dll`（OneNote 不锁此文件，只需重启 OneNote）。
5. 插件代码运行在独立 COM 代理进程（DllHost）里，崩溃不影响 OneNote 本体。

## 开发者：构建与部署

```powershell
# 源码：md2addin\MD2OneNote.cs（单文件，无 PIA 依赖）
C:\Users\...\md2addin\build.ps1          # csc 编译 → md2addin\MD2OneNote.Addin.dll

# 部署（如被代理进程锁定，先结束命令行含 {4A9D4FBD...} 的 dllhost）
Copy-Item md2addin\MD2OneNote.Addin.dll C:\ProgramData\MD2OneNote\MD2OneNote.Addin.dll -Force
```

技术细节见 [docs/技术原理.md](docs/技术原理.md)，完整踩坑过程见 [docs/开发探索历史.md](docs/开发探索历史.md)。
