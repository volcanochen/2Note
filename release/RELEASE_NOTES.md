# 2Note v1.0.0 发布说明

## 简介

2Note（原名 MD2OneNote）是一个 OneNote COM 加载项：把剪贴板里的富文本（HTML）或 Markdown，以 **OneNote 原生结构**（one:OE / one:OEChildren）直接写入当前页面，完整保留嵌套列表缩进、标题、加粗/斜体、链接、行内代码、复选框等格式——绕过 OneNote 内置 HTML 导入器（它会把嵌套列表压平）。

## 下载

| 文件 | 适用场景 |
|---|---|
| `2Note-Setup-1.0.0.msi` | **推荐**。双击一键安装（per-user，无需管理员），自动完成 COM 注册 |
| `2Note-1.0.0-portable.zip` | 免安装版。解压后运行 `INSTALL.ps1` 手动注册，卸载运行 `UNINSTALL.ps1` |

## 安装后

1. 重启 OneNote
2. 入口三选一：
   - 功能区「视图」右侧的 **2note** 选项卡 →「Paste as OneNote structure」
   - 页面内容区右键 →「粘贴为 OneNote 结构 (Alt+Shift+V)」
   - 全局快捷键 **Alt+Shift+V**（可在安装目录 `hotkey.txt` 或运行 `set-hotkey.ps1` 修改）
3. 从任意应用复制内容 → 在 OneNote 页面里粘贴

## 系统要求

- Windows 10/11 x64
- OneNote 2013 及以上（桌面版，含 Office16）
- .NET Framework 4.x（系统自带）

## v1.0.0 更新内容

- **格式保留**：HTML Format 剪贴板读取（CF_HTML 规范、UTF-8、分步日志诊断）
- **原生结构写入**：嵌套列表/标题/加粗/斜体/链接/行内代码/复选框/代码块，绕过 OneNote HTML 导入器
- **晚绑定 COM**：纯 raw IDispatch 调用 OneNote API，无 PIA 依赖；[out] BSTR 所有权安全（防堆损坏）
- **可配置全局热键**：`set-hotkey.ps1` 带占用检测；热键泵运行在专用 STA 线程
- **稳定加载**：IDTExtensibility2 参数级 MarshalAs；OneMore 式 customUI（2note 选项卡 + 页面右键菜单）
- **一键安装**：per-user MSI，无需管理员；卸载干净（删除文件与 HKCU 注册）

## 已知事项

- 从「仅纯文本」的复制来源粘贴时没有格式可保留（来源限制，日志会显示 `htmlFmt=0`）
- 全局热键会截获其他应用里的同组合键；冲突时用 `set-hotkey.ps1` 换一个即可
- 手动/旧版安装在 `C:\ProgramData\MD2OneNote` 的注册会被 MSI 的 HKCU 注册自动覆盖（HKCU Classes 优先级更高），无需手工清理
