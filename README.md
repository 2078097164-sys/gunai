# Desktop Shortcut Tool (桌面快捷方式备份与恢复工具)

<p align="center">
  <b>一款专为 Windows 重装系统场景打造的纯净、开源、绿色便携桌面恢复神器</b><br>
  一键完整备份与还原桌面快捷方式（.lnk / .url）及其九宫格图标坐标，支持新旧用户名与盘符智能重映射。
</p>

---

## ⚠️ 核心使用注意事项（必看！）

> [!CAUTION]
> **重装系统前极重要提醒：**
> 重装系统格式化 C 盘前，**请务必将本工具的整个文件夹（包含备份生成的 `快捷方式备份_*` 文件夹）直接存放于「系统 U 盘」或「移动硬盘」等非系统盘介质中！**
> 绝对不要存放在系统桌面、我的文档或 C 盘任意目录下，否则重装格式化时备份文件将连同旧系统一起被彻底清空！

---

## ✨ 核心特性

- 🗂️ **全域双桌面扫描**：不仅备份当前用户的私有桌面（`%USERPROFILE%\Desktop`），更同步备份包含 80% 常用软件安装快捷方式的**公共桌面**（`Public Desktop`）。
- 🧬 **结构化元数据提取**：深度解析 `.lnk` 目标路径、运行参数、起始工作目录、自定义图标、热键及备注，生成便于审计的 `backup_info.json` 与 `备份清单.csv`。
- 🎯 **图标九宫格坐标定格**：自动导出系统注册表中的桌面图标排列坐标布局，重装后一键导入并刷新 Explorer 即可还原整洁桌面。
- 🔄 **智能路径与盘符重映射**：
  - 用户名变更：原用户名 `Administrator` ➔ 新用户名 `User`
  - 盘符漂移迁移：原软件安装在 `D:\` ➔ 新系统识别为 `E:\`
  - 自动批量修正快捷方式的目标路径，无需逐个手动重置。
- 📋 **失效目标诊断报告**：恢复完成后自动检测目标 EXE 是否存在，生成清晰的 `恢复报告.txt`，未重装的软件一目了然。
- 🛡️ **纯净开源、零外部依赖**：基于原生 Windows PowerShell 5.1 与 .NET Framework 开发，无任何第三方运行库要求，告别传统 PE 维护工具的报毒与捆绑主页困扰。
- 💎 **现代极客暗黑 UI**：Fluent / GitHub Dark 沉浸风格界面，配备 UAC 一键智能提权、单文件 EXE 绿色便携。

---

## 🚀 使用流程指南

### 1. 系统重装前（备份）
1. 双击打开 `DesktopShortcutTool.exe`（或以管理员身份运行 `launch.bat`）。
2. 确保勾选 **“同时备份 / 恢复桌面图标摆放位置”**。
3. 点击 **「💾 1. 备份当前桌面」**。
4. **【关键步骤】** 备份完成后，将**本工具整个文件夹（包含新生成的 `快捷方式备份_*` 文件夹）拷贝并保存在你的系统 U 盘、移动硬盘或网盘中**。

### 2. 系统重装后（还原）
1. 重装好全新 Windows 系统后，将 U 盘中的工具文件夹拷贝回电脑（或直接在 U 盘中运行）。
2. 双击打开 `DesktopShortcutTool.exe`。
3. 如果重装后的用户名或软件盘符发生改变，勾选 **“启用路径 / 盘符重映射”**，填写旧名称与新名称（例如：`旧: D:\` ➔ `新: E:\`）。
4. 点击 **「🚀 2. 一键还原桌面」**，若弹出管理员 UAC 授权请点击“是”。
5. 提示恢复完成后，选择重启资源管理器，桌面快捷方式及图标排布瞬间满血复活！

---

## 🛠 命令行调用 (CLI 支持)

除了图形界面，本工具亦支持在 PE 自动装机脚本或批处理中静默调用：

```powershell
# 执行备份并保存桌面布局
powershell -ExecutionPolicy Bypass -File DesktopShortcutTool.ps1 -Action backup -SaveLayout

# 执行恢复（带路径重映射）
powershell -ExecutionPolicy Bypass -File DesktopShortcutTool.ps1 -Action restore -BackupDir "D:\快捷方式备份_20260910" -OldUser "oldname" -NewUser "newname" -RestoreLayout

# 检查当前桌面失效的快捷方式
powershell -ExecutionPolicy Bypass -File DesktopShortcutTool.ps1 -Action check
```

---

## 💻 兼容性与运行环境

- **操作系统**：Windows 10 / Windows 11 全系列版本（兼容微PE、Edgeless、优启通等各主流 WinPE 系统）
- **环境要求**：系统自带 PowerShell 5.1（无需额外安装 Python、.NET Core 或 Node.js）

---

## 📄 开源许可证

本项目基于 [MIT License](LICENSE) 协议开源，欢迎自由集成至各类装机维护 U 盘与工具箱。

