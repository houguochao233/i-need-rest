# 项目技术交接文档

> 文档目的：让下一位模型/开发者在最少上下文和 token 下快速理解项目、当前实现、已知限制与下一步工作。
>
> 最后核对日期：2026-09-23

## 1. 项目定位

项目目录：`C:\Users\lenovo\Desktop\想摸鱼了`

这是一个 Windows 本地工具，核心功能是：

1. 用户拖入一个文件或 `.lnk` 快捷方式；
2. 备份原对象；
3. 在原位置生成/替换一个外观尽量一致的启动入口；
4. 双击入口后显示全屏“Windows 蓝屏风格”伪蓝屏计时器；
5. 到时间、按热键或手动操作后恢复原对象。

安全边界：只做视觉模拟，不触发真实 BSOD、不重启系统、不修改注册表/驱动/启动项、不联网、不拦截 `Ctrl+Alt+Delete`。

## 2. 文件结构

```text
bsod_shortcut.py          # 主程序：GUI、备份/替换/恢复、伪蓝屏窗口、启动入口
 desktop_icon_positions.py # Explorer 桌面图标位置读取/恢复，使用 Win32 ListView 消息
requirements.txt           # pywin32、Pillow、tkinterdnd2
README.md                  # 面向用户的简要说明，内容可能落后于本交接文档
.gitignore                 # 忽略 data、__pycache__、*.pyc

data/
  profiles/<id>.json       # 每次接管的配置/元数据
  backups/<id>/            # 原始文件或原始 .lnk 备份
```

## 3. 运行环境

- Windows
- Python 3.11（当前环境已验证）
- 依赖：

```powershell
python -m pip install -r .\requirements.txt
python .\bsod_shortcut.py
```

依赖：`pywin32>=306`、`Pillow>=10.0`、`tkinterdnd2>=0.4`。

入口：

```powershell
python .\bsod_shortcut.py
python .\bsod_shortcut.py --launch <profile_id>
```

第二种形式由被接管的 `.lnk` 启动，不显示控制台（配置使用 `pythonw.exe`，若不存在则回退到 `python.exe`）。

## 4. 主程序模块说明

### 4.1 存储和 profile

`ensure_storage()` 创建 `data/profiles` 与 `data/backups`。

`save_profile()` / `read_profile()` / `all_profiles()` 使用 JSON 保存状态。

profile 关键字段：

```json
{
  "id": "uuid hex",
  "source_kind": "shortcut | file",
  "shortcut_path": "当前可双击入口的绝对路径",
  "original_path": "原文件/原快捷方式的绝对路径",
  "backup_path": "备份文件的绝对路径",
  "original": "快捷方式属性或普通文件名称信息",
  "created_at": 0,
  "restored_at": 0,
  "duration_seconds": 1200,
  "start_percent": 0,
  "hotkey": "ctrl+shift+q",
  "restore_mode": "once | after | manual",
  "desktop_position": {"title": "...", "x": 0, "y": 0}
}
```

兼容历史 profile：没有 `source_kind` 时默认按 `shortcut` 处理。

### 4.2 快捷方式模式

适用于拖入的 `.lnk`：

1. `shortcut_data()` 通过 `WScript.Shell` 读取目标、参数、工作目录、图标、描述、窗口模式；
2. 原 `.lnk` 用 `shutil.copy2()` 复制到 `data/backups/<id>/original.lnk`；
3. `replace_shortcut()` 把原 `.lnk` 改成启动当前 Python 程序的 launcher；
4. 原文件名、原路径和图标位置保留；
5. `restore_profile()` 用备份 `.lnk` 原样覆盖回来。

launcher 目标逻辑在 `configure_launcher()`：

- 目标：`pythonw.exe`；
- 参数：`"bsod_shortcut.py" --launch "<profile_id>"`；
- 工作目录：项目目录；
- 图标：原快捷方式的 `IconLocation`，无则使用原目标的 `,0`。

### 4.3 普通文件模式

除 `.lnk` 外，只要是 `Path.is_file()` 就接受；文件夹不接受。

处理逻辑：

1. 原文件移动到 `data/backups/<id>/original<原扩展名>`；
2. 原位置生成 `<原文件名>.lnk`；
3. 该 `.lnk` 由 `create_shadow_shortcut()` 创建，图标源使用备份文件/原文件路径；
4. 恢复时删除代理 `.lnk`，把备份文件复制回 `original_path`；
5. 代理启动后按同一 profile 显示伪蓝屏。

重要限制：普通文件不能在不改文件系统/关联的情况下“原扩展名无痕拦截双击”。如果资源管理器开启显示扩展名，用户会看到额外的 `.lnk`；这是 Windows 限制。默认隐藏已知扩展名时，显示效果接近原文件名。

### 4.4 恢复逻辑

`restore_profile(profile, preserve_current=False)`：

- 检查备份存在；
- 普通文件：删除代理 `.lnk`，备份复制到原路径；
- `.lnk`：备份复制回原路径；
- 使用临时同目录文件 + `os.replace()`，避免半写入；
- `preserve_current=True` 时先保存当前入口副本；
- 恢复后调用 `restore_saved_desktop_position()`。

`profile_is_active()`：通过读取当前 `.lnk` 的参数中是否包含 `--launch` 和 profile id 判断是否仍被接管。对普通文件，它检查代理 `.lnk`。

一次性模式：`launch()` 在显示蓝屏前先恢复原对象，保证程序异常退出时原对象已恢复。

结束后恢复：`BsodWindow.exit()` 在 `restore_mode == "after"` 时恢复。

手动恢复：由主窗口按钮恢复。

### 4.5 伪蓝屏窗口

`BsodWindow`：

- Tk 无边框、置顶、全屏覆盖虚拟桌面；
- 读取 `SM_XVIRTUALSCREEN` 等 Win32 指标，支持多显示器总区域；
- 蓝色背景 `#0078d7`；
- 中文 Windows 10 风格文案、哭脸、百分比、QR-like 图块、停止代码；
- 总时长来自 `duration_seconds`；
- `start_percent` 到 100% 线性递进；
- 配置热键退出，同时绑定 `Esc` 作为紧急退出键；
- 不拦截 Windows 安全组合键。

二维码是 Pillow 绘制的稳定伪二维码样式，不用于真实扫描。

### 4.6 管理器 GUI

`ShortcutManager(TkinterDnD.Tk)`：

- 小型不可拉伸窗口；
- 拖放区接受 Explorer 拖入；
- 点击选择使用 `askopenfilename(..., filetypes=[("所有文件", "*.*")])`；
- `set_shortcut()` 只拒绝文件夹/不存在路径，接受 `.lnk` 和其他普通文件；
- 设置：时长 1–180 分钟、开始进度 0–99%、热键、恢复模式；
- 预览蓝屏不修改文件；
- 列表显示已管理项目；
- 支持恢复选中项和恢复全部。

拖放依赖 `tkinterdnd2`，第一版曾尝试直接替换 Tk 窗口过程并导致拖放闪退，现已改为 TkDND。

## 4.5 BSOD themes

Theme configuration is in `THEMES`. The profile field is `theme`; allowed values are `win11`, `win10`, `win7`, and `winxp`. Missing/unknown values fall back to `win10`.

`BsodWindow.build_layout()` has one visual branch per preset:

- `win11`: modern blue layout, Segoe UI, face, progress, QR-like block, stop code;
- `win10`: Chinese modern layout and current default;
- `win7`: dark-blue classic stop screen with technical information and STOP code;
- `winxp`: solid `#0000AA`, Lucida Console, classic English stop-screen text.

Themes only affect the fake BSOD UI. They do not affect backup, restore, launcher, or desktop-position logic.

## 5. 桌面图标位置

`desktop_icon_positions.py` 通过 Win32 查找 Explorer 的桌面 `SysListView32 / FolderView`，使用远程进程内存和 ListView 消息读取/设置项目位置。

接口：

```python
capture(candidates: list[str]) -> dict | None
restore(position: dict | None, candidates: list[str]) -> bool
```

主程序只对 `path.parent.name.casefold() == "desktop"` 的对象记录坐标。

恢复依据显示标题匹配 `path.name` 和 `path.stem`，因此同名桌面图标可能产生歧义。

限制：

- 只处理桌面目录中的项目；普通目录不处理坐标；
- Explorer 开启“自动排列图标”时，Windows 可能覆盖程序设置的位置；
- 需要 Explorer 桌面窗口可访问，失败时静默跳过，不影响文件恢复；
- 当前实现是 best-effort，不应宣传为 100% 坐标保证。

## 6. 已完成验证

已在当前 Windows/Python 环境做过：

1. `python -m py_compile bsod_shortcut.py desktop_icon_positions.py`；
2. GUI 启动 smoke test；
3. `.lnk`：备份 → 替换 → 恢复，并校验目标、参数、图标；
4. 普通文件代理流程，测试扩展名：`.txt`、`.docx`、`.xlsx`、`.pdf`、`.jpg`、`.zip`、`.mp3`、`.mp4`、`.exe`、`.url`、`.unknown`；
5. 普通文件恢复后校验原始字节内容和代理 `.lnk` 删除；
6. 桌面位置模块导入和不存在项目探测。

这些测试多为脚本级 smoke test，不等价于人工确认所有 Explorer 主题、缩放、多屏、自动排列组合。

## 7. 当前已知问题 / 风险

1. `README.md` 的开头仍主要按“只支持 .lnk”书写，需同步为普通文件 + 桌面坐标的说明。
2. `file_hash()` 当前未使用；`subprocess` 导入也未使用，可后续清理。
3. 恢复时如果原路径被用户占用、文件被锁定、权限不足、磁盘/网络路径不可用，操作会失败并提示。
4. 程序目录、`bsod_shortcut.py`、`data` 不能在仍有 active profile 时移动或删除，否则 launcher 可能无法启动或无法恢复。
5. 普通文件模式使用移动而非复制，适合本地文件；大文件不会产生额外副本，但中断/权限异常需要进一步增强回滚。
6. 快捷方式和普通文件的“完全像原物”受 Windows Explorer 扩展名显示、图标缓存和桌面自动排列影响。
7. `BsodWindow` 当前文本和布局是高仿固定风格，不是内核真实 BSOD，也不是针对所有 Windows 版本的像素级还原。
8. 退出热键解析较轻量，复杂键名/国际键盘布局未专门覆盖。
9. 桌面坐标实现涉及 Explorer 跨进程内存操作，Windows 版本更新后可能需要维护。
10. 当前 profile 没有正式迁移版本号；修改 schema 时应新增 `schema_version` 并兼容旧 JSON。

## 8. 下一步建议（按优先级）

### P0：可靠性

- 接管普通文件前写入 profile 之前先完成所有可逆操作，失败时自动回滚移动；
- 恢复时校验备份与目标冲突，避免覆盖用户新建的同名文件；
- 为 profile 加 `schema_version`、状态字段（`active/restored/error`）；
- 启动时扫描 active profile，检查代理/备份是否存在并提示恢复；
- 把桌面图标位置恢复改为明确返回结果和用户提示，而非静默失败。

### P1：可维护性

- 拆分模块：`storage.py`、`shortcut.py`、`bsod_ui.py`、`desktop_icon_positions.py`、`app.py`；
- 加入 `tests/`，将普通文件和 `.lnk` 的备份、替换、恢复抽成可自动化测试；
- 修正文档、删除未使用导入/函数；
- 增加日志文件，但默认只写本地、不可包含敏感路径之外的数据。

### P2：体验

- 增加“桌面自动排列检测”提示；
- 增加图标缓存刷新/Explorer 刷新后的重试机制；
- 配置界面增加原路径、备份路径、桌面坐标预览；
- 增加 Windows 10/11 主题选择和更多进度节奏；
- 用 PyInstaller 打包成不依赖 Python 的轻量 exe，注意不要把 `data` 备份目录和 launcher 依赖关系弄丢。

## 9. 接手时推荐的最短流程

```powershell
cd "C:\Users\lenovo\Desktop\想摸鱼了"
python -m py_compile .\bsod_shortcut.py .\desktop_icon_positions.py
python .\bsod_shortcut.py
```

先用一个临时 `.txt` 文件测试：

1. 拖入；
2. 选择“手动恢复”；
3. 生成；
4. 确认原位置出现 `.txt.lnk`；
5. 点“恢复选中项”；
6. 确认原 `.txt` 内容和路径恢复。

再用一个临时 `.lnk` 测试快捷方式模式。不要一开始拿重要文件或正在运行的程序测试。

## 10. 给下一模型的工作约束

- 先读本文件，再读 `bsod_shortcut.py` 的相关函数，不要无目的全文重写。
- 修改前优先保留现有 profile 兼容性。
- 不要把普通文件夹加入支持范围，除非用户明确改变需求。
- 不要实现真实蓝屏、驱动、注册表劫持或拦截安全组合键。
- 涉及文件移动/恢复时必须保留可回滚路径，避免覆盖用户原数据。
- 每次改动至少运行语法检查和一个备份/恢复 smoke test。
