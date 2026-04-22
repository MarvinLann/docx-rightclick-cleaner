# macOS 右键工具技术约束（Skill 内部参考）

> 本文档面向 Skill 维护者和对机制感兴趣的使用者。
> 
> **普通用户无需阅读此文档。** 运行 `scripts/install.py` 即可自动完成全部安装。

---

## 核心原则

本 Skill 的右键工具基于 macOS Automator Quick Action 实现。macOS 对 Services/Quick Action 有严格的安全校验机制，**必须从已验证的种子 workflow 复制而来，不可从零创建。**

本 Skill 已内置验证可用的种子 workflow：`assets/DOCX格式整理.workflow/`。

---

## Skill 安装机制

`scripts/install.py` 封装了全部安装逻辑，用户只需执行：

```bash
python3 <skill_dir>/scripts/install.py
```

install.py 内部执行以下步骤：

| 步骤 | 操作 | 目的 |
|---|---|---|
| 1 | 预检依赖（Python、python-docx、lxml、pandoc） | 避免安装后无法运行 |
| 2 | 复制 `scripts/` 下三个 `.py` 到 `~/.docx-cleaner/` | 安装工具脚本 |
| 3 | `cp -R` 复制种子 workflow 到 `~/Library/Services/` | 保留种子元数据 |
| 4 | `plistlib` 修改 `document.wflow` 内的 `COMMAND_STRING` | 注入实际 Shell 脚本 |
| 5 | `touch -t` 同步 workflow 内所有文件时间戳 | 避免系统不信任新文件 |
| 6 | `pbs -flush` + `killall Finder` | 刷新缓存使菜单项生效 |
| 7 | 验证文件存在性和路径正确性 | 确认安装成功 |

---

## 关键技术约束（天条）

以下约束是 install.py 的设计原则，**违反会导致 workflow 被系统拒绝或无法出现在右键菜单中**。

### 1. 绝不从零创建 workflow

- **原因**：macOS 通过 `com.apple.provenance` 等扩展属性标记 GUI 创建的文件。从零创建的目录和 plist 缺少这些标记，会被 Gatekeeper 拒绝。
- **做法**：Skill 内置种子 workflow，install.py 用 `cp -R` 完整复制。
- **禁止**：用 Python/`os.mkdir`/文本编辑器手动创建 `.workflow` 目录结构。

### 2. 绝不用 sed/echo/cat 修改 plist

- **原因**：`document.wflow` 是二进制 plist 格式。文本替换会破坏格式，导致 Automator 无法解析。
- **做法**：install.py 使用 `plistlib` 精确读取和写入二进制 plist。
- **禁止**：任何字符串替换、文本追加、正则替换操作。

### 3. 必须同步时间戳

- **原因**：macOS Services 系统会检查 workflow 文件的时间戳。如果 workflow 内部文件的时间戳与种子差异过大，系统可能不信任该 workflow。
- **做法**：install.py 复制后，用 `touch -t` 将所有内部文件的时间戳统一刷新为当前时间。

### 4. 必须刷新 pbs 缓存

- **原因**：macOS 缓存了 Services 菜单项，新增或修改 workflow 后必须刷新缓存才能生效。
- **做法**：`/System/Library/CoreServices/pbs -flush` + `killall Finder`。
- **注意**：Finder 重启约需 3-5 秒。

### 5. COMMAND_STRING 使用 `$HOME` 变量

- **原因**：不同用户的家目录路径不同（如 `/Users/alice` vs `/Users/bob`），硬编码路径会导致脚本找不到。
- **做法**：workflow 内嵌的 Shell 脚本模板使用 `"$HOME/.docx-cleaner/..."` 和 `"$HOME/.docx-cleaner/logs/..."`。
- **禁止**：在 Shell 脚本中写死绝对路径如 `/Users/xxx/...`。

---

## 故障排查

### 右键菜单不出现

1. 确认 workflow 已复制到 `~/Library/Services/DOCX格式整理.workflow/`
2. 运行 `/System/Library/CoreServices/pbs -flush && killall Finder`
3. 等待 5 秒后重试

### "操作遇到错误"

1. 检查日志：`cat ~/.docx-cleaner/logs/docx_format_cleaner.log`
2. 常见原因：
   - pandoc 未安装 → `brew install pandoc`
   - python-docx 未安装 → `pip3 install python-docx lxml`
   - LibreOffice 未安装（处理 .doc 时需要）→ `brew install --cask libreoffice`

### 输出文件在 Word 中只读

```bash
xattr -d com.apple.quarantine /path/to/xxx_整理.docx
```

---

## 给 Skill 维护者

### 如何更新种子 workflow

种子 workflow 位于 `assets/DOCX格式整理.workflow/`。如需修改（如调整支持的文件类型、修改菜单显示名称）：

1. **不要直接编辑 assets 下的文件**——先复制到 `~/Library/Services/` 测试
2. 用 Automator GUI 打开并修改
3. 测试通过后，用 `cp -R` 覆盖 `assets/` 下的种子
4. 更新 install.py 中的 `WORKFLOW_SHELL_TEMPLATE`（如 Shell 脚本逻辑有变）

### 种子 workflow 的最低要求

一个合格的种子必须满足：

1. 通过 Automator GUI 创建（快速操作类型）
2. 使用 `/bin/bash` 作为 Shell
3. "传递输入"设置为"作为自变量"
4. 保存在 `assets/` 目录下
5. 包含至少一个"运行 Shell 脚本"操作

---

## 常见错误方法及其后果

| 错误方法 | 失败原因 | 现象 |
|---|---|---|
| 脚本从零创建 `.workflow` 目录 | 缺少系统信任标记 | "操作遇到错误" / Gatekeeper 阻止 |
| `sed`/`cat`/`echo` 修改 `document.wflow` | 破坏二进制 plist 格式 | workflow 无法解析 |
| 手动编辑 XML | 格式校验失败 | "workflow 损坏" |
| 硬编码绝对路径 | 不同用户路径不同 | 脚本找不到，处理失败 |
