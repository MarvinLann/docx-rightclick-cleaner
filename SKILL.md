---
name: docx-rightclick-cleaner
description: >
  【macOS Only】一键生成右键菜单 · 整理DOCX格式。
  One-Click Right-Click Menu · Clean DOCX Format (macOS Only).
  当用户提到以下任一意图时，必须立即使用此 Skill：
  「生成右键整理docx」「一键整理docx」「DOCX格式整理」「Word格式清洗」「右键整理Word」
  「安装DOCX右键工具」「Word去格式」「Word文档格式统一」「清除修订痕迹」
  「安装右键工具」「docx整理」「格式整理」「去掉Word格式」「清理docx」
  「Word排版整理」「文档格式清洗」。
  本 Skill 帮助 macOS 用户将 DOCX 整理工具安装为 Finder 右键菜单项，
  安装后用户在任意 .docx/.doc 文件上右键即可一键整理格式，无需打开任何应用。
  也可直接在 AI 助手中运行格式整理（无需安装右键菜单）。
  负向触发：非 macOS 系统、PDF/Excel/PPT 格式整理、Windows 右键工具，不使用此 Skill。
---

# 一键生成右键菜单 · 整理DOCX格式

## 快速决策清单

CRITICAL：遇到以下任一情况，**立即**按对应动作执行，不要询问用户额外信息：

1. 用户要求「安装右键工具」「安装DOCX整理」→ **运行 `scripts/install.py`**
2. 用户提供了 `.docx`/`.doc` 文件路径并要求整理 → **直接调用 `scripts/docx_format_cleaner.py <路径>`**
3. 用户说「右键菜单没出现」「安装失败」→ **执行故障排查步骤**
4. 用户要求「卸载」「删除右键工具」→ **执行卸载命令**

IMPORTANT：本 Skill 仅支持 **macOS**。Windows 用户或 PDF/Excel/PPT 整理需求，不使用此 Skill。

---

## 这个工具做什么

把 Word 文档里乱七八糟的格式、修订痕迹、emoji、反斜杠彻底清掉，
输出一个排版干净、可以直接提交的 `文件名_整理.docx`。

**转换流程：**
```
原始 .docx/.doc
    → 接受所有修订（XML 直接操作）
    → pandoc 转 Markdown（提取纯文字结构）
    → 清洗中间 MD（去反斜杠 + emoji）
    → python-docx 重新排版为 .docx
    → 后处理（去残留 md 符号）
    → 输出 xxx_整理.docx（原文件同目录）
```

---

## 使用场景

| 用户说 | 你该做什么 |
|---|---|
| 「帮我安装右键工具」「安装DOCX格式整理」 | 运行 `scripts/install.py`，完成安装 |
| 「整理这个 docx 文件」+ 提供了文件路径 | 直接调用 `scripts/docx_format_cleaner.py` |
| 「安装失败了」「右键菜单没出现」 | 按本 Skill 故障排查步骤处理 |
| 「卸载」「删除右键工具」 | 执行卸载步骤 |

---

## 一、安装右键工具

### 前置条件

安装前必须具备（否则 install.py 会给出明确提示）：

| 依赖 | 安装命令 |
|---|---|
| Python 3.8+ | macOS 自带或 `brew install python` |
| python-docx | `pip3 install python-docx` |
| lxml | `pip3 install lxml` |
| pandoc | `brew install pandoc` |

### 安装步骤

CRITICAL：必须使用 `scripts/install.py`。绝对禁止：
- 用 `sed`/`echo`/`cat` 修改或生成 workflow 文件
- 手动创建 `.workflow` 目录结构
- 跳过预检直接执行后续步骤

```bash
python3 <skill_dir>/scripts/install.py
```

install.py 自动完成：预检依赖 → 安装脚本到 `~/.docx-cleaner/` → `cp -R` 复制种子 workflow → `plistlib` 修改 COMMAND_STRING → 刷新 pbs 缓存 + 重启 Finder → 验证结果。

IMPORTANT：macOS 右键工具天条（违反会导致 workflow 被系统拒绝或无法触发）：

1. **绝不从零创建 workflow** — 必须从种子 `cp -R` 复制
2. **绝不用 sed/echo/cat 修改 plist** — 必须用 `plistlib` 改 `ActionParameters.COMMAND_STRING`
3. **保留种子元数据** — 复制后同步时间戳，不得破坏扩展属性
4. **必须刷新 pbs 缓存** — `/System/Library/CoreServices/pbs -flush` + `killall Finder`
5. **COMMAND_STRING 用 `$HOME` 变量** — 禁止硬编码绝对路径

### 安装成功的标志

- `~/.docx-cleaner/` 目录下有三个 `.py` 脚本
- `~/Library/Services/DOCX格式整理.workflow/` 存在
- 在 Finder 中右键 `.docx` 文件 → 「快速操作」→ 出现「DOCX格式整理」

---

## 二、直接整理文件（无需安装右键菜单）

用户提供了 `.docx`/`.doc` 文件路径时，直接调用：

```bash
python3 <skill_dir>/scripts/docx_format_cleaner.py /path/to/文件.docx
```

IMPORTANT：直接调用时，脚本会从 `~/.docx-cleaner/` 加载依赖脚本。若目录不存在，先运行 `install.py` 或手动将 `scripts/` 下三个文件复制到 `~/.docx-cleaner/`。

### 端到端示例

**用户输入：**「帮我整理一下 ~/Downloads/合同.docx，里面很多修订痕迹和乱格式」

**你的执行流程：**

1. 确认文件存在：`ls ~/Downloads/合同.docx`
2. 直接调用主脚本：
   ```bash
   python3 /Users/lan/Library/Mobile\ Documents/com~apple~CloudDocs/工作文档/07skill开发/docx-cleaner/scripts/docx_format_cleaner.py ~/Downloads/合同.docx
   ```
3. 检查输出：`ls ~/Downloads/合同_整理.docx`
4. 告知用户：整理完成，输出文件为 `合同_整理.docx`，与原文件同目录。

**禁止做的：**不要打开 Word 手动清理；不要用其他脚本替代 `docx_format_cleaner.py`。

---

## 三、故障排查

### 右键菜单没出现

```bash
# 方法1：手动刷新
/System/Library/CoreServices/pbs -flush
killall Finder

# 方法2：重新安装（会覆盖旧版）
python3 <skill_dir>/scripts/install.py
```

### 处理失败 / 无输出文件

查看日志：
```bash
cat ~/.docx-cleaner/logs/docx_format_cleaner.log
```

常见原因：
- `pandoc 未找到`：`brew install pandoc`
- `python-docx 未安装`：`pip3 install python-docx lxml`
- `.doc 文件处理失败`：需要 LibreOffice（`brew install --cask libreoffice`）

### 输出文件在 Word 中只读

```bash
xattr -d com.apple.quarantine /path/to/xxx_整理.docx
```

---

## 四、卸载

```bash
# 删除工具脚本
rm -rf ~/.docx-cleaner/

# 删除右键菜单
rm -rf ~/Library/Services/DOCX格式整理.workflow

# 刷新服务
/System/Library/CoreServices/pbs -flush
killall Finder
```

---

## 五、目录结构

```
docx-cleaner/
├── SKILL.md                     ← 本文件
├── scripts/
│   ├── install.py               ← 一键安装入口
│   ├── docx_format_cleaner.py   ← 主脚本（全流程）
│   ├── docx2md_converter.py     ← docx → md
│   └── md2docx_plain.py         ← md → docx
├── assets/
│   └── DOCX格式整理.workflow/   ← 种子 workflow（cp -R 复制，勿修改）
└── references/
    └── macos-workflow-rules.md  ← 技术约束与故障排查（Skill 维护者参考）
```
