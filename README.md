# docx-rightclick-cleaner

**【macOS Only】右键一下，让乱掉的 Word 文档焕然一新**

> One right-click to clean up your messy Word documents — macOS only.

---

## 你是否也遇到过这些？

- 多人协作的合同、报告，被改得**修订痕迹遍布全文**，接受/拒绝修订要点半天
- 从 AI 工具（ChatGPT、Claude、Kimi……）复制内容粘进 Word，结果**五光十色、字体大小乱跳**，`**加粗**`、`## 标题` 这些 Markdown 符号原封不动地出现在文档里
- 接手了别人传递过来的 Word 文件，里面字体、段距、缩进**各自为政**，想统一格式要一段一段地手动改
- 打开文档第一步：接受修订 → 全选 → 清除格式 → 重新设字体段落 → 再排版……**光整理格式就要折腾半小时**

---

## 这个工具做什么

**右键点一下，全搞定。**

安装后，在 Finder 中对任意 `.docx` 或 `.doc` 文件点击右键，选择「快速操作 → DOCX格式整理」，工具会在原文件同目录生成一个 `文件名_整理.docx`：

- ✅ 所有修订自动接受，修订痕迹全部清除
- ✅ Markdown 符号（`**`、`##`、`-`、`\`……）自动清除
- ✅ emoji 清除
- ✅ 字体、段落、行距统一重排，输出干净整洁的 Word 格式
- ✅ 原文件保持不变

**全程不用打开 Word，右键一下走人。**

---

## 💡 但这不只是个「清洗 Word」的工具

这个项目的真正价值，是**提供了一套在 macOS 上创建右键快捷工具的方法论**。

DOCX 格式整理，只是这个方法落地的一个场景。它的核心能力是这样的：

> **Finder 里看到文件 → 右键 → 完成处理 → 新文件出现在原地**

整个流程不需要打开任何应用、不需要记住命令行、不需要切窗口。**对每天需要大量处理文档的人——以及受腱鞘炎困扰、想尽量减少点击和键盘操作的人——尤其友好。**

如果你有个性化的文档处理需求，完全可以把这个项目当作**「右键工具种子计划」**：

使用者只需告诉AI工具：你去看看我的鼠标右键工具，有个word整理工具，把它作为种子计划，复制它的工作流（不可以另起炉灶原创），现在我需要制作新的右键工具，我想要实现的功能是……

比如：
- 右键把 `.md` 转成排版好的 Word
- 右键批量给 PDF 加水印
- 右键提取合同里的关键条款生成摘要
- 右键把图片批量压缩并重命名

**思路打开，一切在 Finder 里重复做的操作，都值得一个右键菜单。**

---

## 处理前 vs 处理后

| 处理前 | 处理后 |
|---|---|
| 红绿修订痕迹满屏 | 干净正文，无痕迹 |
| `**重要条款**`、`## 一、总则` 裸露在文档里 | 正常文字，无 Markdown 符号 |
| 字体混杂：宋体/等线/Calibri 各一段 | 统一字体排版 |
| 粘贴自 AI 的内容格式一团乱 | 整洁可读，可直接提交 |

---

## 安装方式

**方式一：通过 AI 助手安装（推荐）**

在 AI 助手（如 Claude Code、WorkBuddy 等）中触发此 Skill，说「帮我安装 DOCX 格式整理右键工具」，安装脚本会自动完成全部步骤，安装完成后直接在 Finder 里右键使用。

**方式二：手动安装**

```bash
git clone https://github.com/MarvinLann/docx-rightclick-cleaner.git
cd docx-rightclick-cleaner
python3 scripts/install.py
```

**依赖：**
- macOS 10.15+
- Python 3.8+
- [pandoc](https://pandoc.org/installing.html)（`brew install pandoc`）
- python-docx（`pip3 install python-docx`）
- LibreOffice（可选，用于 `.doc` → `.docx` 转换）

---

## 直接调用（无需安装右键菜单）

如果只想临时处理文件，可以直接运行主脚本：

```bash
python3 scripts/docx_format_cleaner.py /path/to/你的文件.docx
```

输出文件为 `/path/to/你的文件_整理.docx`，与原文件同目录。

---

## 工作原理

```
原始 .docx / .doc
    → 接受所有修订（XML 直接操作，无需打开 Word）
    → pandoc 提取纯文字结构（转为 Markdown）
    → 清洗 Markdown（去反斜杠、emoji、残留符号）
    → python-docx 按统一样式重新排版
    → 输出 xxx_整理.docx
```

---

## 项目结构

```
docx-rightclick-cleaner/
├── SKILL.md                     ← AI 助手 Skill 定义文件
├── scripts/
│   ├── install.py               ← 一键安装右键工具
│   ├── docx_format_cleaner.py   ← 主脚本（全流程）
│   ├── docx2md_converter.py     ← docx → md（接受修订 + pandoc）
│   └── md2docx_plain.py         ← md → docx（python-docx 重排）
├── assets/
│   └── DOCX格式整理.workflow/   ← Automator 种子 workflow
└── references/
    └── macos-workflow-rules.md  ← 技术约束文档（维护者参考）
```

---

## License

Apache 2.0
