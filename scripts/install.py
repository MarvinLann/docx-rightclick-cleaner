#!/usr/bin/env python3
"""
DOCX 格式整理工具 · 通用安装脚本
========================================
安装流程（严格按顺序执行，不可跳过）：
  Step 1  预检（检查 Python 依赖 + pandoc）
  Step 2  安装工具脚本到 ~/.docx-cleaner/
  Step 3  复制种子 workflow → ~/Library/Services/
  Step 4  用 plistlib 修改 workflow 内 COMMAND_STRING
  Step 5  同步时间戳 + 刷新 pbs 缓存 + 重启 Finder
  Step 6  验证安装结果

注意事项（macOS 右键工具天条）：
  ✅ 必须用 cp -R 复制已验证的种子，不可用 sed/cat/echo 生成 workflow
  ✅ 必须用 plistlib 修改 COMMAND_STRING，不可字符串替换
  ✅ 必须保持种子元数据（时间戳、权限、扩展属性）
"""

import sys
import os
import shutil
import subprocess
import plistlib
import platform
from pathlib import Path


# ── 路径常量 ───────────────────────────────────────────────
SKILL_DIR = Path(__file__).parent.parent.resolve()
SEED_WORKFLOW = SKILL_DIR / "assets" / "DOCX格式整理.workflow"
SCRIPTS_SRC_DIR = SKILL_DIR / "scripts"

INSTALL_DIR = Path.home() / ".docx-cleaner"
SERVICES_DIR = Path.home() / "Library" / "Services"
WORKFLOW_DEST = SERVICES_DIR / "DOCX格式整理.workflow"
WFLOW_PLIST = WORKFLOW_DEST / "Contents" / "document.wflow"
LOG_DIR = Path.home() / ".docx-cleaner" / "logs"


# ── Shell 脚本模板（workflow 内嵌，安装时动态生成）──────────────
WORKFLOW_SHELL_TEMPLATE = r"""#!/bin/bash
export PATH="$HOME/usr/local/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH"

LOG_FILE="$HOME/.docx-cleaner/logs/docx_format_cleaner.log"
mkdir -p "$HOME/.docx-cleaner/logs"
echo "===== DOCX格式整理开始 $(date) =====" >> "$LOG_FILE"

SCRIPT="$HOME/.docx-cleaner/docx_format_cleaner.py"
if [ ! -f "$SCRIPT" ]; then
    echo "错误：找不到脚本 $SCRIPT" >> "$LOG_FILE"
    /usr/bin/osascript -e 'display notification "找不到DOCX格式整理脚本，请重新安装" with title "DOCX格式整理错误"'
    exit 1
fi

# 动态探测 python3
if [ -f "/opt/homebrew/bin/python3" ]; then
    PYTHON="/opt/homebrew/bin/python3"
elif [ -f "/usr/local/bin/python3" ]; then
    PYTHON="/usr/local/bin/python3"
else
    PYTHON="/usr/bin/python3"
fi

for f in "$@"
do
    echo "处理: $f" >> "$LOG_FILE"
    if [[ "$f" == *.docx ]] || [[ "$f" == *.doc ]] || \
       [[ "$f" == *.DOCX ]] || [[ "$f" == *.DOC ]]; then
        "$PYTHON" "$SCRIPT" "$f" >> "$LOG_FILE" 2>&1
        STATUS=$?
        BASENAME=$(basename "$f")
        if [ $STATUS -eq 0 ]; then
            /usr/bin/osascript -e "display notification \"格式整理完成: ${BASENAME}\" with title \"DOCX格式整理\""
        else
            /usr/bin/osascript -e "display notification \"处理失败: ${BASENAME}，请查看日志\" with title \"DOCX格式整理错误\""
        fi
    else
        echo "跳过不支持的文件: $f" >> "$LOG_FILE"
    fi
done

echo "===== DOCX格式整理结束 $(date) =====" >> "$LOG_FILE"
"""


def log(msg: str, indent: int = 0):
    prefix = "  " * indent
    print(f"{prefix}{msg}")


def step(n: int, total: int, title: str):
    print(f"\n>>> Step {n}/{total}：{title}")


def abort(msg: str):
    print(f"\n❌ 安装失败：{msg}")
    sys.exit(1)


# ─────────────────────────────────────────────────────────────
# Step 1：预检
# ─────────────────────────────────────────────────────────────
def preflight():
    step(1, 6, "运行安装前预检")
    errors = []

    # macOS 检查
    if platform.system() != "Darwin":
        abort("本工具仅支持 macOS。")

    # Python 版本
    if sys.version_info < (3, 8):
        errors.append("Python 3.8+ 是必需的，当前版本过低。请升级 Python。")
    else:
        log(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}", 1)

    # python-docx
    try:
        import docx
        log("✅ python-docx 已安装", 1)
    except ImportError:
        errors.append("缺少 python-docx，请运行：pip3 install python-docx")

    # lxml
    try:
        import lxml
        log("✅ lxml 已安装", 1)
    except ImportError:
        errors.append("缺少 lxml，请运行：pip3 install lxml")

    # pandoc
    pandoc_found = any(
        Path(p).exists() for p in [
            "/opt/homebrew/bin/pandoc",
            "/usr/local/bin/pandoc",
            "/usr/bin/pandoc",
        ]
    ) or bool(shutil.which("pandoc"))

    if pandoc_found:
        log("✅ pandoc 已找到", 1)
    else:
        errors.append("缺少 pandoc，请运行：brew install pandoc")

    # 种子 workflow 存在
    if SEED_WORKFLOW.exists():
        log(f"✅ 种子 workflow 已找到：{SEED_WORKFLOW.name}", 1)
    else:
        errors.append(f"种子 workflow 不存在：{SEED_WORKFLOW}")

    if errors:
        print()
        for e in errors:
            log(f"❌ {e}", 1)
        abort("预检失败，请按上述提示修复后重新运行。")

    log("✅ 预检通过", 1)


# ─────────────────────────────────────────────────────────────
# Step 2：安装工具脚本
# ─────────────────────────────────────────────────────────────
def install_scripts():
    step(2, 6, "安装工具脚本到 ~/.docx-cleaner/")
    INSTALL_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)

    scripts = [
        "docx_format_cleaner.py",
        "docx2md_converter.py",
        "md2docx_plain.py",
    ]
    for s in scripts:
        src = SCRIPTS_SRC_DIR / s
        dst = INSTALL_DIR / s
        shutil.copy2(str(src), str(dst))
        os.chmod(str(dst), 0o755)
        log(f"✅ {s} → {dst}", 1)


# ─────────────────────────────────────────────────────────────
# Step 3：复制种子 workflow
# ─────────────────────────────────────────────────────────────
def install_workflow():
    step(3, 6, "复制种子 workflow 到 ~/Library/Services/")
    SERVICES_DIR.mkdir(parents=True, exist_ok=True)

    # 如果已存在，先删除旧版
    if WORKFLOW_DEST.exists():
        shutil.rmtree(str(WORKFLOW_DEST))
        log("♻️  已移除旧版 workflow", 1)

    # ⚠️ 天条：必须用 cp -R 复制，保留所有元数据
    result = subprocess.run(
        ["cp", "-R", str(SEED_WORKFLOW), str(WORKFLOW_DEST)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        abort(f"cp -R 复制种子失败：{result.stderr}")

    log(f"✅ workflow 已复制到：{WORKFLOW_DEST}", 1)


# ─────────────────────────────────────────────────────────────
# Step 4：用 plistlib 修改 COMMAND_STRING
# ─────────────────────────────────────────────────────────────
def patch_workflow():
    step(4, 6, "修改 workflow Shell 脚本（plistlib）")

    # ⚠️ 天条：必须用 plistlib，不可用 sed/echo/cat
    with open(str(WFLOW_PLIST), 'rb') as f:
        data = plistlib.load(f)

    # 找 ActionParameters.COMMAND_STRING
    action = data['actions'][0]['action']
    action['ActionParameters']['COMMAND_STRING'] = WORKFLOW_SHELL_TEMPLATE

    with open(str(WFLOW_PLIST), 'wb') as f:
        plistlib.dump(data, f, fmt=plistlib.FMT_XML)

    log("✅ COMMAND_STRING 已更新", 1)

    # 同步时间戳（保持元数据一致性）
    _sync_timestamps()


def _sync_timestamps():
    """同步 workflow 内部文件的时间戳，避免系统不信任"""
    now_str = subprocess.run(
        ["date", "+%Y%m%d%H%M.%S"],
        capture_output=True, text=True
    ).stdout.strip()

    for dirpath, dirnames, filenames in os.walk(str(WORKFLOW_DEST)):
        for fname in filenames:
            fpath = os.path.join(dirpath, fname)
            subprocess.run(
                ["touch", "-t", now_str, fpath],
                capture_output=True,
            )
    log("✅ 时间戳已同步", 1)


# ─────────────────────────────────────────────────────────────
# Step 5：刷新 pbs 缓存 + 重启 Finder
# ─────────────────────────────────────────────────────────────
def refresh_services():
    step(5, 6, "注册服务（刷新缓存 + 重启 Finder）")

    subprocess.run(
        ["/System/Library/CoreServices/pbs", "-flush"],
        capture_output=True,
    )
    log("✅ pbs 缓存已刷新", 1)

    subprocess.run(
        ["killall", "Finder"],
        capture_output=True,
    )
    log("✅ Finder 已重启（约 3-5 秒后可用）", 1)


# ─────────────────────────────────────────────────────────────
# Step 6：验证安装
# ─────────────────────────────────────────────────────────────
def verify():
    step(6, 6, "验证安装结果")
    ok = True

    checks = [
        (INSTALL_DIR / "docx_format_cleaner.py", "主脚本"),
        (INSTALL_DIR / "docx2md_converter.py", "docx→md 脚本"),
        (INSTALL_DIR / "md2docx_plain.py", "md→docx 脚本"),
        (WFLOW_PLIST, "workflow plist"),
    ]

    for path, name in checks:
        if path.exists():
            log(f"✅ {name}：{path}", 1)
        else:
            log(f"❌ {name} 未找到：{path}", 1)
            ok = False

    # 验证 COMMAND_STRING 已正确写入
    with open(str(WFLOW_PLIST), 'rb') as f:
        data = plistlib.load(f)
    cmd = data['actions'][0]['action']['ActionParameters']['COMMAND_STRING']
    if ".docx-cleaner" in cmd:
        log("✅ workflow Shell 脚本指向正确路径", 1)
    else:
        log("❌ workflow Shell 脚本路径异常，请重新安装", 1)
        ok = False

    return ok


# ─────────────────────────────────────────────────────────────
# 主入口
# ─────────────────────────────────────────────────────────────
def main():
    print("==========================================")
    print("  DOCX格式整理工具 · 一键安装（通用版）")
    print("==========================================")

    preflight()
    install_scripts()
    install_workflow()
    patch_workflow()
    refresh_services()
    ok = verify()

    print()
    print("==========================================")
    if ok:
        print("  ✅ 安装完成！")
        print("==========================================")
        print()
        print("📌 验证方法：")
        print("   1. 等待 Finder 重启（约 3–5 秒）")
        print("   2. 找一个 .docx 文件，右键点击")
        print("   3. 选择「快速操作」→「DOCX格式整理」")
        print("   4. 同目录下会生成「xxx_整理.docx」")
        print()
        print("📌 如遇问题，查看日志：")
        print("   cat ~/.docx-cleaner/logs/docx_format_cleaner.log")
    else:
        print("  ⚠️  安装有异常，请查看上方错误信息。")
        print("==========================================")
        sys.exit(1)


if __name__ == "__main__":
    main()
