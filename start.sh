#!/usr/bin/env bash
set -euo pipefail

# 切换到脚本所在项目根目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# 自动探测可用的 Python 解释器
find_python() {
  if [ -x ".venv/Scripts/python.exe" ]; then
    echo ".venv/Scripts/python.exe"
    return
  fi
  if [ -x ".venv/bin/python" ]; then
    echo ".venv/bin/python"
    return
  fi
  if command -v python >/dev/null 2>&1 && python -c "import sys" >/dev/null 2>&1; then
    echo "python"
    return
  fi
  if command -v python3 >/dev/null 2>&1 && python3 -c "import sys" >/dev/null 2>&1; then
    echo "python3"
    return
  fi
  if [ -n "${LOCALAPPDATA:-}" ]; then
    for candidate in "$LOCALAPPDATA"/Python/pythoncore-*/python.exe "$LOCALAPPDATA"/Programs/Python/Python*/python.exe; do
      if [ -x "$candidate" ]; then
        echo "$candidate"
        return
      fi
    done
  fi
  echo ""
}

PYTHON_BIN="$(find_python)"
if [ -z "$PYTHON_BIN" ]; then
  echo "❌ 未找到可用的 Python 解释器，请先安装 Python 3.10+。"
  exit 1
fi

# 检查核心依赖，若缺失则自动安装
if ! "$PYTHON_BIN" -c "import openpyxl, pymupdf, tqdm, win32com.client" >/dev/null 2>&1; then
  echo "📦 检测到依赖尚未安装完整，正在自动安装 requirements.txt ..."
  "$PYTHON_BIN" -m pip install -r requirements.txt
fi

MODE="${1:-web}"

case "$MODE" in
  web|--web|-w)
    echo "🚀 正在启动 Web 工具 (主页 / · 精简 Sheet /clean · 切字工作台 /workbench · 执行记录看板 /records)..."
    exec "$PYTHON_BIN" web_clip_fixer.py
    ;;
  cli|--cli|-c)
    echo "💻 正在启动 CLI 命令行批处理工具..."
    exec "$PYTHON_BIN" cli_all_in_one.py
    ;;
  menu|--menu|-m)
    echo "================================================"
    echo "  📐 Excel 整理、PDF 导出与切字检查工具"
    echo "================================================"
    echo "  1) 启动 Web 可视化工具 (功能主页 + 切字检查工作台，默认，推荐)"
    echo "  2) 启动 CLI 命令行一键批处理工具"
    echo "================================================"
    read -r -p "请选择启动模式 [1/2，默认 1]: " choice
    if [ "${choice:-1}" = "2" ]; then
      exec "$PYTHON_BIN" cli_all_in_one.py
    else
      exec "$PYTHON_BIN" web_clip_fixer.py
    fi
    ;;
  help|--help|-h)
    echo "用法: ./start.sh [web|cli|menu]"
    echo "  ./start.sh        默认启动 Web 可视化工具 (主页 http://127.0.0.1:8765)"
    echo "  ./start.sh web    启动 Web 可视化工具"
    echo "  ./start.sh cli    启动 CLI 命令行批处理工具"
    echo "  ./start.sh menu   交互式选择启动模式"
    ;;
  *)
    echo "❌ 未知参数: $MODE"
    echo "用法: ./start.sh [web|cli|menu]"
    exit 1
    ;;
esac
