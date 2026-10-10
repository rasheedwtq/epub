#!/usr/bin/env bash
# ============================================================
# docx2epub 项目 —— 环境一键卸载脚本
#
#   用法：
#     bash tools/uninstall_env.sh            # 交互确认后卸载
#     bash tools/uninstall_env.sh --yes      # 跳过确认
#     bash tools/uninstall_env.sh --yes --purge-outputs   # 连同样本与文档一起删
#     bash tools/uninstall_env.sh --yes --clean-uv-python # 额外清理闲置的 uv Python(68MB)
#     bash tools/uninstall_env.sh --yes --clean-uv-all    # 额外清理 uv Python + uv 工具(108MB)
#     bash tools/uninstall_env.sh --yes --only-uv         # 【只清 uv 遗留件，保留 venv】← 推荐
#
#   ⚠️ 只删除「本次安装的虚拟环境」与「明确点名的遗留件」，不会碰：
#      · WorkBuddy 自带的 base Python runtime
#      · 系统已安装的 Microsoft Word
#      · 任何未点名的 AppData 内容
# ============================================================

set -u

VENV="C:/Users/45090/.workbuddy/binaries/python/envs/default"
BASE_PY="C:/Users/45090/.workbuddy/binaries/python/versions/3.13.12/python.exe"
PROJECT="C:/backup/epub"
UV_PY="C:/Users/45090/AppData/Roaming/uv/python"          # uv 托管的 Python 3.12.15 (68MB)
UV_BIN="C:/Users/45090/.local/bin"                          # uv 工具本体 (40MB)
UV_LOCAL="C:/Users/45090/AppData/Local/uv"                 # uv 缓存/回执

ASSUME_YES=0
PURGE_OUTPUTS=0
CLEAN_UV_PY=0
CLEAN_UV_ALL=0
ONLY_UV=0
for arg in "$@"; do
  case "$arg" in
    --yes|-y) ASSUME_YES=1 ;;
    --purge-outputs) PURGE_OUTPUTS=1 ;;
    --clean-uv-python) CLEAN_UV_PY=1 ;;
    --clean-uv-all) CLEAN_UV_ALL=1; CLEAN_UV_PY=1 ;;
    --only-uv) ONLY_UV=1; CLEAN_UV_PY=1 ;;
    --help|-h)
      sed -n '2,15p' "$0" | sed 's/^# \{0,1\}//'
      exit 0 ;;
    *) echo "未知参数：$arg（用 --help 查看用法）"; exit 1 ;;
  esac
done

echo "=============================================="
echo " docx2epub 环境卸载"
echo "=============================================="
echo ""
echo "【将删除】"
if [ "$ONLY_UV" -eq 1 ]; then
  if [ -d "$VENV" ]; then
    echo "  · 虚拟环境：**保留不动**（--only-uv 模式）"
    echo "    $VENV  ($(du -sh "$VENV" 2>/dev/null | cut -f1)，含 10 个包)"
  fi
elif [ -d "$VENV" ]; then
  SIZE=$(du -sh "$VENV" 2>/dev/null | cut -f1)
  echo "  · 虚拟环境：$VENV  (约 $SIZE)"
  echo "    内含 10 个包：python-docx / lxml / Pillow / pywin32 / msoffcrypto-tool"
  echo "                  olefile / cryptography / cffi / pycparser / typing_extensions"
else
  echo "  · 虚拟环境不存在（可能已卸载）：$VENV"
fi

if [ "$PURGE_OUTPUTS" -eq 1 ]; then
  echo ""
  echo "【--purge-outputs：还将删除自产文件】"
  for d in samples tools; do
    [ -d "$PROJECT/$d" ] && echo "  · $PROJECT/$d  ($(du -sh "$PROJECT/$d" 2>/dev/null | cut -f1))"
  done
  N=$(ls "$PROJECT"/*.md 2>/dev/null | wc -l)
  [ "$N" -gt 0 ] && echo "  · $PROJECT/*.md  ($N 份文档)"
fi

echo ""
echo "【不会删除】"
echo "  · WorkBuddy base Python：$BASE_PY"
echo "  · Microsoft Word（系统自带）"
echo "  · LibreOffice（本就未安装）"

# ---- 遗留件：uv 安装的闲置 Python（AI 于 10-05 07:37 装，从未使用）----
UV_SIZE=$(du -sh "$UV_PY" 2>/dev/null | cut -f1)
if [ "$CLEAN_UV_PY" -eq 1 ]; then
  echo ""
  if [ "$ONLY_UV" -eq 1 ]; then
    echo "【--only-uv：只清 uv 遗留件，venv 保持不动】"
  else
    echo "【--clean-uv：额外清理 AI 遗留的闲置件】"
  fi
  if [ -d "$UV_PY" ]; then
    echo "  · uv 托管 Python 3.12.15：$UV_PY  (约 $UV_SIZE)"
  else
    echo "  · uv 托管 Python 3.12.15：**已不存在**（可能已清理过）"
  fi
  echo "    取证依据：site-packages 仅 pip（零第三方包）、全盘无 venv 引用 → 从未使用"
  if [ "$CLEAN_UV_ALL" -eq 1 ]; then
    echo "  · uv 工具本体：$UV_BIN/uv*.exe + $UV_LOCAL  (约 40 MB)"
    echo "    ⚠️ 删除后将无法使用 uv 命令"
  fi
fi

if [ "$ASSUME_YES" -ne 1 ]; then
  echo ""
  read -r -p "确认卸载？(yes/N) " ans
  case "$ans" in
    yes|Y|y|YES) ;;
    *) echo "已取消，未做任何改动。"; exit 0 ;;
  esac
fi

echo ""
echo ">>> 正在卸载..."
if [ "$ONLY_UV" -eq 1 ]; then
  echo "  --only-uv 模式：**跳过 venv**（项目环境保留）"
elif [ -d "$VENV" ]; then
  rm -rf "$VENV" || { echo "✗ 删除失败，请检查文件占用（是否有 Word/LO 进程在跑）"; exit 1; }
  echo "  ✓ 虚拟环境已删除（10 个包一并移除）"
else
  echo "  · 跳过（不存在）"
fi

if [ "$PURGE_OUTPUTS" -eq 1 ]; then
  for d in samples tools; do
    [ -d "$PROJECT/$d" ] && { rm -rf "$PROJECT/$d"; echo "  ✓ 已删除 $d/"; }
  done
  if ls "$PROJECT"/*.md >/dev/null 2>&1; then
    rm -f "$PROJECT"/*.md
    echo "  ✓ 已删除 $PROJECT/*.md"
  fi
fi

# ---- 清理 uv 遗留件 ----
if [ "$CLEAN_UV_PY" -eq 1 ]; then
  if [ -d "$UV_PY" ]; then
    rm -rf "$UV_PY" && echo "  ✓ 已删除 uv 托管 Python 3.12.15（$UV_SIZE）"
  else
    echo "  · uv Python 不存在，跳过"
  fi
  # 清理空壳目录
  rmdir "C:/Users/45090/AppData/Roaming/uv" 2>/dev/null && echo "  ✓ 已清理空目录 AppData\\Roaming\\uv"
  if [ "$CLEAN_UV_ALL" -eq 1 ]; then
    rm -f "$UV_BIN"/uv.exe "$UV_BIN"/uvx.exe "$UV_BIN"/uvw.exe 2>/dev/null
    rm -rf "$UV_LOCAL" 2>/dev/null
    echo "  ✓ 已删除 uv 工具本体（uv/uvx/uvw + 缓存）"
    command -v uv >/dev/null 2>&1 && echo "  ⚠ PATH 中仍能找到 uv，请检查其他安装位置" \
                                || echo "  ✓ uv 命令已不可用（符合预期）"
  fi
fi

echo ""
echo ">>> 验证"
if [ -d "$VENV" ]; then
  echo "  ✗ 仍存在：$VENV"
else
  echo "  ✓ venv 目录已不存在"
fi
if "$BASE_PY" --version >/dev/null 2>&1; then
  echo "  ✓ base Python 仍可用（$("$BASE_PY" --version 2>&1)）—— 未受影响"
else
  echo "  ⚠ base Python 异常，请检查"
fi
if [ -f "C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE" ]; then
  echo "  ✓ Microsoft Word 未受影响"
fi
echo ""
echo "完成。重建方法见 docx2epub_环境与依赖清单.md §5。"
