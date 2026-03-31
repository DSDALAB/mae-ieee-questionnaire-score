#!/usr/bin/env bash
set -euo pipefail

# 在腳本所在資料夾執行
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

VENV_DIR=".venv"
PYTHON_CMD="${PYTHON:-python3}"

if ! command -v "$PYTHON_CMD" >/dev/null 2>&1; then
  echo "找不到 $PYTHON_CMD，請先安裝 Python 3" >&2
  exit 1
fi

if [ ! -d "$VENV_DIR" ]; then
  echo "建立虛擬環境：$VENV_DIR"
  "$PYTHON_CMD" -m venv "$VENV_DIR"
fi

echo "啟用虛擬環境"
# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"

echo "升級 pip 並安裝相依（若有 requirements.txt 則使用之）"
python -m pip install --upgrade pip
if [ -f requirements.txt ]; then
  pip install -r requirements.txt
else
  pip install pandas numpy xlsxwriter openpyxl
fi

# 允許以 -f <filename> 或直接給第一個參數為檔名
if [ "${1-}" = "-f" ] && [ -n "${2-}" ]; then
  FILE="$2"
elif [ -n "${1-}" ]; then
  FILE="$1"
else
  read -rp "請輸入要處理的檔案名（直接 Enter 使用預設: 回收問卷_大學部應屆畢業生_ReWu.xlsx）：" FILE
fi

# 清理引號（若使用者輸入包含單/雙引號）
FILE_CLEAN=$(printf '%s' "$FILE" | sed "s/^['\"]\|['\"]$//g")
FILE="$FILE_CLEAN"

echo "執行 process.py 處理：$FILE"
python process.py -f "$FILE"

# 顯示預期輸出檔名（由 process.py 規則：<原始檔名>_Result.xlsx）
base_name="$(basename "$FILE")"
name_no_ext="${base_name%.*}"
out_file="${name_no_ext}_Result.xlsx"
echo "輸出檔案：$out_file"

echo "完成。"
