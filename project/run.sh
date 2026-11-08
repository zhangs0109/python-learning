#!/bin/zsh
# =============================================================
# 一键运行脚本 —— 自动激活虚拟环境，不用每次手打三行命令
#
# 用法：
#   ./run.sh                    跑全部（台账核算 + 资金计划）
#   ./run.sh settle             只跑台账核算
#   ./run.sh plan               只跑资金计划
#   ./run.sh settle 我的台账.xlsx   指定台账文件跑核算
#   ./run.sh plan   我的台账.xlsx   指定台账文件跑资金计划
# =============================================================

# 切到脚本所在目录（不管你在哪里执行都能找到文件）
cd "$(dirname "$0")" || exit 1

# 检查虚拟环境
if [ ! -d "../.venv" ]; then
  echo "✘ 找不到虚拟环境 ../.venv"
  echo "  请先在 python-learning 目录执行：python3 -m venv .venv"
  exit 1
fi

# 激活虚拟环境（这样才用得上 pandas）
source ../.venv/bin/activate

MODE="${1:-all}"

case "$MODE" in
  settle)
    shift
    python3 subcontract_settle.py "$@"
    ;;
  plan)
    shift
    python3 cash_plan.py "$@"
    ;;
  all)
    FILE="${2:-}"
    if [ -n "$FILE" ]; then
      python3 subcontract_settle.py "$FILE"
      echo
      echo "----------------------------------------"
      echo
      python3 cash_plan.py "$FILE"
    else
      python3 subcontract_settle.py
      echo
      echo "----------------------------------------"
      echo
      python3 cash_plan.py
    fi
    ;;
  *)
    echo "用法："
    echo "  ./run.sh                      跑全部"
    echo "  ./run.sh settle               只跑台账核算"
    echo "  ./run.sh plan                 只跑资金计划"
    echo "  ./run.sh settle 台账.xlsx     指定台账文件"
    exit 1
    ;;
esac

echo
echo "=============================="
echo "  全部完成 ✔"
echo "=============================="
