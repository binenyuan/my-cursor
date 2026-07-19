#!/usr/bin/env bash
# Daily portfolio Excel auto-update script
# Usage: ./run_daily_update.sh
# Cron example (every trading day at 15:30):
#   30 15 * * 1-5 cd /path/to/workspace && ./run_daily_update.sh >> update.log 2>&1

set -euo pipefail
cd "$(dirname "$0")"

echo "===== $(date '+%Y-%m-%d %H:%M:%S') ====="
python3 update_portfolio_excel.py
echo ""
