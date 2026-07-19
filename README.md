# 持仓跟踪表

A股持仓跟踪 Excel 工具，支持公式自动计算 + 行情自动更新。

## 文件说明

| 文件 | 说明 |
|------|------|
| `持仓跟踪表_434050.xlsx` | 主文件（含公式） |
| `generate_portfolio_excel.py` | 重新生成 Excel |
| `update_portfolio_excel.py` | **自动拉取行情并更新 Excel** |
| `portfolio_fetcher.py` | 行情数据接口（akshare） |
| `run_daily_update.sh` | Linux/Mac 一键更新 |
| `run_daily_update.bat` | Windows 一键更新 |

## 快速开始

```bash
pip install -r requirements.txt
python3 update_portfolio_excel.py
```

更新内容：
- 各标的 **现价**、**涨跌幅**
- **上证指数**
- **数据更新日志**（新增工作表）
- 自动刷新市值、盈亏、止损触发状态（Excel 打开后公式自动重算）

## 每日自动更新

### Mac / Linux（cron）

```bash
crontab -e
```

添加（每个交易日 15:30 更新）：

```
30 15 * * 1-5 cd /path/to/workspace && ./run_daily_update.sh >> update.log 2>&1
```

### Windows（任务计划程序）

1. 打开「任务计划程序」→ 创建基本任务
2. 触发器：每周一至五 15:30
3. 操作：启动程序 → `run_daily_update.bat`

## 注意事项

- 数据来源：东方财富 / akshare（免费公开接口）
- **交易时段**运行可获取盘中价；**收盘后/周末**运行获取最近交易日收盘价
- 首次更新约需 1–2 分钟（需逐个拉取个股数据）
- 跌停家数偶发拉取失败，不影响现价更新
- 本工具仅供投资辅助，不构成投资建议
