#!/usr/bin/env python3
"""Generate portfolio tracking Excel workbook with formulas."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/持仓跟踪表_434050.xlsx"

# Styles
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
SUBHEADER_FILL = PatternFill("solid", fgColor="D6E4F0")
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
BOLD = Font(bold=True)
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
CENTER = Alignment(horizontal="center", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
STOP_FILL = PatternFill("solid", fgColor="F8CBAD")
PROFIT_FILL = PatternFill("solid", fgColor="C6EFCE")
NORMAL_FILL = PatternFill("solid", fgColor="FFFFFF")


def style_header_row(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def style_data_area(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN_BORDER
            cell.alignment = CENTER if c <= 3 else RIGHT


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_current_holdings(wb):
    ws = wb.active
    ws.title = "当前持仓"

    ws["A1"] = "A股持仓跟踪表 — 当前持仓（优化前）"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:N1")

    ws["A2"] = "基准日期：2026-07-17  |  账户总市值："
    ws["K2"] = 434050
    ws["K2"].number_format = '#,##0'
    ws["L2"] = "元"
    ws.merge_cells("A2:J2")

    headers = [
        "序号", "标的", "代码", "持股数量", "成本价", "现价",
        "市值(元)", "占比", "累计盈亏(元)", "累计盈亏%", "当日涨跌%",
        "止损价", "止盈价", "备注",
    ]
    for c, h in enumerate(headers, 1):
        ws.cell(row=4, column=c, value=h)
    style_header_row(ws, 4, len(headers))

    data = [
        (1, "科创芯片ETF", "588810", 26100, 1.905, 2.557, -8.19, 2.20, 2.80, "持有"),
        (2, "AI智能ETF", "159819", 32500, 1.500, 1.820, -8.08, 1.60, 1.95, "持有"),
        (3, "东山精密", "002384", 200, 249.537, 241.920, -10.00, 225.00, 260.00, "持有"),
        (4, "立讯精密", "002475", 800, 69.398, 58.050, -5.73, 52.00, 68.00, "持有"),
        (5, "景旺电子", "603228", 600, 79.882, 71.700, 1.29, 66.00, 80.00, "持有"),
        (6, "生益科技", "600183", 300, 185.479, 132.290, -10.00, 120.00, 160.00, "持有"),
        (7, "通信ETF", "515880", 59700, 0.506, 0.659, -9.97, 0.60, 0.72, "持有"),
        (8, "亨通光电", "600487", 500, 109.457, 58.390, -9.96, 0, 0, "建议清仓"),
        (9, "平安电工", "001359", 300, 128.923, 96.190, -10.00, 0, 0, "建议清仓"),
        (10, "通富微电", "002156", 400, 78.211, 68.360, -9.99, 62.00, 82.00, "持有"),
        (11, "福晶科技", "002222", 100, 76.930, 60.050, -8.92, 0, 0, "建议清仓"),
    ]

    start = 5
    for i, row in enumerate(data):
        r = start + i
        ws.cell(r, 1, row[0])
        ws.cell(r, 2, row[1])
        ws.cell(r, 3, row[2])
        ws.cell(r, 4, row[3])
        ws.cell(r, 5, row[4])
        ws.cell(r, 6, row[5])  # 现价 - user updates daily
        # 市值 = 数量 * 现价
        ws.cell(r, 7, f"=D{r}*F{r}")
        # 占比
        ws.cell(r, 8, f"=G{r}/$K$2")
        ws.cell(r, 8).number_format = "0.0%"
        # 累计盈亏
        ws.cell(r, 9, f"=(F{r}-E{r})*D{r}")
        ws.cell(r, 9).number_format = '#,##0'
        # 累计盈亏%
        ws.cell(r, 10, f"=IF(E{r}=0,0,(F{r}-E{r})/E{r})")
        ws.cell(r, 10).number_format = "0.0%"
        ws.cell(r, 11, row[6] / 100)
        ws.cell(r, 11).number_format = "0.00%"
        ws.cell(r, 12, row[7] if row[7] else "")
        ws.cell(r, 13, row[8] if row[8] else "")
        ws.cell(r, 14, row[9])

    total_row = start + len(data)
    ws.cell(total_row, 2, "合计")
    ws.cell(total_row, 2).font = BOLD
    ws.cell(total_row, 7, f"=SUM(G{start}:G{total_row - 1})")
    ws.cell(total_row, 7).number_format = '#,##0'
    ws.cell(total_row, 8, f"=SUM(H{start}:H{total_row - 1})")
    ws.cell(total_row, 8).number_format = "0.0%"
    ws.cell(total_row, 9, f"=SUM(I{start}:I{total_row - 1})")
    ws.cell(total_row, 9).number_format = '#,##0'
    ws.cell(total_row, 10, f"=IF(G{total_row}-I{total_row}=0,0,I{total_row}/(G{total_row}-I{total_row}))")
    ws.cell(total_row, 10).number_format = "0.0%"

    # Number formats
    for r in range(start, total_row):
        ws.cell(r, 4).number_format = '#,##0'
        ws.cell(r, 5).number_format = "0.000"
        ws.cell(r, 6).number_format = "0.000"
        ws.cell(r, 7).number_format = '#,##0'
        if ws.cell(r, 12).value:
            ws.cell(r, 12).number_format = "0.00"
        if ws.cell(r, 13).value:
            ws.cell(r, 13).number_format = "0.00"

    style_data_area(ws, start, total_row, 1, 14)
    set_col_widths(ws, [5, 14, 10, 10, 10, 10, 12, 8, 14, 10, 10, 10, 10, 12])

    ws["A16"] = "说明：每日只需更新 F 列「现价」，其余列自动计算。"
    ws["A16"].font = Font(italic=True, color="666666")


def build_conservative(wb):
    ws = wb.create_sheet("稳健版目标持仓")

    ws["A1"] = "稳健版 — 目标持仓（优化后）"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:O1")

    ws["A2"] = "账户总市值："
    ws["B2"] = 434050
    ws["B2"].number_format = '#,##0'

    headers = [
        "序号", "标的", "代码", "目标数量", "成本价", "现价",
        "市值(元)", "目标占比", "实际占比", "累计盈亏(元)", "累计盈亏%",
        "止损价", "止盈价", "距止损%", "触发状态", "止损/止盈动作",
    ]
    for c, h in enumerate(headers, 1):
        ws.cell(4, c, value=h)
    style_header_row(ws, 4, len(headers))

    # target_qty, cost, stop, profit, action_stop, action_profit, note
    data = [
        (1, "科创芯片ETF", "588810", 14355, 1.905, 2.557, 0.085, 2.20, 2.80, "跌破→清仓", "涨破→减至5%"),
        (2, "AI智能ETF", "159819", 16250, 1.500, 1.820, 0.068, 1.60, 1.95, "跌破→清仓", "涨破→减至4%"),
        (3, "通信ETF", "515880", 26865, 0.506, 0.659, 0.041, 0.60, 0.72, "跌破→清仓", "涨破→减至3%"),
        (4, "景旺电子", "603228", 600, 79.882, 71.700, 0.099, 66.00, 80.00, "跌破→清仓", "涨破→减50%"),
        (5, "立讯精密", "002475", 800, 69.398, 58.050, 0.107, 52.00, 68.00, "跌破→清仓", "涨破→减50%"),
        (6, "东山精密", "002384", 100, 249.537, 241.920, 0.056, 225.00, 260.00, "跌破→清仓", "回本→清仓"),
        (7, "红利ETF", "510880", 16500, 3.15, 3.15, 0.120, 2.90, 3.50, "跌破→观望", "涨破→持有"),
        (8, "沪深300ETF", "510300", 10750, 4.00, 4.00, 0.099, 3.70, 4.40, "跌破→观望", "涨破→持有"),
        (9, "货币基金", "—", 0, 1.00, 1.00, 0.282, "", "", "—", "企稳后分批买入"),
    ]

    start = 5
    cash_amount = 122525
    for i, row in enumerate(data):
        r = start + i
        ws.cell(r, 1, row[0])
        ws.cell(r, 2, row[1])
        ws.cell(r, 3, row[2])
        ws.cell(r, 4, row[3])
        ws.cell(r, 5, row[4])
        ws.cell(r, 6, row[5])
        if row[1] == "货币基金":
            ws.cell(r, 7, cash_amount)
            ws.cell(r, 8, row[6])
            ws.cell(r, 9, f"=G{r}/$B$2")
        else:
            ws.cell(r, 7, f"=D{r}*F{r}")
            ws.cell(r, 8, row[6])
            ws.cell(r, 9, f"=G{r}/$B$2")
        ws.cell(r, 8).number_format = "0.0%"
        ws.cell(r, 9).number_format = "0.0%"
        if row[1] != "货币基金":
            ws.cell(r, 10, f"=(F{r}-E{r})*D{r}")
            ws.cell(r, 11, f"=IF(E{r}=0,0,(F{r}-E{r})/E{r})")
        else:
            ws.cell(r, 10, 0)
            ws.cell(r, 11, 0)
        ws.cell(r, 10).number_format = '#,##0'
        ws.cell(r, 11).number_format = "0.0%"
        ws.cell(r, 12, row[7] if row[7] != "" else "")
        ws.cell(r, 13, row[8] if row[8] != "" else "")
        # 距止损%
        if row[7] != "":
            ws.cell(r, 14, f"=IF(L{r}=\"\",\"\",(F{r}-L{r})/L{r})")
            ws.cell(r, 14).number_format = "0.0%"
            # 触发状态
            ws.cell(r, 15, f'=IF(L{r}="","—",IF(F{r}<=L{r},"🔴止损",IF(F{r}>=M{r},"🟢止盈","⚪正常")))')
        else:
            ws.cell(r, 14, "—")
            ws.cell(r, 15, "—")
        ws.cell(r, 16, f'{row[9]} / {row[10]}')

    total_row = start + len(data)
    ws.cell(total_row, 2, "合计")
    ws.cell(total_row, 2).font = BOLD
    ws.cell(total_row, 7, f"=SUM(G{start}:G{total_row - 1})")
    ws.cell(total_row, 7).number_format = '#,##0'
    ws.cell(total_row, 9, f"=SUM(I{start}:I{total_row - 1})")
    ws.cell(total_row, 9).number_format = "0.0%"
    ws.cell(total_row, 10, f"=SUM(J{start}:J{total_row - 1})")
    ws.cell(total_row, 10).number_format = '#,##0'

    for r in range(start, total_row):
        ws.cell(r, 4).number_format = '#,##0'
        ws.cell(r, 5).number_format = "0.000"
        ws.cell(r, 6).number_format = "0.000"
        ws.cell(r, 7).number_format = '#,##0'
        if ws.cell(r, 12).value and ws.cell(r, 12).value != "":
            ws.cell(r, 12).number_format = "0.00"
        if ws.cell(r, 13).value and ws.cell(r, 13).value != "":
            ws.cell(r, 13).number_format = "0.00"

    style_data_area(ws, start, total_row, 1, 16)
    set_col_widths(ws, [5, 14, 10, 10, 10, 10, 12, 10, 10, 14, 10, 10, 10, 10, 10, 18])

    # Rebalancing section
    r0 = total_row + 3
    ws.cell(r0, 1, "调仓操作清单").font = TITLE_FONT
    rebalance_headers = ["操作", "标的", "代码", "数量", "预估金额(元)", "备注"]
    for c, h in enumerate(rebalance_headers, 1):
        ws.cell(r0 + 1, c, value=h)
    style_header_row(ws, r0 + 1, len(rebalance_headers))

    ops = [
        ("清仓", "亨通光电", "600487", "500股", 29195, "止损 -47%"),
        ("清仓", "福晶科技", "002222", "100股", 6005, "简化组合"),
        ("清仓", "平安电工", "001359", "300股", 28857, "PCB冗余"),
        ("清仓", "生益科技", "600183", "300股", 39687, "成本过高"),
        ("清仓", "通富微电", "002156", "400股", 27344, "用ETF替代"),
        ("减仓50%", "东山精密", "002384", "100股", 24192, "保留100股"),
        ("减仓45%", "科创芯片ETF", "588810", "11745份", 30032, "锁定利润"),
        ("减仓50%", "AI智能ETF", "159819", "16250份", 29575, "锁定利润"),
        ("减仓55%", "通信ETF", "515880", "32835份", 21638, "情绪最弱"),
        ("买入", "红利ETF", "510880", "16500份", 51975, "防御仓位12%"),
        ("买入", "沪深300ETF", "510300", "10750份", 43000, "防御仓位10%"),
        ("持币", "货币基金", "—", "—", 122525, "机动储备28%"),
    ]
    for i, op in enumerate(ops):
        r = r0 + 2 + i
        for c, v in enumerate(op, 1):
            ws.cell(r, c, v)
            ws.cell(r, c).border = THIN_BORDER
        ws.cell(r, 5).number_format = '#,##0'

    sell_row = r0 + 2 + len(ops) + 1
    ws.cell(sell_row, 1, "卖出回笼合计")
    ws.cell(sell_row, 1).font = BOLD
    ws.cell(sell_row, 5, f"=SUMIF(A{r0 + 2}:A{sell_row - 2},\"清仓\",E{r0 + 2}:E{sell_row - 2})+SUMIF(A{r0 + 2}:A{sell_row - 2},\"减仓50%\",E{r0 + 2}:E{sell_row - 2})+SUMIF(A{r0 + 2}:A{sell_row - 2},\"减仓45%\",E{r0 + 2}:E{sell_row - 2})+SUMIF(A{r0 + 2}:A{sell_row - 2},\"减仓55%\",E{r0 + 2}:E{sell_row - 2})")
    ws.cell(sell_row, 5).number_format = '#,##0'


def build_aggressive(wb):
    ws = wb.create_sheet("激进版目标持仓")

    ws["A1"] = "激进版 — 目标持仓（优化后）"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:O1")

    ws["A2"] = "账户总市值："
    ws["B2"] = 434050
    ws["B2"].number_format = '#,##0'

    headers = [
        "序号", "标的", "代码", "目标数量", "成本价", "现价",
        "市值(元)", "目标占比", "实际占比", "累计盈亏(元)", "累计盈亏%",
        "止损价", "止盈价", "距止损%", "触发状态", "止损/止盈动作",
    ]
    for c, h in enumerate(headers, 1):
        ws.cell(4, c, value=h)
    style_header_row(ws, 4, len(headers))

    data = [
        (1, "科创芯片ETF", "588810", 20880, 1.905, 2.557, 0.123, 2.30, 2.90, "跌破→减20%", "涨破→再减20%"),
        (2, "AI智能ETF", "159819", 27625, 1.500, 1.820, 0.116, 1.65, 2.00, "跌破→减20%", "涨破→再减20%"),
        (3, "通信ETF", "515880", 44775, 0.506, 0.659, 0.068, 0.62, 0.75, "跌破→减30%", "涨破→减30%"),
        (4, "景旺电子", "603228", 600, 79.882, 71.700, 0.099, 68.00, 85.00, "跌破→减50%", "涨破→减30%"),
        (5, "东山精密", "002384", 200, 249.537, 241.920, 0.112, 230.00, 270.00, "跌破→清仓", "涨破→减30%"),
        (6, "立讯精密", "002475", 800, 69.398, 58.050, 0.107, 54.00, 72.00, "跌破→减50%", "涨破→减30%"),
        (7, "通富微电", "002156", 400, 78.211, 68.360, 0.063, 62.00, 82.00, "跌破→清仓", "涨破→减30%"),
        (8, "生益科技", "600183", 150, 185.479, 132.290, 0.046, 120.00, 160.00, "跌破→清仓", "回本→清仓"),
        (9, "红利ETF", "510880", 7000, 3.15, 3.15, 0.051, 2.90, 3.50, "跌破→观望", "涨破→持有"),
        (10, "沪深300ETF", "510300", 3250, 4.00, 4.00, 0.030, 3.70, 4.40, "跌破→观望", "涨破→持有"),
        (11, "货币基金", "—", 0, 1.00, 1.00, 0.187, "", "", "—", "企稳后回补ETF"),
    ]

    start = 5
    cash_amount = 80956
    for i, row in enumerate(data):
        r = start + i
        ws.cell(r, 1, row[0])
        ws.cell(r, 2, row[1])
        ws.cell(r, 3, row[2])
        ws.cell(r, 4, row[3])
        ws.cell(r, 5, row[4])
        ws.cell(r, 6, row[5])
        if row[1] == "货币基金":
            ws.cell(r, 7, cash_amount)
            ws.cell(r, 8, row[6])
            ws.cell(r, 9, f"=G{r}/$B$2")
        else:
            ws.cell(r, 7, f"=D{r}*F{r}")
            ws.cell(r, 8, row[6])
            ws.cell(r, 9, f"=G{r}/$B$2")
        ws.cell(r, 8).number_format = "0.0%"
        ws.cell(r, 9).number_format = "0.0%"
        if row[1] != "货币基金":
            ws.cell(r, 10, f"=(F{r}-E{r})*D{r}")
            ws.cell(r, 11, f"=IF(E{r}=0,0,(F{r}-E{r})/E{r})")
        else:
            ws.cell(r, 10, 0)
            ws.cell(r, 11, 0)
        ws.cell(r, 10).number_format = '#,##0'
        ws.cell(r, 11).number_format = "0.0%"
        ws.cell(r, 12, row[7] if row[7] != "" else "")
        ws.cell(r, 13, row[8] if row[8] != "" else "")
        if row[7] != "":
            ws.cell(r, 14, f"=IF(L{r}=\"\",\"\",(F{r}-L{r})/L{r})")
            ws.cell(r, 14).number_format = "0.0%"
            ws.cell(r, 15, f'=IF(L{r}="","—",IF(F{r}<=L{r},"🔴止损",IF(F{r}>=M{r},"🟢止盈","⚪正常")))')
        else:
            ws.cell(r, 14, "—")
            ws.cell(r, 15, "—")
        ws.cell(r, 16, f'{row[9]} / {row[10]}')

    total_row = start + len(data)
    ws.cell(total_row, 2, "合计")
    ws.cell(total_row, 2).font = BOLD
    ws.cell(total_row, 7, f"=SUM(G{start}:G{total_row - 1})")
    ws.cell(total_row, 7).number_format = '#,##0'
    ws.cell(total_row, 9, f"=SUM(I{start}:I{total_row - 1})")
    ws.cell(total_row, 9).number_format = "0.0%"
    ws.cell(total_row, 10, f"=SUM(J{start}:J{total_row - 1})")
    ws.cell(total_row, 10).number_format = '#,##0'

    for r in range(start, total_row):
        ws.cell(r, 4).number_format = '#,##0'
        ws.cell(r, 5).number_format = "0.000"
        ws.cell(r, 6).number_format = "0.000"
        ws.cell(r, 7).number_format = '#,##0'

    style_data_area(ws, start, total_row, 1, 16)
    set_col_widths(ws, [5, 14, 10, 10, 10, 10, 12, 10, 10, 14, 10, 10, 10, 10, 10, 18])

    r0 = total_row + 3
    ws.cell(r0, 1, "调仓操作清单").font = TITLE_FONT
    rebalance_headers = ["操作", "标的", "代码", "数量", "预估金额(元)", "备注"]
    for c, h in enumerate(rebalance_headers, 1):
        ws.cell(r0 + 1, c, value=h)
    style_header_row(ws, r0 + 1, len(rebalance_headers))

    ops = [
        ("清仓", "亨通光电", "600487", "500股", 29195, "止损 -47%"),
        ("清仓", "福晶科技", "002222", "100股", 6005, "简化组合"),
        ("清仓", "平安电工", "001359", "300股", 28857, "PCB冗余"),
        ("减仓50%", "生益科技", "600183", "150股", 19844, "保留150股"),
        ("减仓20%", "科创芯片ETF", "588810", "5220份", 13348, "小幅止盈"),
        ("减仓15%", "AI智能ETF", "159819", "4875份", 8873, "小幅止盈"),
        ("减仓25%", "通信ETF", "515880", "14925份", 9836, "情绪弱"),
        ("买入", "红利ETF", "510880", "7000份", 22050, "防御仓位5%"),
        ("买入", "沪深300ETF", "510300", "3250份", 13000, "防御仓位3%"),
        ("持币", "货币基金", "—", "—", 80956, "机动储备19%"),
    ]
    for i, op in enumerate(ops):
        r = r0 + 2 + i
        for c, v in enumerate(op, 1):
            ws.cell(r, c, v)
            ws.cell(r, c).border = THIN_BORDER
        ws.cell(r, 5).number_format = '#,##0'


def build_daily_monitor(wb):
    ws = wb.create_sheet("每日监控")

    ws["A1"] = "每日监控表 — 更新现价后自动判断"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:J1")

    ws["A2"] = "选择方案："
    ws["B2"] = "稳健版"
    ws["C2"] = "（可改为：激进版）"
    ws["A3"] = "监控日期："
    ws["B3"] = "2026-07-17"
    ws["A4"] = "沪指收盘："
    ws["B4"] = 3764.15
    ws["A5"] = "跌停家数："
    ws["B5"] = 198

    headers = [
        "标的", "代码", "持股数量", "成本价", "现价", "止损价", "止盈价",
        "距止损%", "距止盈%", "触发状态", "建议操作",
    ]
    for c, h in enumerate(headers, 1):
        ws.cell(7, c, value=h)
    style_header_row(ws, 7, len(headers))

    # Conservative holdings for monitoring
    holdings = [
        ("科创芯片ETF", "588810", 14355, 1.905, 2.557, 2.20, 2.80, "跌破→清仓", "涨破→减至5%"),
        ("AI智能ETF", "159819", 16250, 1.500, 1.820, 1.60, 1.95, "跌破→清仓", "涨破→减至4%"),
        ("通信ETF", "515880", 26865, 0.506, 0.659, 0.60, 0.72, "跌破→清仓", "涨破→减至3%"),
        ("景旺电子", "603228", 600, 79.882, 71.700, 66.00, 80.00, "跌破→清仓", "涨破→减50%"),
        ("立讯精密", "002475", 800, 69.398, 58.050, 52.00, 68.00, "跌破→清仓", "涨破→减50%"),
        ("东山精密", "002384", 100, 249.537, 241.920, 225.00, 260.00, "跌破→清仓", "回本→清仓"),
        ("红利ETF", "510880", 16500, 3.15, 3.15, 2.90, 3.50, "跌破→观望", "涨破→持有"),
        ("沪深300ETF", "510300", 10750, 4.00, 4.00, 3.70, 4.40, "跌破→观望", "涨破→持有"),
    ]

    start = 8
    for i, row in enumerate(holdings):
        r = start + i
        ws.cell(r, 1, row[0])
        ws.cell(r, 2, row[1])
        ws.cell(r, 3, row[2])
        ws.cell(r, 4, row[3])
        ws.cell(r, 5, row[4])  # 现价 - update daily
        ws.cell(r, 6, row[5])
        ws.cell(r, 7, row[6])
        # 距止损%
        ws.cell(r, 8, f"=IF(F{r}=0,\"\",(E{r}-F{r})/F{r})")
        ws.cell(r, 8).number_format = "0.0%"
        # 距止盈%
        ws.cell(r, 9, f"=IF(G{r}=0,\"\",(G{r}-E{r})/E{r})")
        ws.cell(r, 9).number_format = "0.0%"
        # 触发状态
        ws.cell(r, 10, f'=IF(E{r}<=F{r},"🔴止损",IF(E{r}>=G{r},"🟢止盈",IF(H{r}<0.03,"🟡预警","⚪正常")))')
        # 建议操作
        ws.cell(r, 11, f'=IF(J{r}="🔴止损","执行：{row[7]}",IF(J{r}="🟢止盈","执行：{row[8]}",IF(J{r}="🟡预警","准备操作，密切关注","持有不动")))')

    style_data_area(ws, start, start + len(holdings) - 1, 1, 11)
    set_col_widths(ws, [14, 10, 10, 10, 10, 10, 10, 10, 10, 10, 18])

    # Market conditions for add position
    r0 = start + len(holdings) + 2
    ws.cell(r0, 1, "稳健版加仓条件检查").font = TITLE_FONT
    cond_headers = ["条件", "标准", "当前值", "是否满足"]
    for c, h in enumerate(cond_headers, 1):
        ws.cell(r0 + 1, c, value=h)
    style_header_row(ws, r0 + 1, 4)

    ws.cell(r0 + 2, 1, "① 沪指站稳")
    ws.cell(r0 + 2, 2, "连续3日 > 3850")
    ws.cell(r0 + 2, 3, "=B4")
    ws.cell(r0 + 2, 4, '=IF(B4>3850,"✅","❌")')

    ws.cell(r0 + 3, 1, "② 跌停减少")
    ws.cell(r0 + 3, 2, "连续2日 < 30家")
    ws.cell(r0 + 3, 3, "=B5")
    ws.cell(r0 + 3, 4, '=IF(B5<30,"✅","❌")')

    ws.cell(r0 + 4, 1, "综合判断")
    ws.cell(r0 + 4, 2, "全部满足才加仓")
    ws.cell(r0 + 4, 4, f'=IF(D{r0 + 2}="✅",IF(D{r0 + 3}="✅","✅ 可以加仓","❌ 暂不加仓"),"❌ 暂不加仓")')
    ws.cell(r0 + 4, 4).font = BOLD

    for r in range(r0 + 2, r0 + 5):
        for c in range(1, 5):
            ws.cell(r, c).border = THIN_BORDER


def build_scenario(wb):
    ws = wb.create_sheet("情景压力测试")

    ws["A1"] = "情景压力测试"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:F1")

    ws["A2"] = "当前账户总市值："
    ws["B2"] = 434050
    ws["B2"].number_format = '#,##0'

    headers = ["情景", "大盘涨跌", "稳健版账户涨跌", "稳健版盈亏(元)", "激进版账户涨跌", "激进版盈亏(元)"]
    for c, h in enumerate(headers, 1):
        ws.cell(4, c, value=h)
    style_header_row(ws, 4, 6)

    scenarios = [
        ("极端下跌", -0.15, -0.075, -0.115),
        ("继续下跌", -0.10, -0.055, -0.085),
        ("震荡", 0.03, 0.015, 0.025),
        ("超跌反弹", 0.10, 0.055, 0.085),
        ("强势修复", 0.15, 0.080, 0.125),
    ]

    start = 5
    for i, (name, market, cons, aggr) in enumerate(scenarios):
        r = start + i
        ws.cell(r, 1, name)
        ws.cell(r, 2, market)
        ws.cell(r, 2).number_format = "0.0%"
        ws.cell(r, 3, cons)
        ws.cell(r, 3).number_format = "0.0%"
        ws.cell(r, 4, f"=$B$2*C{r}")
        ws.cell(r, 4).number_format = '#,##0'
        ws.cell(r, 5, aggr)
        ws.cell(r, 5).number_format = "0.0%"
        ws.cell(r, 6, f"=$B$2*E{r}")
        ws.cell(r, 6).number_format = '#,##0'

    style_data_area(ws, start, start + len(scenarios) - 1, 1, 6)
    set_col_widths(ws, [12, 10, 16, 16, 16, 16])

    r0 = start + len(scenarios) + 2
    ws.cell(r0, 1, "自定义情景测算").font = TITLE_FONT
    ws.cell(r0 + 1, 1, "输入大盘涨跌（如 -0.1 表示跌10%）：")
    ws.cell(r0 + 1, 3, -0.10)
    ws.cell(r0 + 1, 3).number_format = "0.0%"

    ws.cell(r0 + 2, 1, "稳健版预估盈亏：")
    ws.cell(r0 + 2, 3, f"=$B$2*C{r0 + 1}*0.55")
    ws.cell(r0 + 2, 3).number_format = '#,##0'
    ws.cell(r0 + 2, 4, "（科技暴露约45%，贝塔调整系数0.55）")

    ws.cell(r0 + 3, 1, "激进版预估盈亏：")
    ws.cell(r0 + 3, 3, f"=$B$2*C{r0 + 1}*0.85")
    ws.cell(r0 + 3, 3).number_format = '#,##0'
    ws.cell(r0 + 3, 4, "（科技暴露约73%，贝塔调整系数0.85）")


def build_readme(wb):
    ws = wb.create_sheet("使用说明")
    ws.column_dimensions["A"].width = 80

    lines = [
        ("A股持仓跟踪表 — 使用说明", TITLE_FONT),
        ("", None),
        ("【文件结构】", BOLD),
        ("1. 当前持仓：优化前的原始持仓，每日更新 F 列「现价」", None),
        ("2. 稳健版目标持仓：优化后的目标仓位 + 调仓清单 + 止损止盈", None),
        ("3. 激进版目标持仓：优化后的目标仓位 + 调仓清单 + 止损止盈", None),
        ("4. 每日监控：收盘后更新现价，自动判断止损/止盈/预警", None),
        ("5. 情景压力测试：不同市场情景下的盈亏预估", None),
        ("", None),
        ("【每日操作流程】", BOLD),
        ("1. 收盘后打开「每日监控」表", None),
        ("2. 更新 E 列各标的「现价」", None),
        ("3. 更新沪指收盘（B4）和跌停家数（B5）", None),
        ("4. 查看 J 列「触发状态」：", None),
        ("   🔴止损 → 次日开盘执行止损操作", None),
        ("   🟢止盈 → 次日开盘执行止盈操作", None),
        ("   🟡预警 → 距止损不足3%，密切关注", None),
        ("   ⚪正常 → 持有不动", None),
        ("", None),
        ("【调仓执行顺序】", BOLD),
        ("第1步：清仓亨通光电、福晶科技、平安电工（两版共同）", None),
        ("第2步：按选定版本减仓 ETF 和个股", None),
        ("第3步：买入红利ETF(510880)和沪深300ETF(510300)", None),
        ("第4步：剩余资金转入货币基金", None),
        ("", None),
        ("【重要提醒】", BOLD),
        ("• 本表仅为投资辅助工具，不构成投资建议", None),
        ("• 止损价触发后请严格执行，不要抱幻想", None),
        ("• 加仓须满足「每日监控」表底部条件检查", None),
        ("• 红利ETF和沪深300ETF买入价按估算填入，请以实际成交价更新成本", None),
        ("", None),
        ("基准日期：2026-07-17  |  账户总市值：434,050 元", None),
    ]

    for i, (text, font) in enumerate(lines, 1):
        ws.cell(i, 1, text)
        if font:
            ws.cell(i, 1).font = font


def main():
    wb = Workbook()
    build_current_holdings(wb)
    build_conservative(wb)
    build_aggressive(wb)
    build_daily_monitor(wb)
    build_scenario(wb)
    build_readme(wb)
    wb.save(OUTPUT_PATH)
    print(f"已生成: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
