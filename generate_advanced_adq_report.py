import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


wb = openpyxl.Workbook()

title_fill = PatternFill("solid", fgColor="1A3A6B")
header_fill = PatternFill("solid", fgColor="2F4F8F")
green_fill = PatternFill("solid", fgColor="E2EFDA")
yellow_fill = PatternFill("solid", fgColor="FFF2CC")
gray_fill = PatternFill("solid", fgColor="F2F2F2")
input_fill = PatternFill("solid", fgColor="FFFDE7")
orange_fill = PatternFill("solid", fgColor="FAD7A0")
blue_fill = PatternFill("solid", fgColor="AED6F1")
pink_fill = PatternFill("solid", fgColor="FFE0E0")

white_bold = Font(bold=True, color="FFFFFF", size=11)
dark_font = Font(color="1F2D3D", size=10)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
thin = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def st(ws, cell, val=None, fill=None, font=None, align=center, border=thin, fmt=None):
    c = ws[cell]
    if val is not None:
        c.value = val
    c.fill = fill if fill else PatternFill()
    c.font = font if font else dark_font
    c.alignment = align
    c.border = border
    if fmt:
        c.number_format = fmt
    return c


sample_rows = [
    {
        "date": "2024/5/21",
        "anchor": "王雨琪",
        "start": "08:00",
        "end": "10:30",
        "duration": 2.5,
        "session_id": "20240521-王雨琪-1",
        "group_name": "投放组1",
        "channel": "ADQ广告",
        "creative": "左侧截图素材",
        "region": "待补充",
        "viewers": 91,
        "spend": 1732.02,
        "deals": 6,
        "target_cpa": 288.67,
        "product_clicks": 24,
        "submitted_orders": 8,
        "paid_orders": 6,
        "interactions": 8,
        "follows": "",
        "payment_amount": "",
        "refund_amount": "",
        "note": "截图示例数据；收入类字段待补充",
    },
    {
        "date": "2024/5/21",
        "anchor": "王雨琪",
        "start": "08:00",
        "end": "10:30",
        "duration": 2.5,
        "session_id": "20240521-王雨琪-1",
        "group_name": "投放组2",
        "channel": "ADQ广告",
        "creative": "右侧截图素材",
        "region": "待补充",
        "viewers": 361,
        "spend": 8459.93,
        "deals": 36,
        "target_cpa": 235.00,
        "product_clicks": 106,
        "submitted_orders": 43,
        "paid_orders": 35,
        "interactions": 28,
        "follows": "",
        "payment_amount": "",
        "refund_amount": "",
        "note": "截图示例数据；收入类字段待补充",
    },
]


detail_headers = [
    "日期", "主播", "开播时间", "结束时间", "直播时长(h)", "场次ID",
    "投流组名", "投流渠道", "创意名称", "主要地域",
    "观看人数", "花费(元)", "成交量", "转化目标成本(元)",
    "商品点击", "提交订单", "支付成功", "互动人数", "关注人数",
    "支付金额(元)", "退款金额(元)", "净收入(元)", "ROAS", "ROI",
    "商品点击率", "下单率", "支付率", "成交率", "互动率",
    "互动成交率", "每支付成本(元)", "流量/漏斗诊断", "备注",
]

column_widths = [
    12, 10, 10, 10, 12, 20, 12, 12, 16, 12, 10, 12, 10, 14, 10, 10, 10,
    10, 10, 12, 12, 12, 10, 10, 10, 10, 10, 10, 10, 12, 12, 22, 24,
]


ws1 = wb.active
ws1.title = "投流明细"
ws1.merge_cells("A1:AG1")
st(ws1, "A1", "ADQ高级投流分析明细（按投流组纵向追加）", fill=title_fill, font=Font(bold=True, color="FFFFFF", size=14))
ws1.merge_cells("A2:AG2")
st(
    ws1,
    "A2",
    "使用方式：后续收到截图后，每个投流组新增一行；截图里没有的收入、退款、关注等字段先留空，汇总与趋势页会自动沿用该结构。",
    fill=yellow_fill,
    font=Font(bold=True, color="7B3F00", size=10),
    align=left,
)

for idx, header in enumerate(detail_headers, start=1):
    fill = header_fill
    if header in {"支付金额(元)", "退款金额(元)", "净收入(元)", "ROAS", "ROI"}:
        fill = yellow_fill
    elif header in {"商品点击率", "下单率", "支付率", "成交率", "互动率", "互动成交率", "每支付成本(元)", "流量/漏斗诊断"}:
        fill = green_fill
    st(ws1, f"{get_column_letter(idx)}3", header, fill=fill, font=white_bold if fill == header_fill else Font(bold=True, color="375623" if fill == green_fill else "7B5900", size=10))

for idx, width in enumerate(column_widths, start=1):
    ws1.column_dimensions[get_column_letter(idx)].width = width

for row_idx, row in enumerate(sample_rows, start=4):
    base_values = [
        row["date"], row["anchor"], row["start"], row["end"], row["duration"], row["session_id"],
        row["group_name"], row["channel"], row["creative"], row["region"],
        row["viewers"], row["spend"], row["deals"], row["target_cpa"],
        row["product_clicks"], row["submitted_orders"], row["paid_orders"], row["interactions"],
        row["follows"], row["payment_amount"], row["refund_amount"],
    ]
    for col_idx, value in enumerate(base_values, start=1):
        fill = input_fill
        if row_idx == 4:
            fill = orange_fill
        elif row_idx == 5:
            fill = blue_fill
        fmt = None
        if col_idx in {5, 12, 14, 20, 21} and value != "":
            fmt = "#,##0.00"
        st(ws1, f"{get_column_letter(col_idx)}{row_idx}", value, fill=fill, fmt=fmt)

    st(ws1, f"V{row_idx}", f'=IF(OR(T{row_idx}="",U{row_idx}=""),"",T{row_idx}-U{row_idx})', fill=green_fill, fmt="#,##0.00")
    st(ws1, f"W{row_idx}", f'=IFERROR(T{row_idx}/L{row_idx},"")', fill=green_fill, fmt="0.00")
    st(ws1, f"X{row_idx}", f'=IF(OR(V{row_idx}="",L{row_idx}=""),"",(V{row_idx}-L{row_idx})/L{row_idx})', fill=green_fill, fmt="0.00%")
    st(ws1, f"Y{row_idx}", f'=IFERROR(O{row_idx}/K{row_idx},"")', fill=green_fill, fmt="0.00%")
    st(ws1, f"Z{row_idx}", f'=IFERROR(P{row_idx}/O{row_idx},"")', fill=green_fill, fmt="0.00%")
    st(ws1, f"AA{row_idx}", f'=IFERROR(Q{row_idx}/P{row_idx},"")', fill=green_fill, fmt="0.00%")
    st(ws1, f"AB{row_idx}", f'=IFERROR(M{row_idx}/K{row_idx},"")', fill=green_fill, fmt="0.00%")
    st(ws1, f"AC{row_idx}", f'=IFERROR(R{row_idx}/K{row_idx},"")', fill=green_fill, fmt="0.00%")
    st(ws1, f"AD{row_idx}", f'=IFERROR(M{row_idx}/R{row_idx},"")', fill=green_fill, fmt="0.00%")
    st(ws1, f"AE{row_idx}", f'=IFERROR(L{row_idx}/Q{row_idx},"")', fill=green_fill, fmt="#,##0.00")
    st(
        ws1,
        f"AF{row_idx}",
        f'=IF(K{row_idx}="","",IF(Y{row_idx}<0.05,"点击偏低，优先优化素材/定向",IF(AA{row_idx}<0.70,"支付承接偏弱，检查下单到支付链路",IF(AB{row_idx}<0.08,"成交偏低，复盘直播话术与商品匹配","表现稳定，继续观察"))))',
        fill=green_fill,
        align=left,
    )
    st(ws1, f"AG{row_idx}", row["note"], fill=input_fill, align=left)

ws1.freeze_panes = "A4"


ws2 = wb.create_sheet("汇总统计")
summary_headers = [
    "日期", "主播", "场次ID", "直播时长(h)", "投流组数", "总消耗(元)", "总成交量",
    "总支付成功", "总场观", "总商品点击", "总提交订单", "总互动人数", "总关注人数",
    "支付金额(元)", "退款金额(元)", "净收入(元)", "每支付成本(元)", "商品点击率",
    "下单率", "支付率", "成交率", "互动率", "互动成交率", "ROAS", "ROI", "汇总结论",
]
ws2.merge_cells("A1:Z1")
st(ws2, "A1", "ADQ高级投流分析汇总", fill=title_fill, font=Font(bold=True, color="FFFFFF", size=14))
ws2.merge_cells("A2:Z2")
st(ws2, "A2", "示例汇总已按当前截图数据预填；后续新增组数据时，只要保持【投流明细】结构一致，这里的整场汇总可继续复用。", fill=green_fill, font=Font(bold=True, color="375623", size=10), align=left)
for idx, header in enumerate(summary_headers, start=1):
    fill = header_fill
    if header in {"支付金额(元)", "退款金额(元)", "净收入(元)", "ROAS", "ROI"}:
        fill = yellow_fill
    elif header in {"每支付成本(元)", "商品点击率", "下单率", "支付率", "成交率", "互动率", "互动成交率", "汇总结论"}:
        fill = green_fill
    st(ws2, f"{get_column_letter(idx)}3", header, fill=fill, font=white_bold if fill == header_fill else Font(bold=True, color="375623" if fill == green_fill else "7B5900", size=10))

st(ws2, "A4", "2024/5/21", fill=input_fill)
st(ws2, "B4", "王雨琪", fill=input_fill)
st(ws2, "C4", "20240521-王雨琪-1", fill=input_fill)
st(ws2, "D4", 2.5, fill=input_fill, fmt="0.0")
st(ws2, "E4", '=COUNTIF(投流明细!$F:$F,$C4)', fill=green_fill)
st(ws2, "F4", '=SUMIF(投流明细!$F:$F,$C4,投流明细!$L:$L)', fill=green_fill, fmt="#,##0.00")
st(ws2, "G4", '=SUMIF(投流明细!$F:$F,$C4,投流明细!$M:$M)', fill=green_fill)
st(ws2, "H4", '=SUMIF(投流明细!$F:$F,$C4,投流明细!$Q:$Q)', fill=green_fill)
st(ws2, "I4", '=SUMIF(投流明细!$F:$F,$C4,投流明细!$K:$K)', fill=green_fill)
st(ws2, "J4", '=SUMIF(投流明细!$F:$F,$C4,投流明细!$O:$O)', fill=green_fill)
st(ws2, "K4", '=SUMIF(投流明细!$F:$F,$C4,投流明细!$P:$P)', fill=green_fill)
st(ws2, "L4", '=SUMIF(投流明细!$F:$F,$C4,投流明细!$R:$R)', fill=green_fill)
st(ws2, "M4", '=SUMIF(投流明细!$F:$F,$C4,投流明细!$S:$S)', fill=green_fill)
st(ws2, "N4", '=IF(COUNTIFS(投流明细!$F:$F,$C4,投流明细!$T:$T,"<>")=0,"",SUMIF(投流明细!$F:$F,$C4,投流明细!$T:$T))', fill=yellow_fill, fmt="#,##0.00")
st(ws2, "O4", '=IF(COUNTIFS(投流明细!$F:$F,$C4,投流明细!$U:$U,"<>")=0,"",SUMIF(投流明细!$F:$F,$C4,投流明细!$U:$U))', fill=yellow_fill, fmt="#,##0.00")
st(ws2, "P4", '=IF(OR(N4="",O4=""),"",N4-O4)', fill=yellow_fill, fmt="#,##0.00")
st(ws2, "Q4", '=IFERROR(F4/H4,"")', fill=green_fill, fmt="#,##0.00")
st(ws2, "R4", '=IFERROR(J4/I4,"")', fill=green_fill, fmt="0.00%")
st(ws2, "S4", '=IFERROR(K4/J4,"")', fill=green_fill, fmt="0.00%")
st(ws2, "T4", '=IFERROR(H4/K4,"")', fill=green_fill, fmt="0.00%")
st(ws2, "U4", '=IFERROR(G4/I4,"")', fill=green_fill, fmt="0.00%")
st(ws2, "V4", '=IFERROR(L4/I4,"")', fill=green_fill, fmt="0.00%")
st(ws2, "W4", '=IFERROR(G4/L4,"")', fill=green_fill, fmt="0.00%")
st(ws2, "X4", '=IFERROR(N4/F4,"")', fill=yellow_fill, fmt="0.00")
st(ws2, "Y4", '=IF(OR(P4="",F4=""),"",(P4-F4)/F4)', fill=yellow_fill, fmt="0.00%")
st(ws2, "Z4", '=IF(Y4="","待补充支付金额/退款后计算ROI",IF(Y4<0,"ROI为负，优先补收入口径并复盘投放质量",IF(T4<0.70,"支付率偏低，优先检查下单到支付承接","整体结构完整，可继续补充收入与复购数据")))', fill=green_fill, align=left)
for idx, width in enumerate([12, 10, 20, 12, 10, 12, 10, 12, 10, 12, 12, 12, 12, 12, 12, 12, 12, 10, 10, 10, 10, 10, 12, 10, 10, 28], start=1):
    ws2.column_dimensions[get_column_letter(idx)].width = width
ws2.freeze_panes = "A4"


ws3 = wb.create_sheet("漏斗诊断")
funnel_headers = ["场次ID", "投流组名", "场观", "商品点击", "提交订单", "支付成功", "成交量", "商品点击率", "下单率", "支付率", "成交率", "漏斗诊断"]
ws3.merge_cells("A1:L1")
st(ws3, "A1", "漏斗诊断（按投流组纵向比较）", fill=title_fill, font=Font(bold=True, color="FFFFFF", size=14))
for idx, header in enumerate(funnel_headers, start=1):
    fill = green_fill if header in {"商品点击率", "下单率", "支付率", "成交率", "漏斗诊断"} else header_fill
    st(ws3, f"{get_column_letter(idx)}2", header, fill=fill, font=white_bold if fill == header_fill else Font(bold=True, color="375623", size=10))
for detail_row, target_row in zip(range(4, 54), range(3, 53)):
    st(ws3, f"A{target_row}", f'=IF(投流明细!F{detail_row}="","",投流明细!F{detail_row})', fill=input_fill)
    st(ws3, f"B{target_row}", f'=IF(投流明细!G{detail_row}="","",投流明细!G{detail_row})', fill=input_fill)
    st(ws3, f"C{target_row}", f'=IF(A{target_row}="","",投流明细!K{detail_row})', fill=input_fill)
    st(ws3, f"D{target_row}", f'=IF(A{target_row}="","",投流明细!O{detail_row})', fill=input_fill)
    st(ws3, f"E{target_row}", f'=IF(A{target_row}="","",投流明细!P{detail_row})', fill=input_fill)
    st(ws3, f"F{target_row}", f'=IF(A{target_row}="","",投流明细!Q{detail_row})', fill=input_fill)
    st(ws3, f"G{target_row}", f'=IF(A{target_row}="","",投流明细!M{detail_row})', fill=input_fill)
    st(ws3, f"H{target_row}", f'=IF(A{target_row}="","",投流明细!Y{detail_row})', fill=green_fill, fmt="0.00%")
    st(ws3, f"I{target_row}", f'=IF(A{target_row}="","",投流明细!Z{detail_row})', fill=green_fill, fmt="0.00%")
    st(ws3, f"J{target_row}", f'=IF(A{target_row}="","",投流明细!AA{detail_row})', fill=green_fill, fmt="0.00%")
    st(ws3, f"K{target_row}", f'=IF(A{target_row}="","",投流明细!AB{detail_row})', fill=green_fill, fmt="0.00%")
    st(ws3, f"L{target_row}", f'=IF(A{target_row}="","",投流明细!AF{detail_row})', fill=green_fill, align=left)
for idx, width in enumerate([20, 12, 10, 10, 10, 10, 10, 12, 10, 10, 10, 24], start=1):
    ws3.column_dimensions[get_column_letter(idx)].width = width
ws3.freeze_panes = "A3"


ws4 = wb.create_sheet("流量质量_创意表现")
quality_headers = ["场次ID", "投流组名", "投流渠道", "创意名称", "主要地域", "花费(元)", "观看人数", "互动人数", "关注人数", "互动率", "关注率", "每支付成本(元)", "支付金额(元)", "ROAS", "质量诊断"]
ws4.merge_cells("A1:O1")
st(ws4, "A1", "流量质量 / 创意表现结构", fill=title_fill, font=Font(bold=True, color="FFFFFF", size=14))
for idx, header in enumerate(quality_headers, start=1):
    fill = green_fill if header in {"互动率", "关注率", "每支付成本(元)", "ROAS", "质量诊断"} else header_fill
    if header == "支付金额(元)":
        fill = yellow_fill
    st(ws4, f"{get_column_letter(idx)}2", header, fill=fill, font=white_bold if fill == header_fill else Font(bold=True, color="375623" if fill == green_fill else "7B5900", size=10))
for detail_row, target_row in zip(range(4, 54), range(3, 53)):
    st(ws4, f"A{target_row}", f'=IF(投流明细!F{detail_row}="","",投流明细!F{detail_row})', fill=input_fill)
    st(ws4, f"B{target_row}", f'=IF(A{target_row}="","",投流明细!G{detail_row})', fill=input_fill)
    st(ws4, f"C{target_row}", f'=IF(A{target_row}="","",投流明细!H{detail_row})', fill=input_fill)
    st(ws4, f"D{target_row}", f'=IF(A{target_row}="","",投流明细!I{detail_row})', fill=input_fill)
    st(ws4, f"E{target_row}", f'=IF(A{target_row}="","",投流明细!J{detail_row})', fill=input_fill)
    st(ws4, f"F{target_row}", f'=IF(A{target_row}="","",投流明细!L{detail_row})', fill=input_fill, fmt="#,##0.00")
    st(ws4, f"G{target_row}", f'=IF(A{target_row}="","",投流明细!K{detail_row})', fill=input_fill)
    st(ws4, f"H{target_row}", f'=IF(A{target_row}="","",投流明细!R{detail_row})', fill=input_fill)
    st(ws4, f"I{target_row}", f'=IF(A{target_row}="","",投流明细!S{detail_row})', fill=input_fill)
    st(ws4, f"J{target_row}", f'=IF(A{target_row}="","",投流明细!AC{detail_row})', fill=green_fill, fmt="0.00%")
    st(ws4, f"K{target_row}", f'=IF(OR(A{target_row}="",投流明细!S{detail_row}=""),"",IFERROR(投流明细!S{detail_row}/投流明细!K{detail_row},""))', fill=green_fill, fmt="0.00%")
    st(ws4, f"L{target_row}", f'=IF(A{target_row}="","",投流明细!AE{detail_row})', fill=green_fill, fmt="#,##0.00")
    st(ws4, f"M{target_row}", f'=IF(A{target_row}="","",投流明细!T{detail_row})', fill=yellow_fill, fmt="#,##0.00")
    st(ws4, f"N{target_row}", f'=IF(A{target_row}="","",投流明细!W{detail_row})', fill=green_fill, fmt="0.00")
    st(ws4, f"O{target_row}", f'=IF(A{target_row}="","",IF(J{target_row}<0.05,"互动偏弱，优先检查素材与人群匹配",IF(K{target_row}="","待补关注/收入字段",IF(N{target_row}<1,"当前未回本，继续观察素材质量","质量较稳，可继续放量测试"))))', fill=green_fill, align=left)
for idx, width in enumerate([20, 12, 12, 16, 12, 12, 10, 10, 10, 10, 10, 12, 12, 10, 24], start=1):
    ws4.column_dimensions[get_column_letter(idx)].width = width
ws4.freeze_panes = "A3"


ws5 = wb.create_sheet("趋势汇总")
trend_headers = ["日期", "主播", "场次数", "投流组数", "总消耗(元)", "总成交量", "总支付成功", "总场观", "总商品点击", "总提交订单", "总互动人数", "支付金额(元)", "净收入(元)", "商品点击率", "支付率", "成交率", "ROAS", "ROI"]
ws5.merge_cells("A1:R1")
st(ws5, "A1", "趋势汇总（按日期复用同一结构）", fill=title_fill, font=Font(bold=True, color="FFFFFF", size=14))
for idx, header in enumerate(trend_headers, start=1):
    fill = header_fill
    if header in {"支付金额(元)", "净收入(元)", "ROAS", "ROI"}:
        fill = yellow_fill
    elif header in {"商品点击率", "支付率", "成交率"}:
        fill = green_fill
    st(ws5, f"{get_column_letter(idx)}2", header, fill=fill, font=white_bold if fill == header_fill else Font(bold=True, color="375623" if fill == green_fill else "7B5900", size=10))
st(ws5, "A3", "2024/5/21", fill=input_fill)
st(ws5, "B3", "王雨琪", fill=input_fill)
st(ws5, "C3", '=COUNTIF(汇总统计!$A:$A,$A3)', fill=green_fill)
st(ws5, "D3", '=COUNTIF(投流明细!$A:$A,$A3)', fill=green_fill)
st(ws5, "E3", '=SUMIF(投流明细!$A:$A,$A3,投流明细!$L:$L)', fill=green_fill, fmt="#,##0.00")
st(ws5, "F3", '=SUMIF(投流明细!$A:$A,$A3,投流明细!$M:$M)', fill=green_fill)
st(ws5, "G3", '=SUMIF(投流明细!$A:$A,$A3,投流明细!$Q:$Q)', fill=green_fill)
st(ws5, "H3", '=SUMIF(投流明细!$A:$A,$A3,投流明细!$K:$K)', fill=green_fill)
st(ws5, "I3", '=SUMIF(投流明细!$A:$A,$A3,投流明细!$O:$O)', fill=green_fill)
st(ws5, "J3", '=SUMIF(投流明细!$A:$A,$A3,投流明细!$P:$P)', fill=green_fill)
st(ws5, "K3", '=SUMIF(投流明细!$A:$A,$A3,投流明细!$R:$R)', fill=green_fill)
st(ws5, "L3", '=IF(COUNTIFS(投流明细!$A:$A,$A3,投流明细!$T:$T,"<>")=0,"",SUMIF(投流明细!$A:$A,$A3,投流明细!$T:$T))', fill=yellow_fill, fmt="#,##0.00")
st(ws5, "M3", '=IF(COUNTIFS(投流明细!$A:$A,$A3,投流明细!$V:$V,"<>")=0,"",SUMIF(投流明细!$A:$A,$A3,投流明细!$V:$V))', fill=yellow_fill, fmt="#,##0.00")
st(ws5, "N3", '=IFERROR(I3/H3,"")', fill=green_fill, fmt="0.00%")
st(ws5, "O3", '=IFERROR(G3/J3,"")', fill=green_fill, fmt="0.00%")
st(ws5, "P3", '=IFERROR(F3/H3,"")', fill=green_fill, fmt="0.00%")
st(ws5, "Q3", '=IFERROR(L3/E3,"")', fill=yellow_fill, fmt="0.00")
st(ws5, "R3", '=IF(OR(M3="",E3=""),"",(M3-E3)/E3)', fill=yellow_fill, fmt="0.00%")
for idx, width in enumerate([12, 10, 10, 10, 12, 10, 12, 10, 12, 12, 12, 12, 12, 10, 10, 10, 10, 10], start=1):
    ws5.column_dimensions[get_column_letter(idx)].width = width
ws5.freeze_panes = "A3"


ws6 = wb.create_sheet("指标口径说明")
glossary_rows = [
    ("字段/指标", "定义或填写说明", "截图是否可直接获取"),
    ("投流明细", "每个投流组一行，后续新增截图时继续向下追加，不覆盖旧结构。", "是"),
    ("支付金额(元)", "成熟知识付费分析的收入口径；截图没有时先留空，后续补入。", "通常否"),
    ("退款金额(元)", "用于净收入与ROI计算；没有时先留空。", "通常否"),
    ("ROAS", "支付金额 / 花费，用于看广告回收效率。", "自动计算"),
    ("ROI", "(净收入 - 花费) / 花费，用于看真实回本情况。", "自动计算"),
    ("商品点击率", "商品点击 / 观看人数，用于诊断素材与直播间承接。", "自动计算"),
    ("下单率", "提交订单 / 商品点击，用于诊断商品页与转化环节。", "自动计算"),
    ("支付率", "支付成功 / 提交订单，用于诊断下单到支付链路。", "自动计算"),
    ("成交率", "成交量 / 观看人数，用于衡量整体投流转化效率。", "自动计算"),
    ("流量质量_创意表现", "保留渠道、创意、地域、互动、关注、收入位，方便后续补截图和复盘。", "部分可获取"),
    ("趋势汇总", "按日期复用同一结构，适合后续连续截图数据沉淀。", "自动汇总"),
]
for row_idx, row in enumerate(glossary_rows, start=1):
    for col_idx, value in enumerate(row, start=1):
        fill = header_fill if row_idx == 1 else (gray_fill if col_idx == 1 else input_fill)
        font = white_bold if row_idx == 1 else (Font(bold=True, size=10, color="1F2D3D") if col_idx == 1 else dark_font)
        st(ws6, f"{get_column_letter(col_idx)}{row_idx}", value, fill=fill, font=font, align=left)
for idx, width in enumerate([18, 60, 18], start=1):
    ws6.column_dimensions[get_column_letter(idx)].width = width
ws6.freeze_panes = "A2"


wb.save("ADQ高级投流分析_已填数据.xlsx")
print("已生成：ADQ高级投流分析_已填数据.xlsx")
