import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# 样式定义
title_fill = PatternFill("solid", fgColor="1A3A6B")
header_fill = PatternFill("solid", fgColor="2F4F8F")
green_light = PatternFill("solid", fgColor="E2EFDA")
yellow_fill = PatternFill("solid", fgColor="FFF2CC")
gray_fill = PatternFill("solid", fgColor="F2F2F2")
input_fill = PatternFill("solid", fgColor="FFFDE7")
pink_fill = PatternFill("solid", fgColor="FFE0E0")

white_bold = Font(bold=True, color="FFFFFF", size=11)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
thin = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


def st(ws, cell, val=None, fill=None, font=None, align=center, border=thin, fmt=None):
    c = ws[cell]
    if val is not None:
        c.value = val
    c.fill = fill if fill else PatternFill()
    c.font = font if font else Font(size=10, color="1F2D3D")
    c.alignment = align
    c.border = border
    if fmt:
        c.number_format = fmt
    return c


# ================================================
# Sheet 1: ADQ 原始数据录入（纵向多组）
# ================================================
ws1 = wb.active
ws1.title = "ADQ原始数据录入"

ws1.merge_cells("A1:P1")
st(ws1, "A1", "ADQ 投放数据录入表（纵向多组）", fill=title_fill, font=Font(bold=True, color="FFFFFF", size=14))
ws1.row_dimensions[1].height = 36

ws1.merge_cells("A2:P2")
st(
    ws1,
    "A2",
    "说明：每个投放组占一行，可按需新增任意组；黄色单元格填写数据，绿色单元格自动计算。",
    fill=yellow_fill,
    font=Font(bold=True, color="7B3F00", size=10),
    align=left,
)
ws1.row_dimensions[2].height = 20

ws1.merge_cells("A3:P3")
st(ws1, "A3", "基本信息", fill=header_fill, font=white_bold, align=left)
ws1.row_dimensions[3].height = 22

for col, hdr in zip(["A", "B", "C", "D", "E"], ["日期", "主播", "开播时间", "结束时间", "直播时长(h)"]):
    st(ws1, f"{col}4", hdr, fill=header_fill, font=white_bold)

for col in ["A", "B", "C", "D"]:
    st(ws1, f"{col}5", fill=input_fill)
ws1["C5"].number_format = "HH:MM"
ws1["D5"].number_format = "HH:MM"
ws1["E5"].value = "=(D5-C5)*24"
ws1["E5"].number_format = "0.0"
ws1["E5"].fill = green_light
ws1["E5"].font = Font(size=10, color="375623")
ws1["E5"].alignment = center
ws1["E5"].border = thin
ws1.row_dimensions[5].height = 24

ws1.merge_cells("A7:P7")
st(ws1, "A7", "投放组纵向对比（每行一个投放组）", fill=header_fill, font=white_bold, align=left)
ws1.row_dimensions[7].height = 26

adq_headers = [
    "投放组",
    "观看人数",
    "花费(元)",
    "成交量",
    "转化目标成本(元)",
    "商品点击",
    "提交订单",
    "支付成功",
    "互动人数",
    "成本(元/单)",
    "观看成本(元)",
    "商品点击率",
    "提交率",
    "支付率",
    "转化率",
    "互动转化率",
]

for i, field in enumerate(adq_headers):
    col = get_column_letter(i + 1)
    st(ws1, f"{col}8", field, fill=header_fill, font=white_bold)

start_row = 9
max_group_rows = 40
end_row = start_row + max_group_rows - 1

for r in range(start_row, end_row + 1):
    for col in "ABCDEFGHI":
        st(ws1, f"{col}{r}", fill=input_fill)

    st(ws1, f"J{r}", f"=IFERROR(C{r}/D{r},\"\")", fill=green_light, fmt="#,##0.00")
    st(ws1, f"K{r}", f"=IFERROR(C{r}/B{r},\"\")", fill=green_light, fmt="#,##0.00")
    st(ws1, f"L{r}", f"=IFERROR(F{r}/B{r},\"\")", fill=green_light, fmt="0.00%")
    st(ws1, f"M{r}", f"=IFERROR(G{r}/B{r},\"\")", fill=green_light, fmt="0.00%")
    st(ws1, f"N{r}", f"=IFERROR(H{r}/B{r},\"\")", fill=green_light, fmt="0.00%")
    st(ws1, f"O{r}", f"=IFERROR(D{r}/B{r},\"\")", fill=green_light, fmt="0.00%")
    st(ws1, f"P{r}", f"=IFERROR(D{r}/I{r},\"\")", fill=green_light, fmt="0.00%")

ws1.merge_cells(f"A{end_row + 2}:P{end_row + 2}")
st(
    ws1,
    f"A{end_row + 2}",
    "提示：可直接在下方插入新行并复制样式公式以继续扩展投放组。",
    fill=gray_fill,
    font=Font(italic=True, size=9, color="666666"),
    align=left,
)

col_widths1 = [12, 10, 12, 10, 14, 10, 10, 10, 10, 12, 12, 10, 10, 10, 10, 12]
for i, w in enumerate(col_widths1):
    ws1.column_dimensions[get_column_letter(i + 1)].width = w

ws1.freeze_panes = "A9"

# ================================================
# Sheet 2: 汇总统计表（全组合计）
# ================================================
ws2 = wb.create_sheet("汇总统计表")

ws2.merge_cells("A1:Q1")
st(ws2, "A1", "直播 ADQ 投放汇总统计表", fill=title_fill, font=Font(bold=True, color="FFFFFF", size=14))
ws2.row_dimensions[1].height = 36

ws2.merge_cells("A2:Q2")
st(
    ws2,
    "A2",
    "绿色单元格为自动计算，按【ADQ原始数据录入】中全部投放组自动汇总。",
    fill=green_light,
    font=Font(bold=True, color="375623", size=10),
    align=left,
)

summary_headers = [
    "日期", "主播", "时间", "直播时长(h)",
    "总消耗(元)", "总成交量", "平均成本(元/单)",
    "总场观(人)", "观看成本(元)",
    "总商品点击", "商品点击率",
    "总提交订单", "总支付成功", "未支付",
    "总互动人数", "转化率", "互动转化率",
]
for i, hdr in enumerate(summary_headers):
    col = get_column_letter(i + 1)
    st(ws2, f"{col}3", hdr, fill=header_fill, font=white_bold)

ref = "ADQ原始数据录入"
group_view_range = f"B{start_row}:B{end_row}"
group_cost_range = f"C{start_row}:C{end_row}"
group_deal_range = f"D{start_row}:D{end_row}"
group_click_range = f"F{start_row}:F{end_row}"
group_order_range = f"G{start_row}:G{end_row}"
group_paid_range = f"H{start_row}:H{end_row}"
group_interact_range = f"I{start_row}:I{end_row}"
formulas = [
    ("A", f"='{ref}'!A5", None, input_fill),
    ("B", f"='{ref}'!B5", None, input_fill),
    ("C", f"=TEXT('{ref}'!C5,\"HH:MM\")&\"-\"&TEXT('{ref}'!D5,\"HH:MM\")", None, input_fill),
    ("D", f"='{ref}'!E5", "0.0", input_fill),
    ("E", f"=SUM('{ref}'!{group_cost_range})", "#,##0.00", green_light),
    ("F", f"=SUM('{ref}'!{group_deal_range})", None, green_light),
    ("G", "=IFERROR(E4/F4,\"\")", "#,##0.00", green_light),
    ("H", f"=SUM('{ref}'!{group_view_range})", None, green_light),
    ("I", "=IFERROR(E4/H4,\"\")", "#,##0.00", green_light),
    ("J", f"=SUM('{ref}'!{group_click_range})", None, green_light),
    ("K", "=IFERROR(J4/H4,\"\")", "0.00%", green_light),
    ("L", f"=SUM('{ref}'!{group_order_range})", None, green_light),
    ("M", f"=SUM('{ref}'!{group_paid_range})", None, green_light),
    ("N", "=IFERROR(L4-M4,\"\")", None, pink_fill),
    ("O", f"=SUM('{ref}'!{group_interact_range})", None, green_light),
    ("P", "=IFERROR(F4/H4,\"\")", "0.00%", green_light),
    ("Q", "=IFERROR(F4/O4,\"\")", "0.00%", green_light),
]
for col, formula, fmt, fill in formulas:
    st(ws2, f"{col}4", formula, fill=fill, fmt=fmt)

col_widths2 = [12, 10, 16, 12, 13, 10, 13, 10, 13, 12, 12, 12, 12, 10, 12, 10, 12]
for i, w in enumerate(col_widths2):
    ws2.column_dimensions[get_column_letter(i + 1)].width = w
ws2.freeze_panes = "A4"

# ================================================
# Sheet 3: 投放组汇总（新增）
# ================================================
ws3 = wb.create_sheet("投放组汇总")

ws3.merge_cells("A1:L1")
st(ws3, "A1", "投放组对比汇总（纵向）", fill=title_fill, font=Font(bold=True, color="FFFFFF", size=14))
ws3.row_dimensions[1].height = 32

ws3.merge_cells("A2:L2")
st(ws3, "A2", "按投放组逐行展示，并自动计算占比和关键转化指标。", fill=yellow_fill, font=Font(bold=True, color="7B3F00", size=10), align=left)

group_summary_headers = [
    "投放组", "观看人数", "花费(元)", "成交量", "成本(元/单)",
    "商品点击", "提交订单", "支付成功", "互动人数", "转化率", "消耗占比", "成交占比",
]
for i, hdr in enumerate(group_summary_headers):
    col = get_column_letter(i + 1)
    st(ws3, f"{col}3", hdr, fill=header_fill, font=white_bold)

summary_start = 4
summary_end = summary_start + max_group_rows - 1
ws3["N2"] = f"=SUM(C{summary_start}:C{summary_end})"
ws3["O2"] = f"=SUM(D{summary_start}:D{summary_end})"
for idx, r in enumerate(range(summary_start, summary_end + 1)):
    src = start_row + idx
    st(ws3, f"A{r}", f"=IF('{ref}'!A{src}=\"\",\"\",'{ref}'!A{src})", fill=input_fill)
    st(ws3, f"B{r}", f"=IF($A{r}=\"\",\"\",'{ref}'!B{src})", fill=input_fill)
    st(ws3, f"C{r}", f"=IF($A{r}=\"\",\"\",'{ref}'!C{src})", fill=input_fill, fmt="#,##0.00")
    st(ws3, f"D{r}", f"=IF($A{r}=\"\",\"\",'{ref}'!D{src})", fill=input_fill)
    st(ws3, f"E{r}", f"=IFERROR(C{r}/D{r},\"\")", fill=green_light, fmt="#,##0.00")
    st(ws3, f"F{r}", f"=IF($A{r}=\"\",\"\",'{ref}'!F{src})", fill=input_fill)
    st(ws3, f"G{r}", f"=IF($A{r}=\"\",\"\",'{ref}'!G{src})", fill=input_fill)
    st(ws3, f"H{r}", f"=IF($A{r}=\"\",\"\",'{ref}'!H{src})", fill=input_fill)
    st(ws3, f"I{r}", f"=IF($A{r}=\"\",\"\",'{ref}'!I{src})", fill=input_fill)
    st(ws3, f"J{r}", f"=IFERROR(D{r}/B{r},\"\")", fill=green_light, fmt="0.00%")
    st(ws3, f"K{r}", f"=IFERROR(C{r}/$N$2,\"\")", fill=green_light, fmt="0.00%")
    st(ws3, f"L{r}", f"=IFERROR(D{r}/$O$2,\"\")", fill=green_light, fmt="0.00%")

# 总计行
total_row = summary_end + 1
st(ws3, f"A{total_row}", "合计", fill=header_fill, font=white_bold)
st(ws3, f"B{total_row}", f"=SUM(B{summary_start}:B{summary_end})", fill=green_light)
st(ws3, f"C{total_row}", f"=SUM(C{summary_start}:C{summary_end})", fill=green_light, fmt="#,##0.00")
st(ws3, f"D{total_row}", f"=SUM(D{summary_start}:D{summary_end})", fill=green_light)
st(ws3, f"E{total_row}", f"=IFERROR(C{total_row}/D{total_row},\"\")", fill=green_light, fmt="#,##0.00")
st(ws3, f"F{total_row}", f"=SUM(F{summary_start}:F{summary_end})", fill=green_light)
st(ws3, f"G{total_row}", f"=SUM(G{summary_start}:G{summary_end})", fill=green_light)
st(ws3, f"H{total_row}", f"=SUM(H{summary_start}:H{summary_end})", fill=green_light)
st(ws3, f"I{total_row}", f"=SUM(I{summary_start}:I{summary_end})", fill=green_light)
st(ws3, f"J{total_row}", f"=IFERROR(D{total_row}/B{total_row},\"\")", fill=green_light, fmt="0.00%")
st(ws3, f"K{total_row}", "=1", fill=green_light, fmt="0.00%")
st(ws3, f"L{total_row}", "=1", fill=green_light, fmt="0.00%")

col_widths3 = [12, 10, 12, 10, 12, 10, 10, 10, 10, 10, 10, 10]
for i, w in enumerate(col_widths3):
    ws3.column_dimensions[get_column_letter(i + 1)].width = w

ws3.freeze_panes = "A4"

wb.save("ADQ直播数据统计模板.xlsx")
print("已生成：ADQ直播数据统计模板.xlsx")
