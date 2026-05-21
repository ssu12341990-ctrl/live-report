import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── 颜色/样式定义 ─────────────────────────────────────
title_fill   = PatternFill("solid", fgColor="1A3A6B")
header_fill  = PatternFill("solid", fgColor="2F4F8F")
orange_fill  = PatternFill("solid", fgColor="F4A623")  # ADQ组1
blue_fill    = PatternFill("solid", fgColor="4472C4")  # ADQ组2
green_fill   = PatternFill("solid", fgColor="375623")  # 汇总
green_light  = PatternFill("solid", fgColor="E2EFDA")
yellow_fill  = PatternFill("solid", fgColor="FFF2CC")
gray_fill    = PatternFill("solid", fgColor="F2F2F2")
input_fill   = PatternFill("solid", fgColor="FFFDE7")  # 输入区域

white_bold  = Font(bold=True, color="FFFFFF", size=11)
dark_bold   = Font(bold=True, color="1F2D3D", size=10)
dark_normal = Font(color="1F2D3D", size=10)
red_bold    = Font(bold=True, color="C00000", size=10)
orange_bold = Font(bold=True, color="7B3F00", size=10)
green_bold  = Font(bold=True, color="FFFFFF", size=10)

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
thin   = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)
medium = Border(
    left=Side(style="medium"), right=Side(style="medium"),
    top=Side(style="medium"),  bottom=Side(style="medium"),
)

def s(ws, cell, val=None, fill=None, font=None, align=center, border=thin, fmt=None, bold=False):
    c = ws[cell]
    if val is not None: c.value = val
    if fill: c.fill = fill
    if font: c.font = font
    elif bold: c.font = Font(bold=True, size=10, color="1F2D3D")
    else: c.font = Font(size=10, color="1F2D3D")
    c.alignment = align
    c.border = border
    if fmt: c.number_format = fmt
    return c

# ════════════════════════════════════════════════════════
# Sheet 1：ADQ 原始数据录入
# ════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "ADQ原始数据录入"

# 标题
ws1.merge_cells("A1:P1")
s(ws1, "A1", "ADQ 投放数据录入表（每次直播后填入截图数据）", fill=title_fill, font=white_bold)
ws1.row_dimensions[1].height = 36

# 说明行
ws1.merge_cells("A2:P2")
s(ws1, "A2", "说明：黄色单元格为手动填入区域，蓝色单元格为自动计算区域，无需填写",
  fill=yellow_fill, font=Font(bold=True, color="7B3F00", size=10), align=left)
ws1.row_dimensions[2].height = 20

# 基本信息
ws1.merge_cells("A3:P3")
s(ws1, "A3", "【基本信息】", fill=header_fill, font=white_bold, align=left)
ws1.row_dimensions[3].height = 22

basic_headers = ["日期", "主播", "开播时间", "结束时间", "直播时长(h)"]
basic_cols    = ["A", "B", "C", "D", "E"]
for col, hdr in zip(basic_cols, basic_headers):
    s(ws1, f"{col}4", hdr, fill=header_fill, font=white_bold)
ws1.row_dimensions[4].height = 28

# 基本信息输入行（第5行）
for col in basic_cols:
    s(ws1, f"{col}5", fill=input_fill)
ws1.row_dimensions[5].height = 24

# 直播时长自动计算 = 结束-开始（如果用时间格式）
ws1["E5"].value = "=(D5-C5)*24"
ws1["E5"].number_format = "0.0"
ws1["E5"].fill = green_light
ws1["E5"].font = Font(size=10, color="375623")

# ADQ两组数据标题
ws1.merge_cells("A7:H7")
s(ws1, "A7", "投放组 1 数据（从ADQ截图填入）",
  fill=PatternFill("solid", fgColor="F4A623"), font=white_bold)
ws1.merge_cells("I7:P7")
s(ws1, "I7", "投放组 2 数据（从ADQ截图填入）",
  fill=PatternFill("solid", fgColor="4472C4"), font=white_bold)
ws1.row_dimensions[7].height = 28

# 两组共同字段
adq_fields = [
    ("观看人数",       "人",    None),
    ("花费(元)",       "元",    "#,##0.00"),
    ("成交量",         "单",    None),
    ("转化目标成本",   "元",    "#,##0.00"),
    ("商品点击",       "次",    None),
    ("提交订单",       "单",    None),
    ("支付成功",       "单",    None),
    ("互动人数",       "人",    None),
]
group1_cols = ["A","B","C","D","E","F","G","H"]
group2_cols = ["I","J","K","L","M","N","O","P"]

for col, (field, unit, fmt) in zip(group1_cols, adq_fields):
    label = f"{field}\n({unit})"
    s(ws1, f"{col}8", label,
      fill=PatternFill("solid", fgColor="FAD7A0"), font=Font(bold=True, size=9, color="7B3F00"))
    s(ws1, f"{col}9", fill=input_fill, fmt=fmt)

for col, (field, unit, fmt) in zip(group2_cols, adq_fields):
    label = f"{field}\n({unit})"
    s(ws1, f"{col}8", label,
      fill=PatternFill("solid", fgColor="AED6F1"), font=Font(bold=True, size=9, color="1A5276"))
    s(ws1, f"{col}9", fill=input_fill, fmt=fmt)

ws1.row_dimensions[8].height = 36
ws1.row_dimensions[9].height = 28

# 提示行
ws1.merge_cells("A10:P10")
s(ws1, "A10", "互动人数 = 评论人数（ADQ截图中的互动数据，如只有合计则填合计）",
  fill=gray_fill, font=Font(italic=True, size=9, color="666666"), align=left)
ws1.row_dimensions[10].height = 18

# 列宽
for col in "ABCDEFGHIJKLMNOP":
    ws1.column_dimensions[col].width = 14

# ════════════════════════════════════════════════════════
# Sheet 2：汇总统计表
# ════════════════════════════════════════════════════════
ws2 = wb.create_sheet("汇总统计表")

# 标题
ws2.merge_cells("A1:Q1")
s(ws2, "A1", "直播 ADQ 投放汇总统计表",
  fill=title_fill, font=Font(bold=True, color="FFFFFF", size=14))
ws2.row_dimensions[1].height = 36

# 说明
ws2.merge_cells("A2:Q2")
s(ws2, "A2", "本表数据自动从【ADQ原始数据录入】汇总计算，无需手动填写",
  fill=green_light, font=Font(bold=True, color="375623", size=10), align=left)
ws2.row_dimensions[2].height = 20

# 列头定义
summary_headers = [
    "日期", "主播", "时间", "直播时长(h)",
    "消耗(元)", "成交量", "成本(元/单)",
    "场观(人)", "观看成本(元)",
    "商品点击", "商品点击率",
    "提交订单", "支付成功", "未支付",
    "互动人数", "转化率", "互动转化率"
]
header_cols = [get_column_letter(i+1) for i in range(len(summary_headers))]
for col_letter, hdr in zip(header_cols, summary_headers):
    s(ws2, f"{col_letter}3", hdr, fill=header_fill, font=white_bold)
ws2.row_dimensions[3].height = 36

# 第4行：自动汇总公式（引用Sheet1）
ref = "ADQ原始数据录入"

formulas = {
    "A": f"='{ref}'!A5",                                                     # 日期
    "B": f"='{ref}'!B5",                                                     # 主播
    "C": f"=TEXT('{ref}'!C5,\"HH:MM\")&\"-\"&TEXT('{ref}'!D5,\"HH:MM\")",   # 时间
    "D": f"='{ref}'!E5",                                                     # 直播时长
    "E": f"='{ref}'!B9+'{ref}'!J9",                                          # 消耗合计
    "F": f"='{ref}'!C9+'{ref}'!K9",                                          # 成交量合计
    "G": f"=IFERROR(E4/F4,\"\")",                                            # 成本
    "H": f"='{ref}'!A9+'{ref}'!I9",                                          # 场观合计
    "I": f"=IFERROR(E4/H4,\"\")",                                            # 观看成本
    "J": f"='{ref}'!E9+'{ref}'!M9",                                          # 商品点击合计
    "K": f"=IFERROR(J4/H4,\"\")",                                            # 商品点击率
    "L": f"='{ref}'!F9+'{ref}'!N9",                                          # 提交订单合计
    "M": f"='{ref}'!G9+'{ref}'!O9",                                          # 支付成功合计
    "N": f"=IFERROR(L4-M4,\"\")",                                            # 未支付
    "O": f"='{ref}'!H9+'{ref}'!P9",                                          # 互动合计
    "P": f"=IFERROR(F4/H4,\"\")",                                            # 转化率
    "Q": f"=IFERROR(F4/O4,\"\")",                                            # 互动转化率
}

fmts = {
    "A": "YYYY/M/D", "D": "0.0",
    "E": "#,##0.00", "G": "#,##0.00", "I": "#,##0.00",
    "K": "0.00%", "P": "0.00%", "Q": "0.00%",
}
yellow_cols = {"E","F","G","H","I","J","K","L","M","N","O","P","Q"}

for col_letter, formula in formulas.items():
    fill = green_light if col_letter in yellow_cols else input_fill
    s(ws2, f"{col_letter}4", formula, fill=fill, fmt=fmts.get(col_letter))
ws2.row_dimensions[4].height = 24

# 字段说明区
ws2.merge_cells("A6:Q6")
s(ws2, "A6", "字段说明", fill=header_fill, font=white_bold, align=left)
ws2.row_dimensions[6].height = 22

legend_rows = [
    ("消耗(元)",      "组1花费 + 组2花费",                              gray_fill),
    ("成交量",        "组1成交量 + 组2成交量",                          gray_fill),
    ("成本(元/单)",   "消耗 / 成交量，即每单获客成本",                   yellow_fill),
    ("场观(人)",      "组1观看人数 + 组2观看人数",                       gray_fill),
    ("观看成本(元)",  "消��� / 场观人数",                                yellow_fill),
    ("商品点击率",    "商品点击 / 场观人数",                             yellow_fill),
    ("未支付",        "提交订单 - 支付成功",                             yellow_fill),
    ("转化率",        "成交量 / 场观人数，反映整体转化效率",              green_light),
    ("互动转化率",    "成交量 / 互动人数，反映互动后成交效率",            green_light),
]
for j, (term, desc, bg) in enumerate(legend_rows):
    row_n = 7 + j
    ws2.merge_cells(f"A{row_n}:D{row_n}")
    ws2.merge_cells(f"E{row_n}:Q{row_n}")
    c1 = ws2[f"A{row_n}"]
    c1.value = term; c1.font = Font(bold=True, size=10); c1.fill = bg
    c1.alignment = left; c1.border = thin
    c2 = ws2[f"E{row_n}"]
    c2.value = desc; c2.font = Font(size=10, color="444444"); c2.fill = bg
    c2.alignment = left; c2.border = thin
    ws2.row_dimensions[row_n].height = 20

# 列宽
col_widths2 = [12,10,16,12,13,10,13,10,13,10,12,10,10,10,10,10,12]
for i, w in enumerate(col_widths2):
    ws2.column_dimensions[get_column_letter(i+1)].width = w

ws2.freeze_panes = "A4"

# ════════════════════════════════════════════════════════
# Sheet 3：示例数据（本场）
# ════════════════════════════════════════════════════════
ws3 = wb.create_sheet("示例-2024年5月21日")

ws3.merge_cells("A1:Q1")
s(ws3, "A1", "示例数据：2024/5/21 王雨琪 8:00-10:30（已根据ADQ截图填入）",
  fill=PatternFill("solid", fgColor="E67E22"), font=white_bold)
ws3.row_dimensions[1].height = 30

example_headers = summary_headers
for col_letter, hdr in zip(header_cols, example_headers):
    s(ws3, f"{col_letter}2", hdr, fill=header_fill, font=white_bold)
ws3.row_dimensions[2].height = 36

# 示例数据（从截图读取）
# 组1：观看91，花费1732.02，成交6，商品点击24，提交订单8，支付成功6，互动8
# 组2：观看361，花费8459.93，成交36，商品点击106，提交订单43，支付成功35，互动28
example = {
    "A": "2024/5/21",
    "B": "王雨琪",
    "C": "8:00-10:30",
    "D": 2.5,
    "E": 1732.02 + 8459.93,    # 10191.95
    "F": 6 + 36,                # 42
    "G": (1732.02+8459.93)/(6+36),   # 242.67
    "H": 91 + 361,              # 452
    "I": (1732.02+8459.93)/(91+361), # 22.55
    "J": 24 + 106,              # 130
    "K": (24+106)/(91+361),     # 商品点击率
    "L": 8 + 43,                # 51 提交订单
    "M": 6 + 35,                # 41 支付成功
    "N": (8+43)-(6+35),         # 10 未支付
    "O": 8 + 28,                # 36 互动
    "P": (6+36)/(91+361),       # 转化率
    "Q": (6+36)/(8+28),         # 互动转化率
}
ex_fmts = {
    "A": "YYYY/M/D", "D": "0.0",
    "E": "#,##0.00", "G": "#,##0.00", "I": "#,##0.00",
    "K": "0.00%", "P": "0.00%", "Q": "0.00%",
}
for col_letter, val in example.items():
    fill = green_light if col_letter in yellow_cols else input_fill
    s(ws3, f"{col_letter}3", val, fill=fill, fmt=ex_fmts.get(col_letter))
ws3.row_dimensions[3].height = 24

for i, w in enumerate(col_widths2):
    ws3.column_dimensions[get_column_letter(i+1)].width = w
ws3.freeze_panes = "A3"

wb.save("ADQ直播数据统计模板.xlsx")
print("已生成：ADQ直播数据统计模板.xlsx")
