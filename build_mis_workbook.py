"""
Sales MIS Dashboard - workbook generator
=======================================
Builds a formula-driven Excel MIS (Management Information System) report for a
fictional B2B IT & office-supplies distributor, FY 2025-26 (Apr-25 to Mar-26).

* All transaction data is SYNTHETIC (seeded random, reproducible).
* Only raw inputs are hard-coded; every metric in the workbook is a live formula.

Usage:  python build_mis_workbook.py            -> Sales_MIS_Dashboard.xlsx + sales_data.csv
"""
import calendar
import csv
import datetime as dt
import random

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

SEED, N_ORDERS = 42, 600
OUT_XLSX, OUT_CSV = "Sales_MIS_Dashboard.xlsx", "sales_data.csv"
LAST = 2000  # formulas read Sales_Data rows 2..LAST so new rows are picked up

# ----------------------------------------------------------------- master data
PRODUCTS = [  # code, name, category, unit cost, unit price, pick-weight, (qty min, qty max)
    ("P101", 'Business Laptop 14"', "Computers", 46000, 62500, 14, (1, 8)),
    ("P102", "Desktop Workstation", "Computers", 33000, 45000, 10, (1, 10)),
    ("P103", '24" LED Monitor', "Computers", 8400, 11500, 14, (1, 15)),
    ("P201", "Laser Printer A4", "Printers & Peripherals", 10800, 15200, 8, (1, 5)),
    ("P202", "Document Scanner", "Printers & Peripherals", 6300, 9000, 5, (1, 4)),
    ("P203", "Wireless Keyboard & Mouse", "Printers & Peripherals", 850, 1350, 10, (5, 50)),
    ("P301", "Wi-Fi 6 Router", "Networking", 3300, 4800, 8, (2, 20)),
    ("P302", "24-Port Managed Switch", "Networking", 13000, 18500, 5, (1, 8)),
    ("P303", "Wi-Fi Access Point", "Networking", 6800, 9600, 6, (2, 15)),
    ("P401", "Ergonomic Office Chair", "Office Furniture", 5600, 8200, 8, (2, 25)),
    ("P402", "Height-Adjustable Desk", "Office Furniture", 8300, 12000, 6, (1, 15)),
    ("P403", "4-Drawer Filing Cabinet", "Office Furniture", 5900, 8400, 6, (1, 10)),
]
PEOPLE = [  # name, region, activity weight
    ("Rohan Deshmukh", "West", 1.3), ("Sneha Joshi", "West", 1.0), ("Vikram Patil", "West", 0.8),
    ("Arjun Mehra", "North", 1.2), ("Kavita Sharma", "North", 1.0), ("Rahul Verma", "North", 0.7),
    ("Ananya Iyer", "South", 1.2), ("Karthik Reddy", "South", 0.9),
    ("Debashish Roy", "East", 0.9), ("Meera Banerjee", "East", 0.7),
]
REGIONS = ["North", "South", "East", "West"]
CATEGORIES = ["Computers", "Printers & Peripherals", "Networking", "Office Furniture"]
STATUSES = ["Delivered", "Pending", "Cancelled"]
MONTH_W = [.85, .80, .80, .90, .95, 1.0, 1.25, 1.20, 1.0, .90, .95, 1.30]  # festive + FY-end peaks


def fy_months():
    out = []
    for i in range(12):
        m = 4 + i
        out.append(dt.date(2025 + (m - 1) // 12, (m - 1) % 12 + 1, 1))
    return out


def generate_orders():
    rnd = random.Random(SEED)
    months = fy_months()
    rows = []
    for _ in range(N_ORDERS):
        mi = rnd.choices(range(12), MONTH_W)[0]
        m = months[mi]
        day = rnd.randint(1, calendar.monthrange(m.year, m.month)[1])
        p = rnd.choices(PRODUCTS, [x[5] for x in PRODUCTS])[0]
        sp = rnd.choices(PEOPLE, [x[2] for x in PEOPLE])[0]
        qty = rnd.randint(*p[6])
        disc = rnd.choices([0, .02, .05, .08, .10, .12, .15], [20, 10, 25, 15, 15, 10, 5])[0]
        pend_p = 0.14 if mi >= 10 else 0.0
        status = rnd.choices(STATUSES, [1 - pend_p - .05, pend_p, .05])[0]
        od = dt.date(m.year, m.month, day)
        deliv = None
        if status == "Delivered":
            deliv = od + dt.timedelta(days=rnd.choices(range(2, 13), [8, 14, 18, 18, 14, 10, 7, 5, 3, 2, 1])[0])
        rows.append([od, sp[0], p[0], qty, disc, status, deliv])
    rows.sort(key=lambda r: r[0])
    return [[f"ORD-25-{i:04d}"] + r for i, r in enumerate(rows, 1)]


# ---------------------------------------------------------------------- styles
FONT = "Arial"
NAVY, TEAL, LIGHT, GREY, GOLD = "1F3864", "1F7A5C", "EAF0F8", "F2F2F2", "FFF2CC"
INR = '[>=10000000]"₹"#\\,##\\,##\\,##0;[>=100000]"₹"#\\,##\\,##0;"₹"#,##0'
PCT = "0.0%"
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def fnt(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def put(ws, ref, value, fmt=None, font=None, fill_=None, align=None, border=True):
    c = ws[ref]
    c.value = value
    c.font = font or fnt()
    if fmt:
        c.number_format = fmt
    if fill_:
        c.fill = fill(fill_)
    if align:
        c.alignment = Alignment(horizontal=align, vertical="center")
    if border:
        c.border = BOX
    return c


def header(ws, row, col, labels, color=NAVY):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font, c.fill, c.border = fnt(True, "FFFFFF"), fill(color), BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def title(ws, text, sub=None):
    ws["A1"].value, ws["A1"].font = text, fnt(True, NAVY, 16)
    if sub:
        ws["A2"].value, ws["A2"].font = sub, fnt(False, "595959", 9, True)
    ws.row_dimensions[1].height = 26


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def rng(col):  # Sales_Data column range as used in formulas
    return f"Sales_Data!${col}$2:${col}${LAST}"


MONTH, SP, REG, PCODE, CAT, QTY, GROSS, NET, COGS, PROFIT, STATUS, SLA = (
    rng("C"), rng("D"), rng("E"), rng("F"), rng("H"), rng("I"), rng("L"), rng("M"),
    rng("N"), rng("O"), rng("P"), rng("S"))
NOT_CXL = f'{STATUS},"<>Cancelled"'

# --------------------------------------------------------------------- workbook
wb = Workbook()
ws_readme = wb.active
ws_readme.title = "README"
ws_dash = wb.create_sheet("Dashboard")
ws_mis = wb.create_sheet("MIS_Report")
ws_sc = wb.create_sheet("Salesperson_Scorecard")
ws_data = wb.create_sheet("Sales_Data")
ws_m = wb.create_sheet("Masters")
ws_dq = wb.create_sheet("Data_Quality")
ws_calc = wb.create_sheet("Calc")
for w_, col in [(ws_readme, NAVY), (ws_dash, NAVY), (ws_mis, TEAL), (ws_sc, TEAL),
                (ws_data, "7F7F7F"), (ws_m, "7F7F7F"), (ws_dq, "C55A11"), (ws_calc, "BFBFBF")]:
    w_.sheet_properties.tabColor = col

orders = generate_orders()
months = fy_months()

# ------------------------------------------------------------------ Masters
ws = ws_m
title(ws, "Masters - lookup tables & assumptions", "Blue text = editable inputs. Sales_Data looks up everything here.")
header(ws, 4, 1, ["Product Code", "Product", "Category", "Unit Cost (₹)", "Unit Price (₹)"])
for i, p in enumerate(PRODUCTS):
    r = 5 + i
    for j, v in enumerate([p[0], p[1], p[2], p[3], p[4]]):
        put(ws, f"{'ABCDE'[j]}{r}", v, INR if j >= 3 else None, fnt(color="0000FF"))
header(ws, 4, 7, ["Salesperson", "Region", "Monthly Target (₹)"])

# monthly targets derived from the generated data (+/- spread) so achievement varies realistically
tr = random.Random(SEED + 1)
price = {p[0]: p[4] for p in PRODUCTS}
tot = {n: 0.0 for n, _, _ in PEOPLE}
for o in orders:
    if o[6] != "Cancelled":
        tot[o[2]] += o[4] * price[o[3]] * (1 - o[5])
for i, (n, reg, _) in enumerate(PEOPLE):
    r = 5 + i
    target = round(tot[n] / 12 * tr.uniform(0.88, 1.10), -4)
    put(ws, f"G{r}", n, font=fnt(color="0000FF"))
    put(ws, f"H{r}", reg, font=fnt(color="0000FF"))
    put(ws, f"I{r}", target, INR, fnt(color="0000FF"))
put(ws, "G17", "Assumptions", font=fnt(True), border=False)
put(ws, "G18", "Delivery SLA (days)")
put(ws, "H18", 7, "0", fnt(color="0000FF"), GOLD)
put(ws, "I18", "Delivered within this many days of order = on time", font=fnt(italic=True, color="595959"), border=False)
header(ws, 4, 11, ["Regions"]); header(ws, 4, 13, ["Categories"]); header(ws, 4, 15, ["Status"])
for i, v in enumerate(REGIONS): put(ws, f"K{5+i}", v, font=fnt(color="0000FF"))
for i, v in enumerate(CATEGORIES): put(ws, f"M{5+i}", v, font=fnt(color="0000FF"))
for i, v in enumerate(STATUSES): put(ws, f"O{5+i}", v, font=fnt(color="0000FF"))
widths(ws, {"A": 13, "B": 28, "C": 24, "D": 14, "E": 14, "F": 3, "G": 20, "H": 12, "I": 20, "J": 3,
            "K": 11, "L": 3, "M": 24, "N": 3, "O": 12})
ws.freeze_panes = "A5"

# --------------------------------------------------------------- Sales_Data
ws = ws_data
heads = ["Order ID", "Order Date", "Month", "Salesperson", "Region", "Product Code", "Product", "Category",
         "Qty", "Unit Price (₹)", "Discount %", "Gross Sales (₹)", "Net Sales (₹)", "COGS (₹)",
         "Gross Profit (₹)", "Status", "Delivered On", "Delivery Days", "SLA Met?"]
calc_cols = {3, 5, 7, 8, 10, 12, 13, 14, 15, 18, 19}
for i, h in enumerate(heads, 1):
    header(ws, 1, i, [h], TEAL if i in calc_cols else NAVY)
ws.row_dimensions[1].height = 32
PM = "Masters!$A$5:$A$16"
for i, o in enumerate(orders):
    r = 2 + i
    oid, od, sp, pc, qty, disc, status, deliv = o
    vals = {
        "A": oid, "B": od, "C": f"=DATE(YEAR(B{r}),MONTH(B{r}),1)", "D": sp,
        "E": f"=INDEX(Masters!$H$5:$H$14,MATCH(D{r},Masters!$G$5:$G$14,0))", "F": pc,
        "G": f"=INDEX(Masters!$B$5:$B$16,MATCH(F{r},{PM},0))",
        "H": f"=INDEX(Masters!$C$5:$C$16,MATCH(F{r},{PM},0))", "I": qty,
        "J": f"=INDEX(Masters!$E$5:$E$16,MATCH(F{r},{PM},0))", "K": disc,
        "L": f"=I{r}*J{r}", "M": f"=L{r}*(1-K{r})",
        "N": f"=I{r}*INDEX(Masters!$D$5:$D$16,MATCH(F{r},{PM},0))", "O": f"=M{r}-N{r}",
        "P": status, "Q": deliv, "R": f'=IF(Q{r}="","",Q{r}-B{r})',
        "S": f'=IF(R{r}="","",IF(R{r}<=Masters!$H$18,"Yes","No"))',
    }
    fm = {"B": "dd-mmm-yyyy", "C": "mmm-yy", "J": INR, "K": "0%", "L": INR, "M": INR, "N": INR,
          "O": INR, "Q": "dd-mmm-yyyy", "I": "#,##0"}
    for col, v in vals.items():
        c = ws[f"{col}{r}"]
        c.value, c.font = v, fnt()
        if col in fm: c.number_format = fm[col]
        if col in ("S", "P", "R", "A", "F"): c.alignment = Alignment(horizontal="center")
widths(ws, dict(zip("ABCDEFGHIJKLMNOPQRS", [13, 12, 9, 17, 9, 11, 26, 22, 7, 13, 10, 15, 15, 15, 15, 11, 12, 10, 9])))
ws.freeze_panes = "B2"
ws.auto_filter.ref = f"A1:S{1 + len(orders)}"
for col, src in (("D", "Masters!$G$5:$G$14"), ("F", PM), ("P", "Masters!$O$5:$O$7")):
    dv = DataValidation(type="list", formula1=f"={src}", allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"{col}2:{col}{LAST}")
ws.conditional_formatting.add(f"P2:P{LAST}", CellIsRule(operator="equal", formula=['"Cancelled"'], font=Font(name=FONT, color="C00000")))
ws.conditional_formatting.add(f"S2:S{LAST}", CellIsRule(operator="equal", formula=['"No"'], font=Font(name=FONT, color="C00000")))
ws["U1"].value = "Legend"; ws["U1"].font = fnt(True)
put(ws, "U2", "Navy header = raw input column", font=fnt(True, "FFFFFF"), fill_=NAVY, border=False)
put(ws, "U3", "Green header = formula column", font=fnt(True, "FFFFFF"), fill_=TEAL, border=False)
put(ws, "U4", "To add data: append rows below & copy formulas down (C,E,G,H,J,L:O,R,S)", font=fnt(italic=True, color="595959"), border=False)
ws.column_dimensions["U"].width = 60

# --------------------------------------------------------------- MIS_Report
ws = ws_mis
title(ws, "Monthly MIS Report - FY 2025-26", "All figures exclude cancelled orders unless stated. Amounts in ₹ (Indian grouping).")
ws["A3"].value, ws["A3"].font = "A. Monthly performance summary", fnt(True, TEAL, 12)
header(ws, 4, 1, ["Month", "Orders", "Units", "Gross Sales", "Discount Given", "Net Sales", "COGS", "Gross Profit",
                  "Margin %", "Target", "Achieve-ment %", "MoM Growth %", "Cancelled Orders", "Cancel %"])
for i, m in enumerate(months):
    r = 5 + i
    put(ws, f"A{r}", m, "mmm-yy", fnt(color="0000FF"), align="left")
    put(ws, f"B{r}", f'=COUNTIFS({MONTH},$A{r},{NOT_CXL})', "#,##0")
    put(ws, f"C{r}", f'=SUMIFS({QTY},{MONTH},$A{r},{NOT_CXL})', "#,##0")
    put(ws, f"D{r}", f'=SUMIFS({GROSS},{MONTH},$A{r},{NOT_CXL})', INR)
    put(ws, f"E{r}", f"=D{r}-F{r}", INR)
    put(ws, f"F{r}", f'=SUMIFS({NET},{MONTH},$A{r},{NOT_CXL})', INR)
    put(ws, f"G{r}", f'=SUMIFS({COGS},{MONTH},$A{r},{NOT_CXL})', INR)
    put(ws, f"H{r}", f"=F{r}-G{r}", INR)
    put(ws, f"I{r}", f"=IF(F{r}=0,0,H{r}/F{r})", PCT)
    put(ws, f"J{r}", "=SUM(Masters!$I$5:$I$14)", INR, fnt(color="008000"))
    put(ws, f"K{r}", f"=IF(J{r}=0,0,F{r}/J{r})", PCT)
    put(ws, f"L{r}", "n/a" if i == 0 else f"=IF(F{r-1}=0,0,F{r}/F{r-1}-1)", PCT, align="right")
    put(ws, f"M{r}", f'=COUNTIFS({MONTH},$A{r},{STATUS},"Cancelled")', "#,##0")
    put(ws, f"N{r}", f"=IF(B{r}+M{r}=0,0,M{r}/(B{r}+M{r}))", PCT)
put(ws, "A17", "FY Total", font=fnt(True), fill_=LIGHT)
for col in "BCDEFGHJM":
    put(ws, f"{col}17", f"=SUM({col}5:{col}16)", "#,##0" if col in "BCM" else INR, fnt(True), LIGHT)
put(ws, "I17", "=IF(F17=0,0,H17/F17)", PCT, fnt(True), LIGHT)
put(ws, "K17", "=IF(J17=0,0,F17/J17)", PCT, fnt(True), LIGHT)
put(ws, "L17", "", font=fnt(True), fill_=LIGHT)
put(ws, "N17", "=IF(B17+M17=0,0,M17/(B17+M17))", PCT, fnt(True), LIGHT)

green_f, amber_f, red_f = (Font(name=FONT, color="006100"), Font(name=FONT, color="9C5700"), Font(name=FONT, color="9C0006"))
def traffic(ws_, ref):
    ws_.conditional_formatting.add(ref, CellIsRule(operator="greaterThanOrEqual", formula=["1"], fill=fill("C6EFCE"), font=green_f))
    ws_.conditional_formatting.add(ref, CellIsRule(operator="between", formula=["0.9", "0.99999"], fill=fill("FFEB9C"), font=amber_f))
    ws_.conditional_formatting.add(ref, CellIsRule(operator="lessThan", formula=["0.9"], fill=fill("FFC7CE"), font=red_f))
traffic(ws, "K5:K16")
ws.conditional_formatting.add("L6:L16", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, color="C00000")))
ws.conditional_formatting.add("L6:L16", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, color="006100")))

ws["A19"].value, ws["A19"].font = "B. Net sales by region and month", fnt(True, TEAL, 12)
header(ws, 21, 1, ["Month"] + [""] * 4 + ["Total"])
for j in range(4):
    c = ws.cell(row=21, column=2 + j, value=f"=Masters!K{5+j}")
for i, m in enumerate(months):
    r = 22 + i
    put(ws, f"A{r}", f"=A{5+i}", "mmm-yy", fnt(color="008000"), align="left")
    for col in "BCDE":
        put(ws, f"{col}{r}", f'=SUMIFS({NET},{MONTH},$A{r},{REG},{col}$21,{NOT_CXL})', INR)
    put(ws, f"F{r}", f"=SUM(B{r}:E{r})", INR, fnt(True))
put(ws, "A34", "FY Total", font=fnt(True), fill_=LIGHT)
for col in "BCDEF":
    put(ws, f"{col}34", f"=SUM({col}22:{col}33)", INR, fnt(True), LIGHT)
put(ws, "A35", "Share of sales", font=fnt(True))
for col in "BCDEF":
    put(ws, f"{col}35", f"=IF($F$34=0,0,{col}34/$F$34)", PCT)
ws.conditional_formatting.add("B22:E33", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="8FB4E3"))

ws["A37"].value, ws["A37"].font = "C. Category summary", fnt(True, TEAL, 12)
header(ws, 39, 1, ["Category", "Orders", "Units", "Net Sales", "Gross Profit", "Margin %", "Share of Sales"])
for i in range(4):
    r = 40 + i
    put(ws, f"A{r}", f"=Masters!M{5+i}", font=fnt(color="008000"))
    put(ws, f"B{r}", f'=COUNTIFS({CAT},$A{r},{NOT_CXL})', "#,##0")
    put(ws, f"C{r}", f'=SUMIFS({QTY},{CAT},$A{r},{NOT_CXL})', "#,##0")
    put(ws, f"D{r}", f'=SUMIFS({NET},{CAT},$A{r},{NOT_CXL})', INR)
    put(ws, f"E{r}", f'=SUMIFS({PROFIT},{CAT},$A{r},{NOT_CXL})', INR)
    put(ws, f"F{r}", f"=IF(D{r}=0,0,E{r}/D{r})", PCT)
    put(ws, f"G{r}", f"=IF($D$44=0,0,D{r}/$D$44)", PCT)
put(ws, "A44", "Total", font=fnt(True), fill_=LIGHT)
for col in "BCDE":
    put(ws, f"{col}44", f"=SUM({col}40:{col}43)", "#,##0" if col in "BC" else INR, fnt(True), LIGHT)
put(ws, "F44", "=IF(D44=0,0,E44/D44)", PCT, fnt(True), LIGHT)
put(ws, "G44", "=SUM(G40:G43)", PCT, fnt(True), LIGHT)
widths(ws, {"A": 13, **{c: 15 for c in "BCDEFGHIJKLMN"}})
ws.freeze_panes = "B5"

# ------------------------------------------------------------ Scorecard
ws = ws_sc
title(ws, "Salesperson Scorecard - FY 2025-26", "Ranked by net sales. Performance: Exceeds >=100% of annual target, On Track >=90%, else Needs Attention.")
header(ws, 4, 1, ["Salesperson", "Region", "Orders", "Net Sales", "Gross Profit", "Margin %", "Annual Target",
                  "Achieve-ment %", "Rank", "Cancel %", "On-time Delivery %", "Performance"])
for i in range(10):
    r = 5 + i
    put(ws, f"A{r}", f"=Masters!G{r}", font=fnt(color="008000"))
    put(ws, f"B{r}", f"=Masters!H{r}", font=fnt(color="008000"))
    put(ws, f"C{r}", f'=COUNTIFS({SP},$A{r},{NOT_CXL})', "#,##0")
    put(ws, f"D{r}", f'=SUMIFS({NET},{SP},$A{r},{NOT_CXL})', INR)
    put(ws, f"E{r}", f'=SUMIFS({PROFIT},{SP},$A{r},{NOT_CXL})', INR)
    put(ws, f"F{r}", f"=IF(D{r}=0,0,E{r}/D{r})", PCT)
    put(ws, f"G{r}", f"=Masters!I{r}*12", INR, fnt(color="008000"))
    put(ws, f"H{r}", f"=IF(G{r}=0,0,D{r}/G{r})", PCT)
    put(ws, f"I{r}", f"=RANK(D{r},$D$5:$D$14)", "0", align="center")
    put(ws, f"J{r}", f'=IF(COUNTIFS({SP},$A{r})=0,0,COUNTIFS({SP},$A{r},{STATUS},"Cancelled")/COUNTIFS({SP},$A{r}))', PCT)
    put(ws, f"K{r}", f'=IF(COUNTIFS({SP},$A{r},{SLA},"Yes")+COUNTIFS({SP},$A{r},{SLA},"No")=0,0,COUNTIFS({SP},$A{r},{SLA},"Yes")/(COUNTIFS({SP},$A{r},{SLA},"Yes")+COUNTIFS({SP},$A{r},{SLA},"No")))', PCT)
    put(ws, f"L{r}", f'=IF(H{r}>=1,"Exceeds Target",IF(H{r}>=0.9,"On Track","Needs Attention"))', align="center")
put(ws, "A15", "Total", font=fnt(True), fill_=LIGHT)
put(ws, "B15", "", fill_=LIGHT)
for col in "CDEG":
    put(ws, f"{col}15", f"=SUM({col}5:{col}14)", "#,##0" if col == "C" else INR, fnt(True), LIGHT)
put(ws, "F15", "=IF(D15=0,0,E15/D15)", PCT, fnt(True), LIGHT)
put(ws, "H15", "=IF(G15=0,0,D15/G15)", PCT, fnt(True), LIGHT)
for col in "IJKL": put(ws, f"{col}15", "", fill_=LIGHT)
traffic(ws, "H5:H14")
for txt, colr in (("Exceeds Target", "C6EFCE"), ("On Track", "FFEB9C"), ("Needs Attention", "FFC7CE")):
    ws.conditional_formatting.add("L5:L14", CellIsRule(operator="equal", formula=[f'"{txt}"'], fill=fill(colr)))
widths(ws, {"A": 19, "B": 10, **{c: 15 for c in "CDEFGHIJK"}, "L": 18})
ws.freeze_panes = "B5"

# ------------------------------------------------------------ Data_Quality
ws = ws_dq
title(ws, "Data Quality & Reconciliation Checks", "Run these every time data is refreshed. Every check should read PASS before the report is circulated.")
put(ws, "B4", "Records in Sales_Data", font=fnt(True))
put(ws, "C4", f"=COUNTA({rng('A')})", "#,##0", fnt(True))
put(ws, "B5", "Overall status", font=fnt(True))
header(ws, 7, 1, ["#", "Check", "Result", "Expected", "Status"])
def col_(c): return rng(c)
checks = [
    ("Missing values in required fields (date, salesperson, product, qty, status)",
     f"=5*$C$4-(COUNTA({col_('B')})+COUNTA({col_('D')})+COUNTA({col_('F')})+COUNTA({col_('I')})+COUNTA({col_('P')}))"),
    ("Duplicate Order IDs",
     f'=SUMPRODUCT(({col_("A")}<>"")*(COUNTIF({col_("A")},{col_("A")})>1))'),
    ("Salesperson not found in Masters",
     f'=SUMPRODUCT(({col_("D")}<>"")*ISERROR(MATCH({col_("D")},Masters!$G$5:$G$14,0)))'),
    ("Product code not found in Masters",
     f'=SUMPRODUCT(({col_("F")}<>"")*ISERROR(MATCH({col_("F")},{PM},0)))'),
    ("Quantity zero or negative", f'=COUNTIF({col_("I")},"<=0")'),
    ("Discount outside 0%-30%", f'=COUNTIF({col_("K")},">0.3")+COUNTIF({col_("K")},"<0")'),
    ("Order date outside FY 2025-26",
     f'=COUNTIF({col_("B")},"<"&DATE(2025,4,1))+COUNTIF({col_("B")},">"&DATE(2026,3,31))'),
    ("Delivered orders missing a delivery date", f'=COUNTIFS({col_("P")},"Delivered",{col_("Q")},"")'),
    ("Delivery date earlier than order date",
     f'=SUMPRODUCT(({col_("Q")}<>"")*({col_("Q")}<{col_("B")}))'),
    ("Reconciliation: MIS_Report net sales vs raw data (₹ difference)",
     f'=ROUND(MIS_Report!F17-SUMIFS({NET},{NOT_CXL}),2)'),
    ("Reconciliation: region matrix vs monthly total (₹ difference)", "=ROUND(MIS_Report!F34-MIS_Report!F17,2)"),
    ("Reconciliation: scorecard vs monthly total (₹ difference)", "=ROUND(Salesperson_Scorecard!D15-MIS_Report!F17,2)"),
]
for i, (name, f_) in enumerate(checks):
    r = 8 + i
    put(ws, f"A{r}", i + 1, "0", align="center")
    put(ws, f"B{r}", name)
    put(ws, f"C{r}", f_, "#,##0.##")
    put(ws, f"D{r}", 0, "0", fnt(color="0000FF"))
    put(ws, f"E{r}", f'=IF(C{r}=D{r},"PASS","REVIEW")', font=fnt(True), align="center")
last_chk = 7 + len(checks)
put(ws, "C5", f'=IF(COUNTIF(E8:E{last_chk},"REVIEW")=0,"ALL CHECKS PASSED","ATTENTION: "&COUNTIF(E8:E{last_chk},"REVIEW")&" check(s) need review")', font=fnt(True))
ws.merge_cells("C5:E5")
ws.conditional_formatting.add(f"E8:E{last_chk}", CellIsRule(operator="equal", formula=['"PASS"'], fill=fill("C6EFCE"), font=green_f))
ws.conditional_formatting.add(f"E8:E{last_chk}", CellIsRule(operator="equal", formula=['"REVIEW"'], fill=fill("FFC7CE"), font=red_f))
widths(ws, {"A": 5, "B": 72, "C": 16, "D": 12, "E": 12})

# ---------------------------------------------------------------- Calc
ws = ws_calc
title(ws, "Calc - helper tables feeding the Dashboard charts", "Do not edit. Values respond to the region selector on the Dashboard.")
put(ws, "A3", "Region filter criterion"); put(ws, "B3", '=IF(Dashboard!$C$4="All","*",Dashboard!$C$4)')
header(ws, 4, 1, ["Month", "Net Sales (₹ Lakh)", "Target (₹ Lakh)", "Gross Profit (₹ Lakh)"])
CRIT = "Calc!$B$3"
for i in range(12):
    r = 5 + i
    put(ws, f"A{r}", f'=TEXT(MIS_Report!A{5+i},"mmm-yy")')
    put(ws, f"B{r}", f'=SUMIFS({NET},{MONTH},MIS_Report!$A{5+i},{REG},$B$3,{NOT_CXL})/100000', "#,##0.0")
    put(ws, f"C{r}", f"=SUMIF(Masters!$H$5:$H$14,$B$3,Masters!$I$5:$I$14)/100000", "#,##0.0")
    put(ws, f"D{r}", f'=SUMIFS({PROFIT},{MONTH},MIS_Report!$A{5+i},{REG},$B$3,{NOT_CXL})/100000', "#,##0.0")
header(ws, 19, 1, ["Category", "Net Sales (₹ Lakh)"])
for i in range(4):
    r = 20 + i
    put(ws, f"A{r}", f"=Masters!M{5+i}")
    put(ws, f"B{r}", f'=SUMIFS({NET},{CAT},A{r},{REG},$B$3,{NOT_CXL})/100000', "#,##0.0")
header(ws, 26, 1, ["Region", "Net Sales (₹ Lakh)", "Annual Target (₹ Lakh)"])
for i in range(4):
    r = 27 + i
    put(ws, f"A{r}", f"=Masters!K{5+i}")
    put(ws, f"B{r}", f"=MIS_Report!{'BCDE'[i]}34/100000", "#,##0.0")
    put(ws, f"C{r}", f"=SUMIF(Masters!$H$5:$H$14,A{r},Masters!$I$5:$I$14)*12/100000", "#,##0.0")
header(ws, 33, 1, ["Top 5 Salesperson", "Net Sales (₹ Lakh)", "Rank"])
for k in range(5):
    r = 34 + k
    put(ws, f"C{r}", k + 1, "0")
    put(ws, f"B{r}", f"=LARGE(Salesperson_Scorecard!$D$5:$D$14,C{r})/100000", "#,##0.0")
    put(ws, f"A{r}", f"=INDEX(Salesperson_Scorecard!$A$5:$A$14,MATCH(LARGE(Salesperson_Scorecard!$D$5:$D$14,C{r}),Salesperson_Scorecard!$D$5:$D$14,0))")
widths(ws, {"A": 26, "B": 20, "C": 22, "D": 22})

# ------------------------------------------------------------- Dashboard
ws = ws_dash
ws.sheet_view.showGridLines = False
widths(ws, {"A": 2, **{c: 13.5 for c in "BCDEFGHIJKLM"}})
ws.merge_cells("B1:M1")
put(ws, "B1", "SALES MIS DASHBOARD  |  FY 2025-26 (Apr-25 to Mar-26)", font=fnt(True, "FFFFFF", 18), fill_=NAVY, align="left", border=False)
for col in "CDEFGHIJKLM": ws[f"{col}1"].fill = fill(NAVY)
ws.row_dimensions[1].height = 38
ws.merge_cells("B2:M2")
put(ws, "B2", "Apex IT & Office Supplies (fictional company, synthetic data)  -  all amounts in ₹, cancelled orders excluded", font=fnt(False, "595959", 9, True), border=False)
put(ws, "B4", "Select region:", font=fnt(True), align="right", border=False)
put(ws, "C4", "All", font=fnt(True, "0000FF", 11), fill_=GOLD, align="center")
dv = DataValidation(type="list", formula1='"All,North,South,East,West"', allow_blank=False)
ws.add_data_validation(dv); dv.add("C4")
ws.merge_cells("D4:I4")
put(ws, "D4", "< pick All / North / South / East / West - KPIs, monthly trend and category mix update", font=fnt(False, "595959", 9, True), border=False)

def f_sum(rng_): return f'SUMIFS({rng_},{REG},{CRIT},{NOT_CXL})'
orders_f = f'COUNTIFS({REG},{CRIT},{NOT_CXL})'
cxl_f = f'COUNTIFS({REG},{CRIT},{STATUS},"Cancelled")'
yes_f, no_f = f'COUNTIFS({REG},{CRIT},{SLA},"Yes")', f'COUNTIFS({REG},{CRIT},{SLA},"No")'
cards = [
    ("NET SALES", f"={f_sum(NET)}/10000000", '"₹ "0.00" Cr"',
     '="Target ₹ "&TEXT(SUM(Calc!C5:C16)/100,"0.00")&" Cr"'),
    ("TARGET ACHIEVEMENT", "=IF(SUM(Calc!C5:C16)=0,0,SUM(Calc!B5:B16)/SUM(Calc!C5:C16))", "0.0%",
     '=IF(J7>=1,"Target exceeded","Below target")'),
    ("GROSS PROFIT", f"={f_sum(PROFIT)}/10000000", '"₹ "0.00" Cr"', '="After discounts & COGS"'),
    ("PROFIT MARGIN", f"=IF({f_sum(NET)}=0,0,{f_sum(PROFIT)}/{f_sum(NET)})", "0.0%",
     f'="Avg discount "&TEXT(IF({f_sum(GROSS)}=0,0,1-{f_sum(NET)}/{f_sum(GROSS)}),"0.0%")'),
    ("ORDERS", f"={orders_f}", "#,##0", f'=TEXT(IF({orders_f}+{cxl_f}=0,0,{cxl_f}/({orders_f}+{cxl_f})),"0.0%")&" cancelled"'),
    ("ON-TIME DELIVERY", f"=IF({yes_f}+{no_f}=0,0,{yes_f}/({yes_f}+{no_f}))", "0.0%", '="SLA: "&Masters!H18&" days"'),
]
card_cols = [("B", "C"), ("D", "E"), ("F", "G"), ("H", "I"), ("J", "K"), ("L", "M")]
for (a, b), (lab, val, fmt, sub) in zip(card_cols, cards):
    for r_ in (6, 7, 8): ws.merge_cells(f"{a}{r_}:{b}{r_}")
    for r_, col_fill in ((6, LIGHT), (7, "FFFFFF"), (8, "FFFFFF")):
        for c_ in (a, b):
            ws[f"{c_}{r_}"].fill, ws[f"{c_}{r_}"].border = fill(col_fill), BOX
    put(ws, f"{a}6", lab, font=fnt(True, "595959", 9), fill_=LIGHT, align="center")
    put(ws, f"{a}7", val, fmt, fnt(True, NAVY, 20), "FFFFFF", "center")
    put(ws, f"{a}8", sub, font=fnt(False, "595959", 9, True), fill_="FFFFFF", align="center")
ws.row_dimensions[6].height, ws.row_dimensions[7].height, ws.row_dimensions[8].height = 18, 36, 18
# fix the achievement sub-line to point at its own value cell
ws["D8"].value = '=IF(D7>=1,"Target exceeded","Below target")'
ws.conditional_formatting.add("D7", CellIsRule(operator="greaterThanOrEqual", formula=["1"], font=Font(name=FONT, bold=True, size=20, color="006100")))
ws.conditional_formatting.add("D7", CellIsRule(operator="lessThan", formula=["1"], font=Font(name=FONT, bold=True, size=20, color="C00000")))

# charts
def axes_on(ch):
    ch.x_axis.delete = False; ch.y_axis.delete = False; ch.legend.position = "b"

c1 = BarChart(); c1.type = "col"; c1.title = "Monthly Net Sales vs Target (₹ Lakh)"
c1.add_data(Reference(ws_calc, min_col=2, min_row=4, max_row=16), titles_from_data=True)
c1.set_categories(Reference(ws_calc, min_col=1, min_row=5, max_row=16))
c1.series[0].graphicalProperties.solidFill = NAVY
ln = LineChart(); ln.add_data(Reference(ws_calc, min_col=3, min_row=4, max_row=16), titles_from_data=True)
ln.series[0].graphicalProperties.line.solidFill = "C00000"; ln.series[0].graphicalProperties.line.width = 28000
ln.series[0].marker.symbol = "none"; ln.series[0].smooth = False
c1 += ln; axes_on(c1); c1.width, c1.height = 18.6, 8.6
ws.add_chart(c1, "B10")

c2 = PieChart(); c2.title = "Sales Mix by Category"
c2.add_data(Reference(ws_calc, min_col=2, min_row=19, max_row=23), titles_from_data=True)
c2.set_categories(Reference(ws_calc, min_col=1, min_row=20, max_row=23))
for idx, colr in enumerate(["1F3864", "2E75B6", "1F7A5C", "ED7D31"]):
    pt = DataPoint(idx=idx); pt.graphicalProperties.solidFill = colr; c2.series[0].dPt.append(pt)
c2.dataLabels = DataLabelList(); c2.dataLabels.showPercent = True
for a_ in ("showVal", "showCatName", "showSerName", "showLegendKey"): setattr(c2.dataLabels, a_, False)
c2.legend.position = "b"; c2.width, c2.height = 12.4, 8.6
ws.add_chart(c2, "I10")

c3 = BarChart(); c3.type = "col"; c3.title = "Net Sales vs Annual Target by Region (₹ Lakh)"
c3.add_data(Reference(ws_calc, min_col=2, max_col=3, min_row=26, max_row=30), titles_from_data=True)
c3.set_categories(Reference(ws_calc, min_col=1, min_row=27, max_row=30))
c3.series[0].graphicalProperties.solidFill = NAVY; c3.series[1].graphicalProperties.solidFill = "A6A6A6"
axes_on(c3); c3.width, c3.height = 15.2, 8.6
ws.add_chart(c3, "B28")

c4 = BarChart(); c4.type = "bar"; c4.title = "Top 5 Salespersons (₹ Lakh)"
c4.add_data(Reference(ws_calc, min_col=2, min_row=33, max_row=38), titles_from_data=True)
c4.set_categories(Reference(ws_calc, min_col=1, min_row=34, max_row=38))
c4.series[0].graphicalProperties.solidFill = TEAL; c4.x_axis.scaling.orientation = "maxMin"
c4.dataLabels = DataLabelList(); c4.dataLabels.showVal = True
for a_ in ("showCatName", "showSerName", "showLegendKey", "showPercent"): setattr(c4.dataLabels, a_, False)
axes_on(c4); c4.legend = None; c4.width, c4.height = 15.6, 8.6
ws.add_chart(c4, "H28")
ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------------------------------------------------------- README sheet
ws = ws_readme
ws.sheet_view.showGridLines = False
widths(ws, {"A": 2, "B": 26, "C": 90})
ws.merge_cells("B1:C1")
put(ws, "B1", "SALES MIS DASHBOARD - FY 2025-26", font=fnt(True, "FFFFFF", 18), fill_=NAVY, border=False)
ws["C1"].fill = fill(NAVY); ws.row_dimensions[1].height = 36
ws.merge_cells("B2:C2")
put(ws, "B2", "A formula-driven Management Information System (MIS) report built in Excel. Data is synthetic.", font=fnt(False, "595959", 10, True), border=False)
rows = [("SHEET GUIDE", None),
        ("Dashboard", "KPI cards + 4 charts with a region selector (cell C4)."),
        ("MIS_Report", "Monthly summary vs target, MoM growth, region x month matrix, category summary."),
        ("Salesperson_Scorecard", "Rank, target achievement, margin, cancellation and on-time delivery per salesperson."),
        ("Sales_Data", "600 order lines. Navy headers = raw inputs, green headers = formulas (lookups & calculations)."),
        ("Masters", "Product, salesperson, target and SLA assumptions (blue text = editable)."),
        ("Data_Quality", "12 automated validation & reconciliation checks - all should read PASS."),
        ("Calc", "Helper tables that feed the dashboard charts."),
        ("COLOUR CODE", None),
        ("Blue text", "Hard-coded input / assumption"),
        ("Black text", "Formula on the same sheet"),
        ("Green text", "Formula pulling from another sheet"),
        ("Yellow fill", "Cell you are expected to change (region selector, SLA days)"),
        ("KEY ASSUMPTIONS", None),
        ("Revenue basis", "Net Sales = Qty x Unit Price x (1 - Discount %). Cancelled orders are excluded from all sales metrics."),
        ("Targets", "Monthly target per salesperson (Masters). Annual target = monthly x 12."),
        ("On-time delivery", "Delivered within SLA days (Masters!H18) of order date. Pending orders are not counted."),
        ("Currency", "Indian Rupees (₹). Tables use lakh/crore grouping; charts are in ₹ Lakh; KPI cards in ₹ Crore."),
        ("HOW TO REFRESH", None),
        ("1", "Append new rows on Sales_Data and copy the green formula columns down."),
        ("2", "Add new months/people/products on MIS_Report / Masters if needed."),
        ("3", "Open Data_Quality - fix anything marked REVIEW before sharing the report.")]
r = 4
for k, v in rows:
    if v is None:
        ws.merge_cells(f"B{r}:C{r}")
        put(ws, f"B{r}", k, font=fnt(True, "FFFFFF"), fill_=TEAL, border=False)
    else:
        put(ws, f"B{r}", k, font=fnt(True))
        put(ws, f"C{r}", v)
        ws[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="center")
        if k in ("Dashboard", "MIS_Report", "Salesperson_Scorecard", "Sales_Data", "Masters", "Data_Quality", "Calc"):
            ws[f"B{r}"].value = f'=HYPERLINK("#{k}!A1","{k}")'
            ws[f"B{r}"].font = Font(name=FONT, bold=True, color="0563C1", underline="single")
        if k == "Blue text": ws[f"B{r}"].font = fnt(True, "0000FF")
        if k == "Green text": ws[f"B{r}"].font = fnt(True, "008000")
        if k == "Yellow fill": ws[f"B{r}"].fill = fill(GOLD)
    r += 1

for w_ in (ws_readme, ws_mis, ws_sc, ws_dq):
    w_.page_setup.orientation = "landscape"; w_.page_setup.fitToWidth = 1; w_.page_setup.fitToHeight = 0
    w_.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = 1
wb.save(OUT_XLSX)

with open(OUT_CSV, "w", newline="") as fh:
    wr = csv.writer(fh)
    wr.writerow(["Order ID", "Order Date", "Salesperson", "Product Code", "Qty", "Discount %", "Status", "Delivered On"])
    for o in orders:
        wr.writerow([o[0], o[1].isoformat(), o[2], o[3], o[4], o[5], o[6], o[7].isoformat() if o[7] else ""])
print("Saved", OUT_XLSX, "and", OUT_CSV)
