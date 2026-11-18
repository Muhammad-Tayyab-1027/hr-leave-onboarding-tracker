"""Builds hr_leave_onboarding.xlsx: annual leave balances, absence summary and a new-starter checklist.

All employees and requests are made up (fixed random seed).
"""
import random
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

random.seed(2026)
FONT = "Arial"
HEAD = PatternFill("solid", fgColor="1F3864")
INPUT = PatternFill("solid", fgColor="FFF2CC")
RED = PatternFill("solid", bgColor="F8CBAD")
AMBER = PatternFill("solid", bgColor="FFE699")
GREEN = PatternFill("solid", bgColor="C6E0B4")

# Irish public holidays 2026 (source: citizensinformation.ie, Public holidays in Ireland)
HOLIDAYS = [(date(2026, 1, 1), "New Year's Day"), (date(2026, 2, 2), "St Brigid's Day"),
            (date(2026, 3, 17), "St Patrick's Day"), (date(2026, 4, 6), "Easter Monday"),
            (date(2026, 5, 4), "May public holiday"), (date(2026, 6, 1), "June public holiday"),
            (date(2026, 8, 3), "August public holiday"), (date(2026, 10, 26), "October public holiday"),
            (date(2026, 12, 25), "Christmas Day"), (date(2026, 12, 26), "St Stephen's Day")]


def header(ws, names, widths, row=1):
    for i, name in enumerate(names, 1):
        c = ws.cell(row=row, column=i, value=name)
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[col(i)].width = widths[i - 1]
    ws.row_dimensions[row].height = 32
    ws.freeze_panes = ws.cell(row=row + 1, column=1).coordinate


wb = Workbook()
g = wb.active
g.title = "Guide"
g.column_dimensions["A"].width = 110
for i, t in enumerate([
    "HR Leave and Onboarding Tracker", "",
    "Sample data only: every employee and request is made up. No real personal data is stored.", "",
    "Settings: report date and the public holiday list used to count working days.",
    "Employees: staff list with department, start date and annual leave entitlement (days).",
    "Leave Requests: yellow cells are typed in. Working days are counted with NETWORKDAYS, skipping weekends and public holidays.",
    "Leave Balances: approved annual leave taken, pending requests and days remaining for each employee.",
    "Absence Summary: approved days off by department and leave type, with a chart.",
    "Onboarding: checklist for new starters with due dates, completion and overdue flags.", "",
    "Entitlement is entered per employee. Irish law gives most full-time staff at least 4 working weeks a year;",
    "check citizensinformation.ie for the exact rules before using this for real staff.",
], 1):
    g.cell(row=i, column=1, value=t)
g["A1"].font = Font(name=FONT, bold=True, size=14)

# ---------- Settings ----------
st = wb.create_sheet("Settings")
st.column_dimensions["A"].width = 26
st.column_dimensions["B"].width = 26
REPORT = date(2026, 10, 30)
st["A1"], st["B1"] = "Report date", REPORT
st["B1"].number_format = "yyyy-mm-dd"
st["B1"].fill = INPUT
st["B1"].font = Font(name=FONT, color="0000FF")
st["A3"], st["B3"] = "Public holiday", "Name"
for c in ("A3", "B3"):
    st[c].font = Font(name=FONT, bold=True, color="FFFFFF")
    st[c].fill = HEAD
for r, (d, name) in enumerate(HOLIDAYS, 4):
    st[f"A{r}"], st[f"B{r}"] = d, name
    st[f"A{r}"].number_format = "yyyy-mm-dd"
HOL = f"Settings!$A$4:$A${3 + len(HOLIDAYS)}"
st[f"A{5 + len(HOLIDAYS)}"] = "Source: citizensinformation.ie, Public holidays in Ireland."
REPORT_C = "Settings!$B$1"

# ---------- Employees ----------
em = wb.create_sheet("Employees")
header(em, ["Employee ID", "Name", "Department", "Start Date", "Hours per Week", "Annual Leave Entitlement (days)"],
       [12, 22, 16, 12, 10, 14])
first = ["Aoife", "Conor", "Niamh", "Sean", "Maria", "Piotr", "Fatima", "James", "Anna", "Liam", "Grace", "Omar",
         "Ciara", "Darragh", "Sofia", "Eoin", "Hannah", "Tomas"]
last = ["Byrne", "Kelly", "Murphy", "Walsh", "Silva", "Nowak", "Khan", "Ryan", "Doyle", "Nolan", "Ali", "Brennan",
        "Farrell", "Costa", "Quinn", "Lynch", "Kowalski", "Daly"]
depts = ["Finance", "Sales", "Operations", "Customer Care", "IT"]
emps = []
for i, (f, l) in enumerate(zip(first, last), 1):
    hours = random.choice([39, 39, 39, 39, 20, 30])
    ent = 22 if hours == 39 and random.random() < 0.6 else (20 if hours == 39 else round(20 * hours / 39))
    start = [date(2026, 10, 5), date(2026, 10, 19), date(2026, 11, 9)][i - 16] if i > 15 else \
        date(random.randint(2018, 2025), random.randint(1, 12), random.randint(1, 28))
    emps.append(dict(id=f"E{100 + i}", name=f"{f} {l}", dept=random.choice(depts), start=start, hours=hours, ent=ent))
for r, e in enumerate(emps, 2):
    em[f"A{r}"], em[f"B{r}"], em[f"C{r}"], em[f"D{r}"], em[f"E{r}"], em[f"F{r}"] = \
        e["id"], e["name"], e["dept"], e["start"], e["hours"], e["ent"]
    em[f"D{r}"].number_format = "yyyy-mm-dd"
    em[f"F{r}"].fill = INPUT
EN = len(emps) + 1
E_ID, E_NAME, E_DEPT = (f"Employees!${c}$2:${c}${EN}" for c in "ABC")

# ---------- Leave requests ----------
lr = wb.create_sheet("Leave Requests")
header(lr, ["Request ID", "Employee ID", "Employee", "Department", "Leave Type", "First Day", "Last Day",
            "Working Days", "Status", "Month"], [11, 11, 22, 15, 13, 12, 12, 10, 11, 10])
types = ["Annual", "Annual", "Annual", "Annual", "Sick", "Unpaid", "Bereavement"]
reqs = []
current = [e for e in emps if e["start"] <= date(2026, 1, 1)]
for e in current:
    for _ in range(random.randint(2, 5)):
        t = random.choice(types)
        d = date(2026, 1, 5) + timedelta(days=random.randrange(320))
        while d.weekday() >= 5:
            d += timedelta(days=1)
        length = random.choice([1, 1, 2, 3, 5, 5, 10]) if t == "Annual" else random.choice([1, 1, 2, 3])
        end = d + timedelta(days=length - 1)
        status = "Approved" if d <= REPORT else random.choice(["Approved", "Pending", "Pending", "Rejected"])
        reqs.append((e, t, d, end, status))
reqs.sort(key=lambda x: x[2])
for r, (e, t, d, end, status) in enumerate(reqs, 2):
    lr[f"A{r}"] = f"LR{3000 + r - 1}"
    lr[f"B{r}"] = e["id"]
    lr[f"C{r}"] = f'=IFERROR(INDEX({E_NAME},MATCH(B{r},{E_ID},0)),"Unknown ID")'
    lr[f"D{r}"] = f'=IFERROR(INDEX({E_DEPT},MATCH(B{r},{E_ID},0)),"")'
    lr[f"E{r}"], lr[f"F{r}"], lr[f"G{r}"], lr[f"I{r}"] = t, d, end, status
    lr[f"H{r}"] = f'=IF(G{r}<F{r},"Check dates",NETWORKDAYS(F{r},G{r},{HOL}))'
    lr[f"J{r}"] = f'=TEXT(F{r},"yyyy-mm")'
    lr[f"F{r}"].number_format = lr[f"G{r}"].number_format = "yyyy-mm-dd"
    for c in "BEFGI":
        lr[f"{c}{r}"].fill = INPUT
LN = len(reqs) + 1
LIM = LN + 300
for rng, values in ((f"E2:E{LIM}", '"Annual,Sick,Unpaid,Bereavement,Other"'),
                    (f"I2:I{LIM}", '"Approved,Pending,Rejected"')):
    v = DataValidation(type="list", formula1=values, allow_blank=False)
    lr.add_data_validation(v)
    v.add(rng)
v = DataValidation(type="list", formula1=f"Employees!$A$2:$A${EN}", allow_blank=False,
                   showErrorMessage=True, error="Pick an Employee ID from the list.")
lr.add_data_validation(v)
v.add(f"B2:B{LIM}")
lr.auto_filter.ref = f"A1:J{LN}"
lr.conditional_formatting.add(f"I2:I{LN}", CellIsRule(operator="equal", formula=['"Pending"'], fill=AMBER))
lr.conditional_formatting.add(f"I2:I{LN}", CellIsRule(operator="equal", formula=['"Rejected"'], fill=RED))
L_EMP, L_DEPT, L_TYPE, L_DAYS, L_STATUS, L_FIRST = (f"'Leave Requests'!${c}$2:${c}${LIM}" for c in "BDEHIF")

# ---------- Leave balances ----------
lb = wb.create_sheet("Leave Balances")
header(lb, ["Employee ID", "Name", "Department", "Entitlement", "Annual Taken (approved, to date)",
            "Annual Booked (approved, future)", "Annual Pending", "Remaining", "Sick Days", "Flag"],
       [12, 22, 15, 11, 14, 14, 11, 10, 9, 26])
for r in range(2, EN + 1):
    lb[f"A{r}"] = f"=Employees!A{r}"
    lb[f"B{r}"] = f"=Employees!B{r}"
    lb[f"C{r}"] = f"=Employees!C{r}"
    lb[f"D{r}"] = f"=Employees!F{r}"
    lb[f"E{r}"] = f'=SUMIFS({L_DAYS},{L_EMP},A{r},{L_TYPE},"Annual",{L_STATUS},"Approved",{L_FIRST},"<="&{REPORT_C})'
    lb[f"F{r}"] = f'=SUMIFS({L_DAYS},{L_EMP},A{r},{L_TYPE},"Annual",{L_STATUS},"Approved",{L_FIRST},">"&{REPORT_C})'
    lb[f"G{r}"] = f'=SUMIFS({L_DAYS},{L_EMP},A{r},{L_TYPE},"Annual",{L_STATUS},"Pending")'
    lb[f"H{r}"] = f"=D{r}-E{r}-F{r}"
    lb[f"I{r}"] = f'=SUMIFS({L_DAYS},{L_EMP},A{r},{L_TYPE},"Sick",{L_STATUS},"Approved")'
    lb[f"J{r}"] = (f'=IF(H{r}<0,"Over entitlement",IF(G{r}>H{r},"Pending exceeds balance",'
                   f'IF(H{r}>=10,"Large balance left","")))')
lb.conditional_formatting.add(f"J2:J{EN}", FormulaRule(formula=['OR($J2="Over entitlement",$J2="Pending exceeds balance")'], fill=RED))
lb.conditional_formatting.add(f"J2:J{EN}", CellIsRule(operator="equal", formula=['"Large balance left"'], fill=AMBER))

# ---------- Absence summary ----------
ab = wb.create_sheet("Absence Summary")
ab["A1"] = "Approved days off by department (sample data)"
ab["A1"].font = Font(name=FONT, bold=True, size=14)
leave_types = ["Annual", "Sick", "Unpaid", "Bereavement"]
header(ab, ["Department"] + leave_types + ["Total"], [16, 10, 10, 10, 12, 10], row=3)
for r, d in enumerate(depts, 4):
    ab[f"A{r}"] = d
    for i, t in enumerate(leave_types):
        c = col(2 + i)
        ab[f"{c}{r}"] = f'=SUMIFS({L_DAYS},{L_DEPT},$A{r},{L_TYPE},"{t}",{L_STATUS},"Approved")'
    ab[f"F{r}"] = f"=SUM(B{r}:E{r})"
end = 3 + len(depts)
ab[f"A{end + 1}"] = "Total"
for c in "BCDEF":
    ab[f"{c}{end + 1}"] = f"=SUM({c}4:{c}{end})"
    ab[f"{c}{end + 1}"].font = Font(name=FONT, bold=True)
ab[f"A{end + 1}"].font = Font(name=FONT, bold=True)
ch = BarChart()
ch.type = "bar"
ch.grouping = "stacked"
ch.overlap = 100
ch.title = "Approved days off by department"
ch.add_data(Reference(ab, min_col=2, max_col=5, min_row=3, max_row=end), titles_from_data=True)
ch.set_categories(Reference(ab, min_col=1, min_row=4, max_row=end))
for s, colour in zip(ch.series, ["1F3864", "C55A11", "7F7F7F", "548235"]):
    s.graphicalProperties.solidFill = colour
ch.height, ch.width = 8, 16
ab.add_chart(ch, "H3")

# ---------- Onboarding ----------
ob = wb.create_sheet("Onboarding")
header(ob, ["Employee ID", "Name", "Start Date", "Task", "Owner", "Due Date", "Done (Y/N)", "Status"],
       [12, 22, 12, 34, 10, 12, 10, 12])
tasks = [("Contract signed and returned", "HR", -10), ("Payroll and bank details received", "HR", -5),
         ("Laptop and accounts set up", "IT", -2), ("Building access card issued", "Facilities", 0),
         ("Induction and policies walkthrough", "HR", 0), ("Data protection (GDPR) training", "HR", 7),
         ("Health and safety briefing", "Facilities", 7), ("30-day check-in", "Manager", 30)]
new = [e for e in emps if e["start"] > date(2026, 1, 1)]
r = 2
for e in new:
    for t, owner, offset in tasks:
        due = e["start"] + timedelta(days=offset)
        ob[f"A{r}"] = e["id"]
        ob[f"B{r}"] = f'=IFERROR(INDEX({E_NAME},MATCH(A{r},{E_ID},0)),"Unknown ID")'
        ob[f"C{r}"] = e["start"]
        ob[f"D{r}"], ob[f"E{r}"], ob[f"F{r}"] = t, owner, due
        ob[f"G{r}"] = "Y" if due < REPORT - timedelta(days=3) and random.random() < 0.8 else "N"
        ob[f"H{r}"] = f'=IF(G{r}="Y","Done",IF(F{r}<{REPORT_C},"Overdue","Open"))'
        ob[f"C{r}"].number_format = ob[f"F{r}"].number_format = "yyyy-mm-dd"
        ob[f"G{r}"].fill = INPUT
        r += 1
ON = r - 1
v = DataValidation(type="list", formula1='"Y,N"', allow_blank=False)
ob.add_data_validation(v)
v.add(f"G2:G{ON + 200}")
ob.conditional_formatting.add(f"H2:H{ON}", CellIsRule(operator="equal", formula=['"Done"'], fill=GREEN))
ob.conditional_formatting.add(f"H2:H{ON}", CellIsRule(operator="equal", formula=['"Overdue"'], fill=RED))
ob.auto_filter.ref = f"A1:H{ON}"
ob.column_dimensions["J"].width = 22
ob["J1"], ob["K1"], ob["L1"] = "New Starter", "Tasks Done", "% Complete"
for c in ("J1", "K1", "L1"):
    ob[c].font = Font(name=FONT, bold=True, color="FFFFFF")
    ob[c].fill = HEAD
for i, e in enumerate(new, 2):
    ob[f"J{i}"] = e["id"]
    ob[f"K{i}"] = f'=COUNTIFS($A$2:$A${ON},J{i},$G$2:$G${ON},"Y")'
    ob[f"L{i}"] = f'=K{i}/COUNTIF($A$2:$A${ON},J{i})'
    ob[f"L{i}"].number_format = "0%"

wb.move_sheet("Leave Balances", offset=-(wb.sheetnames.index("Leave Balances") - 1))
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and c.font.name != FONT:
                c.font = Font(name=FONT, bold=c.font.bold, color=c.font.color, size=c.font.size)
wb.save("hr_leave_onboarding.xlsx")
print("employees", len(emps), "requests", len(reqs), "onboarding rows", ON - 1, "new starters", len(new))
