# -*- coding: utf-8 -*-
"""Second-pass review: fill edge-case data into a copy of the delivered workbook, recalculate, assert results."""
import datetime as dt, subprocess, json, sys
from openpyxl import load_workbook

SRC, OUT = "final.xlsx", "stress.xlsx"
wb = load_workbook(SRC)
FIRST, CW = 7, 17
col_h = lambda k, j: FIRST + CW * (k - 1) + 2 * (j - 1)
pay = lambda k: FIRST + CW * (k - 1) + 16
def sess(n): return (n - 1) // 8 + 1, (n - 1) % 8 + 1
def tick(ws, row, n, hw=True):
    k, j = sess(n); ws.cell(row=row, column=col_h(k, j), value=True)
    if hw: ws.cell(row=row, column=col_h(k, j) + 1, value=True)
today = dt.date.today()
DAYS = {7: "السبت", 1: "الأحد", 2: "الاثنين", 3: "الثلاثاء", 4: "الأربعاء", 5: "الخميس", 6: "الجمعة"}
dayname = lambda d: DAYS[(d.isoweekday() % 7) + 1]

# clear the shipped example rows
g1 = wb["مجموعة 1"]; g1["B5"] = None; g1["C5"] = None
t = wb["المواعيد"]
for c in "BCDE": t[c + "5"] = None

# S1: month boundary, last row, renamed month label
g1["A1"] = "علوم 4"
g1["B5"], g1["B19"] = "طالب أ", "طالب ب"; g1["C19"] = "01055555555"
for n in range(1, 9):
    tick(g1, 5, n)
    if n != 3: tick(g1, 19, n)
g1.cell(row=1, column=FIRST, value="أكتوبر")
g1.cell(row=5, column=pay(1), value=True)
g1.cell(row=28, column=col_h(1, 1), value=10); g1.cell(row=44, column=col_h(1, 1), value=7)
g1.cell(row=28, column=col_h(1, 1) + 1, value=40); g1.cell(row=44, column=col_h(1, 1) + 1, value=30)
g1.cell(row=44, column=pay(1), value=5)            # a NUMBER typed in the notes column must not count as a score
g1.cell(row=44, column=col_h(1, 2), value=9)      # a score in a column with NO final mark must be ignored
g1.cell(row=30, column=pay(1), value="غاب بعذر - مريض")

# S2: 9 sessions -> month 2 session 1, blank month-2 label, two unpaid months
g2 = wb["مجموعة 2"]; g2["A1"] = "Math 5"
g2["B5"] = "طالب ج"
for n in range(1, 10): tick(g2, 5, n)
g2.cell(row=1, column=FIRST + CW).value = None

# S3: 17 sessions; 3 unpaid months; paid 1 & 3 not 2; stopped student; homework-only tick
g3 = wb["مجموعة 3"]; g3["A1"] = "علوم 6"
for i, nm in enumerate(["س", "ص", "ع", "ف"]): g3.cell(row=5 + i, column=2, value=nm)
for n in range(1, 18):
    tick(g3, 5, n); tick(g3, 6, n)
    if n <= 9: tick(g3, 7, n)
g3.cell(row=6, column=pay(1), value=True); g3.cell(row=6, column=pay(3), value=True)
k, j = sess(17); g3.cell(row=8, column=col_h(k, j) + 1, value=True)   # row 8: homework tick only, never "present"

# S4: 15 students, everyone absent twice and unpaid -> overflow lines on the dashboard
g4 = wb["مجموعة 4"]; g4["A1"] = "رياضة 6"
for i in range(15):
    g4.cell(row=5 + i, column=2, value="ط%02d" % (i + 1)); g4.cell(row=5 + i, column=3, value="0109%07d" % i)
    for n in (1, 2): tick(g4, 5 + i, n)
tick(g4, 5, 3); tick(g4, 5, 4)                                         # only the first student attends 3 and 4

# S5: timetable - ties, a time not in the list, a non-group name, two weekly slots
rows = [(dayname(today), "6:00 م", "علوم 6", "ب"), (dayname(today), "6:00 م", "Math 5", "أ"),
        (dayname(today), "4:15 م", "كريم برايفت", ""), (dayname(today), "9:00 ص", "علوم 4", ""),
        (dayname(today + dt.timedelta(days=2)), "5:00 م", "علوم 4", ""),
        (dayname(today - dt.timedelta(days=1)), "3:00 م", "رياضة 6", "")]
for i, r in enumerate(rows):
    for c, v in zip("BCDE", r): t["%s%d" % (c, 5 + i)] = v
g3.cell(row=3, column=col_h(*sess(17)), value=today)                    # group 3 already recorded today's session
g3["A23"] = "حل تمارين"

# S7: individuals - last row full year, and a mid-pack student; grades need both score and "out of"
on = wb["أونلاين"]
on["B33"] = "آخر صف"
for n in range(1, 97): tick(on, 33, n)
for k in range(1, 13): on.cell(row=33, column=pay(k), value=True)
on["B4"] = "نص باقة"
for n in range(1, 11): tick(on, 4, n)
on.cell(row=4, column=pay(1), value=True); on.cell(row=4, column=pay(2), value=True)
on.cell(row=37, column=col_h(1, 1), value=9)                                   # score without "out of" -> ignored
on.cell(row=37, column=col_h(1, 2), value=6); on.cell(row=37, column=col_h(1, 2) + 1, value=10)
on.cell(row=37, column=pay(1), value="ملاحظة نصية"); on.cell(row=37, column=col_h(1, 3), value="غ")

wb["الرئيسية"]["H27"] = 2
wb.save(OUT)
r = subprocess.run([sys.executable, "/mnt/skills/public/xlsx/scripts/recalc.py", OUT, "240"], capture_output=True, text=True)
res = json.loads(r.stdout); print("recalc:", res["status"], res["total_errors"], res.get("error_summary"))

v = load_workbook(OUT, data_only=True)
ok = True
def check(label, got, exp):
    global ok
    good = got == exp
    ok &= good
    print(("PASS " if good else "FAIL ") + label + ": " + repr(got) + ("" if good else "   expected " + repr(exp)))

a = v["مجموعة 1"]
check("S1 status at 8/8 with renamed month", a["A2"].value, "أكتوبر · حصة 8 من 8")
check("S1 last row absences", a["D19"].value, 1)
check("S1 last row fees uses renamed label", a["F19"].value, "أكتوبر")
check("S1 paid student", a["F5"].value, "تمام")
check("S1 grades last row (number in notes ignored)", a["C44"].value, "37 / 50")
check("S1 rank last row", a["E44"].value, 1)
b = v["مجموعة 2"]
check("S2 status month 2 with blank label", b["A2"].value, "شهر 2 · حصة 1 من 8")
check("S2 two unpaid months", b["F5"].value, "شهرين")
c = v["مجموعة 3"]
check("S3 three unpaid months", c["F5"].value, "3 شهور")
check("S3 paid 1 and 3, not 2", c["F6"].value, "الشهر 2")
check("S3 stopped student", c["D7"].value, "متوقف")
check("S3 homework-only tick counts as a fee month, not attendance", (c["D8"].value, c["F8"].value), ("-", "الشهر 3"))
check("S3 next date skips today once today's session is recorded", c["C23"].value.startswith("النهارده"), False)
d = v["مجموعة 4"]
check("S4 absences rolling (4 held, attended 2)", d["D6"].value, 2)
check("S4 first student no absences", d["D5"].value, 0)
h = v["الرئيسية"]
names_abs = [h["B%d" % r].value for r in range(29, 39) if h["B%d" % r].value]
check("S4 absence table full (10 rows)", len(names_abs), 10)
check("S4 absence overflow line", h["B39"].value, "+ 4 كمان · عمود «غياب» في صفحات المجموعات")
check("KPI absence alerts (14 students with 2 absences)", h["F4"].value, 14)
names_pay = [h["B%d" % r].value for r in range(43, 55) if h["B%d" % r].value]
check("unpaid table full (12 rows)", len(names_pay), 12)
check("KPI unpaid (1 + 1 + 4 + 15 + 1 online renewal)", h["H4"].value, 22)
check("unpaid overflow line", h["B55"].value, "+ 10 كمان · عمود «المصاريف» في صفحات المجموعات")
today_rows = [(h["B%d" % r].value, h["C%d" % r].value, h["G%d" % r].value) for r in range(9, 15) if h["C%d" % r].value]
print("     today's sessions:", today_rows)
check("S5 order: by time, ties keep sheet order, unknown time last",
      [x[1] for x in today_rows], ["علوم 4", "علوم 6", "Math 5", "كريم برايفت"])
check("S5 topic shown for matching group only", [x[2] for x in today_rows], [None, "حل تمارين", None, None])
check("KPI sessions today", h["D4"].value, 4)
check("S5 group with a slot today -> 'النهارده'", a["C23"].value.startswith("النهارده"), True)
check("S5 slot only yesterday -> 6 days ahead", d["C23"].value.startswith(dayname(today - dt.timedelta(days=1))), True)
o = v["أونلاين"]
check("S7 last row full year", (o["D33"].value, o["E33"].value, o["F33"].value), ("8/8", 0, "تجديد"))
check("S7 mid pack", (o["D4"].value, o["E4"].value, o["F4"].value), ("2/8", 6, "تمام"))
check("S7 grades need both score and out-of", o["C37"].value, "6 / 10")
check("students KPI", h["B4"].value, 2 + 1 + 4 + 15 + 2)
print("ALL PASS" if ok else "SOME FAILED")
