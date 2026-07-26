import openpyxl, datetime, collections
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = "/root/.claude/uploads/c0bcfcbe-ebbf-5a45-be22-042c201feaba/312ec2f4-San_Diego_Probate_Real_Estate_2024_2025.xlsx"
OUT = "/home/user/shanghai-restaurant/San_Diego_Aged_Open_Probate_Targets.xlsx"
TODAY = datetime.datetime(2026, 7, 26)
CUTOFF = TODAY - datetime.timedelta(days=730)

wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb["All Real Estate"]
hdr = [c.value for c in ws[1]]
rows = [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=2, values_only=True)]


def num(v):
    try:
        f = float(v)
        return f if f == f else None
    except (TypeError, ValueError):
        return None


def s(v):
    return "" if v is None else str(v).strip()


aged = [r for r in rows
        if isinstance(r["Date Filed"], datetime.datetime) and r["Date Filed"] <= CUTOFF]

SD_COUNTY_PREFIXES = ("919", "920", "921", "922")

def data_flag(r):
    """Source-data valuation problems: the property record matched the wrong parcel."""
    sqft = num(r["Square Feet"]) or 0
    assessed = num(r["Assessed/Estimated Value"]) or 0
    grossfm = num(r["Gross FM Real Property"]) or 0
    mortgage = num(r["Mortgage"]) or 0
    f = []
    if sqft >= 99999:
        f.append("sq ft is a placeholder value")
    elif sqft > 10000:
        f.append("sq ft too large for a home - whole building matched")
    if grossfm > 0 and assessed > 5 * grossfm:
        f.append("assessed value far exceeds court-reported value")
    if mortgage > 0 and assessed > 0 and mortgage > assessed:
        f.append("mortgage exceeds property value")
    if assessed > 5_000_000 and not (num(r["Beds"]) or 0):
        f.append("multi-million valuation with no home characteristics")
    return "; ".join(f)


scored = []
for r in aged:
    months = (TODAY - r["Date Filed"]).days / 30.44
    flag = data_flag(r)
    if flag:
        # Trust only the value reported in the petition itself; assessed data is unreliable here.
        equity = num(r["Gross FM Real Property"]) or 0
    else:
        equity = num(r["Estimated Equity"]) or num(r["Gross FM Real Property"]) or num(r["Assessed/Estimated Value"]) or 0
    absentee = "ABSENTEE" in s(r["Property Status"]).upper()
    limited = s(r["Authority"]).lower().startswith("limited")
    propper = s(r["Attorney Name"]).lower() in ("pro per", "pro se", "in pro per")

    pz, dz = s(r["Petitioner ZIP"])[:5], s(r["Decedent ZIP"])[:5]
    remote = bool(pz) and bool(dz) and pz != dz
    out_of_county = bool(pz) and not pz.startswith(SD_COUNTY_PREFIXES)

    score, why = 0, []
    # Age beyond the 2-year gate
    extra = max(0.0, months - 24)
    age_pts = min(20, round(extra * 2))
    score += age_pts
    if age_pts:
        why.append(f"{months:.0f}mo old")
    # Equity tiers
    for thresh, pts, lbl in ((1_000_000, 30, "$1M+ equity"), (600_000, 24, "$600k+ equity"),
                             (350_000, 18, "$350k+ equity"), (150_000, 10, "$150k+ equity"),
                             (1, 4, "equity present")):
        if equity >= thresh:
            score += pts
            why.append(lbl)
            break
    if absentee:
        score += 15
        why.append("vacant/absentee")
    if limited:
        score += 12
        why.append("limited authority (court confirmation)")
    if propper:
        score += 12
        why.append("self-represented")
    if out_of_county:
        score += 12
        why.append("out-of-county heir")
    elif remote:
        score += 6
        why.append("heir not at property")

    if flag:
        score -= 10
        why.append("VERIFY VALUE FIRST")

    scored.append({
        "score": score, "why": "; ".join(why), "months": round(months, 1),
        "equity": equity, "flag": flag, "r": r,
    })

scored.sort(key=lambda x: (-x["score"], -x["equity"]))

# ---------------- workbook ----------------
out = openpyxl.Workbook()
NAVY = "1F3864"
HFILL = PatternFill("solid", fgColor=NAVY)
HFONT = Font(bold=True, color="FFFFFF", size=10)
TITLE = Font(bold=True, size=14, color=NAVY)
THIN = Border(*[Side(style="thin", color="D9D9D9")] * 4)

# ---- Sheet 1: How to use
sh = out.active
sh.title = "Start Here"
sh.sheet_view.showGridLines = False
lines = [
    ("San Diego Aged Probate Targets - Real Estate", "title"),
    (f"Built {TODAY:%B %d, %Y} from your San Diego Probate Real Estate 2024-2025 workbook.", "sub"),
    ("", ""),
    ("What this is", "h"),
    (f"The {len(scored)} probate cases in your file that were filed on or before {CUTOFF:%B %d, %Y} - i.e. they are now", "b"),
    ("at least 2 years old. Every one has confirmed or probable real property attached. They are scored and", "b"),
    ("ranked so you work the best leads first.", "b"),
    ("", ""),
    ("The one thing this does NOT tell you", "h"),
    ("Whether the case is still OPEN. Your source workbook says the same thing in its own methodology note.", "b"),
    ("Filing date proves age; it does not prove the estate never closed. Column A gives you the case number to", "b"),
    ("check against the court's Register of Actions at odyroa.sdcourt.ca.gov - use the Verified Status column", "b"),
    ("to record what you find. Work top-down and you check the most valuable files first.", "b"),
    ("", ""),
    ("How the score works", "h"),
    ("Age beyond 2 years          up to 20 pts   the longer it sits, the more likely it is stuck", "b"),
    ("Estimated equity            up to 30 pts   deal size - $1M+ scores highest", "b"),
    ("Vacant / absentee              15 pts      nobody living there, carrying costs are mounting", "b"),
    ("Limited authority              12 pts      sale needs court confirmation, they need an expert", "b"),
    ("Self-represented (Pro Per)     12 pts      no attorney steering the file", "b"),
    ("Out-of-county heir             12 pts      remote petitioner, far likelier to sell", "b"),
    ("Heir not at the property        6 pts      lives locally but not in the house", "b"),
    ("", ""),
    ("Check the Data Flag column before you quote a number", "h"),
    ("About 1 in 8 records in the source data has a broken property valuation - a unit number matched the whole", "b"),
    ("apartment building, or the sq ft field holds a placeholder 99999. Those rows show a plain-English warning in", "b"),
    ("the Data Flag column, their equity falls back to the value reported in the petition itself, and they are", "b"),
    ("pushed down the ranking. Never quote equity on a flagged row without pulling the parcel yourself.", "b"),
    ("", ""),
    ("Coverage gaps you should close", "h"),
    ("Your source data is missing 8 months. See the Coverage Gaps tab for the exact list. Filling those in", "b"),
    ("would add roughly 400-500 more aged cases to this pipeline.", "b"),
]
for i, (txt, kind) in enumerate(lines, start=1):
    c = sh.cell(row=i, column=1, value=txt)
    if kind == "title":
        c.font = TITLE
    elif kind == "sub":
        c.font = Font(italic=True, size=10, color="595959")
    elif kind == "h":
        c.font = Font(bold=True, size=11, color=NAVY)
    else:
        c.font = Font(size=10)
sh.column_dimensions["A"].width = 112

# ---- Sheet 2: Priority Targets
COLS = [
    ("Case No", 14, lambda d: d["r"]["Case No"]),
    ("Score", 7, lambda d: d["score"]),
    ("Why It Scores", 46, lambda d: d["why"]),
    ("Verified Status", 15, lambda d: ""),
    ("Months Old", 10, lambda d: d["months"]),
    ("Date Filed", 11, lambda d: d["r"]["Date Filed"]),
    ("Decedent Name", 26, lambda d: d["r"]["Decedent Name"]),
    ("Property Address", 28, lambda d: d["r"]["Decedent Address"]),
    ("City", 14, lambda d: d["r"]["Decedent City"]),
    ("ZIP", 8, lambda d: d["r"]["Decedent ZIP"]),
    ("Est. Equity", 13, lambda d: d["equity"] or None),
    ("Gross FM Real Prop", 15, lambda d: num(d["r"]["Gross FM Real Property"])),
    ("Property Status", 15, lambda d: d["r"]["Property Status"]),
    ("Data Flag", 34, lambda d: d["flag"]),
    ("Beds", 6, lambda d: num(d["r"]["Beds"])),
    ("Baths", 6, lambda d: num(d["r"]["Total Baths"])),
    ("Sq Ft", 8, lambda d: num(d["r"]["Square Feet"])),
    ("Year Built", 9, lambda d: num(d["r"]["Year Built"])),
    ("Petitioner Name", 24, lambda d: d["r"]["Petitioner Name"]),
    ("Petitioner City", 15, lambda d: d["r"]["Petitioner City"]),
    ("Petitioner ST", 8, lambda d: d["r"]["Petitioner State"]),
    ("Petitioner ZIP", 11, lambda d: d["r"]["Petitioner ZIP"]),
    ("Relationship", 14, lambda d: d["r"]["Relationship"]),
    ("Authority", 10, lambda d: d["r"]["Authority"]),
    ("Attorney Name", 22, lambda d: d["r"]["Attorney Name"]),
    ("Attorney Phone", 15, lambda d: d["r"]["Attorney Phone"]),
    ("Attorney Email", 26, lambda d: d["r"]["Attorney Email"]),
]

t = out.create_sheet("Priority Targets")
t.sheet_view.showGridLines = False
for j, (name, w, _) in enumerate(COLS, start=1):
    c = t.cell(row=1, column=j, value=name)
    c.fill, c.font = HFILL, HFONT
    c.alignment = Alignment(vertical="center", wrap_text=True)
    t.column_dimensions[get_column_letter(j)].width = w
t.row_dimensions[1].height = 30

BANDS = [(47, "C6EFCE"), (34, "FFEB9C"), (-999, "FFFFFF")]
for i, d in enumerate(scored, start=2):
    band = next(f for lo, f in BANDS if d["score"] >= lo)
    for j, (_, _, fn) in enumerate(COLS, start=1):
        c = t.cell(row=i, column=j, value=fn(d))
        c.font = Font(size=9)
        c.border = THIN
        if band != "FFFFFF" and j <= 4:
            c.fill = PatternFill("solid", fgColor=band)
    t.cell(row=i, column=6).number_format = "yyyy-mm-dd"
    for col in (11, 12):
        t.cell(row=i, column=col).number_format = '"$"#,##0'
    t.cell(row=i, column=2).font = Font(size=10, bold=True)
    t.cell(row=i, column=3).alignment = Alignment(wrap_text=True, vertical="top")
t.freeze_panes = "E2"
t.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{len(scored)+1}"

# ---- Sheet 3: Coverage Gaps
g = out.create_sheet("Coverage Gaps")
g.sheet_view.showGridLines = False
present = {r["Date Filed"].strftime("%Y-%m") for r in rows if isinstance(r["Date Filed"], datetime.datetime)}
months = []
y, m = 2024, 1
while (y, m) <= (2026, 7):
    months.append(f"{y}-{m:02d}")
    m += 1
    if m == 13:
        y, m = y + 1, 1
g["A1"] = "Coverage Gaps in Your Source Data"
g["A1"].font = TITLE
g["A3"] = "Months with no cases in the source workbook. Pulling these would materially grow the pipeline."
g["A3"].font = Font(size=10, italic=True, color="595959")
for j, h in enumerate(["Month", "Cases in File", "Status", "2+ Years Old Today?"], start=1):
    c = g.cell(row=5, column=j, value=h)
    c.fill, c.font = HFILL, HFONT
cnt = collections.Counter(r["Date Filed"].strftime("%Y-%m") for r in rows
                          if isinstance(r["Date Filed"], datetime.datetime))
rr = 6
for mo in months:
    n = cnt.get(mo, 0)
    dt = datetime.datetime(int(mo[:4]), int(mo[5:]), 1)
    g.cell(row=rr, column=1, value=mo).font = Font(size=10)
    g.cell(row=rr, column=2, value=n).font = Font(size=10)
    st = g.cell(row=rr, column=3, value="MISSING" if n == 0 else "present")
    st.font = Font(size=10, bold=(n == 0), color="C00000" if n == 0 else "375623")
    g.cell(row=rr, column=4, value="Yes" if dt <= CUTOFF else "not yet").font = Font(size=10)
    rr += 1
for col, w in zip("ABCD", (12, 14, 12, 20)):
    g.column_dimensions[col].width = w

out.save(OUT)

# ---- console report
print(f"WROTE {OUT}")
print(f"aged cases: {len(scored)}")
tot = sum(d['equity'] for d in scored)
print(f"total equity: ${tot:,.0f}")
for lo, hi, lbl in ((47, 999, "A-tier (47+)"), (34, 47, "B-tier (34-46)"), (0, 34, "C-tier (<34)")):
    grp = [d for d in scored if lo <= d["score"] < hi]
    if grp:
        print(f"  {lbl}: {len(grp):4} cases  ${sum(d['equity'] for d in grp):>14,.0f}")
print(f"flagged (bad valuation): {sum(1 for d in scored if d['flag'])}")
print("\nTOP 12:")
for d in scored[:12]:
    r = d["r"]
    print(f"  {d['score']:3}  {s(r['Case No']):12} ${d['equity']:>11,.0f}  {s(r['Decedent Name'])[:26]:28} {s(r['Decedent City'])[:12]:13} {d['why'][:60]}")
print("\nMISSING MONTHS:", ", ".join(mo for mo in months if cnt.get(mo, 0) == 0))
