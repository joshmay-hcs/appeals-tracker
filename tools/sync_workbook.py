# Sync builder: rebuilds Appeals_Tracker.xlsx from a JSON export of the tracker's shared database.
# Usage: python build_xlsx2.py <export_dir> <out.xlsx>
import json, os, sys, glob, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment

EXP, OUT = sys.argv[1], sys.argv[2]
def coll(name):
  return [json.load(open(f, encoding="utf-8")) for f in glob.glob(os.path.join(EXP, name, "*.json"))]
CARRIERS = sorted(coll("carriers"), key=lambda c: c["name"])
RULES = coll("rules"); REPORT_RULES = sorted(coll("reportRules"), key=lambda r: (r["carrier"], r["id"]))
REASONS = sorted(coll("reasons"), key=lambda r: r["id"]); TKS = sorted(coll("timekeepers"), key=lambda t: t["name"])
MAPPING = sorted(coll("mapping"), key=lambda m: int(m["id"][1:]) if m["id"][1:].isdigit() else 999)
MATTERS = sorted(coll("matters"), key=lambda m: m["id"]); CUTS = sorted(coll("cuts"), key=lambda c: (c.get("n") or 0, c["id"]))
APPROVALS = sorted(coll("approvals"), key=lambda a: a.get("date", "")); LIBRARY = sorted(coll("library"), key=lambda l: l.get("date", ""), reverse=True)
sp = os.path.join(EXP, "meta", "settings.json"); SETTINGS = json.load(open(sp, encoding="utf-8")) if os.path.exists(sp) else {}

TODAY = dt.date.today()
SYNCED = dt.datetime.now().strftime("%m/%d/%Y %I:%M %p")
N = max(500, len(CUTS) + 200); LAST = N + 1
STATUSES = ["Needs Review", "Appeal Drafted", "Submitted", "Partially Restored", "Fully Restored", "Denied", "Accepted - No Appeal"]
PREVS = ["Preventable", "Potentially Preventable", "Not Preventable"]

NAVY = "1A497C"; WARM = "F1F0EF"
F = "Arial"
def font(**k): return Font(name=F, **k)
HDR = font(bold=True, color="FFFFFF"); HFILL = PatternFill("solid", fgColor=NAVY)
INP = font(color="0000FF"); NORM = font(); BOLD = font(bold=True); GRN = font(color="008000")
TITLE = font(bold=True, size=16, color=NAVY); SUB = font(italic=True, color="5C6068"); H2 = font(bold=True, size=12, color=NAVY)
WRAP = Alignment(wrap_text=True, vertical="top")
thin = Side(style="thin", color="D9D6D0"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
YFILL = PatternFill("solid", fgColor="FFF4D6"); GFILL = PatternFill("solid", fgColor=WARM)
MONEY = '$#,##0.00'; MONEY0 = '$#,##0'; DATE = "mm/dd/yyyy"

CAR = {c["id"]: c for c in CARRIERS}; TK = {t["id"]: t for t in TKS}; MAT = {m["id"]: m for m in MATTERS}
REASON = {r["name"]: r for r in REASONS}
TOPIC_LABEL = {"block":"Block billing","vague":"Vague narratives","clerical":"Clerical / overhead","paralegal":"Paralegal-level work","duplicate":"Duplicates","research":"Research","travel":"Travel","multi":"Multiple attorneys / timekeepers","intraoffice":"Intra-office conferences","review":"File & document review","receipt":"Receipts","approval":"Prior approval","rate":"Rates","budget":"Budget","hours":"Daily hours","nominal":"Nominal tasks","coding":"UTBMS coding","correspondence":"Correspondence","staffing":"Staffing","phase":"Scope / phase","general":"General"}
def d(s):
  try: return dt.date.fromisoformat(s) if s else None
  except ValueError: return None
def cname(cid): return CAR.get(cid, {}).get("name", cid or "")
def tkname(tid): return TK.get(tid, {}).get("name", tid or "")

wb = Workbook()
def header(ws, row, cols, widths=None, col0=1):
  for i, h in enumerate(cols, col0):
    c = ws.cell(row=row, column=i, value=h); c.font = HDR; c.fill = HFILL; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BOX
  if widths:
    for i, w in enumerate(widths, col0): ws.column_dimensions[ws.cell(row=row, column=i).column_letter].width = w
def titled(ws, title, sub):
  ws["A1"] = title; ws["A1"].font = TITLE; ws["A2"] = sub; ws["A2"].font = SUB
def put(ws, r, c, v, f=NORM, fmt=None, wrap=False):
  cell = ws.cell(row=r, column=c, value=v); cell.font = f; cell.border = BOX
  if fmt: cell.number_format = fmt
  if wrap: cell.alignment = WRAP
  return cell

# ------------------------------------------------------------------ INSTRUCTIONS
ws = wb.active; ws.title = "INSTRUCTIONS"
ws["A1"] = "APPEALS TRACKER WORKBOOK"; ws["A1"].font = TITLE
ws["A2"] = f"Synced from the Appeals Tracker on {SYNCED}. {len(CUTS)} cuts, {len(MATTERS)} matters, {len(APPROVALS)} approvals, {len(LIBRARY)} library appeals."; ws["A2"].font = font(bold=True, color="7A4A00")
ws["A3"] = "This workbook is a copy of the tracker for reporting, finance and offline use. Make changes in the tracker, then ask Claude to re-sync. Edits typed here are replaced at the next sync."; ws["A3"].font = SUB
steps = [
 ("Where things live", "The Appeals Tracker web app is the source of truth. Elizabeth manages carriers, guideline rules, reporting rules, cut reasons, timekeepers and intake mapping on its Guidelines & Admin tab. This workbook mirrors all of it."),
 ("How to re-sync", "In Claude, say: \"Sync the Excel workbook from the appeals tracker.\" Claude exports every table from the tracker's database and rebuilds this file with fresh formulas. Takes about a minute."),
 ("Color legend", "Blue text = values copied from the tracker (you may type in blank rows for what-if work). Black text = formulas. Green text = pulled from another tab. Yellow fill = settings or fact cells."),
 ("BILLING CUTS", "One row per reduced line. Platform, role, cut, preventability, appeal deadline, days left, written off and pending fill themselves in. 'Tracker ID' links a row back to the app."),
 ("DASHBOARD / MONTHLY SUMMARY", "Totals, preventable share, recovery rate, and month-by-month trend. Change the year on MONTHLY SUMMARY B2."),
 ("WRITE-OFFS", "For finance: what was lost after appeals. Written off = denied or accepted cuts plus the unrecovered part of partial wins. Pending = cuts still open or awaiting a carrier decision."),
 ("REPORT DEADLINES", "Carrier reporting deadlines per matter, calculated at sync from each matter's assignment, mediation and trial dates. 'Days Left' counts down live."),
 ("APPROVALS / APPEAL LIBRARY", "Prior approvals on file (with expiry status) and appeals that won money back, to reuse their wording."),
 ("APPEAL DRAFT", "Type an Entry ID, fill the yellow fact cells, copy the draft. A ready-made Claude prompt is below it. (The tracker's Appeal Builder does the same with more help.)"),
 ("Reference tabs", "CARRIER GUIDELINES (caps + rules with effective dates; the current version of each rule is marked), REPORTING RULES, REASON MATRIX, INTAKE MAPPING, LOOKUPS."),
 ("Demo data", "All carriers, matters, people and amounts are fictional. Clear them in the tracker (Guidelines & Admin > Data & Excel) and re-sync."),
]
for i, (a, b) in enumerate(steps, 5):
  ws.cell(row=i, column=1, value=a).font = BOLD; ws.cell(row=i, column=2, value=b).alignment = WRAP
  ws.cell(row=i, column=1).alignment = WRAP
ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 110

# ------------------------------------------------------------------ LOOKUPS
lk = wb.create_sheet("LOOKUPS")
header(lk, 1, ["Carriers", "Platforms", "Timekeeper", "Role", "Rate ($/hr)", "Email", "First-year?", "Active?", "Appeal Status", "Preventability", "Appeal Round", "Matters"], [26, 22, 18, 12, 11, 30, 11, 9, 22, 24, 14, 44])
for i, c in enumerate(CARRIERS, 2): lk.cell(row=i, column=1, value=c["name"]).font = INP
for i, p in enumerate(["Legal Solutions Suite", "Legal-X", "Other"], 2): lk.cell(row=i, column=2, value=p).font = INP
for i, t in enumerate(TKS, 2):
  for j, v in enumerate([t["name"], t.get("role", ""), t.get("rate"), t.get("email", ""), "Yes" if t.get("firstYear") else "No", "Yes" if t.get("active", True) else "No"], 3):
    c = lk.cell(row=i, column=j, value=v); c.font = INP
    if j == 5: c.number_format = MONEY
for i, s in enumerate(STATUSES, 2): lk.cell(row=i, column=9, value=s).font = INP
for i, s in enumerate(PREVS, 2): lk.cell(row=i, column=10, value=s).font = INP
for i, s in enumerate(["0 - Not appealed", "1st", "2nd"], 2): lk.cell(row=i, column=11, value=s).font = INP
for i, m in enumerate(MATTERS, 2): lk.cell(row=i, column=12, value=m["name"]).font = INP

# ------------------------------------------------------------------ CARRIER GUIDELINES
cg = wb.create_sheet("CARRIER GUIDELINES")
titled(cg, "CARRIER GUIDELINES", "Fictional carriers. Caps drive the tracker's pre-bill scrub. Edit in the tracker (Guidelines & Admin > Carriers / Guideline rules), then re-sync.")
cgh = ["Carrier", "Platform", "Appeal Window (days)", "Appeal Window Note", "Appeal Rounds", "Billing Cadence", "Stands in for", "Daily Hours Cap", "Research Hours Cap", "Block Billing OK Under (hrs)", "Travel Approval Over (hrs)", "Receipt Over ($)", "Partner Research Needs Approval", "First-Year Associates Paid", "Embedded Time OK", "Carrier ID"]
header(cg, 4, cgh, [24, 20, 11, 44, 36, 36, 36, 9, 9, 10, 10, 9, 11, 10, 9, 8])
cg.row_dimensions[4].height = 45
CG_FIRST = 5
yn = lambda v: "" if v is None else ("Yes" if v else "No")
for i, c in enumerate(CARRIERS, CG_FIRST):
  vals = [c["name"], c.get("platform"), c.get("appealDays"), c.get("appealDaysNote", ""), c.get("rounds", ""), c.get("billing", ""), c.get("source", ""),
          c.get("dailyHours"), c.get("researchHours"), c.get("blockMin"), c.get("travelApprovalHours"), c.get("receiptOver"),
          yn(c.get("partnerResearchNeedsApproval")), yn(c.get("firstYearPaid")), yn(c.get("embeddedTimeOk")), c["id"]]
  for j, v in enumerate(vals, 1):
    cell = put(cg, i, j, v, INP, wrap=True)
    if j == 3: cell.fill = YFILL
    if j == 12: cell.number_format = MONEY0
  note = (c.get("appealDaysNote") or "").lower()
  if "assum" in note or "confirm" in note:
    cg.cell(row=i, column=3).comment = Comment("Assumed for the demo; the guideline does not state a window. Confirm with the carrier or billing vendor.", "Tracker")
CG_LAST = CG_FIRST + 40
RULE_HDR = CG_LAST + 4
cg.cell(row=RULE_HDR - 1, column=1, value="Rules by topic. 'Current' = latest version in effect today; APPEAL DRAFT uses the Current row. Older versions stay for appeals on earlier invoices.").font = H2
header(cg, RULE_HDR, ["Lookup Key", "Carrier", "Topic", "Topic Label", "Section Cite", "Rule", "Effective", "Version"])
by_key = {}
for r in RULES: by_key.setdefault((cname(r["carrier"]), r["topic"]), []).append(r)
row = RULE_HDR + 1
for (cn, topic) in sorted(by_key):
  vers = sorted(by_key[(cn, topic)], key=lambda r: r.get("effective", ""), reverse=True)
  cur = next((v for v in vers if (d(v.get("effective")) or dt.date.min) <= TODAY), None)
  ordered = ([cur] if cur else []) + [v for v in vers if v is not cur]
  for v in ordered:
    eff = d(v.get("effective"))
    status = "Current" if v is cur else ("Future" if eff and eff > TODAY else "Superseded")
    put(cg, row, 1, f'=B{row}&"|"&C{row}' if status == "Current" else "(not used)")
    put(cg, row, 2, cn, INP); put(cg, row, 3, topic, INP); put(cg, row, 4, TOPIC_LABEL.get(topic, topic))
    put(cg, row, 5, v.get("cite", ""), INP, wrap=True); put(cg, row, 6, v.get("text", ""), INP, wrap=True)
    put(cg, row, 7, eff, INP, DATE); put(cg, row, 8, status, BOLD if status == "Current" else SUB)
    row += 1
RULE_FIRST, RULE_LAST = RULE_HDR + 1, row + 40
cg.column_dimensions["H"].width = 11
cg.freeze_panes = "B5"

# ------------------------------------------------------------------ REPORTING RULES
rr_ = wb.create_sheet("REPORTING RULES")
titled(rr_, "CARRIER REPORTING RULES", "When each carrier expects reports. Offset = days after the trigger date (negative = before). 'recurring' repeats from the last completed report.")
header(rr_, 4, ["Carrier", "Report", "Trigger", "Offset (days)", "Rule ID"], [24, 48, 14, 12, 10])
for i, r in enumerate(REPORT_RULES, 5):
  for j, v in enumerate([cname(r["carrier"]), r["name"], r["trigger"], r.get("offset"), r["id"]], 1): put(rr_, i, j, v, INP)
rr_.freeze_panes = "A5"

# ------------------------------------------------------------------ REASON MATRIX
rm = wb.create_sheet("REASON MATRIX")
titled(rm, "CUT REASON MATRIX", "Default Preventability (yellow) drives BILLING CUTS. Appeal Argument placeholders: {rewrite} {scope} {approver} {approvalDate} {invoice} {section}.")
header(rm, 4, ["Cut Reason", "Default Preventability", "Guideline Topic", "Appeal Odds", "Usual Cause", "Prevention Tip", "Appeal Argument Template"], [28, 22, 14, 11, 40, 48, 70])
RM_FIRST = 5
for i, x in enumerate(REASONS, RM_FIRST):
  for j, v in enumerate([x["name"], x.get("prev"), x.get("topic"), x.get("strength"), x.get("cause"), x.get("tip"), x.get("argument")], 1):
    c = put(rm, i, j, v, INP, wrap=True)
    if j == 2: c.fill = YFILL
RM_LAST = RM_FIRST + 60
dvp = DataValidation(type="list", formula1="=LOOKUPS!$J$2:$J$4", allow_blank=True); rm.add_data_validation(dvp); dvp.add(f"B{RM_FIRST}:B{RM_LAST}")
rm.freeze_panes = "A5"

# ------------------------------------------------------------------ INTAKE MAPPING
im = wb.create_sheet("INTAKE MAPPING")
titled(im, "INTAKE MAPPING", "When a reduction report is read, a reviewer category or note containing the pattern is logged under this cut reason. First match wins.")
header(im, 4, ["Pattern (contains)", "Cut Reason"], [40, 30])
for i, m in enumerate(MAPPING, 5):
  put(im, i, 1, m["pattern"], INP); put(im, i, 2, m["reason"], INP)

# ------------------------------------------------------------------ BILLING CUTS
bc = wb.create_sheet("BILLING CUTS", 1)
cols = ["Entry ID","Carrier","Platform","Invoice #","Claim/File #","Matter","Invoice Date","Reduction Date","Timekeeper","TK Role",
        "UTBMS Task","UTBMS Activity","Narrative as Billed","Corrected Narrative","Hours","Original Bill ($)","Allowed ($)","Cut ($)","Cut %","Cut Reason",
        "Default Preventability","Preventability Override","Preventability","Appeal Status","Appeal Round","Appeal Deadline","Days Left",
        "Appeal Amount ($)","Recovered ($)","Net Unrecovered ($)","Written Off ($)","Pending ($)","Notes","Source","Month","Year","Tracker ID"]
widths = [8,22,20,10,13,34,12,12,16,11,9,9,46,46,7,12,11,11,8,24,20,20,20,18,13,13,9,12,12,13,12,12,36,9,7,7,9]
header(bc, 1, cols, widths); bc.row_dimensions[1].height = 32
COL = {h: i + 1 for i, h in enumerate(cols)}
def L(h): return bc.cell(row=1, column=COL[h]).column_letter
ROUND = {0: "0 - Not appealed", 1: "1st", 2: "2nd"}
cg_rng = f"'CARRIER GUIDELINES'!$A${CG_FIRST}:$A${CG_LAST}"
for r in range(2, LAST + 1):
  c = CUTS[r - 2] if r - 2 < len(CUTS) else None
  def p(h, v, f=INP, fmt=None):
    cell = bc.cell(row=r, column=COL[h], value=v); cell.font = f
    if fmt: cell.number_format = fmt
  p("Entry ID", r - 1, NORM)
  for h, fmt in [("Invoice Date", DATE), ("Reduction Date", DATE), ("Hours", "0.0"), ("Original Bill ($)", MONEY), ("Allowed ($)", MONEY), ("Appeal Amount ($)", MONEY), ("Recovered ($)", MONEY)]:
    bc.cell(row=r, column=COL[h]).number_format = fmt
  if c:
    m = MAT.get(c.get("matterId"))
    p("Carrier", cname(c.get("carrier"))); p("Invoice #", c.get("invoice")); p("Claim/File #", c.get("claim")); p("Matter", m["name"] if m else c.get("matter"))
    p("Invoice Date", d(c.get("invDate")), fmt=DATE); p("Reduction Date", d(c.get("redDate")), fmt=DATE)
    p("Timekeeper", tkname(c.get("tk"))); p("UTBMS Task", c.get("task") or None); p("UTBMS Activity", c.get("activity") or None)
    p("Narrative as Billed", c.get("narrative")); p("Corrected Narrative", c.get("rewrite") or None); p("Hours", c.get("hours") or None, fmt="0.0")
    p("Original Bill ($)", c.get("billed"), fmt=MONEY); p("Allowed ($)", c.get("allowed"), fmt=MONEY)
    p("Cut Reason", c.get("reason"))
    rd = REASON.get(c.get("reason"))
    if rd and c.get("prev") and c["prev"] != rd.get("prev"): p("Preventability Override", c["prev"])
    p("Appeal Status", c.get("status")); p("Appeal Round", ROUND.get(int(c.get("round") or 0), "0 - Not appealed"))
    p("Appeal Amount ($)", c.get("appealed") or 0, fmt=MONEY); p("Recovered ($)", c.get("recovered") or 0, fmt=MONEY)
    p("Notes", c.get("notes") or None); p("Source", c.get("source") or None); p("Tracker ID", str(c["id"]))
  B, H, I, P_, Q, R_, T, U, V, X, Z, AC = (L(x) + str(r) for x in ["Carrier","Reduction Date","Timekeeper","Original Bill ($)","Allowed ($)","Cut ($)","Cut Reason","Default Preventability","Preventability Override","Appeal Status","Appeal Deadline","Recovered ($)"])
  p("Platform", f'=IF({B}="","",IFERROR(INDEX(\'CARRIER GUIDELINES\'!$B${CG_FIRST}:$B${CG_LAST},MATCH({B},{cg_rng},0)),""))', GRN)
  p("TK Role", f'=IF({I}="","",IFERROR(INDEX(LOOKUPS!$D$2:$D$60,MATCH({I},LOOKUPS!$C$2:$C$60,0)),""))', GRN)
  p("Cut ($)", f'=IF(OR({P_}="",{Q}=""),"",ROUND({P_}-{Q},2))', NORM, MONEY)
  p("Cut %", f'=IFERROR({R_}/{P_},"")', NORM, '0%')
  p("Default Preventability", f'=IF({T}="","",IFERROR(INDEX(\'REASON MATRIX\'!$B${RM_FIRST}:$B${RM_LAST},MATCH({T},\'REASON MATRIX\'!$A${RM_FIRST}:$A${RM_LAST},0)),""))', GRN)
  p("Preventability", f'=IF({V}<>"",{V},{U})', NORM)
  p("Appeal Deadline", f'=IF(OR({H}="",{B}=""),"",IFERROR({H}+INDEX(\'CARRIER GUIDELINES\'!$C${CG_FIRST}:$C${CG_LAST},MATCH({B},{cg_rng},0)),""))', NORM, DATE)
  p("Days Left", f'=IF({Z}="","",IF(OR({X}="Needs Review",{X}="Appeal Drafted"),{Z}-TODAY(),""))', NORM, "0")
  p("Net Unrecovered ($)", f'=IF({R_}="","",MAX(0,{R_}-N({AC})))', NORM, MONEY)
  p("Written Off ($)", f'=IF({R_}="","",IF(OR({X}="Partially Restored",{X}="Denied",{X}="Accepted - No Appeal"),MAX(0,{R_}-N({AC})),0))', NORM, MONEY)
  p("Pending ($)", f'=IF({R_}="","",IF(OR({X}="Fully Restored",{X}="Partially Restored",{X}="Denied",{X}="Accepted - No Appeal"),0,MAX(0,{R_}-N({AC}))))', NORM, MONEY)
  p("Month", f'=IF({H}="","",TEXT({H},"mmm"))', NORM)
  p("Year", f'=IF({H}="","",YEAR({H}))', NORM)
  for h in ["Narrative as Billed", "Corrected Narrative", "Notes", "Matter"]:
    bc.cell(row=r, column=COL[h]).alignment = WRAP

def rng(h): return f"{L(h)}2:{L(h)}{LAST}"
for h, f1 in [("Carrier", "=LOOKUPS!$A$2:$A$50"), ("Timekeeper", "=LOOKUPS!$C$2:$C$60"), ("Cut Reason", f"='REASON MATRIX'!$A${RM_FIRST}:$A${RM_LAST}"),
              ("Preventability Override", "=LOOKUPS!$J$2:$J$4"), ("Appeal Status", "=LOOKUPS!$I$2:$I$20"), ("Appeal Round", "=LOOKUPS!$K$2:$K$4"), ("Matter", "=LOOKUPS!$L$2:$L$200")]:
  dv = DataValidation(type="list", formula1=f1, allow_blank=True); bc.add_data_validation(dv); dv.add(rng(h))
dl = L("Days Left")
bc.conditional_formatting.add(rng("Days Left"), FormulaRule(formula=[f'AND(ISNUMBER({dl}2),{dl}2<0)'], fill=PatternFill("solid", fgColor="F4C7C3"), font=Font(name=F, color="9C1C14", bold=True)))
bc.conditional_formatting.add(rng("Days Left"), FormulaRule(formula=[f'AND(ISNUMBER({dl}2),{dl}2>=0,{dl}2<=7)'], fill=PatternFill("solid", fgColor="FFE3A3"), font=Font(name=F, color="7A4A00", bold=True)))
pv = L("Preventability")
for val, color in [("Preventable", "F6D3C6"), ("Potentially Preventable", "FBE6B5"), ("Not Preventable", "D6E0EC")]:
  bc.conditional_formatting.add(rng("Preventability"), FormulaRule(formula=[f'{pv}2="{val}"'], fill=PatternFill("solid", fgColor=color)))
bc.freeze_panes = "C2"
bc.auto_filter.ref = f"A1:{L('Tracker ID')}{LAST}"
def BCR(h): return f"'BILLING CUTS'!${L(h)}$2:${L(h)}${LAST}"

# ------------------------------------------------------------------ MATTERS
mt = wb.create_sheet("MATTERS")
titled(mt, "MATTERS", "Matters in the tracker. Dates drive REPORT DEADLINES. Cut and write-off totals are live formulas from BILLING CUTS.")
header(mt, 4, ["Matter", "Type", "Carrier", "Claim #", "Assigned", "Mediation", "Trial / Hearing", "Lead", "Status", "Cuts ($)", "Recovered ($)", "Written Off ($)", "Pending ($)", "Matter ID"], [40, 10, 22, 13, 12, 12, 13, 16, 9, 12, 12, 13, 12, 9])
for i, m in enumerate(MATTERS, 5):
  for j, (v, fmt) in enumerate([(m["name"], None), (m.get("type"), None), (cname(m.get("carrier")), None), (m.get("claim"), None), (d(m.get("assigned")), DATE), (d(m.get("mediation")), DATE),
                                (d(m.get("trial")), DATE), (tkname(m.get("lead")), None), (m.get("status"), None)], 1):
    put(mt, i, j, v, INP, fmt)
  for j, h in enumerate(["Cut ($)", "Recovered ($)", "Written Off ($)", "Pending ($)"], 10):
    put(mt, i, j, f'=SUMIFS({BCR(h)},{BCR("Matter")},$A{i})', NORM, MONEY0)
  put(mt, i, 14, m["id"])
mt.freeze_panes = "B5"

# ------------------------------------------------------------------ REPORT DEADLINES (mirrors reportItems() in the tracker)
items = []
for m in MATTERS:
  if m.get("status") == "Closed": continue
  done = m.get("reportsDone") or {}
  for r in REPORT_RULES:
    if r["carrier"] != m.get("carrier"): continue
    off = int(r.get("offset") or 0)
    if r["trigger"] == "recurring":
      base = d(done.get(r["id"])) or d(m.get("assigned"))
      if not base: continue
      items.append((base + dt.timedelta(days=off), m, r, None, f"Every {off} days (last done {done.get(r['id']) or 'never'})"))
    else:
      anchor = d(m.get("assigned") if r["trigger"] == "assigned" else m.get(r["trigger"]))
      if not anchor: continue
      rule_txt = f"{abs(off)} days {'after' if off >= 0 else 'before'} {r['trigger']}"
      items.append((anchor + dt.timedelta(days=off), m, r, d(done.get(r["id"])), rule_txt))
items.sort(key=lambda x: (x[3] is not None, x[0]))
rp = wb.create_sheet("REPORT DEADLINES")
titled(rp, "REPORT DEADLINES", f"Calculated at sync ({SYNCED}) from matter dates and carrier reporting rules. Mark reports done in the tracker; Days Left updates live.")
header(rp, 4, ["Due", "Days Left", "Report", "Matter", "Carrier", "Lead", "Rule", "Completed"], [12, 10, 40, 40, 22, 16, 34, 12])
for i, (due, m, r, doneD, rule_txt) in enumerate(items, 5):
  put(rp, i, 1, due, INP, DATE)
  put(rp, i, 2, f'=IF(H{i}<>"","Done",A{i}-TODAY())', NORM, "0")
  for j, v in enumerate([r["name"], m["name"], cname(m.get("carrier")), tkname(m.get("lead")), rule_txt], 3): put(rp, i, j, v, INP)
  put(rp, i, 8, doneD, INP, DATE)
RP_LAST = 5 + len(items) + 50
rp.conditional_formatting.add(f"B5:B{RP_LAST}", FormulaRule(formula=['AND(ISNUMBER(B5),B5<0)'], fill=PatternFill("solid", fgColor="F4C7C3"), font=Font(name=F, color="9C1C14", bold=True)))
rp.conditional_formatting.add(f"B5:B{RP_LAST}", FormulaRule(formula=['AND(ISNUMBER(B5),B5>=0,B5<=14)'], fill=PatternFill("solid", fgColor="FFE3A3"), font=Font(name=F, color="7A4A00", bold=True)))
rp.freeze_panes = "A5"; rp.auto_filter.ref = f"A4:H{4 + max(1, len(items))}"

# ------------------------------------------------------------------ APPROVALS
ap = wb.create_sheet("APPROVALS")
titled(ap, "PRIOR APPROVALS ON FILE", "Carrier approvals for research, experts, travel, extra staffing and other pre-approval items. Attach the approval to any appeal that relies on it.")
header(ap, 4, ["Approved On", "Matter", "Carrier", "Type", "Scope", "Approver", "Method", "Hours Cap", "Expires", "Status", "Notes"], [12, 36, 22, 22, 44, 28, 10, 9, 12, 10, 30])
for i, a in enumerate(APPROVALS, 5):
  m = MAT.get(a.get("matterId"))
  for j, (v, fmt) in enumerate([(d(a.get("date")), DATE), (m["name"] if m else a.get("matterId"), None), (cname(a.get("carrier")), None), (a.get("type"), None), (a.get("scope"), None),
                                (a.get("approver"), None), (a.get("method"), None), (a.get("hoursCap") or None, "0.0"), (d(a.get("expires")), DATE)], 1):
    put(ap, i, j, v, INP, fmt, wrap=j in (2, 5))
  put(ap, i, 10, f'=IF(I{i}="","Active",IF(I{i}<TODAY(),"Expired","Active"))')
  put(ap, i, 11, a.get("notes") or None, INP, wrap=True)
ap.conditional_formatting.add(f"J5:J{5 + len(APPROVALS) + 50}", FormulaRule(formula=['J5="Expired"'], fill=PatternFill("solid", fgColor="F4C7C3")))
ap.freeze_panes = "A5"

# ------------------------------------------------------------------ APPEAL LIBRARY
al = wb.create_sheet("APPEAL LIBRARY")
titled(al, "WINNING APPEALS LIBRARY", "Appeals that recovered money. Reuse the wording that worked for the same carrier and reason. New wins are added automatically in the tracker.")
header(al, 4, ["Date", "Carrier", "Cut Reason", "Outcome", "Cut ($)", "Recovered ($)", "Recovery %", "Tags", "Appeal Text", "Tracker Cut ID"], [12, 22, 24, 18, 10, 12, 10, 16, 90, 9])
for i, l in enumerate(LIBRARY, 5):
  for j, (v, fmt) in enumerate([(d(l.get("date")), DATE), (cname(l.get("carrier")), None), (l.get("reason"), None), (l.get("outcome"), None), (l.get("cut"), MONEY), (l.get("recovered"), MONEY)], 1):
    put(al, i, j, v, INP, fmt)
  put(al, i, 7, f'=IFERROR(F{i}/E{i},0)', NORM, '0%')
  put(al, i, 8, l.get("tags") or None, INP); put(al, i, 9, l.get("text"), INP, wrap=True); put(al, i, 10, str(l.get("cutId", "")))
  al.row_dimensions[i].height = 110
al.freeze_panes = "A5"

# ------------------------------------------------------------------ DASHBOARD
db = wb.create_sheet("DASHBOARD", 1)
titled(db, "BILLING REDUCTION & APPEALS DASHBOARD", f"All figures come from BILLING CUTS. Synced {SYNCED}. Demo data is fictional.")
kpis = [("Total Cuts", f"=SUM({BCR('Cut ($)')})", MONEY0),
        ("Preventable Cuts", f'=SUMIFS({BCR("Cut ($)")},{BCR("Preventability")},"Preventable")', MONEY0),
        ("Preventable Share", "=IFERROR(B6/B5,0)", '0%'),
        ("Appealed", f"=SUM({BCR('Appeal Amount ($)')})", MONEY0),
        ("Recovered", f"=SUM({BCR('Recovered ($)')})", MONEY0),
        ("Recovery Rate", "=IFERROR(B9/B8,0)", '0%'),
        ("Written Off", f"=SUM({BCR('Written Off ($)')})", MONEY0),
        ("Pending Decision", f"=SUM({BCR('Pending ($)')})", MONEY0),
        ("Open lines (needs decision)", f'=COUNTIFS({BCR("Appeal Status")},"Needs Review")+COUNTIFS({BCR("Appeal Status")},"Appeal Drafted")', '0'),
        ("Appeals due in 7 days", f'=COUNTIFS({BCR("Days Left")},">=0",{BCR("Days Left")},"<=7")', '0'),
        ("Appeals overdue", f'=COUNTIFS({BCR("Days Left")},"<0")', '0'),
        ("Reports due in 14 days", f"=COUNTIFS('REPORT DEADLINES'!$B$5:$B${RP_LAST},\">=0\",'REPORT DEADLINES'!$B$5:$B${RP_LAST},\"<=14\")", '0'),
        ("Reports overdue", f"=COUNTIFS('REPORT DEADLINES'!$B$5:$B${RP_LAST},\"<0\")", '0')]
db["A4"] = "Key figures"; db["A4"].font = H2
for i, (k, f1, fmt) in enumerate(kpis, 5):
  db.cell(row=i, column=1, value=k).font = BOLD
  c = db.cell(row=i, column=2, value=f1); c.number_format = fmt; c.font = font(bold=True, size=12)
  db.cell(row=i, column=1).fill = GFILL; c.fill = GFILL
def table(ws, top, left, title, heads, rows):
  ws.cell(row=top, column=left, value=title).font = H2
  for j, h in enumerate(heads):
    c = ws.cell(row=top + 1, column=left + j, value=h); c.font = HDR; c.fill = HFILL
  for i, rowv in enumerate(rows):
    for j, (v, fmt) in enumerate(rowv):
      c = ws.cell(row=top + 2 + i, column=left + j, value=v); c.border = BOX
      if fmt: c.number_format = fmt
  return top + 2, top + 1 + len(rows)
R0 = 20
rows = []
for k, x in enumerate(REASONS):
  rr = R0 + 2 + k
  rows.append([(x["name"], None), (f'=SUMIFS({BCR("Cut ($)")},{BCR("Cut Reason")},A{rr})', MONEY0), (f'=COUNTIFS({BCR("Cut Reason")},A{rr})', '0'),
               (f'=SUMIFS({BCR("Cut ($)")},{BCR("Cut Reason")},A{rr},{BCR("Preventability")},"Preventable")', MONEY0), (f'=SUMIFS({BCR("Recovered ($)")},{BCR("Cut Reason")},A{rr})', MONEY0)])
rs, re_ = table(db, R0, 1, "Cuts by reason", ["Reason", "Cuts ($)", "Lines", "Preventable ($)", "Recovered ($)"], rows)
P0 = 4
prow = [[(pp, None), (f'=SUMIFS({BCR("Cut ($)")},{BCR("Preventability")},G{P0 + 2 + k})', MONEY0), (f'=IFERROR(H{P0 + 2 + k}/$B$5,0)', '0%')] for k, pp in enumerate(PREVS)]
ps, pe = table(db, P0, 7, "Cuts by preventability", ["Preventability", "Cuts ($)", "Share"], prow)
C0 = 10
crow = [[(f"=LOOKUPS!A{2 + k}", None), (f'=SUMIFS({BCR("Cut ($)")},{BCR("Carrier")},G{C0 + 2 + k})', MONEY0), (f'=SUMIFS({BCR("Recovered ($)")},{BCR("Carrier")},G{C0 + 2 + k})', MONEY0),
         (f'=IFERROR(I{C0 + 2 + k}/SUMIFS({BCR("Appeal Amount ($)")},{BCR("Carrier")},G{C0 + 2 + k}),0)', '0%')] for k in range(len(CARRIERS))]
cs, ce = table(db, C0, 7, "Cuts by carrier", ["Carrier", "Cuts ($)", "Recovered ($)", "Recovery Rate"], crow)
T0 = C0 + 3 + len(CARRIERS) + 1
trow = [[(f"=LOOKUPS!C{2 + k}", None), (f'=SUMIFS({BCR("Cut ($)")},{BCR("Timekeeper")},G{T0 + 2 + k})', MONEY0), (f'=SUMIFS({BCR("Cut ($)")},{BCR("Timekeeper")},G{T0 + 2 + k},{BCR("Preventability")},"Preventable")', MONEY0),
         (f'=IFERROR(I{T0 + 2 + k}/H{T0 + 2 + k},0)', '0%')] for k in range(len(TKS))]
ts, te = table(db, T0, 7, "Cuts by timekeeper (training targets)", ["Timekeeper", "Cuts ($)", "Preventable ($)", "Preventable %"], trow)
for col, w in zip("ABCDEFGHIJ", [30, 14, 8, 15, 14, 3, 24, 14, 15, 14]): db.column_dimensions[col].width = w
ch = BarChart(); ch.type = "bar"; ch.style = 10; ch.title = "Cuts by reason ($)"
ch.add_data(Reference(db, min_col=2, min_row=rs - 1, max_row=re_), titles_from_data=True)
ch.add_data(Reference(db, min_col=4, min_row=rs - 1, max_row=re_), titles_from_data=True)
ch.set_categories(Reference(db, min_col=1, min_row=rs, max_row=re_)); ch.height = 14; ch.width = 18; ch.overlap = 100
db.add_chart(ch, "L4")
pie = PieChart(); pie.title = "Preventability"
pie.add_data(Reference(db, min_col=8, min_row=ps - 1, max_row=pe), titles_from_data=True); pie.set_categories(Reference(db, min_col=7, min_row=ps, max_row=pe))
pie.dataLabels = DataLabelList(); pie.dataLabels.showPercent = True; pie.height = 8; pie.width = 12
db.add_chart(pie, "L33")
cb = BarChart(); cb.title = "Cuts vs. recovered by carrier"; cb.style = 10
cb.add_data(Reference(db, min_col=8, max_col=9, min_row=cs - 1, max_row=ce), titles_from_data=True); cb.set_categories(Reference(db, min_col=7, min_row=cs, max_row=ce)); cb.height = 8; cb.width = 14
db.add_chart(cb, "S33")

# ------------------------------------------------------------------ WRITE-OFFS
wo = wb.create_sheet("WRITE-OFFS", 2)
titled(wo, "WRITE-OFF REPORT (FOR FINANCE)", "Written off = denied or accepted cuts plus the unrecovered part of partial wins. Pending = cuts still open or awaiting a carrier decision. Live formulas from BILLING CUTS.")
wo["A3"] = "Reduction dates from"; wo["A3"].font = BOLD
ds = [d(c.get("redDate")) for c in CUTS if d(c.get("redDate"))]
wo["B3"] = min(ds) if ds else dt.date(TODAY.year, 1, 1); wo["B3"].number_format = DATE; wo["B3"].font = INP; wo["B3"].fill = YFILL
wo["C3"] = "to"; wo["D3"] = TODAY; wo["D3"].number_format = DATE; wo["D3"].font = INP; wo["D3"].fill = YFILL
rd = BCR("Reduction Date"); per = f'{rd},">="&$B$3,{rd},"<="&$D$3'
WH = ["Cuts ($)", "Recovered ($)", "Written Off ($)", "Pending ($)", "Write-off % of Cuts"]
def wo_table(top, title, key_hdr, keys, key_col):
  rows = []
  for k, key in enumerate(keys):
    r = top + 2 + k
    crit = f'{BCR(key_col)},$A{r},{per}'
    rows.append([(key, None)] + [(f'=SUMIFS({BCR(h)},{crit})', MONEY0) for h in ["Cut ($)", "Recovered ($)", "Written Off ($)", "Pending ($)"]] + [(f'=IFERROR(D{r}/B{r},0)', '0%')])
  r = top + 2 + len(keys)
  rows.append([("Total", None)] + [(f'=SUM({col}{top + 2}:{col}{r - 1})', MONEY0) for col in "BCDE"] + [(f'=IFERROR(D{r}/B{r},0)', '0%')])
  s, e = table(wo, top, 1, title, [key_hdr] + WH, rows)
  for j in range(1, 7): wo.cell(row=e, column=j).font = BOLD
  return s, e
s1, e1 = wo_table(5, "By carrier", "Carrier", [c["name"] for c in CARRIERS], "Carrier")
s2, e2 = wo_table(e1 + 3, "By timekeeper", "Timekeeper", [t["name"] for t in TKS], "Timekeeper")
s3, e3 = wo_table(e2 + 3, "By preventability", "Preventability", PREVS, "Preventability")
s4, e4 = wo_table(e3 + 3, "By matter", "Matter", [m["name"] for m in MATTERS], "Matter")
s5, e5 = wo_table(e4 + 3, "By cut reason", "Cut Reason", [x["name"] for x in REASONS], "Cut Reason")
for col, w in zip("ABCDEF", [40, 13, 13, 14, 13, 12]): wo.column_dimensions[col].width = w
wch = BarChart(); wch.type = "bar"; wch.grouping = "stacked"; wch.overlap = 100; wch.title = "Written off vs. pending by carrier"; wch.style = 10
wch.add_data(Reference(wo, min_col=4, max_col=5, min_row=s1 - 1, max_row=e1 - 1), titles_from_data=True); wch.set_categories(Reference(wo, min_col=1, min_row=s1, max_row=e1 - 1))
wch.height = 8; wch.width = 16
wo.add_chart(wch, "H5")

# ------------------------------------------------------------------ MONTHLY SUMMARY
ms = wb.create_sheet("MONTHLY SUMMARY", 3)
ms["A1"] = "MONTHLY BILLING REDUCTION SUMMARY"; ms["A1"].font = TITLE
ms["A2"] = "Year"; ms["A2"].font = BOLD; ms["B2"] = TODAY.year; ms["B2"].font = INP; ms["B2"].fill = YFILL
ms["C2"] = "Grouped by Reduction Date. Change the year in B2."; ms["C2"].font = SUB
heads = ["Month", "Billed (flagged lines)", "Cuts ($)", "Preventable ($)", "Preventable %", "Appealed ($)", "Recovered ($)", "Written Off ($)", "Pending ($)", "Recovery %"]
header(ms, 4, heads, [12, 18, 13, 15, 13, 13, 13, 14, 12, 11])
for m in range(1, 13):
  r = 4 + m
  ms.cell(row=r, column=1, value=f"=DATE($B$2,{m},1)").number_format = "mmm yyyy"
  cond = f'{rd},">="&A{r},{rd},"<"&EDATE(A{r},1)'
  vals = [(f"=SUMIFS({BCR('Original Bill ($)')},{cond})", MONEY0), (f"=SUMIFS({BCR('Cut ($)')},{cond})", MONEY0),
          (f'=SUMIFS({BCR("Cut ($)")},{cond},{BCR("Preventability")},"Preventable")', MONEY0), (f"=IFERROR(D{r}/C{r},0)", '0%'),
          (f"=SUMIFS({BCR('Appeal Amount ($)')},{cond})", MONEY0), (f"=SUMIFS({BCR('Recovered ($)')},{cond})", MONEY0),
          (f"=SUMIFS({BCR('Written Off ($)')},{cond})", MONEY0), (f"=SUMIFS({BCR('Pending ($)')},{cond})", MONEY0), (f"=IFERROR(G{r}/F{r},0)", '0%')]
  for j, (v, fmt) in enumerate(vals, 2): ms.cell(row=r, column=j, value=v).number_format = fmt
ms.cell(row=17, column=1, value="Total").font = BOLD
for col in "BCDFGHI":
  c = ms[f"{col}17"]; c.value = f"=SUM({col}5:{col}16)"; c.number_format = MONEY0; c.font = BOLD
ms["E17"] = "=IFERROR(D17/C17,0)"; ms["E17"].number_format = '0%'; ms["J17"] = "=IFERROR(G17/F17,0)"; ms["J17"].number_format = '0%'
lc = LineChart(); lc.title = "Cuts, preventable cuts and recoveries by month"; lc.height = 8; lc.width = 22
lc.add_data(Reference(ms, min_col=3, max_col=4, min_row=4, max_row=16), titles_from_data=True)
lc.add_data(Reference(ms, min_col=7, min_row=4, max_row=16), titles_from_data=True)
lc.set_categories(Reference(ms, min_col=1, min_row=5, max_row=16))
ms.add_chart(lc, "A20")

# ------------------------------------------------------------------ APPEAL DRAFT
ad = wb.create_sheet("APPEAL DRAFT", 4)
ad["A1"] = "APPEAL DRAFT"; ad["A1"].font = TITLE
ad["A2"] = "Type an Entry ID from BILLING CUTS and fill the yellow fact cells. Copy the draft from B16. Blanks show as [brackets] to fill in."; ad["A2"].font = SUB
ad.column_dimensions["A"].width = 26; ad.column_dimensions["B"].width = 110
ad["A4"] = "Entry ID"; ad["A4"].font = BOLD
has_rule = lambda c: any(r["carrier"] == c.get("carrier") and r["topic"] == REASON.get(c.get("reason"), {}).get("topic") for r in RULES)
first_open = next((i + 1 for i, c in enumerate(CUTS) if c.get("status") == "Needs Review" and c.get("reason") in ("Block Billing", "Insufficient Detail") and has_rule(c)),
                  next((i + 1 for i, c in enumerate(CUTS) if c.get("status") == "Needs Review" and has_rule(c)), 1))
ad["B4"] = first_open; ad["B4"].font = INP; ad["B4"].fill = YFILL; ad["B4"].alignment = Alignment(horizontal="left")
for lab, r in [("Corrected narrative", 5), ("Why necessary / what it produced", 6), ("Approved by", 7), ("Approval date", 8)]:
  ad.cell(row=r, column=1, value=lab).font = BOLD
  c = ad.cell(row=r, column=2); c.font = INP; c.fill = YFILL; c.alignment = WRAP
if CUTS:
  ex = CUTS[first_open - 1]
  if ex.get("rewrite"): ad["B5"] = ex["rewrite"]
  if ex.get("scope"): ad["B6"] = ex["scope"]
ad["B8"].number_format = DATE
idr = f"MATCH($B$4,{BCR('Entry ID')},0)"
def g(h): return f"INDEX({BCR(h)},{idr})"
pulled = [("Carrier", g("Carrier")), ("Platform", g("Platform")), ("Invoice #", g("Invoice #")), ("Matter", g("Matter")),
          ("Line", f'TEXT({g("Invoice Date")},"mm/dd/yyyy")&" · "&{g("Timekeeper")}&" ("&{g("TK Role")}&") · "&{g("UTBMS Task")}&"/"&{g("UTBMS Activity")}&" · "&TEXT({g("Hours")},"0.0")&" hrs"'),
          ("Narrative as billed", g("Narrative as Billed")), ("Cut ($) / Reason", f'TEXT({g("Cut ($)")},"$#,##0.00")&" · "&{g("Cut Reason")}'),
          ("Preventability", g("Preventability")), ("Appeal deadline", f'TEXT({g("Appeal Deadline")},"mm/dd/yyyy")')]
ad["D3"] = "Pulled from BILLING CUTS"; ad["D3"].font = BOLD
ad.column_dimensions["D"].width = 20; ad.column_dimensions["E"].width = 70
for i, (lab, f1) in enumerate(pulled, 4):
  ad.cell(row=i, column=4, value=lab).font = BOLD
  c = ad.cell(row=i, column=5, value=f"=IFERROR({f1},\"\")"); c.font = GRN; c.alignment = WRAP
ad["A10"] = "Guideline topic"; ad["B10"] = f"=IFERROR(INDEX('REASON MATRIX'!$C${RM_FIRST}:$C${RM_LAST},MATCH({g('Cut Reason')},'REASON MATRIX'!$A${RM_FIRST}:$A${RM_LAST},0)),\"general\")"
rk = f"'CARRIER GUIDELINES'!$A${RULE_FIRST}:$A${RULE_LAST}"
ad["A11"] = "Section cite"; ad["B11"] = f"=IFERROR(INDEX('CARRIER GUIDELINES'!$E${RULE_FIRST}:$E${RULE_LAST},MATCH({g('Carrier')}&\"|\"&B10,{rk},0)),\"[add guideline cite]\")"
ad["A12"] = "Rule"; ad["B12"] = f"=IFERROR(INDEX('CARRIER GUIDELINES'!$F${RULE_FIRST}:$F${RULE_LAST},MATCH({g('Carrier')}&\"|\"&B10,{rk},0)),\"[add guideline rule]\")"
ad["A13"] = "Argument"
arg = f"IFERROR(INDEX('REASON MATRIX'!$G${RM_FIRST}:$G${RM_LAST},MATCH({g('Cut Reason')},'REASON MATRIX'!$A${RM_FIRST}:$A${RM_LAST},0)),\"\")"
ad["B13"] = ("=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" + arg +
  ',"{rewrite}",IF(B5="","[corrected narrative]",B5)),"{scope}",IF(B6="","[facts showing why the work was necessary and what it produced]",B6)),"{approver}",IF(B7="","[approver name]",B7)),"{approvalDate}",IF(B8="","[approval date]",TEXT(B8,"mm/dd/yyyy"))),"{invoice}",' + g("Invoice #") + '),"{section}",B11)')
for r in (10, 11, 12, 13): ad.cell(row=r, column=1).font = BOLD; ad.cell(row=r, column=2).alignment = WRAP; ad.cell(row=r, column=2).font = GRN
ad["A15"] = "DRAFT APPEAL"; ad["A15"].font = H2
NL = "CHAR(10)"
ad["B16"] = ('=IFERROR("RE: Request for reconsideration of invoice reductions"&' + NL + '&"Carrier: "&E4&" ("&E5&")"&' + NL + '&"Matter: "&E7&"   Invoice: "&E6&' + NL +
  '&' + NL + '&"We respectfully request reconsideration of the reduction below. Each line is addressed individually with the applicable guideline provision; supporting documentation is attached where noted."&' + NL + '&' + NL +
  '&"Line 1: "&E8&' + NL + '&"Billed narrative: "&E9&' + NL + '&"Reduced "&E10&' + NL +
  '&"Response: "&B13&' + NL + '&"Guideline: "&E4&" "&B11&". "&B12&' + NL + '&"Request: restore "&TEXT(' + g("Cut ($)") + ',"$#,##0.00")&"."&' + NL + '&' + NL +
  '&"Thank you for reconsidering this item. Please contact our billing department with any questions.","Enter a valid Entry ID in B4.")')
ad["B16"].alignment = WRAP; ad.row_dimensions[16].height = 300
ad["A18"] = "CLAUDE PROMPT"; ad["A18"].font = H2
ad["B19"] = ('=IFERROR("You are helping the billing compliance team at an insurance defense law firm appeal a reduction a carrier bill reviewer made to our invoice."&' + NL + '&' + NL +
  '&"CARRIER: "&E4&" | PLATFORM: "&E5&" | APPEAL DEADLINE: "&E12&' + NL +
  '&"GUIDELINE (quote this; do not invent others): "&B11&" - "&B12&' + NL + '&' + NL +
  '&"REDUCED LINE: "&E8&' + NL + '&"Narrative as billed: "&E9&' + NL + '&"Reduced "&E10&' + NL +
  '&"Facts I can support: "&IF(B5="","","corrected narrative: "&B5&"; ")&IF(B6="","",B6&"; ")&IF(B7="","","approved by "&B7&IF(B8="",""," on "&TEXT(B8,"mm/dd/yyyy")))&' + NL + '&' + NL +
  '&"WHAT I NEED: (1) A concise appeal I can paste into "&E5&" that states the reason, explains why it should be restored, cites the guideline above, and lists attachments. (2) Do not restate the original entry; if the narrative was vague or block billed, give a corrected narrative using only my facts, with time per task. (3) If the guideline clearly prohibits the task, tell me plainly and recommend accepting the cut. (4) Add one prevention tip for the timekeeper. (5) Factual, respectful tone.","")')
ad["B19"].alignment = WRAP; ad.row_dimensions[19].height = 260

order = ["INSTRUCTIONS", "DASHBOARD", "WRITE-OFFS", "MONTHLY SUMMARY", "APPEAL DRAFT", "BILLING CUTS", "REPORT DEADLINES", "MATTERS", "APPROVALS", "APPEAL LIBRARY",
         "CARRIER GUIDELINES", "REPORTING RULES", "REASON MATRIX", "INTAKE MAPPING", "LOOKUPS"]
wb._sheets = [wb[n] for n in order]
wb.active = 0
for s in wb.worksheets:
  s.sheet_view.showGridLines = s.title in ("BILLING CUTS", "LOOKUPS")
wb.save(OUT)
print("saved", OUT, "sheets:", wb.sheetnames, "cuts:", len(CUTS), "report items:", len(items))
