# Appeals Tracker: Handoff

**For:** Elizabeth, Billing Compliance Manager
**What you're getting:** a web tracker you manage yourself (no Excel editing needed), a design for keeping an Excel workbook in sync with it, and a library of copy-paste prompts for your firm's Claude.

**Try the demo:** https://joshmay-hcs.github.io/appeals-tracker/
**Code and files:** https://github.com/joshmay-hcs/appeals-tracker
**Everything in the demo is fictional.** Carrier names, people, matters, claims and dollar amounts are made up. The billing *rules* come from the five guideline sets you shared.

---

## 1. What's in the folder

| File | What it is | Use it for |
|---|---|---|
| `index.html` | The demo, live at https://joshmay-hcs.github.io/appeals-tracker/ | Trying everything out. Your changes save only in your own browser, so nothing you do affects anyone else. **Reset to demo data** (Guidelines & Admin > Data & Excel) starts over. |
| `appeals-tracker.html` | The same app, built to run inside Claude with a shared database | Going live in your firm's Claude (Section 8): everyone with access sees the same data, and the Claude buttons work. |
| `tools/sync_workbook.py` | The script Claude uses to rebuild the Excel workbook from the tracker's data | Nothing for you to run. Claude uses it when you ask for a sync (Claude-hosted version). |
| `HANDOFF.md` | This guide | How it works, how to manage it, the prompts, and how to move it into your firm's Claude. |

---

## 2. The big idea: two kinds of cuts

Every reduction falls into one of three buckets. The tracker sorts them for you based on the cut reason, and you can override any single line.

| Bucket | Meaning | What you do |
|---|---|---|
| **Preventable** | The firm caused it: block billing, vague narratives ("attention to file"), clerical work billed, missing receipts, approval not noted on the line, duplicate entries. | **Train and pre-check.** Most of these can still be appealed once with a corrected narrative, but the real win is stopping them before the bill goes out. |
| **Potentially preventable** | Judgment calls: "excessive" time, research over the cap, two attorneys at one event, budget overruns. | **Appeal with facts** (page counts, approvals, what the work produced) and tighten narratives going forward. |
| **Not preventable** | The reviewer's call or policy: rate disputes, general "outside guidelines" objections, reviewer mistakes (e.g., a "duplicate" that wasn't). | **Appeal.** These are the fights worth having. |

In the demo data, preventable and potentially preventable cuts make up most of the dollars. That is the pattern carriers' AI review tools are built to find, so it's where training pays off fastest.

---

## 3. The workflow

**Before the bill goes out**
1. **Pre-bill Scrub:** paste or upload the draft invoice (LEDES export or a copied table). The tracker flags block billing, vague narratives, clerical work, first-year associates, two attorneys at one event, research over the cap, travel over the approval limit and days over the hour cap, using that carrier's rules and the approvals you have on file. Fix the flagged lines, then send.

**When a reduction report arrives**
2. **Intake:** drop in the reduction report PDF from Legal Solutions Suite or Legal-X, or paste its text. The tracker reads each reduced line, maps the reviewer's category to one of your cut reasons, skips duplicates, and adds the lines to the Cut Log as "Needs Review". If a report's layout isn't recognized, **Read with Claude** pulls the lines out instead.
3. **Decide within the window.** The appeal deadline is set from the reduction date plus the carrier's window (14 to 60 days in the demo). The Dashboard shows what's due in 7 days and anything overdue.
4. **Draft the appeal** in the Appeal Builder. It shows the carrier's rule and section, any approvals on file for that matter, and appeals that already worked for the same carrier and reason. Answer **every line separately** and cite the carrier's own section.
5. **Submit and attach proof:** approval emails, receipts, corrected narratives.
6. **Record the result:** status, round, amount appealed, amount recovered. Wins with appeal text are added to the Appeal Library automatically.

**Ongoing**
7. **Report Deadlines:** carrier reporting deadlines (initial reports, quarterly updates, pre-mediation and pre-trial reports) per matter. Mark each done when it's sent.
8. **Monthly:** review the Dashboard and Training tab with the attorneys, and send finance the Write-offs view (or the synced workbook).

---

## 4. What's on each tab

| Tab | What it does |
|---|---|
| **Dashboard** | Totals, preventable share, recovery rate, appeal and report deadlines, charts by reason, carrier, timekeeper and month. |
| **Cut Log** | Every reduced line. Filter by carrier, matter, platform, reason, preventability, status or timekeeper. Click a row to edit. Tick several rows from one invoice to build one combined appeal. |
| **Intake** | Reads reduction reports (PDF or pasted text) into the Cut Log. |
| **Appeal Builder** | Rule + cite, appeal odds, deadline, approvals on file, "what worked before", a draft appeal, **Copy prompt for Claude**, and **Polish with Claude** (rewrites the draft inside the page). |
| **Appeal Library** | Appeals that won money back, searchable by carrier and reason, ready to reuse. |
| **Approvals** | Prior approvals on file: research over the cap, experts, travel, extra attendees. Each has a scope, approver, date, hour cap and expiry. The Pre-bill Scrub and Appeal Builder both use them. |
| **Pre-bill Scrub** | Checks a draft invoice before it goes out (see the workflow above). **One entry** mode checks a single time entry. **Rewrite with Claude** fixes flagged narratives. |
| **Training** | Built from your actual cuts: a scorecard per timekeeper, before/after examples of rewritten narratives, a printable tip sheet, and a quiz Claude writes from real cuts. |
| **Report Deadlines** | Carrier reporting deadlines per matter, worked out from assignment, mediation and trial dates. Add and edit matters here. |
| **Write-offs** | For finance: what was lost after appeals. Group by carrier, matter, matter type, timekeeper, reason, preventability or month. **Copy as CSV** pastes into Excel. |
| **Guidelines & Admin** | Where you manage everything (Section 5). |

**Guided tour:** opens the first time someone visits (17 short steps). Replay it with **Take the tour** at the top right. Arrow keys move between steps; Esc closes it.

**Claude features inside the page** (Read with Claude, Polish with Claude, Rewrite with Claude, Import from guideline, Training quiz) appear only in the Claude-hosted version, and ask each person for permission the first time they're used. The GitHub demo hides them; **Copy prompt for Claude** works everywhere. Everything else works without them.

---

## 5. Managing the tracker (Guidelines & Admin)

You manage all of the tracker's settings from the **Guidelines & Admin** tab. You never need to edit Excel or code.

| Section | What you can change |
|---|---|
| **Carriers** | Name, platform, appeal window and rounds, billing cadence, and the caps the Pre-bill Scrub enforces: daily hours, research hours, block billing allowed under X hours, travel approval threshold, receipt threshold, whether partner research needs approval, whether first-year associates are paid. |
| **Guideline rules** | Each carrier's rule and section cite by topic. When a carrier updates its guidelines, use **Save as new version** with the effective date. Appeals for older invoices still cite the rule that applied then. **Import from guideline** lets Claude read a new guideline and suggest the rules for you to review. |
| **Reporting rules** | Each carrier's reporting schedule: report name, trigger (assignment, mediation, trial, or recurring), and days before or after. |
| **Cut reasons** | Your 27 reasons: default preventability, appeal odds, usual cause, prevention tip, appeal template. Changing a reason's preventability updates every cut that hasn't been individually overridden. |
| **Timekeepers** | Name, role, rate, email, first-year flag, active. |
| **Intake mapping** | Which reviewer wording maps to which cut reason (e.g., "previously billed" → Duplicate Billing). Add a row when a report uses wording the tracker doesn't recognize. |
| **Data & Excel** | **Reset to demo data**, **Clear demo cuts and start fresh**, Excel sync instructions, and a data-privacy reminder. |

**Who can change what** (Claude-hosted version; in the GitHub demo everyone can change everything, in their own browser only). Access is controlled from the page's **Share** menu:
- **Editor:** can change everything on Guidelines & Admin. Give this to you (and a backup).
- **Contributor:** can log cuts, draft appeals, add approvals and matters, and mark reports done, but can't change carriers, rules, reasons, timekeepers or mapping.
- **Viewer:** read-only.

---

## 6. Keeping the Excel workbook in sync

*This applies once the tracker runs in Claude (Section 8). The workbook isn't in the GitHub repo; ask whoever shared this demo with you for a copy built from the demo data.*

The tracker is the source of truth. The workbook is a copy, rebuilt on request.

**To sync:** in Claude, say *"Sync the Excel workbook from the appeals tracker."* Claude exports every table from the tracker and rebuilds `Appeals_Tracker.xlsx` with fresh formulas, then checks the totals. It takes about a minute. The INSTRUCTIONS tab shows when it was last synced.

**What's in the workbook:**

| Tab | Purpose |
|---|---|
| INSTRUCTIONS | Last sync time, what each tab is for, color legend. |
| DASHBOARD | Key figures (including written off, pending, appeals and reports due), cuts by reason, preventability, carrier and timekeeper, with charts. |
| WRITE-OFFS | For finance: cuts, recovered, written off and pending by carrier, timekeeper, preventability, matter and reason. Set the date range in B3:D3. |
| MONTHLY SUMMARY | Month-by-month cuts, preventable cuts, appeals, recoveries and write-offs. Change the year in B2. |
| APPEAL DRAFT | Type an Entry ID, fill the yellow facts, copy the draft (B16) or the Claude prompt (B19). |
| BILLING CUTS | One row per reduced line, with formulas for platform, cut, preventability, deadline, days left, written off and pending. |
| REPORT DEADLINES | Every reporting deadline per matter; Days Left counts down live. |
| MATTERS | Matters with dates, lead attorney, and live cut and write-off totals. |
| APPROVALS | Prior approvals with Active/Expired status. |
| APPEAL LIBRARY | Winning appeals with full text. |
| CARRIER GUIDELINES | Carrier caps, plus every rule version with its effective date (the current one is marked). |
| REPORTING RULES, REASON MATRIX, INTAKE MAPPING, LOOKUPS | Reference tables, same as in the tracker. |

Edits typed into the workbook are **replaced at the next sync**, so make changes in the tracker.

---

## 7. Prompt library for your firm's Claude

These are the prompts that get good results. The pattern is always the same: **give Claude the carrier's rule, the exact line, and your facts; tell it what to produce; tell it what not to do.** Claude only knows what you paste in, so the guideline text matters most.

> **Before you paste:** follow your firm's policy on what can go into its Claude. Several carrier guidelines also have AI-use clauses (one says counsel "shall not enter confidential or personal data into AI tools"). Strip claimant names and use claim numbers or initials unless your firm has cleared it.

### Prompt A: Appeal one reduced line
```
You are helping the billing compliance team at an insurance defense firm appeal a reduction a carrier's bill reviewer made to our invoice.

CARRIER: [name]   PLATFORM: [Legal Solutions Suite / Legal-X]   APPEAL DEADLINE: [date]
GUIDELINE (quote this; do not invent others): [section number] - [paste the rule text]

REDUCED LINE:
Date: [date] | Timekeeper: [name, role, rate] | UTBMS: [task/activity] | Hours: [x.x]
Narrative as billed: "[exact text]"
Billed $[x] | Allowed $[x] | Reduced $[x]
Reviewer's reason: [exactly as on the report]
Facts I can support: [approval from whom and when / page counts / what the work produced / corrected narrative]

WHAT I NEED:
1. A concise appeal I can paste into the platform: state the reason, explain why it should be restored, cite the section, and list what's attached.
2. Don't restate the original entry. If it was vague or block billed, give a corrected narrative with time per task, using only my facts.
3. If the guideline clearly prohibits the task, say so plainly and recommend accepting the cut.
4. One prevention tip for the timekeeper.
5. Factual, respectful tone. No fairness arguments; only guideline compliance and facts.
```
**Good output looks like:** one short paragraph per line, a section number in each, a clear dollar request, an attachment list, and an honest "don't appeal this one" when that's the right call.

### Prompt B: Appeal a whole reduction report
```
Below is a reduction report from [carrier] on [platform], and the carrier's billing guideline sections that apply.
For EACH reduced line, write a separate response with: line number and date, reduction reason, our response, the guideline section that supports it, and the amount requested.
Group lines I should NOT appeal at the end with one sentence each on why.
Total the amount requested.
Appeal deadline: [date]. Rounds allowed: [e.g., 2 - reviewer, then adjuster].

GUIDELINE SECTIONS:
[paste]

REDUCTION REPORT:
[paste the report text]

FACTS FOR SPECIFIC LINES:
[line #: facts, approvals, page counts]
```

### Prompt C: Fix narratives before the bill goes out (prevention)
```
Rewrite these time entries so they pass [carrier]'s billing guidelines. Rules:
- One task per line (split block-billed entries and divide the time; show your split for me to confirm).
- Who / what / why: name the person or document, the specific task, and its purpose. Include page counts for reviews.
- Flag any entry that is clerical or administrative (scheduling, calendaring, filing, uploading, cover letters, voicemails). Those should be removed, not reworded.
- Flag paralegal-level tasks billed by an attorney.
- Travel: origin, destination, purpose, miles, billed at 50%.
- Don't invent facts. If you need a detail, put it in [brackets] for me to fill.

Carrier-specific rules: [paste, e.g., "research over 3 hrs per topic needs approval"; "no time in parentheses inside one line"]

ENTRIES:
[date | timekeeper | hours | narrative]
```

### Prompt D: Pre-bill review of a draft invoice
```
Act as the carrier's bill reviewer. Review this draft invoice against the guideline text below and list every line you would reduce, with: line, reason, guideline section, the amount you'd cut, and the fix. Then give a corrected version of each flagged line. Sort by dollars at risk.

GUIDELINE: [paste the billing sections]
DRAFT INVOICE: [paste or attach the LEDES/PDF text]
```

### Prompt E: Monthly preventable-cuts summary for attorneys
```
Here are this month's billing cuts (pasted from the tracker). Write a one-page summary for the attorneys:
1. Total cut, preventable share, recovery rate, compared with last month: [numbers].
2. The top 3 preventable habits by dollars, each with one real (anonymized) example entry and a corrected version.
3. Per-timekeeper notes, constructive and brief.
4. Upcoming appeal deadlines.
Plain language; no blame.

DATA: [paste the CSV from the Cut Log or Write-offs tab]
```

### Prompt F: Add a new carrier's guidelines to the tracker
```
Read this carrier billing guideline and extract, as a table: topic | section number | the rule in one or two sentences.
Topics: block billing, vague narratives, clerical/overhead, paralegal-level work, duplicates, research limits, travel, multiple attorneys/timekeepers, intra-office conferences, file/document review, receipt threshold, prior approval (what needs it and how to note it), rates, budget, daily hour cap, nominal tasks, UTBMS coding, correspondence, staffing, scope/phase.
Also give: e-billing platform, appeal window (days and from what event), number of appeal rounds, billing frequency.
If the guideline doesn't address a topic, write "not addressed". Quote section numbers exactly.

GUIDELINE: [attach or paste]
```
Review the result, then enter it under Guidelines & Admin > Guideline rules (or use **Import from guideline**, which runs this for you).

### Tips that make the biggest difference
- **Always include the guideline text.** Without it, Claude guesses at rules.
- **Paste the narrative exactly as billed** and the reviewer's reason exactly as written.
- **Give real facts** (approver and date, page counts, what was produced). Ask Claude to leave [brackets] for anything missing rather than filling gaps.
- **Ask for a "don't appeal" list.** Appealing a clearly excluded task (e.g., scheduling) wastes a round and credibility.
- **One appeal per carrier.** Never mix carriers in one letter.

---

## 8. Moving this into your firm's Claude

When you're ready to run it on real data in your firm's Claude organization:

1. **Create a Project** called "Billing Appeals".
2. **Upload as project knowledge:** each carrier's billing guidelines (one file per carrier, named by carrier), the REASON MATRIX and CARRIER GUIDELINES tabs from the synced workbook (export each as CSV), and a handful of winning appeals from the APPEAL LIBRARY tab.
3. **Paste these project instructions:**
   ```
   You support an insurance defense firm's billing compliance team. When given a reduction report or reduced line:
   - Identify the carrier and use ONLY that carrier's guideline from project knowledge. Quote section numbers exactly.
   - Respond line by line. Never restate the original entry as the argument.
   - Classify each cut as Preventable, Potentially Preventable or Not Preventable using the reason matrix, and say whether it's worth appealing.
   - Never invent facts, approvals, dates or page counts; use [brackets] for anything missing.
   - Note the appeal deadline using the carrier's window.
   - Keep a factual, respectful tone.
   ```
4. Use Prompts A-F inside that project. You'll no longer need to paste the guideline text, because Claude will look it up.
5. Publish the tracker from your firm's own Claude account so the data lives in the firm's workspace, and share it with the billing team (Editor for you, Contributor for the team).

---

## 9. Going live with real data

- **Start clean:** Guidelines & Admin > Data & Excel > **Clear demo cuts and start fresh** (click twice to confirm). Carriers, rules, reasons and timekeepers stay; cuts, matters, approvals and saved appeals are removed.
- **Carriers:** rename the fictional carriers and set each appeal window. Two windows in the demo are assumptions: Keystone's 30 days (the guideline doesn't state one) and Three Rivers' 14 days. Confirm both.
- **Timekeepers and rates:** Guidelines & Admin > Timekeepers.
- **Privacy:** the GitHub demo is public and keeps data only in the visitor's browser. Don't type real client data into it. Before entering real claim data anywhere, confirm with your firm that its Claude workspace is approved for client billing data. Several carrier guidelines also have AI-use clauses (one says counsel "shall not enter confidential or personal data into AI tools"). Use claim numbers or initials rather than claimant names until that's cleared.
- **Firm name and logo:** the demo deliberately doesn't use the firm's name or logo. Add the firm name under Data & Excel (it shows in the header), and IT can add the logo to the HTML. The colors and fonts already follow the firm's website.

---

## Appendix for IT (short)

- **App:** a single HTML file. `index.html` is the standalone demo (demo data embedded, saved per browser in `localStorage` key `appealsTracker.pages.v1`), hosted on GitHub Pages. `appeals-tracker.html` is the Claude-artifact build. Vanilla JavaScript; Chart.js 4.4.1 and pdf.js 3.11.174 from cdnjs (pdf.js loads only when a PDF is opened); Google Fonts (Bitter, Roboto, Roboto Mono).
- **Data (Claude-hosted):** stored in the artifact's shared database, so everyone with access sees the same records. Collections: `carriers`, `rules` (one document per carrier + topic + effective date), `reportRules`, `reasons`, `timekeepers`, `mapping`, `matters` (with `reportsDone`), `cuts`, `approvals`, `library`, and `meta/settings`. Write access to `carriers`, `rules`, `reportRules`, `reasons`, `timekeepers`, `mapping` and `meta` is limited to Editors; Contributors can write the rest. Only the tour's "seen" flag is stored per browser.
- **Cut record:** id, n, carrier, platform, matterId, matter, claim, invoice, invDate, redDate, tk, task, activity, narrative, rewrite, hours, rate, billed, allowed, cut, reason, prev, prevOverride, status, round, appealed, recovered, appealText, scope, notes, source. Appeal deadline = redDate + the carrier's appealDays. Rule lookup uses the latest version with effective date on or before the invoice date.
- **Written off / pending:** Fully Restored → 0 written off; Partially Restored → cut − recovered; Denied or Accepted → cut − recovered; anything else is pending.
- **Workbook sync:** Claude exports each collection to JSON and runs `tools/sync_workbook.py <export_dir> <out.xlsx>` (Python 3 + openpyxl), then recalculates in Excel. INDEX/MATCH and SUMIFS only, so it works in Excel 2010+. Last sync: 7,267 formulas, 0 errors, totals matching the tracker.
- **Statuses:** Needs Review, Appeal Drafted, Submitted, Partially Restored, Fully Restored, Denied, Accepted - No Appeal. "Open" means Needs Review or Appeal Drafted.
- **Rebuilding elsewhere** (Dataverse/Power Apps, SharePoint lists, or a database): the collections above are the whole data model; each maps to one table.
