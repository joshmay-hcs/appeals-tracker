# Billing Appeals Tracker (demo)

A demo tracker for an insurance-defense billing team. It logs carrier bill-review cuts, sorts them into preventable and non-preventable, drafts appeals that cite the carrier's own billing guidelines, and tracks results.

**[Open the live demo →](https://joshmay-hcs.github.io/appeals-tracker/)**

**All data in this demo is fictional.** Carriers, people, matters, claims and amounts are made up. Changes you make save only in your own browser. To start over, go to Guidelines & Admin → Data & Excel → Reset to demo data.

## Start here

- **[HANDOFF.md](HANDOFF.md):** the full guide. Covers the workflow, every tab, managing the tracker, Excel sync, the prompt library for Claude, and going live.

## Files

| File | What it is |
|---|---|
| `index.html` | The standalone demo served by GitHub Pages. Demo data is built in and saved per browser. |
| `appeals-tracker.html` | The same app built to run as a Claude artifact with a shared database and in-page Claude features. |
| `HANDOFF.md` | Plain-English guide and prompt library. |
| `tools/sync_workbook.py` | Rebuilds the Excel workbook from a JSON export of the tracker's database (`python tools/sync_workbook.py <export_dir> <out.xlsx>`; needs openpyxl). |
