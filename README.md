# Billing Appeals Tracker (demo)

A demo tracker for an insurance-defense billing team. It logs carrier bill-review cuts, sorts them into preventable and non-preventable, drafts appeals that cite the carrier's own billing guidelines, and tracks results.

**All data in this demo is fictional.** Carriers, people, matters, claims and amounts are made up.

## Start here

- **[HANDOFF.md](HANDOFF.md):** the full guide. Covers the workflow, every tab, managing the tracker, Excel sync, the prompt library for Claude, and going live.
- **Live demo:** the tracker runs as a Claude artifact with a shared database. Ask the owner for access to the link in HANDOFF.md.

## Files

| File | What it is |
|---|---|
| `appeals-tracker.html` | The tracker app (a single HTML file). Its data comes from the artifact's shared database, so opening the file locally shows an empty tracker. |
| `HANDOFF.md` | Plain-English guide and prompt library. |
| `tools/sync_workbook.py` | Rebuilds the Excel workbook from a JSON export of the tracker's database (`python tools/sync_workbook.py <export_dir> <out.xlsx>`; needs openpyxl). |
