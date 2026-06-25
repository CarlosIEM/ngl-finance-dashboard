# CLAUDE.md

This file provides guidance to Claude Code when working in this repository.

## Project overview

Weekly finance dashboard for NGL (Natural Grass Land LLC) — a tallow and UCO rendering company in Mongolia. Reads a weekly budget vs actual Excel report and generates a single-page HTML dashboard showing cash in/out, budget variance, and major line items.

## How to run

**Install dependencies**

```powershell
pip install -r requirements.txt
```

**Run**

```powershell
python dashboard.py
```

Drop the weekly finance Excel file into `data/` first. Output writes to `output/dashboard.html`.

## File naming

Expected filename pattern: `NGL_finance_report_*.xlsx` or `NGL finance report *.xlsx` (most recent file in `data/` is used automatically).

## Architecture

- `dashboard.py` — single script; no web server, no database
- `data/` — Excel source files (git-ignored; drop new file here each week)
- `output/` — generated HTML (git-ignored; open in browser)
- Column detection is fuzzy (case-insensitive, multiple candidate names) to survive minor Excel header changes
- Charts rendered client-side via Chart.js (inlined for self-contained output)
- Excel serial dates converted via `datetime(1899, 12, 30) + timedelta(days=serial)`

## Weekly update workflow

1. Receive updated Excel from finance team
2. Drop file into `data/`
3. Run `python dashboard.py`
4. Open `output/dashboard.html` in browser

## Version control

Commit and push after each meaningful change. Excel data files and HTML output are git-ignored — only the script and config are tracked.
