#!/usr/bin/env python3
"""NGL Finance Dashboard — CFO/CEO edition.

Drop files in data/ → python dashboard.py → open output/dashboard.html
"""

import sys
import json
from pathlib import Path
from datetime import datetime
import pandas as pd

DATA_DIR       = Path(__file__).parent / "data"
OUTPUT_DIR     = Path(__file__).parent / "output"
CHART_JS_LOCAL = Path(__file__).parent / "chartjs.min.js"
CHART_JS_CDN   = "https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"

PLAN_UNITS = 1_000   # annual plan is in MNT thousands
NGL_BANKS  = {"trade development bank", "tdb", "khan bank", "khanbank",
               "хаан банк", "худалдаа хөгжлийн банк"}
MN_MONTHS  = ["1-р сар", "2-р сар", "3-р сар", "4-р сар", "5-р сар", "6-р сар",
               "7-р сар", "8-р сар", "9-р сар", "10-р сар", "11-р сар", "12-р сар"]
EN_MONTHS  = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

MONTH_KEY = {
    "january": (1, "Jan"), "february": (2, "Feb"), "march": (3, "Mar"),
    "april":   (4, "Apr"), "may":       (5, "May"), "june":  (6, "Jun"),
    "july":    (7, "Jul"), "august":    (8, "Aug"), "september": (9, "Sep"),
    "october": (10, "Oct"), "november": (11, "Nov"), "december": (12, "Dec"),
}

# Raw-data keywords whose values are NEGATIVE = inflow to NGL (sign-flipped for display)
INFLOW_KW = ["fund from", "income of export", "income various",
             "interest income", "ot income ("]

# Display category assignment (evaluated in order; first match wins)
CATEGORIES = [
    ("Inflows",      ["fund from", "income of export", "income various",
                      "interest income", "ot income ("]),
    ("Stock",        ["stock"]),
    ("Fixed assets", ["fixed asset"]),
    ("Taxes",        ["tax"]),
    ("Loans",        ["loan"]),
    ("Labour",       ["labour", "social insurance"]),
    ("Opex",         ["light, power", "fuel expense", "kitchen expense",
                      "phone", "postal", "ot expense"]),
]

# Clean display names for known account codes
DISPLAY_NAMES = {
    "fund from iem":                          "IEM funding",
    "fund from iepl":                         "IEPL funding",
    "income of export":                       "Export income",
    "income various":                         "Other income",
    "interest income":                        "Interest income",
    "ot income (oyu tolgoi llc)":             "Oyu Tolgoi income",
    "ot expense (oyu tolgoi llc)":            "Oyu Tolgoi expense",
    "stock - fat collected":                  "Fat collection",
    "stock - uco collected":                  "UCO collection",
    "stock - production to export":           "Production to export",
    "tax- fat collection / 2024-2025 /":      "Fat collection tax",
    "tax - fat collection":                   "Fat collection tax",
    "tax - corporate income tax":             "Corporate income tax",
    "tax - wage deducted tax":                "Wage deducted tax",
    "fixed asset - gta inprovement, equipment": "Fixed assets & equipment",
    "fixed asset - gta improvement, equipment": "Fixed assets & equipment",
    "loan khanbank":                          "Loan – Khan Bank",
    "loan dbm leasing (+court cost)":         "Loan – DBM leasing",
    "loan to employees & others":             "Loans to employees",
    "labour cost - wage":                     "Wages & salaries",
    "labour cost - social insurance cost":    "Social insurance",
    "light, power, heating expense":          "Utilities (light/heat)",
    "fuel expense":                           "Fuel",
    "employee kitchen expense":               "Kitchen",
    "phone internet postal":                  "Communications",
}


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def safe_float(v):
    try:
        f = float(v)
        return 0.0 if (f != f) else f   # NaN check
    except (TypeError, ValueError):
        return 0.0


def fmt_mnt(v, decimals=1):
    try:
        v = float(v)
        if v == 0:
            return "0"
        sign = "−" if v < 0 else ""
        av = abs(v)
        if av >= 1_000_000_000:
            return f"{sign}{av / 1_000_000_000:.{decimals}f}B"
        if av >= 1_000_000:
            return f"{sign}{av / 1_000_000:.{decimals}f}M"
        if av >= 1_000:
            return f"{sign}{av / 1_000:.0f}k"
        return f"{sign}{av:.0f}"
    except (TypeError, ValueError):
        return "—"


def fmt_usd(v, decimals=0):
    try:
        v = float(v)
        av = abs(v)
        sign = "−" if v < 0 else ""
        if av >= 1_000_000:
            return f"USD {sign}{av / 1_000_000:.2f}m"
        if av >= 1_000:
            return f"USD {sign}{av / 1_000:.{decimals}f}k"
        return f"USD {sign}{av:,.{decimals}f}"
    except (TypeError, ValueError):
        return "—"


def label_has(s, *kw):
    if not isinstance(s, str):
        return False
    sl = s.lower()
    return any(k.lower() in sl for k in kw)


def categorize(label):
    l = label.lower()
    for cat, kws in CATEGORIES:
        if any(k in l for k in kws):
            return cat
    return "Other"


def item_is_inflow(label):
    return any(k in label.lower() for k in INFLOW_KW)


def display_name(label):
    return DISPLAY_NAMES.get(label.lower(), label)


def find_col(df, *candidates):
    norm = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = str(cand).strip().lower()
        if key in norm:
            return norm[key]
    for cand in candidates:
        key = str(cand).strip().lower()
        for col_key, col_name in norm.items():
            if key in col_key:
                return col_name
    return None


def clean_cols(df):
    df.columns = [" ".join(str(c).split()) for c in df.columns]
    return df


# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------

def _is_weekly(f):
    n = f.name.lower()
    return (any(k in n for k in ["budget", "actual"])
            and not any(k in n for k in ["annual", "plan_final", "bank"]))


def find_files():
    all_xlsx = list(DATA_DIR.glob("*.xlsx"))
    weekly = sorted([f for f in all_xlsx if _is_weekly(f)],
                    key=lambda f: f.stat().st_mtime)
    plan   = sorted([f for f in all_xlsx
                     if any(k in f.name.lower()
                            for k in ["annual_plan", "annual plan", "plan_final"])],
                    key=lambda f: f.stat().st_mtime)
    bank   = sorted([f for f in all_xlsx
                     if any(k in f.name.lower()
                            for k in ["bank_balance", "bank balance"])],
                    key=lambda f: f.stat().st_mtime)
    if not weekly:
        sys.exit("ERROR: No weekly budget/actual file found in data/")
    if not plan:
        sys.exit("ERROR: No annual plan file found in data/")
    if not bank:
        sys.exit("ERROR: No bank balance file found in data/")
    print(f"  Weekly files ({len(weekly)}):")
    for f in weekly:
        print(f"    {f.name}")
    print(f"  Plan : {plan[-1].name}")
    print(f"  Bank : {bank[-1].name}")
    return weekly, plan[-1], bank[-1]


def _col_info(path):
    """Return (col_label, sort_key) from filename."""
    stem = path.stem.lower().replace(" ", "_")
    import re
    # Week-level June files: "Jun 1–7", "Jun 1–14", "Jun 1–21"
    m = re.search(r"june_(\d{1,2})_actual", stem)
    if m:
        day = int(m.group(1))
        return f"Jun 1–{day}", (6, day)
    # Monthly files: full-month label e.g. "Jan 2026"
    parts = stem.split("_vs_")
    if len(parts) >= 2:
        actual_part = parts[1]
        for mk, (num, abbr) in MONTH_KEY.items():
            if mk in actual_part:
                return f"{abbr} 2026", (num, 0)
    return path.stem[:6], (99, 0)


# ---------------------------------------------------------------------------
# Raw actual sheet parser (col 0=label, col 1=actual, col 2=budget)
# ---------------------------------------------------------------------------

def parse_raw_sheet(path):
    """Read the raw Actual sheet: returns {label: {actual, budget}}."""
    xl = pd.ExcelFile(path)
    # Prefer a dedicated "[Month] Actual" sheet over "Budget vs Actual Summary"
    actual_sheets = [s for s in xl.sheet_names if "actual" in s.lower()]
    non_summary   = [s for s in actual_sheets
                     if "summary" not in s.lower() and "budget" not in s.lower()]
    sheet = non_summary[0] if non_summary else (actual_sheets[0] if actual_sheets else xl.sheet_names[0])
    df = clean_cols(pd.read_excel(path, sheet_name=sheet, header=0))
    df = df.dropna(how="all")
    if df.shape[1] < 2:
        return {}, sheet

    label_col  = df.columns[0]
    actual_col = df.columns[1]
    budget_col = df.columns[2] if df.shape[1] > 2 else None

    # Row labels that are aggregation/total rows, not line items
    SKIP_LABELS = {"cash", "total", "net", "net cash", "opening balance",
                   "closing balance", "total cash flow", "sub total", "subtotal"}

    rows = {}
    for _, row in df.iterrows():
        lbl = " ".join(str(row.get(label_col, "")).split()).strip()
        if not lbl or lbl.lower() in ("nan", "none", "account code"):
            continue
        ll = lbl.lower()
        if ll in SKIP_LABELS:
            continue
        # Skip any aggregate row that starts with "total" or "grand total"
        if ll.startswith(("total ", "grand total", "total\t")):
            continue
        actual = safe_float(row.get(actual_col, 0))
        budget = safe_float(row.get(budget_col, 0)) if budget_col else 0
        rows[lbl] = {"actual": actual, "budget": budget}

    return rows, sheet


def parse_june_budget(path):
    """Read the June Budget sheet: returns {label: budget_value}."""
    xl = pd.ExcelFile(path)
    sheet = next((s for s in xl.sheet_names if "budget" in s.lower()
                  and "actual" not in s.lower() and "summary" not in s.lower()), None)
    if not sheet:
        return {}
    df = clean_cols(pd.read_excel(path, sheet_name=sheet, header=0))
    df = df.dropna(how="all")
    label_col  = df.columns[0]
    budget_col = df.columns[1] if df.shape[1] > 1 else None
    out = {}
    for _, row in df.iterrows():
        lbl = " ".join(str(row.get(label_col, "")).split()).strip()
        if not lbl or lbl.lower() in ("nan", "none", "account code"):
            continue
        out[lbl] = safe_float(row.get(budget_col, 0)) if budget_col else 0
    print(f"    June budget items: {len(out)}")
    return out


# ---------------------------------------------------------------------------
# Line-item table builder
# ---------------------------------------------------------------------------

def parse_line_items(all_weekly):
    """
    Read all weekly files.  Returns:
      col_headers  : list of str  e.g. ['Jan','Feb','Mar','Apr','May','Jun 7','Jun 14','Jun 21']
      col_meta     : list of (sort_key, label)
      jun_budget   : {raw_label: budget_float}
      rows         : list of {label, display, category, is_inflow,
                               actuals[n], budgets[n]}
    """
    col_data  = []   # list of (sort_key, col_label, {raw_label: {actual, budget}})
    jun_budget = {}
    most_recent_june = None

    for path in all_weekly:
        col_label, sort_key = _col_info(path)
        rows, sheet = parse_raw_sheet(path)
        col_data.append((sort_key, col_label, rows))
        print(f"    {col_label:10s} <- {path.name} [{sheet}] ({len(rows)} rows)")
        if sort_key[0] == 6:
            most_recent_june = path

    # Sort columns chronologically
    col_data.sort(key=lambda x: x[0])
    col_headers = [x[1] for x in col_data]

    # June full-month budget from most recent June file
    if most_recent_june:
        jun_budget = parse_june_budget(most_recent_june)

    # Collect all unique labels across all files
    all_labels = {}
    for _, _, rows in col_data:
        for lbl in rows:
            if lbl not in all_labels:
                all_labels[lbl] = None

    # Build row objects
    result_rows = []
    for lbl in all_labels:
        inflow = item_is_inflow(lbl)
        actuals = []
        budgets = []
        for _, _, rows in col_data:
            entry = rows.get(lbl, {"actual": 0, "budget": 0})
            a = entry["actual"]
            b = entry["budget"]
            # Flip sign for inflows (negative raw = positive received)
            actuals.append(-a if inflow else a)
            budgets.append(-b if inflow else b)

        # Jun budget from the dedicated budget sheet (prefer over raw budget col)
        jb_raw = jun_budget.get(lbl, None)
        jb = (-jb_raw if (inflow and jb_raw is not None) else jb_raw) if jb_raw is not None else budgets[-1]

        result_rows.append({
            "label":    lbl,
            "display":  display_name(lbl),
            "category": categorize(lbl),
            "is_inflow": inflow,
            "actuals":  actuals,
            "budgets":  budgets,
            "jun_budget": abs(jb) if jb is not None else 0,
        })

    # Sort rows: by category order, then by display name
    cat_order = {c: i for i, (c, _) in enumerate(CATEGORIES)}
    cat_order["Other"] = 99
    result_rows.sort(key=lambda r: (cat_order.get(r["category"], 99), r["display"]))

    return {
        "col_headers": col_headers,
        "jun_budget":  jun_budget,
        "rows":        result_rows,
        "last_period": col_headers[-1] if col_headers else "—",
    }


# ---------------------------------------------------------------------------
# Annual plan parser
# ---------------------------------------------------------------------------

def parse_annual_plan(path):
    xl  = pd.ExcelFile(path)
    out = {"vt": {}, "monthly": {}, "file": path.name}

    vt_sheet = next((s for s in xl.sheet_names if "variance" in s.lower()), None)
    if vt_sheet:
        for hr in range(5):
            df = clean_cols(pd.read_excel(path, sheet_name=vt_sheet, header=hr))
            plan_col   = find_col(df, "2026 Төлөвлөгөө", "Plan", "Төлөвлөгөө")
            actual_col = find_col(df, "2026 Гүйцэтгэл", "Performance", "Actual", "Гүйцэтгэл")
            if plan_col and actual_col:
                en_col = (find_col(df, "Indicator (English)", "Indicator")
                          or (df.columns[2] if len(df.columns) > 2 else df.columns[1]))
                for _, row in df.iterrows():
                    lbl = " ".join(str(row.get(en_col, "")).split()).strip()
                    if not lbl or lbl.lower() in ("nan", "none"):
                        continue
                    p = safe_float(row.get(plan_col,   0)) * PLAN_UNITS
                    a = safe_float(row.get(actual_col, 0)) * PLAN_UNITS
                    out["vt"][lbl] = {"plan": p, "actual": a, "variance": a - p}
                keys_safe = [k.encode("ascii", "replace").decode()
                             for k in list(out["vt"].keys())[:5]]
                print(f"    Variance tracker: {len(out['vt'])} rows -> {keys_safe}")
                break

    iscf_sheet = next(
        (s for s in xl.sheet_names if "IS" in s and "CF" in s),
        next((s for s in xl.sheet_names if "2026" in s), None)
    )
    if iscf_sheet:
        for hr in range(5):
            df = clean_cols(pd.read_excel(path, sheet_name=iscf_sheet, header=hr))
            month_cols = []
            for mn in MN_MONTHS:
                col = find_col(df, mn)
                month_cols.append(col)
            if not any(c is not None for c in month_cols):
                continue
            en_col = (find_col(df, "Indicator (English)", "Indicator")
                      or (df.columns[2] if len(df.columns) > 2 else df.columns[1]))

            def monthly_row(*kw):
                for _, row in df.iterrows():
                    if label_has(str(row.get(en_col, "")), *kw):
                        return [safe_float(row.get(mc, 0)) * PLAN_UNITS if mc else 0
                                for mc in month_cols]
                return [0] * 12

            out["monthly"] = {
                "labels":      EN_MONTHS,
                "revenue":     monthly_row("TOTAL REVENUE", "НИЙТ БОРЛУУЛАЛТ", "Total Revenue"),
                "gross_profit":monthly_row("GROSS PROFIT", "Gross Profit"),
                "opex":        monthly_row("TOTAL OPERATING EXPENSES", "TOTAL OPEX", "Total OpEx"),
                "net_profit":  monthly_row("NET PROFIT", "ЦЭВЭР АШИГ", "Net Profit"),
            }
            break

    return out


def vt(vt_dict, *kw):
    for lbl, v in vt_dict.items():
        if label_has(lbl, *kw):
            return v
    return {"plan": 0, "actual": 0, "variance": 0}


# ---------------------------------------------------------------------------
# Bank parser
# ---------------------------------------------------------------------------

def parse_bank(path):
    xl = pd.ExcelFile(path)
    detail = next(
        (s for s in xl.sheet_names if any(k in s.lower() for k in ["group", "mn", "detail"])),
        xl.sheet_names[-1]
    )
    for hr in range(5):
        df = clean_cols(pd.read_excel(path, sheet_name=detail, header=hr))
        entity_col = find_col(df, "Entity", "Аж ахуйн нэгж")
        bank_col   = find_col(df, "BANK", "Bank", "Банк")
        ccy_col    = find_col(df, "Currency", "Валют")
        bal_col    = find_col(df, "Balance in original currency", "Balance")
        mnt_col    = find_col(df, "Converted balance in MNT", "MNT")
        usd_col    = find_col(df, "Converted balance in USD", "USD")
        if not (entity_col and bank_col):
            continue
        accounts = []
        for _, row in df.iterrows():
            entity = str(row.get(entity_col, "")).strip().upper()
            bank   = str(row.get(bank_col,   "")).strip()
            if entity != "NGL":
                continue
            if not any(nb in bank.lower() for nb in NGL_BANKS):
                continue
            ccy      = str(row.get(ccy_col, "MNT")).strip() if ccy_col else "MNT"
            bal_orig = safe_float(row.get(bal_col, 0)) if bal_col else 0
            bal_mnt  = safe_float(row.get(mnt_col, 0)) if mnt_col else 0
            bal_usd  = safe_float(row.get(usd_col, 0)) if usd_col else 0
            accounts.append({"bank": bank, "ccy": ccy,
                             "bal_orig": bal_orig, "bal_mnt": bal_mnt, "bal_usd": bal_usd})
        if accounts:
            total_mnt = sum(a["bal_mnt"] for a in accounts)
            total_usd = sum(a["bal_usd"] for a in accounts)
            print(f"    NGL accounts: {len(accounts)}  MNT={total_mnt:,.0f}  USD={total_usd:,.2f}")
            return {"accounts": accounts, "total_mnt": total_mnt, "total_usd": total_usd,
                    "file": path.name}
        break
    return {"accounts": [], "total_mnt": 0, "total_usd": 0, "file": path.name}


# ---------------------------------------------------------------------------
# KPI helpers — computed from the line-item table
# ---------------------------------------------------------------------------

def period_kpis(tbl, col_idx):
    """Compute total in / out / net for a given column index."""
    total_in = total_out = 0
    for r in tbl["rows"]:
        val = r["actuals"][col_idx] if col_idx < len(r["actuals"]) else 0
        if r["is_inflow"]:
            total_in  += val
        else:
            total_out += val
    return total_in, total_out, total_in - total_out


# ---------------------------------------------------------------------------
# Alert builder
# ---------------------------------------------------------------------------

def build_alerts(tbl):
    """Return list of alert dicts sorted by impact (abs MNT variance)."""
    alerts = []
    last = len(tbl["col_headers"]) - 1  # most recent column
    period_label = tbl["col_headers"][last]   # e.g. "Jun 1–21"
    jun_bgt_idx = last   # for June files, last column is most recent MTD

    for r in tbl["rows"]:
        actual = r["actuals"][last] if r["actuals"] else 0
        budget = r["jun_budget"]

        if budget == 0 and actual == 0:
            continue

        # Unbudgeted: no budget but actual spend/receipt
        if budget == 0 and actual != 0:
            alerts.append({
                "sev": "warn",
                "impact": abs(actual),
                "label": r["display"],
                "cat":   r["category"],
                "text":  f"{fmt_mnt(actual)} MNT in {period_label} — no budget set (unbudgeted item)",
                "tag":   "Unbudgeted",
            })
            continue

        if actual == 0 and budget > 0:
            # Revenue miss: inflow expected but nothing received
            if r["is_inflow"]:
                alerts.append({
                    "sev": "crit",
                    "impact": budget,
                    "label": r["display"],
                    "cat":   r["category"],
                    "text":  (f"0 received in {period_label} vs {fmt_mnt(budget)} MNT full-month budget — "
                              f"100% shortfall"),
                    "tag":   "Revenue miss",
                })
            continue

        if budget == 0:
            continue

        pct = (actual - budget) / budget * 100

        # Over budget (bad for costs, bad for income shortfalls)
        if r["is_inflow"]:
            if pct < -20:   # received less than 80% of expected inflow
                alerts.append({
                    "sev": "crit" if pct < -50 else "warn",
                    "impact": abs(actual - budget),
                    "label": r["display"],
                    "cat":   r["category"],
                    "text":  (f"{fmt_mnt(actual)} received in {period_label} vs {fmt_mnt(budget)} full-month budget "
                              f"({100 + pct:.0f}% of target)"),
                    "tag":   "Shortfall",
                })
        else:
            if pct > 15:    # spending more than 115% of budget
                alerts.append({
                    "sev": "crit" if pct > 50 else "warn",
                    "impact": abs(actual - budget),
                    "label": r["display"],
                    "cat":   r["category"],
                    "text":  (f"{fmt_mnt(actual)} spent in {period_label} vs {fmt_mnt(budget)} full-month budget "
                              f"(+{pct:.0f}% over)"),
                    "tag":   "Over budget",
                })

    alerts.sort(key=lambda a: (-a["impact"], a["sev"]))
    return alerts


# ---------------------------------------------------------------------------
# HTML table builder
# ---------------------------------------------------------------------------

def cell_cls(actual, budget, is_inflow):
    """Return cell CSS class based on variance."""
    if budget == 0 and actual == 0:
        return "cz"         # zero — grey
    if budget == 0 and actual != 0:
        return "cy"         # unbudgeted — amber
    pct = (actual - budget) / abs(budget) * 100
    if is_inflow:
        # More received = good
        return "cg" if pct >= -10 else ("cr" if pct < -25 else "co")
    else:
        # Less spent = good
        return "cg" if pct <= 10 else ("cr" if pct > 25 else "co")


# Groups: (display title, list of category names that belong here)
TABLE_GROUPS = [
    ("Revenue & inflows",
     ["Inflows"],
     "Money received — IEM/IEPL funding, export proceeds, other income"),
    ("Production costs",
     ["Stock", "Labour"],
     "Direct cost of collecting and processing tallow and UCO"),
    ("Overhead & operating expenses",
     ["Opex", "Other"],
     "Ongoing business support costs — utilities, fuel, rent, maintenance, communications"),
    ("Financing & obligations",
     ["Loans", "Fixed assets", "Taxes"],
     "Capital expenditure, debt service, and tax payments"),
]


def _make_table(group_rows, headers, title, subtitle, show_jun_budget=True):
    """Build one sub-table for a group of rows."""
    if not group_rows:
        return ""

    th_cells = "".join(f'<th class="num">{h}</th>' for h in headers)
    if show_jun_budget:
        th_cells += '<th class="num bgt-col">Jun 2026 budget</th>'

    html  = f'<div class="tbl-group-header"><span class="tbl-group-title">{title}</span>'
    html += f'<span class="tbl-group-sub">{subtitle}</span></div>'
    html += '<div class="ft-wrap">'
    html += '<table class="ft">'
    html += (
        "<thead><tr>"
        '<th class="col-item2">Line item</th>'
        + th_cells +
        "</tr></thead><tbody>"
    )

    for r in group_rows:
        actuals   = r["actuals"]
        budgets   = r["budgets"]
        jun_b     = r["jun_budget"]
        is_inflow = r["is_inflow"]

        td_vals = ""
        for i, (a, b) in enumerate(zip(actuals, budgets)):
            cls = cell_cls(a, b, is_inflow)
            txt = fmt_mnt(a) if a != 0 else "—"
            td_vals += f'<td class="num {cls}">{txt}</td>'

        jb_cell = ""
        if show_jun_budget:
            jb_txt = fmt_mnt(jun_b) if jun_b != 0 else "—"
            jb_cell = f'<td class="num bgt-col">{jb_txt}</td>'

        html += (
            f"<tr>"
            f'<td class="col-item2">{r["display"]}'
            f'<span class="item-cat">{r["category"]}</span></td>'
            f"{td_vals}{jb_cell}</tr>"
        )

    # Subtotal row
    n = len(headers)
    sums    = [sum(r["actuals"][i] for r in group_rows) for i in range(n)]
    jun_sum = sum(r["jun_budget"] for r in group_rows)
    td_sums = "".join(
        f'<td class="num tot">{fmt_mnt(s) if s != 0 else "—"}</td>'
        for s in sums
    )
    jb_sum = (f'<td class="num tot bgt-col">{fmt_mnt(jun_sum) if jun_sum else "—"}</td>'
              if show_jun_budget else "")
    html += (
        f"<tr class='total-row'>"
        f'<td class="col-item2 tot">{title} — subtotal</td>'
        f"{td_sums}{jb_sum}</tr>"
    )

    html += "</tbody></table></div>"
    return html


def build_split_tables_html(tbl, show_jun_budget=True):
    headers   = tbl["col_headers"]
    all_rows  = tbl["rows"]
    n         = len(headers)

    # Index rows by category for quick lookup
    by_cat = {}
    for r in all_rows:
        by_cat.setdefault(r["category"], []).append(r)

    html = ""
    group_sums = []   # track per-group totals for the net summary

    for title, cats, subtitle in TABLE_GROUPS:
        group_rows = []
        for cat in cats:
            group_rows.extend(by_cat.get(cat, []))
        if not group_rows:
            continue
        html += _make_table(group_rows, headers, title, subtitle, show_jun_budget)
        html += "<br>"
        grp_sums = [sum(r["actuals"][i] for r in group_rows) for i in range(n)]
        is_inflow_grp = any(r["is_inflow"] for r in group_rows)
        group_sums.append((title, grp_sums, is_inflow_grp))

    # Net cash flow summary bar
    in_row  = [0.0] * n
    out_row = [0.0] * n
    for r in all_rows:
        for i, a in enumerate(r["actuals"]):
            if r["is_inflow"]:
                in_row[i] += a
            else:
                out_row[i] += a
    net_row = [in_row[i] - out_row[i] for i in range(n)]

    in_bgt  = sum(r["jun_budget"] for r in all_rows if r["is_inflow"])
    out_bgt = sum(r["jun_budget"] for r in all_rows if not r["is_inflow"])
    net_bgt = in_bgt - out_bgt

    def net_cell(v):
        cls = "v-good" if v >= 0 else "v-bad"
        return f'<td class="num net-sum {cls}">{fmt_mnt(v) if v != 0 else "—"}</td>'

    th = "".join(f'<th class="num">{h}</th>' for h in headers)
    bgt_th = f'<th class="num bgt-col">Jun 2026 budget</th>' if show_jun_budget else ""
    in_cells  = "".join(f'<td class="num">{fmt_mnt(v) if v else "—"}</td>' for v in in_row)
    out_cells = "".join(f'<td class="num">{fmt_mnt(v) if v else "—"}</td>' for v in out_row)
    net_cells = "".join(net_cell(v) for v in net_row)
    in_bgt_c  = f'<td class="num bgt-col">{fmt_mnt(in_bgt)}</td>' if show_jun_budget else ""
    out_bgt_c = f'<td class="num bgt-col">{fmt_mnt(out_bgt)}</td>' if show_jun_budget else ""
    net_bgt_c = f'<td class="num bgt-col">{fmt_mnt(net_bgt)}</td>' if show_jun_budget else ""

    html += (
        '<div class="tbl-group-header">'
        '<span class="tbl-group-title">Net cash flow summary</span>'
        '<span class="tbl-group-sub">All periods — inflows minus total outflows</span>'
        '</div>'
        '<div class="ft-wrap">'
        '<table class="ft">'
        f'<thead><tr><th class="col-item2">Category</th>{th}{bgt_th}</tr></thead>'
        '<tbody>'
        f'<tr><td class="col-item2">Total inflows</td>{in_cells}{in_bgt_c}</tr>'
        f'<tr><td class="col-item2">Total outflows</td>{out_cells}{out_bgt_c}</tr>'
        f'<tr class="total-row"><td class="col-item2 tot">Net cash flow</td>{net_cells}{net_bgt_c}</tr>'
        '</tbody></table></div>'
    )

    return html


# ---------------------------------------------------------------------------
# HTML template
# ---------------------------------------------------------------------------

HTML = """\
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NGL Finance Dashboard</title>
{chartjs_tag}
<style>
:root {{
  --bg:      #0f1117;
  --surface: #1a1d2e;
  --surf2:   #252840;
  --border:  #2e3150;
  --text:    #e8eaf0;
  --muted:   #8b90a8;
  --good:    #4caf82;
  --warn:    #f5a623;
  --bad:     #e05c5c;
  --plan:    #7c83e0;
  --bank:    #4fc3f7;
  --font:    'Segoe UI', Arial, sans-serif;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ background: var(--bg); color: var(--text); font-family: var(--font); font-size: 13px; line-height: 1.5; }}

/* ── Layout ── */
header {{
  padding: 18px 32px 14px;
  border-bottom: 1px solid var(--border);
  display: flex; align-items: center; gap: 12px;
}}
header h1 {{ font-size: 17px; font-weight: 600; }}
.badge {{
  background: var(--surf2); border: 1px solid var(--border);
  border-radius: 6px; padding: 3px 10px; font-size: 11px; color: var(--muted);
}}
.badge-right {{ margin-left: auto; }}
.main {{ padding: 22px 32px; max-width: 1600px; margin: 0 auto; }}

.section-header {{
  font-size: 10px; font-weight: 700; text-transform: uppercase;
  letter-spacing: 1.2px; color: var(--muted);
  margin: 28px 0 12px; padding-bottom: 7px; border-bottom: 1px solid var(--border);
  display: flex; align-items: center; gap: 8px;
}}
.dot {{ width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; }}
.d-cash {{ background: var(--warn); }}
.d-plan {{ background: var(--plan); }}
.d-bank {{ background: var(--bank); }}
.d-alert {{ background: var(--bad); }}

/* ── Cockpit cards ── */
.cockpit {{
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(190px, 1fr));
  gap: 12px; margin-bottom: 20px;
}}
.card {{
  background: var(--surface); border: 1px solid var(--border);
  border-radius: 10px; padding: 16px 18px;
}}
.card .lbl {{ font-size: 9px; color: var(--muted); text-transform: uppercase;
  letter-spacing: .8px; margin-bottom: 7px; }}
.card .val {{ font-size: 24px; font-weight: 700; line-height: 1.1; }}
.card .sub {{ font-size: 11px; color: var(--muted); margin-top: 5px; }}
.card .tag {{ font-size: 10px; margin-top: 6px; font-weight: 600; }}
.tag-ok   {{ color: var(--good); }}
.tag-warn {{ color: var(--warn); }}
.tag-crit {{ color: var(--bad); }}
.card-crit {{ border-color: #e05c5c66; }}
.card-warn {{ border-color: #f5a62366; }}
.card-ok   {{ border-color: #4caf8266; }}
.v-good {{ color: var(--good); }}
.v-warn {{ color: var(--warn); }}
.v-bad  {{ color: var(--bad); }}

/* ── Alerts ── */
.alerts {{ margin-bottom: 24px; }}
.alert-row {{
  display: flex; align-items: baseline; gap: 10px;
  padding: 7px 14px; border-radius: 6px; margin-bottom: 5px;
  border-left: 3px solid;
}}
.alert-crit {{ background: #e05c5c0d; border-color: var(--bad); }}
.alert-warn {{ background: #f5a6230d; border-color: var(--warn); }}
.alert-dot  {{ font-size: 15px; flex-shrink: 0; }}
.alert-tag  {{ font-size: 10px; font-weight: 700; text-transform: uppercase;
  letter-spacing: .5px; min-width: 90px; flex-shrink: 0; }}
.alert-item {{ font-weight: 600; min-width: 160px; flex-shrink: 0; }}
.alert-text {{ color: var(--muted); font-size: 12px; }}

/* ── P&L cards ── */
.pl-cards {{
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(170px, 1fr));
  gap: 10px; margin-bottom: 24px;
}}
.pl-card {{
  background: var(--surface); border: 1px solid var(--border);
  border-radius: 8px; padding: 14px 16px;
}}
.pl-card .lbl {{ font-size: 9px; color: var(--muted); text-transform: uppercase;
  letter-spacing: .8px; margin-bottom: 6px; }}
.pl-card .val {{ font-size: 18px; font-weight: 700; color: var(--plan); }}
.pl-card .sub {{ font-size: 11px; color: var(--muted); margin-top: 4px; }}
.pl-card .var {{ font-size: 11px; margin-top: 4px; }}
.var-good {{ color: var(--good); }}
.var-bad  {{ color: var(--bad); }}

/* ── Bank cards ── */
.bank-cards {{
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
  gap: 10px; margin-bottom: 24px;
}}
.bank-card {{
  background: var(--surface); border: 1px solid var(--border);
  border-radius: 8px; padding: 14px 16px;
}}
.bank-card .lbl {{ font-size: 9px; color: var(--muted); text-transform: uppercase;
  letter-spacing: .8px; margin-bottom: 6px; }}
.bank-card .val {{ font-size: 18px; font-weight: 700; color: var(--bank); }}
.bank-card .sub {{ font-size: 11px; color: var(--muted); margin-top: 4px; }}
.bank-card-total {{ border-color: #4fc3f766; }}

/* ── Table group headers ── */
.tbl-group-header {{
  display: flex; align-items: baseline; gap: 12px;
  margin: 20px 0 8px;
}}
.tbl-group-title {{
  font-size: 12px; font-weight: 700; color: var(--text);
}}
.tbl-group-sub {{
  font-size: 11px; color: var(--muted);
}}

/* ── Line-item tables ── */
.ft-wrap {{
  overflow-x: auto; border-radius: 8px;
  border: 1px solid var(--border); margin-bottom: 6px;
}}
.ft {{
  border-collapse: collapse;
  width: 100%;
  min-width: 800px;
  font-size: 12px;
}}
.ft thead tr {{ background: var(--surf2); }}
.ft th, .ft td {{
  padding: 5px 10px;
  border-bottom: 1px solid var(--border);
  border-right: 1px solid var(--border);
  white-space: nowrap;
}}
.ft th {{ font-weight: 600; font-size: 10px; text-transform: uppercase;
  letter-spacing: .5px; color: var(--muted); }}
.ft .num {{ text-align: right; }}
.col-item2 {{
  position: sticky; left: 0; z-index: 2;
  background: var(--surface); min-width: 200px;
  border-right: 2px solid var(--border) !important;
}}
.ft thead .col-item2 {{ background: var(--surf2); z-index: 3; }}
.ft .total-row {{ background: var(--surf2); }}
.ft .total-row .col-item2 {{ font-weight: 700; }}
.item-cat {{
  display: block; font-size: 10px; color: var(--muted);
  font-weight: 400; margin-top: 1px;
}}
.net-sum {{ font-weight: 700; font-size: 13px; }}
.bgt-col {{ border-left: 2px solid var(--border) !important; color: var(--muted); }}

/* Cell value classes */
.cg {{ color: var(--good); }}
.co {{ color: var(--warn); }}
.cr {{ color: var(--bad);  }}
.cy {{ color: var(--warn); font-style: italic; }}
.cz {{ color: #3a3f5c; }}
.tot {{ font-weight: 700; }}

/* ── Chart area ── */
.charts {{
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
  gap: 16px; margin-bottom: 28px;
}}
.chart-card {{
  background: var(--surface); border: 1px solid var(--border);
  border-radius: 10px; padding: 18px;
}}
.chart-card h3 {{
  font-size: 10px; font-weight: 600; text-transform: uppercase;
  letter-spacing: .7px; color: var(--muted); margin-bottom: 14px;
}}

footer {{
  padding: 12px 32px; border-top: 1px solid var(--border);
  font-size: 10px; color: var(--muted);
}}
</style>
</head>
<body>

<header>
  <h1>NGL finance dashboard</h1>
  <span class="badge">Period: {period_label}</span>
  <span class="badge badge-right">Generated {generated_at}</span>
</header>

<div class="main">

  <!-- ── Executive cockpit ── -->
  <div class="section-header"><span class="dot d-cash"></span>Cash position — {period_label}</div>
  <div class="cockpit">

    <div class="card {card_in_cls}">
      <div class="lbl">Cash inflows</div>
      <div class="val">{total_in_fmt}</div>
      <div class="sub">Budget: {total_in_bgt_fmt} MNT</div>
      <div class="tag {tag_in_cls}">{tag_in_txt}</div>
    </div>

    <div class="card {card_out_cls}">
      <div class="lbl">Cash outflows</div>
      <div class="val">{total_out_fmt}</div>
      <div class="sub">Budget: {total_out_bgt_fmt} MNT</div>
      <div class="tag {tag_out_cls}">{tag_out_txt}</div>
    </div>

    <div class="card {card_net_cls}">
      <div class="lbl">Net cash flow</div>
      <div class="val {net_color}">{net_fmt}</div>
      <div class="sub">Budget: {net_bgt_fmt} MNT</div>
    </div>

    <div class="card">
      <div class="lbl">Bank balance (NGL)</div>
      <div class="val" style="color:var(--bank)">{bank_mnt_fmt}</div>
      <div class="sub">{bank_usd_fmt} &nbsp;|&nbsp; TDB + Khan Bank</div>
    </div>

    <div class="card {gap_cls}">
      <div class="lbl">June cash gap</div>
      <div class="val {gap_color}">{gap_fmt}</div>
      <div class="sub">Bank {bank_mnt_fmt} vs {remaining_out_fmt} remaining June obligations</div>
      <div class="tag {gap_tag_cls}">{gap_tag}</div>
    </div>

  </div>

  <!-- ── Critical alerts ── -->
  <div class="section-header"><span class="dot d-alert"></span>Alerts — {n_alerts} items &nbsp;<span style="font-weight:400;font-size:9px;color:var(--muted)">comparing {period_label} MTD actuals vs full June 2026 budget</span></div>
  <div class="alerts">
    {alerts_html}
  </div>

  <!-- ── Line-item tracker ── -->
  <div class="section-header"><span class="dot d-cash"></span>Cash flow tracker — line by line</div>
  <div style="font-size:11px;color:var(--muted);margin-bottom:10px;">
    Jan 2026–May 2026 columns: full-month actuals. &nbsp;
    Jun 1–7 / Jun 1–14 / Jun 1–21: month-to-date cumulative actuals. &nbsp;
    Jun 2026 budget: approved full-month budget. &nbsp;
    Colours: <span class="cg">green = on/under budget</span> &nbsp;
    <span class="co">amber = moderately over</span> &nbsp;
    <span class="cr">red = significantly over or revenue missed</span> &nbsp;
    <span class="cy">italic amber = unbudgeted</span>
  </div>
  {table_html}

  <!-- ── P&L vs annual plan ── -->
  <div class="section-header"><span class="dot d-plan"></span>P&amp;L — YTD vs 2026 annual plan</div>
  <div class="pl-cards">
    {pl_cards_html}
  </div>

  <!-- ── Bank balances ── -->
  <div class="section-header"><span class="dot d-bank"></span>NGL bank balances — TDB &amp; Khan Bank</div>
  <div class="bank-cards">
    {bank_cards_html}
  </div>

  <!-- ── P&L trend charts ── -->
  <div class="section-header"><span class="dot d-plan"></span>2026 annual plan — monthly P&amp;L breakdown</div>
  <div class="charts">
    <div class="chart-card"><h3>Revenue &amp; gross profit plan (MNT)</h3><canvas id="c1"></canvas></div>
    <div class="chart-card"><h3>OPEX plan (MNT)</h3><canvas id="c2"></canvas></div>
    <div class="chart-card"><h3>Net profit plan (MNT)</h3><canvas id="c3"></canvas></div>
  </div>

</div>

<footer>
  {weekly_file} &nbsp;|&nbsp; {plan_file} &nbsp;|&nbsp; {bank_file} &nbsp;|&nbsp; Generated {generated_at}
</footer>

<script>
const OPTS = {{
  responsive: true,
  plugins: {{ legend: {{ display: false }} }},
  scales: {{
    x: {{ ticks: {{ color: '#8b90a8', font: {{ size: 10 }} }}, grid: {{ color: '#2e3150' }} }},
    y: {{ ticks: {{ color: '#8b90a8', font: {{ size: 10 }} }}, grid: {{ color: '#2e3150' }} }}
  }}
}};
const GOPTS = {{
  responsive: true,
  plugins: {{ legend: {{ display: true, labels: {{ color: '#8b90a8', font: {{ size: 11 }} }} }} }},
  scales: {{
    x: {{ ticks: {{ color: '#8b90a8', font: {{ size: 10 }} }}, grid: {{ color: '#2e3150' }} }},
    y: {{ ticks: {{ color: '#8b90a8', font: {{ size: 10 }} }}, grid: {{ color: '#2e3150' }} }}
  }}
}};
const P = {plan_json};
const LBL = P.labels.slice(0,6);
function bar(id, data, color, opts) {{
  const nz = LBL.reduce((a,l,i) => {{ if(data[i]!==0) a.push(i); return a; }}, []);
  const fl = nz.length ? nz.map(i=>LBL[i]) : LBL;
  const fd = nz.length ? nz.map(i=>data[i]) : data;
  new Chart(document.getElementById(id), {{
    type:'bar',
    data:{{ labels:fl, datasets:[{{data:fd, backgroundColor:color+'55', borderColor:color, borderWidth:1.5, borderRadius:3}}] }},
    options: opts||OPTS
  }});
}}
function grouped(id, d1,l1,c1, d2,l2,c2) {{
  const nz = LBL.reduce((a,l,i)=>{{if(d1[i]||d2[i]) a.push(i); return a;}},[]);
  const fl = nz.length ? nz.map(i=>LBL[i]) : LBL;
  new Chart(document.getElementById(id), {{
    type:'bar',
    data:{{ labels:fl, datasets:[
      {{label:l1, data:nz.map(i=>d1[i]), backgroundColor:c1+'88', borderColor:c1, borderWidth:1.5, borderRadius:3}},
      {{label:l2, data:nz.map(i=>d2[i]), backgroundColor:c2+'88', borderColor:c2, borderWidth:1.5, borderRadius:3}}
    ]}},
    options: GOPTS
  }});
}}
grouped('c1', P.revenue,'Revenue','#7c83e0', P.gross_profit,'Gross profit','#4caf82');
bar('c2', P.opex,       '#f5a623');
bar('c3', P.net_profit, '#4caf82');
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Renderer
# ---------------------------------------------------------------------------

def _pl_card(label, kpis, is_cost=False):
    a, p = kpis["actual"], kpis["plan"]
    pct  = (a - p) / abs(p) * 100 if p else 0
    good = pct < 0 if is_cost else pct > 0
    var_cls = "var-good" if good else "var-bad"
    sign = "+" if pct > 0 else ""
    var_txt = f"{sign}{pct:.1f}% vs plan" if p else "—"
    return (
        f'<div class="pl-card">'
        f'<div class="lbl">{label}</div>'
        f'<div class="val">{fmt_mnt(a)}</div>'
        f'<div class="sub">Plan: {fmt_mnt(p)} MNT</div>'
        f'<div class="var {var_cls}">{var_txt}</div>'
        f'</div>'
    )


def _bank_cards_html(b):
    if not b["accounts"]:
        return '<div class="bank-card"><div class="lbl">No NGL accounts found</div></div>'
    html = ""
    for acc in b["accounts"]:
        orig = fmt_mnt(acc["bal_orig"]) if acc["ccy"] == "MNT" else fmt_usd(acc["bal_orig"])
        html += (
            f'<div class="bank-card">'
            f'<div class="lbl">{acc["bank"]}</div>'
            f'<div class="val">{orig} {acc["ccy"]}</div>'
            f'<div class="sub">{fmt_mnt(acc["bal_mnt"])} MNT &nbsp;|&nbsp; {fmt_usd(acc["bal_usd"])}</div>'
            f'</div>'
        )
    html += (
        f'<div class="bank-card bank-card-total">'
        f'<div class="lbl">Total NGL cash</div>'
        f'<div class="val">{fmt_mnt(b["total_mnt"])}</div>'
        f'<div class="sub">{fmt_usd(b["total_usd"])} combined</div>'
        f'</div>'
    )
    return html


def _alerts_html(alerts):
    if not alerts:
        return '<div class="alert-row alert-warn"><span class="alert-dot">—</span><span class="alert-text">No alerts generated</span></div>'
    html = ""
    for a in alerts[:12]:   # cap at 12
        cls   = "alert-crit" if a["sev"] == "crit" else "alert-warn"
        dot   = "&#9888;" if a["sev"] == "crit" else "&#x25CB;"
        html += (
            f'<div class="alert-row {cls}">'
            f'<span class="alert-dot">{dot}</span>'
            f'<span class="alert-tag">{a["tag"]}</span>'
            f'<span class="alert-item">{a["label"]}</span>'
            f'<span class="alert-text">{a["text"]}</span>'
            f'</div>'
        )
    return html


def render_html(tbl, plan, bank, all_weekly):
    generated_at = datetime.now().strftime("%d %b %Y %H:%M")

    if CHART_JS_LOCAL.exists():
        chartjs_tag = f"<script>{CHART_JS_LOCAL.read_text(encoding='utf-8')}</script>"
    else:
        chartjs_tag = f'<script src="{CHART_JS_CDN}"></script>'

    # Cockpit KPIs from most recent column
    last = len(tbl["col_headers"]) - 1
    total_in, total_out, net = period_kpis(tbl, last)

    # Jun budget totals for cockpit reference
    in_bgt  = sum(r["jun_budget"] for r in tbl["rows"] if r["is_inflow"])
    out_bgt = sum(r["jun_budget"] for r in tbl["rows"] if not r["is_inflow"])
    net_bgt = in_bgt - out_bgt

    total_bank   = bank["total_mnt"]
    period_label = tbl["last_period"]

    # Cockpit card classes
    in_pct  = total_in / in_bgt * 100 if in_bgt else 0
    out_pct = total_out / out_bgt * 100 if out_bgt else 0

    card_in_cls  = "card-crit" if in_pct < 20 else ("card-warn" if in_pct < 60 else "card-ok")
    card_out_cls = "card-ok" if out_pct < 60 else ("card-warn" if out_pct < 90 else "card-crit")
    card_net_cls = "card-warn" if net < 0 else "card-ok"
    net_color    = "v-bad" if net < 0 else "v-good"

    tag_in_txt  = f"{in_pct:.0f}% of budget received" if in_bgt else "No budget reference"
    tag_in_cls  = "tag-crit" if in_pct < 20 else ("tag-warn" if in_pct < 60 else "tag-ok")
    tag_out_txt = f"{out_pct:.0f}% of budget used" if out_bgt else "No budget reference"
    tag_out_cls = "tag-ok" if out_pct < 60 else ("tag-warn" if out_pct < 85 else "tag-crit")

    # June cash gap: bank balance vs remaining June budget obligations
    # Remaining = full-month budget minus what has already been spent MTD
    remaining_out = max(0.0, out_bgt - total_out)
    cash_gap      = total_bank - remaining_out   # negative = shortfall
    gap_cls       = "card-crit" if cash_gap < 0 else "card-ok"
    gap_color     = "v-bad" if cash_gap < 0 else "v-good"
    gap_tag       = ("Shortfall — bank insufficient for remaining obligations"
                     if cash_gap < 0
                     else "Bank covers remaining June obligations")
    gap_tag_cls   = "tag-crit" if cash_gap < 0 else "tag-ok"

    # Alerts
    alerts = build_alerts(tbl)

    # P&L cards
    vt_dict = plan["vt"]
    pl_items = [
        ("Revenue YTD",          vt(vt_dict, "Revenue", "Total Revenue"),     False),
        ("Gross profit YTD",     vt(vt_dict, "Gross Profit"),                  False),
        ("Total OPEX YTD",       vt(vt_dict, "Total OpEx", "OpEx"),            True),
        ("Operating profit YTD", vt(vt_dict, "Operating Profit"),              False),
        ("Net profit YTD",       vt(vt_dict, "Net Profit"),                    False),
        ("EBITDA YTD",           vt(vt_dict, "EBITDA"),                        False),
    ]
    pl_cards_html = "".join(_pl_card(lbl, kpis, is_cost) for lbl, kpis, is_cost in pl_items)

    # Plan monthly for charts
    plan_monthly = plan.get("monthly") or {
        "labels": EN_MONTHS, "revenue": [0]*12, "gross_profit": [0]*12,
        "opex": [0]*12, "net_profit": [0]*12
    }

    return HTML.format(
        chartjs_tag      = chartjs_tag,
        period_label     = period_label,
        generated_at     = generated_at,
        # Cockpit
        total_in_fmt     = fmt_mnt(total_in),
        total_in_bgt_fmt = fmt_mnt(in_bgt),
        total_out_fmt    = fmt_mnt(total_out),
        total_out_bgt_fmt= fmt_mnt(out_bgt),
        net_fmt          = fmt_mnt(net),
        net_bgt_fmt      = fmt_mnt(net_bgt),
        net_color        = net_color,
        bank_mnt_fmt     = fmt_mnt(bank["total_mnt"]),
        bank_usd_fmt     = fmt_usd(bank["total_usd"]),
        gap_fmt          = fmt_mnt(cash_gap),
        gap_cls          = gap_cls,
        gap_color        = gap_color,
        gap_tag          = gap_tag,
        gap_tag_cls      = gap_tag_cls,
        remaining_out_fmt= fmt_mnt(remaining_out),
        card_in_cls      = card_in_cls,
        card_out_cls     = card_out_cls,
        card_net_cls     = card_net_cls,
        tag_in_txt       = tag_in_txt,
        tag_in_cls       = tag_in_cls,
        tag_out_txt      = tag_out_txt,
        tag_out_cls      = tag_out_cls,
        # Alerts
        n_alerts         = len(alerts),
        alerts_html      = _alerts_html(alerts),
        # Table
        table_html       = build_split_tables_html(tbl),
        # P&L
        pl_cards_html    = pl_cards_html,
        # Bank
        bank_cards_html  = _bank_cards_html(bank),
        # Charts data
        plan_json        = json.dumps(plan_monthly),
        # Footer
        weekly_file      = all_weekly[-1].name,
        plan_file        = plan["file"],
        bank_file        = bank["file"],
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    print("NGL Finance Dashboard — CFO/CEO Edition")
    print("=" * 45)

    all_weekly, plan_path, bank_path = find_files()

    print("\nParsing files...")
    print("  Line-item table:")
    tbl  = parse_line_items(all_weekly)
    print(f"  Annual plan:")
    plan = parse_annual_plan(plan_path)
    print(f"  Bank:")
    bank = parse_bank(bank_path)

    print("\nRendering HTML...")
    html = render_html(tbl, plan, bank, all_weekly)

    out = OUTPUT_DIR / "dashboard.html"
    out.write_text(html, encoding="utf-8")
    print(f"\nDone. Open: {out.resolve()}")


if __name__ == "__main__":
    main()
