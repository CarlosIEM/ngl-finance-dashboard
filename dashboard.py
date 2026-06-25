#!/usr/bin/env python3
"""NGL weekly finance dashboard generator.

Drop files in data/ and run: python dashboard.py
Output: output/dashboard.html
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

PLAN_UNITS = 1_000   # annual plan values are MNT thousands
NGL_BANKS  = {"trade development bank", "tdb", "khan bank", "khanbank",
               "хаан банк", "худалдаа хөгжлийн банк"}
MN_MONTHS  = ["1-р сар", "2-р сар", "3-р сар", "4-р сар", "5-р сар", "6-р сар",
               "7-р сар", "8-р сар", "9-р сар", "10-р сар", "11-р сар", "12-р сар"]
EN_MONTHS  = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

MONTH_NAMES = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
}

# Labels that classify as cash inflows vs outflows in the raw Actual sheets
INFLOW_KW  = ["fund", "income of export", "export income", "income various",
               "income export", "various income"]
OUTFLOW_KW = ["stock -", "tax", "loan khanbank", "loan dbm", "fixed asset",
               "wage", "social insurance", "fuel", "light", "kitchen",
               "opex", "maintenance", "rental", "communication"]


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def safe_float(value):
    try:
        f = float(value)
        return 0.0 if pd.isna(f) else f
    except (TypeError, ValueError):
        return 0.0


def fmt_mnt(value, decimals=1):
    try:
        v = float(value)
        if v == 0:
            return "0"
        sign = "-" if v < 0 else ""
        av = abs(v)
        if av >= 1_000_000_000:
            return f"{sign}{av / 1_000_000_000:.{decimals}f}B"
        if av >= 1_000_000:
            return f"{sign}{av / 1_000_000:.{decimals}f}M"
        return f"{sign}{av:,.0f}"
    except (TypeError, ValueError):
        return "—"


def fmt_usd(value, decimals=0):
    try:
        v = float(value)
        sign = "-" if v < 0 else ""
        av = abs(v)
        if av >= 1_000_000:
            return f"USD {sign}{av / 1_000_000:.2f}m"
        if av >= 1_000:
            return f"USD {sign}{av / 1_000:.{decimals}f}k"
        return f"USD {sign}{av:,.{decimals}f}"
    except (TypeError, ValueError):
        return "—"


def var_pct(actual, budget):
    if budget == 0:
        return None
    return (actual - budget) / abs(budget) * 100


def var_html(actual, budget, is_cost=False):
    pct = var_pct(actual, budget)
    if pct is None:
        return ""
    if abs(pct) < 2:
        return '<div class="variance neu">&rarr; on budget</div>'
    good = (pct < 0) if is_cost else (pct > 0)
    cls  = "good" if good else "bad"
    sym  = ("&#9660;" if is_cost else "&#9650;") if good else ("&#9650;" if is_cost else "&#9660;")
    sign = "+" if pct > 0 else ""
    return f'<div class="variance {cls}">{sym} {sign}{pct:.1f}% vs budget</div>'


def var_plan_html(actual, plan, is_cost=False):
    pct = var_pct(actual, plan)
    if pct is None:
        return ""
    if abs(pct) < 2:
        return '<div class="variance neu">&rarr; on plan</div>'
    good = (pct < 0) if is_cost else (pct > 0)
    cls  = "good" if good else "bad"
    sym  = ("&#9660;" if is_cost else "&#9650;") if good else ("&#9650;" if is_cost else "&#9660;")
    sign = "+" if pct > 0 else ""
    return f'<div class="variance {cls}">{sym} {sign}{pct:.1f}% vs plan</div>'


def label_has(s, *keywords):
    if not isinstance(s, str):
        return False
    sl = s.lower()
    return any(k.lower() in sl for k in keywords)


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


def find_month_cols(df):
    result = []
    for i, mn in enumerate(MN_MONTHS):
        col = find_col(df, mn)
        if col is None:
            for alt in [str(i + 1), f"{i + 1:02d}", f"2026.{i + 1}", f"2026.{i + 1:02d}"]:
                col = find_col(df, alt)
                if col:
                    break
        result.append(col)
    return result


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

    # Sort weekly files by mtime so [-1] = most recently modified
    weekly = sorted([f for f in all_xlsx if _is_weekly(f)], key=lambda f: f.stat().st_mtime)
    plan   = sorted([f for f in all_xlsx
                     if any(k in f.name.lower() for k in ["annual_plan", "annual plan", "plan_final"])],
                    key=lambda f: f.stat().st_mtime)
    bank   = sorted([f for f in all_xlsx
                     if any(k in f.name.lower() for k in ["bank_balance", "bank balance"])],
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
    print(f"  Plan   : {plan[-1].name}")
    print(f"  Bank   : {bank[-1].name}")
    return weekly, plan[-1], bank[-1]


# ---------------------------------------------------------------------------
# Weekly budget vs actual
# ---------------------------------------------------------------------------

def _parse_raw_actual_sheet(path):
    """Parse a raw 'Actual' sheet (ACCOUNT CODE | actual | budget | diff).

    Used for older monthly files and as fallback for newer weekly files.
    Returns dict: label -> {actual, budget, diff}
    """
    xl = pd.ExcelFile(path)
    # Prefer the sheet whose name contains 'actual'; fall back to first sheet
    actual_sheet = next(
        (s for s in xl.sheet_names if "actual" in s.lower()),
        xl.sheet_names[0]
    )
    df = clean_cols(pd.read_excel(path, sheet_name=actual_sheet, header=0))
    df = df.dropna(how="all")

    # Col 0 = label, col 1 = actual cash flow, col 2 = budget (may not exist)
    if df.shape[1] < 2:
        return {}

    label_col  = df.columns[0]
    actual_col = df.columns[1]
    budget_col = df.columns[2] if df.shape[1] > 2 else None
    diff_col   = df.columns[3] if df.shape[1] > 3 else None

    rows = {}
    for _, row in df.iterrows():
        lbl = " ".join(str(row.get(label_col, "")).split()).strip()
        if not lbl or lbl.lower() in ("nan", "none", "account code"):
            continue
        rows[lbl] = {
            "actual": safe_float(row.get(actual_col, 0)),
            "budget": safe_float(row.get(budget_col, 0)) if budget_col else 0,
            "diff":   safe_float(row.get(diff_col,   0)) if diff_col   else 0,
        }
    return rows


def _rows_to_kpis(rows):
    """Extract standard KPI dict from a label->values row map."""
    def row(*kw):
        for lbl, v in rows.items():
            if label_has(lbl, *kw):
                return v
        return {"budget": 0, "actual": 0, "diff": 0}

    # Compute total in/out from line items when not pre-aggregated
    total_in_row  = row("Total Money In", "Total Cash In", "TOTAL IN", "Нийт мөнгө орлого")
    total_out_row = row("Total Money Out", "Total Cash Out", "TOTAL OUT", "Нийт мөнгө зарлага")
    net_row       = row("Net Cash Flow", "Net Cash", "Цэвэр мөнгөн урсгал")

    if total_in_row["actual"] == 0:
        # Sum inflow rows
        total_in_actual  = sum(v["actual"] for lbl, v in rows.items() if label_has(lbl, *INFLOW_KW))
        total_in_budget  = sum(v["budget"] for lbl, v in rows.items() if label_has(lbl, *INFLOW_KW))
    else:
        total_in_actual  = total_in_row["actual"]
        total_in_budget  = total_in_row["budget"]

    if total_out_row["actual"] == 0:
        total_out_actual = sum(v["actual"] for lbl, v in rows.items() if label_has(lbl, *OUTFLOW_KW))
        total_out_budget = sum(v["budget"] for lbl, v in rows.items() if label_has(lbl, *OUTFLOW_KW))
    else:
        total_out_actual = total_out_row["actual"]
        total_out_budget = total_out_row["budget"]

    net_actual = net_row["actual"] if net_row["actual"] != 0 else (total_in_actual - total_out_actual)
    net_budget = net_row["budget"] if net_row["budget"] != 0 else (total_in_budget - total_out_budget)

    return {
        "total_in":      {"actual": total_in_actual,  "budget": total_in_budget},
        "total_out":     {"actual": total_out_actual, "budget": total_out_budget},
        "net_cash":      {"actual": net_actual,       "budget": net_budget},
        "export_income": row("Export Income", "INCOME OF EXPORT", "Экспорт орлого"),
        "funding_iem":   row("Funding from IEM", "FUND from IEM", "IEM"),
        "funding_iepl":  row("Funding from IEPL", "FUND from IEPL", "IEPL"),
        "various_income":row("Various Income", "Income various", "Бусад орлого"),
        "stock_uco":     row("STOCK - UCO", "UCO COLLECTED"),
        "stock_fat":     row("STOCK - FAT", "FAT COLLECTED"),
        "tax":           row("TAX", "Татвар"),
        "loan":          row("LOAN KHANBANK", "LOAN DBM", "LOAN", "Зээл"),
    }


def parse_weekly(path):
    xl = pd.ExcelFile(path)

    # Try the pre-aggregated summary sheet first (newer June weekly files)
    summary_sheet = next((s for s in xl.sheet_names if "summary" in s.lower()), None)

    rows = {}
    if summary_sheet:
        for hr in range(9):
            candidate = clean_cols(pd.read_excel(path, sheet_name=summary_sheet, header=hr))
            cols_l = [c.lower() for c in candidate.columns]
            if any("actual" in c for c in cols_l) and any("budget" in c for c in cols_l):
                label_col  = candidate.columns[0]
                budget_col = find_col(candidate, "Budget (₮)", "Budget", "Төсөв")
                actual_col = find_col(candidate, "Actual (₮)", "Actual", "Гүйцэтгэл")
                diff_col   = find_col(candidate, "Difference (₮)", "Difference", "Зөрүү")
                for _, row in candidate.iterrows():
                    lbl = " ".join(str(row.get(label_col, "")).split()).strip()
                    if not lbl or lbl.lower() in ("nan", "none"):
                        continue
                    rows[lbl] = {
                        "actual": safe_float(row.get(actual_col, 0) if actual_col else 0),
                        "budget": safe_float(row.get(budget_col, 0) if budget_col else 0),
                        "diff":   safe_float(row.get(diff_col,   0) if diff_col   else 0),
                    }
                print(f"    [{path.name}] summary sheet, {len(rows)} rows")
                break

    # Fall back to raw actual sheet
    if not rows:
        rows = _parse_raw_actual_sheet(path)
        print(f"    [{path.name}] raw actual sheet, {len(rows)} rows")

    kpis = _rows_to_kpis(rows)
    kpis["period_label"] = path.stem.replace("_", " ")
    kpis["file"]         = path.name
    kpis["rows"]         = rows
    return kpis


def _month_from_filename(path):
    """Return the actual-data month number (1-12) from a filename, or None."""
    stem = path.stem.lower()
    # Pattern: "[budget_month]_budget_vs_[actual_month]_Actual"
    # The actual month is the second month name in the filename
    found = []
    for word in stem.replace("_", " ").split():
        if word in MONTH_NAMES:
            found.append(MONTH_NAMES[word])
    # Second month mentioned = the actual data month
    return found[1] if len(found) >= 2 else (found[0] if found else None)


def build_monthly_actuals(all_weekly):
    """Read all weekly files and return monthly cash trend series."""
    by_month = {}  # month_num -> {total_in, total_out, net, export}

    for path in all_weekly:
        month = _month_from_filename(path)
        if month is None:
            continue
        rows = _parse_raw_actual_sheet(path)
        if not rows:
            continue

        kpis = _rows_to_kpis(rows)
        # For June we may have multiple weekly files; keep the last (most data)
        by_month[month] = {
            "total_in":  kpis["total_in"]["actual"],
            "total_out": kpis["total_out"]["actual"],
            "net":       kpis["net_cash"]["actual"],
            "export":    kpis["export_income"]["actual"],
        }

    months_present = sorted(by_month.keys())
    print(f"    Monthly actuals: months {months_present}")

    return {
        "labels":    [EN_MONTHS[m - 1] for m in range(1, 13)],
        "total_in":  [by_month.get(m, {}).get("total_in",  0) for m in range(1, 13)],
        "total_out": [by_month.get(m, {}).get("total_out", 0) for m in range(1, 13)],
        "net":       [by_month.get(m, {}).get("net",       0) for m in range(1, 13)],
        "export":    [by_month.get(m, {}).get("export",    0) for m in range(1, 13)],
    }


# ---------------------------------------------------------------------------
# Annual plan
# ---------------------------------------------------------------------------

def parse_annual_plan(path):
    xl  = pd.ExcelFile(path)
    out = {"vt": {}, "monthly": {}, "file": path.name}

    # Variance Tracker -------------------------------------------------------
    vt_sheet = next((s for s in xl.sheet_names if "variance" in s.lower()), None)
    if vt_sheet:
        print(f"    Plan sheet (VT): {vt_sheet}")
        for hr in range(5):
            df = clean_cols(pd.read_excel(path, sheet_name=vt_sheet, header=hr))
            plan_col   = find_col(df, "2026 Төлөвлөгөө", "Plan", "Төлөвлөгөө")
            actual_col = find_col(df, "2026 Гүйцэтгэл", "Performance", "Actual", "Гүйцэтгэл")
            if plan_col and actual_col:
                # English label column: look for "Indicator (English)" by name; fall back to index 2 then 1
                en_col = (find_col(df, "Indicator (English)", "Indicator")
                          or (df.columns[2] if len(df.columns) > 2 else df.columns[1]))
                for _, row in df.iterrows():
                    lbl = " ".join(str(row.get(en_col, "")).split()).strip()
                    if not lbl or lbl.lower() in ("nan", "none"):
                        continue
                    p = safe_float(row.get(plan_col,   0)) * PLAN_UNITS
                    a = safe_float(row.get(actual_col, 0)) * PLAN_UNITS
                    out["vt"][lbl] = {"plan": p, "actual": a, "variance": a - p}
                keys_safe = [k.encode("ascii", "replace").decode() for k in list(out['vt'].keys())[:5]]
                print(f"    Variance tracker: {len(out['vt'])} rows -> {keys_safe}")
                break

    # IS & CF 2026 monthly ---------------------------------------------------
    iscf_sheet = next(
        (s for s in xl.sheet_names if "IS" in s and "CF" in s),
        next((s for s in xl.sheet_names if "2026" in s), None)
    )
    if iscf_sheet:
        print(f"    Plan sheet (IS&CF): {iscf_sheet}")
        for hr in range(5):
            df = clean_cols(pd.read_excel(path, sheet_name=iscf_sheet, header=hr))
            month_cols = find_month_cols(df)
            if not any(c is not None for c in month_cols):
                continue
            en_col = (find_col(df, "Indicator (English)", "Indicator")
                      or (df.columns[2] if len(df.columns) > 2 else df.columns[1]))

            def monthly_row(*kw):
                for _, row in df.iterrows():
                    if label_has(str(row.get(en_col, "")), *kw):
                        return [
                            safe_float(row.get(mc, 0)) * PLAN_UNITS if mc else 0
                            for mc in month_cols
                        ]
                return [0] * 12

            out["monthly"] = {
                "labels":      EN_MONTHS,
                "revenue":     monthly_row("TOTAL REVENUE", "НИЙТ БОРЛУУЛАЛТ", "Total Revenue", "Total Sales"),
                "gross_profit":monthly_row("GROSS PROFIT", "Нийт ашиг", "Gross Profit"),
                "opex":        monthly_row("TOTAL OPEX", "Total OpEx", "Нийт үйл ажиллагааны зардал", "Operating Expenses"),
                "net_profit":  monthly_row("NET PROFIT", "ЦЭВЭР АШИГ", "Net Profit", "Net Income"),
            }
            nz = sum(1 for v in out["monthly"]["revenue"] if v != 0)
            print(f"    Monthly: {nz} non-zero revenue months")
            break

    return out


# ---------------------------------------------------------------------------
# Bank balance
# ---------------------------------------------------------------------------

def parse_bank(path):
    xl = pd.ExcelFile(path)
    detail = next(
        (s for s in xl.sheet_names if any(k in s.lower() for k in ["group", "mn", "detail"])),
        xl.sheet_names[-1]
    )
    print(f"    Bank sheet: {detail}")

    for hr in range(5):
        df = clean_cols(pd.read_excel(path, sheet_name=detail, header=hr))
        entity_col = find_col(df, "Entity", "Аж ахуйн нэгж", "Company")
        bank_col   = find_col(df, "BANK", "Bank", "Банк")
        ccy_col    = find_col(df, "Currency", "Валют", "CCY")
        bal_col    = find_col(df, "Balance in original currency", "Balance", "Үлдэгдэл")
        mnt_col    = find_col(df, "Converted balance in MNT", "MNT balance", "MNT")
        usd_col    = find_col(df, "Converted balance in USD", "USD balance", "USD")

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
            accounts.append({
                "bank": bank, "ccy": ccy,
                "bal_orig": bal_orig, "bal_mnt": bal_mnt, "bal_usd": bal_usd,
            })

        if accounts:
            total_mnt = sum(a["bal_mnt"] for a in accounts)
            total_usd = sum(a["bal_usd"] for a in accounts)
            print(f"    NGL accounts: {len(accounts)}  MNT={total_mnt:,.0f}  USD={total_usd:,.2f}")
            return {
                "accounts": accounts,
                "total_mnt": total_mnt,
                "total_usd": total_usd,
                "file": path.name,
            }
        break

    print("    WARNING: No NGL TDB/Khan accounts found")
    return {"accounts": [], "total_mnt": 0, "total_usd": 0, "file": path.name}


# ---------------------------------------------------------------------------
# VT lookup helper
# ---------------------------------------------------------------------------

def vt(vt_dict, *keywords):
    for lbl, v in vt_dict.items():
        if label_has(lbl, *keywords):
            return v
    return {"plan": 0, "actual": 0, "variance": 0}


# ---------------------------------------------------------------------------
# Executive summary
# ---------------------------------------------------------------------------

def build_exec_summary(w, p, b):
    STATUS = {"ok": "#4caf82", "warning": "#f5a623", "bad": "#e05c5c", "neu": "#8b90a8"}

    net     = w.get("net_cash", {})
    net_a   = net.get("actual", 0)
    net_b   = net.get("budget", 0)
    net_pct = var_pct(net_a, net_b)

    rev_vt  = vt(p["vt"], "Revenue", "Борлуулалт")
    np_vt   = vt(p["vt"], "Net Profit", "Цэвэр ашиг")
    gp_vt   = vt(p["vt"], "Gross Profit", "Нийт ашиг")

    items = []

    # Cash this period
    if net_b != 0:
        good = net_pct is not None and net_pct > 0
        st   = "ok" if good else "bad"
        sign = "+" if (net_pct or 0) > 0 else ""
        pct_str = f" ({sign}{net_pct:.1f}% vs budget)" if net_pct is not None else ""
        items.append({
            "status": st,
            "label":  "Cash position",
            "text":   f"Net cash flow {fmt_mnt(net_a)} MNT{pct_str}. "
                      f"Total in {fmt_mnt(w.get('total_in', {}).get('actual', 0))} MNT, "
                      f"total out {fmt_mnt(w.get('total_out', {}).get('actual', 0))} MNT."
        })

    # Revenue YTD
    if rev_vt["plan"] != 0:
        rev_pct = var_pct(rev_vt["actual"], rev_vt["plan"])
        rev_st  = "ok" if (rev_pct or 0) >= 0 else "bad"
        sign    = "+" if (rev_pct or 0) > 0 else ""
        items.append({
            "status": rev_st,
            "label":  "Revenue YTD",
            "text":   f"{fmt_mnt(rev_vt['actual'])} MNT actual vs {fmt_mnt(rev_vt['plan'])} MNT plan "
                      f"({sign}{rev_pct:.1f}%)." if rev_pct is not None else
                      f"{fmt_mnt(rev_vt['actual'])} MNT actual vs {fmt_mnt(rev_vt['plan'])} MNT plan."
        })

    # Net profit YTD
    if np_vt["plan"] != 0:
        np_pct = var_pct(np_vt["actual"], np_vt["plan"])
        np_st  = "ok" if (np_pct or 0) >= 0 else "bad"
        sign   = "+" if (np_pct or 0) > 0 else ""
        items.append({
            "status": np_st,
            "label":  "Net profit YTD",
            "text":   f"{fmt_mnt(np_vt['actual'])} MNT actual vs {fmt_mnt(np_vt['plan'])} MNT plan "
                      f"({sign}{np_pct:.1f}%)." if np_pct is not None else
                      f"{fmt_mnt(np_vt['actual'])} MNT actual vs {fmt_mnt(np_vt['plan'])} MNT plan."
        })

    # Bank balance
    if b["total_mnt"] != 0:
        items.append({
            "status": "neu",
            "label":  "Bank balance",
            "text":   f"NGL cash on hand {fmt_mnt(b['total_mnt'])} MNT "
                      f"({fmt_usd(b['total_usd'])}) across TDB and Khan Bank."
        })

    rows_html = ""
    for item in items:
        color = STATUS.get(item["status"], "#8b90a8")
        rows_html += (
            f'<div class="exec-item">'
            f'<span class="exec-dot" style="background:{color}"></span>'
            f'<span class="exec-label">{item["label"]}</span>'
            f'<span class="exec-text">{item["text"]}</span>'
            f'</div>'
        )

    return (
        f'<div class="exec-summary">'
        f'<div class="exec-title">Executive summary &mdash; {w.get("period_label", "")}</div>'
        f'{rows_html}'
        f'</div>'
    )


# ---------------------------------------------------------------------------
# Bank cards builder
# ---------------------------------------------------------------------------

def build_bank_cards(b):
    if not b["accounts"]:
        return '<div class="card"><div class="label">Bank balance</div><div class="value">—</div></div>'

    html = ""
    for acc in b["accounts"]:
        bank_short = acc["bank"].split()[-1] if acc["bank"] else "—"
        orig = fmt_mnt(acc["bal_orig"]) if acc["ccy"] == "MNT" else fmt_usd(acc["bal_orig"])
        html += (
            f'<div class="card">'
            f'<div class="label">{acc["bank"]}</div>'
            f'<div class="value cb">{orig} {acc["ccy"]}</div>'
            f'<div class="sub">{fmt_mnt(acc["bal_mnt"])} MNT &nbsp;|&nbsp; {fmt_usd(acc["bal_usd"])}</div>'
            f'</div>'
        )

    html += (
        f'<div class="card card-total">'
        f'<div class="label">Total NGL cash</div>'
        f'<div class="value cb">{fmt_mnt(b["total_mnt"])} MNT</div>'
        f'<div class="sub">{fmt_usd(b["total_usd"])} combined</div>'
        f'</div>'
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
  --cash:    #f5a623;
  --plan:    #7c83e0;
  --bank:    #4fc3f7;
  --good:    #4caf82;
  --bad:     #e05c5c;
  --font:    'Segoe UI', Arial, sans-serif;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ background: var(--bg); color: var(--text); font-family: var(--font); font-size: 14px; }}

header {{
  padding: 22px 32px 16px;
  border-bottom: 1px solid var(--border);
  display: flex; align-items: center; gap: 14px;
}}
header h1 {{ font-size: 18px; font-weight: 600; }}
.badge {{
  background: var(--surf2); border: 1px solid var(--border);
  border-radius: 6px; padding: 3px 10px; font-size: 12px; color: var(--muted);
}}
.badge-right {{ margin-left: auto; }}

.main {{ padding: 24px 32px; max-width: 1400px; margin: 0 auto; }}

.section-header {{
  font-size: 11px; font-weight: 700; text-transform: uppercase;
  letter-spacing: 1.2px; color: var(--muted);
  margin: 28px 0 12px; padding-bottom: 8px; border-bottom: 1px solid var(--border);
  display: flex; align-items: center; gap: 8px;
}}
.section-header:first-child {{ margin-top: 0; }}
.sub-header {{
  font-size: 10px; font-weight: 600; text-transform: uppercase;
  letter-spacing: 1px; color: var(--muted); margin: 16px 0 8px;
}}
.dot {{ width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }}
.d-cash {{ background: var(--cash); }}
.d-plan {{ background: var(--plan); }}
.d-bank {{ background: var(--bank); }}

.cards {{
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(175px, 1fr));
  gap: 12px; margin-bottom: 16px;
}}
.card {{
  background: var(--surface); border: 1px solid var(--border);
  border-radius: 10px; padding: 16px 18px;
}}
.card .label {{ font-size: 10px; color: var(--muted); text-transform: uppercase; letter-spacing: .7px; margin-bottom: 8px; }}
.card .value {{ font-size: 22px; font-weight: 700; line-height: 1.1; }}
.card .sub   {{ font-size: 11px; color: var(--muted); margin-top: 5px; }}
.card .variance {{ font-size: 11px; margin-top: 5px; }}
.variance.good {{ color: var(--good); }}
.variance.bad  {{ color: var(--bad); }}
.variance.neu  {{ color: var(--muted); }}
.cc {{ color: var(--cash); }}
.cp {{ color: var(--plan); }}
.cb {{ color: var(--bank); }}
.cg {{ color: var(--good); }}
.cr {{ color: var(--bad); }}
.card-total {{ border-color: #3a3f5c; background: #1e2235; }}
.card-hi    {{ border-color: var(--cash); }}

.exec-summary {{
  background: var(--surface); border: 1px solid var(--border);
  border-radius: 10px; padding: 20px 24px; margin-bottom: 28px;
}}
.exec-title {{
  font-size: 11px; font-weight: 700; text-transform: uppercase;
  letter-spacing: 1px; color: var(--muted); margin-bottom: 14px;
}}
.exec-item {{
  display: flex; align-items: baseline; gap: 10px;
  margin-bottom: 9px; font-size: 13px; line-height: 1.6;
}}
.exec-item:last-child {{ margin-bottom: 0; }}
.exec-dot  {{ width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; margin-top: 4px; }}
.exec-label {{ font-weight: 600; min-width: 120px; flex-shrink: 0; }}
.exec-text {{ color: var(--muted); }}

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
  font-size: 11px; font-weight: 600; text-transform: uppercase;
  letter-spacing: .7px; color: var(--muted); margin-bottom: 14px;
}}

footer {{
  padding: 14px 32px; border-top: 1px solid var(--border);
  font-size: 11px; color: var(--muted);
}}
</style>
</head>
<body>

<header>
  <h1>NGL finance dashboard</h1>
  <span class="badge">{period_label}</span>
  <span class="badge badge-right">Generated {generated_at}</span>
</header>

<div class="main">

  {exec_summary}

  <!-- ── Cash flow this period ── -->
  <div class="section-header"><span class="dot d-cash"></span>Cash flow &mdash; {period_label}</div>

  <div class="cards">
    <div class="card card-hi">
      <div class="label">Total cash in</div>
      <div class="value cc">{w_total_in_a}</div>
      <div class="sub">Budget: {w_total_in_b} MNT</div>
      {w_total_in_var}
    </div>
    <div class="card card-hi">
      <div class="label">Total cash out</div>
      <div class="value">{w_total_out_a}</div>
      <div class="sub">Budget: {w_total_out_b} MNT</div>
      {w_total_out_var}
    </div>
    <div class="card card-hi">
      <div class="label">Net cash flow</div>
      <div class="value {net_color}">{w_net_a}</div>
      <div class="sub">Budget: {w_net_b} MNT</div>
      {w_net_var}
    </div>
  </div>

  <div class="sub-header">Key inflows</div>
  <div class="cards">
    <div class="card">
      <div class="label">Export income</div>
      <div class="value cc">{w_export_a}</div>
      <div class="sub">Budget: {w_export_b} MNT</div>
      {w_export_var}
    </div>
    <div class="card">
      <div class="label">Funding from IEM</div>
      <div class="value cc">{w_iem_a}</div>
      <div class="sub">Budget: {w_iem_b} MNT</div>
    </div>
    <div class="card">
      <div class="label">Funding from IEPL</div>
      <div class="value cc">{w_iepl_a}</div>
      <div class="sub">Budget: {w_iepl_b} MNT</div>
    </div>
    <div class="card">
      <div class="label">Various income</div>
      <div class="value cc">{w_various_a}</div>
      <div class="sub">Budget: {w_various_b} MNT</div>
    </div>
  </div>

  <div class="sub-header">Key outflows</div>
  <div class="cards">
    <div class="card">
      <div class="label">UCO stock purchases</div>
      <div class="value">{w_uco_a}</div>
      <div class="sub">Budget: {w_uco_b} MNT</div>
      {w_uco_var}
    </div>
    <div class="card">
      <div class="label">Fat stock purchases</div>
      <div class="value">{w_fat_a}</div>
      <div class="sub">Budget: {w_fat_b} MNT</div>
      {w_fat_var}
    </div>
    <div class="card">
      <div class="label">Tax</div>
      <div class="value">{w_tax_a}</div>
      <div class="sub">Budget: {w_tax_b} MNT</div>
      {w_tax_var}
    </div>
    <div class="card">
      <div class="label">Loan repayments</div>
      <div class="value">{w_loan_a}</div>
      <div class="sub">Budget: {w_loan_b} MNT</div>
      {w_loan_var}
    </div>
  </div>

  <!-- ── P&L vs annual plan ── -->
  <div class="section-header"><span class="dot d-plan"></span>P&amp;L &mdash; YTD vs 2026 annual plan</div>
  <div class="cards">
    <div class="card">
      <div class="label">Revenue YTD</div>
      <div class="value cp">{p_rev_a}</div>
      <div class="sub">Plan: {p_rev_p} MNT</div>
      {p_rev_var}
    </div>
    <div class="card">
      <div class="label">Gross profit YTD</div>
      <div class="value cp">{p_gp_a}</div>
      <div class="sub">Plan: {p_gp_p} MNT</div>
      {p_gp_var}
    </div>
    <div class="card">
      <div class="label">Total OPEX YTD</div>
      <div class="value">{p_opex_a}</div>
      <div class="sub">Plan: {p_opex_p} MNT</div>
      {p_opex_var}
    </div>
    <div class="card">
      <div class="label">Operating profit YTD</div>
      <div class="value cp">{p_ebit_a}</div>
      <div class="sub">Plan: {p_ebit_p} MNT</div>
      {p_ebit_var}
    </div>
    <div class="card">
      <div class="label">Net profit YTD</div>
      <div class="value {np_color}">{p_np_a}</div>
      <div class="sub">Plan: {p_np_p} MNT</div>
      {p_np_var}
    </div>
    <div class="card">
      <div class="label">EBITDA YTD</div>
      <div class="value cp">{p_ebitda_a}</div>
      <div class="sub">Plan: {p_ebitda_p} MNT</div>
      {p_ebitda_var}
    </div>
  </div>

  <!-- ── Monthly actuals trend ── -->
  <div class="section-header"><span class="dot d-cash"></span>Monthly cash actuals &mdash; Jan to {last_month}</div>
  <div class="charts">
    <div class="chart-card"><h3>Cash in &amp; cash out by month (MNT)</h3><canvas id="c1"></canvas></div>
    <div class="chart-card"><h3>Net cash flow by month (MNT)</h3><canvas id="c2"></canvas></div>
    <div class="chart-card"><h3>Export income by month (MNT)</h3><canvas id="c3"></canvas></div>
  </div>

  <!-- ── P&L plan breakdown ── -->
  <div class="section-header"><span class="dot d-plan"></span>2026 annual plan &mdash; monthly P&amp;L breakdown</div>
  <div class="charts">
    <div class="chart-card"><h3>Revenue &amp; gross profit plan (MNT)</h3><canvas id="c4"></canvas></div>
    <div class="chart-card"><h3>OPEX plan (MNT)</h3><canvas id="c5"></canvas></div>
    <div class="chart-card"><h3>Net profit plan (MNT)</h3><canvas id="c6"></canvas></div>
  </div>

  <!-- ── Bank balances ── -->
  <div class="section-header"><span class="dot d-bank"></span>NGL bank balances &mdash; TDB &amp; Khan Bank</div>
  <div class="cards">
    {bank_cards}
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
const GROUPED_OPTS = {{
  responsive: true,
  plugins: {{ legend: {{ display: true, labels: {{ color: '#8b90a8', font: {{ size: 11 }} }} }} }},
  scales: {{
    x: {{ ticks: {{ color: '#8b90a8', font: {{ size: 10 }} }}, grid: {{ color: '#2e3150' }} }},
    y: {{ ticks: {{ color: '#8b90a8', font: {{ size: 10 }} }}, grid: {{ color: '#2e3150' }} }}
  }}
}};

const A = {actuals_json};
const P = {plan_json};
const ALABELS = A.labels;
const PLABELS = P.labels.slice(0, 6);

function bar(id, labels, color, data, opts) {{
  const nonzero = labels.reduce((acc, l, i) => {{ if (data[i] !== 0) acc.push(i); return acc; }}, []);
  const fl = nonzero.length ? nonzero.map(i => labels[i]) : labels;
  const fd = nonzero.length ? nonzero.map(i => data[i])   : data;
  new Chart(document.getElementById(id).getContext('2d'), {{
    type: 'bar',
    data: {{
      labels: fl,
      datasets: [{{ data: fd, backgroundColor: color + '55', borderColor: color, borderWidth: 1.5, borderRadius: 3 }}]
    }},
    options: opts || OPTS
  }});
}}

function grouped(id, labels, d1, l1, c1, d2, l2, c2, opts) {{
  const nonzero = labels.reduce((acc, l, i) => {{ if (d1[i] !== 0 || d2[i] !== 0) acc.push(i); return acc; }}, []);
  const fl = nonzero.length ? nonzero.map(i => labels[i]) : labels;
  new Chart(document.getElementById(id).getContext('2d'), {{
    type: 'bar',
    data: {{
      labels: fl,
      datasets: [
        {{ label: l1, data: nonzero.map(i => d1[i]), backgroundColor: c1 + '88', borderColor: c1, borderWidth: 1.5, borderRadius: 3 }},
        {{ label: l2, data: nonzero.map(i => d2[i]), backgroundColor: c2 + '88', borderColor: c2, borderWidth: 1.5, borderRadius: 3 }}
      ]
    }},
    options: opts || GROUPED_OPTS
  }});
}}

// Cash actuals
grouped('c1', ALABELS, A.total_in, 'Cash in', '#4caf82', A.total_out, 'Cash out', '#e05c5c');
bar('c2', ALABELS, '#7c83e0', A.net);
bar('c3', ALABELS, '#f5a623', A.export);

// P&L plan
grouped('c4', PLABELS, P.revenue, 'Revenue', '#7c83e0', P.gross_profit, 'Gross profit', '#4caf82');
bar('c5', PLABELS, '#f5a623', P.opex);
bar('c6', PLABELS, '#4caf82', P.net_profit);
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Renderer
# ---------------------------------------------------------------------------

def w_fmt(item):
    return {"a": fmt_mnt(item["actual"]), "b": fmt_mnt(item["budget"])}


def render_html(w, p, b, actuals):
    generated_at = datetime.now().strftime("%d %b %Y %H:%M")

    if CHART_JS_LOCAL.exists():
        chartjs_tag = f"<script>{CHART_JS_LOCAL.read_text(encoding='utf-8')}</script>"
    else:
        chartjs_tag = f'<script src="{CHART_JS_CDN}"></script>'

    vt_dict = p["vt"]
    rev   = vt(vt_dict, "Revenue",         "Борлуулалт")
    gp    = vt(vt_dict, "Gross Profit",    "Нийт ашиг")
    opex  = vt(vt_dict, "Total OpEx",      "Үйл ажиллагааны", "OpEx")
    ebit  = vt(vt_dict, "Operating Profit","Үйл ажиллагааны ашиг", "EBIT")
    np_   = vt(vt_dict, "Net Profit",      "Цэвэр ашиг")
    ebitda= vt(vt_dict, "EBITDA")

    net_a = w.get("net_cash", {}).get("actual", 0)
    np_a  = np_.get("actual", 0)

    plan_monthly = p.get("monthly") or {
        "labels": EN_MONTHS, "revenue": [0]*12, "gross_profit": [0]*12,
        "opex": [0]*12, "net_profit": [0]*12
    }

    # Last month with actual data for section header label
    act_labels = actuals.get("labels", EN_MONTHS)
    act_totals = actuals.get("total_in", [0]*12)
    last_month = next(
        (act_labels[i] for i in range(11, -1, -1) if act_totals[i] != 0),
        "—"
    )

    return HTML.format(
        chartjs_tag   = chartjs_tag,
        period_label  = w.get("period_label", ""),
        generated_at  = generated_at,
        last_month    = last_month,
        exec_summary  = build_exec_summary(w, p, b),
        bank_cards    = build_bank_cards(b),
        actuals_json  = json.dumps(actuals),
        plan_json     = json.dumps(plan_monthly),
        weekly_file   = w.get("file", ""),
        plan_file     = p.get("file", ""),
        bank_file     = b.get("file", ""),
        # Cash — top 3
        w_total_in_a  = fmt_mnt(w.get("total_in",  {}).get("actual", 0)),
        w_total_in_b  = fmt_mnt(w.get("total_in",  {}).get("budget", 0)),
        w_total_out_a = fmt_mnt(w.get("total_out", {}).get("actual", 0)),
        w_total_out_b = fmt_mnt(w.get("total_out", {}).get("budget", 0)),
        w_net_a       = fmt_mnt(net_a),
        w_net_b       = fmt_mnt(w.get("net_cash",  {}).get("budget", 0)),
        net_color     = "cg" if net_a >= 0 else "cr",
        w_total_in_var = var_html(w.get("total_in",  {}).get("actual", 0), w.get("total_in",  {}).get("budget", 0)),
        w_total_out_var= var_html(w.get("total_out", {}).get("actual", 0), w.get("total_out", {}).get("budget", 0), is_cost=True),
        w_net_var      = var_html(net_a, w.get("net_cash", {}).get("budget", 0)),
        # Cash — inflows
        w_export_a   = fmt_mnt(w.get("export_income",  {}).get("actual", 0)),
        w_export_b   = fmt_mnt(w.get("export_income",  {}).get("budget", 0)),
        w_export_var = var_html(w.get("export_income",  {}).get("actual", 0), w.get("export_income",  {}).get("budget", 0)),
        w_iem_a      = fmt_mnt(w.get("funding_iem",    {}).get("actual", 0)),
        w_iem_b      = fmt_mnt(w.get("funding_iem",    {}).get("budget", 0)),
        w_iepl_a     = fmt_mnt(w.get("funding_iepl",   {}).get("actual", 0)),
        w_iepl_b     = fmt_mnt(w.get("funding_iepl",   {}).get("budget", 0)),
        w_various_a  = fmt_mnt(w.get("various_income", {}).get("actual", 0)),
        w_various_b  = fmt_mnt(w.get("various_income", {}).get("budget", 0)),
        # Cash — outflows
        w_uco_a   = fmt_mnt(w.get("stock_uco", {}).get("actual", 0)),
        w_uco_b   = fmt_mnt(w.get("stock_uco", {}).get("budget", 0)),
        w_uco_var = var_html(w.get("stock_uco", {}).get("actual", 0), w.get("stock_uco", {}).get("budget", 0), is_cost=True),
        w_fat_a   = fmt_mnt(w.get("stock_fat", {}).get("actual", 0)),
        w_fat_b   = fmt_mnt(w.get("stock_fat", {}).get("budget", 0)),
        w_fat_var = var_html(w.get("stock_fat", {}).get("actual", 0), w.get("stock_fat", {}).get("budget", 0), is_cost=True),
        w_tax_a   = fmt_mnt(w.get("tax",  {}).get("actual", 0)),
        w_tax_b   = fmt_mnt(w.get("tax",  {}).get("budget", 0)),
        w_tax_var = var_html(w.get("tax",  {}).get("actual", 0), w.get("tax",  {}).get("budget", 0), is_cost=True),
        w_loan_a  = fmt_mnt(w.get("loan", {}).get("actual", 0)),
        w_loan_b  = fmt_mnt(w.get("loan", {}).get("budget", 0)),
        w_loan_var= var_html(w.get("loan", {}).get("actual", 0), w.get("loan", {}).get("budget", 0), is_cost=True),
        # P&L YTD
        p_rev_a   = fmt_mnt(rev["actual"]),   p_rev_p   = fmt_mnt(rev["plan"]),
        p_rev_var = var_plan_html(rev["actual"],   rev["plan"]),
        p_gp_a    = fmt_mnt(gp["actual"]),    p_gp_p    = fmt_mnt(gp["plan"]),
        p_gp_var  = var_plan_html(gp["actual"],    gp["plan"]),
        p_opex_a  = fmt_mnt(opex["actual"]),  p_opex_p  = fmt_mnt(opex["plan"]),
        p_opex_var= var_plan_html(opex["actual"],  opex["plan"], is_cost=True),
        p_ebit_a  = fmt_mnt(ebit["actual"]),  p_ebit_p  = fmt_mnt(ebit["plan"]),
        p_ebit_var= var_plan_html(ebit["actual"],  ebit["plan"]),
        p_np_a    = fmt_mnt(np_["actual"]),   p_np_p    = fmt_mnt(np_["plan"]),
        p_np_var  = var_plan_html(np_["actual"],   np_["plan"]),
        np_color  = "cg" if np_a >= 0 else "cr",
        p_ebitda_a= fmt_mnt(ebitda["actual"]),p_ebitda_p= fmt_mnt(ebitda["plan"]),
        p_ebitda_var=var_plan_html(ebitda["actual"],ebitda["plan"]),
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    print("NGL Finance Dashboard Generator")
    print("=" * 45)

    all_weekly, plan_path, bank_path = find_files()
    most_recent = all_weekly[-1]

    print("\nParsing files...")
    w = parse_weekly(most_recent)
    p = parse_annual_plan(plan_path)
    b = parse_bank(bank_path)

    print("\nBuilding monthly actuals from all weekly files...")
    actuals = build_monthly_actuals(all_weekly)

    print("\nRendering HTML...")
    html = render_html(w, p, b, actuals)

    out = OUTPUT_DIR / "dashboard.html"
    out.write_text(html, encoding="utf-8")
    print(f"\nDone. Open: {out.resolve()}")


if __name__ == "__main__":
    main()
