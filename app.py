"""CSV / Excel cleaner demo.

DataPhil demo - the sample file is made up. Upload your own file to try it; it is processed
in memory only and never stored.
Run locally:  pip install -r requirements.txt && streamlit run app.py
"""
import csv
import html
import io
import re
from datetime import datetime
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

SAMPLE = Path(__file__).parent / "data" / "sample_orders_messy.csv"

COUNTRY_MAP = {  # common spellings -> one standard name; extend per client
    "at": "Austria", "austria": "Austria", "österreich": "Austria", "oesterreich": "Austria",
    "de": "Germany", "germany": "Germany", "deutschland": "Germany",
    "ch": "Switzerland", "switzerland": "Switzerland", "schweiz": "Switzerland", "suisse": "Switzerland",
}
TOTAL_WORDS = ("total", "sum", "summe", "gesamt")


# ------------------------------------------------------------------ reading
def read_any(name: str, raw: bytes) -> pd.DataFrame:
    """Read CSV (any delimiter, header not necessarily in row 1) or Excel into strings."""
    if name.lower().endswith((".xlsx", ".xls")):
        grid = pd.read_excel(io.BytesIO(raw), header=None, dtype=str).fillna("").values.tolist()
    else:
        text = raw.decode("utf-8-sig", errors="replace")
        try:
            dialect = csv.Sniffer().sniff(text[:5000], delimiters=";,\t|")
            delim = dialect.delimiter
        except csv.Error:
            delim = ","
        grid = list(csv.reader(io.StringIO(text), delimiter=delim))
    width = max(len(r) for r in grid)
    grid = [r + [""] * (width - len(r)) for r in grid]
    # header = first row among the first 15 that has the most filled cells
    filled = [sum(bool(str(c).strip()) for c in r) for r in grid[:15]]
    h = filled.index(max(filled))
    df = pd.DataFrame(grid[h + 1:], columns=[str(c) for c in grid[h]])
    df.attrs["skipped_top_rows"] = h
    return df


# ------------------------------------------------------------------ parsing helpers
def parse_number(s: str):
    s = re.sub(r"[^\d,.\-]", "", str(s))
    if not s or s in "-.,":
        return None
    if "," in s and "." in s:          # 1.234,56 or 1,234.56: the last separator is the decimal one
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if len(s.split(",")[-1]) != 3 else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(s: str, us_slash: bool = True):
    s = str(s).strip()
    if not s:
        return None, False
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date(), False
        except ValueError:
            pass
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        a, b, y = map(int, m.groups())
        month, day = (a, b) if us_slash else (b, a)
        if month > 12:                                   # order impossible -> swap
            month, day = day, month
        ambiguous = a <= 12 and b <= 12 and a != b
        try:
            return datetime(y, month, day).date(), ambiguous
        except ValueError:
            return None, False
    return None, False


def looks_like(series: pd.Series, kind: str) -> bool:
    vals = [v for v in series.astype(str).str.strip() if v]
    if not vals:
        return False
    if kind == "date":
        ok = sum(parse_date(v)[0] is not None for v in vals)
    else:
        ok = sum(parse_number(v) is not None and not re.search(r"[a-df-z]{3,}", v.lower().replace("eur", ""))
                 for v in vals)
    return ok / len(vals) > 0.8


# ------------------------------------------------------------------ cleaning
def clean(df: pd.DataFrame, opts: dict):
    df = df.copy()
    log = []

    def note(step, n, detail=""):
        log.append({"Step": step, "Changed": int(n), "Detail": detail})

    if df.attrs.get("skipped_top_rows"):
        note("Skip title rows above the header", df.attrs["skipped_top_rows"], "rows")

    if opts["trim"]:
        new_cols = [re.sub(r"\s+", " ", c).strip() for c in df.columns]
        n_cols = sum(a != b for a, b in zip(df.columns, new_cols))
        df.columns = new_cols
        before = df.copy()
        df = df.apply(lambda col: col.astype(str).str.replace(r"\s+", " ", regex=True).str.strip())
        note("Trim spaces", (before != df).sum().sum() + n_cols, f"cells (+ {n_cols} column names)")

    if opts["empty"]:
        mask = df.apply(lambda r: all(not str(v).strip() for v in r), axis=1)
        df = df[~mask]
        note("Remove empty rows", mask.sum(), "rows")

    if opts["totals"]:
        first = df.iloc[:, 0].astype(str).str.lower().str.strip()
        mask = first.isin(TOTAL_WORDS)
        df = df[~mask]
        note("Remove total rows", mask.sum(), "rows")

    if opts["dupes"]:
        mask = df.duplicated(keep="first")
        df = df[~mask]
        note("Remove exact duplicates", mask.sum(), "rows")

    date_cols = [c for c in df.columns if looks_like(df[c], "date")] if opts["dates"] else []
    num_cols = [c for c in df.columns if c not in date_cols and looks_like(df[c], "number")] if opts["numbers"] else []
    text_cols = [c for c in df.columns if c not in date_cols + num_cols]

    if opts["case"]:
        changed = 0
        for c in text_cols:
            vals = df[c][df[c] != ""]
            if vals.str.lower().nunique() < vals.nunique():          # same value in different spellings
                canon = vals.groupby(vals.str.lower()).agg(lambda s: s.str.title().mode()[0])
                new = df[c].map(lambda v: canon.get(v.lower(), v) if v else v)
                changed += (new != df[c]).sum()
                df[c] = new
        note("Unify upper/lower case", changed, "cells")

    if opts["country"]:
        changed = 0
        for c in text_cols:
            low = df[c].str.lower()
            if low.isin(COUNTRY_MAP).mean() > 0.5:
                new = low.map(COUNTRY_MAP).fillna(df[c])
                changed += (new != df[c]).sum()
                df[c] = new
        note("Standardise country names", changed, "cells")

    for c in date_cols:
        parsed = df[c].map(lambda v: parse_date(v, opts["us_slash"]))
        amb = sum(p[1] for p in parsed)
        bad = sum(p[0] is None and str(v).strip() != "" for p, v in zip(parsed, df[c]))
        changed = sum(p[0] is not None and str(v).strip() != p[0].isoformat() for p, v in zip(parsed, df[c]))
        df[c] = pd.to_datetime(parsed.map(lambda p: p[0]))
        order = "month/day" if opts["us_slash"] else "day/month"
        note(f"Dates in '{c}' to one format (YYYY-MM-DD)", changed,
             f"{amb} could be either way, read as {order}; {bad} unreadable")

    for c in num_cols:
        new = df[c].map(parse_number)
        plain = df[c].astype(str).str.fullmatch(r"-?\d+(\.\d+)?") | (df[c].astype(str).str.strip() == "")
        note(f"Numbers in '{c}' (currency signs, decimal comma)", (~plain).sum(), f"{new.isna().sum()} missing")
        if new.dropna().apply(float.is_integer).all():
            new = new.astype("Int64")
        df[c] = new

    missing = int(df.isna().sum().sum() + (df == "").sum().sum())
    note("Missing values flagged (not filled)", missing, "cells, see 'Issues' column")
    flags = df.isna() | (df == "")
    df["Issues"] = flags.apply(lambda r: ", ".join(f"missing {c}" for c, v in r.items() if v), axis=1)
    return df.reset_index(drop=True), pd.DataFrame(log), date_cols, num_cols


def to_excel(clean_df, report) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        clean_df.to_excel(xw, sheet_name="Clean data", index=False)
        report.to_excel(xw, sheet_name="Cleaning report", index=False)
    return buf.getvalue()


# ---------------------------------- look & feel (same palette as dataphil.dev) ----------------------------------
NAVY, INK, BODY, MUTED = "#1E3A6E", "#0E1A2F", "#3D4047", "#5A5E66"
GREY, LIGHT, PAPER, WARN = "#9AA3B2", "#D3D9E3", "#F4F5F8", "#B4531A"

CSS = """<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@400;500;600&family=Instrument+Serif&display=swap');
html, body, .stApp, .stMarkdown, p, label, button, input, textarea, select, li {font-family:'IBM Plex Sans',system-ui,sans-serif;}
.block-container {padding-top:2.2rem; max-width:1180px;}
h1 {font-family:'Instrument Serif',Georgia,serif !important; font-weight:400 !important; font-size:2.8rem !important;
    color:#0E1A2F !important; letter-spacing:-.01em; line-height:1.05 !important; padding:0 !important;}
.dp-lead {color:#5A5E66; font-size:1.02rem; margin:.35rem 0 1.5rem; max-width:48rem;}
.dp-lead b {color:#0E1A2F; font-weight:500;}
.dp-chip {font-family:'IBM Plex Mono',monospace; font-size:.7rem; letter-spacing:.08em; color:#1E3A6E;
    background:#E7EBF2; border-radius:20px; padding:.15rem .6rem; margin-left:.5rem; vertical-align:middle; white-space:nowrap;}
.dp-step {font-family:'IBM Plex Mono',monospace; font-size:.72rem !important; letter-spacing:.14em; color:#5A5E66; margin:0 0 .5rem;}
.dp-answer {background:#0E1A2F; color:#fff; border-radius:14px; padding:1.15rem 1.4rem 1.25rem; margin-bottom:.9rem;}
.dp-answer .k {font-family:'IBM Plex Mono',monospace; font-size:.7rem; letter-spacing:.14em; color:#9DB6E3;}
.dp-answer .big {font-family:'Instrument Serif',Georgia,serif; font-size:3.1rem; line-height:1.05; margin:.25rem 0 .55rem;}
.dp-answer .big.warn {color:#F2B48C;}
.dp-answer p {color:#D6DEEB; margin:0; font-size:1rem; line-height:1.55;}
.dp-answer p b {color:#fff; font-weight:500;}
.dp-kpi {background:#fff; border:1px solid #D3D9E3; border-radius:12px; padding:.75rem 1rem; height:100%;}
.dp-kpi .l {font-size:.82rem; color:#5A5E66;}
.dp-kpi .v {font-size:1.5rem; font-weight:500; color:#0E1A2F; line-height:1.25; margin:.1rem 0;}
.dp-kpi .c {font-size:.8rem; color:#5A5E66;} .dp-kpi .c.warn {color:#B4531A;}
.dp-h {font-weight:600; font-size:1.05rem; color:#0E1A2F; margin:0; padding-top:1.3rem;}
.dp-sub {font-size:.84rem; color:#5A5E66; margin:0 0 .4rem;}
.dp-try {font-size:.84rem; color:#3D4047; background:#E7EBF2; border-radius:10px; padding:.6rem .8rem; margin-top:.4rem;}
.dp-note {font-size:.78rem; color:#5A5E66; margin-top:.5rem;}
.dp-list {margin:.2rem 0 0; padding-left:1.1rem; color:#3D4047; font-size:.92rem; line-height:1.7;}
.dp-list b {color:#0E1A2F; font-weight:500;}
@media (max-width:700px){ h1 {font-size:2.1rem !important;} .dp-answer .big {font-size:2.4rem;} }
</style>"""


def esc(s) -> str:
    return html.escape(str(s))


def header(title: str, lead_html: str, sample: bool = True):
    st.title(title)
    chip = '<span class="dp-chip">SAMPLE DATA</span>' if sample else ""
    st.markdown(f'<p class="dp-lead">{lead_html}{chip}</p>', unsafe_allow_html=True)


def step(label: str, where=st):
    where.markdown(f'<p class="dp-step">{esc(label)}</p>', unsafe_allow_html=True)


def answer(label: str, big: str, text_html: str, warn=False, where=st):
    """The one thing the user came for: a big number or verdict, and one sentence that explains it."""
    where.markdown(f'<div class="dp-answer"><span class="k">{esc(label)}</span>'
                   f'<div class="big{" warn" if warn else ""}">{esc(big)}</div><p>{text_html}</p></div>',
                   unsafe_allow_html=True)


def kpi(col, label, value, context="", warn=False):
    col.markdown(f'<div class="dp-kpi"><div class="l">{esc(label)}</div><div class="v">{esc(value)}</div>'
                 f'<div class="c{" warn" if warn else ""}">{"⚠ " if warn else ""}{esc(context)}</div></div>',
                 unsafe_allow_html=True)


def heading(title: str, sub: str = "", where=st):
    where.markdown(f'<p class="dp-h">{esc(title)}</p>' + (f'<p class="dp-sub">{esc(sub)}</p>' if sub else ""),
                   unsafe_allow_html=True)


def try_hint(text_html: str, where=st):
    where.markdown(f'<div class="dp-try"><b>Try it:</b> {text_html}</div>', unsafe_allow_html=True)


def note(text: str, where=st):
    where.markdown(f'<div class="dp-note">{esc(text)}</div>', unsafe_allow_html=True)


def chart_style(chart):
    return (chart.properties(padding={"left": 8, "right": 8, "top": 4, "bottom": 4}).configure_view(strokeWidth=0)
            .configure_axis(labelColor=MUTED, titleColor=MUTED, labelFont="IBM Plex Sans", titleFont="IBM Plex Sans",
                            labelFontSize=12, titleFontSize=12, gridColor="#EEF1F5", domainColor=LIGHT, tickColor=LIGHT)
            .configure_text(font="IBM Plex Sans"))


# ------------------------------------------------------------------ UI
st.set_page_config(page_title="CSV cleaner (demo)", page_icon="🧹", layout="wide", initial_sidebar_state="collapsed")
st.markdown(CSS, unsafe_allow_html=True)

header("Can we trust the numbers in this export?",
       "A monthly order export with title rows, duplicates, mixed spellings and amounts stored as text. "
       "<b>Put the messy file in, get a clean table and the correct totals out.</b>",
       sample=st.session_state.get("upload") is None)

left, right = st.columns([1, 2], gap="large")
with left:
    step("1 · THE FILE")
    with st.container(border=True):
        up = st.file_uploader("Your own CSV or Excel file (optional)", type=["csv", "txt", "xlsx"], key="upload")
        if not up:
            st.caption(f"Using the sample file {SAMPLE.name}")
        us = st.radio("Dates like 3/4/2026 mean", ["month/day (US)", "day/month"], horizontal=True)
        with st.expander("Cleaning steps"):
            opts = {
                "trim": st.checkbox("Trim extra spaces", True),
                "empty": st.checkbox("Remove empty rows", True),
                "totals": st.checkbox("Remove total rows", True),
                "dupes": st.checkbox("Remove exact duplicates", True),
                "case": st.checkbox("Unify upper/lower case", True),
                "country": st.checkbox("Standardise country names", True),
                "dates": st.checkbox("Convert dates to one format", True),
                "numbers": st.checkbox("Convert amounts to numbers", True),
            }
        opts["us_slash"] = us == "month/day (US)"
    try_hint("open <b>Cleaning steps</b> and untick <b>Convert amounts to numbers</b>.", left)
    note("Runs in your browser. Your file is not uploaded anywhere and not stored.", left)

name, raw = (up.name, up.getvalue()) if up else (SAMPLE.name, SAMPLE.read_bytes())
try:
    raw_df = read_any(name, raw)
except Exception as e:  # noqa: BLE001 - show any reading problem to the user
    st.error(f"Could not read this file: {e}")
    st.stop()

clean_df, report, date_cols, num_cols = clean(raw_df, opts)

# ---- impact: what would the numbers have said without cleaning?
raw_cols = list(raw_df.columns)
pos = {c: i for i, c in enumerate(c for c in clean_df.columns if c != "Issues")}
amount_col = num_cols[-1] if num_cols else None
text_cols = [c for c in clean_df.columns if c not in date_cols + num_cols + ["Issues"]]
cust_col = next((c for c in text_cols if re.search(r"customer|client|kunde|name|company", c, re.I)), None)
if cust_col is None:
    cands = [c for c in text_cols if 2 <= clean_df[c].nunique() <= 50]
    cust_col = cands[0] if cands else None
cur = "€" if re.search(rb"\xe2\x82\xac|EUR", raw) else ""
plain = re.compile(r"^\s*-?\d+(\.\d+)?\s*$")
flagged = clean_df["Issues"] != ""
n_check = int(flagged.sum())

excel_sum = clean_sum = None
if amount_col is not None and pos.get(amount_col, 99) < len(raw_cols):
    raw_amount = raw_df.iloc[:, pos[amount_col]].astype(str)
    excel_sum = sum(float(v) for v in raw_amount if plain.match(v))     # what SUM() sees: real numbers only
    clean_sum = float(clean_df[amount_col].sum())
raw_n = clean_n = None
if cust_col is not None and pos.get(cust_col, 99) < len(raw_cols):
    raw_names = raw_df.iloc[:, pos[cust_col]].astype(str)
    raw_n = raw_names[raw_names.str.strip() != ""]
    raw_n = raw_n[~raw_n.str.strip().str.lower().isin(TOTAL_WORDS)].str.lower().nunique()  # like a pivot table
    clean_n = clean_df[cust_col][clean_df[cust_col] != ""].nunique()

with right:
    step("2 · THE RESULT")
    if excel_sum is not None and clean_sum:
        gap = (excel_sum - clean_sum) / clean_sum
        text = (f"Summing the raw export in Excel gives <b>{cur}{excel_sum:,.0f}</b> ({gap:+.0%}): amounts stored "
                f"as text are skipped and duplicate rows count twice." if abs(gap) > 0.005
                else "The raw export already added up correctly; the cleaning mostly fixed names and dates.")
        answer(f"CORRECT TOTAL {amount_col.upper()}", f"{cur}{clean_sum:,.0f}", text)
    else:
        answer("CLEAN ROWS", f"{len(clean_df)}", f"From {len(raw_df)} rows in the export. "
               "No amount column was found, so no total is shown.")

    k1, k2, k3 = st.columns(3)
    kpi(k1, "Clean rows", f"{len(clean_df)}", f"from {len(raw_df)} rows in the export")
    if raw_n is not None:
        kpi(k2, f"{cust_col}s", f"{clean_n}", f"the raw file looks like {raw_n} (different spellings)"
            if raw_n != clean_n else "no duplicate spellings found", warn=raw_n != clean_n)
    else:
        kpi(k2, "Columns", f"{len(clean_df.columns) - 1}", "no name column found")
    kpi(k3, "Rows to check", f"{n_check}", "missing values are flagged, never guessed" if n_check
        else "nothing missing", warn=n_check > 0)

    st.write("")
    d1, d2 = st.columns([3, 2])
    d1.download_button("Download the clean Excel file", to_excel(clean_df, report), file_name="clean_data.xlsx",
                       type="primary", use_container_width=True)
    d2.download_button("Clean CSV", clean_df.to_csv(index=False).encode("utf-8"),
                       file_name="clean_data.csv", mime="text/csv", use_container_width=True)

    heading("What was fixed", "Every change is logged in the Excel file, sheet 'Cleaning report'.")
    got = dict(zip(report["Step"], report["Changed"]))
    def n(prefix):
        return int(sum(v for k, v in got.items() if k.startswith(prefix)))
    removed = [f"{n(k)} {w if n(k) > 1 else w1}" for k, w, w1 in (
        ("Remove exact", "duplicate rows", "duplicate row"), ("Remove empty", "empty rows", "empty row"),
        ("Remove total", "total lines", "total line"), ("Skip title", "title rows", "title row")) if n(k)]
    lines = []
    if removed:
        lines.append(f"<b>Removed</b> {', '.join(removed)}")
    if n("Unify") + n("Standardise"):
        lines.append(f"<b>Unified</b> the spelling in {n('Unify') + n('Standardise')} cells (names and countries)")
    if n("Dates") + n("Numbers"):
        lines.append(f"<b>Converted</b> {n('Dates')} dates and {n('Numbers')} amounts so Excel can calculate with them")
    if n("Trim"):
        lines.append(f"<b>Trimmed</b> {n('Trim')} stray spaces")
    items = "".join(f"<li>{t}</li>" for t in lines)
    st.markdown(f'<ul class="dp-list">{items}</ul>', unsafe_allow_html=True)

heading("The clean table", "Rows that need a look are shaded; the 'Issues' column says why.")
dcfg = {c: st.column_config.DateColumn(format="YYYY-MM-DD") for c in date_cols}
dcfg.update({c: st.column_config.NumberColumn(format="%.2f") for c in num_cols
             if not str(clean_df[c].dtype).startswith("Int")})
order = clean_df.index[flagged].tolist() + clean_df.index[~flagged].tolist()      # rows to check first
show = clean_df.loc[order]
st.dataframe(show.style.apply(lambda r: ["background-color:#FBEDE4" if flagged[r.name] else ""] * len(r), axis=1),
             hide_index=True, use_container_width=True, height=320, column_config=dcfg)
with st.expander("Compare with the raw export"):
    st.dataframe(raw_df, hide_index=True, use_container_width=True, height=320)
