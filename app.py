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
.block-container {padding-top:2rem; max-width:1200px;}
h1 {font-family:'Instrument Serif',Georgia,serif !important; font-weight:400 !important; font-size:2.9rem !important;
    color:#0E1A2F !important; letter-spacing:-.01em; line-height:1.05 !important; padding:0 !important;}
.dp-bar {font-family:'IBM Plex Mono',monospace; font-size:.78rem; color:#C3CCDB; background:#0E1A2F;
    padding:.55rem 1rem; border-radius:10px; display:flex; justify-content:space-between; margin-bottom:1.1rem;}
.dp-bar b {font-weight:400; color:#fff; background:#17253F; border-radius:20px; padding:0 .6rem;}
.dp-lead {color:#5A5E66; font-size:1rem; margin:.2rem 0 1.2rem;}
.dp-kpi {background:#fff; border:1px solid #D3D9E3; border-radius:12px; padding:.85rem 1rem; height:100%;}
.dp-kpi .l {font-size:.82rem; color:#5A5E66;}
.dp-kpi .v {font-size:1.75rem; font-weight:500; color:#0E1A2F; line-height:1.25; margin:.15rem 0;}
.dp-kpi .c {font-size:.8rem; color:#5A5E66;} .dp-kpi .c.warn {color:#B4531A;}
.dp-msg {background:#0E1A2F; color:#fff; border-radius:12px; padding:1rem 1.25rem; margin:1rem 0 1.4rem; font-size:1.02rem; line-height:1.5;}
.dp-msg .k {display:block; font-family:'IBM Plex Mono',monospace; font-size:.7rem; letter-spacing:.14em; color:#9DB6E3; margin-bottom:.2rem;}
.dp-h {font-weight:600; font-size:1.08rem; color:#0E1A2F; margin:.6rem 0 0;}
.dp-sub {font-size:.84rem; color:#5A5E66; margin:0 0 .4rem;}
.dp-steps {display:grid; grid-template-columns:repeat(3,1fr); gap:12px; margin-top:1.4rem;}
.dp-steps div {border:1px solid #D3D9E3; border-radius:12px; padding:.9rem 1rem; font-size:.88rem; color:#3D4047; background:#fff;}
.dp-steps div.r {background:#1E3A6E; color:#fff; border-color:#1E3A6E;}
.dp-steps span {display:block; font-family:'IBM Plex Mono',monospace; font-size:.68rem; letter-spacing:.14em; color:#5A5E66; margin-bottom:.3rem;}
.dp-steps div.r span {color:#9DB6E3;}
.dp-src {font-family:'IBM Plex Mono',monospace; font-size:.78rem; color:#3D4047; line-height:1.9;}
.dp-src i {color:#2E7D4F; font-style:normal; margin-right:.4rem;}
.dp-note {background:#E7EBF2; border-radius:10px; padding:.7rem .9rem; font-size:.8rem; color:#3D4047; margin-top:1rem;}
@media (max-width:700px){ .dp-steps {grid-template-columns:1fr;} h1 {font-size:2.2rem !important;} }
</style>"""


def esc(s) -> str:
    return html.escape(str(s))


def bar(left: str, right: str = "Sample data"):
    st.markdown(f'<div class="dp-bar"><span>{esc(left)}</span><b>{esc(right)}</b></div>', unsafe_allow_html=True)


def kpi(col, label, value, context="", warn=False):
    col.markdown(f'<div class="dp-kpi"><div class="l">{esc(label)}</div><div class="v">{esc(value)}</div>'
                 f'<div class="c{" warn" if warn else ""}">{"⚠ " if warn else ""}{esc(context)}</div></div>',
                 unsafe_allow_html=True)


def message(text_html: str, label="MESSAGE FOR THE MEETING"):
    st.markdown(f'<div class="dp-msg"><span class="k">{label}</span>{text_html}</div>', unsafe_allow_html=True)


def heading(title: str, sub: str = ""):
    st.markdown(f'<p class="dp-h">{esc(title)}</p>' + (f'<p class="dp-sub">{esc(sub)}</p>' if sub else ""),
                unsafe_allow_html=True)


def sources(files):
    st.markdown("<p class='dp-sub' style='letter-spacing:.12em;font-family:IBM Plex Mono'>DATA SOURCES</p>"
                + "".join(f"<div class='dp-src'><i>✓</i>{esc(f)}</div>" for f in files), unsafe_allow_html=True)


def steps(before: str, tool: str, result: str):
    st.markdown(f'<div class="dp-steps"><div><span>1 · BEFORE</span>{before}</div><div><span>2 · THE TOOL</span>{tool}'
                f'</div><div class="r"><span>3 · RESULT</span>{result}</div></div>', unsafe_allow_html=True)


def chart_style(chart):
    return (chart.properties(padding={"left": 8, "right": 8, "top": 4, "bottom": 4}).configure_view(strokeWidth=0)
            .configure_axis(labelColor=MUTED, titleColor=MUTED, labelFont="IBM Plex Sans", titleFont="IBM Plex Sans",
                            labelFontSize=12, titleFontSize=12, gridColor="#EEF1F5", domainColor=LIGHT, tickColor=LIGHT)
            .configure_text(font="IBM Plex Sans"))


# ------------------------------------------------------------------ UI
st.set_page_config(page_title="CSV cleaner (demo)", page_icon="🧹", layout="wide")
st.markdown(CSS, unsafe_allow_html=True)

with st.sidebar:
    up = st.file_uploader("Your own CSV or Excel file (optional)", type=["csv", "txt", "xlsx"])
    sources([up.name if up else SAMPLE.name])
    st.markdown("<p class='dp-h' style='margin-top:1.2rem'>Cleaning steps</p>", unsafe_allow_html=True)
    opts = {
        "trim": st.checkbox("Trim extra spaces", True),
        "empty": st.checkbox("Remove empty rows", True),
        "totals": st.checkbox("Remove total rows", True),
        "dupes": st.checkbox("Remove exact duplicates", True),
        "case": st.checkbox("Unify upper/lower case", True),
        "country": st.checkbox("Standardise country names", True),
        "dates": st.checkbox("Convert dates to one format", True),
        "numbers": st.checkbox("Convert amounts to numbers", True),
        "us_slash": st.radio("Dates like 3/4/2026 mean", ["month/day (US)", "day/month"]) == "month/day (US)",
    }
    st.markdown("<div class='dp-note'>Runs in your browser. Your file is not uploaded anywhere and not stored."
                "</div>", unsafe_allow_html=True)

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
n_check = int((clean_df["Issues"] != "").sum())

if amount_col is not None and pos.get(amount_col, 99) < len(raw_cols):
    raw_amount = raw_df.iloc[:, pos[amount_col]].astype(str)
    excel_sum = sum(float(v) for v in raw_amount if plain.match(v))     # what SUM() sees: real numbers only
    clean_sum = float(clean_df[amount_col].sum())
else:
    excel_sum = clean_sum = None
if cust_col is not None and pos.get(cust_col, 99) < len(raw_cols):
    raw_names = raw_df.iloc[:, pos[cust_col]].astype(str)
    raw_n = raw_names[raw_names.str.strip() != ""]
    raw_n = raw_n[~raw_n.str.strip().str.lower().isin(TOTAL_WORDS)].str.lower().nunique()  # like a pivot table
    clean_n = clean_df[cust_col][clean_df[cust_col] != ""].nunique()
else:
    raw_n = clean_n = None

bar("data-cleaner.local · order export")
st.title("Messy export in, clean table out")
st.markdown("<p class='dp-lead'>A typical order export with title rows, duplicates, four spellings of the same "
            "country and amounts stored as text. Upload your own file in the sidebar to try it.</p>",
            unsafe_allow_html=True)

removed = len(raw_df) - len(clean_df)
k1, k2, k3, k4 = st.columns(4)
kpi(k1, "Rows", f"{len(raw_df)} → {len(clean_df)}", f"{removed} duplicate, empty or total rows removed")
if raw_n is not None:
    kpi(k2, f"Distinct {cust_col.lower()}s", f"{raw_n} → {clean_n}", "a pivot table would list each spelling",
        warn=raw_n != clean_n)
else:
    kpi(k2, "Columns", f"{len(clean_df.columns) - 1}", "no name column found")
if excel_sum is not None:
    gap = (excel_sum - clean_sum) / clean_sum if clean_sum else 0
    kpi(k3, f"Total {amount_col.lower()}", f"{cur}{clean_sum:,.0f}",
        f"Excel SUM on the raw file: {cur}{excel_sum:,.0f} ({gap:+.0%})", warn=abs(gap) > 0.01)
else:
    kpi(k3, "Number columns", str(len(num_cols)), "converted from text")
kpi(k4, "Rows to check", str(n_check), "missing values are flagged, never guessed" if n_check else "nothing missing",
    warn=n_check > 0)

parts = []
if excel_sum is not None and abs(excel_sum - clean_sum) > 0.5:
    parts.append(f"A SUM over the raw file gives <b>{cur}{excel_sum:,.0f}</b>, the correct total is "
                 f"<b>{cur}{clean_sum:,.0f}</b>: amounts written as text were skipped and duplicates counted twice.")
if raw_n is not None and raw_n != clean_n:
    parts.append(f"<b>{raw_n} {cust_col.lower()} names</b> are really <b>{clean_n}</b>.")
parts.append(f"<b>{n_check} rows</b> miss a value and need a short look." if n_check
             else "No missing values.")
message(" ".join(parts), "WHAT CHANGED, IN ONE SENTENCE")

left, right = st.columns([1, 1], gap="large")
with left:
    rep = report[~report["Step"].str.startswith(("Missing", "Skip"))].copy()
    rep = rep[rep["Changed"] > 0]
    rep["Step"] = rep["Step"].str.replace(r" \(.*\)$", "", regex=True)
    rep["Step"] = (rep["Step"].str.replace(r"^Dates in '(.+)' to one format.*$", r"Dates in \1", regex=True)
                   .str.replace(r"^Numbers in '(.+)'.*$", r"Amounts in \1", regex=True))
    if len(rep):
        topstep = rep.sort_values("Changed", ascending=False).iloc[0]["Step"]
        heading("What was wrong in the export", "Cells or rows changed per cleaning step")
        ch = alt.Chart(rep).encode(
            y=alt.Y("Step:N", sort="-x", title=None, axis=alt.Axis(labelLimit=260, ticks=False, domain=False,
                                                                   labelPadding=8)),
            x=alt.X("Changed:Q", axis=None, scale=alt.Scale(domain=[0, rep["Changed"].max() * 1.2])))
        b = ch.mark_bar(cornerRadiusEnd=3, height=18).encode(
            color=alt.condition(alt.datum.Step == topstep, alt.value(NAVY), alt.value(LIGHT)))
        t = ch.mark_text(align="left", dx=6, color=BODY, fontSize=12).encode(text="Changed:Q")
        st.altair_chart(chart_style((b + t).properties(height=34 * len(rep) + 10)), use_container_width=True)
with right:
    if amount_col and cust_col:
        by = clean_df.groupby(cust_col)[amount_col].sum().sort_values(ascending=False).reset_index()
        by.columns = ["Name", "Value"]
        share = by["Value"].iloc[0] / by["Value"].sum() if by["Value"].sum() else 0
        heading(f"{by['Name'].iloc[0]} is the largest {cust_col.lower()} ({share:.0%})",
                f"{amount_col} per {cust_col.lower()}, after cleaning. Only now can you add them up.")
        top10 = by.head(10)
        ch = alt.Chart(top10).encode(
            y=alt.Y("Name:N", sort="-x", title=None, axis=alt.Axis(labelLimit=200, ticks=False, domain=False,
                                                                   labelPadding=8)),
            x=alt.X("Value:Q", axis=None, scale=alt.Scale(domain=[0, top10["Value"].max() * 1.25])))
        b = ch.mark_bar(cornerRadiusEnd=3, height=18).encode(
            color=alt.condition(alt.datum.Name == by["Name"].iloc[0], alt.value(NAVY), alt.value(LIGHT)))
        t = ch.mark_text(align="left", dx=6, color=BODY, fontSize=12).encode(
            text=alt.Text("Value:Q", format=",.0f"))
        st.altair_chart(chart_style((b + t).properties(height=34 * len(top10) + 10)), use_container_width=True)

tab1, tab2, tab3 = st.tabs(["Clean table", "Before / after", "Cleaning log"])
dcfg = {c: st.column_config.DateColumn(format="YYYY-MM-DD") for c in date_cols}
dcfg.update({c: st.column_config.NumberColumn(format="%.2f") for c in num_cols
             if not str(clean_df[c].dtype).startswith("Int")})
with tab1:
    flagged = clean_df["Issues"] != ""
    st.dataframe(clean_df.style.apply(lambda r: ["background-color:#FBEDE4" if flagged[r.name] else ""] * len(r),
                                      axis=1),
                 hide_index=True, use_container_width=True, height=380, column_config=dcfg)
    st.caption("Rows that need a look are shaded. The 'Issues' column says why.")
with tab2:
    l, r = st.columns(2)
    l.markdown("<p class='dp-sub'>Before (as exported)</p>", unsafe_allow_html=True)
    l.dataframe(raw_df, hide_index=True, use_container_width=True, height=380)
    r.markdown("<p class='dp-sub'>After</p>", unsafe_allow_html=True)
    r.dataframe(clean_df, hide_index=True, use_container_width=True, height=380, column_config=dcfg)
with tab3:
    st.dataframe(report, hide_index=True, use_container_width=True, height=(len(report) + 1) * 35 + 3)
    st.caption("Every step is logged, so you can check what was changed and why.")

d1, d2 = st.columns(2)
d1.download_button("Download clean Excel (data + log)", to_excel(clean_df, report),
                   file_name="clean_data.xlsx", type="primary", use_container_width=True)
d2.download_button("Download clean CSV", clean_df.to_csv(index=False).encode("utf-8"),
                   file_name="clean_data.csv", mime="text/csv", use_container_width=True)

steps("Every month someone fixes the same export by hand: delete title rows, merge spellings, retype amounts.",
      "Applies the same rules in seconds, logs every change and flags what it cannot decide.",
      "Numbers you can add up. Try it: untick <b>Convert amounts to numbers</b> and watch the total drop.")
