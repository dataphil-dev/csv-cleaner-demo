"""CSV / Excel cleaner demo.

DataPhil demo - the sample file is made up. Upload your own file to try it; it is processed
in memory only and never stored.
Run locally:  pip install -r requirements.txt && streamlit run app.py
"""
import csv
import io
import re
from datetime import datetime
from pathlib import Path

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


# ------------------------------------------------------------------ UI
st.set_page_config(page_title="CSV cleaner (demo)", page_icon="🧹", layout="wide")
st.title("Messy CSV / Excel cleaner")
st.caption("DataPhil demo · sample data is made up · your own file stays in memory and is never stored")

with st.sidebar:
    st.header("1  Choose a file")
    up = st.file_uploader("CSV or Excel", type=["csv", "txt", "xlsx"])
    st.write("No file? The made-up order export below is used.")
    st.header("2  Cleaning steps")
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

name, raw = (up.name, up.getvalue()) if up else (SAMPLE.name, SAMPLE.read_bytes())
try:
    raw_df = read_any(name, raw)
except Exception as e:  # noqa: BLE001 - show any reading problem to the user
    st.error(f"Could not read this file: {e}")
    st.stop()

clean_df, report, date_cols, num_cols = clean(raw_df, opts)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Rows before", len(raw_df))
c2.metric("Rows after", len(clean_df), f"{len(clean_df) - len(raw_df)} rows", delta_color="off")
c3.metric("Cells changed", int(report.loc[report["Step"] != "Missing values flagged (not filled)", "Changed"].sum()))
c4.metric("Rows needing a look", int((clean_df["Issues"] != "").sum()))

tab1, tab2, tab3 = st.tabs(["Before / after", "Cleaning report", "Quick analysis"])
with tab1:
    l, r = st.columns(2)
    l.subheader("Before")
    l.dataframe(raw_df, hide_index=True, use_container_width=True, height=420)
    r.subheader("After")
    r.dataframe(clean_df, hide_index=True, use_container_width=True, height=420,
                column_config={c: st.column_config.DateColumn(format="YYYY-MM-DD") for c in date_cols})
with tab2:
    st.dataframe(report, hide_index=True, use_container_width=True, height=(len(report) + 1) * 35 + 3)
    st.caption("Every step is logged, so the client can check what was changed and why. "
               "Missing values are flagged, not guessed.")
with tab3:
    cat_cols = [c for c in clean_df.columns if c not in date_cols + num_cols + ["Issues"]]
    if num_cols and cat_cols:
        groupable = [c for c in cat_cols if 2 <= clean_df[c].nunique() <= 30] or cat_cols
        g = st.selectbox("Group by", groupable)
        v = st.selectbox("Sum of", num_cols, index=len(num_cols) - 1)
        st.bar_chart(clean_df.groupby(g)[v].sum().sort_values(), horizontal=True, color="#14a800")
        st.caption("Before cleaning, the same customer appeared under several spellings and amounts were text.")
    else:
        st.info("No numeric and text columns found for a quick chart.")

st.subheader("3  Download")
d1, d2 = st.columns(2)
d1.download_button("Download clean Excel (data + report)", to_excel(clean_df, report),
                   file_name="clean_data.xlsx", type="primary", use_container_width=True)
d2.download_button("Download clean CSV", clean_df.to_csv(index=False).encode("utf-8"),
                   file_name="clean_data.csv", mime="text/csv", use_container_width=True)
