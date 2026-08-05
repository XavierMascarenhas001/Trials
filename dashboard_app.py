import re
import difflib
from datetime import datetime
 
import pandas as pd
import streamlit as st
import plotly.express as px
 
# ============================================================
# CONFIG - adjust these if your real column names differ
# ============================================================
CONFIG = {
    "district_col": "shire",
    "project_col": "project",
    "circuit_col": "segmentcode",
    "item_col": "item",
    "pole_col": "pole",            # -> displayed as "enid"
    "pid_col": "pid_ohl_nr",       # -> displayed as "PID"
    "qsub_col": "qsub",
    "total_col": "total",
    "orig_col": "orig",
    "date_col_candidates": ["datetouse", "plan1", "done"],
    # "job" wasn't in the export script you shared - guessing sourcefile.
    # Change this to the real column if it's something else.
    "job_col": "sourcefile",
}
 
st.set_page_config(page_title="Network Job Tracker", layout="wide")
 
# ============================================================
# YOUR MAPPING DICTIONARIES (unchanged from your script)
# ============================================================
CV7_erect = {
    "Erect Single HV/EHV Pole, up to and including 12 metre pole": "CV7 HV pole",
    "Erect Single HV/EHV Pole, up to and including 12 metre pole.": "CV7 HV pole",
}
CV7_erect_H = {
    "Erect Section Structure 'H' HV/EHV Pole, up to and including 12 metre pole.": "CV7 HV pole"
}
CV7_erect_lv = {
    "Erect LV Structure Single Pole, up to and including 12 metre pole": "CV7 LV pole",
}
CV7_recover = {
    "Recover single pole, up to and including 15 metres in height, and reinstate, all ground conditions": "CV7",
    "Recover 'A' / 'H' pole, up to and including 15 metres in height, and reinstate, all ground conditions": "CV7 HV pole",
}
CV7_Tx = {
    "Erect pole mounted transformer up to 100kVA 1.ph.": "CV7 Tx",
    "Erect pole mounted transformer up to 200kVA 3.p.h.": "CV7 Tx",
    "Erect Voltage Regulator.": "CV7 Tx",
    "Erect Voltage Transformer (VT), RTU or Repeater": "CV7 Tx",
    "Erect 12kV/36kV Surge arrestors ( directly mounted ).": "CV7 Tx",
    "Remove pole mounted tranformer.": "CV7 Tx",
    "Remove platform mounted or 'H' pole mounted transformer.": "CV7 Tx",
}
transformer = {
    "Transformer 1ph 50kVA": "TX 1ph (50kVA)",
    "Transformer 3ph 50kVA": "TX 3ph (50kVA)",
    "Transformer 1ph 100kVA": "TX 1ph (100kVA)",
    "Transformer 1ph 25kVA": "TX 1ph (25kVA)",
    "Transformer 3ph 200kVA": "TX 3ph (200kVA)",
    "Transformer 3ph 100kVA": "TX 3ph (100kVA)",
}
CV7_OHL_CONDUCTOR_instal = {
    "Install bare conductor, run out, sag, terminate, bind in and connect jumpers; <100mm²": "CV7 OHL CONDUCTOR",
    "Install bare conductor, run out, sag, terminate, bind in and connect jumpers; >=100mm² <200mm²": "CV7 OHL CONDUCTOR",
    "Install conductor, run out, sag, terminate, clamp in and form jumper loops; >=200mm²": "CV7 OHL CONDUCTOR",
}
CV7_OHL_CONDUCTOR_recover = {
    "Recover overhead wire and fittings; HV/EHV overhead line or Hardex Pilot (1 conductor)": "CV7 OHL CONDUCTOR",
    "Recover overhead wire and fittings; HV/EHV overhead line or Hardex Pilot (2 conductor)": "CV7 OHL CONDUCTOR",
    "Recover overhead wire and fittings; HV/EHV overhead line or Hardex Pilot (3 conductor)": "CV7 OHL CONDUCTOR",
}
CV7_OHL_CONDUCTOR_LV_instal = {
    "Install conductor, run out, sag, terminate, clamp in and connect jumpers; 2c": "CV7 OHL CONDUCTOR LV",
    "Install conductor, run out, sag, terminate, clamp in and connect jumpers; 4c": "CV7 OHL CONDUCTOR LV",
    "Install conductor, run out, sag, terminate, clamp in and connect jumpers; 2c + Earth": "CV7 OHL CONDUCTOR LV",
    "Install conductor, run out, sag, terminate, clamp in and connect jumpers; 4c + Earth": "CV7 OHL CONDUCTOR LV",
}
CV7_OHL_CONDUCTOR_LV_recover = {
    "Recover overhead wires and fittings; LV openwire overhead line (2 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV openwire overhead line (3 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV openwire overhead line (4 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV openwire overhead line (5 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV service overhead line (open, concentric or ABC, 2 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV service overhead line (open, concentric or ABC, 3 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV service overhead line (open, concentric or ABC, 4 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover overhead wires and fittings; LV service overhead line (open, concentric or ABC, 5 conductors)": "CV7 OHL CONDUCTOR LV",
    "Recover cleated service": "CV7 OHL CONDUCTOR LV",
}
CV7_SWITCHGEAR = {
    "Erect 11kV/33kV ABSW": "CV7 SWITCHGEAR",
    "Erect 11kV Remote Controlled Switch Disconnector ( Soule Auguste ) or Auto Reclosure unit c/w VT, Aerial, RTU & umbilical cable.": "CV7 SWITCHGEAR",
    "Erect 1.ph fuse units at single tee off pole or in line pole.": "CV7 SWITCHGEAR",
    "Erect 3.ph fuse units at single tee off pole or in line pole.": "CV7 SWITCHGEAR",
    "Additional cost for fitting fuse outrigger bracket.": "CV7 SWITCHGEAR",
    "Remove 11kV/33kV ABSW": "CV7 SWITCHGEAR",
}
CV7_UG = {
    "Installation of cable only in trench dug by others; 11kV Cable 3 x 1 core.": "CV7 UG 11 kV",
    "Install cable in existing duct; 11kV Cable 3 x 1 core.": "CV7 UG 11 kV",
    "Installation of cable only in trench dug by others; 33kV Cable 3 x 1 core.": "CV7 UG 33 kV",
    "Install cable in existing duct; 33kV Cable 3 x 1 core.": "CV7 UG 33 kV",
    "Installation of cable only in trench dug by others; LV Cable Large or 11kV Cable 1 x 3 Core": "CV7 UG",
    "Install cable in existing duct; LV Cable Large or 11kV Cable 1 x 3 Core": "CV7 UG",
    "Installation of cable only in trench dug by others; LV Service, Small LV or Pilot Cable.": "CV7 UG LV Service",
    "Install cable in existing duct; LV Service, Small LV or Pilot Cable.": "CV7 UG LV Service",
}
CV7_CB = {"Remove Auto Reclosure.": "CV7 CB"}
Switch = {
    "Noja": "Noja",
    "11kV PMSW (Soule)": "11kV PMSW (Soule)",
    "11kv ABSW Hookstick Standard": "11kv ABSW Hookstick Standard",
    "11kv ABSW Hookstick Spring loaded mech": "11kv ABSW Hookstick Spring loaded mech",
    "33kv ABSW Hookstick Dependant": "33kv ABSW Hookstick Dependant",
}
Fuses = {
    "100A LV Fuse JPU 82.5mm": "100A LV Fuse JPU 82.5mm",
    "160A LV Fuse JPU 82.5mm": "160A LV Fuse JPU 82.5mm",
    "200A LV Fuse JPU 82.5mm": "200A LV Fuse JPU 82.5mm",
    "315A LV Fuse JPU 82.5mm": "315A LV Fuse JPU 82.5mm",
    "400A LV Fuse JPU 82.5mm": "400A LV Fuse JPU 82.5mm",
    "200A LV Fuse JSU 92mm": "200A LV Fuse JSU 92mm",
    "315A LV Fuse JSU 92mm": "315A LV Fuse JSU 92mm",
    "400A LV Fuse JSU 92mm": "400A LV Fuse JSU 92mm",
    "100A LV Fuse - Porcelain screw-in": "100A LV Fuse - Porcelain screw-in",
    "160A LV Fuse - Porcelain screw-in": "160A LV Fuse - Porcelain screw-in",
    "200A LV Fuse - Porcelain screw-in": "200A LV Fuse - Porcelain screw-in",
    "Single Phase cut out kit 100A Henley Series 7": "Single Phase cut out kit 100A Henley Series 7",
    "Three Phase cut out kit 100A Henley Series 7": "Three Phase cut out kit 100A Henley Series 7",
    "Three Phase 200A Cut out": "Three Phase 200A Cut out",
    "Cut out Fuse (MF) 60A": "Cut out Fuse (MF) 60A",
    "Cut out Fuse (MF) 80A": "Cut out Fuse (MF) 80A",
    "Cut out Fuse (MF) 100A": "Cut out Fuse (MF) 100A",
    "11KV FUSE UNIT - C-TYPE": "11KV FUSE UNIT - C-TYPE",
    "11KV SOLID LINK - C-TYPE": "11KV SOLID LINK - C-TYPE",
}
CV31 = {
    "Replace / Fit safety or warning sign, number plates or name plate": "CV31",
    "Barbed Wire Wrap ACD (or Enhanced) single pole or stay - Replace/Repair": "CV31",
    "Steelwork bonding repair / fit.": "CV31",
    "Replace LV/HV/Earth guard missing / damaged.": "CV31",
}
CV8 = {
    "Tighten existing stay.": "CV8",
    "Erect/Replace stay above ground only.": "CV8",
    "Erect/Replace stay complete including block or driven type anchor": "CV8",
    "Erect/Replace stay complete including rock type anchor": "CV8",
}  # (trimmed here for brevity - paste your full CV8 dict back in if you need every line mapped)
 
POLE_CATEGORIES = {
    "CV7_erect": CV7_erect,
    "CV7_erect_H": CV7_erect_H,
    "CV7_erect_lv": CV7_erect_lv,
    "CV7_recover": CV7_recover,
}
ALL_CATEGORIES = {
    **POLE_CATEGORIES,
    "CV7_Tx": CV7_Tx,
    "transformer": transformer,
    "CV7_OHL_CONDUCTOR_instal": CV7_OHL_CONDUCTOR_instal,
    "CV7_OHL_CONDUCTOR_recover": CV7_OHL_CONDUCTOR_recover,
    "CV7_OHL_CONDUCTOR_LV_instal": CV7_OHL_CONDUCTOR_LV_instal,
    "CV7_OHL_CONDUCTOR_LV_recover": CV7_OHL_CONDUCTOR_LV_recover,
    "CV7_SWITCHGEAR": CV7_SWITCHGEAR,
    "CV7_UG": CV7_UG,
    "CV7_CB": CV7_CB,
    "Switch": Switch,
    "Fuses": Fuses,
    "CV31": CV31,
    "CV8": CV8,
}
 
HV_POLE_KEY = "Recover 'A' / 'H' pole, up to and including 15 metres in height, and reinstate, all ground conditions"
HV_POLE_MULTIPLIER = 2
 
# ============================================================
# HELPERS (same normalization logic as your export script)
# ============================================================
def normalize_item(x):
    if pd.isna(x):
        return ""
    s = str(x).replace("\u200b", "").replace("\xa0", "").strip().upper()
    return re.sub(r"\s+", " ", s)
 
 
def clean_job(value):
    """Strip leading 'C -'/'M -' prefix, cut at 'map', drop SPxxxx/GSPxxxx/bare digits."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    text = re.sub(r'^[A-Za-z]\s*-\s*', '', text)
    m = re.search(r'map', text, flags=re.IGNORECASE)
    if m:
        text = text[: m.start()]
    text = re.sub(r'\bGSP\d+\b', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\bSP\d+\b', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\b\d+\b', '', text)
    text = re.sub(r'\s{2,}', ' ', text)
    text = re.sub(r'^[\s\-\u2013_,.:]+|[\s\-\u2013_,.:]+$', '', text)
    return text.strip()
 
 
def dedupe_jobs(values, threshold=0.65):
    kept, is_dup = [], []
    for v in values:
        if not v:
            is_dup.append(False)
            continue
        hit = any(difflib.SequenceMatcher(None, v.lower(), k.lower()).ratio() >= threshold for k in kept)
        is_dup.append(hit)
        if not hit:
            kept.append(v)
    return is_dup
 
 
@st.cache_data
def read_file(file) -> pd.DataFrame:
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_parquet(file)
    df.columns = df.columns.str.strip().str.lower()
    return df
 
 
def process_data(df: pd.DataFrame, cols: dict) -> pd.DataFrame:
    """cols: the resolved {logical_name: real_column_name} mapping picked in the sidebar."""
    df = df.copy()
    item_col = cols["item_col"]
    qsub_col = cols["qsub_col"]
 
    df["_item_norm"] = df[item_col].apply(normalize_item)
 
    # HV pole recovery counts double
    hv_key_norm = normalize_item(HV_POLE_KEY)
    hv_mask = df["_item_norm"] == hv_key_norm
    df["_qsub_adj"] = pd.to_numeric(df[qsub_col], errors="coerce").fillna(0)
    df.loc[hv_mask, "_qsub_adj"] *= HV_POLE_MULTIPLIER
 
    # map every item to its category
    item_to_cat = {}
    for cat_name, mapping in ALL_CATEGORIES.items():
        for desc, label in mapping.items():
            item_to_cat[normalize_item(desc)] = label
    df["_mapped_category"] = df["_item_norm"].map(item_to_cat)
 
    # cleaned job column
    job_col = cols.get("job_col")
    if job_col and job_col in df.columns:
        df["_job_clean"] = df[job_col].apply(clean_job)
    else:
        df["_job_clean"] = ""
 
    # date column
    date_col = cols.get("date_col")
    df["_date"] = pd.to_datetime(df[date_col], errors="coerce") if date_col and date_col in df.columns else pd.NaT
 
    return df
 
 
# ============================================================
# APP
# ============================================================
st.title("Network Job Tracker Dashboard")
 
uploaded = st.file_uploader("Upload your parquet or CSV file", type=["parquet", "csv"])
if not uploaded:
    st.info("Upload the parquet/CSV file your export script normally reads, then filters and charts appear below.")
    st.stop()
 
raw_df = read_file(uploaded)
 
with st.expander("Detected columns in your file (click to view)"):
    st.write(list(raw_df.columns))
 
 
def guess(*candidates):
    for c in candidates:
        if c in raw_df.columns:
            return c
    return raw_df.columns[0]
 
 
st.sidebar.header("Column mapping")
st.sidebar.caption("Match each field to your file's real column names.")
col_options = list(raw_df.columns)
 
def pick(label, default_col):
    idx = col_options.index(default_col) if default_col in col_options else 0
    return st.sidebar.selectbox(label, col_options, index=idx)
 
none_option = ["(none)"] + col_options
 
def pick_optional(label, default_col):
    opts = none_option
    idx = opts.index(default_col) if default_col in opts else 0
    val = st.sidebar.selectbox(label, opts, index=idx)
    return None if val == "(none)" else val
 
cols = {
    "item_col": pick("Description / item", guess("item", "description")),
    "qsub_col": pick("Quantity (qsub)", guess("qsub", "quantity_used")),
    "district_col": pick("District", guess("shire", "district")),
    "project_col": pick("Project", guess("project")),
    "circuit_col": pick("Circuit", guess("segmentcode", "circuit")),
    "pole_col": pick("Pole / enid", guess("pole", "enid")),
    "pid_col": pick("PID", guess("pid_ohl_nr", "pid")),
    "total_col": pick("Total value", guess("total")),
    "orig_col": pick("Original value", guess("orig", "original")),
    "job_col": pick_optional("Job", guess("sourcefile", "job") if "sourcefile" in raw_df.columns or "job" in raw_df.columns else None),
    "date_col": pick_optional("Date", guess("datetouse", "plan1", "done") if any(c in raw_df.columns for c in ["datetouse", "plan1", "done"]) else None),
}
 
df = process_data(raw_df, cols)
 
# ---- Sidebar filters ----
st.sidebar.header("Filters")
 
district_col, project_col, circuit_col = cols["district_col"], cols["project_col"], cols["circuit_col"]
 
if df["_date"].notna().any():
    min_d, max_d = df["_date"].min(), df["_date"].max()
    date_range = st.sidebar.date_input("Date", value=(min_d.date(), max_d.date()))
else:
    date_range = None
 
def multiselect_filter(label, col):
    if col not in df.columns:
        return []
    opts = sorted(df[col].dropna().astype(str).unique())
    return st.sidebar.multiselect(label, opts)
 
districts = multiselect_filter("District", district_col)
projects = multiselect_filter("Project", project_col)
circuits = multiselect_filter("Circuit", circuit_col)
 
f = df.copy()
if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
    start, end = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
    f = f[(f["_date"] >= start) & (f["_date"] <= end) | f["_date"].isna()]
if districts:
    f = f[f[district_col].astype(str).isin(districts)]
if projects:
    f = f[f[project_col].astype(str).isin(projects)]
if circuits:
    f = f[f[circuit_col].astype(str).isin(circuits)]
 
# ---- Pole bar chart ----
st.subheader("Poles")
pole_keys = set()
for mapping in POLE_CATEGORIES.values():
    pole_keys |= {normalize_item(k) for k in mapping}
pole_df = f[f["_item_norm"].isin(pole_keys)]
pole_summary = (
    pole_df.groupby("_mapped_category")["_qsub_adj"].sum().reset_index()
    .rename(columns={"_mapped_category": "Pole type", "_qsub_adj": "Count"})
)
if not pole_summary.empty:
    fig = px.bar(pole_summary, x="Pole type", y="Count", text="Count")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.caption("No pole records for the current filters.")
 
# ---- Job / Circuit / PID table (deduped, scrollable) ----
st.subheader("Jobs")
job_table = f[[cols["job_col"], circuit_col, cols["pid_col"]]].copy() if cols["job_col"] in f.columns else pd.DataFrame()
if not job_table.empty:
    job_table["job"] = f["_job_clean"]
    job_table = job_table[["job", circuit_col, cols["pid_col"]]].rename(
        columns={circuit_col: "Circuit", cols["pid_col"]: "PID"}
    )
    job_table["_dup"] = dedupe_jobs(job_table["job"].tolist())
    job_table = job_table[~job_table["_dup"]].drop(columns="_dup")
    st.dataframe(job_table, height=350, use_container_width=True)  # fixed height -> scrolls internally
else:
    st.caption("No job column selected in the sidebar mapping.")
 
# ---- Per-mapped-item breakdown ----
st.subheader("Mapped items")
for cat_name, mapping in ALL_CATEGORIES.items():
    keys = {normalize_item(k) for k in mapping}
    sub = f[f["_item_norm"].isin(keys)]
    if sub.empty:
        continue
    total_qty = sub["_qsub_adj"].sum()
    with st.expander(f"{cat_name} - {total_qty:,.0f}"):
        detail = sub[[district_col, "_job_clean", circuit_col, cols["pole_col"]]].rename(
            columns={
                district_col: "District",
                "_job_clean": "Job",
                circuit_col: "Circuit",
                cols["pole_col"]: "enid",
            }
        )
        st.dataframe(detail, height=250, use_container_width=True)
 
# ---- Totals & variance ----
st.subheader("Totals")
total_val = pd.to_numeric(f[cols["total_col"]], errors="coerce").sum() if cols["total_col"] in f.columns else None
orig_val = pd.to_numeric(f[cols["orig_col"]], errors="coerce").sum() if cols["orig_col"] in f.columns else None
 
c1, c2 = st.columns(2)
if total_val is not None:
    c1.metric("Total value", f"£{total_val:,.2f}")
if total_val is not None and orig_val is not None:
    c2.metric("Difference vs original", f"£{total_val - orig_val:,.2f}")
 
if total_val is not None and orig_val is not None:
    f["_row_variance"] = pd.to_numeric(f[cols["total_col"]], errors="coerce") - pd.to_numeric(f[cols["orig_col"]], errors="coerce")
    variance_rows = f[f["_row_variance"] != 0]
    variance_table = variance_rows[[district_col, "_job_clean", circuit_col]].rename(
        columns={district_col: "District", "_job_clean": "Job", circuit_col: "Circuit"}
    )
    variance_table["is_dup"] = dedupe_jobs(variance_table["Job"].tolist())
    variance_table = variance_table[~variance_table["is_dup"]].drop(columns="is_dup")
    st.caption("Rows where total differs from the original value")
    st.dataframe(variance_table, height=300, use_container_width=True)
