import os
import re
import difflib
from datetime import datetime
 
import pandas as pd
import streamlit as st
import plotly.express as px
 
st.set_page_config(page_title="Network Job Tracker", layout="wide")
 
# ============================================================
# YOUR MAPPING DICTIONARIES
# (CV8 and Fuses below are now the FULL dictionaries copied from the
#  export tool - the dashboard versions were truncated, which alone
#  explained a chunk of the discrepancy)
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
    "11KV OHL ASL C-TYPE RESET 20A 2 SHOT": "11KV OHL ASL C-TYPE RESET 20A 2 SHOT",
    "11KV OHL ASL C-TYPE RESET 25A 2 SHOT": "11KV OHL ASL C-TYPE RESET 25A 2 SHOT",
    "11KV OHL ASL C-TYPE RESET 40A 1 SHOT": "11KV OHL ASL C-TYPE RESET 40A 1 SHOT",
    "11KV OHL ASL C-TYPE RESET 40A 2 SHOT": "11KV OHL ASL C-TYPE RESET 40A 2 SHOT",
    "11KV OHL ASL C-TYPE RESET 63A 1 SHOT": "11KV OHL ASL C-TYPE RESET 63A 1 SHOT",
    "11KV OHL ASL C-TYPE RESET 63A 2 SHOT": "11KV OHL ASL C-TYPE RESET 63A 2 SHOT",
    "11KV OHL ASL C-TYPE RESET 63A 3 SHOT": "11KV OHL ASL C-TYPE RESET 63A 3 SHOT",
    "11KV OHL ASL C-TYPE RESET 100A 1 SHOT": "11KV OHL ASL C-TYPE RESET 100A 1 SHOT",
    "11KV OHL ASL C-TYPE RESET 100A 2 SHOT": "11KV OHL ASL C-TYPE RESET 100A 2 SHOT",
    "11KV OHL ASL C-TYPE RESET 100A 3 SHOT": "11KV OHL ASL C-TYPE RESET 100A 3 SHOT",
    "11KV OHL FUSE ELEMENT C-TYPE 15A": "11KV OHL FUSE ELEMENT C-TYPE 15A",
    "11KV OHL FUSE ELEMENT C-TYPE 25A": "11KV OHL FUSE ELEMENT C-TYPE 25A",
    "11KV OHL FUSE ELEMENT C-TYPE 30A": "11KV OHL FUSE ELEMENT C-TYPE 30A",
    "11KV OHL FUSE ELEMENT C-TYPE 40A": "11KV OHL FUSE ELEMENT C-TYPE 40A",
    "11KV OHL FUSE ELEMENT C-TYPE 50A": "11KV OHL FUSE ELEMENT C-TYPE 50A",
    "11KV OHL ASL DJP-TYPE 20A 2 SHOT": "11KV OHL ASL DJP-TYPE 20A 2 SHOT",
    "11KV OHL ASL DJP-TYPE 25A 1 SHOT": "11KV OHL ASL DJP-TYPE 25A 1 SHOT",
    "11KV OHL ASL DJP-TYPE 25A 2 SHOT": "11KV OHL ASL DJP-TYPE 25A 2 SHOT",
    "11KV OHL ASL DJP-TYPE 40A 1 SHOT": "11KV OHL ASL DJP-TYPE 40A 1 SHOT",
    "11KV OHL ASL DJP-TYPE 40A 2 SHOT": "11KV OHL ASL DJP-TYPE 40A 2 SHOT",
    "11KV OHL ASL DJP-TYPE 63A 1 SHOT": "11KV OHL ASL DJP-TYPE 63A 1 SHOT",
    "11KV OHL ASL DJP-TYPE 63A 2 SHOT": "11KV OHL ASL DJP-TYPE 63A 2 SHOT",
    "11KV OHL ASL DJP-TYPE 63A 3 SHOT": "11KV OHL ASL DJP-TYPE 63A 3 SHOT",
    "11KV OHL ASL DJP-TYPE 100A 1 SHOT": "11KV OHL ASL DJP-TYPE 100A 1 SHOT",
    "11KV OHL ASL DJP-TYPE 100A 2 SHOT": "11KV OHL ASL DJP-TYPE 100A 2 SHOT",
    "11KV OHL ASL DJP-TYPE 100A 3 SHOT": "11KV OHL ASL DJP-TYPE 100A 3 SHOT",
    "11KV OHL FUSE ELEMENT DJP-TYPE 15A": "11KV OHL FUSE ELEMENT DJP-TYPE 15A",
    "11KV OHL FUSE ELEMENT DJP-TYPE 25A": "11KV OHL FUSE ELEMENT DJP-TYPE 25A",
    "11KV OHL FUSE ELEMENT DJP-TYPE 30A": "11KV OHL FUSE ELEMENT DJP-TYPE 30A",
    "11KV OHL FUSE ELEMENT DJP-TYPE 40A": "11KV OHL FUSE ELEMENT DJP-TYPE 40A",
    "11KV OHL FUSE ELEMENT DJP-TYPE 50A": "11KV OHL FUSE ELEMENT DJP-TYPE 50A",
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
    "Retrofit structure with Anchor Clamp fitting for Section / Angle / Terminal support": "CV8",
    "Erect Single Crossarm to single pole.": "CV8",
    "Erect Double Crossarm 'H' Pole formation": "CV8",
    "Remove Steelwork crossarm item only": "CV8",
    "Change 11kV Insulators to avoid contamination from old conductor": "CV8",
    "Change 33kV Insulators to avoid contamination from old conductor": "CV8",
    "Replace tension insulator, 11kV.": "CV8",
    "Replace tension insulator, 33kV.": "CV8",
    "Additional cost for fitting Stay Outrigger Bracket": "CV8",
    "Additional cost for fitting Angle / Terminal stay attachment plates on Heavy Construction as SP4009862": "CV8",
    "Recover and reinstate stay position,all ground conditions.": "CV8",
    "Fit foundation block to existing pole.": "CV8",
    "Fit bog shoe foundation to existing single pole.": "CV8",
    "Replace jumper / dropper mechanical connection with compression connection": "CV8",
    "Replace jumper / dropper with live line bail and flexible jumper conductor": "CV8",
    "Replace / Repair conductor with mid span joint using compression connection": "CV8",
    "Conductor repair; piece in conductor including compression joints": "CV8",
    "Bind In Conductors; 1.ph 11kV Intermediate / Pin Angle pole.": "CV8",
    "Bind In Conductors; 3.ph 11kV Intermediate / Pin Angle pole.": "CV8",
    "Conductor Terminations - 1.ph 11kV Section pole including jumpers.": "CV8",
    "Conductor Terminations - 3.ph 11kV Section pole including jumpers.": "CV8",
    "Conductor Terminations - 1.ph 11kV Terminal pole.": "CV8",
    "Conductor Terminations - 3.ph 11kV Terminal pole.": "CV8",
    "Unbind and reregulate existing conductors": "CV8",
    "Convert 1.ph 11kV Intermediate pole into Section Pole.": "CV8",
    "Convert 1.ph/3.p.h. 11kV line pole into Terminal Pole.": "CV8",
    "Convert 3.ph 11kV Intermediate pole into Section Pole.": "CV8",
    "Replace 11kV/33kV insulator pin and insulator, including unbinding and binding in": "CV8",
    "Replace 11kV/33kV insulator binder": "CV8",
    "Replace tension insulator, 11kV": "CV8",
    "Replace tension insulator, 33kV": "CV8",
    "Replace 11kV/33kV dead end termination": "CV8",
    "Additional cost for erection of pilot pin and insulator or pilot post insulator (11kV or 33kV)": "CV8",
    "Replace insulated conductor HV/LV earth above ground to first rod": "CV8",
    "Install Copper Covered Green / Yellow HV Earth or Black LV Earth to foot of pole": "CV8",
    "Install EHV/ HV Earth Electrode including excavate & reinstate (up to 8mtrs)": "CV8",
    "Install LV Earth Electrode including excavate & reinstate (up to 28mtrs)": "CV8",
    "Additional extra over for additional earthing excavated, laid & backfilled": "CV8",
    "Install Earth Electrode within cable trench": "CV8",
    "Erect 11kV Cable Termination ( incorporating surge arrestors )": "CV8",
    "Erect 33kV Cable Termination ( incorporating surge arrestors )": "CV8",
    "Steelwork bonding repair / fit": "CV8",
    "Erect 1.ph LV cable pole termination": "CV8",
    "Erect 3.ph LV cable pole termination": "CV8",
    "Remove 11kV/33kV Cable termination": "CV8",
    "Remove LV cable termination": "CV8",
    "Repair pole twist - including unbind / rebind.": "CV8",
}
 
POLE_CATEGORIES = {
    "CV7_erect": CV7_erect,
    "CV7_erect_H": CV7_erect_H,
    "CV7_erect_lv": CV7_erect_lv,
    "CV7_recover": CV7_recover,
}
 
# NOTE: CV7_SWITCHGEAR / CV7_UG / CV7_CB were removed from ALL_CATEGORIES.
# In the export tool, `categories` gets redefined a second time and that
# second definition (the one actually used to build sheets/Summary) never
# includes these three - so the exporter never produces a "true" value for
# them. Showing cards for them here would just be comparing against
# nothing. If you do want them tracked, they need to be added back into
# the export tool's `categories`/`extra_categories` lists first.
ALL_CATEGORIES = {
    **POLE_CATEGORIES,
    "CV7_Tx": CV7_Tx,
    "transformer": transformer,
    "CV7_OHL_CONDUCTOR_instal": CV7_OHL_CONDUCTOR_instal,
    "CV7_OHL_CONDUCTOR_recover": CV7_OHL_CONDUCTOR_recover,
    "CV7_OHL_CONDUCTOR_LV_instal": CV7_OHL_CONDUCTOR_LV_instal,
    "CV7_OHL_CONDUCTOR_LV_recover": CV7_OHL_CONDUCTOR_LV_recover,
    "Switch": Switch,
    "Fuses": Fuses,
    "CV31": CV31,
    "CV8": CV8,
}
 
# Categories that use the exporter's process_cv() logic instead of a plain
# sum: dedupe by pole, drop zero-qty rows, exclude poles already counted
# under a CV7 erect/recover item, then COUNT distinct poles (not sum qty).
POLE_DEDUPE_CATEGORIES = {"CV8", "CV31"}
 
# ============================================================
# IMAGE GROUPS for the Mapped Items tab
# Images live in an "Images" folder alongside this script (same repo path).
# Any category not listed in a group falls through to the "Other items"
# section at the end, in its normal ALL_CATEGORIES order.
# ============================================================
IMAGE_DIR = "Images"
CARD_GROUPS = [
    {
        "title": "Poles",
        "image": os.path.join(IMAGE_DIR, "Poles.png"),
        "categories": ["CV7_recover", "CV7_erect", "CV7_erect_H", "CV7_erect_lv"],
    },
    {
        "title": "Transformers",
        "image": os.path.join(IMAGE_DIR, "Transformer.png"),
        "categories": ["CV7_Tx", "transformer"],
    },
    {
        "title": "Cable",
        "image": os.path.join(IMAGE_DIR, "Cable.png"),
        "categories": [
            "CV7_OHL_CONDUCTOR_instal",
            "CV7_OHL_CONDUCTOR_recover",
            "CV7_OHL_CONDUCTOR_LV_instal",
            "CV7_OHL_CONDUCTOR_LV_recover",
        ],
    },
]
 
HV_POLE_KEY = "Recover 'A' / 'H' pole, up to and including 15 metres in height, and reinstate, all ground conditions"
HV_POLE_MULTIPLIER = 2
 
# ============================================================
# HELPERS (aligned with the export script's normalization)
# ============================================================
def normalize_item(x):
    if pd.isna(x):
        return ""
    s = str(x).replace("\u200b", "").replace("\u200e", "").replace("\u200f", "").replace("\xa0", "").strip().upper()
    return re.sub(r"\s+", " ", s)
 
 
def normalize_pole(p):
    """Mirrors the export tool's normalize_pole(): strips zero-width/
    directional/nbsp characters, uppercases, and removes ALL whitespace
    (not just collapsing it) so pole IDs compare exactly like the exporter."""
    if pd.isna(p):
        return ""
    s = str(p).replace("\u200b", "").replace("\u200e", "").replace("\u200f", "").replace("\xa0", "").strip().upper()
    return re.sub(r"\s+", "", s)
 
 
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
 
 
# Native unit each category's qsub is recorded in - used to format the
# mapped-item cards for conductor lengths.
UNIT_CONFIG = {
    "CV7_OHL_CONDUCTOR_recover": "m",
    "CV7_OHL_CONDUCTOR_LV_recover": "m",
    "CV7_OHL_CONDUCTOR_instal": "km",
    "CV7_OHL_CONDUCTOR_LV_instal": "km",
}
 
 
def format_length(value, native_unit):
    """Converts value (in native_unit) to meters, then displays in km if
    the meters equivalent is >=1000, otherwise in meters."""
    meters = value * 1000 if native_unit == "km" else value
    if meters >= 1000:
        return f"{meters / 1000:,.2f} km"
    return f"{meters:,.0f} m"
 
 
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
 
    # Raw (un-adjusted) quantity - mirrors what process_cv() reads directly
    # from "qsub" in the export tool, before any HV-multiplier adjustment.
    df["_qsub_raw"] = pd.to_numeric(df[qsub_col], errors="coerce").fillna(0)
 
    # HV pole recovery counts double (applied like the export tool's
    # df.loc[hv_pole_mask, col] *= HV_POLE_MULTIPLIER)
    hv_key_norm = normalize_item(HV_POLE_KEY)
    hv_mask = df["_item_norm"] == hv_key_norm
    df["_qsub_adj"] = df["_qsub_raw"]
    df.loc[hv_mask, "_qsub_adj"] *= HV_POLE_MULTIPLIER
 
    # map every item to its category
    item_to_cat = {}
    for cat_name, mapping in ALL_CATEGORIES.items():
        for desc, label in mapping.items():
            item_to_cat[normalize_item(desc)] = label
    df["_mapped_category"] = df["_item_norm"].map(item_to_cat)
 
    # normalized pole/enid, needed for CV8/CV31 pole-dedupe logic
    pole_col = cols.get("pole_col")
    if pole_col and pole_col in df.columns:
        df["_pole_norm"] = df[pole_col].apply(normalize_pole)
    else:
        df["_pole_norm"] = ""
 
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
 
 
def cv7_dedupe_poles(frame: pd.DataFrame) -> set:
    """Poles already covered by a CV7 erect/recover item - mirrors the
    export tool's cv7_poles = cv7_set(df, CV7_erect) | cv7_set(df, CV7_erect_H)
    | cv7_set(df, CV7_erect_lv) | cv7_set(df, CV7_recover)."""
    keys = set()
    for mapping in POLE_CATEGORIES.values():
        keys |= {normalize_item(k) for k in mapping}
    poles = set(frame.loc[frame["_item_norm"].isin(keys), "_pole_norm"].dropna())
    poles.discard("")
    return poles
 
 
def cv_pole_resume(frame: pd.DataFrame, mapping: dict, cv7_poles: set) -> pd.DataFrame:
    """Mirrors the export tool's process_cv(): filter by item, drop
    zero-qty rows, dedupe by pole (keep first), exclude poles already
    counted under CV7. The resulting row count == the exporter's metric."""
    keys = {normalize_item(k) for k in mapping}
    sub = frame[frame["_item_norm"].isin(keys)].copy()
    sub = sub[sub["_qsub_raw"] != 0]
    sub = sub.drop_duplicates(subset="_pole_norm")
    sub = sub[~sub["_pole_norm"].isin(cv7_poles)]
    return sub
 
 
def build_card(frame: pd.DataFrame, cat_name: str, mapping: dict, cv7_poles: set):
    """Returns (cat_name, total_qty, sub) for a single category, or None
    if there are no matching rows under the current filters."""
    if cat_name in POLE_DEDUPE_CATEGORIES:
        sub = cv_pole_resume(frame, mapping, cv7_poles)
        if sub.empty:
            return None
        return (cat_name, len(sub), sub)
    else:
        keys = {normalize_item(k) for k in mapping}
        sub = frame[frame["_item_norm"].isin(keys)]
        if sub.empty:
            return None
        return (cat_name, sub["_qsub_adj"].sum(), sub)
 
 
def render_metric(slot, cat_name: str, total_qty):
    if cat_name in UNIT_CONFIG:
        slot.metric(cat_name, format_length(total_qty, UNIT_CONFIG[cat_name]))
    elif cat_name in POLE_DEDUPE_CATEGORIES:
        slot.metric(cat_name, f"{total_qty:,.0f} poles")
    else:
        slot.metric(cat_name, f"{total_qty:,.0f}")
 
 
# ============================================================
# APP
# ============================================================
st.markdown(
    """
    <style>
    div[data-testid="stMetric"] {
        background-color: #f5f7fa;
        border: 1px solid #e3e7ee;
        border-radius: 10px;
        padding: 14px 16px;
    }
    div[data-testid="stMetricLabel"] { font-weight: 600; }
    </style>
    """,
    unsafe_allow_html=True,
)
 
st.title("⚡ Network Job Tracker Dashboard")
 
uploaded = st.file_uploader("Upload your parquet or CSV file", type=["parquet", "csv"])
if not uploaded:
    st.info("Upload the parquet/CSV file your export script normally reads, then filters and charts appear below.")
    st.stop()
 
raw_df = read_file(uploaded)
 
with st.expander("Detected columns in your file (click to view)"):
    st.write(list(raw_df.columns))
 
 
def guess(*candidates):
    """Returns the first candidate that exists as a real column, or None if none match."""
    for c in candidates:
        if c in raw_df.columns:
            return c
    return None
 
 
with st.sidebar.expander("⚙️ Column mapping (advanced)", expanded=False):
    st.caption("Confirm each field maps to the right column in your file.")
    col_options = list(raw_df.columns)
    none_option = ["(none)"] + col_options
 
    def pick(label, default_col, key):
        opts = none_option
        idx = opts.index(default_col) if default_col in opts else 0
        if default_col is None:
            st.warning(f"Couldn't guess a column for **{label}** - pick one.")
        val = st.selectbox(label, opts, index=idx, key=key)
        return None if val == "(none)" else val
 
    cols = {
        "item_col": pick("Description / item", guess("item", "description"), "map_item"),
        "qsub_col": pick("Quantity (qsub)", guess("qsub", "quantity_used"), "map_qsub"),
        "district_col": pick("District", guess("shire", "district"), "map_district"),
        "project_col": pick("Project", guess("project"), "map_project"),
        "circuit_col": pick("Circuit", guess("segmentcode", "circuit"), "map_circuit"),
        "pole_col": pick("Pole / enid", guess("pole", "enid"), "map_pole"),
        "pid_col": pick("PID", guess("pid_ohl_nr", "pid"), "map_pid"),
        "total_col": pick("Total value", guess("total"), "map_total"),
        "orig_col": pick("Original value", guess("orig", "original"), "map_orig"),
        "job_col": pick("Job", guess("job", "sourcefile"), "map_job"),
        "date_col": pick("Date", guess("datetouse", "date", "plan1", "done"), "map_date"),
    }
 
missing_required = [k for k in ["item_col", "qsub_col", "district_col", "circuit_col"] if cols[k] is None]
if missing_required:
    st.error(f"These required fields still need a column picked in the sidebar 'Column mapping' section: {missing_required}")
    st.stop()
 
if cols.get("pole_col") is None:
    st.sidebar.warning("No Pole/enid column selected - CV8 and CV31 counts can't be deduplicated by pole and will be shown as raw row counts, which may not match the export tool.")
 
df = process_data(raw_df, cols)
 
district_col, project_col, circuit_col = cols["district_col"], cols["project_col"], cols["circuit_col"]
 
# ---- Sidebar filters ----
st.sidebar.header("🔍 Filters")
 
if df["_date"].notna().any():
    min_d, max_d = df["_date"].min(), df["_date"].max()
    date_range = st.sidebar.date_input("Date", value=(min_d.date(), max_d.date()))
else:
    date_range = None
    st.sidebar.caption("No usable dates found in the selected Date column.")
 
 
def multiselect_filter(label, col, key):
    if not col or col not in df.columns:
        return []
    opts = sorted(df[col].dropna().astype(str).unique())
    return st.sidebar.multiselect(label, opts, key=key)
 
 
districts = multiselect_filter("District", district_col, "f_district")
projects = multiselect_filter("Project", project_col, "f_project")
circuits = multiselect_filter("Circuit", circuit_col, "f_circuit")
 
f = df.copy()
if date_range and isinstance(date_range, tuple) and len(date_range) == 2:
    start, end = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    f = f[f["_date"].between(start, end)]
if districts:
    f = f[f[district_col].astype(str).isin(districts)]
if projects:
    f = f[f[project_col].astype(str).isin(projects)]
if circuits:
    f = f[f[circuit_col].astype(str).isin(circuits)]
 
st.sidebar.divider()
st.sidebar.metric("Rows after filters", f"{len(f):,}", delta=f"of {len(df):,} total")
 
tab_overview, tab_jobs, tab_items, tab_totals = st.tabs(
    ["📊 Overview", "🗂️ Jobs", "📦 Mapped Items", "💰 Totals"]
)
 
# ---- Overview: CV7_recover over time ----
with tab_overview:
    st.subheader("CV7_recover — count over time")
    recover_keys = {normalize_item(k) for k in CV7_recover}
    recover_df = f[f["_item_norm"].isin(recover_keys)]
 
    if recover_df.empty or recover_df["_date"].isna().all():
        st.caption("No CV7_recover records (with a date) for the current filters.")
    else:
        granularity = st.radio("Group by", ["Day", "Week", "Month"], index=2, horizontal=True)
        freq = {"Day": "D", "Week": "W", "Month": "MS"}[granularity]
        trend = (
            recover_df.dropna(subset=["_date"])
            .set_index("_date")
            .resample(freq)["_qsub_adj"]
            .sum()
            .reset_index()
            .rename(columns={"_date": "Date", "_qsub_adj": "Count"})
        )
        fig = px.bar(trend, x="Date", y="Count", text="Count")
        fig.update_traces(marker_color="#2563eb")
        fig.update_layout(margin=dict(t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)
 
    st.subheader("All pole categories")
    pole_keys = set()
    for mapping in POLE_CATEGORIES.values():
        pole_keys |= {normalize_item(k) for k in mapping}
    pole_df = f[f["_item_norm"].isin(pole_keys)]
    pole_summary = (
        pole_df.groupby("_mapped_category")["_qsub_adj"].sum().reset_index()
        .rename(columns={"_mapped_category": "Pole type", "_qsub_adj": "Count"})
    )
    if not pole_summary.empty:
        fig2 = px.bar(pole_summary, x="Pole type", y="Count", text="Count", color="Pole type")
        fig2.update_layout(showlegend=False, margin=dict(t=10, b=10))
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.caption("No pole records for the current filters.")
 
# ---- Jobs tab ----
with tab_jobs:
    st.subheader("Jobs (District → Job → Circuit → PID)")
    job_col = cols["job_col"]
    if job_col and job_col in f.columns:
        job_table = pd.DataFrame({
            "District": f[district_col],
            "Job": f["_job_clean"],
            "Circuit": f[circuit_col],
            "PID": f[cols["pid_col"]] if cols["pid_col"] in f.columns else "",
        })
        st.caption(f"{len(job_table):,} rows under the current filters")
        collapse = st.checkbox("Collapse near-duplicate job names (≥65% similar)", value=False, key="jobs_dedupe")
        if collapse:
            job_table["_dup"] = dedupe_jobs(job_table["Job"].tolist())
            job_table = job_table[~job_table["_dup"]].drop(columns="_dup")
            st.caption(f"{len(job_table):,} rows after collapsing")
        st.dataframe(job_table, height=420, use_container_width=True, hide_index=True)
    else:
        st.caption("No job column selected in the sidebar mapping.")
 
# ---- Mapped items tab: image-led groups, then the rest as a card grid ----
with tab_items:
    st.subheader("Mapped items")
 
    cv7_poles = cv7_dedupe_poles(f)
    all_card_data = []  # accumulated in display order, feeds the detail selectbox below
    grouped_cat_names = {c for group in CARD_GROUPS for c in group["categories"]}
 
    for group in CARD_GROUPS:
        title = os.path.splitext(os.path.basename(group["image"]))[0]
        img_l, img_c, img_r = st.columns([1, 1, 1])
        with img_c:
            st.markdown(f"<h3 style='text-align:center; margin-bottom:0.3rem;'>{title}</h3>", unsafe_allow_html=True)
            if os.path.exists(group["image"]):
                st.image(group["image"], width=300)
            else:
                st.caption(f"⚠️ Image not found: {group['image']}")
 
        group_cards = []
        for cat_name in group["categories"]:
            mapping = ALL_CATEGORIES.get(cat_name)
            if mapping is None:
                continue
            card = build_card(f, cat_name, mapping, cv7_poles)
            if card:
                group_cards.append(card)
                all_card_data.append(card)
 
        if group_cards:
            row_cols = st.columns(len(group_cards))
            for slot, (cat_name, total_qty, _sub) in zip(row_cols, group_cards):
                render_metric(slot, cat_name, total_qty)
        else:
            st.caption("No records for this group under the current filters.")
 
        st.divider()
 
    # Everything not already shown in an image group, same grid as before
    remaining_cat_names = [c for c in ALL_CATEGORIES if c not in grouped_cat_names]
    remaining_cards = []
    for cat_name in remaining_cat_names:
        card = build_card(f, cat_name, ALL_CATEGORIES[cat_name], cv7_poles)
        if card:
            remaining_cards.append(card)
            all_card_data.append(card)
 
    if remaining_cards:
        st.markdown("**Other items**")
        n_cols = 4
        rows = [remaining_cards[i:i + n_cols] for i in range(0, len(remaining_cards), n_cols)]
        for row in rows:
            row_cols = st.columns(n_cols)
            for slot, (cat_name, total_qty, _sub) in zip(row_cols, row):
                render_metric(slot, cat_name, total_qty)
 
    if not all_card_data:
        st.caption("No mapped items for the current filters.")
    else:
        st.divider()
        chosen = st.selectbox("View details for", [c[0] for c in all_card_data])
        _, _, sub = next(c for c in all_card_data if c[0] == chosen)
        detail = pd.DataFrame({
            "District": sub[district_col],
            "Job": sub["_job_clean"],
            "Circuit": sub[circuit_col],
            "enid": sub[cols["pole_col"]] if cols["pole_col"] in sub.columns else "",
        })
        st.dataframe(detail, height=320, use_container_width=True, hide_index=True)
 
# ---- Totals tab ----
with tab_totals:
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
 
        st.subheader("Jobs where total ≠ original")
        variance_table = pd.DataFrame({
            "District": variance_rows[district_col],
            "Job": variance_rows["_job_clean"],
            "Circuit": variance_rows[circuit_col],
            "Difference (£)": variance_rows["_row_variance"],
        })
        st.caption(f"{len(variance_table):,} rows with a variance under the current filters")
        st.dataframe(
            variance_table.sort_values("Difference (£)", key=abs, ascending=False),
            height=320, use_container_width=True, hide_index=True,
        )
 
        st.subheader("Difference by Job")
        by_job = (
            variance_rows.assign(Job=variance_rows["_job_clean"])
            .groupby("Job")["_row_variance"].sum()
            .reset_index()
            .rename(columns={"_row_variance": "Difference (£)"})
            .sort_values("Difference (£)", key=abs, ascending=False)
        )
        st.dataframe(by_job, height=280, use_container_width=True, hide_index=True)
 
        st.subheader("Difference by Project")
        if project_col in variance_rows.columns:
            by_project = (
                variance_rows.groupby(project_col)["_row_variance"].sum()
                .reset_index()
                .rename(columns={project_col: "Project", "_row_variance": "Difference (£)"})
                .sort_values("Difference (£)", key=abs, ascending=False)
            )
            st.dataframe(by_project, height=280, use_container_width=True, hide_index=True)
