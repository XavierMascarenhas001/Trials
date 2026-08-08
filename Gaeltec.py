Claude finished the response

In my tab jobs I want to read information from one excel.
"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\1 - Outages Programme\High-level_planning_2026.xlsx"

I want the columns District Outage Date Weekday, Shceme, Outage #, Circuit, PID, SPEN PM, POI.

Should we convert the document to parquet and download into the Dashboard for reading or can we read directly from the dashboard to the path? I want efficiency, dont want the file to be to slow
Weighed direct network reading versus parquet conversion for optimal performance

Direct read is the better call here — you don't need a parquet conversion step for this file.

Why not convert to parquet:

    Converting still requires reading the xlsx at least once, so you're not avoiding the "slow" part — you're just moving it to a separate script.
    You'd now need a scheduler/cron job to keep the parquet copy in sync whenever the planning workbook changes, which is more moving parts to maintain for a file that's presumably updated fairly often (it's a live planning tracker).
    Your existing read_file already handles parquet for the main upload — mixing a second, auto-refreshed parquet pipeline into the same app adds complexity without a real payoff at this file size.

Why direct read is fine — with one important caveat:
The actual slowness with Excel comes from two things: (1) openpyxl parsing the whole workbook including all 8 sheets and formatting, and (2) Streamlit re-running your whole script on every widget interaction, which would re-hit the network share every time if you don't cache it properly.

Fix both by:

    Reading only the one sheet (2026) and only the columns you need, skipping the header junk rows.
    Caching on the file's last-modified time, not just st.cache_data alone — so it re-reads only when the source file actually changes, not on every click.

python

OUTAGE_PATH = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\1 - Outages Programme\High-level_planning_2026.xlsx"

@st.cache_data(show_spinner="Loading outage programme...")
def load_outage_programme(path: str, mtime: float) -> pd.DataFrame:
    # mtime is only here to bust the cache when the source file changes
    df = pd.read_excel(
        path,
        sheet_name="2026",
        header=6,                     # row 7 in Excel is the real header row
        usecols="A,B,C,E,F,G,L,M,N",  # District, Outage Date, Weekday, Scheme, Outage #, Circuit, PID, SPEN PM, POI
        engine="openpyxl",
    )
    df.columns = [
        "District", "Outage Date", "Weekday", "Scheme",
        "Outage #", "Circuit", "PID", "SPEN PM", "POI",
    ]
    df = df.dropna(how="all")
    return df

def get_outage_programme(path: str) -> pd.DataFrame | None:
    if not os.path.exists(path):
        st.sidebar.warning(f"Can't reach outage programme file:\n{path}")
        return None
    mtime = os.path.getmtime(path)
    return load_outage_programme(path, mtime)

Then in the Jobs tab:
python

with tab_jobs:
    st.subheader("Jobs (District → Job → Circuit → PID)")
    ...  # your existing job_table code

    st.divider()
    st.subheader("Outages Programme 2026")
    outage_df = get_outage_programme(OUTAGE_PATH)
    if outage_df is not None:
        st.caption(f"{len(outage_df):,} rows from High-level_planning_2026.xlsx")
        st.dataframe(outage_df, height=420, use_container_width=True, hide_index=True)

Two things to check before this works in practice:

    The machine running Streamlit must be able to reach that UNC path the same way your own PC does (same domain/VPN, and the account running Streamlit needs read access to the share). If you ever deploy this off the local network (e.g. Streamlit Cloud), direct UNC access won't be possible at all — that's the one scenario where you'd have to fall back to a synced local copy or parquet snapshot.
    If several people load the dashboard at once, st.cache_data is shared across sessions by default, so the network file only actually gets hit once per change, not once per user — which is the main efficiency win here.

If this workbook later grows to tens of thousands of rows or gets read very frequently by many users, revisit parquet then — but at the size shown in your upload, direct + mtime-cached read will be plenty fast.

Also the excel information is in a sheet named 2026 and the header is in row 7

Good — that matches what I used. In pandas, header is zero-indexed, so Excel row 7 corresponds to header=6 in read_excel, which is exactly what the snippet already has:
python

df = pd.read_excel(
    path,
    sheet_name="2026",
    header=6,   # Excel row 7 → 0-indexed row 6
    usecols="A,B,C,E,F,G,L,M,N",
    engine="openpyxl",
)

So no change needed there — the code I gave you should line up correctly with your sheet. If you want, paste in a couple of sample rows from the actual "2026" sheet (or re-confirm the exact column letters for District/Outage Date/Weekday/Scheme/Outage #/Circuit/PID/SPEN PM/POI) and I'll double check the usecols mapping matches exactly, since a shifted column would silently pull the wrong data rather than error out.

give me full code
Ran 3 commands, created a file, read a file

Everything else is unchanged from your original app — the only additions are OUTAGE_PATH, load_outage_programme, get_outage_programme near the top, and the new "Outages Programme 2026" block at the end of the Jobs tab (with quick filters for District, SPEN PM, and outage date). It only re-hits the network share when the file's modified time changes, otherwise it serves from Streamlit's cache.
Network job tracker
PY 
Want to be notified when Claude responds?


Claude is AI and can make mistakes. Please double-check responses.
Network job tracker · PY
import os
import re
import difflib
from datetime import datetime
 
import pandas as pd
import streamlit as st
import plotly.express as px
 
st.set_page_config(page_title="Network Job Tracker", layout="wide")
 
# ============================================================
# OUTAGES PROGRAMME (external network file, read directly)
# ============================================================
OUTAGE_PATH = r"\\gaeltec-gl\Gaeltec_Network\62.OHLT.UK\03.SPEN\21.Planning\1 - Outages Programme\High-level_planning_2026.xlsx"
 
 
@st.cache_data(show_spinner="Loading outage programme...")
def load_outage_programme(path: str, mtime: float) -> pd.DataFrame:
    """mtime is only passed in to bust the Streamlit cache when the source
    file on the network share actually changes - it is not used otherwise."""
    df = pd.read_excel(
        path,
        sheet_name="2026",
        header=6,                     # Excel row 7 is the real header row (0-indexed = 6)
        usecols="A,B,C,E,F,G,L,M,N",  # District, Outage Date, Weekday, Scheme, Outage #, Circuit, PID, SPEN PM, POI
        engine="openpyxl",
    )
    df.columns = [
        "District", "Outage Date", "Weekday", "Scheme",
        "Outage #", "Circuit", "PID", "SPEN PM", "POI",
    ]
    df["Outage Date"] = pd.to_datetime(df["Outage Date"], errors="coerce")
    df = df.dropna(how="all")
    return df
 
 
def get_outage_programme(path: str):
    if not os.path.exists(path):
        st.sidebar.warning(f"Can't reach outage programme file:\n{path}")
        return None
    mtime = os.path.getmtime(path)
    return load_outage_programme(path, mtime)
 
 
# ============================================================
# YOUR MAPPING DICTIONARIES
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
 
# Friendly labels for the pole chart/cards, in the "Display (technical_name)"
# style so it's obvious which export-tool category each bar corresponds to.
POLE_DISPLAY_NAMES = {
    "CV7_erect": "CV7 (CV7_erect)",
    "CV7_erect_H": "CV7 H Pole (CV7_erect_H)",
    "CV7_erect_lv": "CV7 LV Pole (CV7_erect_lv)",
    "CV7_recover": "CV7 Recover (CV7_recover)",
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
 
# Switch is split into three sub-types for display instead of one lump
# card. Each key is the display name, each value is the list of raw
# descriptions (as written in the Switch dict above) that count toward it.
SWITCH_SUBTYPES = {
    "Noja": ["Noja"],
    "Soule": ["11kV PMSW (Soule)"],
    "ABSW": [
        "11kv ABSW Hookstick Standard",
        "11kv ABSW Hookstick Spring loaded mech",
        "33kv ABSW Hookstick Dependant",
    ],
}
 
# ============================================================
# IMAGE GROUPS for the Mapped Items tab
# Images live in an "Images" folder alongside this script (same repo path).
# "image": None means the group is shown without an image (no warning).
# Any category not listed in a group's "categories" (or covered by
# "subtypes") falls through to the "Other items" section at the end.
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
        "title": "Conductor",
        "image": os.path.join(IMAGE_DIR, "Cable.png"),
        "categories": [
            "CV7_OHL_CONDUCTOR_instal",
            "CV7_OHL_CONDUCTOR_recover",
            "CV7_OHL_CONDUCTOR_LV_instal",
            "CV7_OHL_CONDUCTOR_LV_recover",
        ],
    },
    {
        "title": "Switch gear",
        "image": os.path.join(IMAGE_DIR, "Switchgear.png"),
        "subtypes": SWITCH_SUBTYPES,
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
 
 
def show_total_banner(label, value_str):
    """A large, centered KPI number - meant to sit directly under a chart's
    subheader so the total is the first thing seen after the title."""
    st.markdown(
        f"""
        <div style="text-align:center; padding: 4px 0 18px 0;">
            <div style="font-size:2.6rem; font-weight:800; color:#1e3a8a; line-height:1.15;">{value_str}</div>
            <div style="font-size:0.9rem; color:#475569; text-transform:uppercase; letter-spacing:0.05em;">{label}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
 
 
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
    """Returns (cat_name, total_qty, sub) for a single category, filtered
    strictly by that category's OWN item keys - never by the shared
    "_mapped_category" label, since several categories map to the same
    label (e.g. CV7_erect and CV7_erect_H both -> "CV7 HV pole") and
    grouping by label would merge their counts together."""
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
 
 
def build_subtype_card(frame: pd.DataFrame, subtype_name: str, descriptions: list):
    keys = {normalize_item(d) for d in descriptions}
    sub = frame[frame["_item_norm"].isin(keys)]
    if sub.empty:
        return None
    return (subtype_name, sub["_qsub_adj"].sum(), sub)
 
 
def render_metric(slot, cat_name: str, total_qty, display_name: str = None):
    label = display_name or cat_name
    if cat_name in UNIT_CONFIG:
        slot.metric(label, format_length(total_qty, UNIT_CONFIG[cat_name]))
    elif cat_name in POLE_DEDUPE_CATEGORIES:
        slot.metric(label, f"{total_qty:,.0f} poles")
    else:
        slot.metric(label, f"{total_qty:,.0f}")
 
 
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
    recover_total = recover_df["_qsub_adj"].sum()
    show_total_banner("Total CV7_recover count", f"{recover_total:,.0f}")
 
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
 
    st.subheader("All pole categories (erect only)")
    # Grouped by SOURCE category (CV7_erect / CV7_erect_H / CV7_erect_lv),
    # not by the shared "_mapped_category" label - several of these
    # categories map to the same label ("CV7 HV pole"), so grouping by
    # label would silently merge their counts. This mirrors how the
    # export tool's Summary sheet keeps each category as its own column.
    # CV7_recover is excluded here since it already has its own chart above.
    pole_rows = []
    for cat_name, mapping in POLE_CATEGORIES.items():
        if cat_name == "CV7_recover":
            continue
        keys = {normalize_item(k) for k in mapping}
        sub = f[f["_item_norm"].isin(keys)]
        if not sub.empty:
            pole_rows.append({
                "Pole type": POLE_DISPLAY_NAMES.get(cat_name, cat_name),
                "Count": sub["_qsub_adj"].sum(),
            })
    pole_summary = pd.DataFrame(pole_rows)
 
    pole_total = pole_summary["Count"].sum() if not pole_summary.empty else 0
    show_total_banner("Total poles (all categories)", f"{pole_total:,.0f}")
 
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
 
    # ---- Outages Programme (external network file, read directly + cached) ----
    st.divider()
    st.subheader("Outages Programme 2026")
 
    outage_df = get_outage_programme(OUTAGE_PATH)
    if outage_df is not None:
        st.caption(f"{len(outage_df):,} rows from High-level_planning_2026.xlsx (sheet '2026')")
 
        oc1, oc2, oc3 = st.columns(3)
        with oc1:
            outage_districts = st.multiselect(
                "District", sorted(outage_df["District"].dropna().unique()), key="outage_district"
            )
        with oc2:
            outage_pms = st.multiselect(
                "SPEN PM", sorted(outage_df["SPEN PM"].dropna().unique()), key="outage_pm"
            )
        with oc3:
            if outage_df["Outage Date"].notna().any():
                min_od, max_od = outage_df["Outage Date"].min(), outage_df["Outage Date"].max()
                outage_date_range = st.date_input(
                    "Outage date", value=(min_od.date(), max_od.date()), key="outage_date_range"
                )
            else:
                outage_date_range = None
 
        outage_f = outage_df.copy()
        if outage_districts:
            outage_f = outage_f[outage_f["District"].isin(outage_districts)]
        if outage_pms:
            outage_f = outage_f[outage_f["SPEN PM"].isin(outage_pms)]
        if outage_date_range and isinstance(outage_date_range, tuple) and len(outage_date_range) == 2:
            o_start = pd.Timestamp(outage_date_range[0])
            o_end = pd.Timestamp(outage_date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            outage_f = outage_f[outage_f["Outage Date"].between(o_start, o_end)]
 
        st.caption(f"{len(outage_f):,} rows after outage filters")
        st.dataframe(outage_f, height=420, use_container_width=True, hide_index=True)
 
# ---- Mapped items tab: image-led groups, then the rest as a card grid ----
with tab_items:
    st.subheader("Mapped items")
 
    cv7_poles = cv7_dedupe_poles(f)
    all_card_data = []  # accumulated in display order, feeds the detail selectbox below
 
    grouped_cat_names = {c for group in CARD_GROUPS for c in group.get("categories", [])}
    grouped_cat_names |= {c for group in CARD_GROUPS if "subtypes" in group for c in ["Switch"]}
 
    for group in CARD_GROUPS:
        title = group["title"]
        img_l, img_c, img_r = st.columns([1, 1, 1])
        with img_c:
            st.markdown(f"<h3 style='text-align:center; margin-bottom:0.3rem;'>{title}</h3>", unsafe_allow_html=True)
            image_path = group.get("image")
            if image_path and os.path.exists(image_path):
                st.image(image_path, width=300)
            elif image_path:
                st.caption(f"⚠️ Image not found: {image_path}")
 
        group_cards = []
 
        if "subtypes" in group:
            for subtype_name, descriptions in group["subtypes"].items():
                card = build_subtype_card(f, subtype_name, descriptions)
                if card:
                    group_cards.append(card)
                    all_card_data.append(card)
        else:
            for cat_name in group.get("categories", []):
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
                render_metric(slot, cat_name, total_qty, display_name=POLE_DISPLAY_NAMES.get(cat_name))
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
 
    if total_val is not None:
        show_total_banner("Total value (£)", f"£{total_val:,.2f}")
 
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
 

