import datetime as dt
 
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
 
# --------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------
 
st.set_page_config(
    page_title="Master Control — Job Costing Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)
 
COLUMNS_NEEDED = [
    "job", "plan1", "done", "datetouse", "invoice date", "total", "orig",
    "district", "project", "project manager", "team", "team lider", "pid",
]
 
COLOR_GREEN = "#4FBE84"
COLOR_RED = "#E2604F"
COLOR_AMBER = "#E3A64A"
COLOR_TEAL = "#4FA6C9"
COLOR_PURPLE = "#8C7AE6"
COLOR_MUTED = "#4A5568"
 
CATEGORY_PALETTE = [
    "#4FA6C9", "#E3A64A", "#8C7AE6", "#4FBE84", "#E2604F",
    "#5DD5D0", "#C97ED9", "#B8C24A", "#E087A8", "#6B8CE6",
]
 
PLOTLY_DARK = dict(
    paper_bgcolor="#131922",
    plot_bgcolor="#131922",
    font=dict(color="#E7ECF2", family="IBM Plex Mono, monospace", size=12),
    legend=dict(bgcolor="rgba(0,0,0,0)"),
    margin=dict(l=10, r=10, t=30, b=10),
)
 
 
# --------------------------------------------------------------------------
# Data loading
# --------------------------------------------------------------------------
 
@st.cache_data(show_spinner="Loading Master Control data…")
def load_data(file) -> pd.DataFrame:
    df = pd.read_parquet(file, columns=COLUMNS_NEEDED)
 
    # total / orig sometimes land as text — coerce to numeric
    df["total"] = pd.to_numeric(df["total"], errors="coerce")
    df["orig"] = pd.to_numeric(df["orig"], errors="coerce")
 
    for c in ["plan1", "done", "datetouse", "invoice date"]:
        df[c] = pd.to_datetime(df[c], errors="coerce")
 
    for c in ["district", "project", "project manager", "team", "team lider", "pid", "job"]:
        df[c] = df[c].astype("string").str.strip()
        df[c] = df[c].replace("", pd.NA)
 
    job_lower = df["job"].str.lower()
    df["flag"] = "other"
    df.loc[job_lower.str.startswith(("m -", "m-"), na=False), "flag"] = "material"
    df.loc[job_lower.str.startswith(("c -", "c-"), na=False), "flag"] = "construction"
    df = df[df["flag"] != "other"].copy()
 
    # datetouse sometimes carries a placeholder ~1900 date for "no date" — treat as missing
    df["datetouse"] = df["datetouse"].where(df["datetouse"].dt.year > 1901)
 
    return df
 
 
# --------------------------------------------------------------------------
# Formatting helpers
# --------------------------------------------------------------------------
 
def fmt_money(v: float) -> str:
    if pd.isna(v):
        v = 0.0
    sign = "-" if v < 0 else ""
    return f"{sign}£{abs(v):,.2f}"
 
 
def fmt_money_short(v: float) -> str:
    if pd.isna(v):
        v = 0.0
    sign = "-" if v < 0 else ""
    a = abs(v)
    if a >= 1_000_000:
        return f"{sign}£{a/1_000_000:.1f}m"
    if a >= 1_000:
        return f"{sign}£{a/1_000:.1f}k"
    return f"{sign}£{a:.0f}"
 
 
def auto_granularity(date_from: dt.date, date_to: dt.date) -> str:
    days = (date_to - date_from).days
    if days <= 31:
        return "Day"
    if days <= 180:
        return "Week"
    if days <= 900:
        return "Month"
    return "Year"
 
 
def bucket_series(dates: pd.Series, granularity: str) -> pd.Series:
    if granularity == "Day":
        return dates.dt.to_period("D").dt.start_time
    if granularity == "Week":
        return dates.dt.to_period("W-MON").dt.start_time
    if granularity == "Month":
        return dates.dt.to_period("M").dt.start_time
    return dates.dt.to_period("Y").dt.start_time
 
 
def bucket_label(ts: pd.Timestamp, granularity: str) -> str:
    if granularity == "Day":
        return ts.strftime("%-d %b") if hasattr(ts, "strftime") else str(ts)
    if granularity == "Week":
        return "w/c " + ts.strftime("%-d %b")
    if granularity == "Month":
        return ts.strftime("%b %Y")
    return ts.strftime("%Y")
 
 
def metric_row(items):
    """items: list of (label, value_str, color_or_None)"""
    cols = st.columns(len(items))
    for col, (label, value, color) in zip(cols, items):
        with col:
            style = f"color:{color};" if color else ""
            st.markdown(
                f"""
                <div style="border:1px solid #2A3442;border-radius:6px;padding:10px 14px;background:#0B0F14;">
                  <div style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.08em;
                              text-transform:uppercase;color:#7C8AA0;margin-bottom:4px;">{label}</div>
                  <div style="font-family:'IBM Plex Mono',monospace;font-size:18px;font-weight:600;{style}">{value}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
 
 
# --------------------------------------------------------------------------
# Sidebar — data source + filters
# --------------------------------------------------------------------------
 
st.sidebar.title("Master Control")
st.sidebar.caption("Job costing & progress — construction / materials")
 
st.sidebar.markdown("### Data source")
uploaded = st.sidebar.file_uploader(
    "Upload Master Control parquet",
    type=["parquet"],
    help="Drop in the latest export each time — nothing is stored between sessions.",
)
 
if uploaded is None:
    st.title("Master Control — Job Costing Dashboard")
    st.info("⬅️ Upload the latest Master Control parquet file in the sidebar to get started.")
    st.stop()
 
df = load_data(uploaded)
 
date_min = df["datetouse"].min()
date_max = df["datetouse"].max()
 
st.sidebar.markdown("### Filters")
 
date_range = st.sidebar.date_input(
    "Date range",
    value=(date_min.date(), date_max.date()),
    min_value=date_min.date(),
    max_value=date_max.date(),
)
if isinstance(date_range, tuple) and len(date_range) == 2:
    date_from, date_to = date_range
else:
    date_from, date_to = date_min.date(), date_max.date()
 
districts = st.sidebar.multiselect("District", sorted(df["district"].dropna().unique()))
projects = st.sidebar.multiselect("Project", sorted(df["project"].dropna().unique()))
pids = st.sidebar.multiselect("PID", sorted(df["pid"].dropna().unique()))
pms = st.sidebar.multiselect("Project Manager", sorted(df["project manager"].dropna().unique()))
 
gran_choice = st.sidebar.selectbox("Date grouping", ["Auto", "Day", "Week", "Month", "Year"])
 
if st.sidebar.button("Reset filters"):
    st.rerun()
 
# --------------------------------------------------------------------------
# Apply filters
# --------------------------------------------------------------------------
 
mask = pd.Series(True, index=df.index)
if districts:
    mask &= df["district"].isin(districts)
if projects:
    mask &= df["project"].isin(projects)
if pids:
    mask &= df["pid"].isin(pids)
if pms:
    mask &= df["project manager"].isin(pms)
 
fdf = df[mask].copy()
 
date_mask = fdf["datetouse"].notna() & (fdf["datetouse"].dt.date >= date_from) & (fdf["datetouse"].dt.date <= date_to)
fdf_dated = fdf[date_mask].copy()
 
granularity = auto_granularity(date_from, date_to) if gran_choice == "Auto" else gran_choice
 
st.sidebar.markdown("---")
st.sidebar.markdown(
    f"<span style='font-family:monospace;font-size:12px;color:#7C8AA0;'>"
    f"RECORDS IN VIEW: <b style='color:#E3A64A;'>{len(fdf_dated):,}</b></span>",
    unsafe_allow_html=True,
)
 
# --------------------------------------------------------------------------
# Header
# --------------------------------------------------------------------------
 
st.title("Master Control — Job Costing Dashboard")
st.caption(
    f"Source: Master parquet · {date_min.date()} → {date_max.date()} · "
    f"grouping: {granularity}{' (auto)' if gran_choice == 'Auto' else ''}"
)
 
tab_trend, tab_pid = st.tabs(["📈 Trends", "📊 PID Breakdown"])
 
# --------------------------------------------------------------------------
# TAB 1 — Trends
# --------------------------------------------------------------------------
 
with tab_trend:
    st.subheader("Panel 01 — Total by Date (Construction)")
    st.caption("Bar height = construction total · colour = variance vs original · label = material value")
 
    construction = fdf_dated[fdf_dated["flag"] == "construction"].copy()
    material = fdf_dated[fdf_dated["flag"] == "material"].copy()
 
    if construction.empty and material.empty:
        st.info("No records under these filters. Widen the date range or clear a filter.")
    else:
        construction["bucket"] = bucket_series(construction["datetouse"], granularity)
        material["bucket"] = bucket_series(material["datetouse"], granularity)
 
        c_agg = construction.groupby("bucket", as_index=False).agg(total=("total", "sum"), orig=("orig", "sum"))
        m_agg = material.groupby("bucket", as_index=False).agg(material=("total", "sum"))
 
        chart_df = pd.merge(c_agg, m_agg, on="bucket", how="outer").fillna(0)
        chart_df = chart_df.sort_values("bucket")
        chart_df["variance"] = chart_df["total"] - chart_df["orig"]
        chart_df["color"] = chart_df["variance"].apply(lambda v: COLOR_GREEN if v >= 0 else COLOR_RED)
        chart_df["label"] = chart_df["bucket"].apply(lambda ts: bucket_label(ts, granularity))
 
        metric_row([
            ("Total (construction)", fmt_money(chart_df["total"].sum()), None),
            ("Original", fmt_money(chart_df["orig"].sum()), None),
            ("Variance", fmt_money(chart_df["variance"].sum()),
             COLOR_GREEN if chart_df["variance"].sum() >= 0 else COLOR_RED),
            ("Materials", fmt_money(chart_df["material"].sum()), COLOR_AMBER),
        ])
 
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=chart_df["label"], y=chart_df["total"],
            marker_color=chart_df["color"],
            text=[fmt_money_short(v) if v else "" for v in chart_df["material"]],
            textposition="outside",
            textfont=dict(color=COLOR_AMBER, size=11),
            name="Total",
            customdata=chart_df[["orig", "variance", "material"]],
            hovertemplate=(
                "<b>%{x}</b><br>Total: £%{y:,.2f}<br>Original: £%{customdata[0]:,.2f}"
                "<br>Variance: £%{customdata[1]:,.2f}<br>Materials: £%{customdata[2]:,.2f}<extra></extra>"
            ),
            showlegend=False,
        ))
        # dummy traces purely to build the 3-colour legend
        fig.add_trace(go.Bar(x=[None], y=[None], marker_color=COLOR_AMBER, name="Materials"))
        fig.add_trace(go.Bar(x=[None], y=[None], marker_color=COLOR_RED, name="Negative variation"))
        fig.add_trace(go.Bar(x=[None], y=[None], marker_color=COLOR_GREEN, name="Positive variation"))
 
        fig.update_layout(
            **PLOTLY_DARK,
            height=380,
            yaxis=dict(tickprefix="£", gridcolor="#1E2733", zeroline=False),
            xaxis=dict(gridcolor="#1E2733"),
            legend=dict(orientation="h", y=-0.18, bgcolor="rgba(0,0,0,0)"),
            bargap=0.25,
        )
        st.plotly_chart(fig, use_container_width=True)
 
    st.markdown("---")
    st.subheader("Panel 02 — Total by Team Leader (Construction)")
    st.caption("Stacked by team leader · same filters & date grouping as above")
 
    if construction.empty:
        st.info("No records under these filters.")
    else:
        totals_by_lider = (
            construction.assign(lider=construction["team lider"].fillna("Unassigned"))
            .groupby("lider")["total"].sum()
            .sort_values(ascending=False)
        )
        top_liders = list(totals_by_lider.index[:9])
        has_other = len(totals_by_lider) > 9
 
        construction["lider_grp"] = construction["team lider"].fillna("Unassigned")
        if has_other:
            construction["lider_grp"] = construction["lider_grp"].where(
                construction["lider_grp"].isin(top_liders), "Other"
            )
 
        lider_agg = construction.groupby(["bucket", "lider_grp"], as_index=False)["total"].sum()
        lider_order = top_liders + (["Other"] if has_other else [])
        color_map = {l: (COLOR_MUTED if l == "Other" else CATEGORY_PALETTE[i % len(CATEGORY_PALETTE)])
                     for i, l in enumerate(lider_order)}
 
        var_sum = construction["total"].sum() - construction["orig"].sum()
        metric_row([
            ("Total (construction)", fmt_money(construction["total"].sum()), None),
            ("Original", fmt_money(construction["orig"].sum()), None),
            ("Variance", fmt_money(var_sum), COLOR_GREEN if var_sum >= 0 else COLOR_RED),
        ])
 
        fig2 = go.Figure()
        for l in lider_order:
            sub = lider_agg[lider_agg["lider_grp"] == l].sort_values("bucket")
            fig2.add_trace(go.Bar(
                x=[bucket_label(ts, granularity) for ts in sub["bucket"]],
                y=sub["total"],
                name=l,
                marker_color=color_map[l],
                hovertemplate=f"<b>{l}</b><br>%{{x}}<br>£%{{y:,.2f}}<extra></extra>",
            ))
        fig2.update_layout(
            **PLOTLY_DARK,
            height=360,
            barmode="stack",
            yaxis=dict(tickprefix="£", gridcolor="#1E2733", zeroline=False),
            xaxis=dict(gridcolor="#1E2733", categoryorder="category ascending"),
            legend=dict(orientation="h", y=-0.22, bgcolor="rgba(0,0,0,0)"),
            bargap=0.25,
        )
        st.plotly_chart(fig2, use_container_width=True)
 
# --------------------------------------------------------------------------
# TAB 2 — PID Breakdown
# --------------------------------------------------------------------------
 
with tab_pid:
    st.subheader("Panel 03 — PID Breakdown")
    st.caption("Grouped by district → project → PID → project manager → most common job line")
 
    mc_mode = st.radio("Show", ["All", "Construction", "Material"], horizontal=True)
 
    pdf = fdf.dropna(subset=["pid"]).copy()
    if mc_mode != "All":
        pdf = pdf[pdf["flag"] == mc_mode.lower()]
 
    if pdf.empty:
        st.info("No PIDs under these filters.")
    else:
        pdf["has_plan"] = pdf["plan1"].notna()
        pdf["has_done"] = pdf["done"].notna()
        pdf["has_inv"] = pdf["invoice date"].notna()
 
        group_cols = ["district", "project", "pid", "project manager"]
 
        def _job_mode(s):
            m = s.mode()
            return m.iloc[0] if not m.empty else None
 
        agg = pdf.groupby(group_cols, dropna=False).apply(
            lambda g: pd.Series({
                "total": g["total"].sum(),
                "planned": g.loc[g["has_plan"], "total"].sum(),
                "done": g.loc[g["has_done"], "total"].sum(),
                "invoiced": g.loc[g["has_inv"], "total"].sum(),
                "job": _job_mode(g["job"]),
            })
        ).reset_index()
 
        agg = agg.sort_values(["district", "project", "pid"])
        agg["row_label"] = (
            agg["district"].fillna("—") + " · " + agg["project"].fillna("—") + " · " +
            agg["pid"].fillna("—") + " · " + agg["project manager"].fillna("—") + " · " +
            agg["job"].fillna("").str.slice(0, 45)
        )
 
        metric_row([
            ("Total", fmt_money(agg["total"].sum()), COLOR_TEAL),
            ("Planned", fmt_money(agg["planned"].sum()), COLOR_AMBER),
            ("Done", fmt_money(agg["done"].sum()), COLOR_GREEN),
            ("Invoiced", fmt_money(agg["invoiced"].sum()), COLOR_PURPLE),
        ])
        st.caption(f"{len(agg)} PID(s) in view")
 
        row_h = 26
        fig3 = go.Figure()
        y = agg["row_label"][::-1]
        fig3.add_trace(go.Bar(y=y, x=agg["total"][::-1], name="Total", orientation="h",
                               marker_color=COLOR_TEAL,
                               hovertemplate="Total: £%{x:,.2f}<extra></extra>"))
        fig3.add_trace(go.Bar(y=y, x=agg["planned"][::-1], name="Planned", orientation="h",
                               marker_color=COLOR_AMBER,
                               hovertemplate="Planned: £%{x:,.2f}<extra></extra>"))
        fig3.add_trace(go.Bar(y=y, x=agg["done"][::-1], name="Done", orientation="h",
                               marker_color=COLOR_GREEN,
                               hovertemplate="Done: £%{x:,.2f}<extra></extra>"))
        fig3.add_trace(go.Bar(y=y, x=agg["invoiced"][::-1], name="Invoiced", orientation="h",
                               marker_color=COLOR_PURPLE,
                               hovertemplate="Invoiced: £%{x:,.2f}<extra></extra>"))
 
        fig3.update_layout(
            **PLOTLY_DARK,
            height=max(320, row_h * len(agg) * 4 * 0.28),
            barmode="group",
            xaxis=dict(tickprefix="£", gridcolor="#1E2733", zeroline=False),
            yaxis=dict(gridcolor="#1E2733", automargin=True, tickfont=dict(size=10)),
            legend=dict(orientation="h", y=1.05, bgcolor="rgba(0,0,0,0)"),
            bargap=0.2,
            bargroupgap=0.1,
        )
        st.plotly_chart(fig3, use_container_width=True)
 
st.markdown("---")
st.caption(f"Master Control Dashboard — data as of {date_max.date()}")
