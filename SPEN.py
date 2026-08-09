import datetime as dt
 
import numpy as np
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
 
COLOR_GREEN = "#1E9E5A"      # positive variation
COLOR_RED = "#D64545"        # negative variation
COLOR_YELLOW = "#F0B429"     # materials
COLOR_NAVY = "#1F3A5F"       # "the rest" / base (construction)
COLOR_TEAL = "#2C7DA0"       # PID tab — total/remaining
COLOR_AMBER = "#E0A02A"      # PID tab — planned
COLOR_PURPLE = "#7C5CBF"     # PID tab — invoiced
COLOR_MUTED = "#B7BFC9"
 
CATEGORY_PALETTE = [
    "#2C7DA0", "#E0A02A", "#7C5CBF", "#1E9E5A", "#D64545",
    "#2F9E9E", "#A85CB0", "#7C8C2A", "#C25C82", "#4C6FC2",
]
 
TEXT_DARK = "#1B2430"
TEXT_MUTED = "#5B6B7C"
GRID_LIGHT = "#E2E7ED"
PANEL_BG = "#FFFFFF"
 
PLOTLY_LIGHT = dict(
    paper_bgcolor=PANEL_BG,
    plot_bgcolor=PANEL_BG,
    font=dict(color=TEXT_DARK, family="IBM Plex Mono, monospace", size=13),
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
    """items: list of (label, value_str, accent_color_or_None, highlight_bool)"""
    cols = st.columns(len(items))
    for col, (label, value, color, highlight) in zip(cols, items):
        with col:
            if highlight:
                box_bg = "#EAF3EC"
                box_border = color or COLOR_NAVY
                value_size = "26px"
                label_color = TEXT_DARK
                border_style = f"border:1px solid {box_border};border-left:5px solid {box_border};"
            else:
                box_bg = "#F6F8FA"
                value_size = "18px"
                label_color = TEXT_MUTED
                border_style = "border:1px solid #DDE3E9;"
            val_style = f"color:{color};" if color else f"color:{TEXT_DARK};"
            st.markdown(
                f"""
                <div style="{border_style}border-radius:6px;padding:10px 14px;background:{box_bg};">
                  <div style="font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:.08em;
                              text-transform:uppercase;color:{label_color};margin-bottom:4px;">{label}</div>
                  <div style="font-family:'IBM Plex Mono',monospace;font-size:{value_size};font-weight:700;{val_style}">{value}</div>
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
 
tab_trend, tab_pid, tab_finance = st.tabs(["📈 Trends", "📊 PID Breakdown", "💰 Finance"])
 
# --------------------------------------------------------------------------
# TAB 1 — Trends
# --------------------------------------------------------------------------
 
with tab_trend:
    st.subheader("Panel 01 — Total by Date (Construction + Materials)")
    st.caption(
        "Each bar = base (navy) + variance vs. original (red/green) + materials (yellow), stacked"
    )
 
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
        # base = the smaller of total/orig ("the rest"); cap = the gap between them, coloured by sign
        chart_df["base"] = chart_df[["total", "orig"]].min(axis=1)
        chart_df["cap"] = chart_df["variance"].abs()
        chart_df["cap_color"] = chart_df["variance"].apply(lambda v: COLOR_GREEN if v >= 0 else COLOR_RED)
        chart_df["label"] = chart_df["bucket"].apply(lambda ts: bucket_label(ts, granularity))
 
        grand_total = chart_df["total"].sum() + chart_df["material"].sum()
        var_sum = chart_df["variance"].sum()
 
        metric_row([
            ("Total (construction + materials)", fmt_money(grand_total), COLOR_NAVY, True),
            ("Construction total", fmt_money(chart_df["total"].sum()), None, False),
            ("Original", fmt_money(chart_df["orig"].sum()), None, False),
            ("Variance", fmt_money(var_sum), COLOR_GREEN if var_sum >= 0 else COLOR_RED, False),
            ("Materials", fmt_money(chart_df["material"].sum()), COLOR_YELLOW, False),
        ])
 
        fig = go.Figure()
 
        fig.add_trace(go.Bar(
            x=chart_df["label"], y=chart_df["base"],
            marker_color=COLOR_NAVY,
            name="Base (original / the rest)",
            customdata=chart_df[["total", "orig", "variance", "material"]],
            hovertemplate="<b>%{x}</b><br>Base: £%{y:,.2f}<extra></extra>",
        ))
        fig.add_trace(go.Bar(
            x=chart_df["label"], y=chart_df["cap"],
            marker_color=chart_df["cap_color"],
            name="Variance (total − original)",
            hovertemplate="<b>%{x}</b><br>Variance: £%{y:,.2f}<extra></extra>",
            showlegend=False,
        ))
        fig.add_trace(go.Bar(
            x=chart_df["label"], y=chart_df["material"],
            marker_color=COLOR_YELLOW,
            marker_line=dict(color="#B98600", width=0.6),
            name="Materials",
            hovertemplate="<b>%{x}</b><br>Materials: £%{y:,.2f}<extra></extra>",
        ))
        # dummy traces purely to add red/green to the legend (variance trace itself is colour-mixed)
        fig.add_trace(go.Bar(x=[None], y=[None], marker_color=COLOR_GREEN, name="Positive variation"))
        fig.add_trace(go.Bar(x=[None], y=[None], marker_color=COLOR_RED, name="Negative variation"))
 
        fig.update_layout(
            **PLOTLY_LIGHT,
            height=400,
            barmode="stack",
            yaxis=dict(tickprefix="£", gridcolor=GRID_LIGHT, zeroline=False, tickfont=dict(size=12)),
            xaxis=dict(gridcolor=GRID_LIGHT, tickfont=dict(size=12)),
            legend=dict(orientation="h", y=-0.18, bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
            bargap=0.25,
            margin=dict(l=10, r=10, t=30, b=10),
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
        TOP_N = 6
        top_liders = list(totals_by_lider.index[:TOP_N])
        has_other = len(totals_by_lider) > TOP_N
 
        construction["lider_grp"] = construction["team lider"].fillna("Unassigned")
        if has_other:
            construction["lider_grp"] = construction["lider_grp"].where(
                construction["lider_grp"].isin(top_liders), "Other"
            )
 
        lider_agg = construction.groupby(["bucket", "lider_grp"], as_index=False)["total"].sum()
        lider_order = top_liders + (["Other"] if has_other else [])
        color_map = {l: (COLOR_MUTED if l == "Other" else CATEGORY_PALETTE[i % len(CATEGORY_PALETTE)])
                     for i, l in enumerate(lider_order)}
 
        def _short_name(name, n=16):
            return name if len(name) <= n else name[:n - 1] + "…"
 
        short_label = {l: (l if l == "Other" else _short_name(l)) for l in lider_order}
 
        var_sum = construction["total"].sum() - construction["orig"].sum()
        metric_row([
            ("Total (construction)", fmt_money(construction["total"].sum()), COLOR_NAVY, True),
            ("Original", fmt_money(construction["orig"].sum()), None, False),
            ("Variance", fmt_money(var_sum), COLOR_GREEN if var_sum >= 0 else COLOR_RED, False),
        ])
 
        fig2 = go.Figure()
        for l in lider_order:
            sub = lider_agg[lider_agg["lider_grp"] == l].sort_values("bucket")
            fig2.add_trace(go.Bar(
                x=[bucket_label(ts, granularity) for ts in sub["bucket"]],
                y=sub["total"],
                name=short_label[l],
                marker_color=color_map[l],
                hovertemplate=f"<b>{l}</b><br>%{{x}}<br>£%{{y:,.2f}}<extra></extra>",
            ))
        fig2.update_layout(
            **PLOTLY_LIGHT,
            height=380,
            barmode="stack",
            yaxis=dict(tickprefix="£", gridcolor=GRID_LIGHT, zeroline=False, tickfont=dict(size=12)),
            xaxis=dict(gridcolor=GRID_LIGHT, categoryorder="category ascending", tickfont=dict(size=12)),
            legend=dict(
                orientation="h", y=-0.22, x=0.5, xanchor="center",
                bgcolor="rgba(0,0,0,0)", font=dict(size=11), tracegroupgap=4,
            ),
            bargap=0.25,
            margin=dict(l=10, r=10, t=30, b=10),
        )
        st.plotly_chart(fig2, use_container_width=True)
        if has_other:
            st.caption(
                "Top " + str(TOP_N) + " team leaders shown individually — the rest are grouped under **Other**."
            )
 
# --------------------------------------------------------------------------
# TAB 2 — PID Breakdown
# --------------------------------------------------------------------------
 
with tab_pid:
    st.subheader("Panel 03 — PID Breakdown")
    st.caption("District → Project → PID, one bar per PID · ordered by total value, highest first")
 
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
 
        # mutually exclusive progress stage per row, so the four segments sum exactly to "total"
        pdf["stage"] = np.select(
            [pdf["has_inv"], pdf["has_done"], pdf["has_plan"]],
            ["invoiced", "done", "planned"],
            default="remaining",
        )
 
        group_cols = ["district", "project", "pid", "project manager"]
 
        def _job_mode(s):
            m = s.mode()
            return m.iloc[0] if not m.empty else None
 
        base = pdf.groupby(group_cols, dropna=False).agg(
            total=("total", "sum"), job=("job", _job_mode)
        ).reset_index()
 
        stage_sums = pdf.groupby(group_cols + ["stage"], dropna=False)["total"].sum().unstack("stage", fill_value=0)
        for s in ["remaining", "planned", "done", "invoiced"]:
            if s not in stage_sums.columns:
                stage_sums[s] = 0.0
        stage_sums = stage_sums.reset_index()
 
        agg = base.merge(stage_sums, on=group_cols, how="left").fillna(0)
 
        # order: district money desc -> project money desc (within district) -> pid money desc (within project)
        agg["district_total"] = agg.groupby("district")["total"].transform("sum")
        agg["project_total"] = agg.groupby(["district", "project"])["total"].transform("sum")
        agg = agg.sort_values(
            ["district_total", "project_total", "total"], ascending=[False, False, False]
        )
 
        def _short(s, n):
            s = s or ""
            return s if len(s) <= n else s[: n - 1] + "…"
 
        agg["leaf_label"] = agg["pid"] + " · " + agg["job"].fillna("").apply(lambda j: _short(j, 42))
 
        metric_row([
            ("Total", fmt_money(agg["total"].sum()), COLOR_NAVY, True),
            ("Remaining", fmt_money(agg["remaining"].sum()), COLOR_MUTED, False),
            ("Planned", fmt_money(agg["planned"].sum()), COLOR_AMBER, False),
            ("Done", fmt_money(agg["done"].sum()), COLOR_GREEN, False),
            ("Invoiced", fmt_money(agg["invoiced"].sum()), COLOR_PURPLE, False),
        ])
        st.caption(f"{len(agg)} PID(s) in view — segments are mutually exclusive and sum to Total")
 
        y_levels = [
            agg["district"].fillna("Unassigned").tolist(),
            agg["project"].fillna("Unassigned").tolist(),
            agg["leaf_label"].tolist(),
        ]
        hover_pm = agg["project manager"].fillna("—")
        hover_job = agg["job"].fillna("")
 
        STAGES = [
            ("remaining", "Remaining", COLOR_MUTED),
            ("planned", "Planned", COLOR_AMBER),
            ("done", "Done", COLOR_GREEN),
            ("invoiced", "Invoiced", COLOR_PURPLE),
        ]
 
        fig3 = go.Figure()
        for col, label, color in STAGES:
            vals = agg[col]
            fig3.add_trace(go.Bar(
                y=y_levels, x=vals, name=label, orientation="h",
                marker_color=color,
                text=[fmt_money_short(v) if v >= 1 else "" for v in vals],
                textposition="inside",
                insidetextanchor="middle",
                textfont=dict(size=10, color="#FFFFFF"),
                customdata=list(zip(hover_pm, hover_job)),
                hovertemplate=(
                    f"<b>{label}</b>: £%{{x:,.2f}}<br>PM: %{{customdata[0]}}<br>Job: %{{customdata[1]}}<extra></extra>"
                ),
            ))
 
        row_h = 30
        fig3.update_layout(
            **PLOTLY_LIGHT,
            height=max(420, row_h * len(agg) + 140),
            barmode="stack",
            xaxis=dict(tickprefix="£", gridcolor=GRID_LIGHT, zeroline=False, tickfont=dict(size=12)),
            yaxis=dict(
                gridcolor=GRID_LIGHT, automargin=True, tickfont=dict(size=13),
                autorange="reversed",
            ),
            legend=dict(orientation="h", y=1.04, x=0.5, xanchor="center", bgcolor="rgba(0,0,0,0)", font=dict(size=12)),
            bargap=0.28,
            margin=dict(l=10, r=40, t=40, b=10),
        )
        st.plotly_chart(fig3, use_container_width=True)
 
# --------------------------------------------------------------------------
# TAB 3 — Finance
# --------------------------------------------------------------------------
 
with tab_finance:
    st.subheader("Panel 04 — Financial Health")
    st.caption("Budget performance, billing pipeline & invoicing speed · same filters as above")
 
    fin = fdf_dated.copy()
    fin_c = fin[fin["flag"] == "construction"].copy()
    fin_m = fin[fin["flag"] == "material"].copy()
 
    if fin.empty:
        st.info("No records under these filters. Widen the date range or clear a filter.")
    else:
        fin["has_plan"] = fin["plan1"].notna()
        fin["has_done"] = fin["done"].notna()
        fin["has_inv"] = fin["invoice date"].notna()
 
        total_value = fin["total"].sum()
        orig_value = fin_c["orig"].sum()
        variance = fin_c["total"].sum() - orig_value
        variance_pct = (variance / orig_value * 100) if orig_value else 0.0
 
        invoiced_sum = fin.loc[fin["has_inv"], "total"].sum()
        wip_sum = fin.loc[fin["has_done"] & ~fin["has_inv"], "total"].sum()          # earned, not yet billed
        backlog_sum = fin.loc[~fin["has_done"], "total"].sum()                        # not yet completed
        invoiced_pct = (invoiced_sum / total_value * 100) if total_value else 0.0
        material_pct = (fin_m["total"].sum() / total_value * 100) if total_value else 0.0
 
        # ---- KPI row 1 : budget performance ----
        metric_row([
            ("Total Value", fmt_money(total_value), COLOR_NAVY, True),
            ("Original Budget", fmt_money(orig_value), None, False),
            ("Variance", fmt_money(variance), COLOR_GREEN if variance >= 0 else COLOR_RED, False),
            ("Variance %", f"{variance_pct:+.2f}%", COLOR_GREEN if variance_pct >= 0 else COLOR_RED, False),
        ])
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        # ---- KPI row 2 : cash / pipeline ----
        metric_row([
            ("Invoiced", fmt_money(invoiced_sum) + f"  ({invoiced_pct:.1f}%)", COLOR_PURPLE, False),
            ("WIP — done, not invoiced", fmt_money(wip_sum), COLOR_GREEN, False),
            ("Backlog — not yet done", fmt_money(backlog_sum), COLOR_MUTED, False),
            ("Material share", f"{material_pct:.1f}%", COLOR_YELLOW, False),
        ])
 
        st.markdown("---")
 
        col_a, col_b = st.columns(2)
 
        # ---- Billing pipeline (single stacked bar) ----
        with col_a:
            st.markdown("**Billing pipeline**")
            st.caption("Where the total value currently sits, end to end")
 
            stage_vals = pd.Series({
                "Remaining\n(not started)": fin.loc[~fin["has_plan"], "total"].sum(),
                "Planned\n(not done)": fin.loc[fin["has_plan"] & ~fin["has_done"], "total"].sum(),
                "Done\n(not invoiced)": wip_sum,
                "Invoiced": invoiced_sum,
            })
            pipe_colors = [COLOR_MUTED, COLOR_AMBER, COLOR_GREEN, COLOR_PURPLE]
 
            fig4 = go.Figure()
            for stage, val, color in zip(stage_vals.index, stage_vals.values, pipe_colors):
                fig4.add_trace(go.Bar(
                    y=["Value"], x=[val], name=stage.replace("\n", " "), orientation="h",
                    marker_color=color,
                    text=[fmt_money_short(val)], textposition="inside", textfont=dict(color="#fff", size=12),
                    hovertemplate=f"<b>{stage.replace(chr(10),' ')}</b>: £%{{x:,.2f}}<extra></extra>",
                ))
            fig4.update_layout(
                **PLOTLY_LIGHT,
                height=180,
                barmode="stack",
                xaxis=dict(tickprefix="£", gridcolor=GRID_LIGHT, zeroline=False, tickfont=dict(size=12)),
                yaxis=dict(visible=False),
                legend=dict(orientation="h", y=-0.35, bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
                margin=dict(l=10, r=10, t=10, b=10),
            )
            st.plotly_chart(fig4, use_container_width=True)
 
        # ---- Invoicing lag ----
        with col_b:
            st.markdown("**Invoicing speed**")
            st.caption("Days between job marked *done* and *invoiced*")
 
            lag_df = fin[fin["has_done"] & fin["has_inv"]].copy()
            lag_df["lag_days"] = (lag_df["invoice date"] - lag_df["done"]).dt.days
            lag_df = lag_df[(lag_df["lag_days"] >= 0) & (lag_df["lag_days"] <= 365)]
 
            if lag_df.empty:
                st.info("Not enough done + invoiced pairs under these filters to measure lag.")
            else:
                med_lag = lag_df["lag_days"].median()
                mean_lag = lag_df["lag_days"].mean()
                st.markdown(
                    f"<span style='font-family:\"IBM Plex Mono\",monospace;font-size:13px;color:{TEXT_DARK};'>"
                    f"Median: <b>{med_lag:.0f} days</b> &nbsp;·&nbsp; Average: <b>{mean_lag:.0f} days</b> "
                    f"&nbsp;·&nbsp; n={len(lag_df):,}</span>",
                    unsafe_allow_html=True,
                )
                fig5 = go.Figure()
                fig5.add_trace(go.Histogram(
                    x=lag_df["lag_days"], nbinsx=30, marker_color=COLOR_TEAL,
                    hovertemplate="%{x} days<br>%{y} jobs<extra></extra>",
                ))
                fig5.add_vline(x=med_lag, line_dash="dash", line_color=COLOR_RED,
                                annotation_text="median", annotation_font_size=11)
                fig5.update_layout(
                    **PLOTLY_LIGHT,
                    height=180,
                    xaxis=dict(title="Days", gridcolor=GRID_LIGHT, tickfont=dict(size=11)),
                    yaxis=dict(title="Jobs", gridcolor=GRID_LIGHT, tickfont=dict(size=11)),
                    margin=dict(l=10, r=10, t=10, b=10),
                    bargap=0.1,
                )
                st.plotly_chart(fig5, use_container_width=True)
 
        st.markdown("---")
        st.markdown("**Variance by project**")
        st.caption("Construction total vs. original budget, worst to best")
 
        proj_var = fin_c.groupby("project", dropna=False).agg(total=("total", "sum"), orig=("orig", "sum"))
        proj_var = proj_var[proj_var["orig"] != 0]
        proj_var["variance"] = proj_var["total"] - proj_var["orig"]
        proj_var["variance_pct"] = proj_var["variance"] / proj_var["orig"] * 100
        proj_var = proj_var.sort_values("variance_pct")
 
        if proj_var.empty:
            st.info("No projects with a non-zero original budget under these filters.")
        else:
            fig6 = go.Figure()
            fig6.add_trace(go.Bar(
                y=proj_var.index.fillna("Unassigned"), x=proj_var["variance_pct"], orientation="h",
                marker_color=[COLOR_GREEN if v >= 0 else COLOR_RED for v in proj_var["variance_pct"]],
                text=[f"{v:+.1f}%" for v in proj_var["variance_pct"]],
                textposition="outside", textfont=dict(size=12),
                customdata=proj_var[["total", "orig", "variance"]],
                hovertemplate=(
                    "<b>%{y}</b><br>Total: £%{customdata[0]:,.2f}<br>Original: £%{customdata[1]:,.2f}"
                    "<br>Variance: £%{customdata[2]:,.2f} (%{x:+.1f}%)<extra></extra>"
                ),
            ))
            fig6.update_layout(
                **PLOTLY_LIGHT,
                height=max(220, 42 * len(proj_var) + 60),
                xaxis=dict(title="Variance %", gridcolor=GRID_LIGHT, zeroline=True, zerolinecolor="#B9C2CC", tickfont=dict(size=12)),
                yaxis=dict(gridcolor=GRID_LIGHT, tickfont=dict(size=13), automargin=True),
                margin=dict(l=10, r=40, t=10, b=10),
            )
            st.plotly_chart(fig6, use_container_width=True)
 
        st.markdown("---")
        st.markdown("**PID variance leaderboard**")
        st.caption("Largest overruns and largest underspends, by £ variance (construction only)")
 
        pid_var = fin_c.dropna(subset=["pid"]).groupby(
            ["pid", "district", "project"], dropna=False
        ).agg(total=("total", "sum"), orig=("orig", "sum")).reset_index()
        pid_var = pid_var[pid_var["orig"] != 0]
        pid_var["variance"] = pid_var["total"] - pid_var["orig"]
        pid_var["variance_pct"] = pid_var["variance"] / pid_var["orig"] * 100
 
        if pid_var.empty:
            st.info("No PIDs with a non-zero original budget under these filters.")
        else:
            lead_a, lead_b = st.columns(2)
            with lead_a:
                st.markdown(f"<span style='color:{COLOR_RED};font-weight:700;'>Top overruns</span>", unsafe_allow_html=True)
                worst = pid_var.sort_values("variance").head(8).copy()
                worst["Variance"] = worst["variance"].apply(fmt_money)
                worst["Variance %"] = worst["variance_pct"].apply(lambda v: f"{v:+.1f}%")
                st.dataframe(
                    worst[["pid", "district", "project", "Variance", "Variance %"]]
                    .rename(columns={"pid": "PID", "district": "District", "project": "Project"}),
                    hide_index=True, use_container_width=True,
                )
            with lead_b:
                st.markdown(f"<span style='color:{COLOR_GREEN};font-weight:700;'>Top underspends</span>", unsafe_allow_html=True)
                best = pid_var.sort_values("variance", ascending=False).head(8).copy()
                best["Variance"] = best["variance"].apply(fmt_money)
                best["Variance %"] = best["variance_pct"].apply(lambda v: f"{v:+.1f}%")
                st.dataframe(
                    best[["pid", "district", "project", "Variance", "Variance %"]]
                    .rename(columns={"pid": "PID", "district": "District", "project": "Project"}),
                    hide_index=True, use_container_width=True,
                )
 
st.markdown("---")
st.caption(f"Master Control Dashboard — data as of {date_max.date()}")
 
