#!/usr/bin/env python3
"""
dashboard.py — Hermes Hospitality STR Performance Dashboard
=============================================================
Interactive Streamlit dashboard visualising hotel performance metrics
from STR_Master.xlsx.

Run:
    streamlit run dashboard.py
"""

from __future__ import annotations

from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

# ──────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).resolve().parent
MASTER_FILE = SCRIPT_DIR / "STR_Master.xlsx"

# Consistent hotel colour palette (colour-blind friendly)
HOTEL_COLORS = {
    "HEZCN": "#4e79a7",
    "JANGM": "#f28e2b",
    "JANTW": "#e15759",
    "LQCHA": "#76b7b2",
    "MSYHV": "#59a14f",
}

HOTEL_NAMES = {
    "HEZCN": "Holiday Inn Natchez",
    "JANGM": "Holiday Inn Jackson",
    "JANTW": "JANTW Property",
    "LQCHA": "La Quinta Chattanooga",
    "MSYHV": "MSYHV Property",
}

INDEX_METRICS = ["MPI", "ARI", "RGI"]
PERIODS = {"7-Day": "7d", "28-Day": "28d"}


# ──────────────────────────────────────────────────────────────
# Page configuration & custom CSS
# ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Hermes STR Dashboard",
    page_icon=":hotel:",
    layout="wide",
    initial_sidebar_state="expanded",
)

_CSS = """
<style>
/* ── KPI Cards ── */
div[data-testid="stMetric"] {
    background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.25);
}
div[data-testid="stMetric"] label {
    color: #94a3b8 !important;
    font-size: 0.82rem !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}
div[data-testid="stMetric"] [data-testid="stMetricValue"] {
    color: #f1f5f9 !important;
    font-weight: 700;
    font-size: 1.6rem !important;
}

/* ── Section dividers ── */
.section-header {
    font-size: 1.15rem;
    font-weight: 600;
    color: #cbd5e1;
    margin-top: 2rem;
    margin-bottom: 0.5rem;
    padding-bottom: 6px;
    border-bottom: 2px solid #334155;
}

/* ── Heatmap table ── */
.heatmap-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.88rem;
    margin-top: 8px;
}
.heatmap-table th {
    background: #1e293b;
    color: #94a3b8;
    padding: 8px 12px;
    text-align: center;
    font-weight: 600;
    text-transform: uppercase;
    font-size: 0.75rem;
    letter-spacing: 0.04em;
}
.heatmap-table td {
    padding: 8px 12px;
    text-align: center;
    font-weight: 600;
    border-bottom: 1px solid #1e293b;
}
.heatmap-table tr:hover {
    background: #1e293b;
}
.hotel-name-cell {
    text-align: left !important;
    color: #e2e8f0;
    font-weight: 700;
}
</style>
"""

st.markdown(_CSS, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# Data loading
# ──────────────────────────────────────────────────────────────


@st.cache_data(ttl=60)
def load_data() -> pd.DataFrame:
    """Load and prepare the master dataset."""
    df = pd.read_excel(MASTER_FILE, engine="openpyxl")
    df["Date"] = pd.to_datetime(df["Date"])
    df.sort_values(["Inn Code", "Date"], inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


# ──────────────────────────────────────────────────────────────
# Sidebar
# ──────────────────────────────────────────────────────────────


def render_sidebar(df: pd.DataFrame) -> tuple[list[str], pd.Timestamp, pd.Timestamp, str]:
    """Render sidebar controls and return filter values."""
    with st.sidebar:
        st.markdown("## :control_knobs: Dashboard Controls")
        st.markdown("---")

        # Hotel selection
        all_hotels = sorted(df["Inn Code"].unique())
        selected = st.multiselect(
            "Select Properties",
            options=all_hotels,
            default=all_hotels,
            help="Filter which hotels appear in charts and KPIs",
        )
        if not selected:
            selected = all_hotels
            st.info("Showing all properties (none selected)")

        st.markdown("---")

        # Date range
        min_date = df["Date"].min().date()
        max_date = df["Date"].max().date()
        date_range = st.slider(
            "Date Range",
            min_value=min_date,
            max_value=max_date,
            value=(min_date, max_date),
            format="MMM DD, YYYY",
        )

        st.markdown("---")

        # Period toggle
        period = st.radio(
            "Metric Period",
            options=list(PERIODS.keys()),
            index=1,  # default 28-Day
            help="Toggle between 7-Day and 28-Day metric windows",
            horizontal=True,
        )

        st.markdown("---")
        st.markdown(
            "<small style='color:#64748b'>Data source: STR_Master.xlsx</small>",
            unsafe_allow_html=True,
        )

    return selected, pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1]), period


# ──────────────────────────────────────────────────────────────
# KPI Cards
# ──────────────────────────────────────────────────────────────


def render_kpis(df: pd.DataFrame, period_key: str) -> None:
    """Render the top KPI summary row."""
    rgi_idx_col = f"RGI_{period_key}_Index"
    rgi_pct_col = f"RGI_{period_key}_PctChg"

    # Use each hotel's most recent row (not all hotels share the same latest date)
    latest = df.sort_values("Date").groupby("Inn Code").tail(1).reset_index(drop=True)

    # For deltas, grab each hotel's second-to-last row
    prior = df.sort_values("Date").groupby("Inn Code", as_index=False).nth(-2).reset_index(drop=True)
    has_prior = not prior.empty

    cols = st.columns(4)

    # Card 1: Portfolio Avg RGI Index
    avg_rgi = latest[rgi_idx_col].mean()
    delta_avg = None
    if has_prior and not prior.empty:
        prior_avg = prior[rgi_idx_col].mean()
        delta_avg = f"{avg_rgi - prior_avg:+.1f}"
    cols[0].metric(
        "Portfolio Avg RGI Index",
        f"{avg_rgi:.1f}",
        delta=delta_avg,
    )

    # Card 2: Properties Above Par (Index > 100)
    above_par = int((latest[rgi_idx_col] > 100).sum())
    total = len(latest)
    delta_par = None
    if has_prior and not prior.empty:
        prior_above = int((prior[rgi_idx_col] > 100).sum())
        diff = above_par - prior_above
        if diff != 0:
            delta_par = f"{diff:+d} vs prior week"
    cols[1].metric(
        "Properties Above Par",
        f"{above_par} / {total}",
        delta=delta_par,
    )

    # Card 3: Best Performer
    if not latest.empty:
        best_row = latest.loc[latest[rgi_idx_col].idxmax()]
        best_code = best_row["Inn Code"]
        best_val = best_row[rgi_idx_col]
        delta_best = None
        if has_prior and not prior.empty:
            prior_best = prior[prior["Inn Code"] == best_code]
            if not prior_best.empty:
                delta_best = f"{best_val - prior_best.iloc[0][rgi_idx_col]:+.1f}"
        cols[2].metric(
            f"Best Performer ({best_code})",
            f"{best_val:.1f}",
            delta=delta_best,
        )

    # Card 4: Biggest WoW Swing
    if has_prior and not prior.empty:
        merged = latest[["Inn Code", rgi_pct_col]].merge(
            prior[["Inn Code", rgi_pct_col]],
            on="Inn Code",
            suffixes=("_now", "_prev"),
        )
        if not merged.empty:
            merged["swing"] = merged[f"{rgi_pct_col}_now"] - merged[f"{rgi_pct_col}_prev"]
            merged["abs_swing"] = merged["swing"].abs()
            biggest = merged.loc[merged["abs_swing"].idxmax()]
            swing_val = biggest["swing"]
            cols[3].metric(
                f"Biggest Swing ({biggest['Inn Code']})",
                f"{swing_val:+.1f}pp",
                delta=f"{'Surge' if swing_val > 0 else 'Drop'} in RGI % Chg",
                delta_color="normal" if swing_val > 0 else "inverse",
            )
    else:
        cols[3].metric("Biggest Swing", "N/A", delta="Insufficient data")


# ──────────────────────────────────────────────────────────────
# Time-Series Chart
# ──────────────────────────────────────────────────────────────


def render_timeseries(df: pd.DataFrame, period_key: str, selected_hotels: list[str]) -> None:
    """Render the 28-Day / 7-Day RGI Index time-series chart."""
    st.markdown(
        '<div class="section-header">📈 RGI Index — Historical Trends</div>',
        unsafe_allow_html=True,
    )

    rgi_col = f"RGI_{period_key}_Index"
    period_label = "28-Day" if period_key == "28d" else "7-Day"

    # Build colour domain/range based on selected hotels
    color_domain = [h for h in selected_hotels if h in HOTEL_COLORS]
    color_range = [HOTEL_COLORS[h] for h in color_domain]

    # Base line chart
    line = (
        alt.Chart(df)
        .mark_line(strokeWidth=2.5, point=alt.OverlayMarkDef(size=50, filled=True))
        .encode(
            x=alt.X("Date:T", title="Week Ending", axis=alt.Axis(format="%b %d", labelAngle=-45)),
            y=alt.Y(f"{rgi_col}:Q", title=f"{period_label} RGI Index", scale=alt.Scale(zero=False)),
            color=alt.Color(
                "Inn Code:N",
                scale=alt.Scale(domain=color_domain, range=color_range),
                legend=alt.Legend(title="Property", orient="top", columns=5),
            ),
            tooltip=[
                alt.Tooltip("Inn Code:N", title="Hotel"),
                alt.Tooltip("Date:T", title="Week", format="%b %d, %Y"),
                alt.Tooltip(f"{rgi_col}:Q", title="RGI Index", format=".1f"),
            ],
        )
    )

    # Par reference line at 100
    par_line = (
        alt.Chart(pd.DataFrame({"y": [100]}))
        .mark_rule(color="#ef4444", strokeDash=[6, 4], strokeWidth=1.5, opacity=0.7)
        .encode(y="y:Q")
    )

    par_text = (
        alt.Chart(pd.DataFrame({"y": [100], "label": ["Par (100)"]}))
        .mark_text(align="left", dx=5, dy=-8, fontSize=11, color="#ef4444", fontStyle="italic")
        .encode(y="y:Q", text="label:N")
    )

    chart = (line + par_line + par_text).properties(height=420).interactive()

    st.altair_chart(chart, use_container_width=True)


# ──────────────────────────────────────────────────────────────
# Comparative Heatmap
# ──────────────────────────────────────────────────────────────


def _heatmap_cell_color(value: float, is_index: bool) -> str:
    """Return a CSS background colour for a heatmap cell."""
    if pd.isna(value):
        return "background-color: #1e293b; color: #64748b;"
    if is_index:
        # Index: par = 100. Green above, red below.
        if value >= 110:
            return "background-color: #166534; color: #bbf7d0;"
        elif value >= 100:
            return "background-color: #14532d; color: #86efac;"
        elif value >= 90:
            return "background-color: #7f1d1d; color: #fecaca;"
        else:
            return "background-color: #991b1b; color: #fca5a5;"
    else:
        # % Change: 0 is neutral
        if value > 5:
            return "background-color: #166534; color: #bbf7d0;"
        elif value > 0:
            return "background-color: #14532d; color: #86efac;"
        elif value > -5:
            return "background-color: #7f1d1d; color: #fecaca;"
        else:
            return "background-color: #991b1b; color: #fca5a5;"


def render_heatmap(df: pd.DataFrame, period_key: str) -> None:
    """Render the latest-week comparative heatmap."""
    st.markdown(
        '<div class="section-header">📊 Latest Week — Cross-Property Comparison</div>',
        unsafe_allow_html=True,
    )

    latest_date = df["Date"].max()
    latest = df[df["Date"] == latest_date].copy()
    period_label = "28-Day" if period_key == "28d" else "7-Day"

    st.caption(f"Reporting week: **{latest_date.strftime('%b %d, %Y')}** | Period: **{period_label}**")

    metric_cols = {
        f"MPI_{period_key}_Index": "MPI Index",
        f"ARI_{period_key}_Index": "ARI Index",
        f"RGI_{period_key}_Index": "RGI Index",
        f"MPI_{period_key}_PctChg": "MPI % Chg",
        f"ARI_{period_key}_PctChg": "ARI % Chg",
        f"RGI_{period_key}_PctChg": "RGI % Chg",
    }

    # Build HTML table
    header = "<tr><th>Property</th>"
    for display in metric_cols.values():
        header += f"<th>{display}</th>"
    header += "</tr>"

    rows_html = ""
    for _, row in latest.sort_values("Inn Code").iterrows():
        code = row["Inn Code"]
        name = HOTEL_NAMES.get(code, code)
        rows_html += f'<tr><td class="hotel-name-cell">{code}<br><small style="color:#64748b;font-weight:400">{name}</small></td>'
        for col, display in metric_cols.items():
            val = row[col]
            is_index = "Index" in display
            style = _heatmap_cell_color(val, is_index)
            if pd.notna(val):
                fmt = f"{val:.1f}" if is_index else f"{val:+.1f}%"
            else:
                fmt = "—"
            rows_html += f'<td style="{style}; border-radius: 4px;">{fmt}</td>'
        rows_html += "</tr>"

    html = f'<table class="heatmap-table">{header}{rows_html}</table>'
    st.markdown(html, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# Trend Sparklines
# ──────────────────────────────────────────────────────────────


def render_sparklines(df: pd.DataFrame, period_key: str) -> None:
    """Render compact sparkline trend charts per hotel."""
    st.markdown(
        '<div class="section-header">✨ Property Trend Snapshots</div>',
        unsafe_allow_html=True,
    )

    rgi_col = f"RGI_{period_key}_Index"
    hotels = sorted(df["Inn Code"].unique())

    cols_per_row = min(len(hotels), 5)
    grid = st.columns(cols_per_row)

    for i, code in enumerate(hotels):
        hdf = df[df["Inn Code"] == code].sort_values("Date")
        col = grid[i % cols_per_row]

        with col:
            latest_val = hdf[rgi_col].iloc[-1] if not hdf.empty else None
            delta = None
            if len(hdf) >= 2:
                delta = hdf[rgi_col].iloc[-1] - hdf[rgi_col].iloc[-2]

            name = HOTEL_NAMES.get(code, code)
            st.markdown(f"**{code}**")
            st.caption(name)

            if latest_val is not None:
                st.metric(
                    "RGI Index",
                    f"{latest_val:.1f}",
                    delta=f"{delta:+.1f}" if delta is not None else None,
                    label_visibility="collapsed",
                )

            # Sparkline chart
            spark = (
                alt.Chart(hdf)
                .mark_area(
                    line={"color": HOTEL_COLORS.get(code, "#8884d8"), "strokeWidth": 2},
                    opacity=0.15,
                    color=alt.Gradient(
                        gradient="linear",
                        stops=[
                            alt.GradientStop(color=HOTEL_COLORS.get(code, "#8884d8"), offset=0),
                            alt.GradientStop(color="transparent", offset=1),
                        ],
                        x1=1, x2=1, y1=0, y2=1,
                    ),
                )
                .encode(
                    x=alt.X("Date:T", axis=None),
                    y=alt.Y(f"{rgi_col}:Q", axis=None, scale=alt.Scale(zero=False)),
                )
                .properties(height=60)
            )
            st.altair_chart(spark, use_container_width=True)


# ──────────────────────────────────────────────────────────────
# Data Explorer
# ──────────────────────────────────────────────────────────────


def render_data_explorer(df: pd.DataFrame) -> None:
    """Expandable raw data viewer with CSV download."""
    st.markdown(
        '<div class="section-header">📁 Data Explorer</div>',
        unsafe_allow_html=True,
    )

    with st.expander("View Raw Data", expanded=False):
        st.dataframe(
            df.style.format(
                {col: "{:.2f}" for col in df.select_dtypes("float").columns}
            ),
            use_container_width=True,
            height=400,
        )

        csv = df.to_csv(index=False)
        st.download_button(
            label="📥 Download CSV",
            data=csv,
            file_name="str_master_export.csv",
            mime="text/csv",
        )


# ──────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────


def main() -> None:
    # Header
    st.markdown(
        """
        <h1 style="
            background: linear-gradient(90deg, #3b82f6, #8b5cf6);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 2rem;
            font-weight: 800;
            margin-bottom: 0;
        ">Hermes Hospitality</h1>
        <p style="color: #64748b; margin-top: 0; font-size: 1.05rem;">
            STR Performance Dashboard
        </p>
        """,
        unsafe_allow_html=True,
    )

    # Load data
    raw_df = load_data()

    if raw_df.empty:
        st.error("No data found in STR_Master.xlsx. Run process_reports.py first.")
        return

    # Sidebar filters
    selected_hotels, start_date, end_date, period_label = render_sidebar(raw_df)
    period_key = PERIODS[period_label]

    # Apply filters
    mask = (
        (raw_df["Inn Code"].isin(selected_hotels))
        & (raw_df["Date"] >= start_date)
        & (raw_df["Date"] <= end_date)
    )
    df = raw_df[mask].copy()

    if df.empty:
        st.warning("No data matches the current filters.")
        return

    # Render sections
    render_kpis(df, period_key)
    st.markdown("")  # spacer
    render_timeseries(df, period_key, selected_hotels)
    render_heatmap(df, period_key)
    render_sparklines(df, period_key)
    render_data_explorer(df)


if __name__ == "__main__":
    main()
