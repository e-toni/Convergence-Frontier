# =============================================================================
#  CONVERGENCE FRONTIER DASHBOARD
#
#  Interactive front-end for the GMC master workbook produced by
#  Run_GMC_pwt_results_v7.R  ->  DCT_GMC_master_pwt_v7.xlsx.
#
#  The dashboard reads the workbook only; it computes no smoothing of its own.
#  All trend extraction, the convergence-frontier quantile regressions and the
#  seven-method robustness battery are done in R. This app is a viewer.
#
#  WHAT IT SHOWS
#  -------------
#    * The convergence frontier (25th / 50th / 75th quantile-regression curves)
#      for a chosen GDP concept, under a chosen trend filter OR the mean
#      ensemble of the seven filters.
#    * Each country's latest growth-income position, classified Accelerator /
#      On Track / Stagnating against that frontier.
#    * The seven-filter uncertainty bar (min..max trend growth) and the
#      robustness flag (Robust / Borderline / Highly uncertain).
#    * Full historical trajectories in the growth-income plane.
#
#  This is the convergence-frontier framework only; the Income-Tier
#  Categorization has been removed.
#
#  Run:  streamlit run app.py
#  Requires DCT_GMC_master_pwt_v7.xlsx in the same folder.
# =============================================================================

import os
import re
import glob
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="Convergence Frontier Dashboard",
    page_icon="🌍",
    layout="wide",
)


# ============================================================
# CONFIGURATION
# ============================================================

DEFAULT_WORKBOOK = "DCT_GMC_master_pwt_v7.xlsx"

# The seven-method robustness battery (order = display order).
METHOD_IDS = ["hp625", "hp100", "hamilton", "llt", "cf", "bn", "ma5"]

# Fallback labels if Methods_legend is unavailable.
METHOD_LABELS_FALLBACK = {
    "hp625":    "HP filter (lambda = 6.25)",
    "hp100":    "HP filter (lambda = 100)",
    "hamilton": "Hamilton (2018) regression filter",
    "llt":      "Local linear trend (state-space)",
    "cf":       "Christiano-Fitzgerald band-pass",
    "bn":       "Beveridge-Nelson decomposition",
    "ma5":      "Trailing 5-year moving average",
}

# Short, unique labels for table column headers and other compact uses.
# These match the labels used in the paper's per-filter classification table.
METHOD_SHORT_LABELS = {
    "hp625":    "HP-6.25",
    "hp100":    "HP-100",
    "hamilton": "Hamilton",
    "llt":      "LLT",
    "cf":       "CF",
    "bn":       "BN",
    "ma5":      "MA-5",
}

# Distinct colours for country trajectories.
# Set EXPLICITLY on every trajectory trace so the figure looks the same in
# the dashboard (where Streamlit applies its theme) and in the exported
# PDF/HTML (where kaleido renders without any Streamlit theme context).
# Without explicit colours, traces rely on the Plotly auto-colorway, which
# is filled in by the theme at display time but can fall back to a
# monochrome default in static export -- the cause of the "all-black PDF"
# bug. The 14 colours match the 14 marker symbols (see below) so colour
# and shape change together as countries are added.
TRAJ_COLORS = [
    "#1f77b4",  # blue
    "#ff7f0e",  # orange
    "#2ca02c",  # green
    "#d62728",  # red
    "#9467bd",  # purple
    "#8c564b",  # brown
    "#e377c2",  # pink
    "#17becf",  # cyan
    "#bcbd22",  # olive
    "#7f7f7f",  # grey
    "#1f9e89",  # teal
    "#e6550d",  # dark orange
    "#756bb1",  # violet
    "#a55194",  # magenta
]

ENSEMBLE_KEY = "__ensemble__"
ENSEMBLE_LABEL = "Ensemble - mean of the seven filters"

# Human-readable GDP-concept labels.
MEASURE_LABELS = {
    "rgdpo":  "rgdpo - output-side real GDP, chained PPPs",
    "rgdpna": "rgdpna - real GDP, constant national prices",
    "rgdpe":  "rgdpe - expenditure-side real GDP, chained PPPs",
    "cgdpe":  "cgdpe - expenditure-side real GDP, current PPPs",
    "cgdpo":  "cgdpo - output-side real GDP, current PPPs",
}

# GMC classes and their plotting colours.
CLASS_ORDER = ["Accelerator", "On Track", "Stagnating"]
CLASS_COLORS = {
    "Accelerator": "#1b7837",
    "On Track":    "#3690c0",
    "Stagnating":  "#b2182b",
}

# Robustness flag colours.
ROBUST_ORDER = ["Robust", "Borderline", "Highly uncertain"]
ROBUST_COLORS = {
    "Robust":           "#1b7837",
    "Borderline":       "#e08214",
    "Highly uncertain": "#b2182b",
}

# Default benchmark country set (PWT 11.0 names).
DEFAULT_COUNTRIES = [
    "Bolivia (Plurinational State of)", "Brazil", "Bulgaria", "Burundi",
    "Chile", "China", "Croatia", "Denmark", "Finland", "France", "Germany",
    "China, Hong Kong SAR", "India", "Ireland", "Italy", "Japan",
    "Republic of Korea", "Kuwait", "Lithuania", "Luxembourg", "Norway",
    "Poland", "Portugal", "Singapore", "Spain", "Sweden", "Switzerland",
    "United Kingdom", "United States", "Uruguay",
]

X_LIMITS = (0, 150000)


# ============================================================
# HELPERS
# ============================================================

def safe_filename(text):
    text = str(text).lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return re.sub(r"_+", "_", text).strip("_")


def find_workbook():
    """Locate the v7 master workbook next to this script."""
    here = os.path.dirname(os.path.abspath(__file__))
    exact = os.path.join(here, DEFAULT_WORKBOOK)
    if os.path.exists(exact):
        return exact
    if os.path.exists(DEFAULT_WORKBOOK):
        return DEFAULT_WORKBOOK
    cands = sorted(glob.glob(os.path.join(here, "DCT_GMC_master_pwt*.xlsx")))
    if not cands:
        cands = sorted(glob.glob("DCT_GMC_master_pwt*.xlsx"))
    return cands[-1] if cands else None


def measure_label(measure_id):
    return MEASURE_LABELS.get(measure_id, measure_id)


# ============================================================
# DATA LOADING
# ============================================================

@st.cache_data(show_spinner=False)
def load_workbook(path):
    """Load the GMC master workbook sheets needed by the dashboard."""
    xls = pd.ExcelFile(path)

    required = {
        "GMC_classification", "GMC_frontier_coefs",
        "GMC_frontier_ensemble", "GMC_trajectory",
    }
    missing = required - set(xls.sheet_names)
    if missing:
        raise ValueError(
            "Workbook is missing required sheet(s): "
            + ", ".join(sorted(missing))
            + ". This dashboard expects a v7 GMC master workbook."
        )

    classification = pd.read_excel(xls, "GMC_classification")
    coefs          = pd.read_excel(xls, "GMC_frontier_coefs")
    ensemble       = pd.read_excel(xls, "GMC_frontier_ensemble")
    trajectory     = pd.read_excel(xls, "GMC_trajectory")

    legend = (pd.read_excel(xls, "Methods_legend")
              if "Methods_legend" in xls.sheet_names else None)
    robustness_summary = (pd.read_excel(xls, "GMC_robustness_summary")
                          if "GMC_robustness_summary" in xls.sheet_names else None)

    # numeric coercion
    for col in ["income_pc", "year", "growth_min", "growth_max",
                "growth_ensemble_mean"] + [f"g_{m}" for m in METHOD_IDS]:
        if col in classification.columns:
            classification[col] = pd.to_numeric(classification[col], errors="coerce")
    for col in ["tau", "intercept", "slope_logpc"]:
        coefs[col]    = pd.to_numeric(coefs[col], errors="coerce")
        ensemble[col] = pd.to_numeric(ensemble[col], errors="coerce")
    for col in ["year", "gdp_pc"] + [f"sm_{m}" for m in METHOD_IDS]:
        if col in trajectory.columns:
            trajectory[col] = pd.to_numeric(trajectory[col], errors="coerce")

    return {
        "classification": classification,
        "coefs": coefs,
        "ensemble": ensemble,
        "trajectory": trajectory,
        "legend": legend,
        "robustness_summary": robustness_summary,
    }


def method_label_map(legend):
    """Build a {method_id: label} map from the Methods_legend sheet."""
    labels = dict(METHOD_LABELS_FALLBACK)
    if legend is not None and {"method_id", "label"}.issubset(legend.columns):
        for _, row in legend.iterrows():
            labels[str(row["method_id"])] = str(row["label"])
    return labels


# ============================================================
# FRONTIER + CLASSIFICATION LOGIC
# ============================================================

def frontier_curve(coef_rows, x_grid):
    """Build 25/50/75 frontier curves from 3 quantile-regression rows.

    coef_rows has columns tau, intercept, slope_logpc (one row per tau).
    Q_tau(y) = intercept + slope * log(y).
    """
    out = {"gdp_pc": x_grid}
    safe_x = np.where(x_grid > 0, x_grid, np.nan)
    for tau, name in [(0.25, "p25"), (0.50, "p50"), (0.75, "p75")]:
        row = coef_rows[np.isclose(coef_rows["tau"], tau)]
        if len(row) == 0:
            out[name] = np.full_like(x_grid, np.nan, dtype=float)
        else:
            a = float(row["intercept"].iloc[0])
            b = float(row["slope_logpc"].iloc[0])
            out[name] = a + b * np.log(safe_x)
    return pd.DataFrame(out)


def selected_frontier(data, measure_id, method_key, x_grid):
    """Return the frontier DataFrame for the chosen measure and method."""
    if method_key == ENSEMBLE_KEY:
        rows = data["ensemble"]
        rows = rows[rows["gdp_measure"] == measure_id]
    else:
        rows = data["coefs"]
        rows = rows[(rows["gdp_measure"] == measure_id)
                    & (rows["method"] == method_key)]
    return frontier_curve(rows[["tau", "intercept", "slope_logpc"]], x_grid)


def latest_points(data, measure_id, method_key):
    """Latest-year country points for the chosen measure and method.

    Returns a DataFrame with: Country Name/Code, year, income_pc, growth
    (the chosen method's trend growth), gmc_class, growth_min/max, robustness.
    """
    cl = data["classification"]
    cl = cl[cl["gdp_measure"] == measure_id].copy()

    if method_key == ENSEMBLE_KEY:
        cl["growth"]    = cl["growth_ensemble_mean"]
        cl["gmc_class"] = cl["class_ensemble_mean"]
    else:
        cl["growth"]    = cl[f"g_{method_key}"]
        cl["gmc_class"] = cl[f"cls_{method_key}"]

    keep = ["Country Name", "Country Code", "year", "income_pc", "growth",
            "gmc_class", "growth_min", "growth_max", "robustness",
            "in_estimation_sample"]
    keep = [c for c in keep if c in cl.columns]
    cl = cl[keep].dropna(subset=["income_pc", "growth", "Country Name"])
    return cl


def trajectory_series(data, measure_id, method_key):
    """Full trajectory panel with a single 'growth' column for the method."""
    tr = data["trajectory"]
    tr = tr[tr["gdp_measure"] == measure_id].copy()
    if method_key == ENSEMBLE_KEY:
        sm_cols = [f"sm_{m}" for m in METHOD_IDS if f"sm_{m}" in tr.columns]
        tr["growth"] = tr[sm_cols].mean(axis=1, skipna=True)
    else:
        tr["growth"] = tr[f"sm_{method_key}"]
    keep = ["Country Name", "Country Code", "year", "gdp_pc", "growth"]
    return tr[[c for c in keep if c in tr.columns]].dropna(
        subset=["gdp_pc", "growth"])


# ============================================================
# CHART
# ============================================================

def build_chart(frontier, pts, traj, sel_latest, sel_traj,
                 sel_classes, show_band, show_median, show_bars,
                 colour_mode, measure_id, method_label_text):
    """Assemble the convergence-frontier Plotly figure."""
    fig = go.Figure()

    # ---- frontier band + lines ---------------------------------------
    if show_band:
        fig.add_trace(go.Scatter(
            x=frontier["gdp_pc"], y=frontier["p75"], mode="lines",
            name="75th percentile", line=dict(dash="dash", width=1.5,
                                               color="#2166ac"),
            hovertemplate="GDP per capita: %{x:,.0f}<br>75th pct: %{y:.2f}%"
                          "<extra></extra>"))
        fig.add_trace(go.Scatter(
            x=frontier["gdp_pc"], y=frontier["p25"], mode="lines",
            name="25th percentile", line=dict(dash="dash", width=1.5,
                                               color="#2166ac"),
            fill="tonexty", fillcolor="rgba(67,147,195,0.15)",
            hovertemplate="GDP per capita: %{x:,.0f}<br>25th pct: %{y:.2f}%"
                          "<extra></extra>"))
    if show_median:
        fig.add_trace(go.Scatter(
            x=frontier["gdp_pc"], y=frontier["p50"], mode="lines",
            name="Convergence frontier (median)",
            line=dict(width=3, color="#b2182b"),
            hovertemplate="GDP per capita: %{x:,.0f}<br>Median fit: %{y:.2f}%"
                          "<extra></extra>"))

    # ---- latest datapoints -------------------------------------------
    pts_plot = pts[pts["Country Name"].isin(sel_latest)].copy()
    if sel_classes:
        pts_plot = pts_plot[pts_plot["gmc_class"].isin(sel_classes)]

    if not pts_plot.empty:
        # uncertainty bars (seven-filter min..max), one trace per robustness
        if show_bars and {"growth_min", "growth_max"}.issubset(pts_plot.columns):
            for flag in ROBUST_ORDER:
                d = pts_plot[pts_plot["robustness"] == flag]
                if d.empty:
                    continue
                fig.add_trace(go.Scatter(
                    x=d["income_pc"], y=d["growth"], mode="markers",
                    name=f"Uncertainty range - {flag}",
                    marker=dict(size=1, color=ROBUST_COLORS[flag]),
                    error_y=dict(
                        type="data", symmetric=False,
                        array=(d["growth_max"] - d["growth"]).clip(lower=0),
                        arrayminus=(d["growth"] - d["growth_min"]).clip(lower=0),
                        color=ROBUST_COLORS[flag], thickness=1.4, width=0),
                    showlegend=False, hoverinfo="skip"))

        # the points themselves
        if colour_mode == "Robustness":
            groups, palette, gcol = ROBUST_ORDER, ROBUST_COLORS, "robustness"
        elif colour_mode == "GMC class":
            groups, palette, gcol = CLASS_ORDER, CLASS_COLORS, "gmc_class"
        else:
            groups, palette, gcol = [None], None, None

        for g in groups:
            if gcol is None:
                d = pts_plot
                colour, nm = "#2166ac", "Latest datapoint"
            else:
                d = pts_plot[pts_plot[gcol] == g]
                colour, nm = palette[g], g
            if d.empty:
                continue
            fig.add_trace(go.Scatter(
                x=d["income_pc"], y=d["growth"],
                mode="markers+text", name=nm,
                text=d["Country Name"], textposition="middle right",
                textfont=dict(size=11),
                marker=dict(size=10, color=colour, opacity=0.85,
                            line=dict(width=0.6, color="white")),
                customdata=np.column_stack([
                    d["Country Name"].astype(str),
                    d["gmc_class"].astype(str),
                    d["robustness"].astype(str) if "robustness" in d else
                    np.array([""] * len(d)),
                    d["growth_min"].round(2).astype(str) if "growth_min" in d
                    else np.array([""] * len(d)),
                    d["growth_max"].round(2).astype(str) if "growth_max" in d
                    else np.array([""] * len(d)),
                ]),
                hovertemplate=(
                    "<b>%{customdata[0]}</b>"
                    "<br>GDP per capita: %{x:,.0f}"
                    "<br>Trend growth: %{y:.2f}%"
                    "<br>Class: %{customdata[1]}"
                    "<br>Robustness: %{customdata[2]}"
                    "<br>Seven-filter range: %{customdata[3]}% .. "
                    "%{customdata[4]}%<extra></extra>")))

    # ---- trajectories -------------------------------------------------
    # Each country trajectory gets a distinct marker symbol AND a distinct
    # colour on top of the auto-colorway, so it can be told apart from the
    # others even in a black-and-white print of the PDF, or in dense regions
    # of the plane. Both attributes are set EXPLICITLY here so the exported
    # figure looks the same as the displayed figure (Streamlit's theme is
    # not visible to kaleido at export time, see the comment on TRAJ_COLORS).
    traj_symbols = [
        "circle", "square", "diamond", "triangle-up", "pentagon",
        "star", "hexagon", "cross", "triangle-down", "x",
        "star-square", "hourglass", "bowtie", "circle-cross",
    ]
    traj_sel = traj[traj["Country Name"].isin(sel_traj)]
    for i, country in enumerate(sel_traj):
        dc = traj_sel[traj_sel["Country Name"] == country].sort_values("year")
        if dc.empty:
            continue
        labels = [""] * (len(dc) - 1) + [country]
        symbol = traj_symbols[i % len(traj_symbols)]
        colour = TRAJ_COLORS[i % len(TRAJ_COLORS)]
        fig.add_trace(go.Scatter(
            x=dc["gdp_pc"], y=dc["growth"], mode="lines+markers+text",
            name=f"{country} trajectory", text=labels,
            textposition="middle right", textfont=dict(size=11),
            marker=dict(size=8, symbol=symbol, color=colour,
                        line=dict(width=0.8, color="white")),
            line=dict(width=2, color=colour),
            customdata=dc["year"].astype(int).astype(str),
            hovertemplate=("<b>" + country + "</b>"
                           "<br>Year: %{customdata}"
                           "<br>GDP per capita: %{x:,.0f}"
                           "<br>Trend growth: %{y:.2f}%<extra></extra>")))

    fig.add_hline(y=0, line_dash="dot", line_width=1, line_color="grey")

    fig.update_layout(
        title=(f"The convergence frontier - {measure_id}  |  {method_label_text}"),
        xaxis_title="GDP per capita",
        yaxis_title="Trend per-capita growth (%)",
        hovermode="closest", height=720,
        margin=dict(l=40, r=30, t=70, b=50),
        legend=dict(orientation="h", yanchor="bottom", y=-0.22,
                    xanchor="left", x=0),
    )
    fig.update_xaxes(tickformat=",", range=list(X_LIMITS))
    fig.update_yaxes(zeroline=True)
    return fig, pts_plot, traj_sel


# ============================================================
# APP BODY
# ============================================================

# ---- Page selector (Dashboard / About) ----------------------------------
# A single-file multi-page setup: a radio in the sidebar selects the view.
# When "About" is selected, we render the About content and stop before
# loading any data, so the dashboard sidebar controls do not appear.
with st.sidebar:
    page = st.radio(
        "View",
        options=["Dashboard", "About"],
        index=0,
    )
    st.divider()

if page == "About":
    st.title("🌍 About the Convergence Platform")
    st.markdown(
        """
        The **Convergence Platform** is the interactive companion to the
        working paper *The Convergence Frontier: Benchmarking Growth Momentum
        and Measuring its Uncertainty* (Toni, 2026). It is a transparent,
        replicable diagnostic for the cross-country convergence debate: every
        economy's latest growth-income position can be examined against an
        empirically estimated benchmark, under any combination of GDP concept
        and trend-extraction filter, with the resulting classification
        uncertainty reported alongside the headline verdict.

        The platform exists so that any claim about a country catching up,
        falling behind, or stagnating can be inspected, contested, and
        replicated --- on the same data, under the same explicit method
        assumptions, in seconds.
        """
    )

    st.divider()

    col1, col2 = st.columns([2, 1])
    with col1:
        st.subheader("The paper")
        st.markdown(
            """
**Title** — *The Convergence Frontier: Benchmarking Growth Momentum and Measuring its Uncertainty*

**Author** — Emiliano Toni, University of St. Gallen

**Contact** — [emiliano.toni@unisg.ch](mailto:emiliano.toni@unisg.ch)

**Website** — https://sites.google.com/view/emilianotoni/home

**Version** — May 2026 (working paper)

The paper introduces a quantile-regression benchmark of trend growth on
income (the *convergence frontier*), estimated on the converged-economy
sample of the Penn World Table, and uses it to classify every economy as
*Accelerator*, *On Track*, or *Stagnating*. Because both the trend and
the income concept are themselves measurement choices, every
classification is re-estimated under a battery of seven trend filters and
across five GDP concepts, and the resulting cross-method disagreement is
reported as a first-class output rather than a methodological footnote.
            """
        )

    with col2:
        st.subheader("Quick facts")
        st.markdown(
            """
**Sample** — 185 economies + EU-17, 1950–2023

**Data source** — Penn World Table 11.0

**Trend filters** — HP-6.25, HP-100, Hamilton, LLT, CF, BN, MA-5

**GDP concepts** — rgdpo, rgdpna, rgdpe, cgdpe, cgdpo

**Quantiles** — 0.25, 0.50, 0.75

**Classes** — Accelerator, On Track, Stagnating

**Robustness flags** — Robust, Borderline, Highly uncertain
            """
        )

    st.divider()

    st.subheader("How to cite")
    st.markdown(
        "If you use this platform or refer to its results, please cite the "
        "paper:"
    )
    st.code(
        """@unpublished{toni2026convergence,
  author = {Toni, Emiliano},
  title  = {The Convergence Frontier: Benchmarking Growth Momentum
            and Measuring its Uncertainty},
  year   = {2026},
  note   = {Working paper, University of St.~Gallen},
  url    = {https://github.com/e-toni/Convergence-Platform}
}""",
        language="bibtex",
    )

    st.subheader("Code and replication")
    st.markdown(
        """
All source code --- the R pipeline that produces the master workbook
(`DCT_GMC_master_pwt_v7.xlsx`), the figure-generation scripts, the LaTeX
paper, and this Streamlit dashboard --- is publicly available on GitHub:

🔗 **[github.com/e-toni/Convergence-Frontier](https://github.com/e-toni/Convergence-Frontier)**

Every number in the paper and every chart in this platform can be
reproduced by running the R pipeline against the public Penn World Table
release.
        """
    )

    st.divider()

    st.subheader("Acknowledgements and intellectual debts")
    st.markdown(
        """
The Penn World Table is maintained by the Groningen Growth and
Development Centre and described in Feenstra, Inklaar and Timmer (2015).
The dashboard is built with Streamlit and Plotly.
        """
    )

    with st.expander("License, disclaimers, and use"):
        st.markdown(
            """
This platform is provided for academic and policy research purposes.
Country classifications carry the seven-method uncertainty range reported
in the dashboard; readers are encouraged to consult the robustness flag
and the per-filter classifications before drawing conclusions about any
single economy. The platform's purpose is not to deliver verdicts, but to
make every verdict checkable.

Penn World Table data are subject to the terms of the original source.
The platform's own source code is released under the MIT license.

**Declaration of interests:** none.
            """
        )

    st.stop()

st.title("🌍 The Convergence Frontier")
st.caption(
    "Each country's latest growth-income position, benchmarked against the "
    "convergence frontier and stress-tested across a seven-method filter "
    "battery. The dashboard reads the GMC master workbook; all estimation is "
    "done in R."
)

workbook_path = find_workbook()
if workbook_path is None:
    st.error(
        f"Could not find `{DEFAULT_WORKBOOK}`. Place the v7 GMC master "
        "workbook in the same folder as `app.py`."
    )
    st.stop()

try:
    data = load_workbook(workbook_path)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not load `{os.path.basename(workbook_path)}`.")
    st.exception(exc)
    st.stop()

labels = method_label_map(data["legend"])

# ---- sidebar -------------------------------------------------------------
with st.sidebar:
    st.header("Income concept")
    measures = sorted(data["classification"]["gdp_measure"].dropna().unique())
    default_measure = "rgdpo" if "rgdpo" in measures else measures[0]
    measure_id = st.selectbox(
        "GDP per capita concept", options=measures,
        index=measures.index(default_measure),
        format_func=measure_label,
    )

    st.header("Trend filter")
    method_options = [ENSEMBLE_KEY] + METHOD_IDS
    method_key = st.selectbox(
        "Smoothing method",
        options=method_options,
        index=0,
        format_func=lambda k: (ENSEMBLE_LABEL if k == ENSEMBLE_KEY
                                else labels.get(k, k)),
        help="Pick one filter, or the mean ensemble of all seven.",
    )
    method_label_text = (ENSEMBLE_LABEL if method_key == ENSEMBLE_KEY
                         else labels.get(method_key, method_key))

    st.header("Chart controls")
    show_band   = st.checkbox("Show 25-75 percentile band", value=True)
    show_median = st.checkbox("Show median frontier", value=True)
    show_bars   = st.checkbox("Show seven-filter uncertainty bars",
                              value=False)
    colour_mode = st.radio(
        "Colour the datapoints by",
        options=["Plain", "Robustness", "GMC class"],
        index=0,
    )

# ---- prepare data for the selected measure/method ------------------------
x_grid = np.linspace(max(500.0, X_LIMITS[0]), X_LIMITS[1], 400)
frontier = selected_frontier(data, measure_id, method_key, x_grid)
pts      = latest_points(data, measure_id, method_key)
traj     = trajectory_series(data, measure_id, method_key)

# ---- sidebar: country selection ------------------------------------------
with st.sidebar:
    st.header("Country selection")

    classes_present = [c for c in CLASS_ORDER
                       if c in set(pts["gmc_class"].dropna())]
    sel_classes = st.multiselect(
        "Filter latest datapoints by GMC class",
        options=classes_present, default=classes_present,
    )

    all_countries = sorted(pts["Country Name"].dropna().unique())
    default_sel = [c for c in DEFAULT_COUNTRIES if c in all_countries]
    sel_latest = st.multiselect(
        "Countries - latest datapoints",
        options=all_countries, default=default_sel,
        help="Labelled dots. Clear to hide all latest datapoints.",
    )

    traj_countries = sorted(traj["Country Name"].dropna().unique())
    sel_traj = st.multiselect(
        "Countries - full trajectories",
        options=traj_countries, default=[],
        help="Full historical paths in the growth-income plane.",
    )

# ---- build chart ---------------------------------------------------------
fig, pts_plot, traj_sel = build_chart(
    frontier, pts, traj, sel_latest, sel_traj, sel_classes,
    show_band, show_median, show_bars, colour_mode,
    measure_id, method_label_text,
)

# ---- headline metrics ----------------------------------------------------
m = st.columns(4)
m[0].metric("GDP concept", measure_id)
m[1].metric("Datapoints shown", f"{len(pts_plot):,}")
m[2].metric("Trajectories shown", f"{len(sel_traj):,}")
if len(traj):
    m[3].metric("Trajectory years",
                f"{int(traj['year'].min())}-{int(traj['year'].max())}")
else:
    m[3].metric("Trajectory years", "n/a")

st.plotly_chart(fig, use_container_width=True)

# ---- figure download buttons ---------------------------------------------
# Two side-by-side options: PDF (needs the kaleido package) and HTML (no
# extra dependency). Both export the exact figure shown above, with every
# trace colour set explicitly so kaleido renders the same lines, markers
# and symbols that the dashboard shows.
#
# Streamlit cannot tell whether the user is viewing the dashboard in Light
# or Dark mode, so we expose a small toggle here. The toggle affects only
# the BACKGROUND of the exported file -- trace colours are explicit and
# the same in both themes -- so the user can match the export to whatever
# they see on their screen.
export_theme = st.radio(
    "PDF / HTML background",
    options=["Light", "Dark"],
    index=0,
    horizontal=True,
    help=("Choose the background of the downloaded file. Use Light if your "
          "dashboard is in light mode, Dark if it is in dark mode. "
          "Trace colours, markers and labels are identical in both."),
)

dl_cols = st.columns([1, 1, 4])

fig_filename_stem = (
    f"convergence_frontier_{safe_filename(measure_id)}"
    f"_{safe_filename(method_label_text)}"
    f"_{export_theme.lower()}"
)

# Force a sensible export aspect ratio (matches the R script figures).
export_width, export_height = 1200, 750

# Build a separate export figure so changing the theme does not mutate the
# displayed chart. go.Figure(fig.to_dict()) is a cheap deep-copy that
# preserves every trace and layout property.
export_fig = go.Figure(fig.to_dict())
if export_theme == "Dark":
    export_fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="#0e1117",   # Streamlit dark default
        plot_bgcolor="#262730",
        font=dict(color="#fafafa"),
    )
else:
    export_fig.update_layout(
        template="plotly_white",
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(color="#222222"),
    )

pdf_bytes = None
pdf_error = None
try:
    pdf_bytes = export_fig.to_image(
        format="pdf",
        width=export_width, height=export_height, scale=2,
    )
except Exception as exc:  # noqa: BLE001
    pdf_error = str(exc)

with dl_cols[0]:
    if pdf_bytes is not None:
        st.download_button(
            "Download figure as PDF",
            pdf_bytes,
            file_name=f"{fig_filename_stem}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    else:
        st.button(
            "PDF unavailable (install kaleido)",
            disabled=True, use_container_width=True,
            help=("PDF export requires the kaleido package. Install it with "
                  "`pip install kaleido==0.2.1`, then restart the app.\n\n"
                  f"Error: {pdf_error}" if pdf_error else
                  "PDF export requires `pip install kaleido==0.2.1`."),
        )

with dl_cols[1]:
    html_bytes = export_fig.to_html(
        include_plotlyjs="cdn", full_html=True
    ).encode("utf-8")
    st.download_button(
        "Download figure as HTML",
        html_bytes,
        file_name=f"{fig_filename_stem}.html",
        mime="text/html",
        use_container_width=True,
    )

# ============================================================
# TABLES
# ============================================================

tab1, tab2, tab3, tab4 = st.tabs(
    ["Latest datapoints", "Per-filter classification",
     "Robustness summary", "Notes"]
)

with tab1:
    st.subheader(f"Latest country datapoints - {measure_id}")
    cols = [c for c in ["Country Name", "Country Code", "year", "income_pc",
                        "growth", "gmc_class", "growth_min", "growth_max",
                        "robustness", "in_estimation_sample"]
            if c in pts_plot.columns]
    show = pts_plot[cols].rename(columns={
        "growth": "trend_growth", "growth_min": "range_lo",
        "growth_max": "range_hi"})
    st.dataframe(show.sort_values(["gmc_class", "Country Name"]),
                 use_container_width=True, hide_index=True)
    st.download_button(
        "Download as CSV", show.to_csv(index=False),
        file_name=f"convergence_{safe_filename(measure_id)}_latest.csv",
        mime="text/csv")

with tab2:
    st.subheader("GMC class under every filter (selected countries)")
    cl = data["classification"]
    cl = cl[(cl["gdp_measure"] == measure_id)
            & (cl["Country Name"].isin(sel_latest))].copy()
    per_filter_cols = (["Country Name"]
                       + [f"cls_{mid}" for mid in METHOD_IDS]
                       + ["class_ensemble_mean", "robustness"])
    per_filter_cols = [c for c in per_filter_cols if c in cl.columns]
    # Use short, unique labels (HP-6.25, HP-100, Hamilton, LLT, CF, BN, MA-5)
    # for the column headers so the table has unique column names.
    rename = {f"cls_{mid}": METHOD_SHORT_LABELS.get(mid, mid)
              for mid in METHOD_IDS}
    rename["class_ensemble_mean"] = "Ensemble (mean)"
    rename["robustness"]          = "Robustness"
    tbl = cl[per_filter_cols].rename(columns=rename)
    if "income_pc" in cl.columns:
        tbl = tbl.assign(_ord=cl["income_pc"].values).sort_values(
            "_ord").drop(columns="_ord")
    st.dataframe(tbl, use_container_width=True, hide_index=True)
    st.caption(
        "Each cell is the Growth-Momentum class under one trend filter. "
        "Disagreement across the row is exactly what the robustness flag "
        "summarises."
    )
    st.download_button(
        "Download as CSV", tbl.to_csv(index=False),
        file_name=f"convergence_{safe_filename(measure_id)}_per_filter.csv",
        mime="text/csv")

with tab3:
    st.subheader("Robustness of the classification")
    rs = data["robustness_summary"]
    if rs is not None:
        st.dataframe(rs, use_container_width=True, hide_index=True)
        st.caption(
            "A country is Robust when all seven filters agree on its class, "
            "Borderline when they straddle one boundary, and Highly uncertain "
            "when they span all three classes."
        )
    else:
        st.info("GMC_robustness_summary sheet not found in the workbook.")

    # robustness composition for the current measure
    if "robustness" in pts.columns and len(pts):
        comp = (pts["robustness"].value_counts()
                .reindex(ROBUST_ORDER).fillna(0).astype(int))
        st.bar_chart(comp)

with tab4:
    st.markdown(
        f"""
        **What this dashboard shows**

        - Workbook: `{os.path.basename(workbook_path)}` (produced by
          `Run_GMC_pwt_results_v7.R`).
        - The **convergence frontier** is the set of 25th / 50th / 75th
          quantile-regression curves of trend growth on log income,
          estimated on the converged-economy sample.
        - **Income concept**: `{measure_id}` - {measure_label(measure_id)}.
        - **Trend filter**: {method_label_text}.
        - A country is an **Accelerator** above the 75th percentile, **On
          Track** inside the 25-75 band, and **Stagnating** below the 25th.
        - The **seven-filter battery** is HP-6.25, HP-100, Hamilton (2018),
          local linear trend, Christiano-Fitzgerald band-pass,
          Beveridge-Nelson and MA-5. The **Ensemble** option uses the mean
          across all seven.
        - The **uncertainty bar** spans the minimum-to-maximum trend growth
          across the seven filters; it is a robustness range across
          estimators, **not** a statistical confidence interval.
        - The **robustness flag** records whether the seven filters agree on
          the class (Robust), straddle one boundary (Borderline) or span all
          three classes (Highly uncertain).
        - The dashboard performs **no smoothing**; every series is read from
          the workbook, which is produced entirely in R.

        **Income concepts available:** rgdpo, rgdpna, rgdpe, cgdpe, cgdpo.
        """
    )
