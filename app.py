import re
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="Dynamic Country Tiers Dashboard",
    page_icon="🌍",
    layout="wide",
)


# ============================================================
# DEFAULT COUNTRY SELECTION
# ============================================================

DEFAULT_LATEST_COUNTRIES = [
    "Bolivia",
    "Brazil",
    "Bulgaria",
    "Burundi",
    "Chile",
    "China",
    "Croatia",
    "Denmark",
    "European Union",
    "Finland",
    "France",
    "Germany",
    "Hong Kong SAR, China",
    "India",
    "Ireland",
    "Japan",
    "Korea, Rep.",
    "Kuwait",
    "Lithuania",
    "Luxembourg",
    "Norway",
    "OECD members",
    "Poland",
    "Portugal",
    "Singapore",
    "Spain",
    "Sweden",
    "Switzerland",
    "United Kingdom",
    "United States",
    "Uruguay",
]


# ============================================================
# DATASET CONFIGURATION
# ============================================================

DATASET_CONFIG = {
    "WB - World Bank indicators": {
        "default_file": "DCT_output_wb_multi.xlsx",
        "fit_sheet": "WBpc_fit_IQR",
        "fit_x_col": "IncomePC",
        "latest_sheet": "Latest_country_data",
        "trajectory_sheet": "Full_country_trajectory",
        "income_col": "income_pc",
        "growth_col": "income_pc_growth",
        "default_measure_id": "gni_atlas_current",
        "default_measure_label": "GNI per capita, Atlas method, current US$",
        "measure_col": "wb_measure",
        "measure_label_col": "wb_measure_label",
        "measure_col_candidates": [
            "wb_measure",
            "income_measure",
            "gdp_measure",
            "gni_measure",
            "measure",
        ],
        "measure_label_col_candidates": [
            "wb_measure_label",
            "income_measure_label",
            "gdp_measure_label",
            "gni_measure_label",
            "measure_label",
        ],
        "income_label_template": "{measure_label}",
        "growth_label_template": "HP-smoothed growth, YoY (%)",
        "raw_growth_label_template": "Raw growth",
        "chart_title_template": "Dynamic Country Tiers: WB {measure_label}",
        "download_prefix": "wb_income_pc",
        "x_axis_range": None,
    },
    "PWT - GDP per capita alternatives": {
        "default_file": "DCT_output_pwt_multi.xlsx",
        "fit_sheet": "GDPpc_fit_IQR",
        "fit_x_col": "GDPpc",
        "latest_sheet": "Latest_country_data",
        "trajectory_sheet": "Full_country_trajectory",
        "income_col": "gdp_pc",
        "growth_col": "gdp_pc_growth",
        "default_measure_id": "rgdpna",
        "default_measure_label": "Real GDP at constant national prices",
        "measure_col": "gdp_measure",
        "measure_label_col": "gdp_measure_label",
        "measure_col_candidates": [
            "gdp_measure",
            "income_measure",
            "measure",
        ],
        "measure_label_col_candidates": [
            "gdp_measure_label",
            "income_measure_label",
            "measure_label",
        ],
        "income_label_template": "GDP per capita: {measure_id} / pop",
        "growth_label_template": "HP-smoothed GDP per capita growth, YoY (%)",
        "raw_growth_label_template": "Raw GDP per capita growth",
        "chart_title_template": "Dynamic Country Tiers: PWT {measure_id} per capita",
        "download_prefix": "pwt_gdp_pc",
        "x_axis_range": [0, 150000],
    },
}


# ============================================================
# GENERAL HELPERS
# ============================================================

def safe_filename(text):
    """Create a safe file-name fragment from a string."""
    text = str(text).lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def safe_key(text):
    """Create a safe Streamlit widget key fragment."""
    return safe_filename(text)


def standardize_measure_columns(df, cfg):
    """
    Standardize measure columns.

    For WB multi-output, this reads wb_measure and wb_measure_label.
    For PWT multi-output, this reads gdp_measure and gdp_measure_label.
    """
    df = df.copy()

    measure_col = cfg.get("measure_col")
    measure_label_col = cfg.get("measure_label_col")

    # Standardize measure ID.
    if measure_col is not None and measure_col in df.columns:
        df = df.rename(columns={measure_col: "measure_id"})
    else:
        found = False

        for candidate in cfg.get("measure_col_candidates", []):
            if candidate in df.columns:
                df = df.rename(columns={candidate: "measure_id"})
                found = True
                break

        if not found:
            df["measure_id"] = cfg["default_measure_id"]

    # Standardize measure label.
    if measure_label_col is not None and measure_label_col in df.columns:
        df = df.rename(columns={measure_label_col: "measure_label"})
    else:
        found = False

        for candidate in cfg.get("measure_label_col_candidates", []):
            if candidate in df.columns:
                df = df.rename(columns={candidate: "measure_label"})
                found = True
                break

        if not found:
            df["measure_label"] = cfg["default_measure_label"]

    df["measure_id"] = df["measure_id"].astype(str)
    df["measure_label"] = df["measure_label"].astype(str)

    return df


def get_measure_options(fit, latest, trajectory):
    """Return measures that are available in all required sheets."""
    fit_measures = set(fit["measure_id"].dropna().unique())
    latest_measures = set(latest["measure_id"].dropna().unique())
    trajectory_measures = set(trajectory["measure_id"].dropna().unique())

    common_measures = sorted(fit_measures & latest_measures & trajectory_measures)

    if len(common_measures) == 0:
        common_measures = sorted(fit_measures | latest_measures | trajectory_measures)

    label_df = pd.concat(
        [
            fit[["measure_id", "measure_label"]],
            latest[["measure_id", "measure_label"]],
            trajectory[["measure_id", "measure_label"]],
        ],
        ignore_index=True,
    ).drop_duplicates()

    label_map = (
        label_df.drop_duplicates("measure_id")
        .set_index("measure_id")["measure_label"]
        .to_dict()
    )

    return common_measures, label_map


def format_measure_option(measure_id, label_map):
    """Format measure options shown in the sidebar."""
    label = label_map.get(measure_id, measure_id)

    if label == measure_id:
        return measure_id

    return f"{measure_id} — {label}"


def default_latest_countries_available(all_countries):
    """Return default latest countries that exist in the selected dataset."""
    available = set(all_countries)
    return [country for country in DEFAULT_LATEST_COUNTRIES if country in available]


def available_columns(df, columns):
    """Return only columns that exist in the DataFrame."""
    return [col for col in columns if col in df.columns]


# ============================================================
# DATA LOADING
# ============================================================

@st.cache_data
def load_dct_output(file, dataset_name):
    """
    Load either the WB multi-measure or PWT multi-measure DCT workbook.

    The function standardizes variable names internally so the rest of
    the dashboard can treat WB and PWT results in the same way.
    """
    cfg = DATASET_CONFIG[dataset_name]

    xls = pd.ExcelFile(file)

    required_sheets = {
        cfg["fit_sheet"],
        cfg["latest_sheet"],
        cfg["trajectory_sheet"],
    }

    available_sheets = set(xls.sheet_names)
    missing_sheets = required_sheets - available_sheets

    if missing_sheets:
        missing = ", ".join(sorted(missing_sheets))
        raise ValueError(
            f"Missing required sheet(s): {missing}. "
            f"Make sure the workbook corresponds to the selected dataset: {dataset_name}."
        )

    fit = pd.read_excel(xls, cfg["fit_sheet"])
    latest = pd.read_excel(xls, cfg["latest_sheet"])
    trajectory = pd.read_excel(xls, cfg["trajectory_sheet"])

    # Standardize measure columns before filtering.
    fit = standardize_measure_columns(fit, cfg)
    latest = standardize_measure_columns(latest, cfg)
    trajectory = standardize_measure_columns(trajectory, cfg)

    # Rename income columns to common internal names.
    fit = fit.rename(
        columns={
            cfg["fit_x_col"]: "income_pc",
        }
    )

    latest = latest.rename(
        columns={
            cfg["income_col"]: "income_pc",
            cfg["growth_col"]: "income_pc_growth",
        }
    )

    trajectory = trajectory.rename(
        columns={
            cfg["income_col"]: "income_pc",
            cfg["growth_col"]: "income_pc_growth",
        }
    )

    # Validate required columns after renaming.
    required_fit_cols = [
        "measure_id",
        "measure_label",
        "income_pc",
        "Lower_25th_Percentile",
        "Median_Fit_Smoothed_Growth",
        "Upper_75th_Percentile",
    ]

    required_data_cols = [
        "measure_id",
        "measure_label",
        "year",
        "Country Name",
        "Country Code",
        "income_pc",
        "income_pc_growth",
        "smoothed_growth",
    ]

    missing_fit_cols = [col for col in required_fit_cols if col not in fit.columns]
    missing_latest_cols = [col for col in required_data_cols if col not in latest.columns]
    missing_trajectory_cols = [col for col in required_data_cols if col not in trajectory.columns]

    if missing_fit_cols:
        raise ValueError(
            f"Fit sheet is missing columns after standardization: {missing_fit_cols}"
        )

    if missing_latest_cols:
        raise ValueError(
            f"Latest_country_data is missing columns after standardization: {missing_latest_cols}"
        )

    if missing_trajectory_cols:
        raise ValueError(
            f"Full_country_trajectory is missing columns after standardization: {missing_trajectory_cols}"
        )

    # Standardize numeric columns.
    for col in [
        "income_pc",
        "Lower_25th_Percentile",
        "Median_Fit_Smoothed_Growth",
        "Upper_75th_Percentile",
    ]:
        fit[col] = pd.to_numeric(fit[col], errors="coerce")

    for col in ["year", "smoothed_growth", "income_pc", "income_pc_growth"]:
        latest[col] = pd.to_numeric(latest[col], errors="coerce")

    for col in ["year", "income_pc", "income_pc_growth", "smoothed_growth"]:
        trajectory[col] = pd.to_numeric(trajectory[col], errors="coerce")

    # Keep only usable observations.
    fit = fit.dropna(
        subset=[
            "measure_id",
            "income_pc",
            "Lower_25th_Percentile",
            "Median_Fit_Smoothed_Growth",
            "Upper_75th_Percentile",
        ]
    ).sort_values(["measure_id", "income_pc"])

    latest = latest.dropna(
        subset=[
            "measure_id",
            "income_pc",
            "smoothed_growth",
            "Country Name",
            "Country Code",
            "year",
        ]
    )

    trajectory = trajectory.dropna(
        subset=[
            "measure_id",
            "income_pc",
            "smoothed_growth",
            "Country Name",
            "Country Code",
            "year",
        ]
    )

    return fit, latest, trajectory


# ============================================================
# PLOT HELPERS
# ============================================================

def add_fit_curves(fig, fit, show_iqr=True, show_median=True):
    """Add fitted DCT quantile curves to the chart."""
    if show_iqr:
        fig.add_trace(
            go.Scatter(
                x=fit["income_pc"],
                y=fit["Upper_75th_Percentile"],
                mode="lines",
                name="75th percentile fit",
                line=dict(dash="dash", width=2),
                hovertemplate=(
                    "Income per capita: %{x:,.0f}"
                    "<br>Upper fit: %{y:.2f}%"
                    "<extra></extra>"
                ),
            )
        )

        fig.add_trace(
            go.Scatter(
                x=fit["income_pc"],
                y=fit["Lower_25th_Percentile"],
                mode="lines",
                name="25th percentile fit",
                line=dict(dash="dash", width=2),
                fill="tonexty",
                fillcolor="rgba(150,150,150,0.12)",
                hovertemplate=(
                    "Income per capita: %{x:,.0f}"
                    "<br>Lower fit: %{y:.2f}%"
                    "<extra></extra>"
                ),
            )
        )

    if show_median:
        fig.add_trace(
            go.Scatter(
                x=fit["income_pc"],
                y=fit["Median_Fit_Smoothed_Growth"],
                mode="lines",
                name="Median fit",
                line=dict(width=3),
                hovertemplate=(
                    "Income per capita: %{x:,.0f}"
                    "<br>Median fit: %{y:.2f}%"
                    "<extra></extra>"
                ),
            )
        )

    return fig


def classify_latest_points(latest: pd.DataFrame, fit: pd.DataFrame) -> pd.DataFrame:
    """Classify latest country points relative to the fitted IQR curves."""
    d = latest.copy()

    fit = fit.sort_values("income_pc").copy()

    fit_x = fit["income_pc"].to_numpy()
    lower = fit["Lower_25th_Percentile"].to_numpy()
    median = fit["Median_Fit_Smoothed_Growth"].to_numpy()
    upper = fit["Upper_75th_Percentile"].to_numpy()

    d["lower_fit"] = np.interp(d["income_pc"], fit_x, lower)
    d["median_fit"] = np.interp(d["income_pc"], fit_x, median)
    d["upper_fit"] = np.interp(d["income_pc"], fit_x, upper)

    conditions = [
        d["smoothed_growth"] > d["upper_fit"],
        d["smoothed_growth"] < d["lower_fit"],
    ]

    choices = [
        "Above 75th percentile",
        "Below 25th percentile",
    ]

    d["DCT_position"] = np.select(
        conditions,
        choices,
        default="Inside IQR",
    )

    return d


def add_latest_points(fig, d, raw_growth_label, marker_size=9):
    """Add labelled latest country datapoints."""
    if d.empty:
        return fig

    hovertemplate = (
        "<b>%{customdata[0]}</b> (%{customdata[1]})"
        "<br>Year: %{customdata[2]}"
        "<br>Income per capita: %{x:,.0f}"
        "<br>HP-smoothed growth: %{y:.2f}%"
        f"<br>{raw_growth_label}: "
        "%{customdata[3]}%"
        "<br>Position: %{customdata[4]}"
        "<extra></extra>"
    )

    fig.add_trace(
        go.Scatter(
            x=d["income_pc"],
            y=d["smoothed_growth"],
            mode="markers+text",
            name="Latest country datapoint",
            text=d["Country Name"],
            textposition="middle right",
            textfont=dict(size=12),
            marker=dict(
                size=marker_size,
                opacity=0.78,
                line=dict(width=0.6),
            ),
            customdata=np.column_stack(
                [
                    d["Country Name"].astype(str),
                    d["Country Code"].astype(str),
                    d["year"].astype(int).astype(str),
                    d["income_pc_growth"].round(2).astype(str),
                    d["DCT_position"].astype(str),
                ]
            ),
            hovertemplate=hovertemplate,
        )
    )

    return fig


def add_country_trajectories(fig, trajectory, selected_countries, raw_growth_label):
    """
    Add full HP-smoothed country trajectories.

    The country name is shown only next to the final available datapoint
    of each trajectory to avoid cluttering the figure.
    """
    hovertemplate = (
        "<b>%{customdata[0]}</b> (%{customdata[1]})"
        "<br>Year: %{customdata[2]}"
        "<br>Income per capita: %{x:,.0f}"
        "<br>HP-smoothed growth: %{y:.2f}%"
        f"<br>{raw_growth_label}: "
        "%{customdata[3]}%"
        "<extra></extra>"
    )

    for country in selected_countries:
        dc = trajectory[trajectory["Country Name"] == country].sort_values("year")

        if dc.empty:
            continue

        final_point_label = [""] * (len(dc) - 1) + [country]

        fig.add_trace(
            go.Scatter(
                x=dc["income_pc"],
                y=dc["smoothed_growth"],
                mode="lines+markers+text",
                name=f"{country} trajectory",
                text=final_point_label,
                textposition="middle right",
                textfont=dict(size=12),
                marker=dict(size=6),
                line=dict(width=2),
                customdata=np.column_stack(
                    [
                        dc["Country Name"].astype(str),
                        dc["Country Code"].astype(str),
                        dc["year"].astype(int).astype(str),
                        dc["income_pc_growth"].round(2).astype(str),
                    ]
                ),
                hovertemplate=hovertemplate,
            )
        )

    return fig


def build_chart(
    fit,
    latest,
    trajectory,
    selected_latest_countries,
    selected_trajectory_countries,
    selected_positions,
    show_iqr,
    show_median,
    show_thresholds,
    cfg,
    measure_id,
    measure_label,
):
    """Build the main Plotly figure."""
    fig = go.Figure()

    income_label = cfg["income_label_template"].format(
        measure_id=measure_id,
        measure_label=measure_label,
    )

    growth_label = cfg["growth_label_template"].format(
        measure_id=measure_id,
        measure_label=measure_label,
    )

    raw_growth_label = cfg["raw_growth_label_template"].format(
        measure_id=measure_id,
        measure_label=measure_label,
    )

    chart_title = cfg["chart_title_template"].format(
        measure_id=measure_id,
        measure_label=measure_label,
    )

    fig = add_fit_curves(
        fig,
        fit,
        show_iqr=show_iqr,
        show_median=show_median,
    )

    latest_plot = latest[
        latest["Country Name"].isin(selected_latest_countries)
        & latest["DCT_position"].isin(selected_positions)
    ].copy()

    trajectory_plot = trajectory[
        trajectory["Country Name"].isin(selected_trajectory_countries)
    ].copy()

    if len(selected_latest_countries) > 0:
        fig = add_latest_points(
            fig,
            latest_plot,
            raw_growth_label=raw_growth_label,
        )

    if len(selected_trajectory_countries) > 0:
        fig = add_country_trajectories(
            fig,
            trajectory_plot,
            selected_trajectory_countries,
            raw_growth_label=raw_growth_label,
        )

    if show_thresholds:
        avg_income = (
            latest_plot["income_pc"].mean()
            if len(latest_plot)
            else latest["income_pc"].mean()
        )

        for multiplier in [0.25, 0.5, 1, 2, 4]:
            x = avg_income * multiplier
            label = "Average" if multiplier == 1 else f"{multiplier:g}× average"

            fig.add_vline(
                x=x,
                line_dash="dot",
                line_width=1.5,
                line_color="white",
                annotation_text=label,
                annotation_position="top",
                annotation_font_color="white",
            )

    fig.add_hline(
        y=0,
        line_dash="dot",
        line_width=1,
    )

    fig.update_layout(
        title=chart_title,
        xaxis_title=income_label,
        yaxis_title=growth_label,
        hovermode="closest",
        height=720,
        margin=dict(l=40, r=30, t=70, b=45),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.23,
            xanchor="left",
            x=0,
        ),
    )

    if cfg.get("x_axis_range") is not None:
        fig.update_xaxes(
            tickformat=",",
            range=cfg["x_axis_range"],
        )
    else:
        fig.update_xaxes(
            tickformat=",",
        )

    fig.update_yaxes(zeroline=True)

    return fig, latest_plot, trajectory_plot


# ============================================================
# APP BODY
# ============================================================

st.title("🌍 Dynamic Country Tiers Dashboard")

st.caption(
    "Interactive DCT result plot using either the World Bank database "
    "or the Penn World Table database with alternative income and GDP per capita concepts."
)


# ============================================================
# SIDEBAR: DATASET SELECTION
# ============================================================

with st.sidebar:
    st.header("Dataset")

    dataset_name = st.radio(
        "Choose data source",
        options=list(DATASET_CONFIG.keys()),
        index=0,
    )

    cfg = DATASET_CONFIG[dataset_name]
    source = cfg["default_file"]


# ============================================================
# LOAD DATA
# ============================================================

try:
    fit_all, latest_all, trajectory_all = load_dct_output(
        source,
        dataset_name,
    )

except Exception as exc:
    st.error(
        f"Could not load the required workbook for `{dataset_name}`. "
        f"Make sure `{cfg['default_file']}` is in the same folder as `app.py`."
    )
    st.exception(exc)
    st.stop()


# ============================================================
# SIDEBAR: MEASURE SELECTION
# ============================================================

measure_options, measure_label_map = get_measure_options(
    fit_all,
    latest_all,
    trajectory_all,
)

if len(measure_options) == 0:
    st.error("No income/GDP measure could be detected in the selected workbook.")
    st.stop()

default_measure = cfg.get("default_measure_id", measure_options[0])
default_index = measure_options.index(default_measure) if default_measure in measure_options else 0

with st.sidebar:
    st.header("Income / GDP measure")

    selected_measure = st.selectbox(
        "Choose income/GDP concept",
        options=measure_options,
        index=default_index,
        format_func=lambda x: format_measure_option(
            x,
            measure_label_map,
        ),
    )

selected_measure_label = measure_label_map.get(
    selected_measure,
    selected_measure,
)


# Filter selected measure.
fit = fit_all[fit_all["measure_id"] == selected_measure].copy()
latest_raw = latest_all[latest_all["measure_id"] == selected_measure].copy()
trajectory = trajectory_all[trajectory_all["measure_id"] == selected_measure].copy()

if fit.empty or latest_raw.empty or trajectory.empty:
    st.error(
        f"The selected measure `{selected_measure}` does not have complete data "
        "across fit, latest datapoints, and trajectory sheets."
    )
    st.stop()


latest = classify_latest_points(
    latest_raw,
    fit,
)


# ============================================================
# SIDEBAR: CHART CONTROLS
# ============================================================

with st.sidebar:
    st.header("Chart controls")

    show_iqr = st.checkbox(
        "Show IQR band",
        value=True,
    )

    show_median = st.checkbox(
        "Show median fit",
        value=True,
    )

    show_thresholds = st.checkbox(
        "Show average-income reference lines",
        value=False,
    )


# ============================================================
# SIDEBAR: COUNTRY SELECTION
# ============================================================

country_widget_suffix = safe_key(f"{dataset_name}_{selected_measure}")

with st.sidebar:
    st.header("Country selection")

    positions = sorted(latest["DCT_position"].dropna().unique())

    selected_positions = st.multiselect(
        "DCT position for latest datapoints",
        options=positions,
        default=positions,
        help="This filter applies only to latest datapoints, not trajectories.",
        key=f"positions_{country_widget_suffix}",
    )

    all_latest_countries = sorted(
        latest["Country Name"].dropna().unique()
    )

    default_latest_countries = default_latest_countries_available(
        all_latest_countries
    )

    selected_latest_countries = st.multiselect(
        "Countries for latest datapoints",
        options=all_latest_countries,
        default=default_latest_countries,
        help=(
            "Select countries whose latest datapoints should appear as labelled dots. "
            "Leaving this empty hides latest datapoints."
        ),
        key=f"latest_countries_{country_widget_suffix}",
    )

    all_trajectory_countries = sorted(
        trajectory["Country Name"].dropna().unique()
    )

    selected_trajectory_countries = st.multiselect(
        "Countries for trajectories",
        options=all_trajectory_countries,
        default=[],
        help=(
            "Select countries whose full HP-smoothed trajectories should appear. "
            "Leaving this empty hides trajectories."
        ),
        key=f"trajectory_countries_{country_widget_suffix}",
    )


# ============================================================
# BUILD AND SHOW CHART
# ============================================================

fig, latest_plot, trajectory_plot = build_chart(
    fit=fit,
    latest=latest,
    trajectory=trajectory,
    selected_latest_countries=selected_latest_countries,
    selected_trajectory_countries=selected_trajectory_countries,
    selected_positions=selected_positions,
    show_iqr=show_iqr,
    show_median=show_median,
    show_thresholds=show_thresholds,
    cfg=cfg,
    measure_id=selected_measure,
    measure_label=selected_measure_label,
)


top = st.columns(4)

top[0].metric(
    "Dataset",
    "WB" if dataset_name.startswith("WB") else "PWT",
)

top[1].metric(
    "Latest datapoints shown",
    f"{len(latest_plot):,}",
)

top[2].metric(
    "Trajectory observations shown",
    f"{len(trajectory_plot):,}",
)

top[3].metric(
    "Trajectory years",
    f"{int(trajectory['year'].min())}–{int(trajectory['year'].max())}",
)


st.plotly_chart(
    fig,
    use_container_width=True,
)


# ============================================================
# TABLES
# ============================================================

tab1, tab2, tab3, tab4 = st.tabs(
    [
        "Latest datapoints",
        "Trajectory data",
        "Fit data",
        "Notes",
    ]
)


with tab1:
    st.subheader("Latest country datapoints")

    latest_cols = available_columns(
        latest_plot,
        [
            "measure_id",
            "measure_label",
            "indicator_code",
            "year",
            "Country Name",
            "Country Code",
            "income_pc",
            "income_pc_growth",
            "smoothed_growth",
            "lower_fit",
            "median_fit",
            "upper_fit",
            "DCT_position",
        ],
    )

    st.dataframe(
        latest_plot[latest_cols].sort_values(
            [
                "DCT_position",
                "Country Name",
            ]
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Download latest datapoints as CSV",
        latest_plot[latest_cols].to_csv(index=False),
        file_name=(
            f"{cfg['download_prefix']}_"
            f"{safe_filename(selected_measure)}_"
            "latest_datapoints_filtered.csv"
        ),
        mime="text/csv",
    )


with tab2:
    st.subheader("Full HP-smoothed trajectory data")

    trajectory_cols = available_columns(
        trajectory_plot,
        [
            "measure_id",
            "measure_label",
            "indicator_code",
            "year",
            "Country Name",
            "Country Code",
            "income_pc",
            "income_pc_growth",
            "smoothed_growth",
            "max_gni",
            "average_max_gni",
            "max_income_pc",
            "average_max_income_pc",
            "max_gdp_pc",
            "average_max_gdp_pc",
            "in_dct_estimation_sample",
            "gdp_level",
            "rgdpna",
            "pop",
        ],
    )

    st.dataframe(
        trajectory_plot[trajectory_cols].sort_values(
            [
                "Country Name",
                "year",
            ]
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Download trajectory data as CSV",
        trajectory_plot[trajectory_cols].to_csv(index=False),
        file_name=(
            f"{cfg['download_prefix']}_"
            f"{safe_filename(selected_measure)}_"
            "trajectory_filtered.csv"
        ),
        mime="text/csv",
    )


with tab3:
    st.subheader("Fitted DCT curves")

    fit_cols = available_columns(
        fit,
        [
            "measure_id",
            "measure_label",
            "indicator_code",
            "income_pc",
            "Lower_25th_Percentile",
            "Median_Fit_Smoothed_Growth",
            "Upper_75th_Percentile",
        ],
    )

    st.dataframe(
        fit[fit_cols],
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Download fitted curves as CSV",
        fit[fit_cols].to_csv(index=False),
        file_name=(
            f"{cfg['download_prefix']}_"
            f"{safe_filename(selected_measure)}_"
            "fit_iqr.csv"
        ),
        mime="text/csv",
    )


with tab4:
    income_label = cfg["income_label_template"].format(
        measure_id=selected_measure,
        measure_label=selected_measure_label,
    )

    growth_label = cfg["growth_label_template"].format(
        measure_id=selected_measure,
        measure_label=selected_measure_label,
    )

    st.markdown(
        f"""
        **How to read the chart**

        - The selected dataset is **{dataset_name}**.
        - The selected income/GDP measure is **{selected_measure}**.
        - The measure label is **{selected_measure_label}**.
        - The x-axis is **{income_label}**.
        - The y-axis is **{growth_label}**.
        - The fitted curves come from the fitted IQR sheet of the selected workbook.
        - The latest country dots come from `Latest_country_data`.
        - The full country trajectories come from `Full_country_trajectory`.
        - The trajectory y-axis is the HP-smoothed growth rate produced by the R code.
        - Streamlit does not compute any smoothing.
        - Latest datapoints are labelled directly next to the selected dots.
        - Country trajectories are labelled at the final available datapoint.
        - To hide latest dots, clear **Countries for latest datapoints**.
        - To hide trajectories, clear **Countries for trajectories**.
        - The DCT position is computed by comparing each latest country point with the fitted 25th and 75th percentile curves.
        - For the PWT version, the x-axis is capped at **150,000** GDP per capita.
        - Average-income reference lines are shown in white when enabled.

        **Expected default files**

        - WB multi-measure version: `DCT_output_wb_multi.xlsx`
        - PWT multi-measure version: `DCT_output_pwt_multi.xlsx`

        **Expected WB measures**

        - `gni_atlas_current`
        - `gdp_current_usd`
        - `gdp_constant_usd`
        - `gdp_ppp_current_intl`
        - `gdp_ppp_constant_intl`

        **Expected PWT measures**

        - `rgdpna`
        - `rgdpe`
        - `rgdpo`
        - `cgdpe`
        - `cgdpo`
        """
    )