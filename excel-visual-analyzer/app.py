import streamlit as st
import pandas as pd
from src.analyzer import (
    load_workbook_sheets,
    profile_dataframe,
    guess_datetime_columns,
    safe_to_datetime,
)
from src.viz import (
    fig_histogram,
    fig_box,
    fig_bar_topk,
    fig_corr_heatmap,
    fig_timeseries,
)

st.set_page_config(page_title="Excel Visual Analyzer", layout="wide")

st.title("📊 Excel Visual Analyzer")
st.caption("Upload an Excel workbook and get quick profiling + visual analysis.")

with st.sidebar:
    st.header("1) Upload")
    uploaded = st.file_uploader("Excel file", type=["xlsx", "xls"])

    st.header("2) Settings")
    max_rows = st.number_input("Max rows to load per sheet (0 = all)", min_value=0, value=200_000, step=10_000)
    top_k = st.slider("Top-K categories to display", 5, 50, 15)

if uploaded is None:
    st.info("Upload an Excel workbook to begin.")
    st.stop()

# Load workbook
try:
    sheets = load_workbook_sheets(uploaded, max_rows=max_rows if max_rows != 0 else None)
except Exception as e:
    st.error(f"Failed to read workbook: {e}")
    st.stop()

sheet_names = list(sheets.keys())
if not sheet_names:
    st.warning("No sheets found.")
    st.stop()

col1, col2 = st.columns([2, 1])
with col1:
    sheet_name = st.selectbox("Select sheet", sheet_names, index=0)
with col2:
    st.download_button(
        "Download selected sheet as CSV",
        data=sheets[sheet_name].to_csv(index=False).encode("utf-8"),
        file_name=f"{sheet_name}.csv",
        mime="text/csv",
    )

df = sheets[sheet_name]

tab_overview, tab_columns, tab_corr, tab_timeseries, tab_all_sheets = st.tabs(
    ["Overview", "Column Explorer", "Correlation", "Time Series", "All Sheets Summary"]
)

with tab_overview:
    st.subheader(f"Sheet: {sheet_name}")
    st.write("Preview")
    st.dataframe(df.head(200), use_container_width=True)

    profile = profile_dataframe(df)

    a, b, c, d = st.columns(4)
    a.metric("Rows", f"{profile['shape'][0]:,}")
    b.metric("Columns", f"{profile['shape'][1]:,}")
    c.metric("Missing cells", f"{profile['missing_cells']:,}")
    d.metric("Duplicate rows", f"{profile['duplicate_rows']:,}")

    st.markdown("### Data types")
    st.dataframe(pd.DataFrame(profile["dtypes"].items(), columns=["column", "dtype"]), use_container_width=True)

    st.markdown("### Missing values by column")
    miss_df = pd.DataFrame(profile["missing_by_col"].items(), columns=["column", "missing"])
    miss_df["missing_pct"] = (miss_df["missing"] / max(profile["shape"][0], 1) * 100).round(2)
    st.dataframe(miss_df.sort_values("missing", ascending=False), use_container_width=True)

    st.markdown("### Numeric summary")
    if profile["numeric_summary_rows"]:
        st.dataframe(pd.DataFrame(profile["numeric_summary_rows"]), use_container_width=True)
    else:
        st.info("No numeric columns found.")

    st.markdown("### Categorical summary (top values)")
    if profile["categorical_summary_rows"]:
        st.dataframe(pd.DataFrame(profile["categorical_summary_rows"]), use_container_width=True)
    else:
        st.info("No categorical columns found.")

with tab_columns:
    st.subheader("Explore columns")

    numeric_cols = df.select_dtypes(include="number").columns.tolist()

    left, right = st.columns([1, 2], gap="large")

    with left:
        st.markdown("#### Choose a column")
        col = st.selectbox("Column", df.columns.tolist())

        st.markdown("#### Basic stats")
        series = df[col]
        st.write({
            "dtype": str(series.dtype),
            "non_null": int(series.notna().sum()),
            "null": int(series.isna().sum()),
            "unique": int(series.nunique(dropna=True)),
        })

        if col in numeric_cols:
            st.markdown("#### Plot controls")
            bins = st.slider("Histogram bins", 5, 200, 30)
            show_box = st.checkbox("Show box plot", value=True)
        else:
            st.markdown("#### Plot controls")
            include_na = st.checkbox("Include NaN as category", value=False)

    with right:
        if col in numeric_cols:
            st.plotly_chart(fig_histogram(df, col, bins=bins), use_container_width=True)
            if show_box:
                st.plotly_chart(fig_box(df, col), use_container_width=True)

            st.markdown("#### Outlier hint (IQR rule)")
            s = df[col].dropna()
            if len(s) >= 10:
                q1 = s.quantile(0.25)
                q3 = s.quantile(0.75)
                iqr = q3 - q1
                lo = q1 - 1.5 * iqr
                hi = q3 + 1.5 * iqr
                outliers = ((s < lo) | (s > hi)).sum()
                st.write({"q1": float(q1), "q3": float(q3), "iqr": float(iqr), "lower": float(lo), "upper": float(hi), "outlier_count": int(outliers)})
            else:
                st.info("Not enough numeric samples to estimate outliers.")
        else:
            st.plotly_chart(fig_bar_topk(df, col, top_k=top_k, include_na=include_na), use_container_width=True)

with tab_corr:
    st.subheader("Correlation heatmap (numeric columns)")
    if len(df.select_dtypes(include="number").columns) < 2:
        st.info("Need at least 2 numeric columns to compute correlations.")
    else:
        method = st.selectbox("Correlation method", ["pearson", "spearman", "kendall"], index=0)
        st.plotly_chart(fig_corr_heatmap(df, method=method), use_container_width=True)

with tab_timeseries:
    st.subheader("Time series")
    dt_cols = guess_datetime_columns(df)

    if not dt_cols:
        st.info("No obvious datetime columns detected. You can still pick a column and try parsing it.")
        dt_cols = df.columns.tolist()

    dt_col = st.selectbox("Datetime column", dt_cols, index=0 if dt_cols else 0)
    y_candidates = df.select_dtypes(include="number").columns.tolist()
    if not y_candidates:
        st.info("No numeric columns available for Y-axis.")
    else:
        y_col = st.selectbox("Value (Y) column", y_candidates, index=0)
        freq = st.selectbox("Resample frequency", ["None", "D", "W", "M", "Q", "Y"], index=0)
        agg = st.selectbox("Aggregation", ["sum", "mean", "median", "min", "max", "count"], index=1)

        parsed = safe_to_datetime(df, dt_col)
        if parsed is None:
            st.error("Could not parse the selected datetime column.")
        else:
            st.plotly_chart(fig_timeseries(df, parsed, y_col=y_col, freq=None if freq == "None" else freq, agg=agg), use_container_width=True)

with tab_all_sheets:
    st.subheader("Workbook summary")
    rows = []
    for name, sdf in sheets.items():
        p = profile_dataframe(sdf)
        rows.append({
            "sheet": name,
            "rows": p["shape"][0],
            "cols": p["shape"][1],
            "missing_cells": p["missing_cells"],
            "duplicate_rows": p["duplicate_rows"],
            "numeric_cols": len(sdf.select_dtypes(include="number").columns),
            "categorical_cols": len([c for c in sdf.columns if c not in sdf.select_dtypes(include="number").columns]),
        })
    st.dataframe(pd.DataFrame(rows).sort_values(["rows", "cols"], ascending=False), use_container_width=True)

st.caption("Tip: If your workbook is huge, reduce the max rows to load in the sidebar for faster profiling.")
