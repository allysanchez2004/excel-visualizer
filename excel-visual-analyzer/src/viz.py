from __future__ import annotations

import pandas as pd
import plotly.express as px


def fig_histogram(df: pd.DataFrame, col: str, bins: int = 30):
    fig = px.histogram(df, x=col, nbins=bins, title=f"Histogram — {col}")
    fig.update_layout(bargap=0.05)
    return fig


def fig_box(df: pd.DataFrame, col: str):
    fig = px.box(df, y=col, points="outliers", title=f"Box plot — {col}")
    return fig


def fig_bar_topk(df: pd.DataFrame, col: str, top_k: int = 15, include_na: bool = False):
    s = df[col]
    counts = (s.fillna("NaN") if include_na else s.dropna()).value_counts().head(top_k)
    plot_df = counts.reset_index()
    plot_df.columns = [col, "count"]
    fig = px.bar(plot_df, x=col, y="count", title=f"Top {top_k} values — {col}")
    fig.update_layout(xaxis_tickangle=-30)
    return fig


def fig_corr_heatmap(df: pd.DataFrame, method: str = "pearson"):
    num = df.select_dtypes(include="number")
    corr = num.corr(method=method)
    fig = px.imshow(corr, text_auto=True, title=f"Correlation heatmap ({method})")
    fig.update_layout(height=700)
    return fig


def fig_timeseries(df: pd.DataFrame, dt_series: pd.Series, y_col: str, freq: str | None, agg: str = "mean"):
    tmp = df.copy()
    tmp["_dt_"] = dt_series
    tmp = tmp.dropna(subset=["_dt_"]).sort_values("_dt_")

    if freq is None:
        plot_df = tmp[["_dt_", y_col]].dropna()
        fig = px.line(plot_df, x="_dt_", y=y_col, title=f"Time series — {y_col}")
        fig.update_layout(xaxis_title="date/time", yaxis_title=y_col)
        return fig

    tmp = tmp.set_index("_dt_")
    rs = getattr(tmp[y_col].resample(freq), agg)()
    plot_df = rs.reset_index()
    plot_df.columns = ["date", y_col]
    fig = px.line(plot_df, x="date", y=y_col, title=f"Time series ({freq}, {agg}) — {y_col}")
    fig.update_layout(xaxis_title="date/time", yaxis_title=y_col)
    return fig
