from __future__ import annotations

import io
from typing import Dict, Optional, List, Any

import pandas as pd
import numpy as np


def load_workbook_sheets(uploaded_file, max_rows: Optional[int] = None) -> Dict[str, pd.DataFrame]:
    \"\"\"Load all sheets from an uploaded Excel file into DataFrames.\"\"\"
    data = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()
    bio = io.BytesIO(data)

    xls = pd.ExcelFile(bio)
    out: Dict[str, pd.DataFrame] = {}

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        if max_rows is not None and len(df) > max_rows:
            df = df.iloc[:max_rows].copy()
        out[sheet_name] = df

    return out


def profile_dataframe(df: pd.DataFrame) -> Dict[str, Any]:
    shape = df.shape
    missing_by_col = df.isna().sum().to_dict()
    missing_cells = int(df.isna().sum().sum())
    duplicate_rows = int(df.duplicated().sum())
    dtypes = {c: str(t) for c, t in df.dtypes.items()}

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    categorical_cols = [c for c in df.columns if c not in numeric_cols]

    numeric_summary_rows = []
    for c in numeric_cols:
        s = df[c]
        numeric_summary_rows.append({
            "column": c,
            "count": int(s.count()),
            "mean": float(s.mean()) if s.count() else np.nan,
            "std": float(s.std()) if s.count() else np.nan,
            "min": float(s.min()) if s.count() else np.nan,
            "p25": float(s.quantile(0.25)) if s.count() else np.nan,
            "median": float(s.median()) if s.count() else np.nan,
            "p75": float(s.quantile(0.75)) if s.count() else np.nan,
            "max": float(s.max()) if s.count() else np.nan,
        })

    categorical_summary_rows = []
    for c in categorical_cols:
        s = df[c]
        nunique = int(s.nunique(dropna=True))
        top = s.value_counts(dropna=True).head(5)
        top_fmt = ", ".join([f"{idx} ({val})" for idx, val in top.items()])
        categorical_summary_rows.append({
            "column": c,
            "count": int(s.count()),
            "unique": nunique,
            "top_values": top_fmt
        })

    return {
        "shape": shape,
        "missing_by_col": missing_by_col,
        "missing_cells": missing_cells,
        "duplicate_rows": duplicate_rows,
        "dtypes": dtypes,
        "numeric_summary_rows": numeric_summary_rows,
        "categorical_summary_rows": categorical_summary_rows,
    }


def guess_datetime_columns(df: pd.DataFrame) -> List[str]:
    dt_cols = []
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            dt_cols.append(c)
            continue
        if pd.api.types.is_object_dtype(df[c]) or pd.api.types.is_string_dtype(df[c]):
            sample = df[c].dropna().astype(str).head(50)
            if sample.empty:
                continue
            parsed = pd.to_datetime(sample, errors="coerce", infer_datetime_format=True)
            success = parsed.notna().mean()
            if success >= 0.7:
                dt_cols.append(c)
    return dt_cols


def safe_to_datetime(df: pd.DataFrame, col: str):
    try:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            return df[col]
        parsed = pd.to_datetime(df[col], errors="coerce", infer_datetime_format=True)
        if parsed.notna().sum() == 0:
            return None
        return parsed
    except Exception:
        return None
