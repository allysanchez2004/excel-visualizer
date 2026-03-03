# Excel Visual Analyzer (Streamlit)

Upload an Excel workbook (`.xlsx`, `.xls`) and get:
- Sheet-by-sheet profiling (rows/cols, dtypes, missing values, uniques)
- Numeric summaries + outlier hints
- Visual analysis:
  - Histograms / box plots for numeric columns
  - Bar charts for categorical columns (top values)
  - Correlation heatmap for numeric columns
  - Date/time trend plot (if a datetime column is detected/selected)

## Quick start

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
# source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

## Notes
- This is a local analytics tool. No data is uploaded anywhere unless you deploy it yourself.
- Large files: start by selecting a sheet and limiting rows in the UI.

## Project structure
- `app.py` — Streamlit UI + analysis
- `src/analyzer.py` — workbook parsing + profiling helpers
- `src/viz.py` — plotting helpers
