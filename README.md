# Finance KPI Dashboard (Python + Streamlit)

A recruiter-ready finance dashboard that loads Actuals and Budget data from Excel and provides executive-style financial analysis.

## Live Demo
https://finance-kpi-dashboard-qsthhxfxjezxo3te28cquy.streamlit.app/

## Features
- KPI cards with deltas vs Budget
- Executive-style Key Insights narrative
- Monthly trends (Revenue, Gross Margin %)
- Budget vs Actual comparison
- Month-by-month Net Income waterfall (Budget → Actual)
- Export filtered data to Excel
- Executive Summary PDF (if reportlab is available)

## Tech Stack
- Python
- Streamlit
- Pandas
- Plotly
- Excel I/O via openpyxl

## Project Structure
finance_kpi_dashboard/
├─ app.py
├─ requirements.txt
├─ README.md
├─ screenshot_kpis.png
├─ screenshot_trends.png
├─ screenshot_variance.png
├─ data/
│  └─ finance_kpi.xlsx
└─ .gitignore

## Run Locally
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py

## Data Format
Excel file: data/finance_kpi.xlsx

Sheets:
- Actuals
- Budget

Required columns:
- Date
- Department
- Account (Revenue, COGS, OpEx)
- Amount (Revenue positive; expenses negative)

## Screenshots

### KPI Overview & Insights
![KPIs](screenshot_kpis.png)

### Trends
![Trends](screenshot_trends.png)

### Variance Analysis
![Variance](screenshot_variance.png)
