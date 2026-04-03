# MIS581 Capstone Project — U.S.-Canada Tariff Impact Analysis

**CSU Global | MIS581 Capstone**  
**Alex Vollen**

---

## What This Is

This project looks at how U.S. tariffs affected trade between the U.S. and Canada across three sectors: steel, aluminum, and automobiles. The analysis covers two tariff periods — the 2018–2019 Section 232 tariffs and the 2025 reimposition — using 15 HS6 commodity codes pulled from USITC and Statistics Canada data going back to 2012.

The three research questions driving the analysis:

1. Did Canada's imports from the U.S. drop during tariff periods?
2. Did Canada shift to non-U.S. suppliers as a result?
3. Did Canada's exports to the U.S. drop during tariff periods?

RQ1 and RQ3 rejected the null hypothesis. RQ2 did not.

---

## HS Codes

| Sector | Codes |
|--------|-------|
| Steel | 720825, 720851, 720916, 721070, 721310 |
| Aluminum | 760110, 760120, 760421, 760612, 760711 |
| Auto | 870323, 870324, 870421, 870840, 870899 |

---

## Folder Structure

```
Inputs/
│
├── Tariff Final Project All Code.py
│
├── US Exports To Canada Jan-2012 to Jan-2026 All Codes.xlsx
├── US Imports From Canada Jan-2012 to Jan-2026 All Codes.xlsx
│
├── Steel YoY 2012-2024 (5 Codes).xlsx
├── Steel Last 24 Months (5 Codes).xlsx
├── Aluminum YoY 2012-2024 (5 Codes).xlsx
├── Aluminum Last 24 Months (5 Codes).xlsx
├── Auto YoY 2012-2024 (5 Codes).xlsx
├── Auto Last 24 Months (5 Codes).xlsx
│
└── Outputs/
    ├── Histogram_Source_Data.xlsx
    ├── Various chart PNGs (histograms, trend lines, forecasts)
```

---

## How to Run It

Make sure all the input files are in the same folder as the script, then:

```bash
pip install pandas numpy matplotlib scipy openpyxl
python "Tariff Final Project All Code.py"
```

The script creates an `Outputs/` folder automatically and saves everything there.

---

## What the Script Does

The Canadian Statistics Canada files came in a non-standard format where year headers appear before the HS code labels, so a custom row-by-row parser handles those. The U.S. USITC files are cleaner and load more straightforwardly.

Once loaded, the script runs:

- **Descriptive statistics** by sector for both monthly U.S. and annual Canadian series
- **Mann-Whitney U tests** comparing tariff-on vs. tariff-off periods (non-parametric since trade distributions are not assumed normal)
- **OLS regression** with a binary tariff dummy to estimate the dollar-value association
- **Two-scenario forecast** projecting 24 months out (2026-2027) — one with tariffs remaining, one without

---

## Outputs

| File | What it is |
|------|------------|
| `Histogram_Source_Data.xlsx` | All aggregated monthly and annual trade data, one tab per sector/series |
| `US_Exports_*_Monthly_Histogram.png` | Distribution of monthly U.S. export values |
| `US_Imports_*_Monthly_Histogram.png` | Distribution of monthly U.S. import values |
| `Canadian_Imports_*_Annual_Histogram.png` | Distribution of Canadian annual import totals |
| `Canadian_Imports_*_US_vs_NonUS_Trend.png` | U.S. vs. non-U.S. sourced imports over time |
| `US_Exports_*_Monthly_Trend.png` | Monthly trend with 12-month rolling average |
| `US_Imports_*_Monthly_Trend.png` | Monthly trend with 12-month rolling average |
| `Forecast_Canada_Exports_*.png` | Two-scenario forecast charts by sector |

---

## Data Sources

- **USITC DataWeb** — Monthly U.S. export (FAS value) and import (general customs value), January 2012 – January 2026
- **Statistics Canada** — Annual import data by country of origin (2012-2024) plus last-24-months monthly files

---

## Tariff Periods

| Period | Dates | Context |
|--------|-------|---------|
| Episode 1 | March 2018 – April 2019 | Section 232 steel and aluminum tariffs |
| Episode 2 | March 2025 – present | Tariff reimposition |
