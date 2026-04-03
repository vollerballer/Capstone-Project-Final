"""
MIS581 Capstone Project — Tariff Impact Analysis
U.S.-Canada Trade: Steel, Aluminum, and Automobile Sectors

Analyzes U.S. and Canadian trade data across three tariff-exposed
sectors using descriptive statistics, Mann-Whitney U tests, OLS
regression, and a two-scenario predictive forecast.

Place this file in the Inputs folder alongside all data files.
Run from terminal: python "Tariff Final Project.py"
"""

import os
import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from scipy import stats

warnings.filterwarnings("ignore")

# PATHS AND OUTPUT FOLDER

BASE_PATH     = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_PATH, "Outputs")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# INPUT FILE LOCATIONS

FILES = {
    "us_exports": os.path.join(BASE_PATH, "US Exports To Canada Jan-2012 to Jan-2026 All Codes.xlsx"),
    "us_imports": os.path.join(BASE_PATH, "US Imports From Canada Jan-2012 to Jan-2026 All Codes.xlsx"),
    "steel_yoy":  os.path.join(BASE_PATH, "Steel YoY 2012-2024 (5 Codes).xlsx"),
    "steel_l24":  os.path.join(BASE_PATH, "Steel Last 24 Months (5 Codes).xlsx"),
    "alum_yoy":   os.path.join(BASE_PATH, "Aluminum YoY 2012-2024 (5 Codes).xlsx"),
    "alum_l24":   os.path.join(BASE_PATH, "Aluminum Last 24 Months (5 Codes).xlsx"),
    "auto_yoy":   os.path.join(BASE_PATH, "Auto YoY 2012-2024 (5 Codes).xlsx"),
    "auto_l24":   os.path.join(BASE_PATH, "Auto Last 24 Months (5 Codes).xlsx"),
}



# SECTOR HS CODE DEFINITIONS
# Each sector maps to five 6-digit HS codes used throughout
# the analysis for filtering and grouping.

SECTORS = {
    "Steel":    [720825, 720851, 720916, 721070, 721310],
    "Aluminum": [760110, 760120, 760421, 760612, 760711],
    "Auto":     [870323, 870324, 870421, 870840, 870899],
}

# Flat lookup: HS code -> sector name, built once and reused throughout
CODE_TO_SECTOR = {code: sector for sector, codes in SECTORS.items() for code in codes}

# REFERENCE LOOKUPS

# Abbreviated month names to zero-padded numbers.
MONTH_MAP = {
    "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
    "May": "05", "June": "06", "Jun": "06", "Jul": "07",
    "Aug": "08", "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12",
}

# All name variants used for the United States in Canadian source files
US_COUNTRY_NAMES = {"United States", "United States of America", "U.S.", "USA", "US"}

# Tariff episode date ranges used across all three RQs.
# Episode 1: (Mar 2018 – Apr 2019)
# Episode 2: 2025 tariff reimposition (Mar 2025 onward)
TARIFF_ON_PERIODS = [("2018-03", "2019-04"), ("2025-03", "2026-01")]
TARIFF_ON_YEARS   = {2018, 2019, 2025}

# Forecast configuration for the RQ3 predictive model
FORECAST_MONTHS = 24
BASELINE_START  = "2012-01"
BASELINE_END    = "2018-02"
LAST_ACTUAL     = "2026-01"



# UTILITY FUNCTIONS THESE FUNCTIONS ARE USED FOR DATA CLEANING AND ENSURING THE STUDY IS REPEATABLE WITH NEW DATA

def money_fmt(x):
    """Format a number as whole-dollar currency string."""
    return f"${x:,.0f}"


def safe_sheet_name(name, max_len=31):
    """Strip characters Excel disallows in sheet names."""
    for ch in ["\\", "/", "*", "[", "]", ":", "?"]:
        name = name.replace(ch, "_")
    return name[:max_len]


def extract_hs_code(s):
    """
    Pull a 6-digit HS code out of a mixed text string such as
    'HS 720825 - Hot-rolled steel'.  Returns None if not found.
    """
    for part in str(s).replace("-", " ").split():
        try:
            code = int(part)
            if 100000 <= code <= 999999:
                return code
        except ValueError:
            pass
    return None


def find_years_in_row(vals):
    """
    Check whether a raw Excel row contains a run of year values
    (2010–2030).  Returns (start_col_index, [years]) or None.
    Requires at least three yearS hits to avoid false positives.
    """
    hits = [
        (ci, int(v))
        for ci, v in enumerate(vals)
        if v is not None
        and not (isinstance(v, float) and np.isnan(v))
        and _is_year(v)
    ]
    return (hits[0][0], [y for _, y in hits]) if len(hits) >= 3 else None


def _is_year(v):
    """Return True if v can be cast to an integer in [2010, 2030]."""
    try:
        return 2010 <= int(v) <= 2030
    except (ValueError, TypeError, OverflowError):
        return False


def parse_period(s):
    """
    Convert a header like '2025 Jan' into the ISO-style string
    '2025-01'.  Returns None if the format is not recognized.
    """
    parts = str(s).strip().split()
    if len(parts) == 2:
        month = MONTH_MAP.get(parts[1], "00")
        return f"{parts[0]}-{month}" if month != "00" else None
    return None


def classify_country_group(country_name):
    """Return 'US' or 'Non-US' for a given country string."""
    return "US" if str(country_name).strip() in US_COUNTRY_NAMES else "Non-US"


def flag_tariff_period(period_str):
    """Return 'Tariff ON' if the YYYY-MM period falls in a tariff episode."""
    for start, end in TARIFF_ON_PERIODS:
        if start <= period_str <= end:
            return "Tariff ON"
    return "Tariff OFF"


def flag_tariff_year(year):
    """Return 'Tariff ON' if the calendar year is a tariff episode year."""
    return "Tariff ON" if int(year) in TARIFF_ON_YEARS else "Tariff OFF"



# AXIS / CHART FORMATTING HELPERS MAKES CHARTS ADJUSTABLE WITHOUT HAVING TO MANUAL KEY IN EACH TIME

def get_compact_scale(values):
    """
    Inspect an array of values and choose the most readable scale
    (billions, millions, thousands, or raw dollars).
    Returns (scale_divisor, suffix_string, axis_label).
    """
    arr = np.asarray(values, dtype=float).ravel()
    arr = arr[~np.isnan(arr)]
    if arr.size == 0:
        return 1, "", "US Dollars"

    max_abs = np.max(np.abs(arr))
    if max_abs >= 1_000_000_000:
        return 1_000_000_000, "B", "Billions of US Dollars"
    elif max_abs >= 1_000_000:
        return 1_000_000, "M", "Millions of US Dollars"
    elif max_abs >= 1_000:
        return 1_000, "K", "Thousands of US Dollars"
    return 1, "", "US Dollars"


def make_currency_formatter(scale, suffix):
    """Return a matplotlib FuncFormatter for compact currency labels."""
    def _fmt(x, pos):
        if scale == 1:
            return f"${x:,.0f}"
        scaled = x / scale
        if abs(scaled) >= 100:
            return f"${scaled:,.0f}{suffix}"
        elif abs(scaled) >= 10:
            return f"${scaled:,.1f}".rstrip("0").rstrip(".") + suffix
        return f"${scaled:,.2f}".rstrip("0").rstrip(".") + suffix
    return _fmt


def apply_compact_currency_axis(ax, values, axis="y"):
    """
    Apply compact currency formatting to a chart axis and suppress
    matplotlib's automatic scientific-notation offset text.
    Returns the unit label string (e.g. 'Millions of US Dollars').
    """
    scale, suffix, unit_label = get_compact_scale(values)
    formatter = FuncFormatter(make_currency_formatter(scale, suffix))
    target = ax.yaxis if axis == "y" else ax.xaxis
    target.set_major_formatter(formatter)
    target.offsetText.set_visible(False)
    return unit_label


def compact_currency(x, pos):
    """Standalone compact currency formatter for forecast charts."""
    if abs(x) >= 1_000_000_000:
        return f"${x / 1_000_000_000:.1f}B"
    elif abs(x) >= 1_000_000:
        return f"${x / 1_000_000:.1f}M"
    elif abs(x) >= 1_000:
        return f"${x / 1_000:.0f}K"
    return f"${x:,.0f}"


# =============================================================
# DATA LOADING FUNCTIONS
# =============================================================

def load_us_monthly(filepath, value_col):
    """
    Load a U.S. ITC monthly trade file (exports or imports).
    Assigns each row to Steel, Aluminum, or Auto based on the
    HTS Number, then drops rows that don't match any sector.
    """
    df = pd.read_excel(filepath, sheet_name=1)
    df = df.dropna(subset=["HTS Number", "Year", "Month"])

    df["HTS Number"] = df["HTS Number"].astype(int)
    df["Year"]       = df["Year"].astype(int)
    df["Month"]      = df["Month"].astype(int)
    df["Period"]     = df["Year"].astype(str) + "-" + df["Month"].astype(str).str.zfill(2)
    df[value_col]    = pd.to_numeric(df[value_col], errors="coerce").fillna(0)
    df["Sector"]     = df["HTS Number"].map(CODE_TO_SECTOR).fillna("Other")

    return df[df["Sector"] != "Other"].copy()


# Row labels to skip when parsing Canadian non-tabular Excel files. The canada data came back not nearly as clean as the US data website
# These appear as metadata or section headers rather than data rows.
_SKIP_LABELS = {"Title", "Products", "Origin", "Destination", "Period", "Units", "Source", "Note", ""}


def _is_skip_row(col_a):
    """Return True if a row label should be ignored during Canadian file parsing."""
    return col_a in _SKIP_LABELS or col_a.startswith(("Source", "Note"))


def _is_subtotal_row(col_a):
    """Return True if a row is an aggregated subtotal that should be excluded."""
    return "Sub-Total" in col_a or "Total All Countries" in col_a


def load_canadian_yoy(filepath, sector_name):
    """
    Parse a Statistics Canada YoY annual Excel file.

    These files use a non-standard layout: year headers appear
    above the data rows and HS code labels appear in column A
    before each block of country rows.  The parser walks row by
    row, updating state variables as it encounters year headers
    and HS code markers.
    """
    df_raw = pd.read_excel(filepath, sheet_name=0, header=None)
    records, year_cols, year_start, in_data, current_hs = [], None, None, False, None

    for _, row in df_raw.iterrows():
        vals  = row.tolist()
        col_a = str(vals[0]).strip() if vals[0] is not None and str(vals[0]) != "nan" else ""

        year_result = find_years_in_row(vals)
        if year_result is not None:
            year_start, year_cols = year_result
            in_data    = True
            current_hs = current_hs or "All"
            continue

        if col_a.upper().startswith("HS "):
            code = extract_hs_code(col_a)
            if code:
                current_hs = code
            continue

        if _is_skip_row(col_a) or _is_subtotal_row(col_a):
            continue

        if in_data and year_cols and col_a and current_hs is not None:
            for j, year in enumerate(year_cols):
                col_idx = year_start + j
                if col_idx < len(vals):
                    val = vals[col_idx]
                    if val is not None and not (isinstance(val, float) and np.isnan(val)):
                        try:
                            records.append({
                                "Sector":   sector_name,
                                "HS Code":  current_hs,
                                "Country":  col_a,
                                "Year":     int(year),
                                "Value":    float(val),
                            })
                        except (ValueError, TypeError):
                            pass

    return pd.DataFrame(records)


def load_canadian_l24(filepath, sector_name):
    """
    Parse a Statistics Canada Last 24 Months monthly Excel file.

    Similar non-standard layout to the YoY files, but period
    headers use 'YYYY Mon' strings instead of bare year integers.
    The parser identifies the period header row by checking whether
    column B contains a recognizable month abbreviation.
    """
    df_raw = pd.read_excel(filepath, sheet_name=0, header=None)
    records, period_cols, in_data, current_hs = [], None, False, None

    for _, row in df_raw.iterrows():
        vals  = row.tolist()
        col_a = str(vals[0]).strip() if vals[0] is not None and str(vals[0]) != "nan" else ""
        col_b = str(vals[1]).strip() if len(vals) > 1 and vals[1] is not None and str(vals[1]) != "nan" else ""

        # Detect the period header row (column B contains a month name)
        if any(m in col_b for m in MONTH_MAP):
            period_cols = [
                parse_period(str(v))
                for v in vals[1:]
                if v is not None and str(v) != "nan" and parse_period(str(v))
            ]
            in_data    = True
            current_hs = current_hs or "All"
            continue

        if col_a.upper().startswith("HS "):
            code = extract_hs_code(col_a)
            if code:
                current_hs = code
            continue

        if _is_skip_row(col_a) or _is_subtotal_row(col_a):
            continue

        if in_data and period_cols and col_a and current_hs is not None:
            for j, period in enumerate(period_cols):
                col_idx = j + 1
                if col_idx < len(vals):
                    val = vals[col_idx]
                    if val is not None and not (isinstance(val, float) and np.isnan(val)):
                        try:
                            records.append({
                                "Sector":   sector_name,
                                "HS Code":  current_hs,
                                "Country":  col_a,
                                "Period":   period,
                                "Value":    float(val),
                            })
                        except (ValueError, TypeError):
                            pass

    return pd.DataFrame(records)

# LOAD ALL DATA I DID OUTPUT THE DATA IN DATAFRAMES TO VALIDATE THIS OPERATED AS INTENDED

print("=" * 55)
print(f"  Loading data from: {BASE_PATH}")
print("=" * 55)

print("\n--- U.S. Trade Files (USITC) ---")
df_exp = load_us_monthly(FILES["us_exports"], "FAS Value")
print(f"  Exports:   {df_exp.shape[0]:,} rows | {df_exp.Period.min()} to {df_exp.Period.max()}")

df_imp = load_us_monthly(FILES["us_imports"], "General Customs Value")
print(f"  Imports:   {df_imp.shape[0]:,} rows | {df_imp.Period.min()} to {df_imp.Period.max()}")

print("\n--- Canadian Annual Files (Statistics Canada YoY) ---")
steel_yoy = load_canadian_yoy(FILES["steel_yoy"], "Steel")
alum_yoy  = load_canadian_yoy(FILES["alum_yoy"],  "Aluminum")
auto_yoy  = load_canadian_yoy(FILES["auto_yoy"],  "Auto")
for name, df in [("Steel", steel_yoy), ("Aluminum", alum_yoy), ("Auto", auto_yoy)]:
    years = sorted(df["Year"].unique())
    print(f"  {name:<10} {df.shape[0]:,} rows | {years[0]} to {years[-1]}")

print("\n--- Canadian Monthly Files (Statistics Canada L24) ---")
steel_l24 = load_canadian_l24(FILES["steel_l24"], "Steel")
alum_l24  = load_canadian_l24(FILES["alum_l24"],  "Aluminum")
auto_l24  = load_canadian_l24(FILES["auto_l24"],  "Auto")
for name, df in [("Steel", steel_l24), ("Aluminum", alum_l24), ("Auto", auto_l24)]:
    periods = sorted(df["Period"].unique())
    print(f"  {name:<10} {df.shape[0]:,} rows | {periods[0]} to {periods[-1]}")


# BUILD CORE AGGREGATIONS

# Monthly totals for U.S. export and import series
us_exp_monthly = (
    df_exp.groupby(["Sector", "Period"], as_index=False)["FAS Value"]
    .sum()
    .sort_values(["Sector", "Period"])
)

us_imp_monthly = (
    df_imp.groupby(["Sector", "Period"], as_index=False)["General Customs Value"]
    .sum()
    .sort_values(["Sector", "Period"])
)

# Combine Canadian YoY files and assign US / Non-US country groupings
can_yoy_all = pd.concat([steel_yoy, alum_yoy, auto_yoy], ignore_index=True)
can_l24_all = pd.concat([steel_l24, alum_l24, auto_l24], ignore_index=True)

can_yoy_all["Country Group"] = can_yoy_all["Country"].apply(classify_country_group)
can_l24_all["Country Group"] = can_l24_all["Country"].apply(classify_country_group)

# Annual totals from YoY files (2012–2024)
can_yoy_totals = (
    can_yoy_all.groupby(["Sector", "Year"], as_index=False)["Value"]
    .sum()
    .sort_values(["Sector", "Year"])
)

# 2025 annual totals aggregated from the L24 monthly rows had to add this because for some reason you can't get yearly for 2025 unless you pull the monthly data. 
# Originally was going to use the monthly data more, but it didn't work easily to use data that transisitioned from yearly to monthly
can_2025_totals = (
    can_l24_all[can_l24_all["Period"].str.startswith("2025-")]
    .groupby("Sector", as_index=False)["Value"]
    .sum()
    .assign(Year=2025)
)[["Sector", "Year", "Value"]]

# Final Canadian annual totals 2012–2025
canadian_annual_totals = (
    pd.concat([can_yoy_totals, can_2025_totals], ignore_index=True)
    .sort_values(["Sector", "Year"])
    .reset_index(drop=True)
)

# Annual US vs Non-US breakdown from YoY files
can_us_nonus_yoy = (
    can_yoy_all.groupby(["Sector", "Year", "Country Group"], as_index=False)["Value"]
    .sum()
    .sort_values(["Sector", "Year", "Country Group"])
)

# 2025 US vs Non-US totals from L24 monthly rows
can_us_nonus_2025 = (
    can_l24_all[can_l24_all["Period"].str.startswith("2025-")]
    .groupby(["Sector", "Country Group"], as_index=False)["Value"]
    .sum()
    .assign(Year=2025)
)[["Sector", "Year", "Country Group", "Value"]]

# Final Canadian annual US vs Non-US breakdown 2012–2025
canadian_us_nonus_annual = (
    pd.concat([can_us_nonus_yoy, can_us_nonus_2025], ignore_index=True)
    .sort_values(["Sector", "Year", "Country Group"])
    .reset_index(drop=True)
)

# Non-US rows only — used directly in RQ2 hypothesis testing
can_nonus_annual = canadian_us_nonus_annual[
    canadian_us_nonus_annual["Country Group"] == "Non-US"
].copy()

# Attach tariff flags to the monthly U.S. trade series
us_exp_monthly["Tariff"] = us_exp_monthly["Period"].apply(flag_tariff_period)
us_imp_monthly["Tariff"] = us_imp_monthly["Period"].apply(flag_tariff_period)

# Attach tariff flags to the Canadian annual series
canadian_us_nonus_annual["Tariff"] = canadian_us_nonus_annual["Year"].apply(flag_tariff_year)
can_nonus_annual["Tariff"]         = can_nonus_annual["Year"].apply(flag_tariff_year)

print("\nCore aggregations complete.")
print(f"  Canadian annual totals: {canadian_annual_totals.shape[0]} rows | "
      f"{canadian_annual_totals.Year.min()}–{canadian_annual_totals.Year.max()}")
print(f"  US vs Non-US annual:    {canadian_us_nonus_annual.shape[0]} rows")



# DESCRIPTIVE STATISTICS

def describe_monthly(df, value_col, hs_col):
    """
    Compute descriptive statistics for a monthly trade dataset,
    grouped by sector.  Returns a dict of {sector: DataFrame}.
    """
    out = {}
    for sector in sorted(df["Sector"].dropna().unique()):
        sub     = df[df["Sector"] == sector]
        monthly = sub.groupby("Period")[value_col].sum().sort_index()
        out[sector] = pd.DataFrame({
            "Metric": [
                "Records (row count)", "Unique HS codes", "Unique periods",
                "Start month", "End month",
                "Total (all records)", "Mean monthly", "Median monthly",
                "Min monthly", "Max monthly", "Std dev monthly",
            ],
            "Value": [
                f"{len(sub):,}", f"{sub[hs_col].nunique():,}", f"{monthly.size:,}",
                str(monthly.index.min()), str(monthly.index.max()),
                money_fmt(sub[value_col].sum()),
                money_fmt(monthly.mean()), money_fmt(monthly.median()),
                money_fmt(monthly.min()),  money_fmt(monthly.max()),
                money_fmt(monthly.std(ddof=1) if len(monthly) > 1 else 0),
            ],
        })
    return out


def describe_annual(df, value_col="Value"):
    """
    Compute descriptive statistics for an annual trade dataset,
    grouped by sector.  Returns a dict of {sector: DataFrame}.
    """
    out = {}
    for sector in sorted(df["Sector"].dropna().unique()):
        sub    = df[df["Sector"] == sector]
        annual = sub.groupby("Year")[value_col].sum().sort_index()
        out[sector] = pd.DataFrame({
            "Metric": [
                "Records (row count)", "Unique years", "Start year", "End year",
                "Total (all records)", "Mean annual", "Median annual",
                "Min annual", "Max annual", "Std dev annual",
            ],
            "Value": [
                f"{len(sub):,}", f"{annual.size:,}",
                str(annual.index.min()), str(annual.index.max()),
                money_fmt(sub[value_col].sum()),
                money_fmt(annual.mean()), money_fmt(annual.median()),
                money_fmt(annual.min()),  money_fmt(annual.max()),
                money_fmt(annual.std(ddof=1) if len(annual) > 1 else 0),
            ],
        })
    return out


def print_stats(results, title):
    """Print sector-level descriptive statistics to the console."""
    print(f"\n{'=' * 65}\n  {title}\n{'=' * 65}")
    for sector, df in results.items():
        print(f"\n--- {sector} ---")
        print(df.to_string(index=False))


us_exp_stats      = describe_monthly(df_exp, "FAS Value", "HTS Number")
us_imp_stats      = describe_monthly(df_imp, "General Customs Value", "HTS Number")
canadian_ann_stats = describe_annual(canadian_annual_totals)

print_stats(us_exp_stats,       "U.S. EXPORTS TO CANADA — DESCRIPTIVE STATISTICS")
print_stats(us_imp_stats,       "U.S. IMPORTS FROM CANADA — DESCRIPTIVE STATISTICS")
print_stats(canadian_ann_stats, "CANADIAN ANNUAL IMPORTS — DESCRIPTIVE STATISTICS (2012–2025)")



# EXPORT SOURCE DATA TO EXCEL

export_path = os.path.join(OUTPUT_FOLDER, "Histogram_Source_Data.xlsx")

with pd.ExcelWriter(export_path, engine="openpyxl") as writer:
    us_exp_monthly.to_excel(writer, sheet_name="US_Exports_Monthly",   index=False)
    us_imp_monthly.to_excel(writer, sheet_name="US_Imports_Monthly",   index=False)
    canadian_annual_totals.to_excel(writer, sheet_name="Canada_Annual_2012_2025", index=False)
    canadian_us_nonus_annual.to_excel(writer, sheet_name="Canada_US_vs_NonUS",    index=False)

    for sector in sorted(us_exp_monthly["Sector"].unique()):
        us_exp_monthly[us_exp_monthly["Sector"] == sector].to_excel(
            writer, sheet_name=safe_sheet_name(f"US_EXP_{sector}"), index=False
        )
    for sector in sorted(us_imp_monthly["Sector"].unique()):
        us_imp_monthly[us_imp_monthly["Sector"] == sector].to_excel(
            writer, sheet_name=safe_sheet_name(f"US_IMP_{sector}"), index=False
        )
    # CAN_ tabs include both US and Non-US rows so each year has
    # a row for each source group
    for sector in sorted(canadian_us_nonus_annual["Sector"].unique()):
        canadian_us_nonus_annual[canadian_us_nonus_annual["Sector"] == sector].to_excel(
            writer, sheet_name=safe_sheet_name(f"CAN_{sector}"), index=False
        )
        canadian_us_nonus_annual[canadian_us_nonus_annual["Sector"] == sector].to_excel(
            writer, sheet_name=safe_sheet_name(f"Trend_{sector}"), index=False
        )

print(f"\nSource data exported: {export_path}")

# HISTOGRAM CHARTS

def save_histogram(series, title, output_path, bins=15):
    """
    Save a frequency histogram with compact currency formatting
    on the x-axis.  Skips silently if the series is empty.
    """
    vals = pd.to_numeric(series, errors="coerce").dropna()
    if vals.empty:
        print(f"  Skipped (no data): {title}")
        return

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.hist(vals, bins=bins, edgecolor="black")
    ax.set_title(title)
    ax.set_xlabel("US Dollars")
    ax.set_ylabel("Frequency")
    apply_compact_currency_axis(ax, vals.values, axis="x")
    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {os.path.basename(output_path)}")


for sector in sorted(us_exp_monthly["Sector"].unique()):
    save_histogram(
        series=us_exp_monthly[us_exp_monthly["Sector"] == sector]["FAS Value"],
        title=f"U.S. Exports to Canada — {sector} Monthly Totals",
        output_path=os.path.join(OUTPUT_FOLDER, f"US_Exports_{sector}_Monthly_Histogram.png"),
    )

for sector in sorted(us_imp_monthly["Sector"].unique()):
    save_histogram(
        series=us_imp_monthly[us_imp_monthly["Sector"] == sector]["General Customs Value"],
        title=f"U.S. Imports from Canada — {sector} Monthly Totals",
        output_path=os.path.join(OUTPUT_FOLDER, f"US_Imports_{sector}_Monthly_Histogram.png"),
    )

# Canadian histograms combine US and Non-US rows to maximize the
# number of data points available in each sector's distribution.
# Due to Aggregate data it isn't very many data points
for sector in sorted(canadian_us_nonus_annual["Sector"].unique()):
    save_histogram(
        series=canadian_us_nonus_annual[canadian_us_nonus_annual["Sector"] == sector]["Value"],
        title=f"Canadian Imports — {sector} Annual Totals (US + Non-US)",
        output_path=os.path.join(OUTPUT_FOLDER, f"Canadian_Imports_{sector}_Annual_Histogram.png"),
        bins=8,
    )


# CANADIAN TREND LINE CHARTS: US VS NON-US

def save_sector_trendline(df_sector, sector_name, output_folder):
    """
    Save an annual trend line chart for Canadian imports broken
    out by source group (US vs Non-US).
    """
    chart_df = (
        df_sector.pivot(index="Year", columns="Country Group", values="Value")
        .fillna(0)
        .sort_index()
    )
    years = chart_df.index.tolist()

    fig, ax = plt.subplots(figsize=(11, 6))
    for group in ["US", "Non-US"]:
        if group in chart_df.columns:
            ax.plot(years, chart_df[group], marker="o", linewidth=2, label=group)

    ax.set_title(f"Canadian Imports by Source — {sector_name} (Annual Totals)")
    ax.set_xlabel("Year")
    ax.set_ylabel("US Dollars")
    ax.set_xticks(years)
    ax.tick_params(axis="x", rotation=45)
    ax.legend()
    apply_compact_currency_axis(ax, chart_df.values, axis="y")

    plt.tight_layout()
    out = os.path.join(output_folder, f"Canadian_Imports_{sector_name}_US_vs_NonUS_Trend.png")
    plt.savefig(out, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {os.path.basename(out)}")


for sector in sorted(canadian_us_nonus_annual["Sector"].unique()):
    save_sector_trendline(
        canadian_us_nonus_annual[canadian_us_nonus_annual["Sector"] == sector],
        sector, OUTPUT_FOLDER
    )

# U.S. MONTHLY TREND LINE CHARTS WITH 12-MONTH ROLLING AVERAGE

def save_us_monthly_trendline(df_sector, value_col, chart_title, output_path):
    """
    Save a U.S. monthly trend line showing raw monthly values and
    a 12-month rolling average.  Tick marks are placed every 12
    months so the x-axis stays legible across the full date range.
    """
    chart_df = df_sector.sort_values("Period").reset_index(drop=True)
    chart_df["Rolling12"] = chart_df[value_col].rolling(window=12, min_periods=1).mean()

    fig, ax = plt.subplots(figsize=(11, 6))
    ax.plot(chart_df["Period"], chart_df[value_col],
            marker="o", linewidth=1.5, alpha=0.6, label="Monthly Total")
    ax.plot(chart_df["Period"], chart_df["Rolling12"],
            linewidth=3, label="12-Month Rolling Average")

    ax.set_title(chart_title)
    ax.set_xlabel("Period")
    ax.set_ylabel("US Dollars")

    tick_pos = list(range(0, len(chart_df), 12))
    if (len(chart_df) - 1) not in tick_pos:
        tick_pos.append(len(chart_df) - 1)

    ax.set_xticks(tick_pos)
    ax.set_xticklabels([chart_df["Period"].iloc[i] for i in tick_pos], rotation=45)
    ax.legend()

    all_vals = np.concatenate([chart_df[value_col].values, chart_df["Rolling12"].values])
    apply_compact_currency_axis(ax, all_vals, axis="y")

    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {os.path.basename(output_path)}")


for sector in sorted(us_exp_monthly["Sector"].unique()):
    save_us_monthly_trendline(
        df_sector=us_exp_monthly[us_exp_monthly["Sector"] == sector],
        value_col="FAS Value",
        chart_title=f"U.S. Exports to Canada — {sector} Monthly Trend",
        output_path=os.path.join(OUTPUT_FOLDER, f"US_Exports_{sector}_Monthly_Trend.png"),
    )

for sector in sorted(us_imp_monthly["Sector"].unique()):
    save_us_monthly_trendline(
        df_sector=us_imp_monthly[us_imp_monthly["Sector"] == sector],
        value_col="General Customs Value",
        chart_title=f"U.S. Imports from Canada — {sector} Monthly Trend",
        output_path=os.path.join(OUTPUT_FOLDER, f"US_Imports_{sector}_Monthly_Trend.png"),
    )

# HYPOTHESIS TESTING — MANN-WHITNEY U TESTS
# All three RQs use the Mann-Whitney U test rather than a t-test
# because the monthly and annual trade distributions are not
# assumed to be normally distributed.  One-sided alternatives
# match the directional hypotheses in each RQ.

def run_mann_whitney(on_vals, off_vals, alternative):
    """
    Wrapper around scipy's mannwhitneyu that returns the statistic,
    p-value, and a plain-English result string.
    """
    u_stat, p_value = stats.mannwhitneyu(on_vals, off_vals, alternative=alternative)
    result = "REJECT H0" if p_value < 0.05 else "FAIL TO REJECT H0"
    return u_stat, p_value, result


def print_mw_result(sector, on_vals, off_vals, u_stat, p_value, result):
    """Print a formatted Mann-Whitney result block to the console."""
    print(f"\n  {sector}:")
    print(f"    Tariff ON  — N={len(on_vals):>3} | Mean=${on_vals.mean():>15,.0f} | Median=${on_vals.median():>12,.0f}")
    print(f"    Tariff OFF — N={len(off_vals):>3} | Mean=${off_vals.mean():>15,.0f} | Median=${off_vals.median():>12,.0f}")
    print(f"    Mann-Whitney U = {u_stat:,.1f}  |  p-value = {p_value:.4f}")
    print(f"    Result: {result}")


# RQ1 — Did Canada's imports FROM the U.S. decline during
# tariff periods?  (Dataset: U.S. Exports to Canada, monthly)
# H1: monthly values LOWER during tariff-on periods (alternative='less')

print(f"\n{'=' * 65}")
print("  HYPOTHESIS TESTING — RQ1")
print("  Canada Imports FROM U.S. | Dataset: U.S. Exports to Canada")
print(f"{'=' * 65}")

rq1_results = []
for sector in ["Steel", "Aluminum", "Auto"]:
    sec      = us_exp_monthly[us_exp_monthly["Sector"] == sector]
    on_vals  = sec[sec["Tariff"] == "Tariff ON"]["FAS Value"]
    off_vals = sec[sec["Tariff"] == "Tariff OFF"]["FAS Value"]

    u_stat, p_value, result = run_mann_whitney(on_vals, off_vals, alternative="less")
    print_mw_result(sector, on_vals, off_vals, u_stat, p_value, result)

    rq1_results.append({
        "RQ": "RQ1", "Sector": sector,
        "Tariff_ON_N": len(on_vals),   "Tariff_OFF_N": len(off_vals),
        "Tariff_ON_Mean": round(on_vals.mean(), 2),
        "Tariff_OFF_Mean": round(off_vals.mean(), 2),
        "Tariff_ON_Median": round(on_vals.median(), 2),
        "Tariff_OFF_Median": round(off_vals.median(), 2),
        "U_Stat": round(u_stat, 2), "P_Value": round(p_value, 4),
        "Significant": p_value < 0.05, "Result": result,
    })

print(f"\n  {'Sector':<12} {'ON Mean':>15} {'OFF Mean':>15} {'p-value':>10}  Result")
for r in rq1_results:
    print(f"  {r['Sector']:<12} ${r['Tariff_ON_Mean']:>13,.0f} ${r['Tariff_OFF_Mean']:>13,.0f} {r['P_Value']:>10.4f}  {r['Result']}")


# RQ2 — Did Canada's imports from NON-U.S. countries increase
# during tariff years?  (Dataset: Canadian annual Non-US, 2012–2025)
# H1: annual Non-US values HIGHER during tariff years (alternative='greater')
# Annual data is used here because the Canadian source files only
# provide monthly breakdowns for the most recent 24 months. Added to get the annual total for 2025

print(f"\n{'=' * 65}")
print("  HYPOTHESIS TESTING — RQ2")
print("  Canada Non-US Imports | Dataset: Canadian Annual 2012–2025")
print(f"{'=' * 65}")

rq2_results = []
for sector in ["Steel", "Aluminum", "Auto"]:
    sec      = can_nonus_annual[can_nonus_annual["Sector"] == sector]
    on_vals  = sec[sec["Tariff"] == "Tariff ON"]["Value"]
    off_vals = sec[sec["Tariff"] == "Tariff OFF"]["Value"]

    if len(on_vals) < 2 or len(off_vals) < 2:
        print(f"\n  {sector}: Insufficient data (ON n={len(on_vals)}, OFF n={len(off_vals)})")
        continue

    u_stat, p_value, result = run_mann_whitney(on_vals, off_vals, alternative="greater")
    print_mw_result(sector, on_vals, off_vals, u_stat, p_value, result)

    rq2_results.append({
        "RQ": "RQ2", "Sector": sector,
        "Tariff_ON_N": len(on_vals),   "Tariff_OFF_N": len(off_vals),
        "Tariff_ON_Mean": round(on_vals.mean(), 2),
        "Tariff_OFF_Mean": round(off_vals.mean(), 2),
        "U_Stat": round(u_stat, 2), "P_Value": round(p_value, 4),
        "Significant": p_value < 0.05, "Result": result,
    })

print(f"\n  {'Sector':<12} {'ON Mean':>18} {'OFF Mean':>18} {'p-value':>10}  Result")
for r in rq2_results:
    print(f"  {r['Sector']:<12} ${r['Tariff_ON_Mean']:>16,.0f} ${r['Tariff_OFF_Mean']:>16,.0f} {r['P_Value']:>10.4f}  {r['Result']}")


# RQ3 — Did Canada's exports TO the U.S. decline during tariff
# periods?  (Dataset: U.S. Imports from Canada, monthly)
# H1: monthly values LOWER during tariff-on periods (alternative='less')

print(f"\n{'=' * 65}")
print("  HYPOTHESIS TESTING — RQ3")
print("  Canada Exports TO U.S. | Dataset: U.S. Imports from Canada")
print(f"{'=' * 65}")

rq3_results = []
for sector in ["Steel", "Aluminum", "Auto"]:
    sec      = us_imp_monthly[us_imp_monthly["Sector"] == sector]
    on_vals  = sec[sec["Tariff"] == "Tariff ON"]["General Customs Value"]
    off_vals = sec[sec["Tariff"] == "Tariff OFF"]["General Customs Value"]

    u_stat, p_value, result = run_mann_whitney(on_vals, off_vals, alternative="less")
    print_mw_result(sector, on_vals, off_vals, u_stat, p_value, result)

    rq3_results.append({
        "RQ": "RQ3", "Sector": sector,
        "Tariff_ON_N": len(on_vals),   "Tariff_OFF_N": len(off_vals),
        "Tariff_ON_Mean": round(on_vals.mean(), 2),
        "Tariff_OFF_Mean": round(off_vals.mean(), 2),
        "Tariff_ON_Median": round(on_vals.median(), 2),
        "Tariff_OFF_Median": round(off_vals.median(), 2),
        "U_Stat": round(u_stat, 2), "P_Value": round(p_value, 4),
        "Significant": p_value < 0.05, "Result": result,
    })

print(f"\n  {'Sector':<12} {'ON Mean':>15} {'OFF Mean':>15} {'p-value':>10}  Result")
for r in rq3_results:
    print(f"  {r['Sector']:<12} ${r['Tariff_ON_Mean']:>13,.0f} ${r['Tariff_OFF_Mean']:>13,.0f} {r['P_Value']:>10.4f}  {r['Result']}")

# OLS REGRESSION
#
# Estimates the association between tariff periods and trade
# values for all three RQs using the model:
#
#   Trade Value = Intercept + Beta * Tariff_Dummy + error
#
# Tariff_Dummy = 1 during tariff-on periods, 0 otherwise.
# Beta is the estimated dollar change in trade value associated
# with a tariff-on period relative to a tariff-off period.
# A negative Beta supports H1 for RQ1 and RQ3; a positive Beta
# supports H1 for RQ2.
# Added layer to further test the hypothesises

def run_ols(y_vals, x_vals, sector, rq_label):
    """
    Fit a simple OLS regression of trade value on a binary tariff
    dummy.  Uses the normal equations directly (no external library)
    and returns key statistics including Beta, SE, t-stat, p-value,
    and R-squared.  Returns None if there are fewer than 4 observations.
    """
    n = len(y_vals)
    if n < 4:
        print(f"  {sector}: Insufficient data for OLS (n={n})")
        return None

    X = np.column_stack([np.ones(n), x_vals])
    y = np.array(y_vals, dtype=float)

    # Normal equations: coefficients = (X'X)^-1 X'y
    XtX_inv          = np.linalg.pinv(X.T @ X)
    coeffs            = XtX_inv @ X.T @ y
    intercept, beta   = coeffs[0], coeffs[1]

    # Residual statistics and standard error of Beta
    resids  = y - X @ coeffs
    ss_res  = np.sum(resids ** 2)
    ss_tot  = np.sum((y - np.mean(y)) ** 2)
    r2      = 1 - ss_res / ss_tot if ss_tot != 0 else 0
    df_res  = n - 2
    mse     = ss_res / df_res if df_res > 0 else 0
    se_beta = np.sqrt(mse * XtX_inv[1, 1]) if XtX_inv[1, 1] >= 0 else 0

    t_stat  = beta / se_beta if se_beta != 0 else 0
    p_value = 2 * stats.t.sf(abs(t_stat), df=df_res)

    return {
        "RQ": rq_label, "Sector": sector, "N": n,
        "Intercept": round(intercept, 2), "Beta": round(beta, 2),
        "SE_Beta": round(se_beta, 2), "T_Stat": round(t_stat, 4),
        "P_Value": round(p_value, 4), "R_Squared": round(r2, 4),
        "Significant": p_value < 0.05,
        "Result": "Significant" if p_value < 0.05 else "Not Significant",
    }


def print_ols_result(r):
    """Print a single OLS result block to the console."""
    direction = "lower" if r["Beta"] < 0 else "higher"
    print(f"\n  {r['Sector']}:")
    print(f"    N={r['N']} | Intercept=${r['Intercept']:,.0f} | Beta=${r['Beta']:,.0f}")
    print(f"    SE=${r['SE_Beta']:,.0f} | t={r['T_Stat']:.4f} | p={r['P_Value']:.4f} | R²={r['R_Squared']:.4f}")
    print(f"    Trade was ${abs(r['Beta']):,.0f} {direction} per period during tariff-on vs tariff-off")
    print(f"    Result: {r['Result']}")


print(f"\n{'=' * 65}")
print("  OLS REGRESSION — ALL RQs")
print(f"{'=' * 65}")

print("\n  RQ1: U.S. Exports to Canada")
ols_rq1 = [
    r for sector in ["Steel", "Aluminum", "Auto"]
    for r in [run_ols(
        us_exp_monthly[us_exp_monthly["Sector"] == sector]["FAS Value"].values,
        (us_exp_monthly[us_exp_monthly["Sector"] == sector]["Tariff"] == "Tariff ON").astype(int).values,
        sector, "RQ1"
    )] if r
]
for r in ols_rq1:
    print_ols_result(r)

print("\n  RQ2: Canadian Non-US Annual Imports")
ols_rq2 = [
    r for sector in ["Steel", "Aluminum", "Auto"]
    for r in [run_ols(
        can_nonus_annual[can_nonus_annual["Sector"] == sector]["Value"].values,
        (can_nonus_annual[can_nonus_annual["Sector"] == sector]["Tariff"] == "Tariff ON").astype(int).values,
        sector, "RQ2"
    )] if r
]
for r in ols_rq2:
    print_ols_result(r)

print("\n  RQ3: U.S. Imports from Canada")
ols_rq3 = [
    r for sector in ["Steel", "Aluminum", "Auto"]
    for r in [run_ols(
        us_imp_monthly[us_imp_monthly["Sector"] == sector]["General Customs Value"].values,
        (us_imp_monthly[us_imp_monthly["Sector"] == sector]["Tariff"] == "Tariff ON").astype(int).values,
        sector, "RQ3"
    )] if r
]
for r in ols_rq3:
    print_ols_result(r)

print(f"\n  {'RQ':<6} {'Sector':<12} {'Beta':>18} {'p-value':>10} {'R²':>8}  Result")
print(f"  {'-'*6} {'-'*12} {'-'*18} {'-'*10} {'-'*8}  {'-'*16}")
for r in ols_rq1 + ols_rq2 + ols_rq3:
    print(f"  {r['RQ']:<6} {r['Sector']:<12} ${r['Beta']:>16,.0f} {r['P_Value']:>10.4f} {r['R_Squared']:>8.4f}  {r['Result']}")


# PREDICTIVE FORECAST — RQ3
#
# Projects Canada's exports to the U.S. over 24 months (2026–2027)
# under two scenarios:
#   Scenario 1 (Tariffs Remain):  baseline trend + OLS Beta
#   Scenario 2 (Tariffs Removed): baseline trend only
#
# The baseline trend is fit using OLS on the pre-tariff period
# (Jan 2012 – Feb 2018).  Extrapolating that trend forward gives
# the inexact "no tariff" trajectory.  Applying the RQ3
# OLS Beta shifts Scenario 1 down by the estimated tariff effect.

def build_forecast_rq3(df, value_col, sector, ols_beta, output_folder):
    """
    Construct and chart a two-scenario trade forecast for one sector.

    Parameters
    ----------
    df            : us_imp_monthly — monthly U.S. imports from Canada
    value_col     : column holding the trade values
    sector        : 'Steel', 'Aluminum', or 'Auto'
    ols_beta      : tariff effect estimate from the RQ3 OLS regression
    output_folder : directory where the PNG chart is saved

    Returns
    -------
    DataFrame with columns Period, Scenario1_Tariffs_Remain,
    Scenario2_Tariffs_Removed, or None if baseline data is insufficient.
    """
    sec = df[df["Sector"] == sector].sort_values("Period").reset_index(drop=True)

    # Fit a linear trend to the pre-tariff baseline period
    baseline = sec[(sec["Period"] >= BASELINE_START) & (sec["Period"] <= BASELINE_END)]
    if len(baseline) < 6:
        print(f"  {sector}: Not enough baseline data for forecast.")
        return None

    slope, intercept_fit = np.polyfit(np.arange(len(baseline)), baseline[value_col].values, 1)

    # Build forecast x-indices continuing from the end of the actual series
    x_start   = len(sec)
    x_forecast = np.arange(x_start, x_start + FORECAST_MONTHS)

    last_period      = pd.Period(LAST_ACTUAL, freq="M")
    forecast_periods = [str(last_period + i + 1) for i in range(FORECAST_MONTHS)]

    # Scenario 2: baseline trend with no tariff adjustment
    scenario2 = np.clip(intercept_fit + slope * x_forecast, 0, None)
    # Scenario 1: baseline trend shifted down by the tariff effect
    scenario1 = np.clip(scenario2 + ols_beta, 0, None)

    forecast_df = pd.DataFrame({
        "Period":                    forecast_periods,
        "Scenario1_Tariffs_Remain":  scenario1,
        "Scenario2_Tariffs_Removed": scenario2,
    })

    # Chart — show the most recent 36 months of actuals plus the forecast
    recent         = sec.tail(36)
    actual_periods = recent["Period"].tolist()
    actual_values  = recent[value_col].tolist()
    all_periods    = actual_periods + forecast_periods
    n_actual       = len(actual_periods)

    fig, ax = plt.subplots(figsize=(13, 6))

    ax.plot(range(n_actual), actual_values,
            color="#2563EB", linewidth=1.8, label="Actual Trade Value")

    ax.plot(range(n_actual - 1, n_actual + FORECAST_MONTHS),
            [actual_values[-1]] + list(scenario1),
            color="#DC2626", linewidth=2, linestyle="--",
            label="Scenario 1: Tariffs Remain")

    ax.plot(range(n_actual - 1, n_actual + FORECAST_MONTHS),
            [actual_values[-1]] + list(scenario2),
            color="#16A34A", linewidth=2, linestyle="--",
            label="Scenario 2: Tariffs Removed")

    # Vertical line and shaded region marking the forecast boundary
    ax.axvline(x=n_actual - 1, color="gray", linestyle=":", linewidth=1.2)
    ax.axvspan(n_actual - 1, n_actual + FORECAST_MONTHS - 1, alpha=0.06, color="gray")
    ax.text(
        n_actual + FORECAST_MONTHS / 2,
        ax.get_ylim()[1] * 0.98 if ax.get_ylim()[1] != 0 else 1,
        "Forecast\n(2026-2027)", ha="center", va="top", fontsize=9, color="gray"
    )

    tick_pos = list(range(0, len(all_periods), 6))
    ax.set_xticks(tick_pos)
    ax.set_xticklabels([all_periods[i] for i in tick_pos], rotation=45, fontsize=8)
    ax.yaxis.set_major_formatter(FuncFormatter(compact_currency))
    ax.set_title(f"Canada Exports to U.S. — {sector}\nTwo-Scenario Forecast (2026-2027)",
                 fontsize=12, fontweight="bold", pad=12)
    ax.set_xlabel("Period", fontsize=10)
    ax.set_ylabel("Trade Value (USD)", fontsize=10)
    ax.legend(loc="upper left", fontsize=9)
    ax.grid(axis="y", linestyle="--", alpha=0.4)

    plt.tight_layout()
    save_path = os.path.join(output_folder, f"Forecast_Canada_Exports_{sector}.png")
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved: Forecast_Canada_Exports_{sector}.png")

    return forecast_df


print(f"\n{'=' * 65}")
print("  PREDICTIVE FORECAST — RQ3")
print("  Canada Exports to U.S. | Horizon: 24 months (2026–2027)")
print(f"{'=' * 65}")

ols_betas_rq3 = {r["Sector"]: r["Beta"] for r in ols_rq3}

forecast_rq3 = {}
for sector in ["Steel", "Aluminum", "Auto"]:
    beta = ols_betas_rq3.get(sector, 0)
    print(f"\n  {sector} (OLS Beta = ${beta:,.0f})")
    forecast_rq3[sector] = build_forecast_rq3(
        df=us_imp_monthly, value_col="General Customs Value",
        sector=sector, ols_beta=beta, output_folder=OUTPUT_FOLDER
    )

print(f"\n  {'Sector':<12} {'Period':<10} {'Tariffs Remain':>20} {'Tariffs Removed':>20}")
print(f"  {'-'*12} {'-'*10} {'-'*20} {'-'*20}")
for sector, fdf in forecast_rq3.items():
    if fdf is None:
        continue
    for row in [fdf.iloc[0], fdf.iloc[-1]]:
        print(f"  {sector:<12} {row['Period']:<10} "
              f"${row['Scenario1_Tariffs_Remain']:>18,.0f} "
              f"${row['Scenario2_Tariffs_Removed']:>18,.0f}")

print(f"\n{'=' * 65}")
print(f"  Analysis complete.  Output folder: {OUTPUT_FOLDER}")
print(f"{'=' * 65}")