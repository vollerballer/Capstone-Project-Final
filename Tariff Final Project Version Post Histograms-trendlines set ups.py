"""
=============================================================
MIS581 Capstone Project — Tariff Final Project
Main Script: Step 1–7
=============================================================
Place this file in the Inputs folder alongside all data files.
Run from terminal: python "Tariff Final Project.py"
=============================================================
"""

# =============================================================
# IMPORT LIBRARIES
# =============================================================
import os
import warnings
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

warnings.filterwarnings("ignore")

# =============================================================
# PATHS / CONFIG
# =============================================================
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FOLDER = os.path.join(BASE_PATH, "Outputs")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# =============================================================
# INPUT FILE LOCATIONS
# =============================================================
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

# =============================================================
# SECTOR DEFINITIONS
# =============================================================
SECTORS = {
    "Steel":    [720825, 720851, 720916, 721070, 721310],
    "Aluminum": [760110, 760120, 760421, 760612, 760711],
    "Auto":     [870323, 870324, 870421, 870840, 870899],
}

# =============================================================
# MONTH NAME MAP
# =============================================================
MONTH_MAP = {
    "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
    "May": "05", "June": "06", "Jun": "06", "Jul": "07",
    "Aug": "08", "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12",
}

# =============================================================
# COUNTRY NAMES TO TREAT AS "US"
# =============================================================
US_COUNTRY_NAMES = {
    "United States",
    "United States of America",
    "U.S.",
    "USA",
    "US"
}

# =============================================================
# HELPER FUNCTIONS
# =============================================================

def money_fmt(x):
    """Format a numeric value as whole-dollar currency."""
    return f"${x:,.0f}"


def extract_hs_code(s):
    """Try to extract a 6-digit HS code from a text string."""
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
    Identify whether a raw Excel row appears to contain years.
    Returns (start_column_index, [list_of_years]) or None.
    """
    hits = []
    for ci, v in enumerate(vals):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            continue
        try:
            iv = int(v)
            if 2010 <= iv <= 2030:
                hits.append((ci, iv))
        except (ValueError, TypeError, OverflowError):
            pass

    if len(hits) >= 3:
        return hits[0][0], [y for _, y in hits]
    return None


def safe_sheet_name(name, max_len=31):
    """Make a string safe for Excel sheet names."""
    invalid = ["\\", "/", "*", "[", "]", ":", "?"]
    for ch in invalid:
        name = name.replace(ch, "_")
    return name[:max_len]


def parse_period(s):
    """Convert a header like '2025 Jan' into '2025-01'."""
    parts = str(s).strip().split()
    if len(parts) == 2:
        month = MONTH_MAP.get(parts[1], "00")
        return f"{parts[0]}-{month}" if month != "00" else None
    return None


def summarize_canadian(df, label, time_col):
    """Print a simple summary for a Canadian dataset after loading."""
    time_vals = sorted(df[time_col].unique().tolist())
    print(f"{label}  {df.shape[0]:,} rows | {time_col}: {time_vals[0]} to {time_vals[-1]}")
    print(f"  HS Codes found:  {sorted(df['HS Code'].unique().tolist())}")


def classify_country_group(country_name):
    """Classify a country into US vs Non-US."""
    country = str(country_name).strip()
    return "US" if country in US_COUNTRY_NAMES else "Non-US"


# =============================================================
# AXIS FORMATTING HELPERS
# =============================================================

def get_compact_scale(values):
    """
    Determine the best display scale for a set of values.
    Returns:
        scale, suffix, unit_label
    """
    arr = np.asarray(values).astype(float).ravel()
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
    else:
        return 1, "", "US Dollars"


def make_currency_formatter(scale, suffix):
    """Build a compact currency formatter for matplotlib axes."""
    def _formatter(x, pos):
        if scale == 1:
            return f"${x:,.0f}"

        scaled = x / scale

        if abs(scaled) >= 100:
            txt = f"{scaled:,.0f}"
        elif abs(scaled) >= 10:
            txt = f"{scaled:,.1f}".rstrip("0").rstrip(".")
        else:
            txt = f"{scaled:,.2f}".rstrip("0").rstrip(".")

        return f"${txt}{suffix}"

    return _formatter


def apply_compact_currency_axis(ax, values, axis="y"):
    """
    Apply compact currency formatting to the x or y axis and
    hide scientific notation.
    """
    scale, suffix, unit_label = get_compact_scale(values)
    formatter = FuncFormatter(make_currency_formatter(scale, suffix))

    if axis == "y":
        ax.yaxis.set_major_formatter(formatter)
        ax.yaxis.offsetText.set_visible(False)
    else:
        ax.xaxis.set_major_formatter(formatter)
        ax.xaxis.offsetText.set_visible(False)

    return unit_label


# =============================================================
# DATA LOADER FUNCTIONS
# =============================================================

def load_us_monthly(filepath, value_col):
    """
    Load a U.S. monthly trade file and map HTS Number into sectors.
    """
    df = pd.read_excel(filepath, sheet_name=1)
    df = df.dropna(subset=["HTS Number", "Year", "Month"])

    df["HTS Number"] = df["HTS Number"].astype(int)
    df["Year"] = df["Year"].astype(int)
    df["Month"] = df["Month"].astype(int)
    df["Period"] = df["Year"].astype(str) + "-" + df["Month"].astype(str).str.zfill(2)
    df[value_col] = pd.to_numeric(df[value_col], errors="coerce").fillna(0)

    code_to_sector = {code: sector for sector, codes in SECTORS.items() for code in codes}
    df["Sector"] = df["HTS Number"].map(code_to_sector).fillna("Other")

    return df[df["Sector"] != "Other"].copy()


def load_canadian_yoy(filepath, sector_name):
    """
    Load a Canadian YoY annual file from its non-tabular layout.
    """
    df_raw = pd.read_excel(filepath, sheet_name=0, header=None)

    records = []
    year_cols = None
    year_start = None
    in_data = False
    current_hs = None

    for _, row in df_raw.iterrows():
        vals = row.tolist()
        col_a = str(vals[0]).strip() if vals[0] is not None and str(vals[0]) != "nan" else ""

        year_result = find_years_in_row(vals)
        if year_result is not None:
            year_start, year_cols = year_result
            in_data = True
            if current_hs is None:
                current_hs = "All"
            continue

        if col_a.upper().startswith("HS "):
            code = extract_hs_code(col_a)
            if code:
                current_hs = code
            continue

        if col_a in {"Title", "Products", "Origin", "Destination", "Period", "Units", "Source", "Note", ""}:
            continue
        if col_a.startswith(("Source", "Note")):
            continue
        if "Sub-Total" in col_a or "Total All Countries" in col_a:
            continue

        if in_data and year_cols and col_a and current_hs is not None:
            for j, year in enumerate(year_cols):
                col_idx = year_start + j
                if col_idx < len(vals):
                    val = vals[col_idx]
                    if val is not None and not (isinstance(val, float) and np.isnan(val)):
                        try:
                            records.append({
                                "Sector": sector_name,
                                "HS Code": current_hs,
                                "Country": col_a,
                                "Year": int(year),
                                "Value": float(val),
                            })
                        except (ValueError, TypeError):
                            pass

    return pd.DataFrame(records)


def load_canadian_l24(filepath, sector_name):
    """
    Load a Canadian Last 24 Months file from its non-tabular layout.
    """
    df_raw = pd.read_excel(filepath, sheet_name=0, header=None)

    records = []
    period_cols = None
    in_data = False
    current_hs = None

    for _, row in df_raw.iterrows():
        vals = row.tolist()
        col_a = str(vals[0]).strip() if vals[0] is not None and str(vals[0]) != "nan" else ""
        col_b = str(vals[1]).strip() if len(vals) > 1 and vals[1] is not None and str(vals[1]) != "nan" else ""

        if any(m in col_b for m in MONTH_MAP):
            period_cols = [
                parse_period(str(v))
                for v in vals[1:]
                if v is not None and str(v) != "nan" and parse_period(str(v))
            ]
            in_data = True
            if current_hs is None:
                current_hs = "All"
            continue

        if col_a.upper().startswith("HS "):
            code = extract_hs_code(col_a)
            if code:
                current_hs = code
            continue

        if col_a in {"Title", "Products", "Origin", "Destination", "Period", "Units", "Source", "Note", ""}:
            continue
        if col_a.startswith(("Source", "Note")):
            continue
        if "Sub-Total" in col_a or "Total All Countries" in col_a:
            continue

        if in_data and period_cols and col_a and current_hs is not None:
            for j, period in enumerate(period_cols):
                col_idx = j + 1
                if col_idx < len(vals):
                    val = vals[col_idx]
                    if val is not None and not (isinstance(val, float) and np.isnan(val)):
                        try:
                            records.append({
                                "Sector": sector_name,
                                "HS Code": current_hs,
                                "Country": col_a,
                                "Period": period,
                                "Value": float(val),
                            })
                        except (ValueError, TypeError):
                            pass

    return pd.DataFrame(records)


# =============================================================
# STEP 1 — LOAD ALL DATA
# =============================================================
print("=" * 55)
print("  STEP 1: LOADING ALL DATA")
print(f"  Folder: {BASE_PATH}")
print("=" * 55)

print("\n--- US Monthly Files ---")
df_exp = load_us_monthly(FILES["us_exports"], "FAS Value")
print(f"US Exports loaded:   {df_exp.shape[0]:,} rows | {df_exp.Period.min()} to {df_exp.Period.max()}")
print(f"  Sectors:           {df_exp.groupby('Sector').size().to_dict()}")
print(f"  HS Codes found:    {sorted(df_exp['HTS Number'].unique().tolist())}")

df_imp = load_us_monthly(FILES["us_imports"], "General Customs Value")
print(f"US Imports loaded:   {df_imp.shape[0]:,} rows | {df_imp.Period.min()} to {df_imp.Period.max()}")
print(f"  Sectors:           {df_imp.groupby('Sector').size().to_dict()}")
print(f"  HS Codes found:    {sorted(df_imp['HTS Number'].unique().tolist())}")

print("\n--- Canadian YoY Annual Files ---")
steel_yoy = load_canadian_yoy(FILES["steel_yoy"], "Steel")
summarize_canadian(steel_yoy, "Steel YoY loaded:   ", "Year")

alum_yoy = load_canadian_yoy(FILES["alum_yoy"], "Aluminum")
summarize_canadian(alum_yoy, "Aluminum YoY loaded:", "Year")

auto_yoy = load_canadian_yoy(FILES["auto_yoy"], "Auto")
summarize_canadian(auto_yoy, "Auto YoY loaded:    ", "Year")

print("\n--- Canadian Last 24 Months Monthly Files ---")
steel_l24 = load_canadian_l24(FILES["steel_l24"], "Steel")
summarize_canadian(steel_l24, "Steel L24 loaded:   ", "Period")

alum_l24 = load_canadian_l24(FILES["alum_l24"], "Aluminum")
summarize_canadian(alum_l24, "Aluminum L24 loaded:", "Period")

auto_l24 = load_canadian_l24(FILES["auto_l24"], "Auto")
summarize_canadian(auto_l24, "Auto L24 loaded:    ", "Period")

print("\n" + "=" * 55)
print("  STEP 1 COMPLETE - All data loaded successfully")
print("  Ready for Step 2")
print("=" * 55)


# =============================================================
# STEP 2 — BUILD CORE AGGREGATIONS
# =============================================================

def build_monthly_sector_totals(df, sector_col, period_col, value_col):
    """Aggregate a dataframe to one row per sector per month."""
    return (
        df.groupby([sector_col, period_col], as_index=False)[value_col]
        .sum()
        .sort_values([sector_col, period_col])
        .copy()
    )


def build_annual_sector_totals(df, sector_col, year_col, value_col):
    """Aggregate a dataframe to one row per sector per year."""
    return (
        df.groupby([sector_col, year_col], as_index=False)[value_col]
        .sum()
        .sort_values([sector_col, year_col])
        .copy()
    )


# U.S. monthly totals
us_exp_monthly = build_monthly_sector_totals(
    df=df_exp, sector_col="Sector", period_col="Period", value_col="FAS Value"
)

us_imp_monthly = build_monthly_sector_totals(
    df=df_imp, sector_col="Sector", period_col="Period", value_col="General Customs Value"
)

# Canadian base datasets
can_yoy_all = pd.concat([steel_yoy, alum_yoy, auto_yoy], ignore_index=True).copy()
can_l24_all = pd.concat([steel_l24, alum_l24, auto_l24], ignore_index=True).copy()

# Canadian annual totals 2012–2024 from YoY
can_yoy_totals = build_annual_sector_totals(
    df=can_yoy_all, sector_col="Sector", year_col="Year", value_col="Value"
)

# Canadian annual totals for 2025 built from L24 monthly rows
can_l24_2025 = can_l24_all[
    can_l24_all["Period"].astype(str).str.startswith("2025-")
].copy()

can_2025_totals = (
    can_l24_2025.groupby("Sector", as_index=False)["Value"]
    .sum()
)

can_2025_totals["Year"] = 2025
can_2025_totals = can_2025_totals[["Sector", "Year", "Value"]]

# Final Canadian annual totals 2012–2025
canadian_annual_totals = pd.concat(
    [can_yoy_totals[["Sector", "Year", "Value"]], can_2025_totals],
    ignore_index=True
).sort_values(["Sector", "Year"]).reset_index(drop=True)

print("\n" + "=" * 60)
print("  CANADIAN ANNUAL TOTALS INCLUDING 2025")
print("=" * 60)
print(f"Rows: {len(canadian_annual_totals):,}")
print(f"Sectors: {sorted(canadian_annual_totals['Sector'].unique().tolist())}")
print(f"Years: {canadian_annual_totals['Year'].min()} to {canadian_annual_totals['Year'].max()}")

# Canadian annual totals by source group: US vs Non-US
can_yoy_all["Country Group"] = can_yoy_all["Country"].apply(classify_country_group)
can_l24_all["Country Group"] = can_l24_all["Country"].apply(classify_country_group)

canadian_us_nonus_yoy = (
    can_yoy_all.groupby(["Sector", "Year", "Country Group"], as_index=False)["Value"]
    .sum()
    .sort_values(["Sector", "Year", "Country Group"])
    .copy()
)

can_l24_2025_grouped = can_l24_all[
    can_l24_all["Period"].astype(str).str.startswith("2025-")
].copy()

canadian_us_nonus_2025 = (
    can_l24_2025_grouped.groupby(["Sector", "Country Group"], as_index=False)["Value"]
    .sum()
)

canadian_us_nonus_2025["Year"] = 2025
canadian_us_nonus_2025 = canadian_us_nonus_2025[["Sector", "Year", "Country Group", "Value"]]

canadian_us_nonus_annual = pd.concat(
    [canadian_us_nonus_yoy, canadian_us_nonus_2025],
    ignore_index=True
).sort_values(["Sector", "Year", "Country Group"]).reset_index(drop=True)

print("\n" + "=" * 60)
print("  CANADIAN ANNUAL TOTALS: US VS NON-US")
print("=" * 60)
print(f"Rows: {len(canadian_us_nonus_annual):,}")
print(f"Sectors: {sorted(canadian_us_nonus_annual['Sector'].unique().tolist())}")
print(f"Years: {canadian_us_nonus_annual['Year'].min()} to {canadian_us_nonus_annual['Year'].max()}")
print(f"Country Groups: {sorted(canadian_us_nonus_annual['Country Group'].unique().tolist())}")


# =============================================================
# STEP 3 — QUICK OVERVIEW TABLES
# =============================================================

def build_quick_overview_monthly(df, value_col, hs_col, sector_col="Sector", period_col="Period"):
    """Build descriptive-statistics overview table for monthly data."""
    out = {}

    for sector in sorted(df[sector_col].dropna().unique()):
        sub = df[df[sector_col] == sector].copy()

        monthly = (
            sub.groupby(period_col, as_index=False)[value_col]
            .sum()
            .sort_values(period_col)
        )

        overview = pd.DataFrame({
            "Metric": [
                "Records (row count)",
                "Unique HS codes",
                "Unique periods",
                "Start month",
                "End month",
                "Total US Dollars (all records)",
                "Average monthly total",
                "Median monthly total",
                "Min monthly total",
                "Max monthly total",
                "Std dev monthly total"
            ],
            "Value": [
                f"{len(sub):,}",
                f"{sub[hs_col].nunique():,}",
                f"{sub[period_col].nunique():,}",
                str(monthly[period_col].min()),
                str(monthly[period_col].max()),
                money_fmt(sub[value_col].sum()),
                money_fmt(monthly[value_col].mean()),
                money_fmt(monthly[value_col].median()),
                money_fmt(monthly[value_col].min()),
                money_fmt(monthly[value_col].max()),
                money_fmt(monthly[value_col].std(ddof=1) if len(monthly) > 1 else 0),
            ]
        })

        out[sector] = {
            "overview": overview,
            "monthly_totals": monthly
        }

    return out


def build_quick_overview_annual(df, value_col="Value", sector_col="Sector", year_col="Year"):
    """Build descriptive-statistics overview table for annual data."""
    out = {}

    for sector in sorted(df[sector_col].dropna().unique()):
        sub = df[df[sector_col] == sector].copy()

        annual = (
            sub.groupby(year_col, as_index=False)[value_col]
            .sum()
            .sort_values(year_col)
        )

        overview = pd.DataFrame({
            "Metric": [
                "Records (row count)",
                "Unique years",
                "Start year",
                "End year",
                "Total US Dollars (all records)",
                "Average annual total",
                "Median annual total",
                "Min annual total",
                "Max annual total",
                "Std dev annual total"
            ],
            "Value": [
                f"{len(sub):,}",
                f"{sub[year_col].nunique():,}",
                str(annual[year_col].min()),
                str(annual[year_col].max()),
                money_fmt(sub[value_col].sum()),
                money_fmt(annual[value_col].mean()),
                money_fmt(annual[value_col].median()),
                money_fmt(annual[value_col].min()),
                money_fmt(annual[value_col].max()),
                money_fmt(annual[value_col].std(ddof=1) if len(annual) > 1 else 0),
            ]
        })

        out[sector] = {
            "overview": overview,
            "annual_totals": annual
        }

    return out


def print_overview_results(results_dict, title):
    """Print overview tables in a readable console format."""
    print("\n" + "=" * 65)
    print(f"  {title}")
    print("=" * 65)

    for sector, result in results_dict.items():
        print(f"\n--- {sector} ---")
        print(result["overview"].to_string(index=False))


us_exports_overview = build_quick_overview_monthly(
    df_exp, value_col="FAS Value", hs_col="HTS Number"
)

us_imports_overview = build_quick_overview_monthly(
    df_imp, value_col="General Customs Value", hs_col="HTS Number"
)

canadian_annual_overview = build_quick_overview_annual(
    canadian_annual_totals, value_col="Value", sector_col="Sector", year_col="Year"
)

print_overview_results(us_exports_overview, "US EXPORTS — QUICK OVERVIEW BY SECTOR")
print_overview_results(us_imports_overview, "US IMPORTS — QUICK OVERVIEW BY SECTOR")
print_overview_results(canadian_annual_overview, "CANADIAN ANNUAL IMPORTS — QUICK OVERVIEW BY SECTOR (2012–2025)")


# =============================================================
# STEP 4 — EXPORT SOURCE DATA
# =============================================================
export_path = os.path.join(OUTPUT_FOLDER, "Histogram_Source_Data.xlsx")

with pd.ExcelWriter(export_path, engine="openpyxl") as writer:
    us_exp_monthly.to_excel(writer, sheet_name="US_Exports_Monthly", index=False)
    us_imp_monthly.to_excel(writer, sheet_name="US_Imports_Monthly", index=False)
    canadian_annual_totals.to_excel(writer, sheet_name="Canada_Annual_2012_2025", index=False)
    canadian_us_nonus_annual.to_excel(writer, sheet_name="Canada_US_vs_NonUS", index=False)

    for sector in sorted(us_exp_monthly["Sector"].unique()):
        us_exp_monthly[us_exp_monthly["Sector"] == sector].to_excel(
            writer,
            sheet_name=safe_sheet_name(f"US_EXP_{sector}"),
            index=False
        )

    for sector in sorted(us_imp_monthly["Sector"].unique()):
        us_imp_monthly[us_imp_monthly["Sector"] == sector].to_excel(
            writer,
            sheet_name=safe_sheet_name(f"US_IMP_{sector}"),
            index=False
        )

    # IMPORTANT:
    # Export CAN_ tabs from canadian_us_nonus_annual so each year has
    # both a US row and a Non-US row.
    for sector in sorted(canadian_us_nonus_annual["Sector"].unique()):
        canadian_us_nonus_annual[canadian_us_nonus_annual["Sector"] == sector].to_excel(
            writer,
            sheet_name=safe_sheet_name(f"CAN_{sector}"),
            index=False
        )

    for sector in sorted(canadian_us_nonus_annual["Sector"].unique()):
        canadian_us_nonus_annual[canadian_us_nonus_annual["Sector"] == sector].to_excel(
            writer,
            sheet_name=safe_sheet_name(f"Trend_{sector}"),
            index=False
        )

print(f"\nSaved: {export_path}")


# =============================================================
# STEP 5 — HISTOGRAMS
# =============================================================

def save_histogram(series, title, output_path, bins=15):
    """
    Save a histogram with compact currency formatting on the x-axis.
    """
    vals = pd.to_numeric(series, errors="coerce").dropna()

    if vals.empty:
        print(f"Skipped: {title} (no numeric data)")
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

    print(f"Saved histogram: {os.path.basename(output_path)}")


# ---------------------------
# U.S. export histograms
# ---------------------------
for sector in sorted(us_exp_monthly["Sector"].unique()):
    sub = us_exp_monthly[us_exp_monthly["Sector"] == sector].copy()
    save_histogram(
        series=sub["FAS Value"],
        title=f"US Exports to Canada - {sector} Monthly Totals",
        output_path=os.path.join(OUTPUT_FOLDER, f"US_Exports_{sector}_Monthly_Histogram.png"),
        bins=15
    )

# ---------------------------
# U.S. import histograms
# ---------------------------
for sector in sorted(us_imp_monthly["Sector"].unique()):
    sub = us_imp_monthly[us_imp_monthly["Sector"] == sector].copy()
    save_histogram(
        series=sub["General Customs Value"],
        title=f"US Imports from Canada - {sector} Monthly Totals",
        output_path=os.path.join(OUTPUT_FOLDER, f"US_Imports_{sector}_Monthly_Histogram.png"),
        bins=15
    )

# ---------------------------
# Canadian annual histograms using US + Non-US together
# One histogram per sector, with both source groups included
# as separate observations to increase the number of points.
# ---------------------------
for sector in sorted(canadian_us_nonus_annual["Sector"].unique()):
    sub = canadian_us_nonus_annual[
        canadian_us_nonus_annual["Sector"] == sector
    ].copy()

    save_histogram(
        series=sub["Value"],
        title=f"Canadian Imports - {sector} Annual Totals (US + Non-US)",
        output_path=os.path.join(OUTPUT_FOLDER, f"Canadian_Imports_{sector}_Annual_Histogram.png"),
        bins=8
    )


# =============================================================
# STEP 6 — CANADIAN TREND LINES: US VS NON-US
# =============================================================

def save_sector_trendline(df_sector, sector_name, output_folder):
    """
    Save an annual trend line chart for Canadian imports by source group:
      - US
      - Non-US
    """
    chart_df = (
        df_sector.pivot(index="Year", columns="Country Group", values="Value")
        .fillna(0)
        .sort_index()
    )

    years = chart_df.index.tolist()

    fig, ax = plt.subplots(figsize=(11, 6))

    if "US" in chart_df.columns:
        ax.plot(years, chart_df["US"], marker="o", linewidth=2, label="US")
    if "Non-US" in chart_df.columns:
        ax.plot(years, chart_df["Non-US"], marker="o", linewidth=2, label="Non-US")

    ax.set_title(f"Canadian Imports by Source Group - {sector_name} (Annual Totals)")
    ax.set_xlabel("Year")
    ax.set_ylabel("US Dollars")
    ax.set_xticks(years)
    ax.tick_params(axis="x", rotation=45)
    ax.legend()

    apply_compact_currency_axis(ax, chart_df.values, axis="y")

    plt.tight_layout()

    output_path = os.path.join(output_folder, f"Canadian_Imports_{sector_name}_US_vs_NonUS_Trend.png")
    plt.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close()

    print(f"Saved trend line: {os.path.basename(output_path)}")


for sector in sorted(canadian_us_nonus_annual["Sector"].unique()):
    df_sector = canadian_us_nonus_annual[canadian_us_nonus_annual["Sector"] == sector].copy()
    save_sector_trendline(df_sector, sector, OUTPUT_FOLDER)


# =============================================================
# STEP 7 — U.S. MONTHLY TREND LINES WITH 12-MONTH ROLLING AVG
# =============================================================

def save_us_monthly_trendline(df_sector, value_col, chart_title, output_path):
    """
    Save a U.S. monthly trend line chart with:
      - raw monthly totals
      - 12-month rolling average
    """
    chart_df = df_sector.copy()
    chart_df = chart_df.sort_values("Period").reset_index(drop=True)

    chart_df["Rolling12"] = chart_df[value_col].rolling(window=12, min_periods=1).mean()

    fig, ax = plt.subplots(figsize=(11, 6))

    ax.plot(
        chart_df["Period"],
        chart_df[value_col],
        marker="o",
        linewidth=1.5,
        alpha=0.6,
        label="Monthly Total"
    )

    ax.plot(
        chart_df["Period"],
        chart_df["Rolling12"],
        linewidth=3,
        label="12-Month Rolling Average"
    )

    ax.set_title(chart_title)
    ax.set_xlabel("Period")
    ax.set_ylabel("US Dollars")

    tick_positions = list(range(0, len(chart_df), 12))
    if len(chart_df) > 0 and (len(chart_df) - 1) not in tick_positions:
        tick_positions.append(len(chart_df) - 1)

    ax.set_xticks(tick_positions)
    ax.set_xticklabels(
        [chart_df["Period"].iloc[i] for i in tick_positions],
        rotation=45
    )

    ax.legend()

    combined_vals = np.concatenate([
        chart_df[value_col].values,
        chart_df["Rolling12"].values
    ])
    apply_compact_currency_axis(ax, combined_vals, axis="y")

    plt.tight_layout()
    plt.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close()

    print(f"Saved trend line: {os.path.basename(output_path)}")


# ---------------------------
# U.S. export monthly trends
# ---------------------------
for sector in sorted(us_exp_monthly["Sector"].unique()):
    df_sector = us_exp_monthly[us_exp_monthly["Sector"] == sector].copy()

    save_us_monthly_trendline(
        df_sector=df_sector,
        value_col="FAS Value",
        chart_title=f"U.S. Exports to Canada - {sector} Monthly Trend",
        output_path=os.path.join(OUTPUT_FOLDER, f"US_Exports_{sector}_Monthly_Trend.png")
    )

# ---------------------------
# U.S. import monthly trends
# ---------------------------
for sector in sorted(us_imp_monthly["Sector"].unique()):
    df_sector = us_imp_monthly[us_imp_monthly["Sector"] == sector].copy()

    save_us_monthly_trendline(
        df_sector=df_sector,
        value_col="General Customs Value",
        chart_title=f"U.S. Imports from Canada - {sector} Monthly Trend",
        output_path=os.path.join(OUTPUT_FOLDER, f"US_Imports_{sector}_Monthly_Trend.png")
    )


# =============================================================
# FINAL STATUS
# =============================================================
print("\n" + "=" * 60)
print("  SCRIPT COMPLETE")
print("=" * 60)
print(f"Output folder: {OUTPUT_FOLDER}")
print("U.S. monthly histograms created.")
print("U.S. monthly trend lines created with 12-month rolling averages.")
print("Canadian annual histograms created with US + Non-US combined in each sector.")
print("Canadian US vs Non-US annual trend lines created with 2025 included.")
print("CAN_ export tabs now include both US and Non-US rows for each year.")
print("All chart axes now use 'US Dollars' and compact currency labels.")
print("=" * 60)

# =============================================================
# STEP 8 — HYPOTHESIS TESTING: RQ1
# Did Canada's monthly imports FROM the U.S. decline
# during tariff periods vs non-tariff periods?
# Dataset: U.S. Exports to Canada (us_exp_monthly)
# =============================================================

from scipy import stats

# -----------------------------------------------------------
# TARIFF PERIOD FLAG
# -----------------------------------------------------------
# Tariff ON: March 2018 - April 2019 AND March 2025 - Jan 2026
# Everything else is Tariff OFF

TARIFF_ON_PERIODS = [
    ("2018-03", "2019-04"),
    ("2025-03", "2026-01"),
]

def flag_tariff(period_str):
    for start, end in TARIFF_ON_PERIODS:
        if start <= period_str <= end:
            return "Tariff ON"
    return "Tariff OFF"

us_exp_monthly["Tariff"] = us_exp_monthly["Period"].apply(flag_tariff)

# -----------------------------------------------------------
# RQ1 MANN-WHITNEY U TEST — BY SECTOR
# H0: No significant difference in monthly trade values
#     between tariff-on and tariff-off periods
# H1: Monthly trade values are LOWER during tariff-on periods
# One-sided test: alternative = 'less'
# -----------------------------------------------------------

print("\n" + "=" * 65)
print("  STEP 8: HYPOTHESIS TESTING — RQ1")
print("  Canada Imports FROM U.S. (U.S. Exports to Canada)")
print("  H1: Monthly values LOWER during tariff-on periods")
print("=" * 65)

rq1_results = []

for sector in ["Steel", "Aluminum", "Auto"]:
    sec = us_exp_monthly[us_exp_monthly["Sector"] == sector].copy()

    on_vals  = sec[sec["Tariff"] == "Tariff ON"]["FAS Value"]
    off_vals = sec[sec["Tariff"] == "Tariff OFF"]["FAS Value"]

    # Descriptive stats by tariff period
    on_mean   = on_vals.mean()
    off_mean  = off_vals.mean()
    on_median = on_vals.median()
    off_mean  = off_vals.mean()
    on_n      = len(on_vals)
    off_n     = len(off_vals)

    # Mann-Whitney U — one-sided (less: testing if ON < OFF)
    u_stat, p_value = stats.mannwhitneyu(on_vals, off_vals, alternative="less")

    # Determine result
    reject = p_value < 0.05
    result = "REJECT H0" if reject else "FAIL TO REJECT H0"

    print(f"\n  {sector}:")
    print(f"    Tariff ON  — N={on_n:>3} | Mean=${on_vals.mean():>12,.0f} | Median=${on_vals.median():>12,.0f}")
    print(f"    Tariff OFF — N={off_n:>3} | Mean=${off_vals.mean():>12,.0f} | Median=${off_vals.median():>12,.0f}")
    print(f"    Mann-Whitney U = {u_stat:,.1f} | p-value = {p_value:.4f}")
    print(f"    Result: {result}")

    rq1_results.append({
        "RQ":               "RQ1",
        "Sector":           sector,
        "Tariff_ON_N":      on_n,
        "Tariff_OFF_N":     off_n,
        "Tariff_ON_Mean":   round(on_vals.mean(), 2),
        "Tariff_OFF_Mean":  round(off_vals.mean(), 2),
        "Tariff_ON_Median": round(on_vals.median(), 2),
        "Tariff_OFF_Median":round(off_vals.median(), 2),
        "U_Stat":           round(u_stat, 2),
        "P_Value":          round(p_value, 4),
        "Significant":      reject,
        "Result":           result,
    })

print("\n" + "-" * 65)
print("  RQ1 SUMMARY")
print("-" * 65)
print(f"  {'Sector':<12} {'ON Mean':>15} {'OFF Mean':>15} {'p-value':>10} {'Result'}")
print(f"  {'-'*12} {'-'*15} {'-'*15} {'-'*10} {'-'*20}")
for r in rq1_results:
    print(f"  {r['Sector']:<12} ${r['Tariff_ON_Mean']:>13,.0f} ${r['Tariff_OFF_Mean']:>13,.0f} {r['P_Value']:>10.4f} {r['Result']}")

print("\n  RQ1 complete. Results stored in rq1_results.")
print("  Ready for RQ2.")

# =============================================================
# STEP 9 — HYPOTHESIS TESTING: RQ2
# Did Canada's imports from NON-U.S. countries INCREASE
# during tariff periods vs non-tariff periods?
# Dataset: Canadian Annual Data 2012-2025 (Non-US rows only)
# =============================================================

# -----------------------------------------------------------
# TARIFF YEAR FLAG
# Tariff ON years: 2018, 2019 (Episode 1) and 2025 (Episode 2)
# Everything else is Tariff OFF
# -----------------------------------------------------------

TARIFF_ON_YEARS = [2018, 2019, 2025]

def flag_tariff_year(year):
    return "Tariff ON" if int(year) in TARIFF_ON_YEARS else "Tariff OFF"

canadian_us_nonus_annual["Tariff"] = canadian_us_nonus_annual["Year"].apply(flag_tariff_year)

# -----------------------------------------------------------
# ISOLATE NON-US ROWS
# -----------------------------------------------------------
can_nonus_annual = canadian_us_nonus_annual[
    canadian_us_nonus_annual["Country Group"] == "Non-US"
].copy()

# -----------------------------------------------------------
# RQ2 MANN-WHITNEY U TEST — BY SECTOR
# H0: No significant difference in annual non-US import values
# H1: Annual non-US import values are HIGHER during tariff years
# One-sided test: alternative = 'greater'
# -----------------------------------------------------------

print("\n" + "=" * 65)
print("  STEP 9: HYPOTHESIS TESTING — RQ2")
print("  Canada Imports from NON-U.S. Countries (Annual 2012-2025)")
print("  H1: Non-US import values HIGHER during tariff years")
print("=" * 65)

rq2_results = []

for sector in ["Steel", "Aluminum", "Auto"]:
    sec = can_nonus_annual[can_nonus_annual["Sector"] == sector].copy()

    on_vals  = sec[sec["Tariff"] == "Tariff ON"]["Value"]
    off_vals = sec[sec["Tariff"] == "Tariff OFF"]["Value"]

    if len(on_vals) < 2 or len(off_vals) < 2:
        print(f"\n  {sector}: Insufficient data (ON n={len(on_vals)}, OFF n={len(off_vals)})")
        continue

    u_stat, p_value = stats.mannwhitneyu(on_vals, off_vals, alternative="greater")

    reject = p_value < 0.05
    result = "REJECT H0" if reject else "FAIL TO REJECT H0"

    print(f"\n  {sector}:")
    print(f"    Tariff ON  — N={len(on_vals):>3} | Years: {sorted(sec[sec['Tariff']=='Tariff ON']['Year'].tolist())} | Mean=${on_vals.mean():>15,.0f}")
    print(f"    Tariff OFF — N={len(off_vals):>3} | Mean=${off_vals.mean():>15,.0f}")
    print(f"    Mann-Whitney U = {u_stat:,.1f} | p-value = {p_value:.4f}")
    print(f"    Result: {result}")

    rq2_results.append({
        "RQ":              "RQ2",
        "Sector":          sector,
        "Tariff_ON_N":     len(on_vals),
        "Tariff_OFF_N":    len(off_vals),
        "Tariff_ON_Mean":  round(on_vals.mean(), 2),
        "Tariff_OFF_Mean": round(off_vals.mean(), 2),
        "U_Stat":          round(u_stat, 2),
        "P_Value":         round(p_value, 4),
        "Significant":     reject,
        "Result":          result,
    })

print("\n" + "-" * 65)
print("  RQ2 SUMMARY")
print("-" * 65)
print(f"  {'Sector':<12} {'ON Mean':>18} {'OFF Mean':>18} {'p-value':>10} {'Result'}")
print(f"  {'-'*12} {'-'*18} {'-'*18} {'-'*10} {'-'*20}")
for r in rq2_results:
    print(f"  {r['Sector']:<12} ${r['Tariff_ON_Mean']:>16,.0f} ${r['Tariff_OFF_Mean']:>16,.0f} {r['P_Value']:>10.4f} {r['Result']}")

print("\n  RQ2 complete. Results stored in rq2_results.")
print("  Ready for RQ3.")

# =============================================================
# STEP 10 — HYPOTHESIS TESTING: RQ3
# Did Canada's exports TO the U.S. decline during tariff
# periods vs non-tariff periods?
# Dataset: U.S. Imports from Canada (us_imp_monthly)
# =============================================================

# -----------------------------------------------------------
# RQ3 MANN-WHITNEY U TEST — BY SECTOR
# H0: No significant difference in monthly trade values
#     between tariff-on and tariff-off periods
# H1: Monthly trade values are LOWER during tariff-on periods
# One-sided test: alternative = 'less'
# -----------------------------------------------------------

us_imp_monthly["Tariff"] = us_imp_monthly["Period"].apply(flag_tariff)

print("\n" + "=" * 65)
print("  STEP 10: HYPOTHESIS TESTING — RQ3")
print("  Canada Exports TO U.S. (U.S. Imports from Canada)")
print("  H1: Monthly values LOWER during tariff-on periods")
print("=" * 65)

rq3_results = []

for sector in ["Steel", "Aluminum", "Auto"]:
    sec = us_imp_monthly[us_imp_monthly["Sector"] == sector].copy()

    on_vals  = sec[sec["Tariff"] == "Tariff ON"]["General Customs Value"]
    off_vals = sec[sec["Tariff"] == "Tariff OFF"]["General Customs Value"]

    u_stat, p_value = stats.mannwhitneyu(on_vals, off_vals, alternative="less")

    reject = p_value < 0.05
    result = "REJECT H0" if reject else "FAIL TO REJECT H0"

    print(f"\n  {sector}:")
    print(f"    Tariff ON  — N={len(on_vals):>3} | Mean=${on_vals.mean():>15,.0f} | Median=${on_vals.median():>15,.0f}")
    print(f"    Tariff OFF — N={len(off_vals):>3} | Mean=${off_vals.mean():>15,.0f} | Median=${off_vals.median():>15,.0f}")
    print(f"    Mann-Whitney U = {u_stat:,.1f} | p-value = {p_value:.4f}")
    print(f"    Result: {result}")

    rq3_results.append({
        "RQ":               "RQ3",
        "Sector":           sector,
        "Tariff_ON_N":      len(on_vals),
        "Tariff_OFF_N":     len(off_vals),
        "Tariff_ON_Mean":   round(on_vals.mean(), 2),
        "Tariff_OFF_Mean":  round(off_vals.mean(), 2),
        "Tariff_ON_Median": round(on_vals.median(), 2),
        "Tariff_OFF_Median":round(off_vals.median(), 2),
        "U_Stat":           round(u_stat, 2),
        "P_Value":          round(p_value, 4),
        "Significant":      reject,
        "Result":           result,
    })

print("\n" + "-" * 65)
print("  RQ3 SUMMARY")
print("-" * 65)
print(f"  {'Sector':<12} {'ON Mean':>15} {'OFF Mean':>15} {'p-value':>10} {'Result'}")
print(f"  {'-'*12} {'-'*15} {'-'*15} {'-'*10} {'-'*20}")
for r in rq3_results:
    print(f"  {r['Sector']:<12} ${r['Tariff_ON_Mean']:>13,.0f} ${r['Tariff_OFF_Mean']:>13,.0f} {r['P_Value']:>10.4f} {r['Result']}")

print("\n  RQ3 complete. Results stored in rq3_results.")
print("  All hypothesis tests complete. Ready for OLS Regression.")

# =============================================================
# STEP 11 — OLS REGRESSION
# Estimates the association between tariff periods and
# trade values for all three research questions.
#
# Model: Trade Value = Intercept + Beta * Tariff_Dummy + error
#
# Tariff_Dummy = 1 if tariff-on period, 0 if tariff-off
#
# Beta tells you the estimated dollar change in trade value
# during tariff-on periods compared to tariff-off periods.
# A negative Beta supports H1 for RQ1 and RQ3.
# A positive Beta supports H1 for RQ2.
# =============================================================
 
import numpy as np
from scipy import stats
 
def run_ols(y_vals, x_vals, sector, rq_label):
    """
    Run simple OLS regression of trade value on tariff dummy.
    Returns a results dict with key statistics.
    """
    n = len(y_vals)
    if n < 4:
        print(f"  {sector}: Insufficient data for regression (n={n})")
        return None
 
    # Build design matrix [1, tariff_dummy]
    X = np.column_stack([np.ones(n), x_vals])
    y = np.array(y_vals, dtype=float)
 
    # OLS solution
    XtX_inv = np.linalg.pinv(X.T @ X)
    coeffs  = XtX_inv @ X.T @ y
    intercept, beta = coeffs[0], coeffs[1]
 
    # Residuals and standard errors
    y_pred  = X @ coeffs
    resids  = y - y_pred
    ss_res  = np.sum(resids ** 2)
    ss_tot  = np.sum((y - np.mean(y)) ** 2)
    r2      = 1 - ss_res / ss_tot if ss_tot != 0 else 0
    df_res  = n - 2
    mse     = ss_res / df_res if df_res > 0 else 0
    se_beta = np.sqrt(mse * XtX_inv[1, 1]) if XtX_inv[1, 1] >= 0 else 0
 
    # t-statistic and p-value for Beta (two-sided)
    t_stat  = beta / se_beta if se_beta != 0 else 0
    p_value = 2 * stats.t.sf(abs(t_stat), df=df_res)
 
    return {
        "RQ":          rq_label,
        "Sector":      sector,
        "N":           n,
        "Intercept":   round(intercept, 2),
        "Beta":        round(beta, 2),
        "SE_Beta":     round(se_beta, 2),
        "T_Stat":      round(t_stat, 4),
        "P_Value":     round(p_value, 4),
        "R_Squared":   round(r2, 4),
        "Significant": p_value < 0.05,
        "Result":      "Significant" if p_value < 0.05 else "Not Significant",
    }
 
 
def print_ols_result(r):
    """Print a single OLS result in a readable format."""
    direction = "lower" if r["Beta"] < 0 else "higher"
    print(f"\n  {r['Sector']}:")
    print(f"    N = {r['N']} | Intercept = ${r['Intercept']:,.0f} | Beta = ${r['Beta']:,.0f}")
    print(f"    SE = ${r['SE_Beta']:,.0f} | t = {r['T_Stat']:.4f} | p-value = {r['P_Value']:.4f}")
    print(f"    R-squared = {r['R_Squared']:.4f}")
    print(f"    Interpretation: Trade was ${abs(r['Beta']):,.0f} {direction} per period during tariff-on vs tariff-off")
    print(f"    Result: {r['Result']}")
 
 
# =============================================================
# RQ1 — OLS: U.S. Exports to Canada
# =============================================================
 
print("\n" + "=" * 65)
print("  STEP 11: OLS REGRESSION")
print("=" * 65)
 
print("\n  --- RQ1: U.S. Exports to Canada ---")
print("  Beta interpretation: estimated change in monthly FAS Value")
print("  during tariff-on vs tariff-off periods (negative = lower)")
 
ols_rq1 = []
for sector in ["Steel", "Aluminum", "Auto"]:
    sec     = us_exp_monthly[us_exp_monthly["Sector"] == sector].copy()
    y_vals  = sec["FAS Value"].values
    x_vals  = (sec["Tariff"] == "Tariff ON").astype(int).values
    result  = run_ols(y_vals, x_vals, sector, "RQ1")
    if result:
        print_ols_result(result)
        ols_rq1.append(result)
 
# =============================================================
# RQ2 — OLS: Canadian Non-US Annual Imports
# =============================================================
 
print("\n  --- RQ2: Canadian Non-US Imports (Annual 2012-2025) ---")
print("  Beta interpretation: estimated change in annual Non-US import value")
print("  during tariff-on vs tariff-off years (positive = higher)")
 
ols_rq2 = []
for sector in ["Steel", "Aluminum", "Auto"]:
    sec     = can_nonus_annual[can_nonus_annual["Sector"] == sector].copy()
    y_vals  = sec["Value"].values
    x_vals  = (sec["Tariff"] == "Tariff ON").astype(int).values
    result  = run_ols(y_vals, x_vals, sector, "RQ2")
    if result:
        print_ols_result(result)
        ols_rq2.append(result)
 
# =============================================================
# RQ3 — OLS: U.S. Imports from Canada
# =============================================================
 
print("\n  --- RQ3: U.S. Imports from Canada ---")
print("  Beta interpretation: estimated change in monthly Customs Value")
print("  during tariff-on vs tariff-off periods (negative = lower)")
 
ols_rq3 = []
for sector in ["Steel", "Aluminum", "Auto"]:
    sec     = us_imp_monthly[us_imp_monthly["Sector"] == sector].copy()
    y_vals  = sec["General Customs Value"].values
    x_vals  = (sec["Tariff"] == "Tariff ON").astype(int).values
    result  = run_ols(y_vals, x_vals, sector, "RQ3")
    if result:
        print_ols_result(result)
        ols_rq3.append(result)
 
# =============================================================
# COMBINED SUMMARY TABLE
# =============================================================
 
print("\n" + "=" * 65)
print("  OLS REGRESSION SUMMARY — ALL RQs")
print("=" * 65)
print(f"\n  {'RQ':<6} {'Sector':<12} {'Beta':>18} {'p-value':>10} {'R²':>8} {'Result'}")
print(f"  {'-'*6} {'-'*12} {'-'*18} {'-'*10} {'-'*8} {'-'*16}")
 
for r in ols_rq1 + ols_rq2 + ols_rq3:
    print(f"  {r['RQ']:<6} {r['Sector']:<12} ${r['Beta']:>16,.0f} {r['P_Value']:>10.4f} {r['R_Squared']:>8.4f} {r['Result']}")
 
print("\n  OLS Regression complete.")
print("  Results stored in ols_rq1, ols_rq2, ols_rq3.")
print("  Ready for predictive forecast.")
 
# =============================================================
# STEP 12 — PREDICTIVE FORECAST
# Projects Canada's exports to the U.S. forward under two scenarios:
#   Scenario 1: Tariffs remain in place
#   Scenario 2: Tariffs are removed
#
# Dataset: U.S. Imports from Canada (us_imp_monthly) — RQ3 only
#
# Method:
#   - Fit a linear trend to the pre-tariff baseline (Jan 2012 - Feb 2018)
#   - Extrapolate that trend forward 24 months (2026-2027)
#   - Scenario 2 = baseline trend projected forward (no tariff effect)
#   - Scenario 1 = baseline trend + OLS Beta (tariff effect applied)
# =============================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from matplotlib.ticker import FuncFormatter

FORECAST_MONTHS = 24   # 2026 through 2027
BASELINE_START  = "2012-01"
BASELINE_END    = "2018-02"
LAST_ACTUAL     = "2026-01"


def compact_currency(x, pos):
    """Compact currency formatter for chart axes."""
    if abs(x) >= 1_000_000_000:
        return f"${x/1_000_000_000:.1f}B"
    elif abs(x) >= 1_000_000:
        return f"${x/1_000_000:.1f}M"
    elif abs(x) >= 1_000:
        return f"${x/1_000:.0f}K"
    return f"${x:,.0f}"


def build_forecast_rq3(df, value_col, sector, ols_beta, output_folder):
    """
    Build a two-scenario forecast for one sector using RQ3 data.

    Parameters:
        df           : monthly dataframe (us_imp_monthly)
        value_col    : 'General Customs Value'
        sector       : Steel, Aluminum, or Auto
        ols_beta     : Beta from RQ3 OLS regression (negative = tariff decline)
        output_folder: where to save the PNG chart
    """
    sec = df[df["Sector"] == sector].copy()
    sec = sec.sort_values("Period").reset_index(drop=True)

    # ── Fit linear trend to pre-tariff baseline ───────────────
    baseline = sec[
        (sec["Period"] >= BASELINE_START) &
        (sec["Period"] <= BASELINE_END)
    ].copy()

    if len(baseline) < 6:
        print(f"  {sector}: Not enough baseline data.")
        return None

    x_base = np.arange(len(baseline))
    y_base = baseline[value_col].values
    slope, intercept_fit = np.polyfit(x_base, y_base, 1)

    # ── Generate forecast periods ─────────────────────────────
    last_period      = pd.Period(LAST_ACTUAL, freq="M")
    forecast_periods = [str(last_period + i + 1) for i in range(FORECAST_MONTHS)]

    # X index continues from end of full series
    x_start   = len(sec)
    x_forecast = np.arange(x_start, x_start + FORECAST_MONTHS)

    # Scenario 2: baseline trend only (tariffs removed)
    scenario2 = intercept_fit + slope * x_forecast

    # Scenario 1: baseline trend + tariff effect (tariffs remain)
    scenario1 = scenario2 + ols_beta

    # Clip to zero
    scenario1 = np.clip(scenario1, 0, None)
    scenario2 = np.clip(scenario2, 0, None)

    # ── Forecast dataframe ────────────────────────────────────
    forecast_df = pd.DataFrame({
        "Period":                    forecast_periods,
        "Scenario1_Tariffs_Remain":  scenario1,
        "Scenario2_Tariffs_Removed": scenario2,
    })

    # ── Chart ─────────────────────────────────────────────────
    recent_actual  = sec.tail(36).copy()
    actual_periods = recent_actual["Period"].tolist()
    actual_values  = recent_actual[value_col].tolist()
    all_periods    = actual_periods + forecast_periods
    n_actual       = len(actual_periods)

    fig, ax = plt.subplots(figsize=(13, 6))

    # Actual data
    ax.plot(
        range(n_actual),
        actual_values,
        color="#2563EB",
        linewidth=1.8,
        label="Actual Trade Value"
    )

    # Scenario 1 — tariffs remain (red dashed)
    ax.plot(
        range(n_actual - 1, n_actual + FORECAST_MONTHS),
        [actual_values[-1]] + list(scenario1),
        color="#DC2626",
        linewidth=2,
        linestyle="--",
        label="Scenario 1: Tariffs Remain"
    )

    # Scenario 2 — tariffs removed (green dashed)
    ax.plot(
        range(n_actual - 1, n_actual + FORECAST_MONTHS),
        [actual_values[-1]] + list(scenario2),
        color="#16A34A",
        linewidth=2,
        linestyle="--",
        label="Scenario 2: Tariffs Removed"
    )

    # Forecast start line
    ax.axvline(
        x=n_actual - 1,
        color="gray",
        linestyle=":",
        linewidth=1.2
    )

    # Shaded forecast region
    ax.axvspan(n_actual - 1, n_actual + FORECAST_MONTHS - 1, alpha=0.06, color="gray")

    # Label the forecast region
    ax.text(
        n_actual + FORECAST_MONTHS / 2,
        ax.get_ylim()[1] * 0.98 if ax.get_ylim()[1] != 0 else 1,
        "Forecast\n(2026-2027)",
        ha="center", va="top", fontsize=9, color="gray"
    )

    # X axis ticks every 6 months
    tick_positions = list(range(0, len(all_periods), 6))
    tick_labels    = [all_periods[i] for i in tick_positions]
    ax.set_xticks(tick_positions)
    ax.set_xticklabels(tick_labels, rotation=45, fontsize=8)

    ax.yaxis.set_major_formatter(FuncFormatter(compact_currency))
    ax.set_title(
        f"Canada Exports to U.S. — {sector}\nTwo-Scenario Forecast (2026-2027)",
        fontsize=12,
        fontweight="bold",
        pad=12
    )
    ax.set_xlabel("Period", fontsize=10)
    ax.set_ylabel("Trade Value (USD)", fontsize=10)
    ax.legend(loc="upper left", fontsize=9)
    ax.grid(axis="y", linestyle="--", alpha=0.4)

    plt.tight_layout()

    filename  = f"Forecast_Canada_Exports_{sector}.png"
    save_path = os.path.join(output_folder, filename)
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved: {filename}")

    return forecast_df


# =============================================================
# RUN FORECASTS — RQ3 ONLY
# =============================================================

print("\n" + "=" * 65)
print("  STEP 12: PREDICTIVE FORECAST — RQ3")
print("  Canada Exports to U.S. (U.S. Imports from Canada)")
print("  Scenario 1: Tariffs Remain | Scenario 2: Tariffs Removed")
print("  Horizon: 24 months (2026-2027)")
print("=" * 65)

ols_betas_rq3 = {r["Sector"]: r["Beta"] for r in ols_rq3}

forecast_rq3 = {}
for sector in ["Steel", "Aluminum", "Auto"]:
    beta = ols_betas_rq3.get(sector, 0)
    print(f"\n  Building forecast for {sector} (OLS Beta = ${beta:,.0f})")
    forecast_rq3[sector] = build_forecast_rq3(
        df=us_imp_monthly,
        value_col="General Customs Value",
        sector=sector,
        ols_beta=beta,
        output_folder=OUTPUT_FOLDER
    )

# =============================================================
# PRINT FORECAST SUMMARY
# =============================================================

print("\n" + "-" * 65)
print("  FORECAST SUMMARY — Canada Exports to U.S.")
print("-" * 65)
print(f"\n  {'Sector':<12} {'Period':<10} {'Tariffs Remain':>20} {'Tariffs Removed':>20}")
print(f"  {'-'*12} {'-'*10} {'-'*20} {'-'*20}")

for sector, fdf in forecast_rq3.items():
    if fdf is None:
        continue
    first = fdf.iloc[0]
    last  = fdf.iloc[-1]
    print(f"  {sector:<12} {first['Period']:<10} ${first['Scenario1_Tariffs_Remain']:>18,.0f} ${first['Scenario2_Tariffs_Removed']:>18,.0f}")
    print(f"  {sector:<12} {last['Period']:<10}  ${last['Scenario1_Tariffs_Remain']:>18,.0f} ${last['Scenario2_Tariffs_Removed']:>18,.0f}")

print("\n" + "=" * 65)
print("  STEP 12 COMPLETE")
print("  3 forecast charts saved to Outputs folder")
print("  All analysis steps complete. Ready to write findings.")
print("=" * 65)