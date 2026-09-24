"""
Blinkit Sales Analysis
-----------------------
Reusable Python analysis pipeline for:
    Blinkit Grocery Data.xlsx

The script:
- Loads the Excel workbook
- Detects and standardizes common Blinkit/Big-Mart style column names
- Cleans missing and inconsistent categorical values
- Calculates core KPIs
- Produces category/outlet/location summaries
- Saves cleaned data, KPI tables, and PNG charts under ./outputs/

Run:
    python blinkit_sales_analysis.py

Optional:
    python blinkit_sales_analysis.py --input "Blinkit Grocery Data.xlsx" --output outputs
"""

from __future__ import annotations

import argparse
from pathlib import Path
import re

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


ALIASES = {
    "item_identifier": ["Item_Identifier", "Item Identifier", "Item ID"],
    "item_weight": ["Item_Weight", "Item Weight"],
    "item_fat_content": ["Item_Fat_Content", "Item Fat Content", "Fat Content"],
    "item_visibility": ["Item_Visibility", "Item Visibility"],
    "item_type": ["Item_Type", "Item Type", "Category"],
    "item_mrp": ["Item_MRP", "Item MRP", "MRP"],
    "outlet_identifier": ["Outlet_Identifier", "Outlet Identifier", "Outlet ID"],
    "outlet_establishment_year": [
        "Outlet_Establishment_Year", "Outlet Establishment Year", "Establishment Year"
    ],
    "outlet_size": ["Outlet_Size", "Outlet Size"],
    "outlet_location_type": [
        "Outlet_Location_Type", "Outlet Location Type", "Location Type"
    ],
    "outlet_type": ["Outlet_Type", "Outlet Type"],
    "sales": [
        "Item_Outlet_Sales", "Item Outlet Sales", "Total Sales", "Sales",
        "Total_Sales", "Item Sales"
    ],
    "rating": ["Rating", "Average Rating", "Customer Rating"],
    "number_of_items": ["Number of Items", "Number_of_Items", "Item Count", "Item_Count"],
}


def normalize_name(value: str) -> str:
    value = str(value).strip().lower()
    value = re.sub(r"[^a-z0-9]+", "_", value)
    return value.strip("_")


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized = {normalize_name(c): c for c in df.columns}
    for candidate in candidates:
        key = normalize_name(candidate)
        if key in normalized:
            return normalized[key]
    return None


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    for target, candidates in ALIASES.items():
        col = find_column(df, candidates)
        if col:
            rename_map[col] = target
    return df.rename(columns=rename_map)


def load_excel(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    # Prefer a sheet containing sales-related fields; otherwise use the first sheet.
    selected = xls.sheet_names[0]
    for sheet in xls.sheet_names:
        preview = pd.read_excel(path, sheet_name=sheet, nrows=5)
        if find_column(preview, ALIASES["sales"]):
            selected = sheet
            break
    return pd.read_excel(path, sheet_name=selected)


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    df = standardize_columns(df.copy())

    # Convert numeric fields where present.
    numeric_cols = [
        "item_weight", "item_visibility", "item_mrp",
        "outlet_establishment_year", "sales", "rating", "number_of_items"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Standardize fat-content labels.
    if "item_fat_content" in df.columns:
        df["item_fat_content"] = (
            df["item_fat_content"]
            .astype("string")
            .str.strip()
            .str.lower()
            .replace({
                "lf": "Low Fat",
                "low fat": "Low Fat",
                "reg": "Regular",
                "regular": "Regular",
            })
            .fillna("Unknown")
        )

    # Fill item weight with median within item type, then global median.
    if "item_weight" in df.columns:
        if "item_type" in df.columns:
            df["item_weight"] = df.groupby("item_type")["item_weight"].transform(
                lambda s: s.fillna(s.median())
            )
        df["item_weight"] = df["item_weight"].fillna(df["item_weight"].median())

    # Outlet size can be inferred from the same outlet identifier where possible.
    if "outlet_size" in df.columns and "outlet_identifier" in df.columns:
        df["outlet_size"] = df.groupby("outlet_identifier")["outlet_size"].transform(
            lambda s: s.ffill().bfill()
        )
    if "outlet_size" in df.columns:
        df["outlet_size"] = df["outlet_size"].fillna("Unknown")

    # Visibility of exactly zero is often a data-quality signal in this dataset.
    if "item_visibility" in df.columns:
        positive = df.loc[df["item_visibility"] > 0, "item_visibility"]
        replacement = positive.median() if not positive.empty else 0
        df["item_visibility"] = df["item_visibility"].replace(0, replacement)

    # Derived fields.
    if "outlet_establishment_year" in df.columns:
        df["outlet_age"] = (
            df["outlet_establishment_year"].max() -
            df["outlet_establishment_year"]
        )

    if "sales" in df.columns and "item_mrp" in df.columns:
        df["sales_to_mrp_ratio"] = np.where(
            df["item_mrp"] > 0, df["sales"] / df["item_mrp"], np.nan
        )

    return df


def safe_group_sum(df, by, value="sales"):
    if value not in df.columns or by not in df.columns:
        return pd.DataFrame()
    return (
        df.groupby(by, dropna=False)[value]
        .agg(["sum", "mean", "count"])
        .sort_values("sum", ascending=False)
        .reset_index()
        .rename(columns={"sum": "total_sales", "mean": "avg_sales", "count": "records"})
    )


def save_bar(data, x, y, title, filename, horizontal=False, top_n=15):
    if data.empty:
        return
    data = data.head(top_n).copy()
    plt.figure(figsize=(10, 6))
    if horizontal:
        plt.barh(data[x].astype(str), data[y])
        plt.gca().invert_yaxis()
        plt.xlabel(y.replace("_", " ").title())
        plt.ylabel(x.replace("_", " ").title())
    else:
        plt.bar(data[x].astype(str), data[y])
        plt.xticks(rotation=45, ha="right")
        plt.ylabel(y.replace("_", " ").title())
        plt.xlabel(x.replace("_", " ").title())
    plt.title(title)
    plt.tight_layout()
    plt.savefig(filename, dpi=180)
    plt.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default="Blinkit Grocery Data.xlsx")
    parser.add_argument("--output", default="outputs")
    args = parser.parse_args()

    input_path = Path(args.input)
    output = Path(args.output)
    output.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise FileNotFoundError(
            f"Input file not found: {input_path}. "
            "Place 'Blinkit Grocery Data.xlsx' beside this script or pass --input."
        )

    raw = load_excel(input_path)
    df = clean_data(raw)

    if "sales" not in df.columns:
        raise ValueError(
            "A sales column could not be detected. Expected one of: "
            + ", ".join(ALIASES["sales"])
        )

    # Core KPIs.
    kpis = {
        "total_sales": df["sales"].sum(),
        "average_sales_per_record": df["sales"].mean(),
        "median_sales_per_record": df["sales"].median(),
        "record_count": len(df),
        "unique_items": df["item_identifier"].nunique() if "item_identifier" in df else np.nan,
        "unique_outlets": df["outlet_identifier"].nunique() if "outlet_identifier" in df else np.nan,
        "average_rating": df["rating"].mean() if "rating" in df else np.nan,
    }
    pd.DataFrame([kpis]).to_csv(output / "kpis.csv", index=False)

    # Data-quality profile.
    quality = pd.DataFrame({
        "column": df.columns,
        "dtype": [str(df[c].dtype) for c in df.columns],
        "missing_values": [df[c].isna().sum() for c in df.columns],
        "unique_values": [df[c].nunique(dropna=True) for c in df.columns],
    })
    quality.to_csv(output / "data_quality.csv", index=False)

    # Summary tables.
    dimensions = [
        "item_type", "item_fat_content", "outlet_identifier",
        "outlet_size", "outlet_location_type", "outlet_type",
        "outlet_establishment_year"
    ]
    for dim in dimensions:
        summary = safe_group_sum(df, dim)
        if not summary.empty:
            summary.to_csv(output / f"{dim}_sales_summary.csv", index=False)

    # Top products if item identifier exists.
    if "item_identifier" in df.columns:
        top_items = safe_group_sum(df, "item_identifier").head(20)
        top_items.to_csv(output / "top_items.csv", index=False)

    # Charts.
    if "item_type" in df.columns:
        save_bar(
            safe_group_sum(df, "item_type"),
            "item_type", "total_sales",
            "Sales by Item Type", output / "sales_by_item_type.png",
            horizontal=True
        )

    if "outlet_type" in df.columns:
        save_bar(
            safe_group_sum(df, "outlet_type"),
            "outlet_type", "total_sales",
            "Sales by Outlet Type", output / "sales_by_outlet_type.png",
            horizontal=True
        )

    if "outlet_location_type" in df.columns:
        save_bar(
            safe_group_sum(df, "outlet_location_type"),
            "outlet_location_type", "total_sales",
            "Sales by Location Type", output / "sales_by_location_type.png"
        )

    if "outlet_establishment_year" in df.columns:
        year = safe_group_sum(df, "outlet_establishment_year")
        if not year.empty:
            plt.figure(figsize=(9, 5))
            plt.plot(year["outlet_establishment_year"], year["total_sales"], marker="o")
            plt.title("Sales by Outlet Establishment Year")
            plt.xlabel("Establishment Year")
            plt.ylabel("Total Sales")
            plt.grid(alpha=0.25)
            plt.tight_layout()
            plt.savefig(output / "sales_by_establishment_year.png", dpi=180)
            plt.close()

    if "item_fat_content" in df.columns:
        save_bar(
            safe_group_sum(df, "item_fat_content"),
            "item_fat_content", "total_sales",
            "Sales by Fat Content", output / "sales_by_fat_content.png"
        )

    # Save cleaned data.
    df.to_csv(output / "cleaned_blinkit_data.csv", index=False)

    # Human-readable summary.
    with open(output / "analysis_summary.txt", "w", encoding="utf-8") as f:
        f.write("BLINKIT SALES ANALYSIS SUMMARY\n")
        f.write("=" * 40 + "\n\n")
        for key, value in kpis.items():
            f.write(f"{key}: {value}\n")

    print("Analysis complete.")
    print(f"Rows analyzed: {len(df):,}")
    print(f"Total sales: {df['sales'].sum():,.2f}")
    print(f"Output folder: {output.resolve()}")


if __name__ == "__main__":
    main()
