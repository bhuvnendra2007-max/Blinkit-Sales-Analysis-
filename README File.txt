Blinkit Sales Analysis

A reusable Python-based sales analytics project built around the Blinkit
Grocery Data workbook from the referenced GitHub repository.

Project Objective

The objective is to transform raw grocery/outlet data into a clean
analytical dataset, calculate business KPIs, compare sales across
product and outlet dimensions, and generate charts that can be reused in
a dashboard or presentation.

The source repository currently contains the Excel dataset, a
requirement PDF, a dashboard image, and a minimal README. The analysis
package below adds a reproducible Python pipeline, dependency file,
documentation, and a formal project report. citeturn0view0

Files

    Blinkit-Sales-Analysis/
    ├── Blinkit Grocery Data.xlsx
    ├── Blinkit Analysis Requirement.pdf
    ├── Blinkit Analysis Dashboard.png
    ├── blinkit_sales_analysis.py
    ├── requirements.txt
    ├── README.md
    └── project_report.pdf

Dashboard

The project includes the supplied Blinkit Sales Analysis dashboard:

[Blinkit Sales Analysis Dashboard]

The dashboard presents: - Total Sales - Average Sales - Number of
Items - Average Rating - Outlet Establishment trend - Fat Content
split - Sales by Item Type - Outlet Size distribution - Outlet Location
performance - Fat Content by Outlet Location - Outlet Type sales,
average sales and item counts - Interactive-style filter categories for
Outlet Size, Outlet Location and Item Type

The dashboard image is included as a project reference/visualization.
The Python analysis script remains the reproducible analytical layer.

Technology Stack

-   Python
-   Pandas
-   NumPy
-   Matplotlib
-   OpenPyXL
-   Excel/XLSX as source data
-   Git/GitHub for version control

Analysis Workflow

1.  Load the Excel workbook.
2.  Detect the relevant worksheet and sales column.
3.  Standardize common column-name variants.
4.  Clean missing numeric and categorical values.
5.  Normalize inconsistent fat-content labels.
6.  Handle outlet-size gaps using outlet-level information where
    available.
7.  Create derived features such as outlet age and sales-to-MRP ratio.
8.  Calculate KPIs.
9.  Generate dimension-level sales summaries.
10. Export cleaned data, summary tables, KPI CSVs, and PNG charts.

Core KPIs

The script calculates:

-   Total Sales
-   Average Sales per Record
-   Median Sales per Record
-   Number of Records
-   Unique Items
-   Unique Outlets
-   Average Rating, when a rating field exists

Business Dimensions

Sales can be analyzed by:

-   Item Type
-   Fat Content
-   Outlet
-   Outlet Size
-   Outlet Location Type
-   Outlet Type
-   Outlet Establishment Year
-   Item Identifier

How to Run

1. Clone the repository

    git clone https://github.com/gagansingh2007/Blinkit-Sales-Analysis.git
    cd Blinkit-Sales-Analysis

2. Install dependencies

    pip install -r requirements.txt

3. Run the analysis

    python blinkit_sales_analysis.py

Or specify custom paths:

    python blinkit_sales_analysis.py --input "Blinkit Grocery Data.xlsx" --output outputs

Generated Outputs

The script creates an outputs/ folder containing:

-   kpis.csv
-   data_quality.csv
-   cleaned_blinkit_data.csv
-   analysis_summary.txt
-   dimension-level sales summaries
-   top-item summary
-   PNG charts for item type, outlet type, location type, establishment
    year, and fat content

Data Notes

The repository’s Excel file is the primary source for the analysis. The
commonly used Blinkit/Big-Mart-style schema contains product attributes,
pricing, visibility, outlet attributes, and sales. The standard public
version of this dataset has 8,523 rows and 12 variables; the exact
workbook in this repository should be treated as authoritative when
running the script. citeturn5search12turn5search9

Limitations

-   The source data is observational sales data, so associations should
    not automatically be interpreted as causal effects.
-   A rating KPI is only calculated if the workbook contains a rating
    field.
-   Profit, margin, delivery time, customer-level behavior, and
    inventory turnover cannot be derived reliably unless those fields
    exist in the source workbook.
-   The project is an analytics/reporting project rather than a
    production forecasting system.

Suggested Dashboard Sections

1.  Executive KPI cards
2.  Sales by Item Type
3.  Sales by Outlet Type
4.  Sales by Location Type
5.  Sales by Outlet Size
6.  Sales by Fat Content
7.  Sales by Establishment Year
8.  Top Items
9.  Filters for outlet, location, item type, and fat content

Reproducibility

All transformations are implemented in blinkit_sales_analysis.py, so the
analysis can be rerun whenever the source workbook is replaced or
updated.

Source

Original repository:
https://github.com/gagansingh2007/Blinkit-Sales-Analysis
