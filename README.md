# Last-Mile Delivery Reliability Diagnostic

![Excel](https://img.shields.io/badge/Excel-Advanced-217346?logo=microsoftexcel&logoColor=white)
![SQL](https://img.shields.io/badge/SQL-SQLite-4479A1?logo=sqlite&logoColor=white)
![Power BI](https://img.shields.io/badge/Power%20BI-Dashboard-F2C811?logo=powerbi&logoColor=black)
![Python](https://img.shields.io/badge/Python-Data%20Generation-3776AB?logo=python&logoColor=white)
![Status](https://img.shields.io/badge/status-self--directed%20portfolio%20project-blue)

A supply chain analytics project diagnosing **why deliveries fail** and **what it's worth to fix them** — combining a simulated SAP-style ERP export, SQL diagnostics, an Excel financial model, and a Power BI dashboard.

---

## The problem

Wrong-address and delayed deliveries cost carriers an estimated **$15B+ per year** across the US parcel market. Companies know deliveries are failing — but rarely have a structured way to diagnose *where* failures concentrate and *whether* fixing them is worth the investment.

## What this project does

1. **Simulates a realistic SAP SD delivery export** — 1,000 delivery records with the kind of messiness a real ERP export has: inconsistent date formats, inconsistent text casing, duplicate rows, missing fields.
2. **Diagnoses the problem with SQL** — first-attempt delivery rate, OTIF, failure rate by region × carrier, and root cause breakdown.
3. **Builds a financial model in Excel** — a market-sizing cost model plus a 3-scenario ROI analysis (conservative / moderate / optimistic) with sensitivity testing for a proposed delivery-verification fix.
4. **Visualizes it in Power BI** — an interactive dashboard with KPI cards, a region × carrier failure breakdown, and a root-cause donut chart.

## Key findings

| Metric | Result |
|---|---|
| Deliveries analyzed | 1,000 |
| First-attempt delivery rate | 83.8% |
| OTIF (On-Time, In-Full) rate | 74.6% |
| Worst-performing combination | **Coastal Express carrier in the Southwest region — 45.7% failure rate** |
| Leading root cause | Carrier delay (51.2% of all failures), followed by address error (38.9%) |
| Addressable market-level cost (US parcel market) | ~$15.2B/year |
| Modeled ROI (moderate adoption of a fix) | 3.1x benefit-cost ratio, ~$1.5B net recoverable value |

The standout insight: one specific carrier/region combination is nearly **2x** worse than the dataset average — and the dominant cause is carrier delay, not address errors, which reframes where a fix should actually be targeted.

## Dashboard

![Excel Dashboard](images/dashboard_screenshot.jpg)

![SQL Diagnostic Summary](images/sql_diagnostic_screenshot.jpg)

A live, interactive version of this is also built in Power BI — see [`powerbi/PowerBI_Build_Guide.md`](powerbi/PowerBI_Build_Guide.md) for the full build (DAX measures included).

## Repo structure

```
├── data/
│   ├── SAP_SD_Delivery_Export_RAW.csv     # raw, messy simulated ERP export
│   ├── Delivery_Data_CLEAN.csv            # cleaned, de-duplicated dataset
│   └── PowerBI_Import_Ready.csv           # typed and ready for BI import
├── sql/
│   └── diagnostic_queries.sql             # the 5 core diagnostic queries
├── excel/
│   └── LastMile_Misdelivery_Model.xlsx    # full model: cost, scenarios, dashboard, SQL summary
├── scripts/
│   ├── generate_data.py                   # generates the simulated ERP dataset
│   ├── sql_analysis.py                    # runs the SQL diagnostics
│   └── extend_workbook.py                 # builds the Excel workbook from the data
├── powerbi/
│   └── PowerBI_Build_Guide.md             # step-by-step Power BI build with DAX
└── README.md
```

## Tools used

**Excel:** XLOOKUP, SUMIFS, COUNTIFS, NETWORKDAYS, TRIM/CLEAN, IFERROR, conditional formatting, scenario modeling, sensitivity tables
**SQL:** aggregations, CASE logic, GROUP BY / HAVING, subqueries
**Power BI:** DAX measures, clustered bar and donut visuals, slicers
**Python:** pandas/numpy for realistic synthetic data generation, sqlite3 for query execution

## How to reproduce this

```bash
pip install pandas numpy openpyxl
python scripts/generate_data.py       # generates the raw + clean CSVs
python scripts/sql_analysis.py        # runs the SQL diagnostics, saves results
python scripts/extend_workbook.py     # builds the full Excel workbook
```

For the Power BI dashboard, follow [`powerbi/PowerBI_Build_Guide.md`](powerbi/PowerBI_Build_Guide.md) — it uses `data/PowerBI_Import_Ready.csv` as the source.

## Note on the data

This project uses a **simulated dataset**, generated to reflect realistic patterns in last-mile delivery operations (including intentionally messy formatting, to mirror a real ERP export). It is not real company data. The cost and ROI figures combine this simulated diagnostic with published industry benchmarks (parcel volume, failure rates, cost per failure) — see the Assumptions tab in the Excel workbook for sources.

---

*Part of a 3-project supply chain analytics portfolio. Built as preparation for Supply Chain / Operations Analyst roles.*
