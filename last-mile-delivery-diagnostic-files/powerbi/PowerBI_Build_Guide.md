# Build This Dashboard Yourself in Power BI Desktop

This walks you through building a real Power BI dashboard from `PowerBI_Import_Ready.csv` — the same 1,000-record delivery dataset behind the Excel/SQL analysis. Doing this yourself (not having it generated for you) is what makes "Power BI" a legitimate, interview-defensible line on your resume.

**Download Power BI Desktop first (free):** https://powerbi.microsoft.com/desktop/

---

## Step 1 — Import the data

1. Open Power BI Desktop → **Get Data** → **Text/CSV**
2. Select `PowerBI_Import_Ready.csv`
3. In the preview window, click **Transform Data** (not Load) — this opens Power Query Editor
4. Check column types: `Ship_Date`, `Promised_Delivery_Date`, `Actual_Delivery_Date` should be **Date**; `Distance_Miles` should be **Whole Number**; everything else **Text**. Fix any that Power BI guessed wrong by clicking the type icon in each column header.
5. Click **Close & Apply**

This step — checking and fixing data types after import — is genuinely what "Power Query" means in a job posting. You just did it.

## Step 2 — Create the core measures (DAX)

In the **Report** view, right-click your table in the Fields pane → **New Measure**. Add each of these:

```dax
First-Attempt Rate =
DIVIDE(
    CALCULATE(COUNTROWS(deliveries), deliveries[Delivery_Status] IN {"Delivered","Delivered - Late"}),
    COUNTROWS(deliveries)
)
```

```dax
OTIF Rate =
DIVIDE(
    CALCULATE(
        COUNTROWS(deliveries),
        deliveries[Delivery_Status] = "Delivered",
        deliveries[Order_Complete] = "Y",
        deliveries[Actual_Delivery_Date] <= deliveries[Promised_Delivery_Date]
    ),
    COUNTROWS(deliveries)
)
```

```dax
Failure Rate =
DIVIDE(
    CALCULATE(COUNTROWS(deliveries), LEFT(deliveries[Delivery_Status], 6) = "Failed"),
    COUNTROWS(deliveries)
)
```

```dax
Avg Lead Time (Days) =
AVERAGEX(
    FILTER(deliveries, deliveries[Delivery_Status] IN {"Delivered","Delivered - Late"}),
    DATEDIFF(deliveries[Ship_Date], deliveries[Actual_Delivery_Date], DAY)
)
```

Format `First-Attempt Rate`, `OTIF Rate`, and `Failure Rate` as **Percentage** (select the measure → Measure Tools tab → Format → Percentage).

## Step 3 — Build the visuals

1. **Card visuals** (×3) for `First-Attempt Rate`, `OTIF Rate`, and `Avg Lead Time (Days)` — drag each measure onto the canvas, Power BI auto-suggests a Card
2. **Clustered bar chart** — Axis: `Region`, Legend: `Carrier`, Values: `Failure Rate`. This reproduces the region × carrier diagnostic.
3. **Donut chart** — Legend: `Failure_Reason_Code`, Values: Count of `Delivery_ID` (filter the visual to `Delivery_Status` starts with "Failed")
4. **Table or matrix** — Rows: `Carrier`, Values: `Avg Lead Time (Days)`, Count of `Delivery_ID`

## Step 4 — Add a slicer

Drag `Region` into a **Slicer** visual on the canvas. This lets anyone viewing the dashboard filter every visual by region — the interactive part that makes it a dashboard rather than a static report.

## Step 5 — Format and title

- Insert a text box: "Last-Mile Delivery Reliability Diagnostic"
- Apply a consistent theme: **View** → **Themes**
- Resize/align visuals into a clean grid

## Step 6 — Save and export for your portfolio

- **File → Save As** → `.pbix` (keep this — it's your real, working file)
- **File → Export → Export to PDF** for a static version you can attach to applications
- If you want it shareable as a link (requires a free Power BI account): **Publish** → Power BI Service → **Share**

---

## What to say about this in an interview

"I built a delivery diagnostic dashboard in Power BI on top of a SQL-analyzed dataset — DAX measures for OTIF, first-attempt rate, and failure rate, with a region-by-carrier breakdown that surfaced a specific hotspot: one carrier had a 45.7% failure rate in one region, mostly driven by carrier delays." That's a true, specific, defensible sentence once you've actually clicked through these steps.

---

*Data source: PowerBI_Import_Ready.csv (1,000 simulated delivery records, generated for this portfolio project). Reflects the same dataset analyzed via SQL in the accompanying Excel workbook.*
