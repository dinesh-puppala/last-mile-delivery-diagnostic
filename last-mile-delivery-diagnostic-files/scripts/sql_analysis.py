import sqlite3
import pandas as pd

conn = sqlite3.connect(":memory:")
clean = pd.read_csv("/home/claude/proj/Delivery_Data_CLEAN.csv", parse_dates=["Ship_Date", "Promised_Delivery_Date", "Actual_Delivery_Date"])
clean.to_sql("deliveries", conn, index=False, if_exists="replace")

cur = conn.cursor()

queries = {}

# 1. Overall first-attempt delivery rate (delivered without a failure code, regardless of on-time)
queries["first_attempt_rate"] = """
SELECT
    ROUND(100.0 * SUM(CASE WHEN Delivery_Status IN ('Delivered','Delivered - Late') THEN 1 ELSE 0 END) / COUNT(*), 2) AS first_attempt_pct
FROM deliveries;
"""

# 2. OTIF = delivered on or before promised date AND order complete
queries["otif_rate"] = """
SELECT
    ROUND(100.0 * SUM(
        CASE WHEN Delivery_Status = 'Delivered'
             AND Order_Complete = 'Y'
             AND date(Actual_Delivery_Date) <= date(Promised_Delivery_Date)
        THEN 1 ELSE 0 END
    ) / COUNT(*), 2) AS otif_pct
FROM deliveries;
"""

# 3. Failure rate by Region x Carrier (find the worst combinations)
queries["failure_by_region_carrier"] = """
SELECT
    Region,
    Carrier,
    COUNT(*) AS total_deliveries,
    SUM(CASE WHEN Delivery_Status LIKE 'Failed%' THEN 1 ELSE 0 END) AS failures,
    ROUND(100.0 * SUM(CASE WHEN Delivery_Status LIKE 'Failed%' THEN 1 ELSE 0 END) / COUNT(*), 2) AS failure_rate_pct
FROM deliveries
GROUP BY Region, Carrier
HAVING total_deliveries >= 15
ORDER BY failure_rate_pct DESC
LIMIT 8;
"""

# 4. Root cause breakdown of failures
queries["root_cause_breakdown"] = """
SELECT
    Failure_Reason_Code,
    COUNT(*) AS failure_count,
    ROUND(100.0 * COUNT(*) / (SELECT COUNT(*) FROM deliveries WHERE Delivery_Status LIKE 'Failed%'), 2) AS pct_of_failures
FROM deliveries
WHERE Delivery_Status LIKE 'Failed%'
GROUP BY Failure_Reason_Code
ORDER BY failure_count DESC;
"""

# 5. Carrier-level on-time performance (for lead time variance, pull raw values to compute in pandas)
queries["carrier_leadtime_raw"] = """
SELECT
    Carrier,
    julianday(Actual_Delivery_Date) - julianday(Ship_Date) AS lead_time_days
FROM deliveries
WHERE Delivery_Status IN ('Delivered','Delivered - Late');
"""

results = {}
for name, q in queries.items():
    results[name] = pd.read_sql_query(q, conn)

print("=== First-Attempt Delivery Rate ===")
print(results["first_attempt_rate"])

print("\n=== OTIF Rate ===")
print(results["otif_rate"])

print("\n=== Failure Rate by Region x Carrier (Top 8) ===")
print(results["failure_by_region_carrier"])

print("\n=== Root Cause Breakdown ===")
print(results["root_cause_breakdown"])

# Lead time variance by carrier (computed in pandas from SQL-extracted raw values)
lt = results["carrier_leadtime_raw"]
leadtime_summary = lt.groupby("Carrier")["lead_time_days"].agg(["mean", "std", "count"]).round(2).reset_index()
leadtime_summary.columns = ["Carrier", "avg_lead_time_days", "stdev_lead_time_days", "n_deliveries"]
print("\n=== Lead Time Variance by Carrier ===")
print(leadtime_summary)

# Save all results for use in Excel build
results["failure_by_region_carrier"].to_csv("/home/claude/proj/sql_failure_by_region_carrier.csv", index=False)
results["root_cause_breakdown"].to_csv("/home/claude/proj/sql_root_cause.csv", index=False)
leadtime_summary.to_csv("/home/claude/proj/sql_leadtime_by_carrier.csv", index=False)

with open("/home/claude/proj/sql_headline_numbers.txt", "w") as f:
    f.write(f"first_attempt_pct={results['first_attempt_rate'].iloc[0,0]}\n")
    f.write(f"otif_pct={results['otif_rate'].iloc[0,0]}\n")
    f.write(f"total_records={len(clean)}\n")
    worst = results["failure_by_region_carrier"].iloc[0]
    f.write(f"worst_combo_region={worst['Region']}\n")
    f.write(f"worst_combo_carrier={worst['Carrier']}\n")
    f.write(f"worst_combo_failure_rate={worst['failure_rate_pct']}\n")
    top_cause = results["root_cause_breakdown"].iloc[0]
    f.write(f"top_failure_cause={top_cause['Failure_Reason_Code']}\n")
    f.write(f"top_failure_cause_pct={top_cause['pct_of_failures']}\n")

print("\nSaved SQL result CSVs and headline numbers.")
