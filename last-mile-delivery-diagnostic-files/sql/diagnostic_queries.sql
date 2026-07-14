-- =====================================================================
-- Last-Mile Delivery Reliability Diagnostic — SQL Queries
-- Run against the cleaned `deliveries` table (see data/Delivery_Data_CLEAN.csv)
-- Dialect: SQLite (portable to Postgres/MySQL with minor date-function changes)
-- =====================================================================

-- 1. Overall first-attempt delivery rate
-- (delivered without a failure code, regardless of on-time)
SELECT
    ROUND(100.0 * SUM(CASE WHEN Delivery_Status IN ('Delivered','Delivered - Late') THEN 1 ELSE 0 END) / COUNT(*), 2) AS first_attempt_pct
FROM deliveries;


-- 2. OTIF (On-Time, In-Full) rate
-- Delivered on/before promised date AND order marked complete
SELECT
    ROUND(100.0 * SUM(
        CASE WHEN Delivery_Status = 'Delivered'
             AND Order_Complete = 'Y'
             AND date(Actual_Delivery_Date) <= date(Promised_Delivery_Date)
        THEN 1 ELSE 0 END
    ) / COUNT(*), 2) AS otif_pct
FROM deliveries;


-- 3. Failure rate by Region x Carrier — surfaces the worst-performing combinations
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


-- 4. Root cause breakdown of all failures
SELECT
    Failure_Reason_Code,
    COUNT(*) AS failure_count,
    ROUND(100.0 * COUNT(*) / (SELECT COUNT(*) FROM deliveries WHERE Delivery_Status LIKE 'Failed%'), 2) AS pct_of_failures
FROM deliveries
WHERE Delivery_Status LIKE 'Failed%'
GROUP BY Failure_Reason_Code
ORDER BY failure_count DESC;


-- 5. Lead time (ship-to-delivery, in days) per delivered record — used to compute
--    avg/stdev lead time variance by carrier in the analysis script
SELECT
    Carrier,
    julianday(Actual_Delivery_Date) - julianday(Ship_Date) AS lead_time_days
FROM deliveries
WHERE Delivery_Status IN ('Delivered','Delivered - Late');
