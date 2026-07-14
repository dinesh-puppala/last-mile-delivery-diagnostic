import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

random.seed(42)
np.random.seed(42)

N = 1000

regions = ["Northeast", "Southeast", "Midwest", "Southwest", "West", "Mid-Atlantic", "Pacific Northwest", "Mountain"]
carriers = ["Apex Freight", "Northline Logistics", "Coastal Express", "Summit Carriers"]

# Build in problem hotspots so the diagnostic has real signal to find:
# Coastal Express + Southwest region -> elevated address-error rate
# Northline Logistics -> elevated carrier-delay rate generally

rows = []
start_date = datetime(2025, 10, 1)

for i in range(N):
    region = random.choice(regions)
    carrier = random.choice(carriers)
    ship_date = start_date + timedelta(days=random.randint(0, 89))
    promised_lead = random.choice([2, 3, 3, 4, 5])
    promised_date = ship_date + timedelta(days=promised_lead)

    # base failure probabilities
    addr_err_p = 0.05
    carrier_delay_p = 0.05
    unavailable_p = 0.02

    if carrier == "Coastal Express" and region == "Southwest":
        addr_err_p = 0.22
    elif carrier == "Coastal Express":
        addr_err_p = 0.10

    if carrier == "Northline Logistics":
        carrier_delay_p = 0.14

    roll = random.random()
    if roll < addr_err_p:
        status = "Failed - Wrong Address"
        failure_reason = "ADDR_ERR"
        actual_date = None
        complete = None
    elif roll < addr_err_p + carrier_delay_p:
        status = "Failed - Carrier Delay"
        failure_reason = "CARRIER_DELAY"
        actual_date = None
        complete = None
    elif roll < addr_err_p + carrier_delay_p + unavailable_p:
        status = "Failed - Customer Unavailable"
        failure_reason = "CUST_UNAVAIL"
        actual_date = None
        complete = None
    else:
        # delivered -- most on-time, some late
        late_chance = 0.12 if carrier == "Northline Logistics" else 0.06
        if random.random() < late_chance:
            actual_date = promised_date + timedelta(days=random.randint(1, 4))
            status = "Delivered - Late"
        else:
            actual_date = promised_date - timedelta(days=random.randint(0, 1))
            status = "Delivered"
        failure_reason = ""
        complete = "Y" if random.random() > 0.04 else "N"  # 4% partial/incomplete orders

    distance = random.randint(15, 620)

    rows.append({
        "Delivery_ID": 100000 + i,
        "Order_ID": 200000 + i,
        "Region": region,
        "Carrier": carrier,
        "Ship_Date": ship_date,
        "Promised_Delivery_Date": promised_date,
        "Actual_Delivery_Date": actual_date,
        "Delivery_Status": status,
        "Failure_Reason_Code": failure_reason,
        "Order_Complete": complete,
        "Distance_Miles": distance,
    })

df = pd.DataFrame(rows)

# ---- Inject realistic ERP-export messiness ----

# 1. Duplicate ~2% of rows (double-scanned records)
dupe_idx = np.random.choice(df.index, size=int(N * 0.02), replace=False)
df = pd.concat([df, df.loc[dupe_idx]], ignore_index=True)

# 2. Randomize date string formats per row (still valid dates, just inconsistent formatting)
date_formats = ["%m/%d/%Y", "%Y-%m-%d", "%d-%b-%Y"]

def messy_date(d):
    if pd.isna(d):
        return ""
    fmt = random.choice(date_formats)
    return d.strftime(fmt)

df["Ship_Date_RAW"] = df["Ship_Date"].apply(messy_date)
df["Promised_Delivery_Date_RAW"] = df["Promised_Delivery_Date"].apply(messy_date)
df["Actual_Delivery_Date_RAW"] = df["Actual_Delivery_Date"].apply(messy_date)

# 3. Inconsistent casing / whitespace on text fields (~15% of rows)
def messy_text(val):
    if random.random() < 0.15:
        style = random.choice(["upper", "lower", "pad"])
        if style == "upper":
            return val.upper()
        elif style == "lower":
            return val.lower()
        else:
            return f"  {val}  "
    return val

df["Delivery_Status_RAW"] = df["Delivery_Status"].apply(messy_text)
df["Region_RAW"] = df["Region"].apply(messy_text)
df["Carrier_RAW"] = df["Carrier"].apply(messy_text)

# 4. Blank out a few Region/Carrier values (~1%)
blank_idx = np.random.choice(df.index, size=int(len(df) * 0.01), replace=False)
df.loc[blank_idx, "Region_RAW"] = ""

# Build the "raw export" CSV (what an analyst would actually pull from SAP SD)
raw_export = df[[
    "Delivery_ID", "Order_ID", "Region_RAW", "Carrier_RAW",
    "Ship_Date_RAW", "Promised_Delivery_Date_RAW", "Actual_Delivery_Date_RAW",
    "Delivery_Status_RAW", "Failure_Reason_Code", "Order_Complete", "Distance_Miles"
]].rename(columns={
    "Region_RAW": "Region", "Carrier_RAW": "Carrier",
    "Ship_Date_RAW": "Ship_Date", "Promised_Delivery_Date_RAW": "Promised_Delivery_Date",
    "Actual_Delivery_Date_RAW": "Actual_Delivery_Date", "Delivery_Status_RAW": "Delivery_Status"
})

raw_export.to_csv("/home/claude/proj/SAP_SD_Delivery_Export_RAW.csv", index=False)

# Also keep the clean version (ground truth) for building the SQL database
clean = df[[
    "Delivery_ID", "Order_ID", "Region", "Carrier", "Ship_Date",
    "Promised_Delivery_Date", "Actual_Delivery_Date", "Delivery_Status",
    "Failure_Reason_Code", "Order_Complete", "Distance_Miles"
]].drop_duplicates(subset=["Delivery_ID"])  # simulate de-dup during cleaning

clean.to_csv("/home/claude/proj/Delivery_Data_CLEAN.csv", index=False)

print(f"Raw export rows (with dupes/messiness): {len(raw_export)}")
print(f"Clean rows after de-dup: {len(clean)}")
print(clean["Delivery_Status"].value_counts())
