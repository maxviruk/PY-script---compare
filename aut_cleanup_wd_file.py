import os
import pandas as pd
from datetime import datetime


# === CONFIGURATION ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
input_folder = os.path.join(BASE_DIR, "PY - Data - WD original")
output_folder = os.path.join(BASE_DIR, "PY - Data - EOPWD")
log_dir = os.path.join(BASE_DIR, "PY - Logs")
log_file = os.path.join(log_dir, "processing_log_2.txt")
mapping_filename = "WD - ColumnMapping.xlsx"


# === Ensure folders exist ===
os.makedirs(output_folder, exist_ok=True)
os.makedirs(log_dir, exist_ok=True)


# === Logging function ===
def log(msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {msg}\n")
    print(f"[{timestamp}] {msg}")

try:
    
    # === Find the input Excel file (excluding the mapping file) ===
    input_files = [f for f in os.listdir(input_folder)
                   if f.endswith(".xlsx") and f != mapping_filename]
    
    if not input_files:
        raise FileNotFoundError("❌ No Excel files found in the input folder.")

    input_file_path = os.path.join(input_folder, input_files[0])
    mapping_file_path = os.path.join(input_folder, mapping_filename)

    log(f"📂 Found input file")


    # === Load column mapping ===
    mapping_df = pd.read_excel(mapping_file_path, usecols=[0, 1], header=None, names=["column", "action"])
    columns_to_delete = mapping_df[mapping_df["action"].str.lower() == "delete"]["column"].tolist()
    log(f"🔧 Columns marked for deletion")


    # === Load input file and skip first 13 rows ===
    df = pd.read_excel(input_file_path, skiprows=13)
    log(f"📊 File loaded. Columns before cleanup")


    # === Delete columns based on mapping ===
    not_found = [col for col in columns_to_delete if col not in df.columns]
    found_to_delete = [col for col in columns_to_delete if col in df.columns]

    if not_found:
        log(f"⚠️ Some columns in the mapping were not found in the file: {not_found}")
    df = df.drop(columns=found_to_delete)
    log(f"🧹 Deleted columns")


    # === Filter: keep only rows with Employment Status ID == 3 ===
    before_count = len(df)
    df = df[df["Employment Status ID"] == 3]
    log(f"🧹 Filtered Employment Status ID != 3 — removed {before_count - len(df)} rows")


    # === Filter: remove rows with empty Time Off type ===
    before_count = len(df)
    df = df[df["Time Off type"].notna() & (df["Time Off type"].astype(str).str.strip() != "")]
    log(f"🧹 Removed rows with empty Time Off type — removed {before_count - len(df)} rows")



    # === Filter: Keep only Time Off date dates <= today + 3 months ===
    today = pd.Timestamp.today().normalize()
    cutoff_date = today + pd.DateOffset(months=3)

    # Convert with specific format: day/month/year
    df["Time Off date"] = pd.to_datetime(df["Time Off date"], format="%d/%m/%Y", errors='coerce')

    invalid_dates = df["Time Off date"].isna().sum()
    log(f"🧪 Failed to parse 'Time Off date' in {invalid_dates} rows")

    # Apply filter
    before_count = len(df)
    df = df[df["Time Off date"] <= cutoff_date]
    log(f"🧹 Removed rows where Time Off date > {cutoff_date.date()} — removed {before_count - len(df)} rows")


    # === Filter: Keep only selected countries ===
    allowed_countries = ["Netherlands", "Germany", "Luxembourg"]
    before_count = len(df)
    df = df[df["Work Location Country"].isin(allowed_countries)]
    log(f"🧹 Removed rows not in allowed countries — removed {before_count - len(df)} rows")


    # === Final save ===
    date_suffix = datetime.now().strftime("%d%m")
    output_filename = f"Table_WD_{date_suffix}.xlsx"
    output_path = os.path.join(output_folder, output_filename)

    df.to_excel(output_path, index=False) 
    log(f"💾 File successfully saved: {output_filename}")

except Exception as e:
    log(f"❌ Error: {str(e)}")
