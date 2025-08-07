import os
import time
import pandas as pd
from datetime import datetime

# === CONFIGURATION ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(BASE_DIR, "PY - Data - WD original")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "PY - Data - EOPWD")
LOG_DIR = os.path.join(BASE_DIR, "PY - Logs")
LOG_FILE = os.path.join(LOG_DIR, "processing_log_2.txt")
MAPPING_FILENAME = "WD - ColumnMapping.xlsx"
ALLOWED_COUNTRIES = {"Netherlands", "Germany", "Luxembourg"}
FILE_NAME_PART = "Absence - EUR - Time Offs Report"
CHECK_INTERVAL = 10  # seconds


# === Ensure folders exist ===
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)


def log(msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{timestamp}] {msg}"
    print(entry)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(entry + "\n")


# === Find latest matching input file ===
def find_latest_matching_file():
    candidates = []
    for f in os.listdir(INPUT_FOLDER):
        if f.endswith(".xlsx") and MAPPING_FILENAME not in f and FILE_NAME_PART in f:
            full_path = os.path.join(INPUT_FOLDER, f)
            mod_time = os.path.getmtime(full_path)
            candidates.append((mod_time, full_path))
    return max(candidates, key=lambda x: x[0])[1] if candidates else None


# === Load column mapping for deletion ===
def load_column_mapping():
    mapping_path = os.path.join(INPUT_FOLDER, MAPPING_FILENAME)
    mapping_df = pd.read_excel(mapping_path, usecols=[0, 1], header=None, names=["column", "action"])
    return mapping_df[mapping_df["action"].str.lower() == "delete"]["column"].tolist()


# === Main cleanup function ===
def process_file(input_path):
    try:
        columns_to_delete = load_column_mapping()
        log(f"📂 Loaded column mapping")

        df = pd.read_excel(input_path, skiprows=13)
        log(f"📊 Loaded file '{os.path.basename(input_path)}' with {len(df)} rows")

        # Drop columns
        actual_cols = set(df.columns)
        drop_cols = [col for col in columns_to_delete if col in actual_cols]
        missing_cols = [col for col in columns_to_delete if col not in actual_cols]
        df.drop(columns=drop_cols, inplace=True)
        if missing_cols:
            log(f"⚠️ Columns to delete not found: {missing_cols}")
        log(f"🧹 Dropped {len(drop_cols)} columns")

        # Filter 1
        if "Employment Status ID" in df.columns:
            before = len(df)
            df = df[df["Employment Status ID"] == 3]
            log(f"🧹 Employment Status ID != 3 — removed {before - len(df)} rows")

        # Filter 2
        if "Time Off type" in df.columns:
            before = len(df)
            df = df[df["Time Off type"].notna() & (df["Time Off type"].astype(str).str.strip() != "")]
            log(f"🧹 Empty Time Off type — removed {before - len(df)} rows")

        # Filter 3
        if "Time Off date" in df.columns:
            df["Time Off date"] = pd.to_datetime(df["Time Off date"], format="%d/%m/%Y", errors="coerce")
            failed = df["Time Off date"].isna().sum()
            log(f"🧪 Failed to parse 'Time Off date' in {failed} rows")

            before = len(df)
            cutoff = pd.Timestamp.today().normalize() + pd.DateOffset(months=3)
            df = df[df["Time Off date"] <= cutoff]
            log(f"🧹 Time Off date > {cutoff.date()} — removed {before - len(df)} rows")

        # Filter 4
        if "Work Location Country" in df.columns:
            before = len(df)
            df = df[df["Work Location Country"].isin(ALLOWED_COUNTRIES)]
            log(f"🧹 Not in allowed countries — removed {before - len(df)} rows")

        # Save output
        date_suffix = datetime.now().strftime("%d%m")
        output_file = f"Table_WD_{date_suffix}.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_file)
        df.to_excel(output_path, index=False)
        log(f"💾 File saved: {output_file}")

    except Exception as e:
        log(f"❌ Error during processing: {e}")


# === Wait and run ===
def wait_for_file():
    log("🚀 Script started. Waiting for input file")
    while True:
        latest_file = find_latest_matching_file()
        if latest_file:
            log(f"📥 Detected file: {os.path.basename(latest_file)}")
            process_file(latest_file)
            break
        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    wait_for_file()
