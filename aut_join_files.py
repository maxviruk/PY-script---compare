"""
aut_join_files.py

This script waits for a new Excel file to appear in the folder, then merges it with the main SAP file.
Main features:
1. Watches for a new file (excluding specified files).
2. Loads SAP and new Excel files.
3. Filters by allowed CoCd values.
4. Removes duplicates using Pers.No. + Start Date + A/AType as key.
5. Saves the combined file with an incremental filename.
All steps are logged to file and console.

"""

import os
import time
from datetime import datetime
import pandas as pd

# === CONFIGURATION ===
watch_dir = os.path.join(os.getcwd(), "PY - Data - EOPWD")
log_dir = os.path.join(os.getcwd(), "PY - Logs")
os.makedirs(log_dir, exist_ok=True)
file_sap = "Table_SAP.xlsx"
log_file = "processing_log_3.txt"
log_path = os.path.join(log_dir, log_file)
check_interval = 10  # seconds


def log(msg):
    """
    Logs a message with a timestamp to both the log file and the console.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {msg}\n")
    print(f"[{timestamp}] {msg}")


def get_incremental_filename(base_path):
    """
    Generates a new filename with an incremental number suffix (file_1.xlsx, file_2.xlsx, etc.).
    """
    name, ext = os.path.splitext(base_path)
    counter = 1
    while True:
        new_name = f"{name}_{counter}{ext}"
        if not os.path.exists(new_name):
            return new_name
        counter += 1


def find_new_file(watch_dir, exclude_files):
    """
    Finds the most recently modified file in watch_dir, excluding those in exclude_files.
    Returns the filename or None if no new files are found.
    """
    files = [f for f in os.listdir(watch_dir) if f not in exclude_files and os.path.isfile(os.path.join(watch_dir, f))]
    if not files:
        return None
    files.sort(key=lambda f: os.path.getmtime(os.path.join(watch_dir, f)), reverse=True)
    return files[0]


def load_excel_files(new_file_path, sap_file_path):
    """
    Loads the SAP file and new file into pandas DataFrames.
    Returns: sap_df, new_df, and sap_columns (list of column names).
    """
    try:
        log("⏳ Loading SAP file")
        sap_df = pd.read_excel(sap_file_path)
        sap_columns = sap_df.columns.tolist()
    except Exception as e:
        log(f"❌ ERROR loading SAP file: {e}")
        return None, None, None
    try:
        log("⏳ Loading new file")
        new_df = pd.read_excel(new_file_path)
    except Exception as e:
        log(f"❌ ERROR loading new file: {e}")
        return None, None, None
    return sap_df, new_df, sap_columns


def filter_cocd(combined_df):
    """
    Filters the DataFrame to only include rows where CoCd is in the allowed values.
    Logs how many rows are removed.
    """
    valid_cocds = [
        "DE11", "DE14", "DE15", "DE19", "DE20", "DE43", "DE78", "DE84", "DE85", "DE86", "DE91", "DE92", "DE93", "DE94",
        "HQ01", "HQ02", "HQ06", "HQ76", "HQ78", "HQ79", "HQ80", "HQ81", "HQ82", "HQ83", "HQ86", "HQ87", "HQ93", "HQ95", "HQ96",
        "LU01", "NL11", "NL84"
    ]
    if "CoCd" in combined_df.columns:
        before = len(combined_df)
        combined_df = combined_df[combined_df["CoCd"].isin(valid_cocds)]
        removed = before - len(combined_df)
        log(f"🧹 Filtered CoCd: remaining {len(combined_df)}, removed {removed} rows")
    else:
        log("❗ Column 'CoCd' not found in combined_df — no filtering applied")
    return combined_df


def remove_duplicates(combined_df):
    """
    Removes duplicate rows based on the composite key: Pers.No. + Start Date + A/AType.
    Logs how many duplicates are removed.
    """
    required_cols = ["Pers.No.", "Start Date", "A/AType"]
    if all(col in combined_df.columns for col in required_cols):
        combined_df["#"] = combined_df["Pers.No."].astype(str) + \
                           combined_df["Start Date"].astype(str) + \
                           combined_df["A/AType"].astype(str)
        before_dedup = len(combined_df)
        combined_df.drop_duplicates(subset=["#"], inplace=True)
        after_dedup = len(combined_df)
        combined_df.drop(columns=["#"], inplace=True)
        log(f"🧹 Pers.No. + Start Date + A/AType: Removed duplicates - {before_dedup - after_dedup}")
    else:
        log("⚠️ One or more required columns for deduplication ('Pers.No.', 'Start Date', 'A/AType') are missing")
    return combined_df


def save_combined_file(combined_df, new_file_path, sap_file_path):
    """
    Saves the combined DataFrame as a new Excel file with an incremental filename.
    Logs the save status.
    """
    new_sap_path = get_incremental_filename(sap_file_path)
    try:
        combined_df.to_excel(new_sap_path, index=False)
        log(f"✅ Saved combined file as {os.path.basename(new_sap_path)} (total {len(combined_df)} rows)")
    except Exception as e:
        log(f"❌ ERROR saving combined file: {e}")


def append_new_to_sap(new_file_path, sap_file_path):
    """
    Main workflow:
    1. Loads SAP and new file.
    2. Aligns columns (including "PY" if present).
    3. Concatenates data, filters by CoCd, removes duplicates, and saves result.
    """
    sap_df, new_df, sap_columns = load_excel_files(new_file_path, sap_file_path)
    if sap_df is None or new_df is None:
        return

    # Handle "PY" column
    keep_columns = sap_columns + (["PY"] if "PY" in new_df.columns else [])
    new_df = new_df[[col for col in keep_columns if col in new_df.columns]]

    if "PY" in new_df.columns and "PY" not in sap_df.columns:
        sap_df["PY"] = "-"

    final_columns = sap_columns + (["PY"] if "PY" in new_df.columns else [])
    sap_df = sap_df[final_columns] if "PY" in sap_df.columns else sap_df[sap_columns]
    new_df = new_df[final_columns] if "PY" in new_df.columns else new_df[sap_columns]

    combined_df = pd.concat([sap_df, new_df], ignore_index=True)
    log(f"🧾 Combined rows before CoCd filter: {len(combined_df)}")

    combined_df = filter_cocd(combined_df)
    combined_df = remove_duplicates(combined_df)
    save_combined_file(combined_df, new_file_path, sap_file_path)


def wait_for_new_file_and_process():
    """
    Main watcher loop. Waits for a new file, then starts the append/merge process.
    """
    log("🚀 Script started. Waiting for new file")
    exclude_files = [file_sap, "Table_WD.xlsx", log_file]
    while True:
        new_file = find_new_file(watch_dir, exclude_files)
        if new_file:
            log(f"📂 New file detected: {new_file}")
            new_file_path = os.path.join(watch_dir, new_file)
            sap_file_path = os.path.join(watch_dir, file_sap)
            append_new_to_sap(new_file_path, sap_file_path)
            break
        time.sleep(check_interval)

if __name__ == "__main__":
    wait_for_new_file_and_process()
