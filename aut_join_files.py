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


# === Logging function ===
def log(msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {msg}\n")
    print(f"[{timestamp}] {msg}")


# === Generate new output filename with incremental number ===
def get_incremental_filename(base_path):
    name, ext = os.path.splitext(base_path)
    counter = 1
    while True:
        new_name = f"{name}_{counter}{ext}"
        if not os.path.exists(new_name):
            return new_name
        counter += 1


# === Find new file excluding specified files ===
def find_new_file(watch_dir, exclude_files):
    files = [f for f in os.listdir(watch_dir) if f not in exclude_files and os.path.isfile(os.path.join(watch_dir, f))]
    if not files:
        return None
    # Sort by modification time descending (newest first)
    files.sort(key=lambda f: os.path.getmtime(os.path.join(watch_dir, f)), reverse=True)
    return files[0]


# === Append new file rows to SAP file and save as new file with incremented name ===
def append_new_to_sap(new_file_path, sap_file_path):
    try:
        log(f"⏳ Loading SAP file")
        sap_df = pd.read_excel(sap_file_path)
    except Exception as e:
        log(f"ERROR loading SAP file: {e}")
        return

    try:
        log(f"⏳ Loading new file")
        new_df = pd.read_excel(new_file_path)
    except Exception as e:
        log(f"ERROR loading new file: {e}")
        return

    combined_df = pd.concat([sap_df, new_df], ignore_index=True)
    
    new_sap_path = get_incremental_filename(sap_file_path)
    
    try:
        combined_df.to_excel(new_sap_path, index=False)
        log(f"Saved combined file as {os.path.basename(new_sap_path)} (added {len(new_df)} rows from {os.path.basename(new_file_path)})")
    except Exception as e:
        log(f"ERROR saving combined file: {e}")


# === Main watcher loop ===
def wait_for_new_file_and_process():
    log("Script started. Waiting for new file...")

    exclude_files = [file_sap, "Table_WD.xlsx", log_file]

    while True:
        new_file = find_new_file(watch_dir, exclude_files)
        if new_file:
            log(f"New file detected: {new_file}")
            new_file_path = os.path.join(watch_dir, new_file)
            sap_file_path = os.path.join(watch_dir, file_sap)
            append_new_to_sap(new_file_path, sap_file_path)
            break
        time.sleep(check_interval)


if __name__ == "__main__":
    wait_for_new_file_and_process()
    