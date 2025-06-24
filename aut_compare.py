import os
import time
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# === CONFIGURATION ===
watch_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "PY - Data")
file_sap = "Table_SAP.xlsx"                              # SAP input file
file_wd = "Table_WD.xlsx"                                # Workday input file
output_file = "AS01,AX04,AS03,AH01 - SAP_vs_WD.xlsx"     # Output Excel file
log_file = "processing_log.txt"                          # Log file
check_interval = 15                                      # Seconds to wait before checking for files again


# Required fields to keep and populate
required_columns = [
    "Pers.No.", "Personnel Number", "EEGrp", "Employee Group", "S", "Employment Status",
    "CoCd", "Company Code", "PA", "Personnel Area", "ESgrp", "Employee Subgroup", "Start Date", "End Date", "Changed by", "Start", "End time", "A/AType", "Attendance or Absence Type"
]
max_valid_date = pd.Timestamp("2262-04-11")


# === Logging utility ===
def log(msg):
    with open(os.path.join(watch_dir, log_file), "a", encoding="utf-8") as logf:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        logf.write(f"[{timestamp}] {msg}\n")
    print(f"[{timestamp}] {msg}")


# === Main comparison logic ===
def process_files():
    try:
        sap_path = os.path.join(watch_dir, file_sap)
        wd_path = os.path.join(watch_dir, file_wd)
        out_path = os.path.join(watch_dir, output_file)


        log("📂 Loading input files...") 
        # Load SAP and WD files
        sap_df = pd.read_excel(sap_path)
        wd_df = pd.read_excel(wd_path)


        # Filter for absence type below:
        # "AS01", "AX04","AS03" - GE
        # "AH01" - LUX
        sap_df = sap_df[sap_df["A/AType"].isin(["AS01", "AX04","AS03", "AH01"])].copy()

        
        all_columns = sap_df.columns.tolist()
        wd_df["Key_WD"] = wd_df["Employee ID"].astype(str) + "_" + wd_df["Time Off date"].dt.strftime("%Y%m%d")

        rows = []

        for _, row in sap_df.iterrows():
            start = row["Start Date"]
            end = row["End Date"]
            if pd.isnull(start) or pd.isnull(end):
                continue


            # Fallback row generator with dashes for optional fields
            def build_row(overrides):
                full_row = {}
                for col in all_columns:
                    if col in required_columns:
                        full_row[col] = row[col]
                    else:
                        full_row[col] = "-"
                full_row.update(overrides)
                return full_row

            if end > max_valid_date:
                original = build_row({
                    "End Date": None,
                    "AbsenceDate_SAP": pd.NaT,
                    "Key_SAP": f"{row['Personnel Number']}_{start.strftime('%Y%m%d')}",
                    "Status": "ORIGINAL"
                })
                rows.append(original)

            elif start == end:
                one_day = build_row({
                    "Start Date": start,
                    "End Date": end,
                    "AbsenceDate_SAP": start,
                    "Key_SAP": f"{row['Personnel Number']}_{start.strftime('%Y%m%d')}",
                    "Status": None
                })
                rows.append(one_day)

            else:
                original = build_row({
                    "AbsenceDate_SAP": pd.NaT,
                    "Key_SAP": f"{row['Personnel Number']}_{start.strftime('%Y%m%d')}",
                    "Status": "ORIGINAL"
                })
                rows.append(original)

                for d in pd.date_range(start, end):
                    split = build_row({
                        "Start Date": d,
                        "End Date": d,
                        "AbsenceDate_SAP": d,
                        "Key_SAP": f"{row['Personnel Number']}_{d.strftime('%Y%m%d')}",
                        "Status": None
                    })
                    rows.append(split)


        # Convert to DataFrame
        df = pd.DataFrame(rows)


        # Mark which SAP rows are found in WD
        df["Status"] = df["Key_SAP"].isin(wd_df["Key_WD"]).map(
            {True: "OK", False: "Missing in WD"}
        ).where(df["Status"] != "ORIGINAL", "ORIGINAL")


        # Remove duplicates v2
        df = df.sort_values(by="Status", ascending=True)  # keep ORIGINAL at the bottom
        df = df.drop_duplicates(subset=["Key_SAP"], keep="first")
        
        # Remove duplicates v1
        #df = df.drop_duplicates(subset=["Key_SAP", "Status"], keep="first")
        #df = df.drop_duplicates(subset=["Key_SAP"], keep="first")


        # Save to Excel
        df.to_excel(out_path, index=False)


        # === Highlight rows with "ORIGINAL" status ===
        wb = load_workbook(out_path)
        ws = wb.active
        fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        status_col = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}.get("Status", None)

        if status_col:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                if row[status_col - 1].value == "ORIGINAL":
                    for cell in row:
                        cell.fill = fill
        wb.save(out_path)
        wb.close()
        log("✅ Processing completed successfully.")
        
    except Exception as e:
        log(f"❌ ERROR during processing: {e}")


# === Wait until required files appear ===
def wait_for_files():
    log("🔍 Watching for input files...")
    while True:
        files = os.listdir(watch_dir)
        if file_sap in files and file_wd in files:
            log("📂 Detected both files. Starting processing.")
            process_files()
            break
        time.sleep(check_interval)


if __name__ == "__main__":
    wait_for_files()
    