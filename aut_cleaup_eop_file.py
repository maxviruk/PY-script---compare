import os
import time
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# === CONFIGURATION ===
watch_dir = os.path.join(os.getcwd(), "PY - Data - EOPWD")
log_dir = os.path.join(os.getcwd(), "PY - Logs")
max_valid_date = pd.Timestamp("2262-04-11")
file_sap = "Table_SAP.xlsx"
file_wd = "Table_WD.xlsx"
output_file = "SAP_Expanded.xlsx"  # kept as-is to avoid breaking downstream
log_file = "processing_log_1.txt"
check_interval = 10

required_columns = [
    "Pers.No.", "Personnel Number", "EEGrp", "Employee Group", "S", "Employment Status",
    "CoCd", "Company Code", "PA", "Personnel Area", "ESgrp", "Employee Subgroup", "Start Date",
    "End Date", "Changed by", "Start", "End time", "A/AType", "Attendance or Absence Type"
]

def log(msg):
    with open(os.path.join(log_dir, log_file), "a", encoding="utf-8") as logf:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        logf.write(f"[{timestamp}] {msg}\n")
    print(f"[{timestamp}] {msg}")

def get_unique_output_path(watch_dir, base_filename):
    base_path = os.path.join(watch_dir, base_filename)
    if not os.path.exists(base_path):
        return base_path

    name, ext = os.path.splitext(base_filename)
    date_suffix = datetime.now().strftime("%d%m%Y")
    counter = 0

    while True:
        candidate_name = f"{name}_{date_suffix}-{counter}{ext}" if counter > 0 else f"{name}_{date_suffix}{ext}"
        candidate_path = os.path.join(watch_dir, candidate_name)
        if not os.path.exists(candidate_path):
            return candidate_path
        counter += 1

def add_formula_columns(file_path):
    wb = load_workbook(file_path)
    ws = wb.active
    max_row = ws.max_row
    last_col_idx = ws.max_column

    headers = [cell.value for cell in ws[1]]
    if "PY" not in headers:
        col_idx = last_col_idx + 1
        ws.cell(row=1, column=col_idx).value = "PY"
        col_letter = get_column_letter(col_idx)
        for row in range(2, max_row + 1):
            ws[f"{col_letter}{row}"] = "Compared"

    wb.save(file_path)
    wb.close()

def reorder_columns(file_path):
    wb = load_workbook(file_path)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    cols_to_move = ["AbsenceDate_SAP", "Key_SAP", "PY"]
    indices_to_move = [headers.index(col) + 1 for col in cols_to_move if col in headers]
    max_col = ws.max_column
    max_row = ws.max_row
    new_col_start = max_col + 1

    for i, col_idx in enumerate(indices_to_move):
        new_col = new_col_start + i
        ws.cell(row=1, column=new_col).value = ws.cell(row=1, column=col_idx).value
        for row in range(2, max_row + 1):
            ws.cell(row=row, column=new_col).value = ws.cell(row=row, column=col_idx).value

    for col_idx in sorted(indices_to_move, reverse=True):
        ws.delete_cols(col_idx)

    wb.save(file_path)
    wb.close()

def add_formula_and_remove_duplicates(file_path, column_title="#"):
    wb = load_workbook(file_path)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    if column_title not in headers:
        col_idx = ws.max_column + 1
        ws.cell(row=1, column=col_idx).value = column_title
    else:
        col_idx = headers.index(column_title) + 1

    col_letter = get_column_letter(col_idx)
    base_formula = '=IFERROR(IF(AK2="SAP Payroll systems";CONCAT(A2;V2;M2);"-");"-")'

    for row in range(2, ws.max_row + 1):
        formula = base_formula.replace("2", str(row))
        ws[f"{col_letter}{row}"] = formula

    seen = set()
    rows_to_delete = []
    for row in range(2, ws.max_row + 1):
        val = ws[f"{col_letter}{row}"].value
        if val in seen:
            rows_to_delete.append(row)
        else:
            seen.add(val)

    for r in reversed(rows_to_delete):
        ws.delete_rows(r)

    ws.delete_cols(col_idx)

    wb.save(file_path)
    wb.close()
    log(f"✅ Column '{column_title}' formula added and removed after deleting {len(rows_to_delete)} duplicates.")

def process_files():
    try:
        out_path = get_unique_output_path(watch_dir, output_file)
        log(f"📁 Saving result to: {os.path.basename(out_path)}")

        sap_path = os.path.join(watch_dir, file_sap)
        wd_path = os.path.join(watch_dir, file_wd)

        log("📂 Loading input files")
        sap_df = pd.read_excel(sap_path)
        wd_df = pd.read_excel(wd_path)

        # Clean potential leftovers
        for col_to_drop in ["AbsenceDate_SAP", "Key_SAP"]:
            if col_to_drop in sap_df.columns:
                sap_df.drop(columns=[col_to_drop], inplace=True)

        # === CHANGE === remove A/AType filtering — use ALL absence types now
        # sap_df = sap_df[sap_df["A/AType"].isin(["AS01", "AX04", "AS03", "AH01", "AH02" ])].copy()

        all_columns = sap_df.columns.tolist()
        # Keep how WD key is built (not used later but left as in your version)
        if "Employee ID" in wd_df.columns and "Time Off date" in wd_df.columns:
            wd_df["Key_WD"] = wd_df["Employee ID"].astype(str) + "_" + wd_df["Time Off date"].dt.strftime("%Y%m%d")

        rows = []

        for _, row in sap_df.iterrows():
            start = row.get("Start Date")
            end = row.get("End Date")

            if pd.isnull(start) or pd.isnull(end):
                continue
            if pd.to_datetime(end) > max_valid_date:
                # Skip unrealistic "infinite" end dates
                continue

            def build_row(overrides):
                # Preserve your original pattern: keep 'required_columns' values; others as '-'
                # (not changing behavior intentionally)
                full_row = {col: row[col] if col in required_columns else "-" for col in all_columns}
                full_row.update(overrides)
                return full_row

            start = pd.to_datetime(start)
            end = pd.to_datetime(end)

            # === CHANGE === expand EVERY row by days (for all absence types)
            if start == end:
                rows.append(build_row({
                    "Start Date": start,
                    "End Date": end,
                    "AbsenceDate_SAP": start,
                    "Key_SAP": f"{row['Personnel Number']}_{start.strftime('%Y%m%d')}",
                    "PY": None
                }))
            else:
                for d in pd.date_range(start, end, freq="D"):
                    rows.append(build_row({
                        "Start Date": d,
                        "End Date": d,
                        "AbsenceDate_SAP": d,
                        "Key_SAP": f"{row['Personnel Number']}_{d.strftime('%Y%m%d')}",
                        "PY": None
                    }))

        df = pd.DataFrame(rows)

        # Mark PY as before
        if "PY" in df.columns:
            df["PY"] = df["PY"].apply(lambda x: "Python script" if pd.isna(x) else x)

        # Unique by Key_SAP
        if "Key_SAP" in df.columns:
            df = df.drop_duplicates(subset=["Key_SAP"], keep="first")

        # Normalize Start/End time
        for col in ["Start", "End time"]:
            if col in df.columns:
                df[col] = df[col].astype(str).replace({":  :": "-"})

        # Keep service columns at end (same behavior)
        cols = df.columns.tolist()
        for col in ["AbsenceDate_SAP", "Key_SAP", "PY"]:
            if col in cols:
                cols.remove(col)
        cols.extend(["AbsenceDate_SAP", "Key_SAP", "PY"])
        df = df[cols]

        # Before saving — remove service columns (as in your version)
        for col_to_remove in ["AbsenceDate_SAP", "Key_SAP"]:
            if col_to_remove in df.columns:
                df.drop(columns=[col_to_remove], inplace=True)

        log("📊 Saving results to Excel")
        df.to_excel(out_path, index=False)

        # Post-processing in Excel
        add_formula_columns(out_path)
        reorder_columns(out_path)
        add_formula_and_remove_duplicates(out_path, "#")

        log("✅ Processing completed successfully.")
    except Exception as e:
        log(f"❌ ERROR during processing: {e}")

def wait_for_files():
    log("🚀 Script started. Waiting for files")
    log("🔍 Watching for input files")
    while True:
        files = [f.lower() for f in os.listdir(watch_dir)]
        if file_sap.lower() in files and file_wd.lower() in files:
            log("📂 Detected both files. Starting processing.")
            process_files()
            break
        time.sleep(check_interval)

if __name__ == "__main__":
    os.makedirs(log_dir, exist_ok=True)
    wait_for_files()
