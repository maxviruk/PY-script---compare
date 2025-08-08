import os
import subprocess
from datetime import datetime


# === LOG SETUP ===
log_dir = os.path.join(os.getcwd(), "PY - Logs")
os.makedirs(log_dir, exist_ok=True)
log_path = os.path.join(log_dir, "processing_log_0.txt")

def write_log(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as log_file:
        log_file.write(f"[{timestamp}] {message}\n")

def run_script(script_name, step_desc):
    try:
        print(f"🚀 {step_desc}")
        write_log(f"⏳ Running {script_name}")
        subprocess.run(["python", script_name], check=True)
        write_log(f"✅ Finished {script_name}")
        print(f"✅ {script_name} completed")
        return True
    except subprocess.CalledProcessError as e:
        error_msg = f"[ERROR] {script_name} failed: {e}"
        print(error_msg)
        write_log(error_msg)
        print("\n[STOP] Execution stopped due to error above.")
        return False
    except Exception as e:
        general_error = f"[ERROR] {script_name} unexpected error: {e}"
        print(general_error)
        write_log(general_error)
        print("\n[STOP] Execution stopped due to unexpected error above.")
        return False

if __name__ == "__main__":
    # Ask about WD cleanup
    run_cleanup = input("Do you want to run aut_cleanup_wd_file.py? (y/n): ").strip().lower()

    if run_cleanup == "y":
        if not run_script("aut_cleanup_wd_file.py", "Step 1: Running aut_cleanup_wd_file.py"):
            exit(1)
    else:
        write_log("Skipped aut_cleanup_wd_file.py")
        print("Skipped aut_cleanup_wd_file.py")

    # Step 2: Run EOP cleanup
    if not run_script("aut_cleaup_eop_file.py", "Step 2: Running aut_cleaup_eop_file.py"):
        exit(2)

    # Step 3: Join files
    if not run_script("aut_join_files.py", "Step 3: Running aut_join_files.py"):
        exit(3)

    print("\n[INFO] ✅  All steps completed successfully ✅")
    write_log("✅ All scripts completed successfully ✅")
