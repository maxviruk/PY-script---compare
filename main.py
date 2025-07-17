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

# === SCRIPT EXECUTION ===
try:
    # Ask whether to run aut_cleanup_wd_file.py first
    run_cleanup = input("Do you want to run aut_cleanup_wd_file.py? (y/n): ").strip().lower()
    if run_cleanup == "y":
        print("[INFO] Step 1: Running aut_cleanup_wd_file.py")
        write_log("Running aut_cleanup_wd_file.py")
        subprocess.run(["python", "aut_cleanup_wd_file.py"], check=True)
        write_log("Finished aut_cleanup_wd_file.py")
    else:
        write_log("Skipped aut_cleanup_wd_file.py")

    print("[INFO] Step 2: Running aut_compare.py")
    write_log("Running aut_compare.py")
    subprocess.run(["python", "aut_compare.py"], check=True)
    write_log("Finished aut_compare.py")

    print("[INFO] Step 3: Running aut_join_files.py")
    write_log("Running aut_join_files.py")
    subprocess.run(["python", "aut_join_files.py"], check=True)
    write_log("Finished aut_join_files.py")

    print("[INFO] All steps completed.")
    write_log("All scripts completed successfully.")

except subprocess.CalledProcessError as e:
    error_msg = f"Script failed: {e}"
    print("[ERROR]", error_msg)
    write_log(error_msg)

except Exception as e:
    general_error = f"Unexpected error: {e}"
    print("[ERROR]", general_error)
    write_log(general_error)
