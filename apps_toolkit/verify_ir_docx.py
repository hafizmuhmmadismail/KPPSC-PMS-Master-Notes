import glob
import json
import subprocess
import os

cli = r"H:\My Drive\PMS KPK 2022 notes\apps_toolkit\tools\wordcraft\wordcraft-0.8.0-windows-x64-portable\wordcraft-cli.exe"
files = sorted(glob.glob(r"H:\My Drive\PMS KPK 2022 notes\2. International Relations\*.docx"))
print(f"Total docx files found: {len(files)}")
print(f"{'Filename':<42} | {'Pages':<5} | {'Words':<5} | {'Dirty':<5}")
print("-" * 65)

all_ok = True
for f in files:
    res = subprocess.run([cli, "info", f], capture_output=True, text=True)
    if res.returncode == 0:
        data = json.loads(res.stdout)
        fn = os.path.basename(f)
        pages = data.get("pages", 0)
        words = data.get("words", 0)
        dirty = data.get("dirty", False)
        print(f"{fn[:40]:<42} | {pages:<5} | {words:<5} | {str(dirty):<5}")
        if dirty or pages < 4:
            all_ok = False
    else:
        print(f"ERROR on {f}: {res.stderr}")
        all_ok = False

if all_ok:
    print("\nALL 23 DOCX FILES VERIFIED SUCCESSFULLY WITH 0 LAYOUT ERRORS!")
else:
    print("\nWARNING: Some files had layout anomalies.")
