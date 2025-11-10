import time
import win32com.client as win32
import numpy as np
import pandas as pd
from openpyxl import Workbook

BKPPATH = r"C:\Users\erikv\OneDrive\Documents\CSC411\aspen_api\Framework Testing\Simulation\AspenFile.bkp"
PROGID = "Apwn.Document" 
aspen = win32.Dispatch(PROGID)
aspen.InitFromArchive2(BKPPATH)


def set(path, value):
    node = aspen.Tree.FindNode(path)
    if node is None:
        raise KeyError(f"Path not found: {path}")
    node.Value = float(value)

def get(path):
    node = aspen.Tree.FindNode(path)
    if node is None:
        raise KeyError(f"Path not found: {path}")
    val = node.Value
    try:
        return float(val)
    except Exception as e:
        print(f"Error Getting {path}: {e}")
        return np.nan

RR1_Path = r"\Data\Blocks\AB-CDE\Input\BASIS_RR"
RR2_Path = r"\Data\Blocks\A-B\Input\BASIS_RR"
RR3_Path = r"\Data\Blocks\C-DE\Input\BASIS_RR"
RR4_Path = r"\Data\Blocks\D-E\Input\BASIS_RR"
Duty1_Path = r"\Data\Blocks\AB-CDE\Output\REB_DUTY"
Duty2_Path = r"\Data\Blocks\A-B\Output\REB_DUTY"
Duty3_Path = r"\Data\Blocks\C-DE\Output\REB_DUTY"
Duty4_Path = r"\Data\Blocks\D-E\Output\REB_DUTY"


results = []

for i in (1.1,1.2,1.3,1.4,1.5,1.6,1.7,1.8,1.9,2.0,2.1,2.2,2.3,2.4,2.5,2.6,2.7,2.8,2.9,3.0):
    t0 = time.perf_counter()
    set(RR1_Path,i)
    t1 = time.perf_counter()
    aspen.Engine.Run2() 
    t2 = time.perf_counter()
    duty1 = get(Duty1_Path)
    t3 = time.perf_counter()
    aspen.Engine.Reinit()
    t4 = time.perf_counter()
    aspen.Engine.Run2() 
    t5 = time.perf_counter()
    duty2 = get(Duty1_Path)
    t6 = time.perf_counter()
    results.append({
        "RR Value": i,
        "Duty Value after initial run": duty1,
        "Duty Value after reinit run": duty1,
        "Time to Set RR": t1-t0,
        "Time to run hot": t2-t1,
        "Time to get duty before": t3-t2,
        "Time to reinit": t4-t3,
        "Time to run cold": t5-t4,
        "Time to get duty after": t6-t5,
    })
    


# Export to Excel
wb = Workbook()
ws = wb.active
ws.title = "Aspen Timing 2"

# Write header
headers = list(results[0].keys())
ws.append(headers)

# Write rows
for row in results:
    ws.append([row[h] for h in headers])

# Save Excel file
outfile = r"C:\Users\erikv\OneDrive\Documents\CSC411\aspen_api\Framework Testing\TimingResults.xlsx"
wb.save(outfile)

print(f"Results exported to {outfile}")
aspen.Quit()