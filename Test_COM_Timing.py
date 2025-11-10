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
    RR1_node = aspen.Tree.FindNode(RR1_Path)
    t1 = time.perf_counter()
    RR1_node.Value = i
    t2 = time.perf_counter()
    aspen.Tree.FindNode(RR2_Path).Value = i
    t3 = time.perf_counter()
    set(RR3_Path,i)
    t4 = time.perf_counter()
    aspen.Engine.Run2() 
    t5 = time.perf_counter()
    Duty1_Node = aspen.Tree.FindNode(Duty1_Path)
    t6 = time.perf_counter()
    Duty1_Value = Duty1_Node.Value
    t7 = time.perf_counter()
    Duty2_Value = aspen.Tree.FindNode(Duty2_Path).Value
    t8 = time.perf_counter()
    Duty3_Value = get(Duty3_Path)
    t9 = time.perf_counter()
    t10 = time.perf_counter()
    
    results.append({
        "RR Value": i,
        "Duty1 Value": Duty1_Value,
        "Duty2 Value": Duty2_Value,
        "Duty3 Value": Duty3_Value,
        "Time RR1_node = aspen.Tree.FindNode(RR1_Path)": t1 - t0,
        "Time RR1_node.Value = i": t2 - t1,
        "Time aspen.Tree.FindNode(RR2_Path).Value = i": t3 - t2,
        "Time set(RR3_Path,i)": t4 - t3,
        "Time aspen.Engine.Run2() ": t5 - t4,
        "Time Duty1_Node = aspen.Tree.FindNode(Duty1_Path)": t6 - t5,
        "Time Duty1_Value = Duty1_Node.Value": t7 - t6,
        "Time Duty2_Value = aspen.Tree.FindNode(Duty2_Path).Value": t8 - t7,
        "Time Duty3_Value = get(Duty3_Path)": t9 - t8,
        "Time get time": t10 - t9
    })
    

# Export to Excel
wb = Workbook()
ws = wb.active
ws.title = "Aspen Timing"

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