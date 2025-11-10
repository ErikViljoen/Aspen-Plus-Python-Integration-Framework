import os
import time
import shutil
import tempfile
import threading
import numpy as np
import pandas as pd
import psutil
import pythoncom
from threading import Lock
import win32com.client as win32
import matplotlib.pyplot as plt
from collections import defaultdict


PROGID = "Apwn.Document"  
BKPPATH = r"C:\Users\erikv\OneDrive\Documents\CSC411\aspen_api\Framework Testing\Simulation\AspenFile.bkp"

RR_Path   = r"\Data\Blocks\AB-CDE\Input\BASIS_RR"
Duty_Path = r"\Data\Blocks\AB-CDE\Output\REB_DUTY"

def run_aspen_instance(instance_id, bkppath, elapsed_sink):
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
        
    
    pythoncom.CoInitialize()
    tmpdir = tempfile.mkdtemp(prefix=f"aspen_{instance_id}_")
    try:
        # make a private copy of the bkp in this directory
        local_bkp = os.path.join(tmpdir, "model.bkp")
        shutil.copy2(bkppath, local_bkp)

        # launch Aspen and load the private copy
        aspen = win32.Dispatch(PROGID)
        try:
            aspen.Visible = False
        except Exception:
            pass
        aspen.InitFromArchive2(local_bkp)

        print(f"[Instance {instance_id}] Aspen launched in {tmpdir}")
        
        t0 = time.perf_counter()
        for i in [round(x, 1) for x in np.arange(1.1, 3.1, 0.1)]:
            set(RR_Path,i)
            aspen.Engine.Run2() 
            duty1 = get(Duty_Path)
        elapsed = time.perf_counter() - t0
        print(f"[Instance {instance_id}] Finished in {elapsed:.2f} seconds")
        elapsed_sink.append(elapsed)
        try:
            aspen.Quit()
        except Exception:
            pass
        # cleanup the private folder
        if tmpdir and os.path.isdir(tmpdir):
            shutil.rmtree(tmpdir, ignore_errors=True)
        pythoncom.CoUninitialize()
    except Exception:
        # cleanup if something fails
        shutil.rmtree(tmpdir, ignore_errors=True)
        pythoncom.CoUninitialize()
        raise
    
def sample_system_while(threads, interval_s=0.5):
    """Sample system CPU and RAM until all threads finish. Returns dict with averages and peaks."""
    cpu_vals = []
    mem_vals = []

    # prime cpu_percent
    psutil.cpu_percent(interval=None)

    while any(t.is_alive() for t in threads):
        cpu_vals.append(psutil.cpu_percent(interval=None))           # percent of total system
        mem_vals.append(psutil.virtual_memory().percent)             # system RAM percent
        time.sleep(interval_s)

    if not cpu_vals:
        # edge case if batch finished too fast
        cpu_vals = [psutil.cpu_percent(interval=None)]
        mem_vals = [psutil.virtual_memory().percent]

    return {
        "cpu_avg": float(np.mean(cpu_vals)),
        "cpu_peak": float(np.max(cpu_vals)),
        "ram_avg": float(np.mean(mem_vals)),
        "ram_peak": float(np.max(mem_vals)),
        "samples": len(cpu_vals),
    }
def pad_column(values, target_len):
    # pad with NaN so all columns have equal length
    return values + [float("nan")] * (target_len - len(values))

# choose which K values to test
k_values = list(range(1, 15))  # 1 through 14

all_timings = {}  # K -> list of per instance times
util_by_k  = []     
for K in k_values:
    per_instance_times = []
    threads = []
    for i in range(K):
        t = threading.Thread(target=run_aspen_instance, args=(i + 1, BKPPATH, per_instance_times))
        threads.append(t)
        t.start()
    util_stats = sample_system_while(threads, interval_s=0.5)
    for t in threads:
        t.join()
    all_timings[K] = per_instance_times
    util_by_k.append({
        "K": K,
        "CPU_avg_percent": util_stats["cpu_avg"],
        "CPU_peak_percent": util_stats["cpu_peak"],
        "RAM_avg_percent": util_stats["ram_avg"],
        "RAM_peak_percent": util_stats["ram_peak"],
        "Samples": util_stats["samples"],
        "Avg_time_per_instance_s": float(np.nanmean(per_instance_times)) if per_instance_times else np.nan,
        "Total_instances_finished": int(np.sum(~pd.isna(per_instance_times))),
    })
    print(f"K={K} times: {per_instance_times}")
    print(f"K={K} util: CPU avg {util_stats['cpu_avg']:.1f}% peak {util_stats['cpu_peak']:.1f}%, "
          f"RAM avg {util_stats['ram_avg']:.1f}% peak {util_stats['ram_peak']:.1f}%")

# Build per instance timing table for Excel
max_rows = max(len(v) for v in all_timings.values())
timing_sheet = {f"K={K}": pad_column(times, max_rows) for K, times in all_timings.items()}
df_timings = pd.DataFrame(timing_sheet)

# Build CPU and RAM vs K table
df_util = pd.DataFrame(util_by_k).set_index("K")

# Save to Excel
out_xlsx = r"C:\Users\erikv\OneDrive\Documents\CSC411\aspen_api\Framework Testing\ParallelBenchmark.xlsx"
with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xw:
    df_timings.to_excel(xw, sheet_name="Per_instance_times", index=False)
    df_util.reset_index().to_excel(xw, sheet_name="CPU_RAM_vs_K", index=False)

print(f"Wrote results to {out_xlsx}")

# Plot CPU and RAM vs K
plt.figure(figsize=(8,5))
plt.plot(df_util.index, df_util["CPU_avg_percent"], marker="o", label="CPU average")
plt.plot(df_util.index, df_util["CPU_peak_percent"], marker="o", label="CPU peak")
plt.plot(df_util.index, df_util["RAM_avg_percent"], marker="s", label="RAM average")
plt.plot(df_util.index, df_util["RAM_peak_percent"], marker="s", label="RAM peak")
plt.xlabel("Number of Aspen instances, K")
plt.ylabel("Percent")
plt.title("System CPU and RAM vs number of Aspen instances")
plt.legend()
plt.tight_layout()

plot_path = r"C:\Users\erikv\OneDrive\Documents\CSC411\aspen_api\Framework Testing\CPU_RAM_vs_K.png"
plt.savefig(plot_path, dpi=150)
print(f"Saved plot to {plot_path}")