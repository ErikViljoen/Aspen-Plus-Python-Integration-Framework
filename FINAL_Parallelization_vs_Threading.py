import os
import time
import shutil
import tempfile
import threading
import multiprocessing as mp
import numpy as np
import pandas as pd
import psutil
import pythoncom
import win32com.client as win32
import matplotlib.pyplot as plt

# =====================
# User configuration
# =====================
PROGID = "Apwn.Document"
BKPPATH = r"C:\Users\erikv\OneDrive\Documents\CSC411\aspen_api\Framework Testing\Simulation\AspenFile.bkp"
OUTDIR  = r"C:\Users\erikv\OneDrive\Documents\CSC411\aspen_api\Framework Testing"

# Aspen Data Browser paths
RR_Path   = r"\Data\Blocks\AB-CDE\Input\BASIS_RR"
Duty_Path = r"\Data\Blocks\AB-CDE\Output\REB_DUTY"

# Concurrency values to test
K_VALUES = list(range(1, 15))  # 1 through 14

# System sampling interval
SAMPLE_INTERVAL_S = 0.5

# Sweep points for the Aspen run
RR_POINTS = [round(x, 1) for x in np.arange(1.1, 3.1, 0.01)]  # 1.1 to 3.0 inclusive
N_CASES_PER_INSTANCE = len(RR_POINTS)


# =====================
# Worker logic, shared
# =====================
def _aspen_set(aspen, path, value):
    node = aspen.Tree.FindNode(path)
    if node is None:
        raise KeyError(f"Path not found: {path}")
    node.Value = float(value)

def _aspen_get(aspen, path):
    node = aspen.Tree.FindNode(path)
    if node is None:
        raise KeyError(f"Path not found: {path}")
    val = node.Value
    try:
        return float(val)
    except Exception:
        return np.nan

def _run_aspen_session(instance_id, bkppath):
    """
    Launch one Aspen session, run the RR sweep, return elapsed seconds.
    This function is safe to call from a thread or a process.
    """
    pythoncom.CoInitialize()
    tmpdir = tempfile.mkdtemp(prefix=f"aspen_{instance_id}_")
    try:
        local_bkp = os.path.join(tmpdir, "model.bkp")
        shutil.copy2(bkppath, local_bkp)

        aspen = win32.Dispatch(PROGID)
        try:
            aspen.Visible = False
        except Exception:
            pass
        aspen.InitFromArchive2(local_bkp)

        t0 = time.perf_counter()
        for rr in RR_POINTS:
            _aspen_set(aspen, RR_Path, rr)
            aspen.Engine.Run2()
            _ = _aspen_get(aspen, Duty_Path)
        elapsed = time.perf_counter() - t0

        try:
            aspen.Quit()
        except Exception:
            pass
        shutil.rmtree(tmpdir, ignore_errors=True)
        pythoncom.CoUninitialize()
        return elapsed
    except Exception as e:
        shutil.rmtree(tmpdir, ignore_errors=True)
        pythoncom.CoUninitialize()
        raise e


# =====================
# Threaded benchmark
# =====================
def _sample_system_while(threads_or_procs, interval_s=SAMPLE_INTERVAL_S):
    cpu_vals = []
    mem_vals = []
    psutil.cpu_percent(interval=None)
    while any(t.is_alive() for t in threads_or_procs):
        cpu_vals.append(psutil.cpu_percent(interval=None))
        mem_vals.append(psutil.virtual_memory().percent)
        time.sleep(interval_s)
    if not cpu_vals:
        cpu_vals = [psutil.cpu_percent(interval=None)]
        mem_vals = [psutil.virtual_memory().percent]
    return {
        "cpu_avg": float(np.mean(cpu_vals)),
        "cpu_peak": float(np.max(cpu_vals)),
        "ram_avg": float(np.mean(mem_vals)),
        "ram_peak": float(np.max(mem_vals)),
        "samples": len(cpu_vals),
    }

def run_threading_suite(k_values, bkppath):
    all_timings = {}  # K -> list of times
    util_rows = []

    for K in k_values:
        per_instance_times = []
        threads = []
        for i in range(K):
            t = threading.Thread(
                target=lambda idx: per_instance_times.append(_run_aspen_session(idx, bkppath)),
                args=(i + 1,)
            )
            threads.append(t)
            t.start()

        util_stats = _sample_system_while(threads, interval_s=SAMPLE_INTERVAL_S)
        for t in threads:
            t.join()

        all_timings[K] = per_instance_times
        util_rows.append({
            "K": K,
            "CPU_avg_percent": util_stats["cpu_avg"],
            "CPU_peak_percent": util_stats["cpu_peak"],
            "RAM_avg_percent": util_stats["ram_avg"],
            "RAM_peak_percent": util_stats["ram_peak"],
            "Samples": util_stats["samples"],
            "Avg_time_per_instance_s": float(np.nanmean(per_instance_times)) if per_instance_times else np.nan,
            "Median_time_per_instance_s": float(np.nanmedian(per_instance_times)) if per_instance_times else np.nan,
            "Total_instances_finished": int(np.sum(~pd.isna(per_instance_times))),
        })
        print(f"[Threading] K={K} times: {per_instance_times}")

    df_util = pd.DataFrame(util_rows).set_index("K")
    return all_timings, df_util


# =====================
# Process benchmark
# =====================
def _proc_worker(idx, bkppath, q):
    try:
        elapsed = _run_aspen_session(idx, bkppath)
        q.put(("ok", elapsed))
    except Exception as e:
        q.put(("err", str(e)))

def run_process_suite(k_values, bkppath):
    all_timings = {}
    util_rows = []

    for K in k_values:
        q = mp.Queue()
        procs = []
        for i in range(K):
            p = mp.Process(target=_proc_worker, args=(i + 1, bkppath, q))
            procs.append(p)
            p.start()

        util_stats = _sample_system_while(procs, interval_s=SAMPLE_INTERVAL_S)

        per_instance_times = []
        errs = 0
        for _ in range(K):
            status, payload = q.get()
            if status == "ok":
                per_instance_times.append(payload)
            else:
                errs += 1
                print(f"[Process] Worker error: {payload}")

        for p in procs:
            p.join()

        all_timings[K] = per_instance_times
        util_rows.append({
            "K": K,
            "CPU_avg_percent": util_stats["cpu_avg"],
            "CPU_peak_percent": util_stats["cpu_peak"],
            "RAM_avg_percent": util_stats["ram_avg"],
            "RAM_peak_percent": util_stats["ram_peak"],
            "Samples": util_stats["samples"],
            "Avg_time_per_instance_s": float(np.nanmean(per_instance_times)) if per_instance_times else np.nan,
            "Median_time_per_instance_s": float(np.nanmedian(per_instance_times)) if per_instance_times else np.nan,
            "Total_instances_finished": int(np.sum(~pd.isna(per_instance_times))),
            "Errors": int(errs),
        })
        print(f"[Process] K={K} times: {per_instance_times}")

    df_util = pd.DataFrame(util_rows).set_index("K")
    return all_timings, df_util


# =====================
# Helpers
# =====================
def pad_column(values, target_len):
    return values + [float("nan")] * (target_len - len(values))

def build_timing_sheet(all_timings):
    max_rows = max(len(v) for v in all_timings.values())
    sheet = {f"K={K}": pad_column(times, max_rows) for K, times in all_timings.items()}
    return pd.DataFrame(sheet)

def add_derived_metrics(df_util):
    df = df_util.copy()
    # Average elapsed per instance already present
    # Compute total time to finish the batch, approximate as max of instance times
    # We do not have the exact wall per K, so estimate as avg per instance
    # You can refine by recording wall time around each suite if needed
    df["Throughput_cases_per_min"] = (N_CASES_PER_INSTANCE * df["Total_instances_finished"]) / (df["Avg_time_per_instance_s"] * 60.0)
    # Speedup vs K=1
    if 1 in df.index and not np.isnan(df.loc[1, "Avg_time_per_instance_s"]):
        t1 = df.loc[1, "Avg_time_per_instance_s"]
        df["Speedup_vs_K1"] = t1 / df["Avg_time_per_instance_s"]
        df["Efficiency"] = df["Speedup_vs_K1"] / df.index
    else:
        df["Speedup_vs_K1"] = np.nan
        df["Efficiency"] = np.nan
    return df


# =====================
# Plotting
# =====================
def plot_cpu_ram(df_util, title, outpath):
    plt.figure(figsize=(8, 5))
    plt.plot(df_util.index, df_util["CPU_avg_percent"], marker="o", label="CPU average")
    plt.plot(df_util.index, df_util["CPU_peak_percent"], marker="o", label="CPU peak")
    plt.plot(df_util.index, df_util["RAM_avg_percent"], marker="s", label="RAM average")
    plt.plot(df_util.index, df_util["RAM_peak_percent"], marker="s", label="RAM peak")
    plt.xlabel("Number of Aspen instances, K")
    plt.ylabel("Percent")
    plt.title(title)
    plt.legend()
    plt.tight_layout()
    plt.savefig(outpath, dpi=150)
    print(f"Saved plot to {outpath}")

def plot_avg_time(df_util_thread, df_util_proc, outpath):
    plt.figure(figsize=(8, 5))
    plt.plot(df_util_thread.index, df_util_thread["Avg_time_per_instance_s"], marker="o", label="Threading")
    plt.plot(df_util_proc.index, df_util_proc["Avg_time_per_instance_s"], marker="o", label="Processes")
    plt.xlabel("Number of Aspen instances, K")
    plt.ylabel("Average time per instance, s")
    plt.title("Average time per instance vs K (threading vs processes)")
    plt.legend()
    plt.tight_layout()
    plt.savefig(outpath, dpi=150)
    print(f"Saved plot to {outpath}")

def plot_speedup(df_util_thread, df_util_proc, outpath):
    plt.figure(figsize=(8, 5))
    if "Speedup_vs_K1" in df_util_thread.columns:
        plt.plot(df_util_thread.index, df_util_thread["Speedup_vs_K1"], marker="o", label="Threading")
    if "Speedup_vs_K1" in df_util_proc.columns:
        plt.plot(df_util_proc.index, df_util_proc["Speedup_vs_K1"], marker="o", label="Processes")
    plt.xlabel("Number of Aspen instances, K")
    plt.ylabel("Speedup vs K = 1")
    plt.title("Speedup vs K (threading vs processes)")
    plt.legend()
    plt.tight_layout()
    plt.savefig(outpath, dpi=150)
    print(f"Saved plot to {outpath}")

def plot_throughput(df_util_thread, df_util_proc, outpath):
    plt.figure(figsize=(8, 5))
    plt.plot(df_util_thread.index, df_util_thread["Throughput_cases_per_min"], marker="o", label="Threading")
    plt.plot(df_util_proc.index, df_util_proc["Throughput_cases_per_min"], marker="o", label="Processes")
    plt.xlabel("Number of Aspen instances, K")
    plt.ylabel("Throughput, Aspen cases per minute")
    plt.title("Throughput vs K (threading vs processes)")
    plt.legend()
    plt.tight_layout()
    plt.savefig(outpath, dpi=150)
    print(f"Saved plot to {outpath}")


# =====================
# Main
# =====================
def main():
    os.makedirs(OUTDIR, exist_ok=True)
    out_xlsx = os.path.join(OUTDIR, "ParallelBenchmark.xlsx")

    # Threading suite
    thr_timings, df_thr_util = run_threading_suite(K_VALUES, BKPPATH)
    df_thr_util = add_derived_metrics(df_thr_util)
    df_thr_times = build_timing_sheet(thr_timings)

    # Process suite
    # Important on Windows: use spawn and guard with if __name__ == "__main__"
    proc_timings, df_proc_util = run_process_suite(K_VALUES, BKPPATH)
    df_proc_util = add_derived_metrics(df_proc_util)
    df_proc_times = build_timing_sheet(proc_timings)

    # Combined comparison sheet
    df_compare = pd.DataFrame({
        "Avg_time_thread_s": df_thr_util["Avg_time_per_instance_s"],
        "Avg_time_proc_s": df_proc_util["Avg_time_per_instance_s"],
        "CPU_avg_thread_percent": df_thr_util["CPU_avg_percent"],
        "CPU_avg_proc_percent": df_proc_util["CPU_avg_percent"],
        "RAM_avg_thread_percent": df_thr_util["RAM_avg_percent"],
        "RAM_avg_proc_percent": df_proc_util["RAM_avg_percent"],
        "Throughput_thread_cases_per_min": df_thr_util["Throughput_cases_per_min"],
        "Throughput_proc_cases_per_min": df_proc_util["Throughput_cases_per_min"],
        "Speedup_thread_vs_K1": df_thr_util["Speedup_vs_K1"],
        "Speedup_proc_vs_K1": df_proc_util["Speedup_vs_K1"],
        "Efficiency_thread": df_thr_util["Efficiency"],
        "Efficiency_proc": df_proc_util["Efficiency"],
    })

    # Save to Excel
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as xw:
        df_thr_times.to_excel(xw, sheet_name="Thread_per_instance_times", index=False)
        df_thr_util.reset_index().to_excel(xw, sheet_name="Thread_CPU_RAM_vs_K", index=False)
        df_proc_times.to_excel(xw, sheet_name="Proc_per_instance_times", index=False)
        df_proc_util.reset_index().to_excel(xw, sheet_name="Proc_CPU_RAM_vs_K", index=False)
        df_compare.reset_index(names="K").to_excel(xw, sheet_name="Comparison", index=False)

    print(f"Wrote results to {out_xlsx}")

    # Plots
    plot_cpu_ram(
        df_thr_util,
        title="System CPU and RAM vs K, threading",
        outpath=os.path.join(OUTDIR, "CPU_RAM_vs_K_threading.png"),
    )
    plot_cpu_ram(
        df_proc_util,
        title="System CPU and RAM vs K, processes",
        outpath=os.path.join(OUTDIR, "CPU_RAM_vs_K_processes.png"),
    )
    plot_avg_time(
        df_thr_util,
        df_proc_util,
        outpath=os.path.join(OUTDIR, "Avg_time_per_instance_vs_K.png"),
    )
    plot_speedup(
        df_thr_util,
        df_proc_util,
        outpath=os.path.join(OUTDIR, "Speedup_vs_K.png"),
    )
    plot_throughput(
        df_thr_util,
        df_proc_util,
        outpath=os.path.join(OUTDIR, "Throughput_vs_K.png"),
    )

if __name__ == "__main__":
    mp.set_start_method("spawn", force=True)
    main()
