import multiprocessing
import numpy as np
import pythoncom
import win32com.client
import tempfile
import shutil
import os

BKPPATH = r"C:\Users\erikv\OneDrive\Documents\CSC411\aspen_api\Framework Testing\Simulation\AspenFile.bkp"
PROGID = "Apwn.Document" 
RR_Path = r"\Data\Blocks\AB-CDE\Input\BASIS_RR"
Duty_Path = r"\Data\Blocks\AB-CDE\Output\REB_DUTY"

aspen = None
temp_dir = None

def init_aspen():
    global aspen, temp_dir
    pythoncom.CoInitialize()
    temp_dir = tempfile.mkdtemp(prefix="aspen_")
    shutil.copy2(BKPPATH, os.path.join(temp_dir, "model.bkp"))
    aspen = win32com.client.Dispatch(PROGID)
    aspen.Visible = False
    aspen.InitFromArchive2(os.path.join(temp_dir, "model.bkp"))

def cleanup_aspen():
    global aspen, temp_dir
    if aspen:
        aspen.Quit()
    pythoncom.CoUninitialize()
    if temp_dir:
        shutil.rmtree(temp_dir, ignore_errors=True)

def set(path, value):
    aspen.Tree.FindNode(path).Value = float(value)

def get(path):
    return float(aspen.Tree.FindNode(path).Value)

def simulate(arg):
    global aspen
    if arg is None:
        cleanup_aspen()
        return None
    set(RR_Path, arg)
    aspen.Engine.Run2()
    return get(Duty_Path)

if __name__ == "__main__":
    values = [round(x, 1) for x in np.arange(1.1, 3.1, 0.1)]
    with multiprocessing.Pool(processes=4, initializer=init_aspen) as pool:
        duties = pool.map(simulate, values, chunksize=1)
        pool.map(simulate, [None] * 4)