from AspenController import AspenController

BKPPATH = r"C:\Users\erikv\OneDrive\Documents\CSC411\aspen_api\Framework Testing\Simulation\AspenFile.bkp"
RR_Path1 = r"\Data\Blocks\AB-CDE\Input\BASIS_RR"
RR_Path2 = r"\Data\Blocks\AB-CDE\Input\BASIS_RR2"
RR_Path3 = r"\Data\Blocks\AB-CDE\Input\BASIS_RR3"
Duty_Path = r"\Data\Blocks\AB-CDE\Output\REB_DUTY"

def compute(v):
    i, j, k = v
    sim1.set(RR_Path1, i)
    sim1.run()
    return float(sim1.get(Duty_Path))

ac = AspenController(BKPPATH, visible=False, suppress_dialogs=True)
x = ([3, 2, 5], [2, 5, 6], [8, 5, 6])

results = ac.evaluate(function=compute, values=x, num_workers=2)
print(results)
