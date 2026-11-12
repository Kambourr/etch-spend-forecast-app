"""Recalculate an exported workbook with an independent formula engine and compare with the app engine."""
import json, os, shutil, subprocess, sys, warnings
warnings.filterwarnings("ignore")
import formulas
from openpyxl import load_workbook
xlsx, expected = sys.argv[1], json.load(open(sys.argv[2]))
out = xlsx + "_calc"; shutil.rmtree(out, ignore_errors=True)
xl = formulas.ExcelModel().loads(xlsx).finish(); xl.calculate(); xl.write(dirpath=out)
w = load_workbook(os.path.join(out, os.listdir(out)[0]), data_only=True)
S, M = w["SUMMARY"], w["MODEL"]; N = expected["nprog"]
bad = 0
def chk(name, got, want, tol=0.01):
    global bad
    ok = got is not None and abs(got - want) <= tol
    if not ok: bad += 1
    print(("ok   " if ok else "FAIL ") + name, got, want)
k = expected["keys"]
for cell, key in (("B6", "aop"), ("B7", "fc"), ("B8", "var"), ("B10", "opex"), ("B11", "capex"), ("B12", "ytd"), ("B13", "fte"), ("B14", "dep"), ("B15", "runrate")):
    chk(f"Summary {cell} {key}", S[cell].value, k[key])
fp = 9 + 9 * N + 10 + 7 * N  # FP row: FB[0]+7N, FB[0]=SP+10, SP=9+9N
fp = (9 + 9 * N) + 10 + 7 * N
for i, v in enumerate(expected["forTotal"]): chk(f"forecast total month {i+1}", M.cell(fp + 5, 3 + i).value, v)
dp = fp + 9 + 2 + 5 * N  # DP = FP+11+5N
dp = fp + 11 + 5 * N
for i, v in enumerate(expected["dep"]): chk(f"depreciation month {i+1}", M.cell(dp + 1, 3 + i).value, v)
for i, v in enumerate(expected["bridge"][1:5]): chk(f"bridge step {i+1}", S.cell(27 + i, 2).value, v)
print("Summary overall:", S["F3"].value, "| expected:", expected["overall"])
if S["F3"].value != expected["overall"]: bad += 1
errs = 0
for ws in w:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and (c.value.startswith("#") or "Error" in c.value): errs += 1
print("error cells:", errs); bad += errs
print("EXPORT VERIFIED" if not bad else f"{bad} PROBLEMS"); sys.exit(1 if bad else 0)
