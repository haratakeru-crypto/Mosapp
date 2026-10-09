# -*- coding: utf-8 -*-
"""Export VBA via Excel COM if available."""
import sys
from pathlib import Path

out = Path(r"c:\Users\kouza\source\repos\MOSapp\task\_vocab_extract\vba_export")
out.mkdir(parents=True, exist_ok=True)
log = []

try:
    import win32com.client  # type: ignore
except Exception as e:
    log.append(f"no win32com: {e}")
    # try late binding via comtypes or Dispatch via powershell-written path
    try:
        import comtypes.client  # type: ignore
        log.append("comtypes available")
    except Exception as e2:
        log.append(f"no comtypes: {e2}")

desktop = Path(r"c:\Users\kouza\Desktop")
src = next(p for p in desktop.glob("*.xlsm") if "MOSExcel" in p.name)

# Prefer win32com
excel = None
try:
    import win32com.client as win32
    excel = win32.DispatchEx("Excel.Application")
except Exception as e:
    log.append(f"DispatchEx failed: {e}")
    try:
        import win32com.client as win32
        excel = win32.gencache.EnsureDispatch("Excel.Application")
    except Exception as e2:
        log.append(f"EnsureDispatch failed: {e2}")

if excel is None:
    # fallback: write a note and exit - PowerShell COM will be used separately
    (out / "_log.txt").write_text("\n".join(log) + "\nNeed PowerShell COM", encoding="utf-8")
    print("NO_EXCEL_COM")
    sys.exit(0)

excel.Visible = False
excel.DisplayAlerts = False
# Trust access to VBA project must be enabled by user; try anyway
wb = excel.Workbooks.Open(str(src), ReadOnly=True)
try:
    vbproj = wb.VBProject
    for comp in vbproj.VBComponents:
        name = comp.Name
        code = ""
        try:
            cm = comp.CodeModule
            n = cm.CountOfLines
            if n > 0:
                code = cm.Lines(1, n)
        except Exception as e:
            code = f"<<error reading: {e}>>"
        (out / f"{name}.bas.txt").write_text(code, encoding="utf-8")
        log.append(f"exported {name} lines={code.count(chr(10))+1 if code else 0}")
        print("EXPORTED", name)
except Exception as e:
    log.append(f"VBProject access failed: {e}")
    print("VBPROJECT_FAIL", e)
    # still dump sheet formulas via COM
    for sh in wb.Sheets:
        try:
            used = sh.UsedRange
            vals = used.Value
            fmls = used.Formula
            (out / f"sheetvals_{sh.Name}.txt").write_text(repr(vals), encoding="utf-8")
            (out / f"sheetfml_{sh.Name}.txt").write_text(repr(fmls), encoding="utf-8")
            log.append(f"sheet dump {sh.Name}")
            print("SHEET", sh.Name)
        except Exception as e2:
            log.append(f"sheet fail {sh.Name}: {e2}")
finally:
    wb.Close(False)
    excel.Quit()

(out / "_log.txt").write_text("\n".join(log), encoding="utf-8")
print("DONE")
