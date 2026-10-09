# -*- coding: utf-8 -*-
from pathlib import Path
import re

base = Path(r"c:/Users/kouza/source/repos/MOSapp/task/_vocab_extract/extracted/xl/worksheets")
out = Path(r"c:/Users/kouza/source/repos/MOSapp/task/_vocab_extract/sheet_formulas.txt")
lines = []
for f in sorted(base.glob("sheet*.xml")):
    t = f.read_text(encoding="utf-8")
    lines.append(f"==== {f.name} ====")
    for m in re.finditer(r'<c r="([^"]+)"[^>]*>.*?(?:<f[^>]*>([^<]*)</f>)?(?:<v>([^<]*)</v>)?', t, re.S):
        pass
    for m in re.finditer(r'<f(?: [^>]*)?>([^<]+)</f>', t):
        lines.append(f"  FORMULA: {m.group(1)}")
    for tag in ["legacyDrawing", "controls", "oleObjects", "drawing", "dataValidations", "conditionalFormatting"]:
        if tag in t:
            lines.append(f"  HAS:{tag}")
    # dump full for test sheets
    if f.name in ("sheet2.xml", "sheet4.xml", "sheet5.xml"):
        lines.append(t)
    lines.append("")
out.write_text("\n".join(lines), encoding="utf-8")
print("wrote", out)

# drawings
dbase = Path(r"c:/Users/kouza/source/repos/MOSapp/task/_vocab_extract/extracted/xl/drawings")
dout = Path(r"c:/Users/kouza/source/repos/MOSapp/task/_vocab_extract/drawings_dump.txt")
parts = []
for f in sorted(dbase.glob("*")):
    if f.is_file():
        parts.append(f"==== {f.name} ====")
        parts.append(f.read_text(encoding="utf-8", errors="replace"))
dout.write_text("\n".join(parts), encoding="utf-8")
print("wrote", dout)
