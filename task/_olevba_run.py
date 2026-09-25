# -*- coding: utf-8 -*-
from oletools.olevba import VBA_Parser
from pathlib import Path

src = next(p for p in Path(r"c:\Users\kouza\Desktop").glob("*.xlsm") if "0904" in p.name)
print("SRC", src)
vp = VBA_Parser(str(src))
out = Path(r"c:\Users\kouza\source\repos\MOSapp\task\_vocab_extract\vba_olevba.txt")
parts = []
for filename, stream_path, vba_filename, vba_code in vp.extract_macros():
    parts.append("==== " + str(vba_filename) + " ====")
    parts.append(vba_code or "")
    print("FILE", vba_filename, "len", len(vba_code or ""))
out.write_text("\n".join(parts), encoding="utf-8")
print("wrote", out)
vp.close()

# older workbook
old_dir = Path(r"C:\Users\kouza\Desktop\MOSExcel365模擬試験①")
old = next(old_dir.glob("*.xlsm"))
print("OLD", old, old.stat().st_size)
import zipfile, xml.etree.ElementTree as ET
with zipfile.ZipFile(old) as z:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    for sh in wb.findall("m:sheets/m:sheet", NS):
        print("OLD_SHEET", sh.attrib.get("name"), sh.attrib.get("state", "visible"))
