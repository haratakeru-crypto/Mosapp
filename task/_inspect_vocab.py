# -*- coding: utf-8 -*-
import zipfile, os, re, xml.etree.ElementTree as ET
from pathlib import Path

desktop = Path(r"c:\Users\kouza\Desktop")
src = next(p for p in desktop.glob("*.xlsm") if "MOSExcel" in p.name)
print("SRC", src, src.stat().st_size)

out_dir = Path(r"c:\Users\kouza\source\repos\MOSapp\task\_vocab_extract")
out_dir.mkdir(exist_ok=True)
extract = out_dir / "extracted"
if extract.exists():
    import shutil; shutil.rmtree(extract)
with zipfile.ZipFile(src, "r") as z:
    z.extractall(extract)

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

# shared strings
ss_path = extract / "xl" / "sharedStrings.xml"
ss_root = ET.parse(ss_path).getroot()
strings = []
for si in ss_root.findall("m:si", NS):
    texts = [t.text or "" for t in si.findall(".//m:t", NS)]
    strings.append("".join(texts))
print("SHARED", len(strings))
(out_dir / "shared_strings.txt").write_text("\n".join(f"{i}\t{s}" for i,s in enumerate(strings)), encoding="utf-8")

# workbook sheets
wb = ET.parse(extract / "xl" / "workbook.xml").getroot()
rels = ET.parse(extract / "xl" / "_rels" / "workbook.xml.rels").getroot()
RNS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
sheets_info = []
for sh in wb.findall("m:sheets/m:sheet", NS):
    name = sh.attrib["name"]
    rid = sh.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
    state = sh.attrib.get("state", "visible")
    target = rid_to_target[rid]
    sheets_info.append((name, state, target))
    print("SHEET", name, state, target)

def col_index(ref):
    m = re.match(r"([A-Z]+)", ref)
    col = 0
    for ch in m.group(1):
        col = col*26 + (ord(ch)-64)
    return col

def dump_sheet(target, name):
    path = extract / "xl" / target
    root = ET.parse(path).getroot()
    rows = {}
    max_col = 0
    for c in root.findall("m:sheetData/m:row/m:c", NS):
        ref = c.attrib.get("r")
        if not ref: continue
        m = re.match(r"([A-Z]+)(\d+)", ref)
        if not m: continue
        col = col_index(m.group(1)); row = int(m.group(2))
        t = c.attrib.get("t")
        v = c.find("m:v", NS)
        is_el = c.find("m:is", NS)
        val = ""
        if t == "s" and v is not None and v.text is not None:
            val = strings[int(v.text)]
        elif t == "inlineStr" and is_el is not None:
            val = "".join(x.text or "" for x in is_el.findall(".//m:t", NS))
        elif v is not None and v.text is not None:
            val = v.text
        rows.setdefault(row, {})[col] = val
        max_col = max(max_col, col)
    lines = [f"# {name} | rows={len(rows)} maxcol={max_col}"]
    for r in sorted(rows):
        cells = [rows[r].get(c, "") for c in range(1, max_col+1)]
        lines.append(f"R{r}\t" + "\t".join(cells))
    outp = out_dir / f"sheet_{name.replace(' ','_').replace('/','_')}.tsv"
    outp.write_text("\n".join(lines), encoding="utf-8")
    print("WROTE", outp, "rows", len(rows))

for name, state, target in sheets_info:
    dump_sheet(target, name)

# table
table = (extract / "xl" / "tables" / "table1.xml").read_text(encoding="utf-8")
(out_dir / "table1.xml").write_text(table, encoding="utf-8")
print("TABLE", table[:500])

# vba strings
vba = (extract / "xl" / "vbaProject.bin").read_bytes()
(out_dir / "vbaProject.bin").write_bytes(vba)

def ascii_strings(b, n=4):
    out=[]; cur=[]
    for x in b:
        if 32 <= x <= 126: cur.append(chr(x))
        else:
            if len(cur) >= n: out.append("".join(cur))
            cur=[]
    if len(cur)>=n: out.append("".join(cur))
    return out

def utf16le_strings(b, n=3):
    out=[]; cur=[]
    i=0
    while i+1 < len(b):
        lo, hi = b[i], b[i+1]
        if hi==0 and 32 <= lo <= 126:
            cur.append(chr(lo)); i+=2; continue
        # Japanese-ish BMP
        if hi != 0 or lo < 32:
            cp = lo + (hi<<8)
            if 0x3000 <= cp <= 0x9FFF or 0xFF00 <= cp <= 0xFFEF:
                cur.append(chr(cp)); i+=2; continue
        if len(cur) >= n: out.append("".join(cur))
        cur=[]; i+=2 if hi==0 or True else 1
        # always advance 2 for utf16 scan when broken - already advanced in continue; here advance 2 from failed pair start... messy
    # simpler re-scan:
    return out

# better utf16 extract
def utf16_scan(b, n=3):
    out=[]; i=0
    while i+1 < len(b):
        if b[i+1]==0 and 32<=b[i]<=126:
            j=i; s=[]
            while j+1 < len(b) and b[j+1]==0 and 32<=b[j]<=126:
                s.append(chr(b[j])); j+=2
            if len(s)>=n: out.append("".join(s))
            i=j if j>i else i+2
        else:
            # try JP char
            cp = b[i] + (b[i+1]<<8)
            if 0x3040 <= cp <= 0x30FF or 0x4E00 <= cp <= 0x9FFF or 0xFF66 <= cp <= 0xFF9F or 0x3000 <= cp <= 0x303F:
                j=i; s=[]
                while j+1 < len(b):
                    cp2 = b[j] + (b[j+1]<<8)
                    if (0x3040 <= cp2 <= 0x30FF or 0x4E00 <= cp2 <= 0x9FFF or 0xFF66 <= cp2 <= 0xFF9F
                        or 0x3000 <= cp2 <= 0x303F or (b[j+1]==0 and 32<=b[j]<=126)):
                        s.append(chr(cp2)); j+=2
                    else: break
                if len(s)>=n: out.append("".join(s))
                i=j if j>i else i+2
            else:
                i+=2
    return out

asc = ascii_strings(vba, 4)
u16 = utf16_scan(vba, 3)
interesting = [s for s in asc if re.search(r"Sub |Function |Module|ThisWorkbook|Attribute|MsgBox|Range|Cells|Button|CommandButton|UserForm|Random|Answer|Check|Score|Sheet|正解|不正解|Value|Click|Private|Public|Option|Dim ", s)]
(out_dir / "vba_ascii_all.txt").write_text("\n".join(dict.fromkeys(x for x in asc if len(x)>=5)), encoding="utf-8")
(out_dir / "vba_ascii_interesting.txt").write_text("\n".join(dict.fromkeys(interesting)), encoding="utf-8")
(out_dir / "vba_utf16.txt").write_text("\n".join(dict.fromkeys(u16)), encoding="utf-8")
print("VBA ascii interesting count", len(interesting))
print("VBA utf16 sample:")
for s in list(dict.fromkeys(u16))[:80]:
    print(" ", s)

# ctrl / vml
for p in (extract / "xl" / "ctrlProps").glob("*.xml"):
    print("CTRL", p.name, p.read_text(encoding="utf-8"))
for p in (extract / "xl" / "drawings").glob("*.vml"):
    print("VML", p.name)
    print(p.read_text(encoding="utf-8")[:2000])
for p in (extract / "xl" / "worksheets" / "_rels").glob("*.rels"):
    print("RELS", p.name, p.read_text(encoding="utf-8"))

print("DONE")