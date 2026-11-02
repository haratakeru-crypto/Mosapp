# -*- coding: utf-8 -*-
import json, re, sys, zipfile, xml.etree.ElementTree as ET

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

def col_of(ref):
    m = re.match(r"([A-Z]+)", ref or "")
    letters = m.group(1) if m else "A"
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n

def load_shared(z):
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root.findall("a:si", NS):
        out.append("".join(t.text or "" for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")))
    return out

def cell_text(c, shared):
    t = c.get("t")
    v = c.find("a:v", NS)
    if v is None or v.text is None:
        is_el = c.find("a:is", NS)
        if is_el is not None:
            return "".join(n.text or "" for n in is_el.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
        return ""
    if t == "s":
        return shared[int(v.text)]
    return v.text

def read_sheet(path):
    with zipfile.ZipFile(path) as z:
        shared = load_shared(z)
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = wb.find("a:sheets", NS)
        first = sheets.find("a:sheet", NS)
        rid = first.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = None
        for rel in rels:
            if rel.get("Id") == rid:
                target = rel.get("Target")
                break
        if not target:
            target = "worksheets/sheet1.xml"
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        root = ET.fromstring(z.read(target))
        rows = []
        for row in root.find("a:sheetData", NS).findall("a:row", NS):
            cells = {}
            for c in row.findall("a:c", NS):
                cells[col_of(c.get("r"))] = cell_text(c, shared)
            rows.append(cells)
        return rows

def parse_tasks(path):
    rows = read_sheet(path)
    header = rows[0]
    inv = { (header.get(k) or "").strip(): k for k in header }
    pcol = inv.get("プロジェクト")
    tcol = inv.get("タスク")
    qcol = inv.get("問題文")
    scol = inv.get("解答手順")
    out = []
    for r in rows[1:]:
        try:
            pid = int(float(str(r.get(pcol, "")).strip()))
            tid = int(float(str(r.get(tcol, "")).strip()))
        except Exception:
            continue
        out.append({
            "projectId": pid,
            "taskId": tid,
            "q": (r.get(qcol) or "").replace("\r\n", "\n").strip(),
            "steps": (r.get(scol) or "").strip()
        })
    return out

if __name__ == "__main__":
    path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else None
    rows = read_sheet(path)
    header = {k: rows[0].get(k) for k in sorted(rows[0])} if rows else {}
    tasks = parse_tasks(path)
    payload = {"header": header, "count": len(tasks), "tasks": tasks}
    text = json.dumps(payload, ensure_ascii=False)
    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(text)
    else:
        sys.stdout.reconfigure(encoding="utf-8")
        print(text)
