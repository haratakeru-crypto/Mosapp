# -*- coding: utf-8 -*-
from _read_xlsx import parse_tasks, read_sheet
import json
from collections import Counter

pairs = [
    (r"c:\Users\kouza\source\repos\MOSapp\MOSapp\MOS Word app\References\Word類題_問題と解答.xlsx", "_word_ruitai.json"),
    (r"c:\Users\kouza\source\repos\MOSapp\MOSapp\Mos PowerPoint Mogi App\References\PowerPoint類題.xlsx", "_ppt_ruitai.json"),
]
for src, dst in pairs:
    rows = read_sheet(src)
    header = {k: rows[0].get(k) for k in sorted(rows[0])} if rows else {}
    tasks = parse_tasks(src)
    with open(dst, "w", encoding="utf-8") as f:
        json.dump({"header": header, "count": len(tasks), "tasks": tasks}, f, ensure_ascii=False)
    print(dst, len(tasks), header)
    print(dict(Counter(x["projectId"] for x in tasks)))
    if tasks:
        print("sample", tasks[0]["projectId"], tasks[0]["taskId"], tasks[0]["q"][:80])
