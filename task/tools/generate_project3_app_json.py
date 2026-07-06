import json
import re
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
src = json.loads(
    (BASE / "類題Json" / "MOS_類題_project3_配置別_5セット_問題文.json").read_text(
        encoding="utf-8"
    )
)
app_sets = []
for s in src["sets"]:
    items = []
    for i, t in enumerate(s["tasks"], 1):
        desc = re.sub(r"^タスク3-\d+　", "", t)
        if i == 1:
            desc = s["problemStatement"] + "\n" + desc
        items.append({"taskId": i, "description": desc})
    app_sets.append(
        {
            "setNo": s["setNo"],
            "theme": s["theme"],
            "workbook": s["workbook"],
            "projectId": 3,
            "tasks": items,
        }
    )
out = {"sourceFile": "project3.xlsx", "variantType": "layout", "sets": app_sets}
path = BASE / "類題Json" / "MOS_類題_project3_配置別_5セット_問題文_アプリ用.json"
path.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
print(f"wrote {path}")
