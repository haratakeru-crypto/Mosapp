"""Generate MD / app JSON from variant problem JSON. Validate consistency."""
import argparse
import json
import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]

PROJECT_CONFIG = {
    2: {"task_count": 7, "forbidden": {"下半期売上", "社員リスト", "担当者別売上", "業務予定", "参加者一覧"}},
    3: {"task_count": 7, "forbidden": {"売上一覧", "販売実績", "スキルアップ検定結果"}},
    4: {"task_count": 4, "forbidden": {"上半期売上", "５年間売上", "下半期売上", "商品別売上"}},
    5: {"task_count": 6, "forbidden": {"売上実績", "商品別売上", "月別売上"}},
    6: {"task_count": 4, "forbidden": {"売上一覧", "販売実績"}},
    7: {"task_count": 7, "forbidden": {"売上報告", "受注明細", "下半期売上"}},
    8: {"task_count": 7, "forbidden": {"イベント売上", "試験結果", "申込一覧"}},
    9: {"task_count": 6, "forbidden": {"学生名簿", "売上報告", "担当者リスト", "申込一覧"}},
    10: {"task_count": 8, "forbidden": {"担当者リスト", "出張精算", "売上一覧", "業務予定", "売上集計", "在庫管理"}},
}


def json_path(project_id: int) -> Path:
    return BASE / "類題Json" / f"MOS_類題_project{project_id}_配置別_5セット_問題文.json"


def md_path(project_id: int) -> Path:
    return BASE / f"MOS_類題_project{project_id}_配置別_5セット_問題文.md"


def app_json_path(project_id: int) -> Path:
    return BASE / "類題Json" / f"MOS_類題_project{project_id}_配置別_5セット_問題文_アプリ用.json"


def write_md(data: dict) -> str:
    pid = data["projectId"]
    lines = [
        f"# MOSスペシャリスト Excel 類題 — Project{pid} 配置別（5セット）",
        "",
        f"元ファイル {data['sourceFile']} の操作をもとに、",
        "レイアウトを差別化した 5 セットの類題ブック用問題文です。",
        "",
    ]
    for s in data["sets"]:
        lines.append(f"## セット{s['setNo']}：{s['workbook']}")
        lines.append("")
        lines.append(f"{s['problemStatement']}  ")
        lines.append("問題）  ")
        for t in s["tasks"]:
            lines.append(f"{t}  ")
        lines.append("")
    return "\n".join(lines)


def write_app_json(data: dict) -> dict:
    pid = data["projectId"]
    app_sets = []
    for s in data["sets"]:
        items = []
        for i, t in enumerate(s["tasks"], 1):
            desc = re.sub(rf"^タスク{pid}-\d+　", "", t)
            if i == 1:
                desc = s["problemStatement"] + "\n" + desc
            items.append({"taskId": i, "description": desc})
        app_sets.append(
            {
                "setNo": s["setNo"],
                "theme": s["theme"],
                "workbook": s["workbook"],
                "projectId": pid,
                "tasks": items,
            }
        )
    return {
        "sourceFile": data["sourceFile"],
        "variantType": data.get("variantType", "layout"),
        "sets": app_sets,
    }


def validate(data: dict, cfg: dict) -> list[str]:
    issues: list[str] = []
    pid = data["projectId"]
    forbidden = cfg["forbidden"]
    task_count = cfg["task_count"]
    md = write_md(data)
    layouts = []
    table_starts = []
    header_rows = []
    col_width_sets = []
    visuals = []
    formats = []

    for s in data["sets"]:
        for name in s["layout"].get("sheets", []):
            if name in forbidden:
                issues.append(f"set{s['setNo']}: forbidden sheet {name}")
        if len(s["tasks"]) != task_count:
            issues.append(f"set{s['setNo']}: task count {len(s['tasks'])} != {task_count}")
        if s["workbook"] not in md:
            issues.append(f"set{s['setNo']}: workbook missing in md")
        for t in s["tasks"]:
            if t not in md:
                issues.append(f"set{s['setNo']}: task missing in md: {t[:50]}")
        layout = s["layout"]
        layouts.append((layout.get("tableStart"), layout.get("headerRow"), layout.get("tabColor")))
        table_starts.append(layout.get("tableStart"))
        header_rows.append(layout.get("headerRow"))
        col_width_sets.append(tuple(layout.get("colWidths", [])))
        visuals.append(layout.get("visual"))
        formats.append(layout.get("format"))

    if len(set(layouts)) != 5:
        issues.append(f"layout signatures not unique: {layouts}")
    if len(set(table_starts)) != 5:
        issues.append(f"tableStart not unique: {table_starts}")
    if len(set(header_rows)) != 5:
        issues.append(f"headerRow not unique: {header_rows}")
    if len(set(col_width_sets)) != 5:
        issues.append("colWidths not unique across sets")
    if len(set(visuals)) != 5:
        issues.append(f"visual not unique: {visuals}")
    if len(set(formats)) != 5:
        issues.append(f"format not unique: {formats}")

    for s in data["sets"]:
        expected_keys = [f"task{pid}_{i}" for i in range(1, task_count + 1)]
        for key in expected_keys:
            if key not in s.get("taskParams", {}):
                issues.append(f"set{s['setNo']}: missing {key}")
    return issues


def build_project(project_id: int) -> list[str]:
    path = json_path(project_id)
    if not path.exists():
        return [f"JSON not found: {path}"]
    data = json.loads(path.read_text(encoding="utf-8"))
    cfg = PROJECT_CONFIG[project_id]
    issues = validate(data, cfg)
    if issues:
        return issues
    md_path(project_id).write_text(write_md(data), encoding="utf-8")
    app_json_path(project_id).write_text(
        json.dumps(write_app_json(data), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return []


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--project", type=int, action="append", help="project id (repeatable)")
    parser.add_argument("--all", action="store_true")
    parser.add_argument("--validate-only", action="store_true")
    args = parser.parse_args()
    ids = list(PROJECT_CONFIG.keys()) if args.all else (args.project or [])
    if not ids:
        parser.print_help()
        sys.exit(1)
    failed = False
    for pid in ids:
        if args.validate_only:
            data = json.loads(json_path(pid).read_text(encoding="utf-8"))
            issues = validate(data, PROJECT_CONFIG[pid])
        else:
            issues = build_project(pid)
        if issues:
            print(f"Project{pid}: FAIL", issues)
            failed = True
        else:
            print(f"Project{pid}: OK")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
