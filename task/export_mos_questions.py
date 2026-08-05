# -*- coding: utf-8 -*-
"""Export latest MOS question texts to CSV (1 per subject) and Excel (3 sheets)."""
from __future__ import annotations

import csv
import json
from pathlib import Path

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "task" / "export_問題文"

SOURCES = [
    (
        "Word",
        ROOT / "MOSapp" / "MOS Word app" / "References" / "JSON" / "MOS模擬アプリ問題文一覧_Word.json",
    ),
    (
        "Excel",
        ROOT / "MOSapp" / "mos_xaml_app" / "References" / "JSON" / "MOS模擬アプリ問題文一覧.json",
    ),
    (
        "PowerPoint",
        ROOT
        / "MOSapp"
        / "Mos PowerPoint Mogi App"
        / "References"
        / "JSON"
        / "MOS模擬アプリ問題文一覧_PowerPoint.json",
    ),
]

HEADERS = ["プロジェクト番号", "タスク番号", "問題文", "問題文"]


def load_rows(json_path: Path) -> list[tuple[int, int, str]]:
    data = json.loads(json_path.read_text(encoding="utf-8-sig"))
    rows: list[tuple[int, int, str]] = []
    for project in data.get("projects", []):
        project_id = project.get("projectId")
        for task in project.get("tasks", []):
            task_id = task.get("taskId")
            description = task.get("description") or ""
            rows.append((project_id, task_id, description))
    return rows


def write_csv(path: Path, rows: list[tuple[int, int, str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)
        for project_id, task_id, description in rows:
            writer.writerow([project_id, task_id, description, description])


def write_excel(path: Path, sheets: dict[str, list[tuple[int, int, str]]]) -> None:
    if Workbook is None:
        raise RuntimeError("openpyxl is not installed")
    wb = Workbook()
    # Remove default sheet after creating real ones
    default = wb.active
    first = True
    for name, rows in sheets.items():
        if first:
            ws = default
            ws.title = name
            first = False
        else:
            ws = wb.create_sheet(name)
        ws.append(HEADERS)
        for project_id, task_id, description in rows:
            ws.append([project_id, task_id, description, description])
        # Widen columns a bit for readability
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 80
        ws.column_dimensions["D"].width = 80
    wb.save(path)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    sheets: dict[str, list[tuple[int, int, str]]] = {}

    for subject, json_path in SOURCES:
        if not json_path.exists():
            raise FileNotFoundError(json_path)
        rows = load_rows(json_path)
        sheets[subject] = rows
        csv_path = OUT_DIR / f"MOS問題文_{subject}.csv"
        write_csv(csv_path, rows)
        print(f"{subject}: {len(rows)} tasks from {json_path}")
        print(f"  -> {csv_path}")

    if Workbook is not None:
        xlsx_path = OUT_DIR / "MOS問題文_3科目.xlsx"
        write_excel(xlsx_path, sheets)
        print(f"Excel: {xlsx_path}")
    else:
        print("openpyxl not available; CSV only")


if __name__ == "__main__":
    main()
