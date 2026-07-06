"""Parse PP修正版問題文一覧.csv into operation master records."""

from __future__ import annotations

import csv
import json
import re
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
CSV_PATH = BASE / "PP修正版問題文一覧.csv"
OUT_PATH = BASE / "類題Json" / "_pp_csv_master.json"

_BRACKET_UI = re.compile(r"\[([^\]]+)\]")


def _extract_ui_tokens(steps: str) -> list[str]:
    return list(dict.fromkeys(_BRACKET_UI.findall(steps or "")))


def parse_csv(path: Path | None = None) -> list[dict]:
    path = path or CSV_PATH
    rows: list[dict] = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("科目") != "PowerPoint":
                continue
            pid = int(row["プロジェクトID"])
            if pid > 6:
                continue
            tid = int(row["タスクID"])
            problem = (row.get("問題文") or "").strip()
            steps = (row.get("解答手順") or "").strip()
            intro = ""
            body = problem
            if "\n" in problem:
                intro, body = problem.split("\n", 1)
                intro = intro.strip()
                body = body.strip()
            rows.append(
                {
                    "projectId": pid,
                    "taskId": tid,
                    "intro": intro,
                    "problemBody": body,
                    "problemText": problem,
                    "answerSteps": steps,
                    "uiTokens": _extract_ui_tokens(steps),
                }
            )
    return rows


def group_by_project(rows: list[dict]) -> dict[int, list[dict]]:
    out: dict[int, list[dict]] = {i: [] for i in range(1, 7)}
    for r in rows:
        out[r["projectId"]].append(r)
    for pid in out:
        out[pid].sort(key=lambda x: x["taskId"])
    return out


def main() -> None:
    rows = parse_csv()
    grouped = group_by_project(rows)
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(
        json.dumps({"tasks": rows, "byProject": grouped}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    for pid in range(1, 7):
        print(f"Project {pid}: {len(grouped[pid])} tasks")
    print(f"wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
