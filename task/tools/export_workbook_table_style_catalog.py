#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""開いている xlsx からテーブルスタイル全一覧（ギャラリー表示名付き）を CSV 出力。"""
from __future__ import annotations

import argparse
import csv
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path

import win32com.client
import win32gui
from pywinauto import Application

DEFAULT_WORKBOOK = Path(r"C:\Users\kouza\Desktop\MOS_類題_再作成_セット1_配送管理.xlsx")
DEFAULT_PREFIX = "MOS_類題_セット1_配送管理"


@dataclass
class StyleRecord:
    internal_name: str
    name_local: str
    gallery_name_raw: str
    gallery_name_mos: str


def log(msg: str) -> None:
    print(msg, flush=True)


def normalize_to_mos_format(raw: str) -> str:
    s = raw.strip()
    s = s.replace(", ", "、").replace(",", "、")
    s = re.sub(r"テーブル\s*スタイル\s*", "テーブルスタイル", s)
    s = re.sub(r"[（(](淡色|中間|濃色)[）)]\s*(\d+)", r"（\1）\2", s)
    return s


def family_of(internal: str) -> str:
    if internal.startswith("TableStyleLight"):
        return "淡色"
    if internal.startswith("TableStyleMedium"):
        return "中間"
    if internal.startswith("TableStyleDark"):
        return "濃色"
    return ""


def is_valid_gallery_name(text: str) -> bool:
    return bool(text) and text != "なし" and ("テーブル" in text or "スタイル" in text)


def attach_workbook(workbook_path: Path, sheet_index: int, table_index: int):
    if not workbook_path.exists():
        raise FileNotFoundError(f"workbook not found: {workbook_path}")

    wb = win32com.client.GetObject(str(workbook_path.resolve()))
    excel = wb.Application
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        excel.WindowState = -4137
    except Exception:
        pass

    ws = wb.Worksheets(sheet_index)
    ws.Activate()
    if ws.ListObjects.Count < table_index:
        raise RuntimeError(
            f"sheet {sheet_index} has {ws.ListObjects.Count} tables; need index {table_index}"
        )
    lo = ws.ListObjects(table_index)
    lo.Range.Cells(2, 1).Select()
    time.sleep(2.0)
    log(f"  接続: {wb.FullName}")
    log(f"  選択: {ws.Name} / {lo.Name}")
    return excel, wb


def get_table_styles(wb) -> list[tuple[str, str]]:
    styles: list[tuple[str, str]] = []
    for ts in wb.TableStyles:
        if ts.ShowAsAvailableTableStyle:
            styles.append((str(ts.Name), str(ts.NameLocal)))
    return styles


class ExcelGallerySession:
    def __init__(self, excel_hwnd: int):
        self.hwnd = excel_hwnd
        self.app = Application(backend="uia").connect(handle=excel_hwnd)
        self.win = self.app.window(handle=excel_hwnd)
        self._ensure_table_design_tab()

    def _ensure_table_design_tab(self) -> None:
        try:
            win32gui.SetForegroundWindow(self.hwnd)
        except Exception:
            pass
        try:
            self.win.set_focus()
        except Exception:
            pass
        time.sleep(0.3)
        for title in ("テーブル デザイン", "テーブルデザイン"):
            try:
                tab = self.win.child_window(title=title, control_type="TabItem")
                if tab.exists(timeout=1):
                    tab.select()
                    time.sleep(0.4)
                    return
            except Exception:
                continue

    def get_gallery_buttons(self) -> list[str]:
        items: list[str] = []
        seen: set[str] = set()
        for item in self.win.descendants(control_type="ListItem"):
            cls = item.element_info.class_name or ""
            if cls != "NetUIGalleryButton":
                continue
            text = (item.window_text() or "").strip()
            if not is_valid_gallery_name(text) or text in seen:
                continue
            seen.add(text)
            items.append(text)
        if not items:
            raise RuntimeError("ギャラリー ListItem が見つかりませんでした")
        return items


def write_catalog(path: Path, records: list[StyleRecord], source: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(
            [
                "no",
                "style_family",
                "internal_name",
                "name_local",
                "gallery_name_raw",
                "gallery_name_mos",
                "source_workbook",
            ]
        )
        for i, r in enumerate(records, start=1):
            w.writerow(
                [
                    i,
                    family_of(r.internal_name),
                    r.internal_name,
                    r.name_local,
                    r.gallery_name_raw,
                    r.gallery_name_mos,
                    source,
                ]
            )


def write_family_splits(out_dir: Path, records: list[StyleRecord], prefix: str, source: str) -> None:
    for family in ("淡色", "中間", "濃色"):
        rows = [r for r in records if family_of(r.internal_name) == family]
        csv_path = out_dir / f"{prefix}_テーブルスタイル全一覧_{family}.csv"
        txt_path = out_dir / f"{prefix}_テーブルスタイル全一覧_{family}.txt"
        with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["no", "style_family", "internal_name", "name_local", "gallery_name_mos"])
            for i, r in enumerate(rows, start=1):
                w.writerow([i, family, r.internal_name, r.name_local, r.gallery_name_mos])
        with txt_path.open("w", encoding="utf-8") as f:
            f.write(f"# {prefix} — {family}\n")
            f.write(f"# 取得元: {source}\n\n")
            for r in rows:
                f.write(r.gallery_name_mos + "\n")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--sheet-index", type=int, default=1)
    parser.add_argument("--table-index", type=int, default=1)
    parser.add_argument("--prefix", default=DEFAULT_PREFIX)
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=Path(__file__).resolve().parents[1],
    )
    args = parser.parse_args()

    task_dir = args.out_dir
    prefix = args.prefix

    log("開いているブックに接続...")
    excel, wb = attach_workbook(args.workbook, args.sheet_index, args.table_index)
    source = str(wb.FullName)
    styles = get_table_styles(wb)
    log(f"  TableStyles: {len(styles)} 件")

    session = ExcelGallerySession(int(excel.Hwnd))
    gallery_names = session.get_gallery_buttons()
    log(f"  ギャラリー項目: {len(gallery_names)} 件")

    if len(gallery_names) != len(styles):
        log(
            f"  警告: 件数不一致 COM={len(styles)} gallery={len(gallery_names)} "
            "(min 件数で対応)"
        )

    pair_count = min(len(styles), len(gallery_names))
    records: list[StyleRecord] = []
    for i in range(pair_count):
        internal, name_local = styles[i]
        raw = gallery_names[i]
        records.append(
            StyleRecord(
                internal_name=internal,
                name_local=name_local,
                gallery_name_raw=raw,
                gallery_name_mos=normalize_to_mos_format(raw),
            )
        )

    main_csv = task_dir / f"{prefix}_テーブルスタイル全一覧.csv"
    capture_dir = task_dir / "table_style_capture"
    capture_csv = capture_dir / f"{prefix}_テーブルスタイル全一覧.csv"

    write_catalog(main_csv, records, source)
    write_catalog(capture_csv, records, source)
    write_family_splits(task_dir, records, prefix, source)
    write_family_splits(capture_dir, records, prefix, source)

    log(f"\n完了: {len(records)} 件")
    log(f"  {main_csv}")
    log(f"  {capture_csv}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
