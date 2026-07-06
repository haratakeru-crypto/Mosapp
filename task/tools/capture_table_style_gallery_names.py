#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel クイックスタイル（リボン埋め込みギャラリー）から
色付きギャラリー表示名を UI Automation で取得する。
名称を UI で確認 → ホバーで名称表示 → スクリーンショット保存。

禁止: 新規ブック作成・新規テーブル作成は行わない。
既存 xlsx 内の ListObject を選択してギャラリーを表示する。
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path

import mss
import win32com.client
import win32gui
from PIL import Image, ImageDraw, ImageFont
from pywinauto import Application, Desktop, mouse

DEFAULT_WORKBOOK = Path(r"C:\MOSTest\Excel365\Tab1\project1.xlsx")


@dataclass
class StyleRecord:
    internal_name: str
    name_local: str
    gallery_name_raw: str
    gallery_name_mos: str
    screenshot_path: str


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


def get_table_styles(wb) -> list[tuple[str, str]]:
    styles: list[tuple[str, str]] = []
    for ts in wb.TableStyles:
        if ts.ShowAsAvailableTableStyle:
            styles.append((str(ts.Name), str(ts.NameLocal)))
    return styles


def setup_excel(workbook_path: Path, sheet_index: int, table_index: int):
    """既存ブック・既存テーブルのみ使用（新規作成しない）。"""
    if not workbook_path.exists():
        raise FileNotFoundError(f"workbook not found: {workbook_path}")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.WindowState = -4137  # xlMaximized

    wb = excel.Workbooks.Open(str(workbook_path.resolve()))
    ws = wb.Worksheets(sheet_index)
    ws.Activate()
    if ws.ListObjects.Count < table_index:
        raise RuntimeError(
            f"sheet {sheet_index} has {ws.ListObjects.Count} tables; need index {table_index}"
        )
    lo = ws.ListObjects(table_index)
    lo.Range.Cells(2, 1).Select()
    time.sleep(2.0)
    log(f"  使用ブック: {workbook_path.name}")
    log(f"  使用シート: {ws.Name} / テーブル: {lo.Name}")
    return excel, wb, lo


def grab_region(left: int, top: int, width: int, height: int) -> Image.Image:
    width = max(1, int(width))
    height = max(1, int(height))
    with mss.MSS() as sct:
        shot = sct.grab({"left": int(left), "top": int(top), "width": width, "height": height})
        return Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")


def save_with_name_banner(img: Image.Image, path: Path, internal_name: str, name_mos: str) -> None:
    banner_h = 44
    font = ImageFont.load_default()
    out = Image.new("RGB", (max(img.width, 520), img.height + banner_h), (255, 255, 255))
    draw = ImageDraw.Draw(out)
    draw.rectangle((0, 0, out.width, banner_h), fill=(32, 32, 32))
    draw.text((8, 6), internal_name, fill=(200, 200, 200), font=font)
    draw.text((8, 22), name_mos, fill=(255, 255, 255), font=font)
    out.paste(img, (0, banner_h))
    out.save(path, format="PNG")


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

    def get_gallery_buttons(self) -> list[tuple[str, object]]:
        items: list[tuple[str, object]] = []
        seen: set[str] = set()
        for item in self.win.descendants(control_type="ListItem"):
            cls = item.element_info.class_name or ""
            if cls != "NetUIGalleryButton":
                continue
            text = (item.window_text() or "").strip()
            if not is_valid_gallery_name(text) or text in seen:
                continue
            seen.add(text)
            items.append((text, item))
        if not items:
            raise RuntimeError("ギャラリー ListItem が見つかりませんでした")
        return items

    def get_quick_styles_pane_rect(self):
        try:
            gallery = self.win.child_window(title="クイック スタイル", control_type="Pane")
            if gallery.exists(timeout=1):
                return gallery.rectangle()
        except Exception:
            pass
        return None

    def capture_gallery_overview(self, shot_path: Path) -> None:
        pane = self.get_quick_styles_pane_rect()
        if pane:
            img = grab_region(pane.left, pane.top, pane.width(), pane.height())
        else:
            rect = self.win.rectangle()
            img = grab_region(rect.left, rect.top + 80, min(rect.width(), 1400), 400)
        img.save(shot_path, format="PNG")

    def find_visible_tooltip_rect(self, name_raw: str):
        name_mos = normalize_to_mos_format(name_raw)
        desktop = Desktop(backend="uia")
        candidates = []
        for w in desktop.windows():
            cls = w.element_info.class_name or ""
            title = (w.window_text() or "").strip()
            if not title:
                continue
            if "ToolTip" in cls or "NetUITooltip" in cls:
                if name_raw in title or name_mos in title:
                    candidates.append(w)
            elif name_raw in title or name_mos in title:
                if "Excel" not in title and "Cursor" not in title:
                    candidates.append(w)
        if candidates:
            return candidates[0].rectangle()
        return None

    def capture_after_name_displayed(
        self,
        index: int,
        name_raw: str,
        internal_name: str,
        name_mos: str,
        shot_path: Path,
        tooltip_wait: float = 0.9,
    ) -> None:
        buttons = self.get_gallery_buttons()
        if index >= len(buttons):
            raise IndexError(f"gallery index {index} out of range ({len(buttons)})")

        confirmed_name, item = buttons[index]
        if confirmed_name != name_raw:
            log(f"    警告: 名称不一致 index={index} expected={name_raw!r} got={confirmed_name!r}")

        try:
            win32gui.SetForegroundWindow(self.hwnd)
            self.win.set_focus()
        except Exception:
            pass

        try:
            item.scroll_into_view()
            time.sleep(0.25)
        except Exception:
            pass

        try:
            item_rect = item.rectangle()
        except Exception:
            pane = self.get_quick_styles_pane_rect()
            if pane:
                img = grab_region(pane.left, pane.top, pane.width(), pane.height())
                save_with_name_banner(img, shot_path, internal_name, name_mos)
            return

        cx = (item_rect.left + item_rect.right) // 2
        cy = (item_rect.top + item_rect.bottom) // 2
        mouse.move(coords=(cx, cy))
        time.sleep(tooltip_wait)

        tooltip_rect = self.find_visible_tooltip_rect(confirmed_name)

        if tooltip_rect and tooltip_rect.width() > 10 and tooltip_rect.height() > 5:
            left = min(tooltip_rect.left, item_rect.left) - 12
            top = min(tooltip_rect.top, item_rect.top) - 12
            right = max(tooltip_rect.right, item_rect.right) + 12
            bottom = max(tooltip_rect.bottom, item_rect.bottom) + 12
        else:
            left = item_rect.left - 40
            top = max(0, item_rect.top - 100)
            right = item_rect.right + 40
            bottom = item_rect.bottom + 20

        width = right - left
        height = bottom - top

        if width < 200 or height < 60:
            pane = self.get_quick_styles_pane_rect()
            if pane:
                left, top, width, height = pane.left, pane.top, pane.width(), pane.height()
            else:
                width = max(width, 520)
                height = max(height, 140)

        img = grab_region(left, top, width, height)
        if img.width < 50 or img.height < 30:
            pane = self.get_quick_styles_pane_rect()
            if pane:
                img = grab_region(pane.left, pane.top, pane.width(), pane.height())

        save_with_name_banner(img, shot_path, internal_name, name_mos)


def write_csv(path: Path, records: list[StyleRecord]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(
            [
                "internal_name",
                "name_local",
                "gallery_name_raw",
                "gallery_name_mos",
                "screenshot_path",
            ]
        )
        for r in records:
            w.writerow(
                [
                    r.internal_name,
                    r.name_local,
                    r.gallery_name_raw,
                    r.gallery_name_mos,
                    r.screenshot_path,
                ]
            )


def write_project1_catalog(path: Path, records: list[StyleRecord]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["no", "style_family", "internal_name", "name_local", "gallery_name_mos"])
        for i, r in enumerate(records, start=1):
            w.writerow([i, family_of(r.internal_name), r.internal_name, r.name_local, r.gallery_name_mos])


def write_family_splits(out_dir: Path, records: list[StyleRecord], prefix: str) -> None:
    for family in ("淡色", "中間", "濃色"):
        rows = [r for r in records if family_of(r.internal_name) == family]
        csv_path = out_dir / f"{prefix}_{family}.csv"
        txt_path = out_dir / f"{prefix}_{family}.txt"
        with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["no", "style_family", "internal_name", "name_local", "gallery_name_mos"])
            for i, r in enumerate(rows, start=1):
                w.writerow([i, family, r.internal_name, r.name_local, r.gallery_name_mos])
        with txt_path.open("w", encoding="utf-8") as f:
            f.write(f"# project1.xlsx — {prefix}_{family}\n")
            f.write(f"# 取得元: {DEFAULT_WORKBOOK}\n\n")
            for r in rows:
                f.write(r.gallery_name_mos + "\n")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=0, help="0=全件")
    parser.add_argument(
        "--workbook",
        type=Path,
        default=DEFAULT_WORKBOOK,
        help="既存 xlsx（新規ブック禁止）",
    )
    parser.add_argument("--sheet-index", type=int, default=1, help="1=試験結果")
    parser.add_argument("--table-index", type=int, default=1, help="シート内テーブル番号")
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "table_style_capture",
    )
    args = parser.parse_args()

    out_dir = args.out_dir
    shot_dir = out_dir / "screenshots"
    shot_dir.mkdir(parents=True, exist_ok=True)
    task_dir = Path(__file__).resolve().parents[1]

    csv_auto = out_dir / "MOS_project1_ギャラリー名_自動取得.csv"
    csv_catalog = out_dir / "MOS_project1_スタイル一覧_ギャラリー表示名.csv"
    gallery_shot = shot_dir / "quick_styles_gallery_overview.png"

    log("Excel セットアップ（既存 project1 テーブル使用）...")
    excel, wb, lo = setup_excel(args.workbook, args.sheet_index, args.table_index)
    hwnd = int(excel.Hwnd)
    styles = get_table_styles(wb)
    log(f"  COM テーブルスタイル数: {len(styles)}")

    session = ExcelGallerySession(hwnd)
    gallery_buttons = session.get_gallery_buttons()
    log(f"  ギャラリー項目数: {len(gallery_buttons)}")

    session.capture_gallery_overview(gallery_shot)
    log(f"  ギャラリー全体スクショ: {gallery_shot}")

    if len(gallery_buttons) != len(styles):
        log(
            f"  警告: 件数不一致 COM={len(styles)} gallery={len(gallery_buttons)} "
            "(順序対応は min 件数で実施)"
        )

    pair_count = min(len(styles), len(gallery_buttons))
    if args.limit and args.limit > 0:
        pair_count = min(pair_count, args.limit)

    records: list[StyleRecord] = []
    log("名称確認 → ホバー表示 → スクリーンショット...")
    for i in range(pair_count):
        internal, name_local = styles[i]
        raw, _ = gallery_buttons[i]
        mos = normalize_to_mos_format(raw)

        item_shot = shot_dir / f"{internal}.png"
        session.capture_after_name_displayed(i, raw, internal, mos, item_shot)

        records.append(
            StyleRecord(
                internal_name=internal,
                name_local=name_local,
                gallery_name_raw=raw,
                gallery_name_mos=mos,
                screenshot_path=str(item_shot),
            )
        )
        log(f"  [{i + 1}/{pair_count}] {internal} | {mos} -> {item_shot.name}")

    write_csv(csv_auto, records)
    write_project1_catalog(csv_catalog, records)
    write_family_splits(out_dir, records, "MOS_project1_スタイル一覧_ギャラリー表示名")
    write_family_splits(task_dir, records, "MOS_project1_スタイル一覧_ギャラリー表示名")

    # task/ 直下にもメイン CSV をコピー
    write_project1_catalog(task_dir / "MOS_project1_スタイル一覧_ギャラリー表示名.csv", records)

    log(f"\n完了: {len(records)} 件")
    log(f"  {csv_auto}")
    log(f"  {csv_catalog}")

    known = {
        "TableStyleMedium10": "オレンジ、テーブルスタイル（中間）10",
        "TableStyleDark11": "緑、テーブルスタイル（濃色）11",
    }
    log("\n教材照合:")
    for internal, expected in known.items():
        hit = next((r for r in records if r.internal_name == internal), None)
        if hit:
            ok = "OK" if hit.gallery_name_mos == expected else "要確認"
            log(f"  {internal}: {hit.gallery_name_mos} [{ok}]")
        else:
            log(f"  {internal}: 未取得")

    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        excel.Quit()
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
