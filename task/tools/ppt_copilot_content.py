"""Parse and validate PowerPoint Copilot fixed-content MD files."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]

GENERIC_PLACEHOLDERS = frozenset(
    {
        "項目Aの確認",
        "項目Bの確認",
        "項目Cの確認",
        "要点A",
        "要点B",
        "要点C",
        "ダミー",
        "サンプル",
        "xxx",
        "...",
        "（記入）",
        "（ここに記入）",
    }
)

SET_HEADER_RE = re.compile(r"^##\s*セット\s*(\d+)\s*$", re.MULTILINE)
SLIDE_HEADER_RE = re.compile(r"^###\s*スライド\s*(\d+)\s*:\s*(.+?)\s*$", re.MULTILINE)
SUBTITLE_RE = re.compile(r"^\*\*サブタイトル:\*\*\s*(.+?)\s*$", re.MULTILINE)
SPEAKER_NOTES_RE = re.compile(r"^\*\*スピーカーノート:\*\*\s*(.+?)\s*$", re.MULTILINE)
BULLETS_HEADER_RE = re.compile(r"^\*\*箇条書き（そのまま転記）:\*\*\s*$", re.MULTILINE)
BODY_HEADER_RE = re.compile(r"^\*\*本文段落（そのまま転記）:\*\*\s*$", re.MULTILINE)


@dataclass
class SlideContent:
    slide_no: int
    title: str
    subtitle: str | None = None
    bullets: list[str] = field(default_factory=list)
    body_paragraphs: list[str] = field(default_factory=list)
    speaker_notes: str | None = None


def content_md_path(project_id: int) -> Path:
    return BASE / f"PowerPoint_類題_Copilot_Content_project{project_id}.md"


def _parse_bullets_or_body(block: str, header_re: re.Pattern[str]) -> list[str]:
    m = header_re.search(block)
    if not m:
        return []
    rest = block[m.end() :]
    items: list[str] = []
    for line in rest.splitlines():
        stripped = line.strip()
        if stripped.startswith("### ") or stripped.startswith("## "):
            break
        if stripped.startswith("**") and stripped.endswith(":**"):
            break
        if stripped.startswith("- "):
            items.append(stripped[2:].strip())
        elif stripped and not items:
            continue
        elif stripped and items:
            break
    return items


def _parse_slide_block(slide_no: int, title: str, block: str) -> SlideContent:
    subtitle_m = SUBTITLE_RE.search(block)
    notes_m = SPEAKER_NOTES_RE.search(block)
    bullets = _parse_bullets_or_body(block, BULLETS_HEADER_RE)
    body = _parse_bullets_or_body(block, BODY_HEADER_RE)
    return SlideContent(
        slide_no=slide_no,
        title=title.strip(),
        subtitle=subtitle_m.group(1).strip() if subtitle_m else None,
        bullets=bullets,
        body_paragraphs=body,
        speaker_notes=notes_m.group(1).strip() if notes_m else None,
    )


def parse_content_md(path: Path | str) -> dict[int, dict[int, SlideContent]]:
    text = Path(path).read_text(encoding="utf-8")
    result: dict[int, dict[int, SlideContent]] = {}
    set_matches = list(SET_HEADER_RE.finditer(text))
    if not set_matches:
        return result

    for idx, set_m in enumerate(set_matches):
        set_no = int(set_m.group(1))
        start = set_m.end()
        end = set_matches[idx + 1].start() if idx + 1 < len(set_matches) else len(text)
        set_block = text[start:end]
        slides: dict[int, SlideContent] = {}
        slide_matches = list(SLIDE_HEADER_RE.finditer(set_block))
        for sidx, slide_m in enumerate(slide_matches):
            slide_no = int(slide_m.group(1))
            title = slide_m.group(2)
            s_start = slide_m.end()
            s_end = slide_matches[sidx + 1].start() if sidx + 1 < len(slide_matches) else len(set_block)
            slides[slide_no] = _parse_slide_block(slide_no, title, set_block[s_start:s_end])
        result[set_no] = slides
    return result


def _collect_strings(content: SlideContent) -> list[str]:
    parts: list[str] = []
    if content.subtitle:
        parts.append(content.subtitle)
    parts.extend(content.bullets)
    parts.extend(content.body_paragraphs)
    if content.speaker_notes:
        parts.append(content.speaker_notes)
    return parts


def validate_content_md(
    project_id: int,
    content_by_set: dict[int, dict[int, SlideContent]],
    *,
    expected_sets: int = 5,
    expected_slides_per_set: dict[int, int] | None = None,
    slide_titles_by_set: dict[int, dict[int, str]] | None = None,
    forbidden_texts: set[str] | None = None,
    require_speaker_notes: bool = False,
    min_bullets: int = 3,
) -> list[str]:
    issues: list[str] = []
    forbidden = forbidden_texts or set()

    if len(content_by_set) < expected_sets:
        issues.append(f"content: expected {expected_sets} sets, found {len(content_by_set)}")

    all_strings: list[str] = []
    for set_no in range(1, expected_sets + 1):
        slides = content_by_set.get(set_no)
        if not slides:
            issues.append(f"content set{set_no}: missing")
            continue

        expected_count = (expected_slides_per_set or {}).get(set_no)
        if expected_count and len(slides) < expected_count:
            issues.append(f"content set{set_no}: expected {expected_count} slides, found {len(slides)}")

        titles = (slide_titles_by_set or {}).get(set_no, {})
        for slide_no, sc in sorted(slides.items()):
            if titles and slide_no in titles and sc.title != titles[slide_no]:
                issues.append(
                    f"content set{set_no} slide{slide_no}: title mismatch "
                    f"'{sc.title}' != JSON '{titles[slide_no]}'"
                )
            if len(sc.bullets) < min_bullets:
                issues.append(
                    f"content set{set_no} slide{slide_no}: bullets need>={min_bullets}, "
                    f"found {len(sc.bullets)}"
                )
            for b in sc.bullets:
                if b in GENERIC_PLACEHOLDERS or b.startswith("（"):
                    issues.append(f"content set{set_no} slide{slide_no}: placeholder bullet: {b!r}")
            if require_speaker_notes and not sc.speaker_notes:
                issues.append(f"content set{set_no} slide{slide_no}: speakerNotes required")
            for part in _collect_strings(sc):
                all_strings.append(part)
                for f in forbidden:
                    if f and f in part:
                        issues.append(f"content set{set_no} slide{slide_no}: forbidden text: {f}")

    seen: dict[str, list[str]] = {}
    for set_no, slides in content_by_set.items():
        for slide_no, sc in slides.items():
            for part in _collect_strings(sc):
                if len(part) < 8:
                    continue
                key = part.strip()
                loc = f"set{set_no}/slide{slide_no}"
                seen.setdefault(key, []).append(loc)
    for text, locs in seen.items():
        if len(locs) > 1:
            issues.append(f"content duplicate text across slides: {text!r} in {', '.join(locs)}")

    return issues


def load_content_for_project(project_id: int) -> dict[int, dict[int, SlideContent]]:
    path = content_md_path(project_id)
    if not path.exists():
        return {}
    return parse_content_md(path)


def slide_titles_from_json_sets(sets: list[dict]) -> dict[int, dict[int, str]]:
    out: dict[int, dict[int, str]] = {}
    for s in sets:
        sm = s.get("contentBlocks", {}).get("slideMap", [])
        out[s["setNo"]] = {slide["slideNo"]: slide["title"] for slide in sm}
    return out


def expected_slide_count(sets: list[dict]) -> dict[int, int]:
    return {s["setNo"]: len(s.get("contentBlocks", {}).get("slideMap", [])) for s in sets}
