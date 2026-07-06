"""Slide map rendering for PowerPoint variant JSON."""

from __future__ import annotations


def _format_shape(shape: dict) -> str:
    parts = [shape.get("type", "図形")]
    if shape.get("insertName") and shape["insertName"] != shape.get("type"):
        parts.append(f"挿入名={shape['insertName']}")
    if shape.get("count", 1) > 1:
        parts.append(f"×{shape['count']}")
    if shape.get("fillColorDesc"):
        parts.append(f"塗り={shape['fillColorDesc']}")
    if shape.get("text"):
        parts.append(f"文字={shape['text']!r}")
    if shape.get("label"):
        parts.append(f"ラベル={shape['label']!r}")
    if shape.get("notes"):
        parts.append(f"({shape['notes']})")
    return " / ".join(parts)


def render_slide_map_md(slide_map: list[dict], pre_task: dict | None = None) -> list[str]:
    lines: list[str] = []
    pre = pre_task or {}
    for s in slide_map:
        no = s["slideNo"]
        title = s.get("title", "")
        layout = s.get("layout", "")
        parts = [f"- **スライド{no}**（{layout}）: {title}"]
        if s.get("subtitle"):
            parts.append(f"  - サブタイトル: {s['subtitle']}")
        if s.get("bullets"):
            for b in s["bullets"]:
                parts.append(f"  - {b}")
        if s.get("objects"):
            parts.append(f"  - オブジェクト: {', '.join(s['objects'])}")
        if s.get("shapes"):
            parts.append("  - 配置図形:")
            for sh in s["shapes"]:
                parts.append(f"    - {_format_shape(sh)}")
        if s.get("notes"):
            parts.append(f"  - ※ {s['notes']}")
        lines.extend(parts)
    if pre.get("hiddenSlides"):
        lines.append(f"- **非表示スライド（事前）**: なし（{pre['hiddenSlides']}は表示のまま）")
    else:
        lines.append("- **非表示スライド（事前）**: なし")
    return lines


def render_slide_content_block(slides: dict[int, object]) -> list[str]:
    """Render fixed slide body text for Copilot paste blocks."""
    lines: list[str] = []
    for slide_no in sorted(slides.keys()):
        sc = slides[slide_no]
        title = getattr(sc, "title", "") if not isinstance(sc, dict) else sc.get("title", "")
        lines.append(f"  スライド{slide_no}「{title}」")
        subtitle = getattr(sc, "subtitle", None) if not isinstance(sc, dict) else sc.get("subtitle")
        bullets = getattr(sc, "bullets", []) if not isinstance(sc, dict) else sc.get("bullets", [])
        body = (
            getattr(sc, "body_paragraphs", [])
            if not isinstance(sc, dict)
            else sc.get("body_paragraphs", [])
        )
        notes = getattr(sc, "speaker_notes", None) if not isinstance(sc, dict) else sc.get("speaker_notes")
        if subtitle:
            lines.append(f"    サブタイトル: {subtitle}")
        if bullets:
            lines.append("    箇条書き:")
            for b in bullets:
                lines.append(f"    - {b}")
        if body:
            lines.append("    本文段落:")
            for p in body:
                lines.append(f"    - {p}")
        if notes:
            lines.append(f"    スピーカーノート: {notes}")
        lines.append("")
    return lines
