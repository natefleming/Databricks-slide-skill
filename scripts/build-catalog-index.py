#!/usr/bin/env python3
"""
Build Architecture Catalog Index

Reads system_architecture_catalog.pptx and outputs architecture_catalog.json
with metadata for each slide: index, title, subtitle, kind (section vs diagram),
tags, talk track preview, and section grouping.

Usage:
    python build-catalog-index.py
"""

import json
import re
import sys
from pathlib import Path

try:
    from pptx import Presentation
except ImportError:
    print("Error: python-pptx is required. Install with: pip3 install python-pptx")
    sys.exit(1)

SCRIPT_DIR = Path(__file__).parent
SKILL_DIR = SCRIPT_DIR.parent
CATALOG_PPTX = SKILL_DIR / "assets" / "databricks" / "system_architecture_catalog.pptx"
OUTPUT_JSON = SKILL_DIR / "assets" / "databricks" / "architecture_catalog.json"

# Slides to skip (1-based): slide 1 = disclaimer, slide 2 = changelog
SKIP_SLIDES = {1, 2}

# Section header: low shape count + typically a title-only layout
SECTION_MAX_SHAPES = 4


def extract_tags_and_talk_track(slide):
    """Extract tags and talk track from slide notes.

    Tags are lines starting with 'tag:' (comma-separated values).
    Talk track is everything else in the notes.
    """
    tags = []
    talk_track_lines = []

    try:
        if not slide.has_notes_slide:
            return tags, ""
        notes_text = slide.notes_slide.notes_text_frame.text
    except Exception:
        return tags, ""

    if not notes_text:
        return tags, ""

    for line in notes_text.split("\n"):
        stripped = line.strip()
        # Match "tag:" or "tag: " at start of line
        if re.match(r'^tags?\s*:', stripped, re.IGNORECASE):
            tag_part = re.sub(r'^tags?\s*:\s*', '', stripped, flags=re.IGNORECASE)
            for t in tag_part.split(","):
                t = t.strip()
                if t:
                    tags.append(t)
        elif stripped:
            talk_track_lines.append(stripped)

    # Also check for standalone level tags like "L200" at start of notes
    if not tags and talk_track_lines:
        first = talk_track_lines[0]
        if re.match(r'^L\d{3}', first):
            tags.append(first)
            talk_track_lines = talk_track_lines[1:]

    talk_track = " ".join(talk_track_lines)
    return tags, talk_track


def get_slide_title(slide):
    """Extract the title text from a slide."""
    if slide.shapes.title:
        return slide.shapes.title.text.strip()
    # Fallback: look for the first text shape
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            return shape.text_frame.text.strip()[:100]
    return ""


def get_slide_subtitle(slide):
    """Extract subtitle text (second text element) from a slide."""
    texts = []
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            texts.append(shape.text_frame.text.strip())
    if len(texts) >= 2:
        # First is title, second might be subtitle
        candidate = texts[1]
        # Only return if it's short enough to be a subtitle
        if len(candidate) < 200:
            return candidate
    return ""


def classify_slide(slide):
    """Classify slide as 'section' or 'diagram'."""
    shape_count = len(slide.shapes)
    if shape_count <= SECTION_MAX_SHAPES:
        return "section"
    return "diagram"


def build_catalog():
    if not CATALOG_PPTX.exists():
        print(f"Error: Catalog not found at {CATALOG_PPTX}")
        sys.exit(1)

    prs = Presentation(str(CATALOG_PPTX))
    total_slides = len(prs.slides)
    print(f"Reading {total_slides} slides from {CATALOG_PPTX.name}")

    slides = []
    current_section = None
    section_order = 0

    for idx, slide in enumerate(prs.slides):
        slide_num = idx + 1  # 1-based

        if slide_num in SKIP_SLIDES:
            continue

        title = get_slide_title(slide)
        subtitle = get_slide_subtitle(slide)
        kind = classify_slide(slide)
        tags, talk_track = extract_tags_and_talk_track(slide)
        shape_count = len(slide.shapes)

        # Track sections
        if kind == "section":
            current_section = title
            section_order += 1

        entry = {
            "slide": slide_num,
            "title": title,
            "kind": kind,
            "tags": tags,
            "shape_count": shape_count,
        }

        if subtitle and subtitle != title:
            entry["subtitle"] = subtitle

        if talk_track:
            # Truncate talk track preview to 300 chars
            entry["talk_track"] = talk_track[:300] + ("..." if len(talk_track) > 300 else "")

        if current_section and kind != "section":
            entry["section"] = current_section

        slides.append(entry)

    # Build section index
    sections = []
    seen_sections = {}
    for s in slides:
        if s["kind"] == "section":
            if s["title"] not in seen_sections:
                seen_sections[s["title"]] = len(sections)
                sections.append({
                    "title": s["title"],
                    "slide": s["slide"],
                    "diagrams": []
                })
        elif s["kind"] == "diagram" and s.get("section"):
            sec_idx = seen_sections.get(s["section"])
            if sec_idx is not None:
                sections[sec_idx]["diagrams"].append(s["slide"])

    catalog = {
        "source": CATALOG_PPTX.name,
        "total_slides": total_slides,
        "importable_slides": len([s for s in slides if s["kind"] == "diagram"]),
        "sections": sections,
        "slides": slides,
    }

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(catalog, f, indent=2, ensure_ascii=False)

    diagram_count = catalog["importable_slides"]
    section_count = len(sections)
    print(f"Wrote {OUTPUT_JSON}")
    print(f"  Sections: {section_count}")
    print(f"  Diagrams: {diagram_count}")
    print(f"  Total indexed: {len(slides)} (skipped {len(SKIP_SLIDES)} meta slides)")


if __name__ == "__main__":
    build_catalog()
