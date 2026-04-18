"""
build_pptx.py — Builds a 14-slide .pptx from ppt_content.json + screenshots
fetched from the website-decoder repo via raw GitHub URLs.

Usage:
    python build_pptx.py --domain wisprflow.ai \
                         --run-id 20260418-1244 \
                         --repo1 Rahul-peacenik159/website-decoder
"""

import argparse
import io
import json
import re
import sys
from pathlib import Path

import requests
from PIL import Image as PILImage
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ── Slide dimensions: 16:9 widescreen ────────────────────────────────────────
SLIDE_W = Inches(10)
SLIDE_H = Inches(5.625)

# ── Type scale ────────────────────────────────────────────────────────────────
FONT_COVER_TITLE = Pt(40)
FONT_TITLE = Pt(26)
FONT_SUBTITLE = Pt(16)
FONT_BODY = Pt(14)
FONT_LABEL = Pt(11)
FONT_QUOTE = Pt(26)

# ── Geometry ──────────────────────────────────────────────────────────────────
MARGIN = Inches(0.55)

# ── Fallback colors ───────────────────────────────────────────────────────────
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK_NAVY = RGBColor(0x1A, 0x1A, 0x2E)
LIGHT_BG = RGBColor(0xF8, 0xF8, 0xF8)
MID_GRAY = RGBColor(0x44, 0x44, 0x55)


# ── Helpers ───────────────────────────────────────────────────────────────────

def hex_to_rgb(hex_str: str) -> RGBColor:
    try:
        h = hex_str.lstrip("#")
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        return DARK_NAVY


def is_dark(hex_str: str) -> bool:
    try:
        h = hex_str.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.45
    except Exception:
        return True


def fetch_bytes(url: str, retries: int = 3) -> bytes:
    for attempt in range(retries):
        try:
            r = requests.get(url, timeout=25)
            r.raise_for_status()
            return r.content
        except Exception as e:
            if attempt < retries - 1:
                import time; time.sleep(10)
            else:
                print(f"  ! Failed to fetch {url}: {e}")
    return None


def fetch_screenshot(base_url: str, filename: str, cache_dir: Path) -> str:
    cache_dir.mkdir(parents=True, exist_ok=True)
    local = cache_dir / filename
    if local.exists():
        return str(local)
    data = fetch_bytes(f"{base_url}/{filename}")
    if data is None:
        return None
    local.write_bytes(data)
    return str(local)


def blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])  # blank layout


def set_bg(slide, hex_color: str):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(hex_color)


def add_rect(slide, left, top, width, height, fill_hex: str, line=False):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    if line:
        shape.line.color.rgb = WHITE
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    return shape


def add_text(slide, text: str, left, top, width, height,
             size=None, bold=False, color: RGBColor = WHITE,
             align=PP_ALIGN.LEFT, wrap=True) -> None:
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = size or FONT_BODY
    run.font.bold = bold
    run.font.color.rgb = color


def add_bullets(slide, items: list, left, top, width, height,
                size=None, color: RGBColor = WHITE) -> None:
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf = txb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_before = Pt(5)
        if item.startswith("✓"):
            c = RGBColor(0x4C, 0xAF, 0x50)
        elif item.startswith("✗"):
            c = RGBColor(0xF4, 0x43, 0x36)
        else:
            c = color
        run = p.add_run()
        run.text = item
        run.font.size = size or FONT_BODY
        run.font.color.rgb = c


def add_image(slide, img_path, left, top, width, height, accent_hex: str):
    if img_path and Path(img_path).exists():
        try:
            # Quick validation with Pillow
            with PILImage.open(img_path) as im:
                im.verify()
            slide.shapes.add_picture(img_path, left, top, width=width, height=height)
            return
        except Exception as e:
            print(f"  ! Image load failed: {e}")
    # Placeholder
    shape = add_rect(slide, left, top, width, height, accent_hex, line=True)
    tf = shape.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = "[ screenshot ]"
    run.font.color.rgb = WHITE
    run.font.size = FONT_LABEL


def accent_line(slide, accent_hex: str):
    """Thin horizontal accent line at very top of slide."""
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Pt(5), accent_hex)


# ── Layout builders ───────────────────────────────────────────────────────────

def build_cover(prs, s, img_path):
    slide = blank_slide(prs)
    bg = s.get("bg_color", "#1A1A2E")
    set_bg(slide, bg)
    accent = s.get("accent_color", "#7C3AED")
    txt_color = WHITE

    # Screenshot: right side
    add_image(slide, img_path, Inches(5.3), Inches(0.3), Inches(4.35), Inches(5.0), accent)

    # Accent bar left edge
    add_rect(slide, MARGIN, Inches(1.1), Pt(4), Inches(2.5), accent)

    # Title
    add_text(slide, s.get("title", "Website Decode"),
             Inches(0.75), Inches(1.1), Inches(4.3), Inches(1.4),
             size=FONT_COVER_TITLE, bold=True, color=txt_color)

    # Subtitle (domain)
    add_text(slide, s.get("subtitle", ""),
             Inches(0.75), Inches(2.6), Inches(4.3), Inches(0.6),
             size=FONT_SUBTITLE, color=hex_to_rgb(accent))

    # Footer: "Website Decode"
    add_text(slide, "Website Decode",
             MARGIN, Inches(5.1), Inches(5.0), Inches(0.4),
             size=Pt(10), color=RGBColor(0x88, 0x88, 0x99))


def build_two_column(prs, s, img_path):
    slide = blank_slide(prs)
    accent = s.get("accent_color", "#7C3AED")
    set_bg(slide, "#FFFFFF")
    txt_color = DARK_NAVY

    accent_line(slide, accent)

    # Title
    add_text(slide, s.get("title", ""),
             MARGIN, Inches(0.25), Inches(9.0), Inches(0.7),
             size=FONT_TITLE, bold=True, color=txt_color)

    # Divider line under title
    add_rect(slide, MARGIN, Inches(1.0), Inches(4.2), Pt(1.5), accent)

    # Bullets (left column)
    body = s.get("body", [])
    add_bullets(slide, body,
                MARGIN, Inches(1.15), Inches(4.2), Inches(4.1),
                size=FONT_BODY, color=txt_color)

    # Screenshot (right column)
    add_image(slide, img_path, Inches(5.1), Inches(0.9), Inches(4.4), Inches(4.4), accent)


def build_bullets(prs, s):
    slide = blank_slide(prs)
    accent = s.get("accent_color", "#7C3AED")
    set_bg(slide, "#FFFFFF")
    txt_color = DARK_NAVY

    accent_line(slide, accent)

    # Title
    add_text(slide, s.get("title", ""),
             MARGIN, Inches(0.25), Inches(9.0), Inches(0.7),
             size=FONT_TITLE, bold=True, color=txt_color)

    subtitle = s.get("subtitle")
    body_top = Inches(1.15)
    if subtitle:
        add_text(slide, subtitle,
                 MARGIN, Inches(1.05), Inches(9.0), Inches(0.45),
                 size=Pt(15), color=hex_to_rgb(accent))
        body_top = Inches(1.6)

    # Divider
    add_rect(slide, MARGIN, Inches(1.0), Inches(9.0), Pt(1.5), accent)

    body = s.get("body", [])
    add_bullets(slide, body,
                MARGIN, body_top, Inches(9.0), Inches(4.0),
                size=FONT_BODY, color=txt_color)


def build_color_palette(prs, s):
    slide = blank_slide(prs)
    accent = s.get("accent_color", "#7C3AED")
    set_bg(slide, "#0D0D1A")

    accent_line(slide, accent)

    add_text(slide, s.get("title", "Brand Color System"),
             MARGIN, Inches(0.25), Inches(9.0), Inches(0.7),
             size=FONT_TITLE, bold=True, color=WHITE)

    add_rect(slide, MARGIN, Inches(1.0), Inches(9.0), Pt(1.5), accent)

    colors = s.get("colors", [])
    swatch_w = Inches(1.5)
    swatch_h = Inches(2.0)
    gap = Inches(0.28)
    top = Inches(1.2)

    for i, c in enumerate(colors[:8]):
        left = MARGIN + i * (swatch_w + gap)
        hex_val = c.get("hex", "#888888")
        role = c.get("role", "")

        add_rect(slide, left, top, swatch_w, swatch_h, hex_val)

        add_text(slide, hex_val.upper(),
                 left, top + swatch_h + Pt(4), swatch_w, Inches(0.3),
                 size=FONT_LABEL, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        add_text(slide, role,
                 left, top + swatch_h + Inches(0.32), swatch_w, Inches(0.28),
                 size=Pt(10), color=RGBColor(0xAA, 0xAA, 0xBB), align=PP_ALIGN.CENTER)


def build_quote(prs, s):
    slide = blank_slide(prs)
    bg = s.get("bg_color", "#1A1A2E")
    accent = s.get("accent_color", "#7C3AED")
    set_bg(slide, bg)

    accent_line(slide, accent)

    # Decorative quote mark
    add_text(slide, "\u201c",
             MARGIN, Inches(0.6), Inches(1.5), Inches(1.5),
             size=Pt(80), bold=True, color=hex_to_rgb(accent))

    # Quote text
    add_text(slide, s.get("quote", ""),
             Inches(1.2), Inches(1.2), Inches(7.8), Inches(3.2),
             size=FONT_QUOTE, bold=False, color=WHITE, align=PP_ALIGN.LEFT)

    # Attribution
    add_text(slide, "— " + s.get("attribution", ""),
             Inches(1.2), Inches(4.6), Inches(8.0), Inches(0.55),
             size=Pt(15), color=hex_to_rgb(accent))


# ── Dispatcher ────────────────────────────────────────────────────────────────

BUILDERS = {
    "cover": build_cover,
    "two-column": build_two_column,
    "bullets": build_bullets,
    "color-palette": build_color_palette,
    "quote": build_quote,
}

SCREENSHOT_LAYOUTS = {"cover", "two-column"}


def build_presentation(content: dict, cache_dir: Path, screenshots_base: str) -> Presentation:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    for s in content.get("slides", []):
        layout = s.get("layout", "bullets")
        builder = BUILDERS.get(layout)
        if builder is None:
            print(f"  ! Unknown layout '{layout}' on slide {s.get('id')} — falling back to bullets")
            build_bullets(prs, s)
            continue

        img_path = None
        if layout in SCREENSHOT_LAYOUTS:
            fname = s.get("screenshot")
            if fname:
                img_path = fetch_screenshot(screenshots_base, fname, cache_dir)

        if layout in ("cover", "two-column"):
            builder(prs, s, img_path)
        else:
            builder(prs, s)

    return prs


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--domain", required=True)
    parser.add_argument("--run-id", required=True)
    parser.add_argument("--repo1", required=True, help="owner/repo of website-decoder")
    args = parser.parse_args()

    domain = args.domain
    run_id = args.run_id
    repo1 = args.repo1

    raw_base = f"https://raw.githubusercontent.com/{repo1}/main"
    json_url = f"{raw_base}/output/decodes/{domain}/ppt_content.json"

    print(f"Fetching ppt_content.json from {json_url}")
    data = fetch_bytes(json_url)
    if data is None:
        print("ERROR: Could not fetch ppt_content.json")
        sys.exit(1)

    try:
        content = json.loads(data.decode("utf-8"))
    except json.JSONDecodeError as e:
        print(f"ERROR: Invalid JSON: {e}")
        sys.exit(1)

    screenshots_base = f"{raw_base}/output/screenshots/{domain}/{run_id}"
    cache_dir = Path("tmp") / domain / run_id

    print(f"Building presentation: {domain} / {run_id}")
    prs = build_presentation(content, cache_dir, screenshots_base)

    out_dir = Path("output")
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{domain}-{run_id}.pptx"
    prs.save(str(out_path))
    print(f"Saved: {out_path}")


if __name__ == "__main__":
    main()
