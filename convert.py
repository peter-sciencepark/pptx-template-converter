#!/usr/bin/env python3
"""
pptx-template-converter

Converts PowerPoint presentations from an old layout/template
to a new one, mapping slide types and transferring content.

Usage:
    python3 convert.py input.pptx --template template.potx -o output.pptx
"""

import argparse
import copy
import io
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import CONTENT_TYPE as CT


def open_potx_as_presentation(potx_path: str) -> Presentation:
    """Open a .potx template file as a Presentation by patching the content type."""
    buf = io.BytesIO()
    with zipfile.ZipFile(potx_path, "r") as z_in, zipfile.ZipFile(buf, "w") as z_out:
        for item in z_in.infolist():
            data = z_in.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
                    b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                )
            z_out.writestr(item, data)
    buf.seek(0)
    return Presentation(buf)


def get_layout_map(template_prs: Presentation) -> dict:
    """Build a name -> layout mapping from the template."""
    layouts = {}
    for master in template_prs.slide_masters:
        for layout in master.slide_layouts:
            layouts[layout.name] = layout
    return layouts


def _safe_color(font):
    """Safely extract RGB color string from a font, or return None."""
    try:
        if font.color and font.color.type is not None:
            return str(font.color.rgb)
    except AttributeError:
        pass
    return None


def extract_all_text(shape) -> list[dict]:
    """Extract text runs from a shape, preserving paragraph structure."""
    paragraphs = []
    if not shape.has_text_frame:
        return paragraphs
    for para in shape.text_frame.paragraphs:
        runs = []
        for run in para.runs:
            runs.append({
                "text": run.text,
                "bold": run.font.bold,
                "italic": run.font.italic,
                "size": run.font.size,
                "color": _safe_color(run.font),
            })
        if not runs and para.text:
            runs.append({"text": para.text, "bold": None, "italic": None, "size": None, "color": None})
        paragraphs.append({
            "runs": runs,
            "alignment": para.alignment,
            "level": para.level,
        })
    return paragraphs


def classify_slide(slide, index: int, total: int) -> str:
    """
    Classify a slide into a category based on its content and position.

    Returns one of:
        'title'       - Title/cover slide
        'chapter'     - Chapter heading (short text, no bullets)
        'quote'       - Quote slide
        'content'     - Content with text/bullets
        'content_img' - Content with image
        'closing'     - Last slide / closing
        'blank'       - Empty or minimal content
    """
    layout_name = slide.slide_layout.name.lower()
    shapes = list(slide.shapes)

    has_picture = any(s.shape_type == 13 for s in shapes)  # MSO_SHAPE_TYPE.PICTURE
    text_shapes = [s for s in shapes if hasattr(s, "text") and s.text.strip()]
    all_text = " ".join(s.text for s in text_shapes).strip()
    word_count = len(all_text.split()) if all_text else 0

    # First slide is typically a title
    if index == 0:
        return "title"

    # Last slide is typically closing
    if index == total - 1 and word_count < 30:
        return "closing"

    # Quote detection
    if all_text.startswith('"') or all_text.startswith('\u201c') or all_text.startswith('\u201d'):
        return "quote"
    if "citat" in layout_name:
        return "quote"

    # Title-like slides (rubrikbild) with very little text
    if "rubrik" in layout_name and "content" not in layout_name:
        if word_count < 20 and not has_picture:
            return "chapter"

    # Blank
    if word_count == 0 and not has_picture:
        return "blank"

    # Content with image
    if has_picture:
        return "content_img"

    # Short text without bullets = chapter
    if word_count < 15 and len(text_shapes) <= 2:
        return "chapter"

    return "content"


def map_to_new_layout(category: str, layouts: dict, has_image: bool = False) -> str:
    """Map a slide category to the best matching new template layout name."""
    mapping = {
        "title": ["1 - Rubrikbild logo", "1 - Rubrikbild blank"],
        "chapter": ["1 - Kapitelrubrik", "1 - Kapitelrubrik med underrubrik"],
        "quote": ["11 - Midicitat blå", "11 - Maxicitat blå"],
        "content": ["4 - Innehåll blank", "5 - Innehåll blank", "7 - Innehåll blank"],
        "content_img": ["4 - Bild höger", "5 - Bild höger", "7 - Bild höger"],
        "closing": ["14 - Slogan hav", "14 - Slogan skog", "13 - Bakgrund hav"],
        "blank": ["4 - Innehåll blank", "5 - Innehåll blank"],
    }

    candidates = mapping.get(category, ["4 - Innehåll blank"])
    for name in candidates:
        if name in layouts:
            return name

    # Fallback: first available content layout
    for name in layouts:
        if "Innehåll" in name:
            return name
    return list(layouts.keys())[0]


def apply_text_to_placeholder(placeholder, paragraphs: list[dict]):
    """Write extracted paragraph data into a placeholder's text frame."""
    tf = placeholder.text_frame
    tf.clear()

    for i, para_data in enumerate(paragraphs):
        if not para_data["runs"]:
            continue

        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        if para_data["alignment"] is not None:
            p.alignment = para_data["alignment"]
        p.level = para_data.get("level", 0) or 0

        for j, run_data in enumerate(para_data["runs"]):
            if j == 0:
                r = p.runs[0] if p.runs else p.add_run()
            else:
                r = p.add_run()
            r.text = run_data["text"]
            if run_data["bold"] is not None:
                r.font.bold = run_data["bold"]
            if run_data["italic"] is not None:
                r.font.italic = run_data["italic"]


def find_title_text(slide) -> str:
    """Find the most likely title text from a slide."""
    # Check placeholders first
    for shape in slide.placeholders:
        if shape.placeholder_format.type in (1, 3):  # TITLE, CENTER_TITLE
            return shape.text.strip()

    # Check named shapes
    for shape in slide.shapes:
        name = shape.name.lower()
        if "rubrik" in name or "title" in name or "titel" in name:
            if hasattr(shape, "text") and shape.text.strip():
                return shape.text.strip()

    return ""


def find_body_text(slide) -> list[dict]:
    """Find the main body/content text from a slide (excluding title)."""
    title_text = find_title_text(slide)
    all_paragraphs = []

    for shape in slide.shapes:
        if not hasattr(shape, "text") or not shape.text.strip():
            continue
        # Skip if this is the title
        if shape.text.strip() == title_text and title_text:
            continue
        # Skip small reference/source texts
        if shape.width and shape.width < Inches(4) and len(shape.text) < 50:
            text_lower = shape.text.lower()
            if any(kw in text_lower for kw in ["källa", "source", "ref"]):
                continue

        paragraphs = extract_all_text(shape)
        all_paragraphs.extend(paragraphs)

    return all_paragraphs


def find_images(slide) -> list:
    """Find all picture shapes in a slide."""
    return [s for s in slide.shapes if s.shape_type == 13]


def convert_slide(old_slide, new_prs, layouts, category, slide_index):
    """Convert a single slide from old to new template."""
    has_image = bool(find_images(old_slide))
    layout_name = map_to_new_layout(category, layouts, has_image)
    layout = layouts[layout_name]
    new_slide = new_prs.slides.add_slide(layout)

    title_text = find_title_text(old_slide)
    body_paragraphs = find_body_text(old_slide)
    images = find_images(old_slide)

    # Find available placeholders in new slide
    title_ph = None
    body_ph = None
    picture_ph = None

    for ph in new_slide.placeholders:
        ph_type = ph.placeholder_format.type
        if ph_type in (1, 3):  # TITLE, CENTER_TITLE
            title_ph = ph
        elif ph_type == 18:  # PICTURE
            picture_ph = ph
        elif ph_type in (2, 4, 7):  # BODY, SUBTITLE, OBJECT
            if body_ph is None:
                body_ph = ph

    # Also check layout placeholders that might need to be accessed
    for ph in layout.placeholders:
        ph_type = ph.placeholder_format.type
        idx = ph.placeholder_format.idx
        if ph_type in (1, 3) and title_ph is None:
            try:
                title_ph = new_slide.placeholders[idx]
            except KeyError:
                pass
        elif ph_type in (2, 4, 7) and body_ph is None:
            try:
                body_ph = new_slide.placeholders[idx]
            except KeyError:
                pass

    # Apply title
    if title_text and title_ph:
        title_ph.text = title_text
    elif title_text and not title_ph:
        # No title placeholder — add as textbox
        from pptx.util import Inches, Pt
        txBox = new_slide.shapes.add_textbox(
            Inches(0.92), Inches(1.02), Inches(11.5), Inches(0.83)
        )
        txBox.text_frame.paragraphs[0].text = title_text
        txBox.text_frame.paragraphs[0].font.size = Pt(28)
        txBox.text_frame.paragraphs[0].font.bold = True

    # Apply body content
    if body_paragraphs and body_ph:
        apply_text_to_placeholder(body_ph, body_paragraphs)
    elif body_paragraphs and not body_ph:
        # No body placeholder — add as textbox
        from pptx.util import Inches, Pt
        top = Inches(2.1) if title_text else Inches(1.5)
        txBox = new_slide.shapes.add_textbox(
            Inches(0.92), top, Inches(11.5), Inches(4.5)
        )
        apply_text_to_placeholder(txBox, body_paragraphs)

    # Handle images
    if images and picture_ph:
        img = images[0]
        try:
            image_stream = io.BytesIO(img.image.blob)
            # For picture placeholders, insert the image
            picture_ph.insert_picture(image_stream)
        except Exception:
            pass  # Skip if image insertion fails
    elif images:
        # No picture placeholder — add image as a free shape
        for img in images:
            try:
                image_stream = io.BytesIO(img.image.blob)
                new_slide.shapes.add_picture(
                    image_stream,
                    img.left, img.top,
                    img.width, img.height,
                )
            except Exception:
                pass

    return new_slide, layout_name


def convert_presentation(input_path: str, template_path: str, output_path: str):
    """Main conversion function."""
    print(f"Opening input:    {input_path}")
    old_prs = Presentation(input_path)

    print(f"Opening template: {template_path}")
    if template_path.endswith(".potx"):
        new_prs = open_potx_as_presentation(template_path)
    else:
        new_prs = Presentation(template_path)

    # Remove example slides from template
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    while len(new_prs.slides) > 0:
        sldId = new_prs.slides._sldIdLst[0]
        rId = sldId.get(f"{{{ns_r}}}id") or sldId.get("r:id")
        if rId:
            new_prs.part.drop_rel(rId)
        new_prs.slides._sldIdLst.remove(sldId)

    layouts = get_layout_map(new_prs)
    print(f"\nAvailable layouts in new template:")
    for name in sorted(layouts.keys()):
        print(f"  - {name}")

    total_slides = len(old_prs.slides)
    print(f"\nConverting {total_slides} slides...\n")

    for i, slide in enumerate(old_prs.slides):
        category = classify_slide(slide, i, total_slides)
        old_layout = slide.slide_layout.name
        new_slide, new_layout = convert_slide(slide, new_prs, layouts, category, i)
        print(f"  Slide {i+1}/{total_slides}: [{category:12s}] '{old_layout}' -> '{new_layout}'")

    new_prs.save(output_path)
    print(f"\nSaved: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Convert PowerPoint presentations to a new template layout."
    )
    parser.add_argument("input", help="Input .pptx file to convert")
    parser.add_argument(
        "--template", "-t", required=True,
        help="New template file (.potx or .pptx)"
    )
    parser.add_argument(
        "--output", "-o", default=None,
        help="Output .pptx file (default: input_converted.pptx)"
    )
    args = parser.parse_args()

    if args.output is None:
        stem = Path(args.input).stem
        args.output = str(Path(args.input).parent / f"{stem}_converted.pptx")

    convert_presentation(args.input, args.template, args.output)


if __name__ == "__main__":
    main()
