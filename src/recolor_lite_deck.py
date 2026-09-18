"""Recolor the existing LITE deck from green to UChicago maroon without changing content."""
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import shutil
import tempfile
import zipfile

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "final_slideshow" / "LITE_Round1_Pitch.pptx"
OUTPUT = ROOT / "final_slideshow" / "Final_Lumentum_Bull_Thesis.pptx"

COLOR_MAP = {
    "2F7A38": "800000",  # primary green -> UChicago maroon
    "1B4D1E": "5C0000",  # dark green -> dark maroon
    "66A36B": "B13F3F",  # secondary green -> muted red
    "E8F5E9": "F5ECEC",  # pale green -> pale warm red
    "C8E6C9": "E7C7C7",  # light green -> light warm red
}


def recolor_color(color) -> None:
    try:
        if color.type == MSO_COLOR_TYPE.RGB and color.rgb is not None:
            old = str(color.rgb).upper()
            if old in COLOR_MAP:
                color.rgb = RGBColor.from_string(COLOR_MAP[old])
    except (AttributeError, TypeError, ValueError):
        pass


def recolor_text_frame(tf) -> None:
    for para in tf.paragraphs:
        recolor_color(para.font.color)
        for run in para.runs:
            recolor_color(run.font.color)


def recolor_shape(shape) -> None:
    try:
        recolor_color(shape.fill.fore_color)
    except (AttributeError, TypeError):
        pass
    try:
        recolor_color(shape.line.color)
    except (AttributeError, TypeError):
        pass
    if getattr(shape, "has_text_frame", False):
        recolor_text_frame(shape.text_frame)
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                recolor_color(cell.fill.fore_color)
                recolor_text_frame(cell.text_frame)
    if getattr(shape, "shape_type", None) == 6:  # group
        for child in shape.shapes:
            recolor_shape(child)


def recolor_raster(blob: bytes, suffix: str) -> bytes:
    try:
        im = Image.open(BytesIO(blob)).convert("RGBA")
    except Exception:
        return blob
    pix = im.load()
    palettes = [
        ((47, 122, 56), (128, 0, 0)),
        ((27, 77, 30), (92, 0, 0)),
        ((102, 163, 107), (177, 63, 63)),
        ((232, 245, 233), (245, 236, 236)),
        ((200, 230, 201), (231, 199, 199)),
    ]
    for y in range(im.height):
        for x in range(im.width):
            r, g, b, a = pix[x, y]
            # Only touch pixels that are visibly green or extremely close to a known pale green.
            if not (g > r * 1.06 and g > b * 1.04):
                continue
            best = min(palettes, key=lambda p: (r-p[0][0])**2 + (g-p[0][1])**2 + (b-p[0][2])**2)
            src, dst = best
            dist = ((r-src[0])**2 + (g-src[1])**2 + (b-src[2])**2) ** .5
            if dist > 85:
                continue
            # Preserve antialiasing/tonal variation around the source palette color.
            delta = ((r-src[0]) + (g-src[1]) + (b-src[2])) / 3
            pix[x, y] = tuple(max(0, min(255, round(v + delta))) for v in dst) + (a,)
    out = BytesIO()
    fmt = "PNG" if suffix.lower() == ".png" else (im.format or "PNG")
    im.save(out, format=fmt)
    return out.getvalue()


prs = Presentation(SOURCE)
for slide in prs.slides:
    for shape in slide.shapes:
        recolor_shape(shape)

with tempfile.TemporaryDirectory() as td:
    td = Path(td)
    intermediate = td / "shapes.pptx"
    prs.save(intermediate)
    unpacked = td / "pkg"
    with zipfile.ZipFile(intermediate) as zf:
        zf.extractall(unpacked)
    media = unpacked / "ppt" / "media"
    if media.exists():
        for item in media.iterdir():
            if item.suffix.lower() in {".png", ".jpg", ".jpeg"}:
                item.write_bytes(recolor_raster(item.read_bytes(), item.suffix))
    rebuilt = td / "rebuilt.pptx"
    with zipfile.ZipFile(rebuilt, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in unpacked.rglob("*"):
            if item.is_file():
                zf.write(item, item.relative_to(unpacked).as_posix())
    shutil.copy2(rebuilt, OUTPUT)

print(OUTPUT)
print(f"slides={len(prs.slides)}")
